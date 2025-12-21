#!/usr/bin/env python3
"""
OCR_line_geometry_solver.core

Snapshot of the line-geometry-based OCR core used to suppress table/box lines.
This module proxies through to eidp_term_scanner.core.py and overrides the
EasyOCR boxes path with geometric line detection + masking.
"""
from __future__ import annotations

from pathlib import Path as _Path
import importlib.util as _importlib
import math as _math
import os as _os
import sys as _sys
from typing import Any as _Any, Dict as _Dict, List as _List, Sequence as _Sequence, Tuple as _Tuple

import numpy as _np  # type: ignore
try:
    import cv2  # type: ignore
except Exception:
    cv2 = None  # type: ignore

try:
    import fitz  # type: ignore
except Exception:
    fitz = None  # type: ignore


_ORIG_CORE_PATH = _Path(__file__).with_name("eidp_term_scanner.core.py")

_spec = _importlib.spec_from_file_location("eidp_core_orig", _ORIG_CORE_PATH)
if _spec is None or _spec.loader is None:
    raise RuntimeError(f"Unable to load original core module from {_ORIG_CORE_PATH}")
_orig = _importlib.module_from_spec(_spec)
_spec.loader.exec_module(_orig)  # type: ignore[attr-defined]

# Re-export everything from the original core so this module is a drop-in.
for _name in dir(_orig):
    if _name.startswith("__"):
        continue
    globals()[_name] = getattr(_orig, _name)


_DEFAULT_MAX_EASYOCR_DPI = 800

# Threshold that governs when we try backup OCR retries.
_BACKUP_CONF_THRESHOLD = 0.5
_TESS_BIN: str | None = None


def _backup_enabled() -> bool:
    """Return whether multi-engine backup OCR logic should be used.

    Controlled via OCR_ENABLE_BACKUP environment variable so we can easily
    disable this path while experimenting with other refinements.
    """
    try:
        val = (_os.environ.get("OCR_ENABLE_BACKUP") or "").strip().lower()
        return val in ("1", "true", "yes")
    except Exception:
        return False


def _clamp_ocr_dpi(requested: int) -> int:
    """Clamp requested OCR DPI to avoid excessive raster sizes."""
    max_dpi = _DEFAULT_MAX_EASYOCR_DPI
    try:
        env_val = (_os.environ.get("OCR_MAX_DPI") or "").strip()
        if env_val:
            cand = int(env_val)
            if cand > 0:
                max_dpi = cand
    except Exception:
        pass
    if requested <= max_dpi:
        return requested
    try:
        print(f"[INFO] OCR DPI capped from {requested} to {max_dpi}", flush=True)
    except Exception:
        pass
    return max_dpi


def _tesseract_backup_enabled() -> bool:
    """Return True if Tesseract backup OCR is enabled and the binary is available.

    Controlled via OCR_ENABLE_TESSERACT environment variable. We keep this separate
    from the main app's Tesseract wiring so the line-geometry core can experiment
    independently.
    """
    try:
        val = (_os.environ.get("OCR_ENABLE_TESSERACT") or "").strip().lower()
        if val not in ("1", "true", "yes"):
            return False
    except Exception:
        return False
    global _TESS_BIN
    if _TESS_BIN:
        return True
    try:
        import shutil as _sh  # type: ignore

        _TESS_BIN = _sh.which("tesseract")
    except Exception:
        _TESS_BIN = None
    return bool(_TESS_BIN)


def _angle_diff_deg(a: float, b: float) -> float:
    """Return the smallest absolute difference between two angles in degrees (0-180)."""
    diff = abs(a - b) % 180.0
    return diff if diff <= 90.0 else 180.0 - diff


def _bbox_center(box: _Sequence[_Sequence[float]]) -> _Tuple[float, float]:
    xs = [float(p[0]) for p in box]
    ys = [float(p[1]) for p in box]
    return sum(xs) / 4.0, sum(ys) / 4.0


def _bbox_long_short(box: _Sequence[_Sequence[float]]) -> _Tuple[float, float, float]:
    lens_ang: _List[_Tuple[float, float]] = []
    for i in range(4):
        x0, y0 = box[i]
        x1, y1 = box[(i + 1) % 4]
        dx, dy = float(x1) - float(x0), float(y1) - float(y0)
        l = _math.hypot(dx, dy)
        ang = _math.degrees(_math.atan2(dy, dx)) if l > 0 else 0.0
        lens_ang.append((l, ang))
    long_len, long_ang = max(lens_ang, key=lambda t: t[0])
    short_len = min(l for l, _ in lens_ang)
    return long_len, short_len, long_ang


def _dist_point_to_line(cx: float, cy: float, line: _Tuple[float, float, float, float, float, float]) -> float:
    x1, y1, x2, y2, _, _ = line
    num = abs((y2 - y1) * cx - (x2 - x1) * cy + x2 * y1 - y2 * x1)
    den = _math.hypot(y2 - y1, x2 - x1)
    return num / den if den else 1e9


def _tesseract_ocr_crop(
    crop: _np.ndarray,
    *,
    allowlist: str | None = None,
    debug: bool = False,
) -> _Tuple[str, float]:
    """Run Tesseract on a small RGB crop and return (text, conf[0-1])."""
    if crop is None or crop.size == 0:
        return "", 0.0
    if not _tesseract_backup_enabled():
        return "", 0.0
    global _TESS_BIN
    tess_bin = _TESS_BIN
    if not tess_bin:
        return "", 0.0
    try:
        import tempfile as _tmp  # type: ignore
        import subprocess as _sp  # type: ignore
        from PIL import Image as _Image  # type: ignore
    except Exception:
        return "", 0.0
    try:
        with _tmp.NamedTemporaryFile(suffix=".png", delete=False) as tf:
            _Image.fromarray(crop).save(tf.name)
            tf_path = tf.name
    except Exception:
        return "", 0.0
    cmd = [
        tess_bin,
        tf_path,
        "stdout",
        "-l",
        "eng",
        "--psm",
        "7",
        "--oem",
        "3",
    ]
    if allowlist:
        cmd.extend(["-c", f"tessedit_char_whitelist={allowlist}"])
    cmd.append("tsv")
    try:
        proc = _sp.run(cmd, capture_output=True, text=True, check=False)
    except Exception:
        proc = None  # type: ignore[assignment]
    try:
        _os.remove(tf_path)
    except Exception:
        pass
    if proc is None or proc.returncode != 0 or not proc.stdout:
        return "", 0.0
    best_text = ""
    best_conf = 0.0
    for line in proc.stdout.splitlines():
        parts = line.split("\t")
        if len(parts) < 12:
            continue
        txt = parts[11].strip()
        if not txt:
            continue
        try:
            cval = float(parts[10]) / 100.0
        except Exception:
            cval = 0.0
        if cval > best_conf:
            best_conf = cval
            best_text = txt
    if debug:
        try:
            _sys.stderr.write(
                f"[OCR TESS BACKUP] text={best_text!r} conf={best_conf:.3f}\n"
            )
        except Exception:
            pass
    return best_text.strip(), best_conf


def _detect_lines(img_rgb: _np.ndarray) -> _List[_Tuple[float, float, float, float, float, float]]:
    lines: _List[_Tuple[float, float, float, float, float, float]] = []
    if cv2 is None:
        return lines
    try:
        gray = cv2.cvtColor(img_rgb, cv2.COLOR_RGB2GRAY)
        edges = cv2.Canny(gray, 40, 120, apertureSize=3)
        h, w = edges.shape
        min_len = max(int(min(h, w) * 0.05), 12)
        lines_p = cv2.HoughLinesP(edges, 1, _math.pi / 180.0, threshold=40, minLineLength=min_len, maxLineGap=6)
        if lines_p is not None:
            for x1, y1, x2, y2 in lines_p[:, 0, :]:
                dx, dy = float(x2 - x1), float(y2 - y1)
                length = _math.hypot(dx, dy)
                if length < min_len:
                    continue
                angle = _math.degrees(_math.atan2(dy, dx))
                lines.append((float(x1), float(y1), float(x2), float(y2), length, angle))
        # Optional: try LSD for missed lines
        try:
            lsd = cv2.createLineSegmentDetector(cv2.LSD_REFINE_STD)
            lsd_lines, _, _, _ = lsd.detect(edges)
            if lsd_lines is not None:
                for l in lsd_lines:
                    x1, y1, x2, y2 = l[0]
                    dx, dy = float(x2 - x1), float(y2 - y1)
                    length = _math.hypot(dx, dy)
                    if length < min_len:
                        continue
                    angle = _math.degrees(_math.atan2(dy, dx))
                    lines.append((float(x1), float(y1), float(x2), float(y2), length, angle))
        except Exception:
            pass
    except Exception:
        pass
    return lines


def _is_line_artifact(
    box: _Sequence[_Sequence[float]],
    text: str,
    lines: _List[_Tuple[float, float, float, float, float, float]],
    *,
    confidence: float = 0.0,
) -> bool:
    if not lines:
        return False
    cx, cy = _bbox_center(box)
    long_len, short_len, ang_token = _bbox_long_short(box)
    text_single = len(text.strip()) <= 2
    xs = [float(p[0]) for p in box]
    ys = [float(p[1]) for p in box]
    w_box = max(xs) - min(xs)
    h_box = max(ys) - min(ys)
    aspect = (w_box / h_box) if h_box else 0.0
    for line in lines:
        x1, y1, x2, y2, line_len, ang_line = line
        if line_len <= long_len * 1.2:
            continue
        dist = _dist_point_to_line(cx, cy, line)
        ang_diff = _angle_diff_deg(ang_line, ang_token)
        if ang_diff < 15.0 and dist < max(short_len * 0.6, 3.0):
            return True
        # Single glyphs that sit on a long line with tiny distance
        if (
            text_single
            and dist < 2.0
            and ang_diff < 25.0
            and aspect < 0.15
            and confidence < 0.65
        ):
            return True
    return False


def _rerun_low_conf_token(
    rdr_main: _Any,
    langs: _List[str],
    arr_rgb: _np.ndarray,
    box: _Sequence[_Sequence[float]],
    text_orig: str,
    conf_orig: float,
    scale: float = 2.0,
) -> _Tuple[str, float, _Dict[str, _Any] | None]:
    if arr_rgb is None or arr_rgb.size == 0 or rdr_main is None:
        return text_orig, conf_orig, None
    try:
        debug = False
        try:
            debug = (_os.environ.get("OCR_DEBUG_BACKUP") or "").strip().lower() in ("1", "true", "yes")
        except Exception:
            debug = False
        try:
            use_tess = (_os.environ.get("OCR_ENABLE_TESSERACT") or "").strip().lower() in ("1", "true", "yes")
        except Exception:
            use_tess = False
        best_text = text_orig
        best_conf = conf_orig
        xs = [float(pt[0]) for pt in box]
        ys = [float(pt[1]) for pt in box]
        # Slightly larger padding, with extra vertical context for tall glyphs
        pad_x = 12.0
        pad_y = 16.0
        x0 = max(0, int(min(xs) - pad_x))
        y0 = max(0, int(min(ys) - pad_y))
        x1 = min(arr_rgb.shape[1], int(max(xs) + pad_x))
        y1 = min(arr_rgb.shape[0], int(max(ys) + pad_y))
        crop_base = arr_rgb[y0:y1, x0:x1]
        if crop_base.size == 0:
            return text_orig, conf_orig, None
        # Build a small family of scale factors so we effectively vary the
        # "local DPI" of this token region.
        scales: _List[float] = [1.0]
        if scale > 1.0:
            scales.append(scale)
        scales = [s for s in scales if s > 0.0]
        # For each scaled crop, try both raw and lightly sharpened versions,
        # and vary EasyOCR's internal contrast handling a bit. All retries
        # stay within the primary EasyOCR reader.
        cand_log: _List[Tuple[str, float, str]] = []
        for sc in scales:
            if sc == 1.0:
                crop = crop_base
            else:
                try:
                    crop = cv2.resize(crop_base, None, fx=sc, fy=sc, interpolation=cv2.INTER_CUBIC)
                except Exception:
                    crop = crop_base
            if crop.size == 0:
                continue
            # Variant 1: raw RGB crop
            variants: _List[Tuple[_np.ndarray, float, float, str]] = [
                (crop, 0.10, 0.3, "beamsearch"),
            ]
            # Variant 2: equalized + unsharp, if possible
            try:
                gray = cv2.cvtColor(crop, cv2.COLOR_RGB2GRAY)
                eq = cv2.equalizeHist(gray)
                blur = cv2.GaussianBlur(eq, (3, 3), 0)
                sharp = cv2.addWeighted(eq, 1.5, blur, -0.5, 0)
                proc_sharp = cv2.cvtColor(sharp, cv2.COLOR_GRAY2RGB)
                variants.append((proc_sharp, 0.05, 0.7, "beamsearch"))
            except Exception:
                pass
            for img, contrast_ths, adjust_contrast, decoder in variants:
                try:
                    res_retry = rdr_main.readtext(
                        img,
                        detail=1,
                        contrast_ths=contrast_ths,
                        adjust_contrast=adjust_contrast,
                        decoder=decoder,
                    )  # type: ignore[attr-defined]
                    if not res_retry:
                        continue
                    best = max(
                        res_retry,
                        key=lambda r: (r[2] if len(r) > 2 and r[2] is not None else 0.0),
                    )
                    text = best[1].strip() if len(best) > 1 and isinstance(best[1], str) else ""
                    conf = float(best[2]) if len(best) > 2 and best[2] is not None else 0.0
                except Exception:
                    continue
                if not text:
                    continue
                tag = f"scale={sc:.2f},ct={contrast_ths},ac={adjust_contrast}"
                cand_log.append((text, conf, tag))
                # Only accept if clearly better, not wildly different in length,
                # and does not reduce the character count relative to the
                # original token.
                if (
                    conf > best_conf + 0.05
                    and abs(len(text) - len(text_orig)) <= 2
                    and len(text) >= len(text_orig)
                ):
                    best_text = text
                    best_conf = conf
        best_easy_text = best_text
        best_easy_conf = best_conf
        tess_text = ""
        tess_conf = 0.0
        if use_tess and crop_base.size > 0:
            # Let Tesseract see the same padded crop; we do not restrict
            # characters here so that it can help with both digits and words.
            tess_text, tess_conf = _tesseract_ocr_crop(crop_base, allowlist=None, debug=debug)
            if tess_text and tess_conf > best_conf + 0.02:
                best_text = tess_text
                best_conf = tess_conf
        if debug and (best_text != text_orig):
            try:
                _sys.stderr.write(
                    f"[OCR BACKUP] orig={text_orig!r} conf={conf_orig:.3f} -> best={best_text!r} conf={best_conf:.3f}; "
                    f"candidates={cand_log}\n"
                )
            except Exception:
                pass
        debug_info: _Dict[str, _Any] | None = {
            "text_orig": text_orig,
            "conf_orig": conf_orig,
            "easy_text": best_easy_text,
            "easy_conf": best_easy_conf,
            "tess_text": tess_text,
            "tess_conf": tess_conf,
            "final_text": best_text,
            "final_conf": best_conf,
            "used_tess": bool(tess_text and tess_conf >= best_easy_conf),
        }
        return best_text, best_conf, debug_info
    except Exception:
        return text_orig, conf_orig, None


def _reconstruct_token_from_glyphs(
    rdr_main: _Any,
    arr_rgb: _np.ndarray,
    box: _Sequence[_Sequence[float]],
    text_orig: str,
    conf_orig: float,
    scale: float = 2.0,
) -> _Tuple[str, float]:
    """Attempt to re-read a low-confidence token by segmenting into glyphs.

    This stays within the "pure OCR" phase: we do not infer semantics or
    swap glyphs based on context, we simply try to give EasyOCR a cleaner,
    per-character view of the same pixels and aggregate the per-glyph scores.
    """
    if cv2 is None or rdr_main is None or arr_rgb is None or arr_rgb.size == 0:
        return text_orig, conf_orig
    try:
        debug = False
        try:
            debug = (_os.environ.get("OCR_DEBUG_GLYPH") or "").strip().lower() in ("1", "true", "yes")
        except Exception:
            debug = False
        # Restrict glyph-level characters to a conservative set derived from the
        # original token so that we do not “invent” obviously unrelated symbols.
        # Debug-only glyph inspection: at this stage we do not try to
        # fix or filter the token, only to blow up likely character
        # strips and see what OCR returns for each.
        xs = [float(pt[0]) for pt in box]
        ys = [float(pt[1]) for pt in box]
        pad_x = 12.0
        pad_y = 16.0
        x0 = max(0, int(min(xs) - pad_x))
        y0 = max(0, int(min(ys) - pad_y))
        x1 = min(arr_rgb.shape[1], int(max(xs) + pad_x))
        y1 = min(arr_rgb.shape[0], int(max(ys) + pad_y))
        crop = arr_rgb[y0:y1, x0:x1]
        if crop.size == 0:
            return text_orig, conf_orig
        try:
            gray = cv2.cvtColor(crop, cv2.COLOR_RGB2GRAY)
        except Exception:
            return text_orig, conf_orig
        # Simple binarization + per-column projection to isolate glyphs.
        try:
            _, bw = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        except Exception:
            return text_orig, conf_orig
        inv = 255 - bw
        h, w = inv.shape
        # Use vertical whitespace to segment into character-like strips so that
        # adjacent digits/symbols cannot interfere with each other.
        try:
            col_nonzero = (inv > 0).sum(axis=0)
        except Exception:
            return text_orig, conf_orig
        fg_thresh = max(1, int(0.01 * h))
        is_gap = col_nonzero <= fg_thresh
        segments: _List[_Tuple[int, int]] = []
        in_seg = False
        seg_start = 0
        for x in range(w):
            if not in_seg and not is_gap[x]:
                in_seg = True
                seg_start = x
            elif in_seg and is_gap[x]:
                seg_end = x
                if seg_end - seg_start >= 2:
                    segments.append((seg_start, seg_end))
                in_seg = False
        if in_seg:
            seg_end = w
            if seg_end - seg_start >= 2:
                segments.append((seg_start, seg_end))
        if not segments:
            return text_orig, conf_orig
        # If we detected fewer segments than characters, fall back to
        # simple equal-width slices so that each character in the token
        # gets its own strip, even when there is little internal whitespace.
        if len(segments) < len(text_orig):
            segments = []
            n = max(1, len(text_orig))
            for i in range(n):
                sx = int(i * w / n)
                ex = int((i + 1) * w / n)
                if ex - sx >= 2:
                    segments.append((sx, ex))
            if not segments:
                return text_orig, conf_orig
        glyph_chars: _List[str] = []
        glyph_confs: _List[float] = []
        for g_idx, (sx, ex) in enumerate(segments):
            # Widen each strip horizontally so that EasyOCR sees more
            # context around each nominal character slot. This uses an
            # overlapping window whose half-width is at least one-Nth
            # of the token box, where N ~= len(text_orig).
            if len(text_orig) > 0:
                target_half = max((ex - sx) // 2, w // (2 * len(text_orig)))
            else:
                target_half = (ex - sx) // 2
            cx_strip = (sx + ex) // 2
            gx0 = max(0, int(cx_strip - target_half))
            gx1 = min(w, int(cx_strip + target_half))
            if gx1 - gx0 < 2:
                continue
            # Derive vertical extent from foreground pixels within this widened strip.
            col_slice = inv[:, gx0:gx1]
            ys_fg = _np.where(col_slice > 0)[0]
            if ys_fg.size == 0:
                continue
            gy0 = max(0, int(ys_fg.min()) - 1)
            gy1 = min(h, int(ys_fg.max()) + 2)
            g_crop = crop[gy0:gy1, gx0:gx1]
            if g_crop.size == 0:
                continue
            if scale > 1.0:
                try:
                    g_crop = cv2.resize(g_crop, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
                except Exception:
                    pass
            # Let EasyOCR handle its own preprocessing; we only provide
            # a cleaner per-strip RGB crop.
            proc = g_crop
            try:
                res_g = rdr_main.readtext(
                    proc,
                    detail=1,
                )  # type: ignore[attr-defined]
            except Exception:
                res_g = []
            orig_ch = text_orig[g_idx] if g_idx < len(text_orig) else ""
            if not res_g:
                if debug:
                    try:
                        _sys.stderr.write(
                            f"[OCR GLYPH CHAR] token={text_orig!r} idx={g_idx} "
                            f"orig_char={orig_ch!r} raw_text=None conf=0.000 "
                            f"bbox=({gx0},{gy0},{gx1},{gy1}) reason=no_read\n"
                        )
                    except Exception:
                        pass
                glyph_chars.append("")
                glyph_confs.append(0.0)
                continue
            best = max(res_g, key=lambda r: (r[2] if len(r) > 2 and r[2] is not None else 0.0))
            txt_g = best[1].strip() if len(best) > 1 and isinstance(best[1], str) else ""
            conf_g = float(best[2]) if len(best) > 2 and best[2] is not None else 0.0
            glyph_ch = txt_g[0] if txt_g else ""
            if debug:
                try:
                    _sys.stderr.write(
                        f"[OCR GLYPH CHAR] token={text_orig!r} idx={g_idx} "
                        f"orig_char={orig_ch!r} raw_text={txt_g!r} ch={glyph_ch!r} "
                        f"conf={conf_g:.3f} bbox=({gx0},{gy0},{gx1},{gy1})\n"
                    )
                except Exception:
                    pass
            glyph_chars.append(glyph_ch)
            glyph_confs.append(conf_g)
        if debug:
            try:
                recon = "".join(ch or "?" for ch in glyph_chars) if glyph_chars else ""
                _sys.stderr.write(
                    f"[OCR GLYPH] orig={text_orig!r} conf={conf_orig:.3f} -> "
                    f"new={recon!r} conf={conf_orig:.3f} accepted=False "
                    f"chars={glyph_chars} glyph_confs={[round(c, 3) for c in glyph_confs]} glyphs={len(glyph_chars)}\n"
                )
            except Exception:
                pass
        return text_orig, conf_orig
    except Exception:
        return text_orig, conf_orig


def _refine_backslash_value(
    items: _List[_Dict[str, float]],
    arr_rgb: _np.ndarray,
) -> None:
    if arr_rgb is None or arr_rgb.size == 0:
        return
    # Find the Backslash Noise label
    labels = [it for it in items if str(it.get("text", "")).strip().lower() == "backslash noise"]
    if not labels:
        return
    row_tol = 25.0
    for lab in labels:
        cy = float(lab.get("cy", 0.0))
        cx = float(lab.get("cx", 0.0))
        # Candidate values: same row, to the right of the label
        cand_vals = [
            it for it in items
            if it is not lab
            and float(it.get("cx", 0.0)) > cx
            and abs(float(it.get("cy", 0.0)) - cy) <= row_tol
        ]
        if not cand_vals:
            continue
        cand_vals.sort(key=lambda d: float(d.get("cx", 0.0)))
        val = cand_vals[0]
        base_conf = float(val.get("conf", 0.0))
        # Only retry clearly shaky reads
        if base_conf >= 0.6:
            continue
        # Segment characters within the value bbox
        try:
            xs = [float(val.get("x0", 0.0)), float(val.get("x1", 0.0))]
            ys = [float(val.get("y0", 0.0)), float(val.get("y1", 0.0))]
            pad_x = 12.0
            pad_y = 16.0
            x0 = max(0, int(min(xs) - pad_x))
            y0 = max(0, int(min(ys) - pad_y))
            x1 = min(arr_rgb.shape[1], int(max(xs) + pad_x))
            y1 = min(arr_rgb.shape[0], int(max(ys) + pad_y))
            crop = arr_rgb[y0:y1, x0:x1]
            if crop.size == 0:
                continue
            # Binarize and find connected components
            gray = cv2.cvtColor(crop, cv2.COLOR_RGB2GRAY)
            _, bw = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            inv = 255 - bw
            num_labels, _, stats, _ = cv2.connectedComponentsWithStats(inv, connectivity=8)
            glyphs: _List[_Tuple[float, float, float, float]] = []
            h, w = inv.shape
            for idx in range(1, num_labels):
                x = stats[idx, cv2.CC_STAT_LEFT]
                y = stats[idx, cv2.CC_STAT_TOP]
                w_cc = stats[idx, cv2.CC_STAT_WIDTH]
                h_cc = stats[idx, cv2.CC_STAT_HEIGHT]
                area = stats[idx, cv2.CC_STAT_AREA]
                if w_cc <= 1 or h_cc <= 5:
                    continue
                if area < max(20, (h * w) * 0.002):
                    continue
                glyphs.append((float(x), float(y), float(w_cc), float(h_cc)))
            if not glyphs:
                continue
            glyphs.sort(key=lambda g: g[0])
            chars: _List[str] = []
            confs: _List[float] = []
            for gx, gy, gw, gh in glyphs:
                gx0 = int(max(0, gx - 2))
                gy0 = int(max(0, gy - 2))
                gx1 = int(min(w, gx + gw + 2))
                gy1 = int(min(h, gy + gh + 2))
                g_crop = crop[gy0:gy1, gx0:gx1]
                if g_crop.size == 0:
                    continue
                # Upscale small glyphs
                try:
                    g_crop = cv2.resize(g_crop, None, fx=2.0, fy=2.0, interpolation=cv2.INTER_CUBIC)
                except Exception:
                    pass
                try:
                    alt_rdr = _get_alt_easyocr_reader(["en"])
                except Exception:
                    alt_rdr = None
                engine = alt_rdr
                if engine is None:
                    engine = None
                try:
                    if engine is not None:
                        res_g = engine.readtext(
                            g_crop,
                            detail=1,
                            contrast_ths=0.05,
                            adjust_contrast=0.7,
                            decoder="beamsearch",
                        )  # type: ignore[attr-defined]
                    else:
                        res_g = []
                except Exception:
                    res_g = []
                if not res_g:
                    continue
                best = max(res_g, key=lambda r: (r[2] if len(r) > 2 and r[2] is not None else 0.0))
                txt_g = best[1].strip() if len(best) > 1 and isinstance(best[1], str) else ""
                conf_g = float(best[2]) if len(best) > 2 and best[2] is not None else 0.0
                if not txt_g:
                    continue
                # Only keep the first reasonable character from this glyph region
                ch = txt_g[0]
                if ch not in "0123456789/\\-":
                    continue
                chars.append(ch)
                confs.append(conf_g)
            if not chars:
                continue
            new_text = "".join(chars)
            new_conf = min(confs) if confs else base_conf
            if new_conf > base_conf and len(new_text) >= 2:
                val["text"] = new_text
                val["conf"] = new_conf
        except Exception:
            continue


def _get_easyocr_boxes_page(pdf_path: _Path, page: int, dpi: int, langs: _List[str]) -> _List[_Dict[str, float]]:
    """Line-geometry solver: detect long lines, erase them, OCR cleaned image, drop residual line artifacts."""
    if not (_orig._HAVE_EASYOCR and _orig._HAVE_PYMUPDF):  # type: ignore[attr-defined]
        return []
    rdr = _orig._get_easyocr_reader(langs)  # type: ignore[attr-defined]
    if rdr is None or fitz is None:
        return []
    try:
        doc = fitz.open(str(pdf_path))  # type: ignore[attr-defined]
    except Exception:
        return []
    try:
        dpi_eff = _clamp_ocr_dpi(int(dpi))
        if not (1 <= page <= doc.page_count):
            return []
        try:
            pg = doc.load_page(page - 1)
            pix = pg.get_pixmap(dpi=dpi_eff)
        except Exception:
            return []
        try:
            from PIL import Image as _Image  # type: ignore
            mode = "RGB" if pix.n >= 3 else "L"
            img = _Image.frombytes(mode, [pix.width, pix.height], pix.samples)
            if mode != "RGB":
                img = img.convert("RGB")
            arr = _np.array(img)
        except Exception:
            return []

        lines = _detect_lines(arr)
        arr_clean = arr.copy()
        if cv2 is not None and lines:
            try:
                mask = _np.zeros(arr_clean.shape[:2], dtype=_np.uint8)
                h, w = mask.shape
                thickness = max(1, min(h, w) // 600)
                for x1, y1, x2, y2, _, _ in lines:
                    cv2.line(mask, (int(x1), int(y1)), (int(x2), int(y2)), color=255, thickness=thickness)
                if thickness > 1:
                    kernel = cv2.getStructuringElement(
                        cv2.MORPH_RECT,
                        (max(1, thickness // 2), max(1, thickness // 2)),
                    )
                    mask = cv2.erode(mask, kernel, iterations=1)
                try:
                    arr_clean = cv2.inpaint(arr_clean, mask, 3, cv2.INPAINT_TELEA)
                except Exception:
                    arr_clean[mask > 0] = 255
            except Exception:
                pass
        try:
            res = rdr.readtext(arr_clean, detail=1)  # type: ignore[attr-defined]
        except Exception:
            res = []
        items: _List[_Dict[str, float]] = []
        use_backup = _backup_enabled()
        for it in res:
            try:
                box, text, conf = it
                if not isinstance(text, str):
                    continue
                t = text.strip()
                if not t:
                    continue
                if _is_line_artifact(box, t, lines, confidence=float(conf) if conf is not None else 0.0):
                    continue
                cval = float(conf) if conf is not None else 0.0
                debug_backup: _Dict[str, _Any] | None = None
                # Optional: as a final resort for very low-confidence tokens,
                # allow a single fallback pass using the primary EasyOCR
                # reader with more aggressive settings. This is disabled
                # unless OCR_ENABLE_BACKUP is set.
                if use_backup and cval < _BACKUP_CONF_THRESHOLD:
                    t, cval, debug_backup = _rerun_low_conf_token(rdr, langs, arr_clean, box, t, cval)
                xs = [float(pt[0]) for pt in box]
                ys = [float(pt[1]) for pt in box]
                x0, y0, x1, y1 = min(xs), min(ys), max(xs), max(ys)
                cx = (x0 + x1) / 2.0
                cy = (y0 + y1) / 2.0
                item: _Dict[str, _Any] = {
                    "x0": x0, "y0": y0, "x1": x1, "y1": y1, "cx": cx, "cy": cy,
                    "text": t, "conf": cval
                }
                if debug_backup is not None:
                    item["debug_backup"] = debug_backup
                items.append(item)  # type: ignore[list-item]
            except Exception:
                pass
        return items
    finally:
        try:
            doc.close()
        except Exception:
            pass


_orig._get_easyocr_boxes_page = _get_easyocr_boxes_page  # type: ignore[attr-defined]
