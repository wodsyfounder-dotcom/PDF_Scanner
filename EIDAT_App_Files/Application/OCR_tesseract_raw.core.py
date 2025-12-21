#!/usr/bin/env python3
"""
OCR_tesseract_raw.core

Clean, "first principles" Tesseract OCR baseline for scanned PDFs.

Goal:
- Render PDF pages to images (PyMuPDF).
- Run Tesseract directly on the rendered page image (no line removal, no masking).
- Consume native Tesseract structure via TSV output (word boxes + line ids).
- Build a structured view: lines -> positions -> words.

No image preprocessing is performed in this module.
"""
from __future__ import annotations

import json
import os
import subprocess
from pathlib import Path
from statistics import median
from typing import Any, Dict, List, Sequence, Tuple

try:
    import fitz  # type: ignore
except Exception:
    fitz = None  # type: ignore


DEFAULT_STRESS_PDF = Path(
    r"C:\Users\zachs\Documents\DevProjects\PDF_Scanner\Data Packages\ocr_stress_case.pdf"
)


def _which_tesseract() -> str | None:
    try:
        import shutil

        return shutil.which("tesseract")
    except Exception:
        return None


def _render_page_png_bytes(pdf_path: Path, page_1_based: int, dpi: int) -> bytes:
    if fitz is None:
        raise RuntimeError("PyMuPDF (fitz) is not available")
    doc = fitz.open(str(pdf_path))  # type: ignore[attr-defined]
    try:
        if not (1 <= page_1_based <= doc.page_count):
            raise ValueError(f"Page out of range: {page_1_based} (1..{doc.page_count})")
        page = doc.load_page(page_1_based - 1)
        pix = page.get_pixmap(dpi=int(dpi))
        # PNG bytes for Tesseract stdin
        return pix.tobytes("png")
    finally:
        try:
            doc.close()
        except Exception:
            pass


def _parse_tesseract_tsv(tsv_text: str) -> List[Dict[str, Any]]:
    """Parse Tesseract TSV output into rows (dicts) for word-level entries."""
    rows: List[Dict[str, Any]] = []
    if not tsv_text:
        return rows
    lines = tsv_text.splitlines()
    if not lines:
        return rows
    # Header: level page_num block_num par_num line_num word_num left top width height conf text
    for ln in lines[1:]:
        parts = ln.split("\t")
        if len(parts) < 12:
            continue
        try:
            level = int(parts[0])
        except Exception:
            continue
        if level != 5:
            continue
        text = (parts[11] or "").strip()
        if not text:
            continue
        try:
            conf_raw = float(parts[10])
        except Exception:
            conf_raw = -1.0
        # Tesseract uses -1 for "not a real confidence" sometimes.
        conf = max(0.0, min(1.0, conf_raw / 100.0)) if conf_raw >= 0.0 else 0.0
        try:
            left = int(parts[6])
            top = int(parts[7])
            width = int(parts[8])
            height = int(parts[9])
        except Exception:
            continue
        try:
            block_num = int(parts[2])
            par_num = int(parts[3])
            line_num = int(parts[4])
            word_num = int(parts[5])
        except Exception:
            continue
        rows.append(
            {
                "block_num": block_num,
                "par_num": par_num,
                "line_num": line_num,
                "word_num": word_num,
                "left": left,
                "top": top,
                "width": width,
                "height": height,
                "x0": left,
                "y0": top,
                "x1": left + width,
                "y1": top + height,
                "conf": conf,
                "text": text,
            }
        )
    return rows


def _group_words_into_lines(words: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Group word rows into lines using Tesseract's own (block,par,line) ids."""
    buckets: Dict[Tuple[int, int, int], List[Dict[str, Any]]] = {}
    for w in words:
        key = (int(w["block_num"]), int(w["par_num"]), int(w["line_num"]))
        buckets.setdefault(key, []).append(w)
    lines: List[Dict[str, Any]] = []
    for (block_num, par_num, line_num), ws in buckets.items():
        ws_sorted = sorted(ws, key=lambda r: (int(r["top"]), int(r["left"]), int(r["word_num"])))
        y0 = min(int(w["y0"]) for w in ws_sorted)
        y1 = max(int(w["y1"]) for w in ws_sorted)
        x0 = min(int(w["x0"]) for w in ws_sorted)
        x1 = max(int(w["x1"]) for w in ws_sorted)
        cy = (y0 + y1) / 2.0
        lines.append(
            {
                "block_num": block_num,
                "par_num": par_num,
                "line_num": line_num,
                "bbox": {"x0": x0, "y0": y0, "x1": x1, "y1": y1, "cy": cy},
                "words": ws_sorted,
            }
        )
    # Stable reading order: top-to-bottom, then left-to-right.
    lines.sort(key=lambda d: (float(d["bbox"]["cy"]), int(d["bbox"]["x0"])))
    return lines


def _group_line_positions(
    line_words: List[Dict[str, Any]],
    *,
    gap_mult: float,
) -> List[Dict[str, Any]]:
    """Group a line's words into "positions" based on x-gaps between word boxes."""
    if not line_words:
        return []
    words = sorted(line_words, key=lambda w: int(w["x0"]))
    heights = [max(1, int(w["height"])) for w in words]
    h_med = float(median(heights)) if heights else 10.0
    gap_thresh = max(1.0, gap_mult * h_med)
    positions: List[List[Dict[str, Any]]] = []
    cur: List[Dict[str, Any]] = [words[0]]
    for prev, nxt in zip(words, words[1:]):
        gap = float(int(nxt["x0"]) - int(prev["x1"]))
        if gap >= gap_thresh:
            positions.append(cur)
            cur = [nxt]
        else:
            cur.append(nxt)
    positions.append(cur)
    out: List[Dict[str, Any]] = []
    for idx, group in enumerate(positions, 1):
        out.append(
            {
                "pos": idx,
                "text": " ".join(str(w["text"]) for w in group),
                "words": [
                    {
                        "text": w["text"],
                        "conf": float(w["conf"]),
                        "bbox": {"x0": int(w["x0"]), "y0": int(w["y0"]), "x1": int(w["x1"]), "y1": int(w["y1"])},
                    }
                    for w in group
                ],
            }
        )
    return out


def tesseract_ocr_structured(
    pdf_path: Path,
    pages: Sequence[int],
    *,
    dpi: int = 800,
    lang: str = "eng",
    psm: int = 6,
    tess_dpi: int | None = None,
    gap_mult: float = 1.5,
) -> Dict[str, Any]:
    pdf_path = Path(pdf_path).expanduser().resolve()
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    tess = _which_tesseract()
    if not tess:
        raise RuntimeError("tesseract binary not found on PATH")

    out_pages: Dict[str, Any] = {}
    for p in pages:
        png_bytes = _render_page_png_bytes(pdf_path, int(p), int(dpi))
        use_tess_dpi = int(tess_dpi) if tess_dpi is not None else int(dpi)
        cmd = [
            tess,
            "stdin",
            "stdout",
            "-l",
            str(lang),
            "--psm",
            str(int(psm)),
            "--oem",
            "3",
            "--dpi",
            str(use_tess_dpi),
            "tsv",
        ]
        proc = subprocess.run(
            cmd,
            input=png_bytes,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=False,
        )
        tsv = ""
        try:
            tsv = proc.stdout.decode("utf-8", errors="replace")
        except Exception:
            tsv = ""

        words = _parse_tesseract_tsv(tsv)
        line_groups = _group_words_into_lines(words)
        structured_lines: List[Dict[str, Any]] = []
        for idx, ln in enumerate(line_groups, 1):
            positions = _group_line_positions(ln["words"], gap_mult=float(gap_mult))
            structured_lines.append(
                {
                    "line_index": idx,
                    "tess_line": {
                        "block_num": ln["block_num"],
                        "par_num": ln["par_num"],
                        "line_num": ln["line_num"],
                    },
                    "bbox": ln["bbox"],
                    "positions": positions,
                }
            )
        out_pages[str(int(p))] = {
            "page": int(p),
            "tsv_word_count": len(words),
            "lines": structured_lines,
            "tesseract": {
                "rc": int(proc.returncode),
                "cmd": cmd,
                "stderr": (proc.stderr.decode("utf-8", errors="replace") if proc.stderr else ""),
            },
        }

    return {
        "pdf_path": str(pdf_path),
        "dpi": int(dpi),
        "lang": str(lang),
        "psm": int(psm),
        "gap_mult": float(gap_mult),
        "pages": out_pages,
    }


def main() -> int:
    pdf_path = Path(os.environ.get("OCR_PDF", str(DEFAULT_STRESS_PDF)))
    try:
        pages_raw = (os.environ.get("OCR_PAGES") or "1").strip()
        pages = [int(x.strip()) for x in pages_raw.split(",") if x.strip()]
    except Exception:
        pages = [1]
    try:
        dpi = int((os.environ.get("OCR_DPI") or "800").strip())
    except Exception:
        dpi = 800
    try:
        psm = int((os.environ.get("OCR_TESS_PSM") or "6").strip())
    except Exception:
        psm = 6
    try:
        gap_mult = float((os.environ.get("OCR_POS_GAP_MULT") or "1.5").strip())
    except Exception:
        gap_mult = 1.5

    obj = tesseract_ocr_structured(pdf_path, pages, dpi=dpi, psm=psm, gap_mult=gap_mult)
    print(json.dumps(obj, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
