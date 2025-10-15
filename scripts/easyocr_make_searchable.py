#!/usr/bin/env python3
"""
Make a searchable PDF by overlaying EasyOCR text on top of the original pages.

Behavior
- Renders each page with PyMuPDF at a chosen DPI
- Runs EasyOCR on the raster image (optionally resized/preprocessed)
- For each detected text line, inserts invisible text at the corresponding bbox
- Saves a visually-identical but text-searchable PDF

Usage (from project root):
  py -3.11 scripts/easyocr_make_searchable.py \
      --pdf user_inputs/EIDP_Import_Docs/sn\ 3232.pdf \
      --out user_inputs/EIDP_Import_Docs/sn\ 3232.easy.ocr.pdf \
      --dpi 800 --langs en --min-conf 0.3

Requirements
- PyMuPDF (pymupdf), easyocr (and its CPU torch deps)
"""
from __future__ import annotations

import argparse
import sys
import tempfile
from pathlib import Path
from typing import List, Tuple, Optional
import json

try:
    import fitz  # type: ignore  # PyMuPDF
except Exception:
    print("[ERROR] PyMuPDF (fitz) is required: pip install pymupdf", file=sys.stderr)
    sys.exit(1)

try:
    import easyocr  # type: ignore
except Exception:
    print("[ERROR] EasyOCR is required: pip install easyocr", file=sys.stderr)
    sys.exit(1)


def _resize_for_ocr(src_png: Path, dst_png: Path, max_side: int = 4000) -> Tuple[Path, int, int]:
    """Ensure max(image width,height) <= max_side. Returns (dst_path, width, height)."""
    try:
        import cv2  # type: ignore
        img = cv2.imread(str(src_png))
        if img is None:
            # Fallback: copy original and read size via PyMuPDF pixmap
            import shutil
            shutil.copyfile(str(src_png), str(dst_png))
            pm = fitz.Pixmap(str(src_png))  # type: ignore[name-defined]
            w, h = pm.width, pm.height
            pm = None  # type: ignore
            return dst_png, w, h
        h, w = img.shape[:2]
        m = max(w, h)
        if m > max_side and m > 0:
            scale = max_side / float(m)
            new_w = max(1, int(round(w * scale)))
            new_h = max(1, int(round(h * scale)))
            img = cv2.resize(img, (new_w, new_h), interpolation=cv2.INTER_AREA)
            w, h = new_w, new_h
        cv2.imwrite(str(dst_png), img)
        return dst_png, w, h
    except Exception:
        import shutil
        shutil.copyfile(str(src_png), str(dst_png))
        pm = fitz.Pixmap(str(src_png))  # type: ignore[name-defined]
        w, h = pm.width, pm.height
        pm = None  # type: ignore
        return dst_png, w, h


def _preprocess_contrast(src_png: Path, dst_png: Path) -> Path:
    """Light pre-processing (grayscale+CLAHE+adaptive threshold). Returns dst path."""
    try:
        import cv2  # type: ignore
        img = cv2.imread(str(src_png), cv2.IMREAD_GRAYSCALE)
        if img is None:
            raise RuntimeError("cv2.imread returned None")
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
        g = clahe.apply(img)
        thr = cv2.adaptiveThreshold(g, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 35, 11)
        cv2.imwrite(str(dst_png), thr)
        return dst_png
    except Exception:
        # On any failure, just return original
        return src_png


def _bbox_to_rect(bbox: List[List[float]]) -> Tuple[float, float, float, float]:
    # EasyOCR bbox: [[x1,y1],[x2,y2],[x3,y3],[x4,y4]]
    xs = [p[0] for p in bbox]
    ys = [p[1] for p in bbox]
    return min(xs), min(ys), max(xs), max(ys)


def make_searchable_pdf(
    in_pdf: Path,
    out_pdf: Path,
    dpi: int = 800,
    langs: List[str] | None = None,
    min_conf: float = 0.3,
    preprocess: bool = True,
    visible: bool = False,
    dump_dir: Optional[Path] = None,
    dump_json: bool = False,
) -> None:
    out_pdf.parent.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(str(in_pdf))  # type: ignore[name-defined]

    dpi = max(200, min(900, int(dpi)))
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)  # type: ignore[name-defined]

    reader = easyocr.Reader(langs or ["en"], gpu=False, verbose=False)  # type: ignore
    tmp_dir = Path(tempfile.mkdtemp(prefix="easyocr_pdf_"))
    if dump_dir is not None:
        dump_dir.mkdir(parents=True, exist_ok=True)
    combined_lines: List[str] = []
    combined_json: List[dict] = []

    try:
        for pno in range(doc.page_count):
            page = doc.load_page(pno)
            pix = page.get_pixmap(matrix=mat)
            raw_png = tmp_dir / f"p{pno+1}.png"
            pix.save(str(raw_png))

            # Preprocess and/or resize for OCR
            png_for_ocr = raw_png
            if preprocess:
                pre = tmp_dir / f"p{pno+1}.pre.png"
                png_for_ocr = _preprocess_contrast(raw_png, pre)
            fit_png = tmp_dir / f"p{pno+1}.fit.png"
            fit_png, used_w, used_h = _resize_for_ocr(png_for_ocr, fit_png, max_side=4000)

            # OCR
            try:
                results = reader.readtext(str(fit_png), detail=1)  # list of [bbox, text, conf]
            except Exception:
                results = []

            if not results:
                # Fallback: try raw image too
                try:
                    results = reader.readtext(str(raw_png), detail=1)
                    used_w, used_h = pix.width, pix.height
                except Exception:
                    results = []

            if not results:
                continue

            # Scale factors (OCR image -> page points)
            sx = page.rect.width / float(used_w or 1)
            sy = page.rect.height / float(used_h or 1)

            # If debug-visible, drop a small stamp so we know overlay executed
            if visible:
                try:
                    page.insert_text(
                        fitz.Point(page.rect.x0 + 10, page.rect.y0 + 14),  # type: ignore[name-defined]
                        "DEBUG overlay",
                        fontname="helv",
                        fontsize=10,
                        color=(1, 0, 0),
                        render_mode=0,
                        overlay=True,
                    )
                except Exception:
                    pass

            # Collect per-page text for optional dumps
            page_lines: List[str] = []
            page_json: List[dict] = []

            for item in results:
                try:
                    bbox, text, conf = item
                except Exception:
                    continue
                try:
                    conf_val = float(conf)
                except Exception:
                    conf_val = 0.0
                if conf_val < float(min_conf):
                    continue
                if not isinstance(text, str) or not text.strip():
                    continue

                try:
                    x0, y0, x1, y1 = _bbox_to_rect(bbox)
                except Exception:
                    continue

                rx0, ry0 = x0 * sx, y0 * sy
                rx1, ry1 = x1 * sx, y1 * sy
                rect = fitz.Rect(rx0, ry0, rx1, ry1)  # type: ignore[name-defined]
                # Increase minimum font size to avoid invisible glyphs
                fontsize = max(12.0, min(72.0, rect.height * 0.9))
                text_color = (1, 0, 0) if visible else (0, 0, 0)
                rmode = 0 if visible else 3
                page_lines.append(text)
                # Cast to JSON-serializable built-in types (avoid numpy dtypes)
                x0p, y0p, x1p, y1p = float(x0), float(y0), float(x1), float(y1)
                page_json.append({
                    "page": int(pno + 1),
                    "bbox": [x0p, y0p, x1p, y1p],
                    "text": str(text),
                    "conf": float(conf_val),
                })

                try:
                    ret = page.insert_textbox(
                        rect,
                        text,
                        fontname="helv",
                        fontsize=fontsize,
                        color=text_color,
                        align=0,
                        render_mode=rmode,
                        overlay=True,
                    )
                    if not ret:
                        page.insert_text(
                            fitz.Point(rect.x0, rect.y0),  # type: ignore[name-defined]
                            text,
                            fontname="helv",
                            fontsize=fontsize,
                            color=text_color,
                            render_mode=rmode,
                            overlay=True,
                        )
                except Exception:
                    try:
                        page.insert_text(
                            fitz.Point(rect.x0, rect.y0),  # type: ignore[name-defined]
                            text,
                            fontname="helv",
                            fontsize=fontsize,
                            color=text_color,
                            render_mode=rmode,
                            overlay=True,
                        )
                    except Exception:
                        pass

                # Draw rectangle outlines in visible mode to visualize OCR boxes
                if visible:
                    try:
                        page.draw_rect(rect, color=(1, 0, 0), width=0.7, overlay=True)
                    except Exception:
                        pass

            # Optional dumps per page
            if dump_dir is not None:
                (dump_dir / f"{in_pdf.stem}_page_{pno+1}.txt").write_text("\n".join(page_lines), encoding="utf-8")
                if dump_json:
                    (dump_dir / f"{in_pdf.stem}_page_{pno+1}.json").write_text(json.dumps(page_json, ensure_ascii=False, indent=2), encoding="utf-8")
            combined_lines.extend(page_lines)
            if dump_json:
                combined_json.extend(page_json)

        doc.save(str(out_pdf), garbage=4, deflate=True)
        print(f"[DONE] Wrote searchable PDF -> {out_pdf}")
        if dump_dir is not None:
            (dump_dir / f"{in_pdf.stem}.txt").write_text("\n".join(combined_lines), encoding="utf-8")
            if dump_json:
                (dump_dir / f"{in_pdf.stem}.json").write_text(json.dumps(combined_json, ensure_ascii=False, indent=2), encoding="utf-8")
    finally:
        try:
            doc.close()
        except Exception:
            pass
        try:
            import shutil
            shutil.rmtree(str(tmp_dir), ignore_errors=True)
        except Exception:
            pass


def main() -> None:
    ap = argparse.ArgumentParser(description="Make PDF searchable via EasyOCR overlays")
    ap.add_argument("--pdf", required=True, help="Input PDF path")
    ap.add_argument("--out", required=True, help="Output searchable PDF path")
    ap.add_argument("--dpi", type=int, default=800, help="Render DPI (200..900, default 800)")
    ap.add_argument("--langs", default="en", help="Comma-separated language codes (default: en)")
    ap.add_argument("--min-conf", type=float, default=0.3, help="Min confidence (0..1)")
    ap.add_argument("--no-preprocess", action="store_true", help="Disable image pre-processing")
    ap.add_argument("--visible", action="store_true", help="Render overlays visibly (debug)")
    ap.add_argument("--dump-dir", help="Directory to dump recognized text per page and combined")
    ap.add_argument("--dump-json", action="store_true", help="Also dump OCR results as JSON")
    args = ap.parse_args()

    in_pdf = Path(args.pdf)
    out_pdf = Path(args.out)
    if not in_pdf.exists():
        print(f"[ERROR] PDF not found: {in_pdf}", file=sys.stderr)
        sys.exit(2)

    langs = [s.strip() for s in args.langs.replace(";", ",").split(",") if s.strip()]
    dump_dir = Path(args.dump_dir) if args.dump_dir else None
    make_searchable_pdf(
        in_pdf,
        out_pdf,
        dpi=args.dpi,
        langs=langs or ["en"],
        min_conf=args.min_conf,
        preprocess=(not args.no_preprocess),
        visible=bool(args.visible),
        dump_dir=dump_dir,
        dump_json=bool(args.dump_json),
    )


if __name__ == "__main__":
    main()
