#!/usr/bin/env python3
"""
Make a searchable PDF by overlaying PaddleOCR text on top of the original pages.

Behavior
- Opens the input PDF with PyMuPDF (fitz) and for each page:
  - Renders the page to an image at the requested DPI
  - Runs PaddleOCR on the image
  - For each detected text line, scales its bounding box back to page coordinates
  - Inserts the recognized text into that box with an invisible rendering mode
- Saves a new PDF that remains visually identical but is text-searchable.

Usage (from project root):
  py -3.11 scripts/paddle_ocr_make_searchable.py \
      --pdf user_inputs/EIDP_Import_Docs/sn\ 3232.pdf \
      --out user_inputs/EIDP_Import_Docs/sn\ 3232.ocr.pdf \
      --dpi 600 --lang en --min-conf 0.5

Requirements
- PyMuPDF (pymupdf)
- paddlepaddle, paddleocr

Notes
- We only add an invisible text layer; we do not try to reflow tables or layout.
- If a page already had some text, we still add overlays (they are invisible).
"""
from __future__ import annotations

import argparse
import math
import os
import sys
import tempfile
from pathlib import Path
from typing import Iterable, List, Tuple

try:
    import fitz  # type: ignore  # PyMuPDF
except Exception:
    print("[ERROR] PyMuPDF (fitz) is required: pip install pymupdf", file=sys.stderr)
    sys.exit(1)

try:
    from paddleocr import PaddleOCR  # type: ignore
except Exception:
    print("[ERROR] PaddleOCR is required: pip install paddlepaddle paddleocr", file=sys.stderr)
    sys.exit(1)


def _page_points_to_rect(quad: List[List[float]]) -> Tuple[float, float, float, float]:
    xs = [pt[0] for pt in quad]
    ys = [pt[1] for pt in quad]
    x0, y0 = min(xs), min(ys)
    x1, y1 = max(xs), max(ys)
    return x0, y0, x1, y1


def _resize_for_paddle(src_png: Path, dst_png: Path, max_side: int = 4000) -> Tuple[Path, int, int]:
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
        # No OpenCV — just copy and use original size
        import shutil
        shutil.copyfile(str(src_png), str(dst_png))
        pm = fitz.Pixmap(str(src_png))  # type: ignore[name-defined]
        w, h = pm.width, pm.height
        pm = None  # type: ignore
        return dst_png, w, h


def make_searchable_pdf(
    in_pdf: Path,
    out_pdf: Path,
    dpi: int = 600,
    lang: str = "en",
    min_conf: float = 0.5,
) -> None:
    out_pdf.parent.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(str(in_pdf))  # type: ignore[name-defined]
    dpi = max(200, min(800, int(dpi)))
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)  # type: ignore[name-defined]
    # Prefer textline orientation flag; fallback to older arg for older versions
    try:
        ocr = PaddleOCR(use_textline_orientation=True, lang=lang)  # type: ignore
    except Exception:
        ocr = PaddleOCR(use_angle_cls=True, lang=lang)  # type: ignore

    tmp_dir = Path(tempfile.mkdtemp(prefix="paddle_ocr_pdf_"))

    try:
        for pno in range(doc.page_count):
            page = doc.load_page(pno)
            # Render page to raster image for OCR
            pix = page.get_pixmap(matrix=mat)
            img_path = tmp_dir / f"p{pno+1}.png"
            pix.save(str(img_path))

            # Prepare image for Paddle: cap max side to 4000 so coordinates match our scaling
            proc_path = tmp_dir / f"p{pno+1}.fit.png"
            proc_path, used_w, used_h = _resize_for_paddle(img_path, proc_path, max_side=4000)

            # Run OCR using the stable API; if empty, try raw image too
            def _ocr_boxes(png: Path):
                try:
                    return ocr.ocr(str(png), cls=True)  # type: ignore[attr-defined]
                except Exception:
                    return None

            result = _ocr_boxes(proc_path)
            if not result:
                result = _ocr_boxes(img_path)

            if not result:
                continue

            # Scale factors (processed image -> page points)
            sx = page.rect.width / float(used_w or 1)
            sy = page.rect.height / float(used_h or 1)

            # Overlay recognized text as invisible text on the page
            for block in result:
                for item in (block or []):
                    try:
                        quad, (text, conf) = item
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
                        x0, y0, x1, y1 = _page_points_to_rect(quad)
                    except Exception:
                        continue

                    # Map image pixel coords -> page points
                    rx0, ry0 = x0 * sx, y0 * sy
                    rx1, ry1 = x1 * sx, y1 * sy
                    # Fit text inside this box; rendering mode 3 = invisible text
                    rect = fitz.Rect(rx0, ry0, rx1, ry1)  # type: ignore[name-defined]
                    # Font size as box height (points), clamped
                    fontsize = max(4.0, min(48.0, rect.height * 0.9))
                    try:
                        _ = page.insert_textbox(
                            rect,
                            text,
                            fontname="helv",
                            fontsize=fontsize,
                            color=(0, 0, 0),
                            align=0,
                            render_mode=3,  # invisible text layer
                        )
                    except Exception:
                        # Fallback: single-point insertion at top-left of rect
                        try:
                            page.insert_text(
                                fitz.Point(rect.x0, rect.y0),  # type: ignore[name-defined]
                                text,
                                fontname="helv",
                                fontsize=fontsize,
                                color=(0, 0, 0),
                                render_mode=3,
                            )
                        except Exception:
                            pass

        # Save new searchable PDF
        doc.save(str(out_pdf), garbage=4, deflate=True)
        print(f"[DONE] Wrote searchable PDF -> {out_pdf}")
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
    ap = argparse.ArgumentParser(description="Make PDF searchable via PaddleOCR overlays")
    ap.add_argument("--pdf", required=True, help="Path to input PDF")
    ap.add_argument("--out", required=True, help="Path to output searchable PDF")
    ap.add_argument("--dpi", type=int, default=600, help="Render DPI (200..800, default 600)")
    ap.add_argument("--lang", default="en", help="PaddleOCR language code (default en)")
    ap.add_argument("--min-conf", type=float, default=0.5, help="Min confidence for text (0..1)")
    args = ap.parse_args()

    in_pdf = Path(args.pdf)
    out_pdf = Path(args.out)
    if not in_pdf.exists():
        print(f"[ERROR] PDF not found: {in_pdf}", file=sys.stderr)
        sys.exit(2)

    make_searchable_pdf(in_pdf, out_pdf, dpi=args.dpi, lang=args.lang, min_conf=args.min_conf)


if __name__ == "__main__":
    main()
