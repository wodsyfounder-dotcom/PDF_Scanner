#!/usr/bin/env python3
"""
OCR a PDF with PaddleOCR (via PyMuPDF render) and write a combined text file.

Usage (from project root):
  py -3.11 scripts/paddle_ocr_dump.py --pdf user_inputs/EIDP_Import_Docs/sn\ 3232.pdf \
      --out Product_Data_File/paddle_text/sn_3232.txt --dpi 600

Notes:
- Requires: fitz (PyMuPDF), paddleocr, paddlepaddle installed in the Python used.
- Writes a single text file; optionally save per-page text files via --out-dir.
"""
from __future__ import annotations

import argparse
import sys
import os
import tempfile
from pathlib import Path
from typing import List, Sequence

try:
    import fitz  # type: ignore  # PyMuPDF
except Exception as e:
    print("[ERROR] PyMuPDF (fitz) is required: pip install pymupdf", file=sys.stderr)
    sys.exit(1)

try:
    from paddleocr import PaddleOCR  # type: ignore
except Exception as e:
    print("[ERROR] PaddleOCR is required: pip install paddlepaddle paddleocr", file=sys.stderr)
    sys.exit(1)


def parse_pages(pages_raw: str | None, max_page: int) -> List[int]:
    if not pages_raw:
        return list(range(1, max_page + 1))
    pages: set[int] = set()
    for part in pages_raw.replace(" ", "").split(","):
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            try:
                start = int(a)
                end = int(b)
            except ValueError:
                continue
            if start <= end:
                for p in range(start, end + 1):
                    if 1 <= p <= max_page:
                        pages.add(p)
        else:
            try:
                p = int(part)
                if 1 <= p <= max_page:
                    pages.add(p)
            except ValueError:
                continue
    return sorted(pages) or list(range(1, max_page + 1))


def _preprocess_image(src_png: Path, dst_png: Path) -> None:
    """Light pre-processing to help OCR: grayscale + CLAHE + adaptive threshold."""
    try:
        import cv2  # type: ignore
        img = cv2.imread(str(src_png), cv2.IMREAD_GRAYSCALE)
        if img is None:
            # Fallback: copy as-is
            import shutil
            shutil.copyfile(str(src_png), str(dst_png))
            return
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
        g = clahe.apply(img)
        thr = cv2.adaptiveThreshold(g, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 35, 11)
        cv2.imwrite(str(dst_png), thr)
    except Exception:
        # If OpenCV isn't available, fall back to original
        import shutil
        shutil.copyfile(str(src_png), str(dst_png))


def ocr_pdf_to_text(pdf_path: Path, out_txt: Path, dpi: int, lang: str, out_dir: Path | None = None, preprocess: bool = True) -> None:
    out_txt.parent.mkdir(parents=True, exist_ok=True)
    if out_dir is not None:
        out_dir.mkdir(parents=True, exist_ok=True)

    # Avoid deprecated/unsupported args like show_log in newer releases
    # Prefer textline orientation flag; fallback to older arg for compatibility
    try:
        ocr = PaddleOCR(use_textline_orientation=True, lang=lang)  # type: ignore
    except Exception:
        ocr = PaddleOCR(use_angle_cls=True, lang=lang)  # type: ignore
    doc = fitz.open(str(pdf_path))  # type: ignore[name-defined]
    pages = parse_pages(None, doc.page_count)

    # Clamp DPI
    dpi = max(200, min(800, int(dpi)))
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)  # type: ignore[name-defined]

    tmp_dir = Path(tempfile.mkdtemp(prefix="paddle_dump_"))
    combined: List[str] = []
    try:
        for pno in pages:
            try:
                page = doc.load_page(pno - 1)
                pix = page.get_pixmap(matrix=mat)
                img_path = tmp_dir / f"page_{pno}.png"
                pix.save(str(img_path))
            except Exception:
                if out_dir:
                    (out_dir / f"page_{pno}.txt").write_text("", encoding="utf-8")
                continue

            # Try multiple OCR passes: preprocessed first, then raw, then inverted-pre
            def _run_ocr(png: Path) -> str:
                try:
                    res = ocr.ocr(str(png), cls=True)  # type: ignore[attr-defined]
                except Exception:
                    return ""
                out_lines: List[str] = []
                for block in (res or []):
                    for item in (block or []):
                        try:
                            txt = item[1][0]
                            if isinstance(txt, str) and txt.strip():
                                out_lines.append(txt)
                        except Exception:
                            pass
                return "\n".join(out_lines)

            page_text = ""
            tried: List[str] = []
            if preprocess:
                pre_path = tmp_dir / f"page_{pno}.pre.png"
                _preprocess_image(img_path, pre_path)
                tried.append("preprocess")
                page_text = _run_ocr(pre_path)
            if not page_text:
                tried.append("raw")
                page_text = _run_ocr(img_path)
            if not page_text and preprocess:
                # Invert the preprocessed image and try again
                try:
                    import cv2  # type: ignore
                    inv_path = tmp_dir / f"page_{pno}.pre.inv.png"
                    m = cv2.imread(str(pre_path), cv2.IMREAD_GRAYSCALE)
                    if m is not None:
                        m = 255 - m
                        cv2.imwrite(str(inv_path), m)
                        tried.append("pre_inv")
                        page_text = _run_ocr(inv_path)
                except Exception:
                    pass

            if out_dir:
                (out_dir / f"page_{pno}.txt").write_text(page_text, encoding="utf-8")
            combined.append(f"\n--- PAGE {pno} ---\n{page_text}\n")
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

    out_txt.write_text("".join(combined), encoding="utf-8")
    print(f"[DONE] Wrote text -> {out_txt}")


def main() -> None:
    ap = argparse.ArgumentParser(description="Dump PDF text via PaddleOCR (PyMuPDF render)")
    ap.add_argument("--pdf", required=True, help="Path to input PDF")
    ap.add_argument("--out", help="Output .txt file (combined). Default under Product_Data_File/paddle_text")
    ap.add_argument("--out-dir", help="Optional folder to write per-page text files")
    ap.add_argument("--dpi", type=int, default=600, help="Render DPI (200..800, default 600)")
    ap.add_argument("--lang", default="en", help="PaddleOCR language code (default en)")
    ap.add_argument("--no-preprocess", action="store_true", help="Disable image pre-processing before OCR")
    args = ap.parse_args()

    pdf_path = Path(args.pdf).resolve()
    if not pdf_path.exists():
        print(f"[ERROR] PDF not found: {pdf_path}", file=sys.stderr)
        sys.exit(2)

    root = Path.cwd()
    default_out = root / "Product_Data_File" / "paddle_text" / (pdf_path.stem + ".txt")
    out_txt = Path(args.out).resolve() if args.out else default_out
    out_dir = Path(args.out_dir).resolve() if args.out_dir else None

    # argparse maps "--no-preprocess" to attribute name "no_preprocess"
    ocr_pdf_to_text(pdf_path, out_txt, dpi=args.dpi, lang=args.lang, out_dir=out_dir, preprocess=(not args.no_preprocess))


if __name__ == "__main__":
    main()
