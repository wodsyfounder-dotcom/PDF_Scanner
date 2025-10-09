#!/usr/bin/env python3
"""
OCR Export Utility
------------------
Render one or more PDF pages to images and export OCR output as:
- Searchable PDF (text layer)
- Plain text
- hOCR (HTML with positions)

Examples:
  py Application\ocr_export.py --pdf user_inputs\EIDP_Import_Docs\SN 1111.pdf \
     --pages 1 --format pdf --dpi 600 --psm 4 --out Product_Data_File\run_data\ocr_SN1111_p1.pdf

  py Application\ocr_export.py --pdf user_inputs\EIDP_Import_Docs\SN 1111.pdf \
     --pages 1 --format text --dpi 600 --psm 4 --out Product_Data_File\run_data\ocr_SN1111_p1.txt
"""

import argparse
import os
from pathlib import Path
from typing import List


def parse_pages(s: str) -> List[int]:
    out = set()
    s = (s or "").strip()
    if not s:
        return []
    s = s.replace("–", "-").replace("—", "-")
    for part in [p for p in s.replace(";", ",").split(",") if p.strip()]:
        if "-" in part:
            try:
                a, b = part.split("-", 1)
                a, b = int(a), int(b)
                if a > b:
                    a, b = b, a
                for p in range(a, b + 1):
                    out.add(p)
            except Exception:
                continue
        else:
            try:
                out.add(int(part))
            except Exception:
                continue
    return sorted(out)


def main() -> None:
    ap = argparse.ArgumentParser(description="Export OCR for PDF pages to PDF/text/hOCR")
    ap.add_argument("--pdf", required=True, help="Path to PDF file")
    ap.add_argument("--pages", default="1", help="Pages, e.g. '1' or '1-3,5'")
    ap.add_argument("--format", choices=["pdf", "text", "hocr"], default="pdf", help="Output format")
    ap.add_argument("--out", required=True, help="Output file path (single page) or prefix (multi-page)")
    ap.add_argument("--dpi", type=int, default=int(os.environ.get("OCR_DPI", "600")), help="Render DPI (200-1200)")
    ap.add_argument("--psm", type=int, default=4, help="Tesseract PSM (e.g., 4=columns, 6=block)")
    ap.add_argument("--renderer", choices=["auto", "pymupdf", "pdf2image"], default=os.environ.get("OCR_RENDERER", "auto"), help="Rendering backend")
    # Optional image pre-processing controls
    ap.add_argument("--scale", type=float, default=1.0, help="Upscale factor before OCR (e.g., 1.5)")
    ap.add_argument("--contrast", type=float, default=1.0, help="Contrast factor (1.0=no change; try 1.2-1.8)")
    ap.add_argument("--threshold", type=int, default=None, help="Binarize threshold 0-255 (omit for none)")
    ap.add_argument("--sharpen", action="store_true", help="Apply a sharpening filter before OCR")
    args = ap.parse_args()

    pdf_path = Path(args.pdf)
    if not pdf_path.exists():
        raise SystemExit(f"PDF not found: {pdf_path}")

    pages = parse_pages(args.pages) or [1]
    dpi = max(200, min(1200, args.dpi))
    tess_cfg = os.environ.get("TESSERACT_ARGS") or f"--psm {args.psm}"

    # Optional Tesseract path override
    try:
        import pytesseract
        tc = os.environ.get("TESSERACT_CMD") or os.environ.get("TESSERACT_PATH")
        if tc and os.path.exists(tc):
            pytesseract.pytesseract.tesseract_cmd = tc
        _ = pytesseract.get_tesseract_version()
    except Exception:
        raise SystemExit("Tesseract engine not available. Install it or set TESSERACT_CMD to tesseract.exe")

    # Choose renderer
    renderer = args.renderer.lower() if args.renderer else "auto"
    use_pymupdf = False
    use_pdf2image = False
    if renderer in ("auto", "pymupdf"):
        try:
            import fitz  # type: ignore
            use_pymupdf = True
        except Exception:
            use_pymupdf = False
    if (renderer == "pdf2image") or (renderer == "auto" and not use_pymupdf):
        try:
            from pdf2image import convert_from_path  # type: ignore
            use_pdf2image = True
        except Exception:
            use_pdf2image = False

    if not (use_pymupdf or use_pdf2image):
        raise SystemExit("No renderer available. Install PyMuPDF or pdf2image (with Poppler).")

    out = Path(args.out)
    multi = len(pages) > 1

    def _preprocess(img):
        from PIL import Image, ImageEnhance, ImageFilter
        out_img = img
        # Ensure RGB base
        if out_img.mode not in ("RGB", "L"):
            out_img = out_img.convert("RGB")
        # Optional upscale
        if args.scale and args.scale != 1.0:
            try:
                resample = getattr(Image, 'Resampling', Image).LANCZOS
            except Exception:
                resample = Image.LANCZOS
            w, h = out_img.size
            out_img = out_img.resize((int(w*args.scale), int(h*args.scale)), resample)
        # Optional contrast
        if args.contrast and args.contrast != 1.0:
            try:
                out_img = ImageEnhance.Contrast(out_img).enhance(args.contrast)
            except Exception:
                pass
        # Optional binarization
        if args.threshold is not None:
            try:
                gray = out_img.convert('L')
                t = max(0, min(255, int(args.threshold)))
                out_img = gray.point(lambda x: 255 if x >= t else 0, mode='1').convert('L')
            except Exception:
                pass
        # Optional sharpen
        if args.sharpen:
            try:
                from PIL import ImageFilter
                out_img = out_img.filter(ImageFilter.SHARPEN)
            except Exception:
                pass
        return out_img

    def write_page_output(pnum: int, img) -> None:
        import pytesseract
        from PIL import Image
        if not isinstance(img, Image.Image):
            raise RuntimeError("Renderer did not return a PIL Image")
        img = _preprocess(img)
        if args.format == "pdf":
            data = pytesseract.image_to_pdf_or_hocr(img, extension='pdf', lang='eng', config=tess_cfg)
            target = (out.parent / f"{out.stem}_p{pnum}.pdf") if multi else out
            target.write_bytes(data)
        elif args.format == "hocr":
            data = pytesseract.image_to_pdf_or_hocr(img, extension='hocr', lang='eng', config=tess_cfg)
            target = (out.parent / f"{out.stem}_p{pnum}.hocr.html") if multi else out
            target.write_bytes(data)
        else:
            text = pytesseract.image_to_string(img, lang='eng', config=tess_cfg)
            target = (out.parent / f"{out.stem}_p{pnum}.txt") if multi else out
            target.write_text(text or "", encoding="utf-8")

    if use_pymupdf:
        import fitz  # type: ignore
        doc = fitz.open(str(pdf_path))
        try:
            for p in pages:
                if 1 <= p <= doc.page_count:
                    page = doc.load_page(p - 1)
                    pix = page.get_pixmap(dpi=dpi)
                    from PIL import Image
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    write_page_output(p, img)
        finally:
            doc.close()
        return

    # pdf2image path
    from pdf2image import convert_from_path  # type: ignore
    images = convert_from_path(str(pdf_path), dpi=dpi, first_page=min(pages), last_page=max(pages))
    page_set = set(pages)
    start = min(pages)
    for idx, img in enumerate(images, start=start):
        if idx in page_set:
            write_page_output(idx, img)


if __name__ == "__main__":
    main()
