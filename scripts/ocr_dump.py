#!/usr/bin/env python3
"""
Quick EasyOCR dump for debugging pages that might need OCR.

Usage examples (from project root):
  py -3.11 scripts/ocr_dump.py --pdf "user_inputs/EIDP_Import_Docs/sn 4444.pdf" --pages 1 --dpi 800 --langs en
  py -3.11 scripts/ocr_dump.py --pdf "user_inputs/EIDP_Import_Docs/sn 4444.pdf" --pages 1-3 --dpi 600 --langs en --min-conf 0.2 --out-dir Product_Data_File/ocr_dump/sn_4444

Outputs:
  - Prints conf|text lines to stdout for each requested page.
  - If --out-dir is provided, writes per-page JSON lists of tokens with coordinates:
      [ {x0,y0,x1,y1,cx,cy,w,h,conf,text}, ... ]
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Dict, List


def parse_pages(pages_raw: str | None, max_page: int) -> List[int]:
    if not pages_raw:
        return [i for i in range(1, max_page + 1)]
    s = pages_raw.replace(" ", "")
    out: set[int] = set()
    for part in s.split(','):
        if not part:
            continue
        if '-' in part:
            a, b = part.split('-', 1)
            try:
                a, b = int(a), int(b)
            except Exception:
                continue
            if a > b:
                a, b = b, a
            for p in range(a, b + 1):
                if 1 <= p <= max_page:
                    out.add(p)
        else:
            try:
                p = int(part)
                if 1 <= p <= max_page:
                    out.add(p)
            except Exception:
                pass
    return sorted(out)


def main() -> None:
    ap = argparse.ArgumentParser(description="Dump EasyOCR tokens per page for a PDF")
    ap.add_argument("--pdf", required=True, help="Path to input PDF")
    ap.add_argument("--pages", default=None, help="Pages to OCR (e.g., '1' or '1-3,5')")
    ap.add_argument("--dpi", type=int, default=800, help="Render DPI (200..900)")
    ap.add_argument("--langs", default="en", help="Comma/semicolon list of languages (default en)")
    ap.add_argument("--min-conf", type=float, default=0.0, help="Min confidence to keep a token")
    ap.add_argument("--out-dir", default=None, help="Output directory for per-page JSON")
    args = ap.parse_args()

    try:
        import fitz  # type: ignore  # PyMuPDF
    except Exception as e:
        print("[ERROR] PyMuPDF is required: pip install pymupdf", file=sys.stderr)
        sys.exit(1)
    try:
        import easyocr  # type: ignore
    except Exception:
        print("[ERROR] EasyOCR is required: pip install easyocr (and CPU torch/torchvision)", file=sys.stderr)
        sys.exit(1)

    pdf = Path(args.pdf)
    out_dir = Path(args.out_dir) if args.out_dir else None
    langs = [s.strip() for s in args.langs.replace(";", ",").split(",") if s.strip()]

    try:
        doc = fitz.open(str(pdf))  # type: ignore[name-defined]
    except Exception as e:
        print(f"[ERROR] Unable to open PDF: {e}", file=sys.stderr)
        sys.exit(1)

    pages = parse_pages(args.pages, doc.page_count)
    if not pages:
        print("[WARN] No pages selected.")
        try:
            doc.close()
        except Exception:
            pass
        return

    if out_dir:
        out_dir.mkdir(parents=True, exist_ok=True)

    # Initialize EasyOCR reader once
    try:
        reader = easyocr.Reader(langs or ["en"], gpu=False, verbose=False)  # type: ignore
    except Exception as e:
        print(f"[ERROR] EasyOCR init failed: {type(e).__name__}: {e}", file=sys.stderr)
        try:
            doc.close()
        except Exception:
            pass
        sys.exit(1)

    for p in pages:
        if not (1 <= p <= doc.page_count):
            continue
        try:
            page = doc.load_page(p - 1)
            pix = page.get_pixmap(dpi=args.dpi)
        except Exception:
            print(f"[WARN] Failed to render page {p}")
            continue
        import tempfile, shutil
        tmp_dir = Path(tempfile.mkdtemp(prefix="ocr_dump_"))
        img_path = tmp_dir / f"p{p}.png"
        tokens: List[Dict[str, float]] = []
        try:
            pix.save(str(img_path))
            try:
                res = reader.readtext(str(img_path), detail=1)  # list of [bbox, text, conf]
            except Exception:
                res = []
            for it in res:
                try:
                    bbox, text, conf = it
                    if not isinstance(text, str) or not text.strip():
                        continue
                    cval = 0.0
                    try:
                        cval = float(conf) if conf is not None else 0.0
                    except Exception:
                        cval = 0.0
                    if cval < float(args.min_conf):
                        continue
                    xs = [float(pt[0]) for pt in bbox]
                    ys = [float(pt[1]) for pt in bbox]
                    x0, y0, x1, y1 = min(xs), min(ys), max(xs), max(ys)
                    cx = (x0 + x1) / 2.0
                    cy = (y0 + y1) / 2.0
                    tokens.append({
                        "x0": x0, "y0": y0, "x1": x1, "y1": y1,
                        "cx": cx, "cy": cy, "w": (x1 - x0), "h": (y1 - y0),
                        "conf": cval, "text": text.strip(),
                    })
                except Exception:
                    # Defensive: skip malformed item
                    pass
        finally:
            try:
                shutil.rmtree(str(tmp_dir), ignore_errors=True)
            except Exception:
                pass

        print(f"\n=== Page {p} (tokens >= {args.min_conf}) ===")
        for t in sorted(tokens, key=lambda x: (-x["conf"], x["cy"], x["cx"]))[:1000]:
            print(f"{t['conf']:.3f} | {t['text']}")

        if out_dir is not None:
            out_path = out_dir / f"{pdf.stem}_page_{p}.ocr.json"
            out_path.write_text(json.dumps(tokens, ensure_ascii=False, indent=2), encoding="utf-8")

    try:
        doc.close()
    except Exception:
        pass


if __name__ == "__main__":
    main()
