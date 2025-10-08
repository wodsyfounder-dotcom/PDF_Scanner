#!/usr/bin/env python3
"""
Simple PDF text dumper (self-contained in /Application).

Usage:
  py Application/dumppdf.py path/to/file.pdf [page]

If page is omitted, dumps page 1. Uses PyMuPDF if available, else pypdf.
"""
from __future__ import annotations

import sys
from pathlib import Path


def dump_with_pymupdf(pdf: Path, page_index: int) -> str:
    try:
        import fitz  # PyMuPDF
    except Exception:
        return ""
    doc = fitz.open(str(pdf))
    try:
        if not (0 <= page_index < doc.page_count):
            return ""
        return doc.load_page(page_index).get_text("text")
    finally:
        doc.close()


def dump_with_pypdf(pdf: Path, page_index: int) -> str:
    try:
        from pypdf import PdfReader
    except Exception:
        try:
            from PyPDF2 import PdfReader  # legacy
        except Exception:
            return ""
    r = PdfReader(str(pdf))
    if not (0 <= page_index < len(r.pages)):
        return ""
    return r.pages[page_index].extract_text() or ""


def main() -> None:
    if len(sys.argv) < 2:
        print("Usage: py Application/dumppdf.py file.pdf [page]", file=sys.stderr)
        sys.exit(2)
    pdf = Path(sys.argv[1])
    p = int(sys.argv[2]) if len(sys.argv) > 2 else 1
    idx = max(0, p - 1)
    txt = dump_with_pymupdf(pdf, idx) or dump_with_pypdf(pdf, idx)
    print(txt)


if __name__ == "__main__":
    main()

