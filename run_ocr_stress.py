#!/usr/bin/env python3
"""
Run OCR on the stress-case PDF and print the lines for inspection.
Uses the existing EasyOCR pipeline from eidp_term_scanner.core.
"""
from pathlib import Path
import importlib.util
import os
import sys

PDF_PATH = Path(r"C:\Users\zachs\Documents\DevProjects\PDF_Scanner\Data Packages\ocr_stress_case.pdf")
CORE_PATH = Path(r"C:\Users\zachs\Documents\DevProjects\PDF_Scanner\EIDAT_App_Files\Application\eidp_term_scanner.core.py")


def load_core(path: Path):
    spec = importlib.util.spec_from_file_location("eidp_core", path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Unable to load core module from {path}")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)  # type: ignore[attr-defined]
    return mod


def main():
    if not PDF_PATH.exists():
        print(f"[ERROR] PDF not found: {PDF_PATH}", file=sys.stderr)
        sys.exit(1)
    if not CORE_PATH.exists():
        print(f"[ERROR] Core not found: {CORE_PATH}", file=sys.stderr)
        sys.exit(1)

    # Set a default DPI suitable for stress testing; override via env if desired
    os.environ.setdefault("OCR_DPI", "1200")

    core = load_core(CORE_PATH)
    text_map, pipe = core.ocr_pages_with_easyocr(PDF_PATH, [1])
    text = text_map.get(1, "") or ""
    print(f"pipeline: {pipe}")
    print("----- OCR LINES -----")
    for i, line in enumerate(text.splitlines(), 1):
        print(f"{i:02d}: {line}")


if __name__ == "__main__":
    main()
