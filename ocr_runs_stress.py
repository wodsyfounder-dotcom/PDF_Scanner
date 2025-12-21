#!/usr/bin/env python3
"""
Run OCR on the stress-case PDF using the alternate geometric line-handling core.

This script mirrors run_ocr_stress.py but loads EIDAT_App_Files/Application/
eidp_term_scanner.core_alt.py instead, so we can iterate on the OCR pipeline
without affecting the primary core module.
"""
from pathlib import Path
import importlib.util
import os
import sys


PDF_PATH = Path(r"C:\Users\zachs\Documents\DevProjects\PDF_Scanner\Data Packages\ocr_stress_case.pdf")
CORE_PATH = Path(r"C:\Users\zachs\Documents\DevProjects\PDF_Scanner\EIDAT_App_Files\Application\OCR_line_geometry_solver.core.py")


def load_core(path: Path):
    spec = importlib.util.spec_from_file_location("eidp_core_alt", path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Unable to load alt core module from {path}")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)  # type: ignore[attr-defined]
    return mod


def main():
    if not PDF_PATH.exists():
        print(f"[ERROR] PDF not found: {PDF_PATH}", file=sys.stderr)
        sys.exit(1)
    if not CORE_PATH.exists():
        print(f"[ERROR] Alt core not found: {CORE_PATH}", file=sys.stderr)
        sys.exit(1)

    # Use same default DPI as the main stress runner; override via env if desired.
    os.environ.setdefault("OCR_DPI", "1200")

    core = load_core(CORE_PATH)
    text_map, pipe = core.ocr_pages_with_easyocr(PDF_PATH, [1])
    text = text_map.get(1, "") or ""
    print(f"pipeline: {pipe}")
    print("----- OCR LINES (ALT CORE) -----")
    for i, line in enumerate(text.splitlines(), 1):
        print(f"{i:02d}: {line}")


if __name__ == "__main__":
    main()
