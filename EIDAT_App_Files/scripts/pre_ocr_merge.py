#!/usr/bin/env python3
"""
Pre-OCR a PDF (or several) and emit merged text artifacts for downstream extraction.

Outputs per PDF:
- page_{n}.txt files with header/footer lines removed
- headers.txt / footers.txt
- combined.txt containing headers, footers, and body as a single stream
- manifest.json with page spans for offset->page mapping
"""

from __future__ import annotations

import argparse
import importlib.util
import sys
from pathlib import Path
from typing import List

APP_ROOT = Path(__file__).resolve().parents[1]
CORE_PATH = APP_ROOT / "Application" / "eidp_term_scanner.core.py"


def _load_core():
    spec = importlib.util.spec_from_file_location("eidp_core", CORE_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Unable to load core from {CORE_PATH}")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)  # type: ignore[attr-defined]
    return mod


def main(argv: List[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Pre-OCR and merge PDFs into single-text artifacts.")
    parser.add_argument("--pdf", action="append", dest="pdfs", required=True, help="Path to a PDF (can repeat).")
    parser.add_argument("--out", dest="out_root", help="Optional root directory for merged artifacts.")
    parser.add_argument("--dpi", type=int, default=None, help="Optional OCR DPI override for this run.")
    args = parser.parse_args(argv)

    core = _load_core()
    out_root = Path(args.out_root).expanduser() if args.out_root else None

    pdf_paths = [Path(p).expanduser() for p in args.pdfs]
    rc = 0
    for pdf in pdf_paths:
        print(f"[INFO] Pre-OCR + merge starting -> {pdf}")
        try:
            res = core.pre_ocr_and_merge_pdf(pdf, dpi=args.dpi, out_dir=out_root)
            print(f"[DONE] Combined: {res.get('combined')}")
            print(f"[INFO] Manifest: {res.get('manifest')}")
        except Exception as exc:  # noqa: BLE001
            rc = 1
            print(f"[ERROR] Failed for {pdf}: {exc}")
    return rc


if __name__ == "__main__":
    sys.exit(main())
