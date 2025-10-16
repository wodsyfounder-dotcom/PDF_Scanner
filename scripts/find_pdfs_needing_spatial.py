#!/usr/bin/env python3
"""
Scan the per-PDF JSONs under a run_data/<timestamp>/by_pdf folder and
print the full paths of PDFs that have any missing/empty results.

Criteria per record:
 - found is false OR number is null/empty string

Use --pdf-dir to resolve the filenames listed in the JSON to full paths.

Usage:
  py -3.11 scripts/find_pdfs_needing_spatial.py \
     --by-pdf-dir Product_Data_File/run_data/20251015_173211/by_pdf \
     --pdf-dir user_inputs/EIDP_Import_Docs
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path


def needs_spatial(rec: dict) -> bool:
    if not isinstance(rec, dict):
        return False
    if not rec.get('found', False):
        return True
    num = rec.get('number', None)
    if num is None:
        return True
    if isinstance(num, str) and not num.strip():
        return True
    return False


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument('--by-pdf-dir', required=True, help='Path to run_data/.../by_pdf folder')
    ap.add_argument('--pdf-dir', required=True, help='Folder containing the source PDFs')
    args = ap.parse_args()

    by_pdf = Path(args.by_pdf_dir)
    pdf_dir = Path(args.pdf_dir)
    out = []
    for j in sorted(by_pdf.glob('*.json')):
        try:
            data = json.loads(j.read_text(encoding='utf-8'))
        except Exception:
            continue
        hit = False
        if isinstance(data, list):
            for rec in data:
                if needs_spatial(rec):
                    hit = True
                    break
        if hit:
            # Resolve PDF path from record filename or from JSON name
            # Prefer rec['pdf_file'] if present
            pdf_name = None
            for rec in (data or []):
                if isinstance(rec, dict) and rec.get('pdf_file'):
                    pdf_name = str(rec['pdf_file'])
                    break
            if not pdf_name:
                pdf_name = j.stem + '.pdf'
            pdf_path = pdf_dir / pdf_name
            if pdf_path.exists():
                out.append(str(pdf_path))
            else:
                # Print anyway (caller can handle missing)
                out.append(str(pdf_path))
    for p in out:
        print(p)


if __name__ == '__main__':
    main()

