#!/usr/bin/env python3
"""
Setup scaffold for EIDP Term Scanner (no OCR dependencies).

Creates a clean folder layout:

- user_inputs/
    - EIDP_Import_Docs/   # drop PDFs here
    - Scanned_Docs/       # processed PDFs moved here
    - terms.csv           # input terms (Term,Pages)

- Product_Data_File/
    - EIDP_data.csv       # aggregate export (growing wide matrix)
    - run_data/           # per-run snapshots of outputs

Usage:
  py Application/setup_scaffold.py
"""

from __future__ import annotations

import csv
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def ensure_dirs() -> None:
    (ROOT / "user_inputs" / "EIDP_Import_Docs").mkdir(parents=True, exist_ok=True)
    (ROOT / "user_inputs" / "Scanned_Docs").mkdir(parents=True, exist_ok=True)
    (ROOT / "Product_Data_File" / "run_data").mkdir(parents=True, exist_ok=True)


def ensure_terms_csv() -> None:
    dst = ROOT / "user_inputs" / "terms.csv"
    if dst.exists():
        return
    # Copy existing root terms.csv if present, otherwise create a sample
    src = ROOT / "terms.csv"
    if src.exists():
        dst.write_text(src.read_text(encoding="utf-8"), encoding="utf-8")
        return
    with dst.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Term", "Pages"])
        w.writerow(["Thrust", "1"])  # sample
        w.writerow(["Specific Impulse", "1"])  # sample


def ensure_aggregate_csv() -> None:
    agg = ROOT / "Product_Data_File" / "EIDP_data.csv"
    if agg.exists():
        return
    with agg.open("w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerow(["Term", "Pages"])  # start with headers only


def main() -> None:
    ensure_dirs()
    ensure_terms_csv()
    ensure_aggregate_csv()
    print("[READY] Created scaffold under:")
    print(f"  {ROOT / 'user_inputs'}")
    print(f"  {ROOT / 'Product_Data_File'}")
    print("")
    print("Run the scanner from repo root, for example:")
    print("  py .\\Application\\eidp_term_scanner.py \\")
    print("     --input .\\user_inputs\\terms.csv \\")
    print("     --pdf-folder .\\user_inputs\\EIDP_Import_Docs \\")
    print("     --scanned-folder .\\user_inputs\\Scanned_Docs \\")
    print("     --output-xlsx .\\Product_Data_File\\scan_results.xlsx ")


if __name__ == "__main__":
    main()
