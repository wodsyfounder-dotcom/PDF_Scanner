#!/usr/bin/env python3
"""
Compile a master extraction workbook from the per-run caches recorded in run_registry.

Inputs:
  - Product_Data_File/run_registry.xlsx (preferred) or run_registry.csv
  - For each row: (serial_number, run_folder)
  - Each run folder contains scan_results.json which holds per-term values per SN.

Output:
  - Product_Data_File/master.xlsx (preferred, pandas + xlsxwriter)
  - Fallback: Product_Data_File/master.csv

No external executables required. Uses pandas/xlsxwriter if available; otherwise falls back to CSV.
"""
from __future__ import annotations

import csv
import json
import sys
from pathlib import Path
from typing import Dict, List, Tuple, Optional


ROOT = Path(__file__).resolve().parents[1]
EXPORTS = ROOT / "Product_Data_File"
REG_XLSX = EXPORTS / "run_registry.xlsx"
REG_CSV = EXPORTS / "run_registry.csv"
OUT_XLSX = EXPORTS / "master.xlsx"
OUT_CSV = EXPORTS / "master.csv"


def load_registry() -> List[Tuple[str, Path]]:
    """Return list of (serial_number, run_folder) from registry.
    Priority: Excel -> CSV. Duplicates are unlikely; if present, keep last occurrence.
    """
    rows: List[Tuple[str, Path]] = []
    if REG_XLSX.exists():
        # Try pandas first
        try:
            import pandas as pd  # type: ignore
            df = pd.read_excel(REG_XLSX)
            for _, r in df.iterrows():
                sn = str(r.get("serial_number") or "").strip()
                rf = str(r.get("run_folder") or "").strip()
                if sn and rf:
                    rows.append((sn, Path(rf)))
            return rows
        except Exception:
            # Fallback: openpyxl without pandas
            try:
                from openpyxl import load_workbook  # type: ignore
                wb = load_workbook(str(REG_XLSX), data_only=True)
                ws = wb["runs"] if "runs" in wb.sheetnames else wb.active
                # Build header map (case-insensitive)
                header_map: Dict[str, int] = {}
                for ci, cell in enumerate(next(ws.iter_rows(min_row=1, max_row=1)), start=1):
                    key = (str(cell.value) if cell.value is not None else "").strip().lower()
                    if key:
                        header_map[key] = ci
                def col_idx(name: str) -> Optional[int]:
                    name = name.lower()
                    for k, v in header_map.items():
                        if k == name:
                            return v
                    return None
                sn_col = col_idx("serial_number")
                rf_col = col_idx("run_folder")
                if sn_col and rf_col:
                    for row in ws.iter_rows(min_row=2):
                        sn_val = row[sn_col - 1].value if sn_col else None
                        rf_val = row[rf_col - 1].value if rf_col else None
                        sn = (str(sn_val) if sn_val is not None else "").strip()
                        rf = (str(rf_val) if rf_val is not None else "").strip()
                        if sn and rf:
                            rows.append((sn, Path(rf)))
                    if rows:
                        return rows
            except Exception:
                pass
    if REG_CSV.exists():
        try:
            with REG_CSV.open("r", encoding="utf-8", newline="") as f:
                r = csv.DictReader(f)
                for row in r:
                    sn = (row.get("serial_number") or "").strip()
                    rf = (row.get("run_folder") or "").strip()
                    if sn and rf:
                        rows.append((sn, Path(rf)))
        except Exception:
            pass
    return rows


def load_results_json(run_folder: Path) -> List[Dict]:
    path = run_folder / "scan_results.json"
    try:
        with path.open("r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def build_master() -> Tuple[List[str], Dict[str, Dict[str, Optional[str]]]]:
    """Return (serials, matrix) where matrix[term][sn] = value string or None."""
    reg = load_registry()
    if not reg:
        print("[WARN] No registry entries found. Nothing to compile.")
        return [], {}

    # Keep last occurrence per SN (registry is already latest-first, but be safe)
    last_for_sn: Dict[str, Path] = {}
    for sn, rf in reg:
        last_for_sn[sn] = rf

    terms_order: List[str] = []
    matrix: Dict[str, Dict[str, Optional[str]]] = {}

    for sn, rf in last_for_sn.items():
        rows = load_results_json(rf)
        if not rows:
            print(f"[WARN] No scan_results.json in {rf}")
            continue
        # Build a best-value per term for this SN (first occurrence wins per run)
        seen_terms = set()
        for row in rows:
            if (row.get("serial_number") or "").strip() != sn:
                continue
            term = (row.get("term") or "").strip()
            if not term or term in seen_terms:
                continue
            val = row.get("number")
            seen_terms.add(term)
            if term not in matrix:
                matrix[term] = {}
                terms_order.append(term)
            matrix[term][sn] = val

    serials = list(last_for_sn.keys())
    return serials, matrix


def write_master(serials: List[str], matrix: Dict[str, Dict[str, Optional[str]]]) -> None:
    EXPORTS.mkdir(parents=True, exist_ok=True)
    # Try Excel via pandas/xlsxwriter
    try:
        import pandas as pd  # type: ignore
        import xlsxwriter  # noqa: F401
        rows = []
        for term in matrix.keys():
            row = {"Term": term}
            for sn in serials:
                row[sn] = matrix.get(term, {}).get(sn)
            rows.append(row)
        df = pd.DataFrame(rows, columns=["Term"] + serials)
        with pd.ExcelWriter(OUT_XLSX, engine="xlsxwriter") as writer:
            df.to_excel(writer, sheet_name="master", index=False)
            ws = writer.sheets["master"]
            ws.freeze_panes(1, 1)
            # Auto-size columns
            for i, col in enumerate(df.columns):
                try:
                    max_len = int(df[col].astype(str).map(len).max()) if not df.empty else len(col)
                except Exception:
                    max_len = len(col)
                ws.set_column(i, i, min(60, max(10, max_len + 2)))
        print(f"[DONE] Master workbook -> {OUT_XLSX}")
        return
    except Exception as e:
        print(f"[WARN] Excel write unavailable ({e}); falling back to CSV")

    # CSV fallback
    try:
        with OUT_CSV.open("w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["Term"] + serials)
            for term in matrix.keys():
                row = [term]
                for sn in serials:
                    row.append(matrix.get(term, {}).get(sn))
                w.writerow(row)
        print(f"[DONE] Master CSV -> {OUT_CSV}")
    except Exception as e:
        print(f"[ERROR] Could not write master CSV: {e}")


def main() -> None:
    serials, matrix = build_master()
    if not serials:
        sys.exit(0)
    write_master(serials, matrix)


if __name__ == "__main__":
    main()
