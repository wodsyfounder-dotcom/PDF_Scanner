#!/usr/bin/env python3
"""
Enrich scan_results.json and scan_results_flat.xlsx with Program and Space Vehicle
parsed from the EIDP filename convention:

  Program_Space Vehicle_SNXXXX.pdf

Also supports space-delimited names like: Program SpaceVehicle SNXXXX.pdf

Usage:
  python scripts/enrich_scan_results.py --run-dir Product_Data_File/run_data/20250101_120000
"""
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Dict, Tuple


SN_REGEX = re.compile(r"\bSN\W*([A-Za-z0-9][A-Za-z0-9_\-]*)", re.IGNORECASE)


def parse_eidp_parts(filename_stem: str) -> Tuple[str, str]:
    """Return (Program, Space Vehicle) parsed from a filename stem.

    Rules:
      - Identify SN position using a full-name regex (allows 'SN 1234' or 'SN-1234').
      - Only set Program/SV when there is meaningful content before the SN.
      - Underscore style: Program_Space Vehicle_SNXXXX -> Program, Space Vehicle.
      - Space style: Program SpaceVehicle SNXXXX -> Program, SpaceVehicle.
      - If no content before SN or ambiguous, return blanks.
    """
    name = filename_stem.strip()
    m = SN_REGEX.search(name)
    if not m:
        return "", ""  # no explicit SN in the name; leave blank
    pre = name[: m.start()].strip(" _-\t")
    if not pre:
        return "", ""
    if "_" in pre:
        parts = [p.strip() for p in pre.split("_") if p.strip()]
        if len(parts) >= 2:
            return parts[0], " ".join(parts[1:])
        return "", ""
    # Space-delimited
    tokens = [t for t in re.split(r"\s+", pre) if t]
    if len(tokens) >= 2:
        return tokens[0], " ".join(tokens[1:])
    return "", ""


def enrich_run_dir(run_dir: Path) -> None:
    run_dir = run_dir.resolve()
    json_path = run_dir / "scan_results.json"
    if not json_path.exists():
        return
    try:
        data = json.loads(json_path.read_text(encoding="utf-8"))
    except Exception:
        return
    if not isinstance(data, list):
        return

    # Compute program/SV per serial by filename
    prog_by_sn: Dict[str, str] = {}
    sv_by_sn: Dict[str, str] = {}
    for row in data:
        pdf_file = str(row.get("pdf_file") or "").strip()
        stem = Path(pdf_file).stem
        program, space_vehicle = parse_eidp_parts(stem)
        sn = str(row.get("serial_number") or "").strip()
        if program and sn and sn not in prog_by_sn:
            prog_by_sn[sn] = program
        if space_vehicle and sn and sn not in sv_by_sn:
            sv_by_sn[sn] = space_vehicle
    # Apply enrichment to every row
    for row in data:
        sn = str(row.get("serial_number") or "").strip()
        row["program"] = prog_by_sn.get(sn, "")
        row["space_vehicle"] = sv_by_sn.get(sn, "")
    try:
        json_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass

    # Enrich flat workbook if present
    xlsx_path = run_dir / "scan_results_flat.xlsx"
    if not xlsx_path.exists():
        return
    try:
        import pandas as pd  # type: ignore
    except Exception:
        return
    try:
        df = pd.read_excel(xlsx_path, sheet_name="extraction")
    except Exception:
        return
    # Insert columns after serial_number if present, else after pdf_file
    cols = list(df.columns)
    def insert_cols(df):
        if "serial_number" in cols:
            idx = cols.index("serial_number") + 1
        elif "pdf_file" in cols:
            idx = cols.index("pdf_file") + 1
        else:
            idx = 0
        df.insert(idx, "Program", df.get("serial_number").map(lambda sn: prog_by_sn.get(str(sn).strip(), "")))
        df.insert(idx + 1, "Space Vehicle", df.get("serial_number").map(lambda sn: sv_by_sn.get(str(sn).strip(), "")))
        return df
    df = insert_cols(df)
    # Errors sheet enrichment (best-effort)
    try:
        df_err = pd.read_excel(xlsx_path, sheet_name="errors")
        if not df_err.empty:
            ce = list(df_err.columns)
            if "serial_number" in ce:
                df_err.insert(ce.index("serial_number") + 1, "Program", df_err.get("serial_number").map(lambda sn: prog_by_sn.get(str(sn).strip(), "")))
                df_err.insert(ce.index("serial_number") + 2, "Space Vehicle", df_err.get("serial_number").map(lambda sn: sv_by_sn.get(str(sn).strip(), "")))
    except Exception:
        df_err = None

    try:
        with pd.ExcelWriter(xlsx_path, engine="xlsxwriter") as writer:
            df.to_excel(writer, sheet_name="extraction", index=False)
            if df_err is not None:
                df_err.to_excel(writer, sheet_name="errors", index=False)
    except Exception:
        pass


def main() -> None:
    ap = argparse.ArgumentParser(description="Enrich scan results with Program and Space Vehicle")
    ap.add_argument("--run-dir", required=True, help="Path to a run_data folder containing scan_results.json")
    args = ap.parse_args()
    run_dir = Path(args.run_dir)
    enrich_run_dir(run_dir)


if __name__ == "__main__":
    main()
