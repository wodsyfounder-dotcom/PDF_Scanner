#!/usr/bin/env python3
"""
Enrich the run registry with Program and Space Vehicle columns when available.

Derives Program and Space Vehicle from each run's scan_results.json or
from the source PDF filename convention: Program_SpaceVehicle_Serial.pdf

Updates Product_Data_File/run_registry.xlsx (preferred) or run_registry.csv.
Safe to run multiple times; keeps existing values and fills blanks.
"""
from __future__ import annotations

import csv
import json
import re
from pathlib import Path
from typing import Dict, List, Tuple

ROOT = Path(__file__).resolve().parents[1]
EXPORTS = ROOT / "Product_Data_File"
REG_XLSX = EXPORTS / "run_registry.xlsx"
REG_CSV = EXPORTS / "run_registry.csv"


def norm(s):
    if s is None:
        return ""
    return str(s).strip()


def derive_from_filename(pdf_name: str) -> Tuple[str, str]:
    stem = Path(pdf_name).stem
    if "_" in stem:
        parts = [p.strip() for p in stem.split("_") if p.strip()]
        if len(parts) >= 3:
            return parts[0], parts[1]
    toks = [t for t in re.split(r"\s+", stem) if t]
    if len(toks) >= 2:
        return toks[0], " ".join(toks[1:])
    return "", ""


def normalize_sn(sn: str) -> str:
    s = norm(sn)
    if not s:
        return ""
    # Match SN, optionally separated by space/underscore/dash, then the id
    try:
        m = re.match(r"^(?i:sn)[\s_\-]*([A-Za-z0-9]+)$", s)
    except Exception:
        m = None
    if m:
        return f"SN {m.group(1)}"
    # Already looks like "SN 1234" -> keep
    try:
        if re.match(r"^(?i:sn)\s+.+", s):
            return s if s.startswith("SN ") else f"SN {s.split(None,1)[1]}"
    except Exception:
        pass
    return s


def read_registry_rows() -> List[Dict[str, str]]:
    rows: List[Dict[str, str]] = []
    if REG_XLSX.exists():
        try:
            import pandas as pd  # type: ignore
            df = pd.read_excel(REG_XLSX)
            for _, r in df.iterrows():
                rows.append({k: norm(r.get(k)) for k in df.columns})
            return rows
        except Exception:
            try:
                from openpyxl import load_workbook  # type: ignore
                wb = load_workbook(str(REG_XLSX), data_only=True)
                ws = wb.active
                header = [norm(c.value) for c in next(ws.iter_rows(min_row=1, max_row=1))]
                for row in ws.iter_rows(min_row=2):
                    rec = {}
                    for i, c in enumerate(row):
                        if i < len(header):
                            rec[header[i]] = norm(c.value)
                    rows.append(rec)
                return rows
            except Exception:
                pass
    if REG_CSV.exists():
        try:
            with REG_CSV.open("r", encoding="utf-8", newline="") as f:
                r = csv.DictReader(f)
                for row in r:
                    rows.append({k: norm(v) for k, v in row.items()})
        except Exception:
            pass
    return rows


def write_registry_rows(rows: List[Dict[str, str]]) -> None:
    # Prefer Excel if pandas + xlsxwriter
    try:
        import pandas as pd  # type: ignore
        import xlsxwriter  # noqa: F401
        cols = ["serial_number", "program", "space_vehicle", "run_date", "run_folder"]
        for c in cols:
            for r in rows:
                r.setdefault(c, "")
        df = pd.DataFrame(rows, columns=cols)
        with pd.ExcelWriter(REG_XLSX, engine="xlsxwriter") as writer:
            df.to_excel(writer, sheet_name="runs", index=False)
            ws = writer.sheets["runs"]
            ws.freeze_panes(1, 0)
            for i, col in enumerate(df.columns):
                try:
                    max_len = int(df[col].astype(str).map(len).max()) if not df.empty else len(col)
                except Exception:
                    max_len = len(col)
                ws.set_column(i, i, min(80, max(12, max_len + 2)))
        return
    except Exception:
        pass
    # CSV fallback
    with REG_CSV.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["serial_number", "program", "space_vehicle", "run_date", "run_folder"])
        w.writeheader()
        for r in rows:
            out = {k: r.get(k, "") for k in ["serial_number", "program", "space_vehicle", "run_date", "run_folder"]}
            w.writerow(out)


def load_results_json(run_folder: Path) -> list[dict]:
    p = run_folder / "scan_results.json"
    try:
        with p.open("r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def main() -> None:
    rows = read_registry_rows()
    if not rows:
        # nothing to do
        return
    # Build map serial_number -> run_folder
    last_for_sn: Dict[str, str] = {}
    for r in rows:
        sn = normalize_sn(norm(r.get("serial_number")))
        rf = norm(r.get("run_folder"))
        if sn:
            last_for_sn[sn] = rf
    # derive metadata
    program_by_sn: Dict[str, str] = {}
    sv_by_sn: Dict[str, str] = {}
    for sn, rf in last_for_sn.items():
        run_folder = Path(rf)
        entries = load_results_json(run_folder)
        prog = ""
        sv = ""
        for e in entries:
            if normalize_sn(norm(e.get("serial_number"))) != sn:
                continue
            prog = norm(e.get("program"))
            sv = norm(e.get("space_vehicle"))
            if not prog or not sv:
                pdf = norm(e.get("pdf_file"))
                if pdf:
                    dprog, dsv = derive_from_filename(pdf)
                    prog = prog or dprog
                    sv = sv or dsv
            if prog or sv:
                break
        program_by_sn[sn] = prog
        sv_by_sn[sn] = sv

    # merge into rows (keeping existing non-blank values)
    # Build deduplicated map by normalized SN
    merged: Dict[str, Dict[str, str]] = {}
    for r in rows:
        sn = normalize_sn(norm(r.get("serial_number")))
        if not sn:
            continue
        r["serial_number"] = sn
        if not norm(r.get("program")) and program_by_sn.get(sn):
            r["program"] = program_by_sn[sn]
        if not norm(r.get("space_vehicle")) and sv_by_sn.get(sn):
            r["space_vehicle"] = sv_by_sn[sn]
        prev = merged.get(sn)
        if not prev or norm(r.get("run_date")) > norm(prev.get("run_date")):
            merged[sn] = r

    write_registry_rows(list(merged.values()))


if __name__ == "__main__":
    main()
