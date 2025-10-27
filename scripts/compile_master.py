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

Workbook layout:
  - Leading rows reserved for "Program" and "Space Vehicle" metadata.
  - Columns include Grouping (group_after), Units, Row Label, and Column Label to align multi-row terms.

No external executables required. Uses pandas/xlsxwriter if available; otherwise falls back to CSV.
"""
from __future__ import annotations

import csv
import json
import sys
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Any


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


def build_master() -> Tuple[List[str], List[Dict[str, Any]], Dict[str, str], Dict[str, str]]:
    """Return (serials, rows) including per-term row/column breakdown and captured values."""
    reg = load_registry()
    if not reg:
        print("[WARN] No registry entries found. Nothing to compile.")
        return [], []

    # Keep last occurrence per SN (registry is already latest-first, but be safe)
    last_for_sn: Dict[str, Path] = {}
    for sn, rf in reg:
        last_for_sn[sn] = rf

    serials = list(last_for_sn.keys())
    terms_order: List[str] = []
    term_map: Dict[str, Dict[str, Any]] = {}
    program_by_sn: Dict[str, str] = {}
    sv_by_sn: Dict[str, str] = {}

    def norm(value: Any) -> str:
        if value is None:
            return ""
        if isinstance(value, str):
            return value.strip()
        return str(value).strip()

    def extract_units(entry: Dict[str, Any]) -> str:
        direct = norm(entry.get("units"))
        if direct:
            return direct
        hints = entry.get("units_hint")
        if isinstance(hints, (list, tuple)):
            parts = [norm(h) for h in hints if norm(h)]
            if parts:
                # Deduplicate while preserving order
                unique: List[str] = []
                seen: set[str] = set()
                for part in parts:
                    key = part.lower()
                    if key in seen:
                        continue
                    seen.add(key)
                    unique.append(part)
                return " | ".join(unique)
        return ""

    def extract_value(entry: Dict[str, Any]) -> Optional[str]:
        for key in ("number", "text", "string", "value"):
            if key in entry:
                val = entry.get(key)
                if val is None:
                    continue
                text = norm(val)
                return text
        return None

    for sn, rf in last_for_sn.items():
        rows = load_results_json(rf)
        if not rows:
            print(f"[WARN] No scan_results.json in {rf}")
            continue
        # Capture every row for this SN, grouped by term with row/column detail
        for row in rows:
            if (row.get("serial_number") or "").strip() != sn:
                continue
            # Capture metadata per SN if present (or derive from filename)
            prog = norm(row.get("program"))
            sv = norm(row.get("space_vehicle"))
            if not prog:
                # Derive from filename
                pdf_file = norm(row.get("pdf_file"))
                stem = Path(pdf_file).stem if pdf_file else ""
                # simple parse mirroring enrich script behavior
                if stem:
                    if "_" in stem:
                        parts = [p.strip() for p in stem.split("_") if p.strip()]
                        # find SN part index
                        sn_idx = None
                        for i, p in enumerate(parts):
                            if re.search(r"\bSN\b", p, flags=re.IGNORECASE) or re.search(r"\bSN\W*", p, flags=re.IGNORECASE):
                                sn_idx = i
                                break
                        if sn_idx is None:
                            sn_idx = len(parts)
                        if sn_idx >= 2:
                            prog = parts[0]
                            sv = " ".join(parts[1:sn_idx])
                    else:
                        toks = [t for t in re.split(r"\s+", stem) if t]
                        # locate token matching SN*
                        si = None
                        for i, t in enumerate(toks):
                            if t.lower().startswith("sn"):
                                si = i
                                break
                        if si is None and len(toks) >= 2:
                            prog = toks[0]
                            sv = " ".join(toks[1:])
                        elif si is not None and si >= 2:
                            prog = toks[0]
                            sv = " ".join(toks[1:si])
            if prog and sn not in program_by_sn:
                program_by_sn[sn] = prog
            if sv and sn not in sv_by_sn:
                sv_by_sn[sn] = sv
            term = (row.get("term") or "").strip()
            if not term:
                continue
            value = extract_value(row)
            if value is None:
                value = ""

            if term not in term_map:
                term_map[term] = {
                    "order": [],
                    "entries": {}
                }
                terms_order.append(term)
            term_info = term_map[term]

            group_after = norm(row.get("group_after"))
            units = extract_units(row)
            row_label = norm(row.get("line") or row.get("row_label"))
            column_label = norm(row.get("column") or row.get("column_label"))
            group_key = group_after.lower()
            entry_key = (row_label.lower(), column_label.lower(), group_key)

            entries: Dict[Tuple[str, str, str], Dict[str, Any]] = term_info["entries"]
            if entry_key not in entries:
                entries[entry_key] = {
                    "group": group_after,
                    "row_label": row_label,
                    "column_label": column_label,
                    "units": units,
                    "values": {}
                }
                term_info["order"].append(entry_key)
            entry = entries[entry_key]
            if not entry.get("group"):
                entry["group"] = group_after
            if units and not entry.get("units"):
                entry["units"] = units
            entry["values"][sn] = value

    term_rows: List[Dict[str, Any]] = []
    for term in terms_order:
        info = term_map.get(term)
        if not info:
            continue
        for entry_key in info["order"]:
            entry = info["entries"][entry_key]
            term_rows.append({
                "term": term,
                "group": entry.get("group", ""),
                "units": entry.get("units", ""),
                "row_label": entry["row_label"],
                "column_label": entry["column_label"],
                "values": entry["values"],
            })

    return serials, term_rows, program_by_sn, sv_by_sn


def write_master(serials: List[str], term_rows: List[Dict[str, Any]], program_by_sn: Dict[str, str] | None = None, sv_by_sn: Dict[str, str] | None = None) -> None:
    EXPORTS.mkdir(parents=True, exist_ok=True)
    base_columns = ["Term", "Grouping", "Units", "Row Label", "Column Label"]
    header = base_columns + serials

    def blank_meta(label: str) -> Dict[str, Any]:
        row = {col: "" for col in header}
        row["Term"] = label
        return row

    structured_rows: List[Dict[str, Any]] = [blank_meta("Program"), blank_meta("Space Vehicle")]

    # Fill Program / Space Vehicle rows per serial
    program_by_sn = program_by_sn or {}
    sv_by_sn = sv_by_sn or {}
    # Row 0 -> Program, Row 1 -> Space Vehicle
    for sn in serials:
        structured_rows[0][sn] = program_by_sn.get(sn, "")
        structured_rows[1][sn] = sv_by_sn.get(sn, "")

    for entry in term_rows:
        row: Dict[str, Any] = {col: "" for col in header}
        row["Term"] = entry.get("term", "")
        row["Grouping"] = entry.get("group", "")
        row["Units"] = entry.get("units", "")
        row["Row Label"] = entry.get("row_label", "")
        row["Column Label"] = entry.get("column_label", "")
        values = entry.get("values", {}) or {}
        for sn in serials:
            row[sn] = values.get(sn, "")
        structured_rows.append(row)

    # Try Excel via pandas/xlsxwriter
    try:
        import pandas as pd  # type: ignore
        import xlsxwriter  # noqa: F401
        df = pd.DataFrame(structured_rows, columns=header)
        with pd.ExcelWriter(OUT_XLSX, engine="xlsxwriter") as writer:
            df.to_excel(writer, sheet_name="master", index=False)
            ws = writer.sheets["master"]
            # Freeze header + Program/Space Vehicle rows, and keep core columns visible.
            ws.freeze_panes(3, 3)
            # Auto-size columns
            for i, col in enumerate(df.columns):
                try:
                    max_len = int(df[col].astype(str).map(len).max()) if not df.empty else len(col)
                except Exception:
                    max_len = len(col)
                ws.set_column(i, i, min(60, max(12, max_len + 2)))
        print(f"[DONE] Master workbook -> {OUT_XLSX}")
        return
    except Exception as e:
        print(f"[WARN] Excel write unavailable ({e}); falling back to CSV")

    # CSV fallback
    try:
        with OUT_CSV.open("w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(header)
            for row_dict in structured_rows:
                row = [row_dict.get(col, "") for col in header]
                w.writerow(row)
        print(f"[DONE] Master CSV -> {OUT_CSV}")
    except Exception as e:
        print(f"[ERROR] Could not write master CSV: {e}")


def main() -> None:
    serials, rows, prog_map, sv_map = build_master()
    if not serials:
        sys.exit(0)
    write_master(serials, rows, program_by_sn=prog_map, sv_by_sn=sv_map)


if __name__ == "__main__":
    main()
