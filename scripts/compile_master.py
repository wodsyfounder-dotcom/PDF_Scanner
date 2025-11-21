#!/usr/bin/env python3
"""
Compile a master extraction workbook from the per-run caches recorded in run_registry.

Inputs:
  - Product_Data_File/run_registry.xlsx (preferred) or run_registry.csv
  - For each row: serial_component, run_folder, program_name, vehicle_number
  - Each run folder contains scan_results.json which holds per-term values per SN.

Output:
  - Product_Data_File/master.xlsx (preferred, pandas + xlsxwriter)
  - Fallback: Product_Data_File/master.csv

Workbook layout:
  - Leading rows reserved for "Program", "Space Vehicle", and "Data" metadata.
  - Columns include Term Label, Data Group, Units, Min, Max, followed by one column per serial number.

No external executables required. Uses pandas/xlsxwriter if available; otherwise falls back to CSV.
"""
from __future__ import annotations

import csv
import re
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


def load_registry() -> List[Tuple[str, Path, Dict[str, str]]]:
    """Return list of (serial_component, run_folder, metadata) from registry.
    Priority: Excel -> CSV. Duplicates are unlikely; if present, keep last occurrence.
    """
    def clean_cell(value: Any) -> str:
        if value is None:
            return ""
        return str(value).strip()

    def build_meta(source: Dict[str, Any] | None) -> Dict[str, str]:
        source = source or {}
        return {
            "program_name": clean_cell(source.get("program_name")),
            "vehicle_number": clean_cell(source.get("vehicle_number")),
            "serial_component": clean_cell(source.get("serial_component")),
        }

    rows: List[Tuple[str, Path, Dict[str, str]]] = []
    if REG_XLSX.exists():
        # Try pandas first
        try:
            import pandas as pd  # type: ignore
            df = pd.read_excel(REG_XLSX)
            for _, r in df.iterrows():
                row_dict = r.to_dict() if hasattr(r, "to_dict") else dict(r)
                sc = clean_cell(row_dict.get("serial_component") or row_dict.get("serial_number"))
                rf = clean_cell(row_dict.get("run_folder"))
                if sc and rf:
                    rows.append((sc, Path(rf), build_meta(row_dict)))
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
                sc_col = col_idx("serial_component") or col_idx("serial_number")
                rf_col = col_idx("run_folder")
                prog_col = col_idx("program_name")
                veh_col = col_idx("vehicle_number")
                data_col = col_idx("serial_component")
                if sc_col and rf_col:
                    for row in ws.iter_rows(min_row=2):
                        sc_val = row[sc_col - 1].value if sc_col else None
                        rf_val = row[rf_col - 1].value if rf_col else None
                        sc = clean_cell(sc_val)
                        rf = clean_cell(rf_val)
                        if sc and rf:
                            meta = {
                                "program_name": clean_cell(row[prog_col - 1].value) if prog_col else "",
                                "vehicle_number": clean_cell(row[veh_col - 1].value) if veh_col else "",
                                "serial_component": clean_cell(row[data_col - 1].value) if data_col else "",
                            }
                            rows.append((sc, Path(rf), meta))
                    if rows:
                        return rows
            except Exception:
                pass
    if REG_CSV.exists():
        try:
            with REG_CSV.open("r", encoding="utf-8", newline="") as f:
                r = csv.DictReader(f)
                for row in r:
                    sc = clean_cell(row.get("serial_component") or row.get("serial_number"))
                    rf = clean_cell(row.get("run_folder"))
                    if sc and rf:
                        rows.append((sc, Path(rf), build_meta(row)))
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


def build_master() -> Tuple[List[str], List[Dict[str, Any]], Dict[str, str], Dict[str, str], Dict[str, str]]:
    """Return serial list, term rows, and program/vehicle/data mappings."""
    reg = load_registry()
    if not reg:
        print("[WARN] No registry entries found. Nothing to compile.")
        return [], [], {}, {}, {}

    # Keep last occurrence per SN (registry is already latest-first, but be safe)
    last_for_sn: Dict[str, Tuple[Path, Dict[str, str]]] = {}
    for sn, rf, meta in reg:
        last_for_sn[sn] = (rf, meta or {})

    serials = list(last_for_sn.keys())
    terms_order: List[Tuple[str, str]] = []
    term_map: Dict[Tuple[str, str], Dict[str, Any]] = {}
    program_by_sn: Dict[str, str] = {}
    sv_by_sn: Dict[str, str] = {}
    data_by_sn: Dict[str, str] = {}

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
        """Return a normalized value for a term.

        If the extractor explicitly failed to find a value (found == False or
        an error_reason is present), return 'N/A' so the master workbook
        distinguishes between "not present in this EIDP" and "not yet run".
        """
        # Prefer any concrete extracted value
        for key in ("extracted_value", "number", "text", "string", "value"):
            if key in entry:
                val = entry.get(key)
                if val is None:
                    continue
                text = norm(val)
                if text != "":
                    return text
        # No concrete value: if the extractor ran but found nothing, mark N/A
        try:
            found_flag = entry.get("found", None)
        except Exception:
            found_flag = None
        if found_flag is False:
            return "N/A"
        err = norm(entry.get("error_reason"))
        if err:
            return "N/A"
        return None

    for sn, (rf, registry_meta) in last_for_sn.items():
        registry_meta = registry_meta or {}
        reg_prog = norm(registry_meta.get("program_name"))
        reg_sv = norm(registry_meta.get("vehicle_number"))
        reg_data = norm(registry_meta.get("serial_component"))
        rows = load_results_json(rf)
        if not rows:
            print(f"[WARN] No scan_results.json in {rf}")
            if reg_prog and sn not in program_by_sn:
                program_by_sn[sn] = reg_prog
            if reg_sv and sn not in sv_by_sn:
                sv_by_sn[sn] = reg_sv
            if reg_data and sn not in data_by_sn:
                data_by_sn[sn] = reg_data
            continue
        # Capture every row for this SN, grouped by displayed term/data group
        for row in rows:
            row_id = (row.get("serial_component") or row.get("serial_number") or "").strip()
            if row_id and row_id != sn:
                continue
            # Capture metadata per SN if present (or derive from filename)
            prog = norm(row.get("program_name") or row.get("program"))
            sv = norm(row.get("vehicle_number") or row.get("space_vehicle"))
            serial_component = norm(row.get("serial_component")) or reg_data
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
            if not prog and reg_prog:
                prog = reg_prog
            if not sv and reg_sv:
                sv = reg_sv
            if prog and sn not in program_by_sn:
                program_by_sn[sn] = prog
            if sv and sn not in sv_by_sn:
                sv_by_sn[sn] = sv
            if serial_component and sn not in data_by_sn:
                data_by_sn[sn] = serial_component
            term_label = norm(row.get("term_label") or row.get("term"))
            if not term_label:
                continue
            value = extract_value(row)
            if value is None:
                value = ""

            data_group = norm(row.get("data_group"))
            key = (term_label.lower(), data_group.lower())
            if key not in term_map:
                term_map[key] = {
                    "term_label": term_label,
                    "data_group": data_group,
                    "units": "",
                    "range_min": "",
                    "range_max": "",
                    "values": {},
                }
                terms_order.append(key)
            entry = term_map[key]
            units = extract_units(row)
            if units and not entry["units"]:
                entry["units"] = units
            rng_min = norm(row.get("range_min"))
            if rng_min and not entry["range_min"]:
                entry["range_min"] = rng_min
            rng_max = norm(row.get("range_max"))
            if rng_max and not entry["range_max"]:
                entry["range_max"] = rng_max
            entry["values"][sn] = value

    term_rows: List[Dict[str, Any]] = []
    for key in terms_order:
        info = term_map.get(key)
        if not info:
            continue
        term_rows.append({
            "term_label": info.get("term_label", ""),
            "data_group": info.get("data_group", ""),
            "units": info.get("units", ""),
            "range_min": info.get("range_min", ""),
            "range_max": info.get("range_max", ""),
            "values": info.get("values", {}),
        })

    return serials, term_rows, program_by_sn, sv_by_sn, data_by_sn


def write_master(serials: List[str], term_rows: List[Dict[str, Any]], program_by_sn: Dict[str, str] | None = None, sv_by_sn: Dict[str, str] | None = None, data_by_sn: Dict[str, str] | None = None) -> None:
    EXPORTS.mkdir(parents=True, exist_ok=True)
    base_columns = ["Term Label", "Data Group", "Units", "Min", "Max"]
    header = base_columns + serials

    def blank_meta(label: str) -> Dict[str, Any]:
        row = {col: "" for col in header}
        row["Term Label"] = label
        return row

    structured_rows: List[Dict[str, Any]] = [blank_meta("Program"), blank_meta("Space Vehicle"), blank_meta("Data")]

    # Fill Program / Space Vehicle rows per serial
    program_by_sn = program_by_sn or {}
    sv_by_sn = sv_by_sn or {}
    data_by_sn = data_by_sn or {}
    # Row 0 -> Program, Row 1 -> Space Vehicle, Row 2 -> Data
    for sn in serials:
        structured_rows[0][sn] = program_by_sn.get(sn, "")
        structured_rows[1][sn] = sv_by_sn.get(sn, "")
        structured_rows[2][sn] = data_by_sn.get(sn, "")

    for entry in term_rows:
        row: Dict[str, Any] = {col: "" for col in header}
        row["Term Label"] = entry.get("term_label", "")
        row["Data Group"] = entry.get("data_group", "")
        row["Units"] = entry.get("units", "")
        row["Min"] = entry.get("range_min", "")
        row["Max"] = entry.get("range_max", "")
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
            # Freeze header + metadata rows, and keep core columns visible.
            ws.freeze_panes(4, 5)
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
    serials, rows, prog_map, sv_map, data_map = build_master()
    if not serials:
        sys.exit(0)
    write_master(serials, rows, program_by_sn=prog_map, sv_by_sn=sv_map, data_by_sn=data_map)


if __name__ == "__main__":
    main()

