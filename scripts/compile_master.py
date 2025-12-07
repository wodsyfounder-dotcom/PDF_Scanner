#!/usr/bin/env python3
"""
Compile a master extraction workbook from the per-run caches recorded in run_registry.

Inputs:
  - Product_Data_File/run_registry.xlsx (preferred) or run_registry.csv
  - For each row: serial_component, run_folder, program_name, vehicle_number
  - Each run folder contains scan_results.json which holds per-term values per SN.

Output:
  - Product_Data_File/master.xlsx (requires pandas + xlsxwriter)

Workbook layout:
  - Leading rows reserved for "Program", "Space Vehicle", and "Data" metadata.
  - Columns include Term Label, Data Group, Units, Min, Max, followed by one column per serial number.

Requires pandas and xlsxwriter packages.
"""
from __future__ import annotations

import csv
import re
import json
import sys
import math
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Any, Set


ROOT = Path(__file__).resolve().parents[1]
EXPORTS = ROOT / "Product_Data_File"
REG_XLSX = EXPORTS / "run_registry.xlsx"
REG_CSV = EXPORTS / "run_registry.csv"
OUT_XLSX = EXPORTS / "master.xlsx"
OUT_CSV = EXPORTS / "master.csv"
TERMS_XLSX = ROOT / "user_inputs" / "terms.schema.smartsnap.xlsx"
TERMS_SHEET = "Template"


# ===== Intelligent Term Ordering Configuration =====

# Data group ordering rules: (priority, keywords)
# Lower priority = earlier in the list (front-loaded)
DATA_GROUP_ORDER_RULES = [
    (1, ["document", "profile", "info", "label", "id", "metadata", "component info"]),  # Identification
    (2, ["pre", "initial", "baseline", "before", "as-received"]),  # Pre-test
    (3, ["environmental", "test setup", "conditions", "ambient", "thermal", "vib"]),  # Test conditions
    (4, ["performance", "kpi", "functional", "acceptance", "results", "output"]),  # Performance/Results
    (5, ["calibration", "trim"]),  # Calibration (contains both pre and post)
    (6, ["post", "final", "after"]),  # Post-test
    (7, ["quality", "status", "compliance", "verification"]),  # Quality/Status
]


def _categorize_data_group(data_group: str) -> int:
    """
    Categorize a data group into priority buckets for intelligent ordering.
    Lower number = earlier in the list (front-loaded).

    Returns priority (1-7), with 4 as default for uncategorized groups.
    """
    dg_lower = data_group.lower()

    for priority, keywords in DATA_GROUP_ORDER_RULES:
        if any(kw in dg_lower for kw in keywords):
            return priority

    # Default: uncategorized groups go in the middle (after pre-test, before post-test)
    return 4


def _sort_data_group_key(data_group: str) -> tuple:
    """
    Generate sort key for data groups.
    Returns (priority, alphabetical_name).
    """
    priority = _categorize_data_group(data_group)
    return (priority, data_group.lower())


def _extract_base_metric_name(term_label: str) -> str:
    """
    Extract the base metric name from a term label to group related terms.

    Examples:
    - "Thrust Nominal Value" → "thrust nominal"
    - "Thrust Nominal Target" → "thrust nominal"
    - "Power Draw Peak" → "power draw"
    - "LVDT-2 Pre-Trim" → "lvdt"
    """
    tl_lower = term_label.lower()

    # Remove common suffixes/modifiers
    suffixes = [
        'value', 'target', 'confidence', 'tolerance', 'actual',
        'units', 'quality', 'status', 'source', 'over',
        'pre-trim', 'post-trim', 'pre trim', 'post trim',
        'high', 'low', 'max', 'min',
        'duration', 'amplitude', 'path',
    ]

    # Remove numbered component suffixes (e.g., "-2", "-02")
    base = re.sub(r'[-_]\d+', '', tl_lower).strip()

    # Remove known suffixes
    for suffix in suffixes:
        if base.endswith(suffix):
            base = base[:-len(suffix)].strip()

    # Remove "nominal" if it's a modifier (e.g., "thrust nominal" → "thrust")
    if base.endswith('nominal'):
        base = base[:-7].strip()

    return base if base else tl_lower


def _sort_term_label_key(term_label: str) -> tuple:
    """
    Generate sort key for term labels within a data group.

    Ordering rules:
    1. Base metric name (groups related terms together)
    2. Numerical component prefixes (TC-01, TC-02, LVDT-1, LVDT-2)
    3. Pre before Post (for paired measurements)
    4. Value → Target → Confidence → Units (metric hierarchy)
    5. High → Low, Max → Min (range ordering)
    6. Main field before Status/Quality
    7. Alphabetical as final tiebreaker
    """
    tl_lower = term_label.lower()

    # Extract base metric name to group related terms
    base_name = _extract_base_metric_name(term_label)

    # Extract numerical component prefix for sorting (e.g., "TC-02" → 2, "LVDT-1" → 1)
    component_num = 0
    match = re.search(r'[-_](\d+)', term_label)
    if match:
        component_num = int(match.group(1))

    # Pre/Post ordering (Pre=0, neither=1, Post=2)
    pre_post_order = 0
    if 'pre' in tl_lower:
        pre_post_order = 0
    elif 'post' in tl_lower:
        pre_post_order = 2
    else:
        pre_post_order = 1

    # Metric hierarchy ordering
    metric_order = 1
    if 'value' in tl_lower or 'actual' in tl_lower:
        metric_order = 0
    elif 'target' in tl_lower:
        metric_order = 1
    elif 'confidence' in tl_lower or 'tolerance' in tl_lower:
        metric_order = 2
    elif 'units' in tl_lower:
        metric_order = 3

    # Status/Quality goes last (within the base metric group)
    status_order = 1 if any(kw in tl_lower for kw in ['status', 'quality']) else 0

    # Range ordering (High before Low)
    range_order = 0
    if 'high' in tl_lower or 'max' in tl_lower:
        range_order = 0
    elif 'low' in tl_lower or 'min' in tl_lower:
        range_order = 1

    # Sort priority: base_name, component_num, status (last!), pre_post, metric, range, alphabetical
    return (base_name, component_num, status_order, pre_post_order, metric_order, range_order, tl_lower)


def norm(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    return str(value).strip()


def _clean_master_cell(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    return str(value).strip()


def _read_existing_master() -> Tuple[List[str], List[Dict[str, str]], Optional[float]]:
    """Return (header, rows, mtime) from an existing master workbook if present."""
    target: Optional[Path] = None
    if OUT_XLSX.exists():
        target = OUT_XLSX
    elif OUT_CSV.exists():
        target = OUT_CSV
    if not target:
        return [], [], None

    mtime = target.stat().st_mtime
    header: List[str] = []
    rows: List[Dict[str, str]] = []
    suffix = target.suffix.lower()

    if suffix == ".xlsx":
        try:
            import pandas as pd  # type: ignore
            df = pd.read_excel(target, dtype=object, keep_default_na=False)
            header = [str(col) for col in df.columns]
            raw_rows = df.to_dict(orient="records")
            for record in raw_rows:
                cleaned = {str(k): _clean_master_cell(v) for k, v in record.items()}
                rows.append(cleaned)
            return header, rows, mtime
        except Exception:
            try:
                from openpyxl import load_workbook  # type: ignore
                wb = load_workbook(str(target), data_only=True)
                ws = wb.active
                first_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
                header = [(_clean_master_cell(val) or f"column_{idx}") for idx, val in enumerate(first_row)]
                for values in ws.iter_rows(min_row=2, values_only=True):
                    record: Dict[str, str] = {}
                    for idx, val in enumerate(values):
                        if idx >= len(header):
                            continue
                        record[header[idx]] = _clean_master_cell(val)
                    rows.append(record)
                wb.close()
                return header, rows, mtime
            except Exception:
                pass

    if suffix == ".csv":
        try:
            with target.open("r", encoding="utf-8", newline="") as f:
                reader = csv.DictReader(f)
                header = reader.fieldnames or []
                for row in reader:
                    cleaned = {str(k): _clean_master_cell(v) for k, v in row.items()}
                    rows.append(cleaned)
            return header, rows, mtime
        except Exception:
            pass

    return [], [], mtime


def _load_schema_units(path: Path = TERMS_XLSX) -> Dict[Tuple[str, str], str]:
    """Return {(term_label, data_group): units} from the Smart-Snap schema."""
    mapping: Dict[Tuple[str, str], str] = {}
    if not path.exists():
        return mapping
    try:
        import pandas as _pd  # type: ignore

        df = _pd.read_excel(path, sheet_name=TERMS_SHEET)
        df = df.fillna("")
        for _, row in df.iterrows():
            term = str(row.get("Term Label") or "").strip()
            if not term:
                continue
            group = str(row.get("Data Group") or "").strip()
            units = str(row.get("Units") or "").strip()
            if not units:
                continue
            mapping[(term.lower(), group.lower())] = units
        return mapping
    except Exception:
        try:
            from openpyxl import load_workbook  # type: ignore

            wb = load_workbook(str(path), data_only=True)
            ws = wb[TERMS_SHEET] if TERMS_SHEET in wb.sheetnames else wb.active
            header = [str(v or "").strip() for v in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
            lookup: Dict[str, int] = {name: idx for idx, name in enumerate(header)}
            def val(row_vals, name: str) -> str:
                idx = lookup.get(name)
                if idx is None or idx >= len(row_vals):
                    return ""
                return str(row_vals[idx] or "").strip()
            for row in ws.iter_rows(min_row=2, values_only=True):
                term = val(row, "Term Label")
                if not term:
                    continue
                group = val(row, "Data Group")
                units = val(row, "Units")
                if not units:
                    continue
                mapping[(term.lower(), group.lower())] = units
            wb.close()
        except Exception:
            return mapping
    return mapping


def _parse_run_datetime(run_dir: Path) -> datetime:
    """Parse run folder name as timestamp; fallback to filesystem mtime."""
    try:
        return datetime.strptime(run_dir.name, "%Y%m%d_%H%M%S")
    except Exception:
        return datetime.fromtimestamp(run_dir.stat().st_mtime)


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
    """Return serial list, term rows, and program/vehicle/data mappings.

    This implementation seeds from any existing master.xlsx so user edits /
    overrides persist. Only run_data folders newer than the master file are
    applied, and they only fill blank cells—non-blank cells are treated as
    canonical (user-entered or previously accepted values).
    """
    base_columns = ["Term Label", "Data Group", "Units", "Min", "Max"]
    schema_units = _load_schema_units()

    # --- Seed from existing master (if present) so user edits persist
    header_existing, rows_existing, master_mtime = _read_existing_master()
    serials: List[str] = [col for col in header_existing if col not in base_columns]
    terms_order: List[Tuple[str, str]] = []
    term_map: Dict[Tuple[str, str], Dict[str, Any]] = {}
    program_by_sn: Dict[str, str] = {}
    sv_by_sn: Dict[str, str] = {}
    data_by_sn: Dict[str, str] = {}

    for row in rows_existing:
        term_label = norm(row.get("Term Label"))
        if not term_label:
            continue
        tl_lower = term_label.lower()
        data_group = norm(row.get("Data Group"))
        if tl_lower == "program":
            for sn in serials:
                val = norm(row.get(sn))
                if val:
                    program_by_sn[sn] = val
            continue
        if tl_lower == "space vehicle":
            for sn in serials:
                val = norm(row.get(sn))
                if val:
                    sv_by_sn[sn] = val
            continue
        if tl_lower == "data":
            for sn in serials:
                val = norm(row.get(sn))
                if val:
                    data_by_sn[sn] = val
            continue

        key = (tl_lower, data_group.lower())
        if key not in term_map:
            term_map[key] = {
                "term_label": term_label,
                "data_group": data_group,
                "units": norm(row.get("Units")),
                "range_min": norm(row.get("Min")),
                "range_max": norm(row.get("Max")),
                "values": {},
            }
            terms_order.append(key)
        entry = term_map[key]
        entry_units = norm(row.get("Units"))
        if entry_units and not entry["units"]:
            entry["units"] = entry_units
        rng_min = norm(row.get("Min"))
        if rng_min and not entry["range_min"]:
            entry["range_min"] = rng_min
        rng_max = norm(row.get("Max"))
        if rng_max and not entry["range_max"]:
            entry["range_max"] = rng_max
        for sn in serials:
            entry["values"][sn] = norm(row.get(sn))

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

    reg = load_registry()
    meta_by_sn: Dict[str, Dict[str, str]] = {}
    for sn, _rf, meta in reg:
        meta_by_sn[sn] = meta or {}

    runs_seen = False
    runs_root = EXPORTS / "run_data"
    if runs_root.exists():
        for run_dir in sorted(p for p in runs_root.iterdir() if p.is_dir()):
            # Process all run_data folders regardless of timestamp. This ensures
            # that if a user deletes a value from master.xlsx (making it blank),
            # it will be refilled from ANY available run_data on the next compile.
            # Since we only fill blank cells (line 494), existing values are safe.
            rows = load_results_json(run_dir)
            if not rows:
                continue
            runs_seen = True
            for row in rows:
                sn = norm(row.get("serial_component") or row.get("serial_number"))
                if not sn:
                    continue
                if sn not in serials:
                    serials.append(sn)

                meta = meta_by_sn.get(sn, {}) or {}
                reg_prog = norm(meta.get("program_name"))
                reg_sv = norm(meta.get("vehicle_number"))
                reg_data = norm(meta.get("serial_component"))

                # Capture metadata per SN if present (or derive from filename)
                prog = norm(row.get("program_name") or row.get("program"))
                sv = norm(row.get("vehicle_number") or row.get("space_vehicle"))
                serial_component = norm(row.get("serial_component")) or reg_data

                if not prog:
                    pdf_file = norm(row.get("pdf_file"))
                    stem = Path(pdf_file).stem if pdf_file else ""
                    if stem:
                        if "_" in stem:
                            parts = [p.strip() for p in stem.split("_") if p.strip()]
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
                if prog and not program_by_sn.get(sn):
                    program_by_sn[sn] = prog
                if sv and not sv_by_sn.get(sn):
                    sv_by_sn[sn] = sv
                if serial_component and not data_by_sn.get(sn):
                    data_by_sn[sn] = serial_component

                term_label = norm(row.get("term_label") or row.get("term"))
                if not term_label:
                    continue
                value = extract_value(row)
                if value is None:
                    value = ""
                elif value == "N/A":
                    data_group_preview = norm(row.get("data_group"))
                    print(f"[INFO] N/A recorded for {sn} :: {term_label} [{data_group_preview}]")

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
                if not units and key in schema_units:
                    units = schema_units.get(key, "")
                if units and not entry["units"]:
                    entry["units"] = units
                rng_min = norm(row.get("range_min"))
                if rng_min and not entry["range_min"]:
                    entry["range_min"] = rng_min
                rng_max = norm(row.get("range_max"))
                if rng_max and not entry["range_max"]:
                    entry["range_max"] = rng_max

                # ALWAYS OVERRIDE with latest scan result
                # Newest scanned values replace old ones
                entry["values"][sn] = value

    # Apply intelligent ordering to terms
    # Group by data_group, sort groups intelligently, then sort terms within each group
    def _sort_term_key(key: Tuple[str, str]) -> tuple:
        term_label_lower, data_group_lower = key
        # Get the actual term_label and data_group from term_map
        info = term_map.get(key, {})
        term_label = info.get("term_label", term_label_lower)
        data_group = info.get("data_group", data_group_lower)
        # Sort by: (data_group_priority, data_group_name, term_label_priority)
        return _sort_data_group_key(data_group) + _sort_term_label_key(term_label)

    terms_order = sorted(terms_order, key=_sort_term_key)

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

    # Ensure serials present in registry are preserved even if no new runs were ingested.
    for sn, meta in meta_by_sn.items():
        if sn not in serials:
            serials.append(sn)
        if not program_by_sn.get(sn):
            program_by_sn[sn] = norm(meta.get("program_name"))
        if not sv_by_sn.get(sn):
            sv_by_sn[sn] = norm(meta.get("vehicle_number"))
        if not data_by_sn.get(sn):
            data_by_sn[sn] = norm(meta.get("serial_component"))

    if not serials and not term_rows and not runs_seen and not rows_existing:
        print("[WARN] No registry entries or run_data found. Nothing to compile.")
        return [], [], {}, {}, {}

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

    # Write to Excel (required - no CSV fallback)
    try:
        import pandas as pd  # type: ignore
        import xlsxwriter  # noqa: F401
    except ImportError as e:
        raise RuntimeError(
            "pandas and xlsxwriter are required to create master.xlsx. "
            "Install them with: pip install pandas xlsxwriter"
        ) from e

    try:
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
    except Exception as e:
        raise RuntimeError(f"Failed to write master.xlsx: {e}") from e


def build_master_from_state() -> Tuple[List[str], List[Dict[str, Any]], Dict[str, str], Dict[str, str], Dict[str, str]]:
    """
    Rebuild master workbook from master_cell_state.json, respecting existing master.xlsx.

    Master.xlsx is the source of truth for which terms should exist.
    - If master.xlsx exists: only include terms that are present in it
    - If master.xlsx doesn't exist: include all terms from cell state
    This ensures user deletions from master.xlsx are respected.

    Returns: (serials, term_rows, program_by_sn, sv_by_sn, data_by_sn)
    """
    from scripts.master_cell_state import load_cell_state

    # Load cell state (single source of truth for extracted values)
    state = load_cell_state()

    if not state:
        print("[WARN] Cell state is empty. No data to compile.")
        return [], [], {}, {}, {}

    # Load schema to get term metadata (units, min, max, data_group)
    schema_units = _load_schema_units()

    # Load schema to get full term details
    schema_terms = _load_schema_terms()

    # Get serial components from state
    serials = sorted(state.keys())

    # Read existing master.xlsx to determine which terms should be included
    # Master.xlsx is the source of truth - if user deleted a row, respect it
    header_existing, rows_existing, master_mtime = _read_existing_master()
    allowed_terms_set: Optional[Set[str]] = None

    if rows_existing:
        # Extract (term_label, data_group) keys from existing master
        # Skip the first 3 metadata rows (Program, Space Vehicle, Data)
        allowed_terms_set = set()
        for row in rows_existing[3:]:  # Skip metadata rows
            term_label = norm(row.get("Term Label", ""))
            data_group = norm(row.get("Data Group", ""))
            if term_label:
                # Store term_name (from schema) - we'll match by term_label
                allowed_terms_set.add(term_label.lower())
        print(f"[INFO] Filtering to {len(allowed_terms_set)} terms from existing master.xlsx")
    else:
        print("[INFO] No existing master.xlsx found - including all terms from cell state")

    # Collect all unique terms across all serial components
    all_terms_set = set()
    for serial_component, terms in state.items():
        all_terms_set.update(terms.keys())

    # Filter terms based on existing master.xlsx (if it exists)
    if allowed_terms_set is not None:
        # Match terms by their term_label from schema
        filtered_terms = set()
        for term_name in all_terms_set:
            term_meta = schema_terms.get(term_name.lower(), {})
            term_label = term_meta.get("term_label", term_name)
            if term_label.lower() in allowed_terms_set:
                filtered_terms.add(term_name)
        all_terms_set = filtered_terms
        print(f"[INFO] Filtered to {len(all_terms_set)} terms present in master.xlsx")

    # Apply intelligent ordering to terms
    # Group by data_group, sort groups intelligently, then sort terms within each group
    def _sort_term_name_key(term_name: str) -> tuple:
        # Get term metadata from schema
        term_meta = schema_terms.get(term_name.lower(), {})
        term_label = term_meta.get("term_label", term_name)
        data_group = term_meta.get("data_group", "")
        # Sort by: (data_group_priority, data_group_name, term_label_priority)
        return _sort_data_group_key(data_group) + _sort_term_label_key(term_label)

    terms_order = sorted(all_terms_set, key=_sort_term_name_key)

    # Build term_rows (organized by term)
    term_rows: List[Dict[str, Any]] = []

    for term_name in terms_order:
        # Look up term metadata from schema
        term_meta = schema_terms.get(term_name.lower(), {})

        # Build values dict for this term across all serial components
        values = {}
        for serial_component in serials:
            cell_data = state.get(serial_component, {}).get(term_name)
            if cell_data:
                values[serial_component] = cell_data.get("value")
            else:
                values[serial_component] = ""

        # Create term row
        row = {
            "term_label": term_meta.get("term_label", term_name),
            "data_group": term_meta.get("data_group", ""),
            "units": term_meta.get("units", ""),
            "range_min": term_meta.get("range_min", ""),
            "range_max": term_meta.get("range_max", ""),
            "values": values,
        }
        term_rows.append(row)

    # Get metadata (program_name, vehicle_number) from run_registry
    # since cell state doesn't track metadata
    program_by_sn: Dict[str, str] = {}
    sv_by_sn: Dict[str, str] = {}
    data_by_sn: Dict[str, str] = {}

    try:
        reg = load_registry()
        for sn, _rf, meta in reg:
            if sn in serials:
                program_by_sn[sn] = meta.get("program_name", "")
                sv_by_sn[sn] = meta.get("vehicle_number", "")
                data_by_sn[sn] = meta.get("serial_component", sn)
    except Exception as e:
        print(f"[WARN] Could not load metadata from registry: {e}")

    print(f"[INFO] Built master from state: {len(serials)} serial components, {len(term_rows)} terms")
    return serials, term_rows, program_by_sn, sv_by_sn, data_by_sn


def _load_schema_terms() -> Dict[str, Dict[str, str]]:
    """
    Load term metadata from schema file.
    Returns: {term_label_lower: {term_label, data_group, units, range_min, range_max}}
    """
    mapping: Dict[str, Dict[str, str]] = {}
    if not TERMS_XLSX.exists():
        return mapping

    try:
        import pandas as pd  # type: ignore
        df = pd.read_excel(TERMS_XLSX, sheet_name=TERMS_SHEET)
        df = df.fillna("")

        for _, row in df.iterrows():
            term_label = str(row.get("Term Label") or "").strip()
            if not term_label:
                continue

            mapping[term_label.lower()] = {
                "term_label": term_label,
                "data_group": str(row.get("Data Group") or "").strip(),
                "units": str(row.get("Units") or "").strip(),
                "range_min": str(row.get("Min") or "").strip(),
                "range_max": str(row.get("Max") or "").strip(),
            }
        return mapping
    except Exception:
        # Fallback to openpyxl
        try:
            from openpyxl import load_workbook  # type: ignore
            wb = load_workbook(str(TERMS_XLSX), data_only=True)
            ws = wb[TERMS_SHEET] if TERMS_SHEET in wb.sheetnames else wb.active

            # Parse header
            header = [str(v or "").strip() for v in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
            lookup = {name: idx for idx, name in enumerate(header)}

            def val(row_vals, name: str) -> str:
                idx = lookup.get(name)
                if idx is None or idx >= len(row_vals):
                    return ""
                return str(row_vals[idx] or "").strip()

            for row in ws.iter_rows(min_row=2, values_only=True):
                term_label = val(row, "Term Label")
                if not term_label:
                    continue

                mapping[term_label.lower()] = {
                    "term_label": term_label,
                    "data_group": val(row, "Data Group"),
                    "units": val(row, "Units"),
                    "range_min": val(row, "Min"),
                    "range_max": val(row, "Max"),
                }

            wb.close()
        except Exception as e:
            print(f"[WARN] Could not load schema terms: {e}")

        return mapping


def main() -> None:
    serials, rows, prog_map, sv_map, data_map = build_master()
    if not serials:
        sys.exit(0)
    write_master(serials, rows, program_by_sn=prog_map, sv_by_sn=sv_map, data_by_sn=data_map)


if __name__ == "__main__":
    main()

