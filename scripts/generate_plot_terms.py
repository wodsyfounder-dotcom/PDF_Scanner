#!/usr/bin/env python3
"""
Generate or refresh a plotting configuration sheet (plot_terms.xlsx).

Source: Product_Data_File/master.xlsx
Output: user_inputs/plot_terms.xlsx

Each row in the sheet corresponds to a unique (Term Label, Data Group)
combination emitted by the master workbook compiled from run data.
For every combination we provide:
  - Plot? (Y/N) toggle to include the row in downstream plotting
  - Plot Name / Tie To Plot / X Axis controls for chart generation
  - Term Label, Data Group, Units, Min, Max identifiers copied from master
  - Series Label default (can be edited per-row)

If plot_terms.xlsx already exists, preserves user edits by merging on the
identifier columns (Term Label + Data Group). Legacy sheets that still use
Term/Grouping will be translated automatically when possible.
"""
from __future__ import annotations

import sys
from pathlib import Path
from typing import List, Dict, Any, Tuple


ROOT = Path(__file__).resolve().parents[1]
EXPORTS = ROOT / "Product_Data_File"
MASTER_XLSX = EXPORTS / "master.xlsx"
MASTER_CSV = EXPORTS / "master.csv"
USER = ROOT / "user_inputs"
OUT_XLSX = USER / "plot_terms.xlsx"
OUT_CSV = USER / "plot_terms.csv"


PREFERRED_AXIS = ("SN", "Program", "Space Vehicle")
Y_AXIS_COL = "Y Axis Label"


BASE_COLS = ["Term Label", "Data Group", "Units", "Min", "Max"]
LEGACY_STATIC = ["Term", "Grouping", "Row Label", "Column Label", "Units"]
ID_COLS = ["Term Label", "Data Group"]
_META_LABELS = {"program", "space vehicle", "data"}


def _prepare_master_df(df: "pd.DataFrame") -> Tuple[List[str], List[Dict[str, Any]]]:  # type: ignore[name-defined]
    """Normalize master columns and drop metadata rows."""
    import pandas as pd  # type: ignore

    rename_map: Dict[str, str] = {}
    if "Term Label" not in df.columns and "Term" in df.columns:
        rename_map["Term"] = "Term Label"
    if "Data Group" not in df.columns and "Grouping" in df.columns:
        rename_map["Grouping"] = "Data Group"
    if rename_map:
        df = df.rename(columns=rename_map)

    for col in BASE_COLS:
        if col not in df.columns:
            df[col] = ""

    df = df.fillna("")
    term_series = df["Term Label"].astype(str).str.strip()
    mask_meta = term_series.str.lower().isin(_META_LABELS)
    mask_blank = term_series == ""
    df = df[~(mask_meta | mask_blank)]
    if df.empty:
        return [], []

    exclude = set(BASE_COLS) | set(LEGACY_STATIC)
    serials = [c for c in df.columns if c not in exclude]
    df = df.reset_index(drop=True)
    records = df.to_dict(orient="records")
    return serials, records


def read_master() -> Tuple[List[str], List[Dict[str, Any]]]:
    """Return (serials, rows) from master workbook (Term Label/Data Group schema)."""
    try:
        import pandas as pd  # type: ignore
    except Exception:
        print("[ERROR] pandas is required to read the master workbook.")
        return [], []

    if not MASTER_XLSX.exists():
        print("[ERROR] master.xlsx not found. Compile master first.")
        return [], []

    try:
        df = pd.read_excel(MASTER_XLSX, sheet_name="master")
        serials, records = _prepare_master_df(df)
        return serials, records
    except Exception as e:
        print(f"[ERROR] Failed to read master.xlsx: {e}")
        return [], []


def default_plot_name(row: Dict[str, Any]) -> str:
    term = str(row.get("Term Label") or "").strip()
    group = str(row.get("Data Group") or "").strip()
    if term and group:
        return f"{term} - {group}"
    return term or group or "Plot"


def default_series_label(row: Dict[str, Any]) -> str:
    term = str(row.get("Term Label") or "").strip()
    group = str(row.get("Data Group") or "").strip()
    parts = [p for p in (term, group) if p]
    return " - ".join(parts) or term or "series"


def _row_key(row: Dict[str, Any]) -> Tuple[str, str] | None:
    """Return normalized (Term Label, Data Group) tuple or None if empty."""
    key = tuple(str(row.get(k) or "").strip() for k in ID_COLS)
    if any(key):
        return key  # type: ignore[return-value]
    # Legacy fallback: Term + Grouping
    legacy_term = str(row.get("Term") or "").strip()
    legacy_group = str(row.get("Grouping") or "").strip()
    if not legacy_group:
        legacy_group = str(row.get("Row Label") or "").strip()
    if legacy_term or legacy_group:
        return (legacy_term, legacy_group)
    return None


def load_existing() -> Dict[Tuple[str, str], Dict[str, Any]]:
    """Return mapping from ID tuple to existing row, if a prior sheet exists."""
    import pandas as pd  # type: ignore
    path = OUT_XLSX if OUT_XLSX.exists() else OUT_CSV
    if not path.exists():
        return {}
    try:
        if path.suffix.lower() == ".xlsx":
            df = pd.read_excel(path)
        else:
            df = pd.read_csv(path)
        df = df.fillna("")
        out: Dict[Tuple[str, str], Dict[str, Any]] = {}
        for _, r in df.iterrows():
            key = _row_key(r)
            if not key:
                continue
            out[key] = {k: r.get(k) for k in df.columns}
        return out
    except Exception:
        return {}


def write_sheet(rows: List[Dict[str, Any]]) -> None:
    import pandas as pd  # type: ignore
    USER.mkdir(parents=True, exist_ok=True)
    cols = [
        "Plot?",
        "Plot Name",
        "Tie To Plot",
        "X Axis",
        *ID_COLS,
        "Units",
        "Min",
        "Max",
        Y_AXIS_COL,
        "Series Label",
    ]
    df = pd.DataFrame(rows, columns=cols)
    # Write Excel only (xlsxwriter preferred, openpyxl fallback)
    try:
        import xlsxwriter  # noqa: F401
        with pd.ExcelWriter(OUT_XLSX, engine="xlsxwriter") as writer:
            df.to_excel(writer, sheet_name="plot_terms", index=False)
            ws = writer.sheets["plot_terms"]
            ws.freeze_panes(1, 1)
            # Auto-size columns
            for i, col in enumerate(df.columns):
                try:
                    max_len = int(df[col].astype(str).map(len).max()) if not df.empty else len(col)
                except Exception:
                    max_len = len(col)
                ws.set_column(i, i, min(60, max(12, max_len + 2)))
            # Data validations for toggles and axis
            last_row = max(1, len(df) + 1)
            ws.data_validation(1, 0, last_row, 0, {"validate": "list", "source": ["", "Y", "N"]})
            ws.data_validation(1, 3, last_row, 3, {"validate": "list", "source": list(PREFERRED_AXIS)})
        if OUT_CSV.exists():
            try:
                OUT_CSV.unlink()
            except Exception:
                pass
        print(f"[DONE] Plot terms workbook -> {OUT_XLSX}")
        return
    except Exception:
        # Fallback: write Excel via openpyxl engine (no validations)
        try:
            with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="plot_terms", index=False)
            if OUT_CSV.exists():
                try:
                    OUT_CSV.unlink()
                except Exception:
                    pass
            print(f"[DONE] Plot terms workbook -> {OUT_XLSX}")
            return
        except Exception as e2:
            raise SystemExit(f"[ERROR] Could not write plot_terms.xlsx: {e2}")


def main() -> None:

    serials, master_rows = read_master()
    if not serials and not master_rows:
        sys.exit(1)

    # Build unique identifier rows
    seen: set[Tuple[str, str]] = set()
    candidates: List[Dict[str, Any]] = []
    for row in master_rows:
        term_label = str(row.get("Term Label") or "").strip()
        data_group = str(row.get("Data Group") or "").strip()
        if not term_label:
            continue
        key = (term_label, data_group)
        if key in seen:
            continue
        seen.add(key)
        candidates.append({
            "Term Label": term_label,
            "Data Group": data_group,
            "Units": str(row.get("Units") or "").strip(),
            "Min": str(row.get("Min") or "").strip(),
            "Max": str(row.get("Max") or "").strip(),
        })

    existing = load_existing()
    merged: List[Dict[str, Any]] = []
    for base in candidates:
        key = (base["Term Label"], base["Data Group"])
        prev = existing.get(key, {})
        plot_flag = str(prev.get("Plot?") or "").strip()
        plot_name = str(prev.get("Plot Name") or "").strip() or default_plot_name(base)
        tie = str(prev.get("Tie To Plot") or "").strip()
        x_axis = str(prev.get("X Axis") or "").strip() or "SN"
        units = str(prev.get("Units") or "").strip() or base.get("Units", "")
        min_v = str(prev.get("Min") or "").strip() or base.get("Min", "")
        max_v = str(prev.get("Max") or "").strip() or base.get("Max", "")
        axis_label = str(prev.get(Y_AXIS_COL) or "").strip()
        if not axis_label:
            axis_label = units
        series_label = str(prev.get("Series Label") or "").strip() or default_series_label(base)
        merged.append({
            "Plot?": plot_flag,
            "Plot Name": plot_name,
            "Tie To Plot": tie,
            "X Axis": x_axis,
            "Term Label": base.get("Term Label", ""),
            "Data Group": base.get("Data Group", ""),
            "Units": units,
            "Min": min_v,
            "Max": max_v,
            Y_AXIS_COL: axis_label,
            "Series Label": series_label,
        })

    write_sheet(merged)


if __name__ == "__main__":
    main()
