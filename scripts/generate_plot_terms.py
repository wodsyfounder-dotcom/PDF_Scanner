#!/usr/bin/env python3
"""
Generate or refresh a plotting configuration sheet (plot_terms.xlsx).

Source: Product_Data_File/master.xlsx (preferred) or master.csv
Output: user_inputs/plot_terms.xlsx (preferred) or plot_terms.csv

The sheet contains one row per unique measurement row in the master table:
  - Plot? (Y/N): toggle to include this row for plotting
  - Plot Name: name for a new plot this row starts (auto-suggested)
  - Tie To Plot: optional, tie this row as an additional series to an
                 existing plot name
  - X Axis: SN | Program | Space Vehicle (only considered for rows that
            define a new plot via Plot?=Y and a Plot Name)
  - Term, Grouping, Row Label, Column Label, Units: identifiers
  - Min, Max: optional numeric bounds for horizontal reference lines
  - Series Label: legend label override for this row (optional)

If plot_terms.xlsx already exists, preserves user edits by merging on the
identifier columns (Term, Grouping, Row Label, Column Label, Units).
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


def read_master() -> Tuple[List[str], List[Dict[str, Any]]]:
    """Return (serials, rows) from master workbook.

    rows are dicts with keys: Term, Grouping, Units, Row Label, Column Label, and each SN column.
    Skips the first two metadata rows (Program, Space Vehicle).
    """
    if MASTER_XLSX.exists():
        try:
            import pandas as pd  # type: ignore
            df = pd.read_excel(MASTER_XLSX, sheet_name="master")
            serials = [c for c in df.columns if c not in ("Term", "Grouping", "Units", "Row Label", "Column Label")]
            # Skip metadata rows
            df_rows = df.iloc[2:].fillna("")
            records = df_rows.to_dict(orient="records")
            return serials, records
        except Exception:
            pass
    if MASTER_CSV.exists():
        try:
            import pandas as pd  # type: ignore
            df = pd.read_csv(MASTER_CSV)
            serials = [c for c in df.columns if c not in ("Term", "Grouping", "Units", "Row Label", "Column Label")]
            df_rows = df.iloc[2:].fillna("")
            records = df_rows.to_dict(orient="records")
            return serials, records
        except Exception:
            pass
    print("[ERROR] No master workbook found. Compile master first.")
    return [], []


ID_COLS = ["Term", "Grouping", "Row Label", "Column Label", "Units"]


def default_plot_name(row: Dict[str, Any]) -> str:
    g = str(row.get("Grouping") or "").strip()
    t = str(row.get("Term") or "").strip()
    if g and t:
        return f"{g} - {t}"
    return t or g or "Plot"


def default_series_label(row: Dict[str, Any]) -> str:
    t = str(row.get("Term") or "").strip()
    r = str(row.get("Row Label") or "").strip()
    c = str(row.get("Column Label") or "").strip()
    parts = [t]
    if r and r.lower() not in ("value",):
        parts.append(r)
    if c:
        parts.append(c)
    return " - ".join([p for p in parts if p]) or t or "series"


def load_existing() -> Dict[Tuple[str, str, str, str, str], Dict[str, Any]]:
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
        out: Dict[Tuple[str, str, str, str, str], Dict[str, Any]] = {}
        for _, r in df.iterrows():
            key = tuple(str(r.get(k) or "").strip() for k in ID_COLS)
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
        "Min",
        "Max",
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
    if not serials:
        sys.exit(1)

    # Build unique identifier rows
    seen: set[Tuple[str, str, str, str, str]] = set()
    candidates: List[Dict[str, Any]] = []
    for row in master_rows:
        term = str(row.get("Term") or "").strip()
        if not term or term.lower() in ("program", "space vehicle"):
            continue
        key = tuple(str(row.get(k) or "").strip() for k in ID_COLS)
        if key in seen:
            continue
        seen.add(key)
        candidates.append({
            **{k: row.get(k, "") for k in ID_COLS},
        })

    existing = load_existing()
    merged: List[Dict[str, Any]] = []
    for base in candidates:
        key = tuple(str(base.get(k) or "").strip() for k in ID_COLS)
        prev = existing.get(key, {})
        plot_flag = str(prev.get("Plot?") or "").strip()
        plot_name = str(prev.get("Plot Name") or "").strip() or default_plot_name(base)
        tie = str(prev.get("Tie To Plot") or "").strip()
        x_axis = str(prev.get("X Axis") or "").strip() or "SN"
        min_v = prev.get("Min") if prev.get("Min") not in (None, "") else ""
        max_v = prev.get("Max") if prev.get("Max") not in (None, "") else ""
        series_label = str(prev.get("Series Label") or "").strip() or default_series_label(base)
        merged.append({
            "Plot?": plot_flag,
            "Plot Name": plot_name,
            "Tie To Plot": tie,
            "X Axis": x_axis,
            **base,
            "Min": min_v,
            "Max": max_v,
            "Series Label": series_label,
        })

    write_sheet(merged)


if __name__ == "__main__":
    main()
