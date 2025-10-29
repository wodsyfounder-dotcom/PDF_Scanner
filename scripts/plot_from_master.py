#!/usr/bin/env python3
"""
Generate plots based on user_inputs/plot_terms.xlsx and Product_Data_File/master.

Outputs PNG figures to Product_Data_File/plots/<plot_name>.png using matplotlib.

Behavior:
  - Builds plot definitions from rows where Plot? == 'Y' (case-insensitive) OR Tie To Plot is set.
  - A row with Plot? == 'Y' starts (or contributes to) a plot named by 'Plot Name'.
  - A row with 'Tie To Plot' adds a series to that named plot.
  - If both Plot Name and Tie To Plot are present, the row contributes to Plot Name
    and is also allowed to tie to the same name (Tie wins if provided).
  - X Axis (only read from rows that define/own a plot): 'SN' | 'Program' | 'Space Vehicle'.
  - Adds horizontal Min/Max reference lines if provided on any contributing row.
  - Y-axis label uses a common Units value if unique across series; otherwise units are appended to legend labels.
"""
from __future__ import annotations

import math
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Tuple, Any, Optional


# Use non-interactive backend
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402


ROOT = Path(__file__).resolve().parents[1]
EXPORTS = ROOT / "Product_Data_File"
PLOTS_DIR = EXPORTS / "plots"
MASTER_XLSX = EXPORTS / "master.xlsx"
MASTER_CSV = EXPORTS / "master.csv"
USER = ROOT / "user_inputs"
PLOT_TERMS_XLSX = USER / "plot_terms.xlsx"
PLOT_TERMS_CSV = USER / "plot_terms.csv"


ID_COLS = ["Term", "Grouping", "Row Label", "Column Label", "Units"]


def read_master() -> Tuple[List[str], List[Dict[str, Any]], Dict[str, str], Dict[str, str]]:
    """Return (serials, data_rows, program_by_sn, sv_by_sn) from master workbook.

    Skips first two metadata rows in data_rows.
    """
    import pandas as pd  # type: ignore
    programs: Dict[str, str] = {}
    vehicles: Dict[str, str] = {}

    if MASTER_XLSX.exists():
        df = pd.read_excel(MASTER_XLSX, sheet_name="master")
    elif MASTER_CSV.exists():
        df = pd.read_csv(MASTER_CSV)
    else:
        print("[ERROR] No master workbook found. Compile master first.")
        return [], [], {}, {}

    serials = [c for c in df.columns if c not in ("Term", "Grouping", "Units", "Row Label", "Column Label")]
    # Extract program/vehicle rows
    if len(df) >= 2:
        row_programs = df.iloc[0]
        row_vehicles = df.iloc[1]
        for sn in serials:
            programs[sn] = str(row_programs.get(sn) or "").strip()
            vehicles[sn] = str(row_vehicles.get(sn) or "").strip()
    data_rows = df.iloc[2:].fillna("").to_dict(orient="records")
    return serials, data_rows, programs, vehicles


def read_plot_terms() -> List[Dict[str, Any]]:
    import pandas as pd  # type: ignore
    path = PLOT_TERMS_XLSX if PLOT_TERMS_XLSX.exists() else PLOT_TERMS_CSV
    if not path.exists():
        print("[ERROR] plot_terms workbook not found. Generate it first.")
        return []
    if path.suffix.lower() == ".xlsx":
        df = pd.read_excel(path)
    else:
        df = pd.read_csv(path)
    df = df.fillna("")
    return df.to_dict(orient="records")


@dataclass
class SeriesSpec:
    id_tuple: Tuple[str, str, str, str, str]
    series_label: str
    units: str
    bounds_min: Optional[float] = None
    bounds_max: Optional[float] = None


@dataclass
class PlotSpec:
    name: str
    x_axis: str  # SN | Program | Space Vehicle
    series: List[SeriesSpec] = field(default_factory=list)


def slugify(name: str) -> str:
    s = re.sub(r"[^A-Za-z0-9_. -]+", "_", name.strip())
    s = re.sub(r"[\s]+", "_", s)
    return s[:150] or "plot"


def build_plot_specs(rows: List[Dict[str, Any]]) -> List[PlotSpec]:
    # Collect plots by name
    plots: Dict[str, PlotSpec] = {}

    # First pass: any row with Plot? == 'Y' starts/declares a plot
    for r in rows:
        plot_flag = str(r.get("Plot?") or "").strip().lower()
        if plot_flag not in ("y", "yes", "1", "true"):  # only declared ones here
            continue
        pname = str(r.get("Plot Name") or "").strip() or "Plot"
        x_axis = str(r.get("X Axis") or "SN").strip()
        if pname not in plots:
            plots[pname] = PlotSpec(name=pname, x_axis=x_axis)

    # Second pass: add series based on either declared plot or tie-to-plot
    def to_float(v) -> Optional[float]:
        s = str(v).strip()
        if not s:
            return None
        try:
            return float(s.replace(",", ""))
        except Exception:
            return None

    for r in rows:
        plot_flag = str(r.get("Plot?") or "").strip().lower()
        pname = str(r.get("Plot Name") or "").strip()
        tie = str(r.get("Tie To Plot") or "").strip()
        if not pname and not tie and plot_flag not in ("y", "yes", "1", "true"):
            continue
        # Determine target plot name
        target = tie or pname
        if not target:
            continue
        if target not in plots:
            # permit tying to a new name
            plots[target] = PlotSpec(name=target, x_axis=str(r.get("X Axis") or "SN").strip())
        key = tuple(str(r.get(k) or "").strip() for k in ID_COLS)
        series_label = str(r.get("Series Label") or "").strip()
        units = str(r.get("Units") or "").strip()
        s = SeriesSpec(
            id_tuple=key,
            series_label=series_label or "series",
            units=units,
            bounds_min=to_float(r.get("Min")),
            bounds_max=to_float(r.get("Max")),
        )
        plots[target].series.append(s)

    # Drop empty plots
    out = [p for p in plots.values() if p.series]
    # Sort series per plot for stability
    for p in out:
        p.series.sort(key=lambda s: s.series_label.lower())
    # Stable sort plots by name
    out.sort(key=lambda p: p.name.lower())
    return out


def extract_series(master_rows: List[Dict[str, Any]], serials: List[str], spec: SeriesSpec) -> Tuple[List[int], List[float]]:
    # Find matching row in master
    tgt = {k: v for k, v in zip(ID_COLS, spec.id_tuple)}
    row = next((r for r in master_rows if all(str(r.get(k) or "").strip() == tgt[k] for k in ID_COLS)), None)
    if row is None:
        return [], []
    xs = list(range(len(serials)))
    ys: List[float] = []
    for sn in serials:
        raw = row.get(sn)
        if raw is None:
            ys.append(math.nan)
            continue
        s = str(raw).strip()
        if not s:
            ys.append(math.nan)
            continue
        # Try float conversion; tolerate commas
        try:
            y = float(s.replace(",", ""))
        except Exception:
            y = math.nan
        ys.append(y)
    return xs, ys


def pick_y_units(series_list: List[SeriesSpec]) -> Tuple[Optional[str], List[str]]:
    uniq = []
    seen = set()
    for s in series_list:
        u = (s.units or "").strip()
        k = u.lower()
        if k in seen:
            continue
        seen.add(k)
        if u:
            uniq.append(u)
    if len(uniq) == 1:
        return uniq[0], uniq
    return None, uniq


def main() -> None:
    serials, master_rows, program_by_sn, vehicle_by_sn = read_master()
    if not serials:
        sys.exit(1)
    rows = read_plot_terms()
    if not rows:
        sys.exit(1)
    plots = build_plot_specs(rows)
    if not plots:
        print("[WARN] No plots selected. Toggle 'Plot?' or set 'Tie To Plot'.")
        sys.exit(0)

    PLOTS_DIR.mkdir(parents=True, exist_ok=True)

    for p in plots:
        # Build x tick labels by axis preference
        if p.x_axis.lower() == "program":
            xticklabels = [program_by_sn.get(sn, sn) or sn for sn in serials]
            xlabel = "Program"
        elif p.x_axis.lower() in ("space vehicle", "vehicle", "sv"):
            xticklabels = [vehicle_by_sn.get(sn, sn) or sn for sn in serials]
            xlabel = "Space Vehicle"
        else:
            xticklabels = serials
            xlabel = "Serial Number"

        fig, ax = plt.subplots(figsize=(10, 5.5), constrained_layout=True)
        y_label, _all_units = pick_y_units(p.series)
        cmap = plt.get_cmap("tab10")

        for i, s in enumerate(p.series):
            xs, ys = extract_series(master_rows, serials, s)
            if not xs:
                print(f"[WARN] Series not found in master: {s.series_label}")
                continue
            label = s.series_label
            if y_label is None:
                u = (s.units or "").strip()
                if u:
                    label = f"{label} ({u})"
            color = cmap(i % 10)
            ax.plot(xs, ys, marker="o", linestyle="-", label=label, color=color)
            # Bounds lines per series if provided
            if s.bounds_min is not None and not math.isnan(s.bounds_min):
                ax.axhline(s.bounds_min, color=color, linestyle="--", linewidth=1, alpha=0.6)
            if s.bounds_max is not None and not math.isnan(s.bounds_max):
                ax.axhline(s.bounds_max, color=color, linestyle=":", linewidth=1, alpha=0.6)

        ax.set_title(p.name)
        ax.set_xlabel(xlabel)
        if y_label:
            ax.set_ylabel(y_label)
        ax.set_xticks(list(range(len(xticklabels))))
        ax.set_xticklabels(xticklabels, rotation=20, ha="right")
        ax.grid(True, which="major", axis="both", alpha=0.25)
        ax.legend(loc="best", fontsize=9)

        out_path = PLOTS_DIR / f"{slugify(p.name)}.png"
        fig.savefig(out_path, dpi=160)
        plt.close(fig)
        print(f"[DONE] Plot -> {out_path}")


if __name__ == "__main__":
    main()

