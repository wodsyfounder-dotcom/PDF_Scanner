#!/usr/bin/env python3
"""
Build an Excel summary workbook that embeds generated plot images.

Inputs:
  - Product_Data_File/plots/*.png (produced by plot_from_master.py)

Output:
  - Product_Data_File/plots_summary.xlsx (preferred; xlsxwriter)
    Fallback: Product_Data_File/plots_summary.xlsx via openpyxl if xlsxwriter unavailable.

Behavior:
  - One sheet per plot image, sheet named from the plot filename (sanitized, <=31 chars).
  - Adds an Index sheet with links to each plot sheet.
  - If no plots present, prints a warning and exits 0.
"""
from __future__ import annotations

import os
import re
import sys
from pathlib import Path
from typing import List, Tuple


ROOT = Path(__file__).resolve().parents[1]
EXPORTS = ROOT / "Product_Data_File"
PLOTS_DIR = EXPORTS / "plots"
OUT_XLSX = EXPORTS / "plots_summary.xlsx"


def list_plot_images(plots_dir: Path) -> List[Path]:
    if not plots_dir.exists():
        return []
    imgs = [p for p in sorted(plots_dir.iterdir()) if p.suffix.lower() in (".png", ".jpg", ".jpeg")]
    return imgs


def sheet_name_from_file(path: Path, used: set[str]) -> str:
    base = path.stem.strip() or "Plot"
    # Sanitize to Excel sheet name constraints
    name = re.sub(r"[\\/*?:\[\]]+", "_", base)
    # Trim to 31 characters
    name = name[:31] if len(name) > 31 else name
    if not name:
        name = "Plot"
    orig = name
    i = 2
    while name in used:
        suffix = f"_{i}"
        name = (orig[: max(0, 31 - len(suffix))] + suffix)[:31]
        i += 1
    used.add(name)
    return name


def write_with_xlsxwriter(images: List[Path]) -> None:
    import pandas as pd  # type: ignore
    import xlsxwriter  # noqa: F401

    EXPORTS.mkdir(parents=True, exist_ok=True)
    used: set[str] = set()
    # Build index data
    rows: List[Tuple[str, str]] = []
    # Create workbook
    with pd.ExcelWriter(OUT_XLSX, engine="xlsxwriter") as writer:
        ws_index = writer.book.add_worksheet("Index")
        writer.sheets["Index"] = ws_index
        # Header
        ws_index.write(0, 0, "Plot")
        ws_index.write(0, 1, "Sheet")
        # Per-plot sheets
        for idx, img in enumerate(images, start=1):
            sheet = sheet_name_from_file(img, used)
            ws = writer.book.add_worksheet(sheet)
            writer.sheets[sheet] = ws
            # Put title and image
            ws.write(0, 0, img.stem)
            # Leave a row, insert image starting at A3
            ws.insert_image(2, 0, str(img), {"x_scale": 1.0, "y_scale": 1.0})
            # Index row with hyperlink to sheet
            ws_index.write(idx, 0, img.name)
            ws_index.write_url(idx, 1, f"internal:'{sheet}'!A1", string=sheet)
        # Autofit index columns
        ws_index.set_column(0, 0, 50)
        ws_index.set_column(1, 1, 24)
    print(f"[DONE] Plot summary workbook -> {OUT_XLSX}")


def write_with_openpyxl(images: List[Path]) -> None:
    from openpyxl import Workbook  # type: ignore
    from openpyxl.drawing.image import Image as XLImage  # type: ignore

    EXPORTS.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    # Use active as Index
    ws_index = wb.active
    ws_index.title = "Index"
    ws_index.cell(row=1, column=1, value="Plot")
    ws_index.cell(row=1, column=2, value="Sheet")

    used: set[str] = set()
    for idx, img in enumerate(images, start=2):
        sheet = sheet_name_from_file(img, used)
        ws = wb.create_sheet(title=sheet)
        ws.cell(row=1, column=1, value=img.stem)
        try:
            xlimg = XLImage(str(img))
            ws.add_image(xlimg, "A3")
        except Exception:
            pass
        ws_index.cell(row=idx, column=1, value=img.name)
        # No internal hyperlink API here; write sheet name
        ws_index.cell(row=idx, column=2, value=sheet)

    # Remove default empty sheet if present and unused
    # (openpyxl created active already used as Index)
    wb.save(OUT_XLSX)
    print(f"[DONE] Plot summary workbook -> {OUT_XLSX}")


def main() -> int:
    imgs = list_plot_images(PLOTS_DIR)
    if not imgs:
        print(f"[WARN] No plot images found in {PLOTS_DIR}. Generate plots first.")
        return 0
    # Try xlsxwriter path
    try:
        write_with_xlsxwriter(imgs)
        return 0
    except Exception as e:
        print(f"[WARN] xlsxwriter path unavailable ({e}); trying openpyxl")
    try:
        write_with_openpyxl(imgs)
        return 0
    except Exception as e:
        print(f"[ERROR] Could not write Excel summary: {e}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

