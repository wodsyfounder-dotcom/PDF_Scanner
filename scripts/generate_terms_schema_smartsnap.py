#!/usr/bin/env python3
"""
Generate the Smart-Snap terms schema workbook with helpful validation, formatting,
and instructions. Output: user_inputs/terms.schema.smartsnap.xlsx
"""

from __future__ import annotations

from pathlib import Path
import sys


HEADERS = [
    "Data Group",
    "Term Label",
    "Term",
    "Pages",
    "Mode",
    "Line",
    "Column",
    "Anchor",
    "FieldIndex",
    "FieldSplit",
    "Return",
    "Units",
    "Range (min)",
    "Range (max)",
    "Format",
    "GroupAfter",
    "GroupBefore",
    "Smart Snap Type",
    "Secondary Term",
    "Smart Position",
]

MODE_OPTIONS = ["smart", "nearest", "line", "table(xy)", "full table"]
RETURN_OPTIONS = ["number", "string"]
SPLIT_OPTIONS = ["auto", "groups", "tokens"]
SMART_TYPES = ["", "auto", "number", "date", "time", "title"]


def main() -> None:
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.worksheet.datavalidation import DataValidation
        from openpyxl.formatting.rule import FormulaRule
    except Exception:
        print(
            "[ERROR] openpyxl is required to generate the Smart-Snap schema. "
            "Install with: py -m pip install openpyxl",
            file=sys.stderr,
        )
        sys.exit(1)

    root = Path(__file__).resolve().parents[1]
    out_dir = root / "user_inputs"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / "terms.schema.smartsnap.xlsx"

    wb = Workbook()
    ws = wb.active
    if ws is None:
        ws = wb.create_sheet()
    ws.title = "Template"

    header_font = Font(bold=True)
    ws.append(HEADERS)
    for idx, header in enumerate(HEADERS, start=1):
        cell = ws.cell(row=1, column=idx)
        cell.font = header_font
        ws.column_dimensions[cell.column_letter].width = max(14, len(header) + 2)

    # Data validations
    def _dv(values: list[str]) -> DataValidation:
        return DataValidation(type="list", formula1=f'"{",".join(values)}"', allow_blank=True)

    dv_mode = _dv(MODE_OPTIONS)
    dv_return = _dv(RETURN_OPTIONS)
    dv_split = _dv(SPLIT_OPTIONS)
    dv_smart = _dv([v or " " for v in SMART_TYPES])

    for dv in (dv_mode, dv_return, dv_split, dv_smart):
        ws.add_data_validation(dv)
    dv_mode.add("E2:E2000")
    dv_return.add("K2:K2000")
    dv_split.add("J2:J2000")
    dv_smart.add("R2:R2000")

    # Conditional formatting greys-out irrelevant fields
    grey = PatternFill(start_color="00E6E6E6", end_color="00E6E6E6", fill_type="solid")
    ws.conditional_formatting.add("F2:G2000", FormulaRule(formula=['$E2<>"table(xy)"'], fill=grey))
    ws.conditional_formatting.add("H2:J2000", FormulaRule(formula=['$E2<>"line"'], fill=grey))
    ws.conditional_formatting.add("K2:M2000", FormulaRule(formula=['$K2<>"number"'], fill=grey))
    ws.conditional_formatting.add("N2:N2000", FormulaRule(formula=['$E2="table(xy)"'], fill=grey))

    # Example rows
    examples = [
        ["Performance", "Pre-Test ISP", "isp", "1", "smart", "", "", "isp", "", "", "number", "sec", "40", "44", "", "Pre Test", "Post Test", "number", "", ""],
        ["Metadata", "Title", "title", "1", "line", "", "", "Title:", "2", "groups", "string", "", "", "", "", "", "", "title", "", ""],
        ["Tables", "Proof Load", "proof load", "2", "table(xy)", "Proof Load", "Value", "", "", "", "number", "psi", "", "", "", "Acceptance", "", "number", "", ""],
    ]
    for row in examples:
        ws.append(row)

    ws.freeze_panes = "A2"

    inst = wb.create_sheet("Instructions")
    inst["A1"] = "Smart-Snap Terms Schema"
    inst["A1"].font = Font(bold=True, size=13)
    instructions = [
        "Fill rows in the Template sheet. Rows are processed in order.",
        "",
        "Required:",
        " - Term: Friendly name for outputs.",
        " - Pages: One or more ranges (1-indexed). e.g., 5-7; 10; 12-13",
        "",
        "Key columns:",
        " - Mode: smart, nearest, line, table(xy), full table.",
        " - Line/Column: Row/column labels when Mode=table(xy).",
        " - Anchor/FieldIndex/FieldSplit: Used for Mode=line. FieldIndex is 1-based.",
        " - Return: number or string.",
        " - Range min/max & Units apply when Return=number.",
        " - Format: Optional mask or /regex/ for string validation.",
        " - GroupAfter / GroupBefore: Limit search window between these anchors.",
        " - Smart Snap Type/Secondary Term/Smart Position: Controls Smart-Snap resolution.",
        "",
        "Tips:",
        " - Leave columns blank when not applicable; conditional shading shows which ones are ignored.",
        " - Use consistent Term Labels/Data Groups to segment dashboards.",
        " - Save the file as .xlsx; the scanner also accepts .csv with the same headers.",
    ]
    for idx, line in enumerate(instructions, start=3):
        cell = inst[f"A{idx}"]
        cell.value = line
        cell.alignment = Alignment(wrap_text=True)
    inst.column_dimensions["A"].width = 110

    wb.save(out_path)
    print(f"[DONE] Smart-Snap schema written -> {out_path}")


if __name__ == "__main__":
    main()
