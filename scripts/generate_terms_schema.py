#!/usr/bin/env python3
"""
Generate a user-friendly Excel schema with dropdowns and conditional formatting:
- Sheet "Template" with headers + validation lists and greying of non-applicable fields by Mode/Return
- Sheet "Instructions" describing each column and examples

Writes: user_inputs/terms.schema.xlsx
"""
from pathlib import Path
import sys


def main() -> None:
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.worksheet.datavalidation import DataValidation
        from openpyxl.formatting.rule import FormulaRule
    except Exception:
        print("[ERROR] openpyxl is required to generate the template. Install with: py -m pip install openpyxl", file=sys.stderr)
        sys.exit(1)

    root = Path(__file__).resolve().parents[1]
    out_dir = root / 'user_inputs'
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / 'terms.schema.xlsx'

    wb = Workbook()
    ws = wb.active
    ws.title = 'Template'

    # Headers
    headers = [
        'Term', 'Pages', 'Mode', 'Line (x)', 'Column (y)',
        'Anchor', 'FieldIndex', 'FieldSplit', 'Return',
        'Units', 'Range (min)', 'Range (max)'
    ]
    ws.append(headers)
    header_font = Font(bold=True)
    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        ws.column_dimensions[cell.column_letter].width = max(12, len(h) + 2)

    # Data validation lists
    dv_mode = DataValidation(type="list", formula1='"nearest,table(xy),line"', allow_blank=True)
    dv_split = DataValidation(type="list", formula1='"auto,groups,tokens"', allow_blank=True)
    dv_return = DataValidation(type="list", formula1='"number,string"', allow_blank=True)
    ws.add_data_validation(dv_mode)
    ws.add_data_validation(dv_split)
    ws.add_data_validation(dv_return)
    dv_mode.add('C2:C2000')
    dv_split.add('H2:H2000')
    dv_return.add('I2:I2000')

    # Conditional formatting to grey non-applicable fields
    grey = PatternFill(start_color='00DDDDDD', end_color='00DDDDDD', fill_type='solid')
    # If Mode <> "table(xy)", grey Line (D) and Column (E)
    ws.conditional_formatting.add('D2:D2000', FormulaRule(formula=["$C2<>\"table(xy)\""], fill=grey))
    ws.conditional_formatting.add('E2:E2000', FormulaRule(formula=["$C2<>\"table(xy)\""], fill=grey))
    # If Mode <> "line", grey Anchor (F), FieldIndex (G), FieldSplit (H)
    for col in ('F','G','H'):
        ws.conditional_formatting.add(f'{col}2:{col}2000', FormulaRule(formula=["$C2<>\"line\""], fill=grey))
    # If Return <> "number", grey Units (J), Range min(K), Range max(L)
    for col in ('J','K','L'):
        ws.conditional_formatting.add(f'{col}2:{col}2000', FormulaRule(formula=["$I2<>\"number\""], fill=grey))

    # Example rows
    examples = [
        ['Thrust', '1', 'table(xy)', 'Thrust', 'Value|Nominal', '', '', 'groups', 'number', 'lbf', '200', '3000'],
        ['Title', '1', 'line', '', '', 'Title:', '2', 'auto', 'string', '', '', ''],
        ['Ignition Temperature', '1-2', 'nearest', '', '', '', '', '', 'number', 'K|degC', '200', '5000'],
    ]
    for row in examples:
        ws.append(row)

    ws.freeze_panes = 'A2'

    # Instructions
    inst = wb.create_sheet('Instructions')
    inst['A1'] = 'EIDP Terms Schema Template'
    inst['A1'].font = Font(bold=True)
    lines = [
        'Fill rows in the Template sheet. Rows appear in outputs in the same order.',
        '',
        'Columns:',
        ' - Term: Friendly row name (required).',
        ' - Pages: Page ranges (1-indexed), e.g., 5-10; 22 (required).',
        ' - Mode: nearest | table(xy) | line (defaults to nearest).',
        '   * table(xy): Use Line (row header) and Column (column header; allow Value|Nominal).',
        '   * line     : Use Anchor (defaults to Term), FieldIndex (1-based), FieldSplit (auto|groups|tokens), Return (number|string).',
        '   * nearest  : No extra fields; optional Units and Range (min/max).',
        '',
        'Header aliases:',
        ' - XY headers accepted as "Line (x)" and "Column (y)".',
        ' - Units: Preferred units (pipe list) for number return.',
        ' - Range (min)/(max): Numeric bounds (scientific notation allowed).',
    ]
    for i, t in enumerate(lines, start=2):
        inst[f'A{i}'] = t
        inst[f'A{i}'].alignment = Alignment(wrap_text=True)
    inst.column_dimensions['A'].width = 110

    wb.save(out_path)
    print(f"[DONE] Wrote template -> {out_path}")


if __name__ == '__main__':
    main()

