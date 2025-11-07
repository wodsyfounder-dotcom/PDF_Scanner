#!/usr/bin/env python3
"""
Generate a Smart-Snap-only Excel schema with:
- Sheet "Template" restricted to Smart mode columns
- Dropdowns for limited-choice fields
- A second header row with short explanations

Writes: user_inputs/terms.schema.smartsnap.xlsx
"""

from pathlib import Path
import sys


def main() -> None:
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, Alignment
        from openpyxl.worksheet.datavalidation import DataValidation
    except Exception:
        print("[ERROR] openpyxl is required to generate the template. Install with: py -m pip install openpyxl", file=sys.stderr)
        sys.exit(1)

    root = Path(__file__).resolve().parents[1]
    out_dir = root / 'user_inputs'
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / 'terms.schema.smartsnap.xlsx'

    wb = Workbook()
    ws = wb.active
    ws.title = 'Template'

    # Headers (Smart Snap only)
    headers = [
        'Data Group',         # A
        'Term Label',         # B
        'Term',               # C (search anchor)
        'Pages',              # D
        'Mode',               # E (locked to "smart")
        'Return',             # F number|string
        'Units',              # G numeric units hint(s)
        'Range (min)',        # H
        'Range (max)',        # I
        'Format',             # J optional mask or /regex/
        'GroupAfter',         # K
        'GroupBefore',        # L
        'Smart Snap Type',    # M auto|number|date|time|title
        'Secondary Term',     # N secondary label for disambiguation
        'Smart Position',     # O 1-based positional override
    ]
    ws.append(headers)

    # Explanatory second header row
    expl = [
        'User-defined category for reporting/rollups',
        'Friendly label shown in outputs (leave blank to reuse Term)',
        'Search term/anchor used by the scanner (required)',
        'Page ranges (e.g., 1 or 2-4; comma/semicolon supported)',
        'Fixed: smart',
        'Output type for value (number or string)',
        'Preferred units token(s) for numbers (pipe list ok)',
        'Lower numeric bound (optional)',
        'Upper numeric bound (optional)',
        'Optional value mask (e.g., tpl-xxxx) or /regex/',
        'Only search after this anchor text (optional)',
        'Stop searching once this anchor is reached (optional)',
        'auto|number|date|time|title (auto if blank)',
        'Optional qualifier expected near the chosen value',
        'Pick the Nth value to the right (1-based) - overrides scoring',
    ]
    ws.append(expl)

    # Style headers
    header_font = Font(bold=True)
    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        ws.column_dimensions[cell.column_letter].width = max(14, len(h) + 2)
    # Explanation row styling
    for col_idx in range(1, len(headers) + 1):
        c = ws.cell(row=2, column=col_idx)
        c.alignment = Alignment(wrap_text=True)

    # Data validations
    dv_mode = DataValidation(type="list", formula1='"smart"', allow_blank=False)
    dv_return = DataValidation(type="list", formula1='"number,string"', allow_blank=False)
    dv_smart = DataValidation(type="list", formula1='"auto,number,date,time,title"', allow_blank=True)
    ws.add_data_validation(dv_mode)
    ws.add_data_validation(dv_return)
    ws.add_data_validation(dv_smart)
    dv_mode.add('E3:E2000')
    dv_return.add('F3:F2000')
    dv_smart.add('M3:M2000')

    # Example rows
    examples = [
        ['Propulsion', 'Resistance (Pre-Test)', 'resistance', '1', 'smart', 'number', '', '5', '7', '', 'pre test', 'post test', 'number', '', ''],
        ['Operations', 'Test Date', 'test date', '1', 'smart', 'string', '', '', '', '', '', '', 'date', '', ''],
    ]
    for row in examples:
        ws.append(row)

    ws.freeze_panes = 'A3'

    # Instructions sheet
    inst = wb.create_sheet('Instructions')
    inst['A1'] = 'Smart Snap Terms Schema Template'
    inst['A1'].font = Font(bold=True)
    notes = [
        'This template configures ONLY Smart Snap extraction. Columns are reduced to what Smart Snap needs.',
        'Data Group and Term Label are for reporting only—they flow to run outputs and the master workbook.',
        'For numbers, Units and Range help scoring; for strings/dates, set Return and Smart Snap Type.',
        'Smart Position overrides scoring by selecting the Nth value to the right of the label.',
    ]
    for i, t in enumerate(notes, start=3):
        inst[f'A{i}'] = t
        inst[f'A{i}'].alignment = Alignment(wrap_text=True)
    inst.column_dimensions['A'].width = 110

    wb.save(out_path)
    print(f"[DONE] Wrote template -> {out_path}")


if __name__ == '__main__':
    main()

