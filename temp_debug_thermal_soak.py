#!/usr/bin/env python3
"""
Debug script to investigate thermal_soak_high Smart Position issue for OCR.
"""
import sys
import os
import importlib.util

# Load the scanner module using importlib
spec = importlib.util.spec_from_file_location('scanner', 'Application/eidp_term_scanner.core.py')
scanner = importlib.util.module_from_spec(spec)
spec.loader.exec_module(scanner)

# Enable debug mode
os.environ['DEBUG_MODE'] = '1'
os.environ['OCR_MODE'] = 'easyocr'
os.environ['OCR_ROW_EPS'] = '2'

# Set up paths
pdf_path = r'C:\Users\zachs\Documents\DevProjects\PDF_Scanner\Data Packages\FakeProgram_SV1_SN1111.pdf'
terms_file = r'C:\Users\zachs\Documents\DevProjects\PDF_Scanner\user_inputs\terms.schema.smartsnap.xlsx'

print("Loading scanner...")
print(f"PDF: {pdf_path}")
print(f"Terms: {terms_file}")
print()

# Load the scanner's internal functions to get OCR boxes
try:
    # Get OCR boxes for page 3 (where thermal soak should be)
    boxes = scanner._get_easyocr_boxes_page(pdf_path, 3, dpi=500, langs=['en'])
    print(f"\nFound {len(boxes)} OCR boxes on page 3:")

    # Filter for boxes containing "thermal" and "soak"
    thermal_boxes = [b for b in boxes if 'thermal' in b.get('text', '').lower() and 'soak' in b.get('text', '').lower()]
    print(f"\nBoxes with 'thermal' and 'soak': {len(thermal_boxes)}")
    for b in thermal_boxes:
        print(f"  {b.get('text')!r:30} x0={b.get('x0'):.1f} x1={b.get('x1'):.1f} y0={b.get('y0'):.1f} y1={b.get('y1'):.1f}")

    # Get all boxes near "Thermal Soak High" specifically
    thermal_high_boxes = [b for b in thermal_boxes if 'high' in b.get('text', '').lower()]
    if thermal_high_boxes:
        ref_box = thermal_high_boxes[0]
        print(f"\nFocusing on 'Thermal Soak High' at y={ref_box.get('y0'):.1f}-{ref_box.get('y1'):.1f}")
        ref_y = (ref_box['y0'] + ref_box['y1']) / 2

        # Try different EPS values to see where the data is
        for eps in [2, 5, 10, 20, 50]:
            same_row = [b for b in boxes if abs((b['y0'] + b['y1'])/2 - ref_y) <= eps]
            same_row = sorted(same_row, key=lambda b: b['x0'])

            print(f"\nWith EPS={eps}:")
            print(f"  Found {len(same_row)} boxes in row")
            if len(same_row) > 1:
                for b in same_row[:10]:  # Show first 10
                    print(f"    {b.get('text')!r:30} x0={b.get('x0'):.1f} y={(b.get('y0') + b.get('y1'))/2:.1f}")

        # Now use the actual EPS value from config
        eps = 2  # OCR_ROW_EPS
        print(f"\nUsing configured EPS={eps}:")

        same_row = [b for b in boxes if abs((b['y0'] + b['y1'])/2 - ref_y) <= eps]
        same_row = sorted(same_row, key=lambda b: b['x0'])

        print(f"\nAll boxes in same row as 'Thermal Soak' (y={ref_y:.1f}, eps={eps}):")
        for b in same_row:
            print(f"  {b.get('text')!r:30} x0={b.get('x0'):.1f} x1={b.get('x1'):.1f}")

        # Show which would be to the right of the label
        label_right_x = ref_box['x1']
        right_boxes = [b for b in same_row if b['x0'] >= label_right_x]
        print(f"\nBoxes to the right of label (x >= {label_right_x:.1f}):")
        for i, b in enumerate(right_boxes):
            print(f"  Position {i}: {b.get('text')!r:30} x0={b.get('x0'):.1f}")

except Exception as e:
    print(f"Error: {e}")
    import traceback
    traceback.print_exc()
