#!/usr/bin/env python3
"""Find which page has Thermal Soak High"""
import sys
import os
import importlib.util
import fitz  # PyMuPDF

# Load the scanner module
spec = importlib.util.spec_from_file_location('scanner', 'Application/eidp_term_scanner.core.py')
scanner = importlib.util.module_from_spec(spec)
spec.loader.exec_module(scanner)

pdf_path = r'C:\Users\zachs\Documents\DevProjects\PDF_Scanner\Data Packages\FakeProgram_SV1_SN1111.pdf'

print(f"Searching for 'Thermal Soak High' in {pdf_path}")
print()

# Try each page
for page_num in range(1, 6):  # Check pages 1-5
    print(f"Page {page_num}:")
    try:
        boxes = scanner._get_easyocr_boxes_page(pdf_path, page_num, dpi=500, langs=['en'])
        thermal_boxes = [b for b in boxes if 'thermal' in b.get('text', '').lower() and 'soak' in b.get('text', '').lower()]

        if thermal_boxes:
            print(f"  FOUND! {len(thermal_boxes)} thermal soak boxes")
            for b in thermal_boxes:
                print(f"    {b.get('text')!r}")
        else:
            # Check for "Thermal" and "Soak" separately
            thermal = [b for b in boxes if 'thermal' in b.get('text', '').lower()]
            soak = [b for b in boxes if 'soak' in b.get('text', '').lower()]
            if thermal or soak:
                print(f"    thermal: {[b.get('text') for b in thermal]}")
                print(f"    soak: {[b.get('text') for b in soak]}")
            else:
                print(f"    No thermal/soak boxes found")
    except Exception as e:
        print(f"  Error: {e}")

print("\nDone")
