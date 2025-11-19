#!/usr/bin/env python3
"""
Test _fields_from_items to see if it's creating extra fields
"""
import sys
import importlib.util

# Load scanner module
spec = importlib.util.spec_from_file_location('scanner', 'Application/eidp_term_scanner.core.py')
scanner = importlib.util.module_from_spec(spec)
spec.loader.exec_module(scanner)

# Simulate ordered_right_items after filtering (should NOT include label)
ordered_right_items = [
    {'text': '+95', 'x0': 1210.0, 'x1': 1325.0, 'y0': 975.0, 'y1': 1005.0},
    {'text': '1.5', 'x0': 1596.0, 'x1': 1698.0, 'y0': 977.0, 'y1': 1007.0},
    {'text': 'NO', 'x0': 2056.0, 'x1': 2206.0, 'y0': 975.0, 'y1': 1005.0},
    {'text': 'A-118', 'x0': 3461.0, 'x1': 3654.0, 'y0': 977.0, 'y1': 1007.0},
]

print("ordered_right_items:", [it['text'] for it in ordered_right_items])
print()

# Call _fields_from_items
fields = scanner._fields_from_items(ordered_right_items)

print(f"fields_from_items returned {len(fields)} fields:")
for i, field in enumerate(fields, start=1):
    print(f"  Position {i}: {field!r}")

print()
print("Expected:")
print("  Position 1: '+95'")
print("  Position 2: '1.5' (Duration)")
print("  Position 3: 'NO' (Status)")
print("  Position 4: 'A-118'")
