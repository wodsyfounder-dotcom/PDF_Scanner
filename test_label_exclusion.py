#!/usr/bin/env python3
"""
Test script to verify label box exclusion from right_items
"""

# Simulate the row_items we'd see for Thermal Soak High
row_items = [
    {'text': 'Thermal Soak High', 'x0': 534.0, 'x1': 1137.0, 'y0': 938.0, 'y1': 1025.0},
    {'text': '+95', 'x0': 1210.0, 'x1': 1325.0, 'y0': 975.0, 'y1': 1005.0},
    {'text': '1.5', 'x0': 1596.0, 'x1': 1698.0, 'y0': 977.0, 'y1': 1007.0},
    {'text': 'NO', 'x0': 2056.0, 'x1': 2206.0, 'y0': 975.0, 'y1': 1005.0},
    {'text': 'A-118', 'x0': 3461.0, 'x1': 3654.0, 'y0': 977.0, 'y1': 1007.0},
]

# Simulate finding the label via span matching
span = (0, 0)  # Label is at index 0
_, j = span
label_box_index = j
label_right_x = float(row_items[j]['x1'])

print(f"Label box index: {label_box_index}")
print(f"Label right x: {label_right_x}")
print()

# OLD METHOD (buggy):
print("OLD METHOD (before fix):")
right_items_old = [it for it in row_items if float(it.get('x0',0.0)) >= label_right_x - 1.0]
print(f"  right_items: {[it['text'] for it in right_items_old]}")
print(f"  Positions: {list(enumerate([it['text'] for it in right_items_old], start=1))}")
print()

# NEW METHOD (with fix):
print("NEW METHOD (after fix):")
right_items_new = [it for i, it in enumerate(row_items)
                   if float(it.get('x0',0.0)) >= label_right_x - 1.0
                   and (label_box_index is None or i != label_box_index)]
print(f"  right_items: {[it['text'] for it in right_items_new]}")
print(f"  Positions: {list(enumerate([it['text'] for it in right_items_new], start=1))}")
print()

print("Expected:")
print("  Position 1: '+95'")
print("  Position 2: '1.5' (Duration)")
print("  Position 3: 'NO' (Status)")
print("  Position 4: 'A-118'")
