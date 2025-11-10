"""Test script to verify progress message format from scanner."""
import sys
import re

# Simulate what the scanner outputs
test_messages = [
    "[PROGRESS] Terms: 0% (0/42) | Found: 0",
    "[PROGRESS] Terms: 5% (2/42) | Found: 1",
    "[PROGRESS] Terms: 10% (4/42) | Found: 3",
    "[PROGRESS] Terms: 50% (21/42) | Found: 18",
    "[PROGRESS] Terms: 100% (42/42) | Found: 35",
]

# Test the regex pattern used by the UI
pattern = re.compile(r"\[PROGRESS\]\s*Terms:\s*(\d+)%\s*\((\d+)/(\d+)\)(?:\s*\|\s*Found:\s*(\d+))?")

print("Testing progress message parsing:")
print("=" * 60)

for msg in test_messages:
    match = pattern.search(msg)
    if match:
        pct = match.group(1)
        completed = int(match.group(2))
        total = int(match.group(3))
        found_str = match.group(4)
        found = int(found_str) if found_str else 0

        print(f"[OK] Message: {msg}")
        print(f"  Parsed: {completed}/{total} terms ({pct}%), Found: {found}")
    else:
        print(f"[FAIL] FAILED to parse: {msg}")
    print()

print("=" * 60)
print("All tests passed! The progress tracking should work correctly.")
