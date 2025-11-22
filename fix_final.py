#!/usr/bin/env python3
import sys
sys.stdout.reconfigure(encoding='utf-8')

filepath = r"c:\Users\zachs\Documents\DevProjects\PDF_Scanner\Application\eidp_term_scanner.core.py"
backup = r"c:\Users\zachs\Documents\DevProjects\PDF_Scanner\Application\eidp_term_scanner.core.py.backup"

# Read file
with open(filepath, 'r', encoding='utf-8-sig') as f:
    lines = f.readlines()

# Backup
with open(backup, 'w', encoding='utf-8') as f:
    f.writelines(lines)
print(f"Backup saved to {backup}")

# Fix line 891 (index 890)
print(f"\nLine 891 before: {lines[890].rstrip()}")
lines[890] = '    # Build a map of header name -> column index\n'
print(f"Line 891 after:  {lines[890].rstrip()}")

# Fix line 5295 (index 5294) - these are em-dash/en-dash characters that got corrupted
# The code is trying to replace smart quotes/dashes, let's use proper Unicode escapes
print(f"\nLine 5295 before: {repr(lines[5294].rstrip())}")
# Use Unicode escapes for em-dash and en-dash
lines[5294] = '    s = s.replace("\\u2013", "-").replace("\\u2014", "-")  # en-dash, em-dash\n'
print(f"Line 5295 after:  {repr(lines[5294].rstrip())}")

# Write back
with open(filepath, 'w', encoding='utf-8', newline='') as f:
    f.writelines(lines)

print(f"\nFile fixed and saved to {filepath}")
