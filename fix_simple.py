#!/usr/bin/env python3
"""Simple fix for encoding issues"""

filepath = r"c:\Users\zachs\Documents\DevProjects\PDF_Scanner\Application\eidp_term_scanner.core.py"

# Read file
with open(filepath, 'r', encoding='utf-8-sig') as f:  # utf-8-sig auto-removes BOM
    content = f.read()

# Count lines before
line_count = len(content.split('\n'))
print(f"File has {line_count} lines")

# Look for lines with non-ASCII Latin characters that look corrupted
import re

# Find lines with patterns that look like corrupted UTF-8
# These typically have character combinations like Ã followed by other accented characters
corrupted_pattern = re.compile(r'[ÃÂ][ÃÂƒÆ][^a-zA-Z0-9\s]{2,}')

fixes_made = []
for i, line in enumerate(content.split('\n'), 1):
    if re.search(r'[ÃÂÆ][^a-zA-Z0-9\s,.:;(){}\[\]"\']', line):
        print(f"Line {i}: Potential corruption found")
        print(f"  Content: {line[:80]}")

        # Apply common fixes
        original = line

        # Fix: Build a map of header name -> column index
        if 'Build a map of header name' in line and 'column index' in line:
            line = re.sub(r'header name[^c]+column index', 'header name -> column index', line)
            if line != original:
                fixes_made.append((i, 'Fixed arrow in comment'))
                content = content.replace(original, line)
                print(f"  Fixed: {line[:80]}")

# Write back without BOM
with open(filepath, 'w', encoding='utf-8', newline='\n') as f:
    f.write(content)

print(f"\nMade {len(fixes_made)} fixes:")
for line_num, desc in fixes_made:
    print(f"  Line {line_num}: {desc}")

print(f"\nFile saved successfully")
