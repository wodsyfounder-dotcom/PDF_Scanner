#!/usr/bin/env python3
"""Find all problematic characters"""

filepath = r"c:\Users\zachs\Documents\DevProjects\PDF_Scanner\Application\eidp_term_scanner.core.py"

with open(filepath, 'rb') as f:
    data = f.read()

# Skip the UTF-8 BOM if present
start = 3 if data.startswith(b'\xef\xbb\xbf') else 0

# Look for multi-byte UTF-8 sequences that might be problematic
i = start
line_num = 1
issues = []

while i < len(data):
    byte = data[i]

    # Track line numbers
    if byte == ord('\n'):
        line_num += 1
        i += 1
        continue

    # Check for multi-byte UTF-8 sequences (bytes starting with 0xC0-0xF7)
    if byte >= 0x80:
        # Get context
        line_start = data.rfind(b'\n', 0, i) + 1
        line_end = data.find(b'\n', i)
        if line_end == -1:
            line_end = len(data)
        line_content = data[line_start:line_end]

        try:
            # Try to decode as UTF-8
            line_content.decode('utf-8')
        except UnicodeDecodeError as e:
            issues.append((line_num, i, byte, line_content))
            print(f"Line {line_num}: Decode error at byte position {i}")
            print(f"  Byte: 0x{byte:02x}")
            print(f"  Context: {line_content[:100]}")
            print()

    i += 1

# Also check for U+FFFF specifically
if b'\xef\xbf\xbf' in data:
    pos = data.index(b'\xef\xbf\xbf')
    line_num = data[:pos].count(b'\n') + 1
    print(f"U+FFFF found at line {line_num}, byte position {pos}")

print(f"\nTotal issues found: {len(issues)}")
