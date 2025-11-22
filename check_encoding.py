#!/usr/bin/env python3
"""Check for encoding issues"""

filepath = r"c:\Users\zachs\Documents\DevProjects\PDF_Scanner\Application\eidp_term_scanner.core.py"

# Check for BOM and encoding
with open(filepath, 'rb') as f:
    first_bytes = f.read(4)
    print(f"First 4 bytes: {first_bytes.hex()}")

    # Read full file
    f.seek(0)
    data = f.read()

    # Look for UTF-8 BOM
    if data.startswith(b'\xef\xbb\xbf'):
        print("UTF-8 BOM found")

    # Look for any byte sequences that might be problematic
    # U+FFFF in UTF-8 is: EF BF BF
    if b'\xef\xbf\xbf' in data:
        pos = data.index(b'\xef\xbf\xbf')
        line_num = data[:pos].count(b'\n') + 1
        print(f"U+FFFF (UTF-8: EF BF BF) found at byte {pos}, approximately line {line_num}")

    # Check for other non-ASCII characters
    for i, byte in enumerate(data):
        if byte > 127 and byte not in [0xef, 0xbb, 0xbf]:  # Allow UTF-8 sequences
            line_num = data[:i].count(b'\n') + 1
            print(f"Non-ASCII byte 0x{byte:02x} at position {i}, line ~{line_num}")
            # Show context
            start = max(0, i-20)
            end = min(len(data), i+20)
            print(f"Context: {data[start:end]}")
            break
