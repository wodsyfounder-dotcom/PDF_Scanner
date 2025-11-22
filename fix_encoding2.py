#!/usr/bin/env python3
"""Fix all encoding issues in the core file"""
import re

filepath = r"c:\Users\zachs\Documents\DevProjects\PDF_Scanner\Application\eidp_term_scanner.core.py"
backup_path = r"c:\Users\zachs\Documents\DevProjects\PDF_Scanner\Application\eidp_term_scanner.core.py.bak"

# Read the file in binary mode
with open(filepath, 'rb') as f:
    data = f.read()

# Make a backup
with open(backup_path, 'wb') as f:
    f.write(data)
print(f"Backup created at {backup_path}")

original_size = len(data)

# Remove UTF-8 BOM if present
if data.startswith(b'\xef\xbb\xbf'):
    print("Removing UTF-8 BOM")
    data = data[3:]

# Find all sequences that look like corrupted UTF-8
# These are typically double-encoded or mangled characters
# Pattern: bytes in range 0xC0-0xFF that form invalid UTF-8 sequences

# Specific fixes based on what we found:
# Fix 1: Remove corrupted arrow-like sequences (multiple bytes starting with 0xC3)
# We'll replace any sequence of 0xC3 + other high bytes with "->"

# Look for patterns like: C3 83 C6 92 etc. (corrupted multi-byte sequences)
# Replace them with simple ASCII equivalents

# Strategy: Find any stretch of bytes >= 0x80 that don't form valid UTF-8
# and replace with a placeholder

import codecs

# Try to decode, tracking where errors occur
fixed = bytearray()
i = 0
errors_found = 0

while i < len(data):
    # Try to decode from this position
    try:
        # Attempt to decode a single character
        char_bytes = bytes([data[i]])

        # Check if this is a multi-byte sequence
        if data[i] >= 0x80:
            # Determine how many bytes this UTF-8 character should have
            if data[i] >= 0xF0:
                char_bytes = data[i:i+4]
                expected_len = 4
            elif data[i] >= 0xE0:
                char_bytes = data[i:i+3]
                expected_len = 3
            elif data[i] >= 0xC0:
                char_bytes = data[i:i+2]
                expected_len = 2
            else:
                # Invalid start byte
                raise UnicodeDecodeError('utf-8', data[i:i+1], 0, 1, 'invalid start byte')

            # Try to decode
            try:
                decoded = char_bytes.decode('utf-8')
                # Check if it's a suspicious character (likely corrupted)
                if len(decoded) == 1 and ord(decoded) > 0x7F:
                    # Check if next bytes also look corrupted
                    # If we see multiple high bytes in a row, it's likely corrupted
                    if i + expected_len < len(data) and data[i + expected_len] >= 0xC0:
                        # Skip this corrupted sequence
                        # Peek ahead to see the pattern
                        end = i
                        while end < len(data) and data[end] >= 0x80:
                            end += 1

                        # Skip the whole corrupted block
                        # Look at the context to decide what to replace with
                        before = fixed[-20:].decode('utf-8', errors='ignore')

                        if 'header name' in before or 'map of' in before:
                            fixed.extend(b'->')
                        else:
                            fixed.extend(b'-')

                        i = end
                        errors_found += 1
                        continue

                fixed.extend(char_bytes)
                i += expected_len
            except UnicodeDecodeError:
                # Skip corrupted byte(s) and insert placeholder
                print(f"Corrupted sequence at position {i}: {data[i:i+10].hex()}")
                fixed.extend(b'-')
                # Skip ahead to next ASCII or valid UTF-8 start
                i += 1
                while i < len(data) and data[i] >= 0x80 and data[i] < 0xC0:
                    i += 1
                errors_found += 1
        else:
            # Regular ASCII character
            fixed.append(data[i])
            i += 1
    except Exception as e:
        print(f"Error at position {i}: {e}")
        fixed.append(ord('-'))
        i += 1
        errors_found += 1

print(f"Fixed {errors_found} corrupted sequences")
print(f"Original size: {original_size}, New size: {len(fixed)}")

# Write the fixed version
with open(filepath, 'wb') as f:
    f.write(bytes(fixed))

print(f"File cleaned and saved to {filepath}")

# Verify it's valid UTF-8 now
try:
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()
    print("File is now valid UTF-8!")
except UnicodeDecodeError as e:
    print(f"Warning: File still has UTF-8 errors: {e}")
