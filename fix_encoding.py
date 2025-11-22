#!/usr/bin/env python3
"""Fix all encoding issues in the core file"""

filepath = r"c:\Users\zachs\Documents\DevProjects\PDF_Scanner\Application\eidp_term_scanner.core.py"
backup_path = r"c:\Users\zachs\Documents\DevProjects\PDF_Scanner\Application\eidp_term_scanner.core.py.bak"

# Read the file
with open(filepath, 'rb') as f:
    data = f.read()

# Make a backup
with open(backup_path, 'wb') as f:
    f.write(data)
print(f"Backup created at {backup_path}")

# Remove UTF-8 BOM if present
if data.startswith(b'\xef\xbb\xbf'):
    print("Removing UTF-8 BOM")
    data = data[3:]

# Now decode to text and look for corrupted sequences
text = data.decode('utf-8', errors='replace')

# Count replacement characters
replacement_count = text.count('\ufffd')
print(f"Found {replacement_count} replacement characters (\ufffd)")

# Find and show corrupted parts
lines = text.split('\n')
for i, line in enumerate(lines, 1):
    if '\ufffd' in line or any(ord(c) > 0x7F and ord(c) not in [0xA0, 0xB1] for c in line if not c.isalpha()):
        # Check for specific corrupted patterns
        if 'ÃƒÂ' in line or 'Ã¢â' in line or 'Ã‚Â' in line:
            print(f"Line {i}: {line[:100]}")

# Manually fix known corrupted strings
fixes = [
    # These are the corrupted encodings we found
    ('ÃƒÆ'Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢', '->'),  # arrow
    ('ÃƒÂ¢Ã¢â‚¬Â� Ã¢â‚¬â„¢', '->'),  # arrow variant
    ('Ãƒâ€šÃ‚Â±', '±'),  # plus-minus (if any remain)
    ('ÃƒÆ'Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â±', '±'),  # plus-minus variant
]

for old, new in fixes:
    if old in text:
        count = text.count(old)
        text = text.replace(old, new)
        print(f"Replaced {count} occurrence(s) of corrupted text with '{new}'")

# Write back
with open(filepath, 'w', encoding='utf-8', newline='\n') as f:
    f.write(text)

print(f"File cleaned and saved to {filepath}")
