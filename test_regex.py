import re

_AERO_UNITS = (
    "%|ppm|ppb|ms|s|sec|kg|g|mg|ug|lb|lbm|lbf|lbs|"
    "N|kN|mN|Ns|bar|mbar|Pa|kPa|MPa|psi|psia|psig|"
    "mm|cm|m|in|ft|K|degC|degF|C|F"
)
NUMBER_REGEX = re.compile(
    rf"""
    (?<![A-Za-z0-9_.-])           # left boundary
    [-+]?                         # optional sign
    (?:\d{{1,3}}(?:,\d{{3}})+|\d+)    # integer with thousands or plain digits
    (?:\.\d+)?                    # optional decimal part
    (?:\s?(?:{_AERO_UNITS}))?      # optional aerospace units
    (?![A-Za-z0-9_.-])            # right boundary
    """,
    re.VERBOSE,
)

def _fix_ocr_in_numbers(text: str) -> str:
    """Fix common OCR errors in numeric strings."""
    if not text:
        return text
    # Fix O → 0 when adjacent to digits or decimal points
    text = re.sub(r'(\d)O(?=\d|[a-z]|\s|$)', r'\g<1>0', text)  # digitO
    text = re.sub(r'(^|\s|\.)O(?=\d)', r'\g<1>0', text)  # Odigit or .O
    text = re.sub(r'(\d|\.)O(?=\D|$)', r'\g<1>0', text)  # digit.O or O at end
    text = re.sub(r'\.O(?=[a-z])', '.0', text)  # .Opsig → .0psig
    return text

print("Testing OCR-corrupted values:")
test_values = ['151.Opsig', '12OO.Opsig', '124O.7psig', '151.0psig', '1200.0psig']
for val in test_values:
    fixed = _fix_ocr_in_numbers(val)
    match = NUMBER_REGEX.search(fixed)
    result = match.group(0) if match else 'NO MATCH'
    print(f"{val:20} -> fixed: {fixed:20} -> {result}")
