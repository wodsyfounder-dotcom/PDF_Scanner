#!/usr/bin/env python3
"""Find U+FFFF character in file"""
import sys

filepath = r"c:\Users\zachs\Documents\DevProjects\PDF_Scanner\Application\eidp_term_scanner.core.py"

with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
    for line_num, line in enumerate(f, 1):
        if '\uffff' in line:
            print(f"Line {line_num}: Found U+FFFF character")
            print(f"Content: {repr(line)}")
            col = line.index('\uffff')
            print(f"Column: {col}")
        # Also check for other problematic characters
        for i, char in enumerate(line):
            if ord(char) > 0xFFFD or (0xD800 <= ord(char) <= 0xDFFF):
                print(f"Line {line_num}, Col {i}: Found problematic character U+{ord(char):04X}")
