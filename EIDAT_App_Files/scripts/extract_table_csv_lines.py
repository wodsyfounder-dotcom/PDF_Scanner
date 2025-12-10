#!/usr/bin/env python3
"""
Extract tables from PDF pages using line-by-line CSV field detection.

Algorithm:
- Extracts text from specified pages (PyMuPDF or EasyOCR fallback)
- Reads each line and attempts to parse it as CSV-like data (detecting field count)
- When 3+ consecutive lines have the same number of fields (X), identifies it as a table with X columns
- Extracts all rows matching that pattern as a table
- Outputs to a single Excel file in Data Packages folder

Usage:
  python scripts/extract_table_csv_lines.py --pdf path/to/file.pdf --pages "1,3-5"
  python scripts/extract_table_csv_lines.py --pdf path/to/file.pdf --pages "1,3-5" --ocr
  python scripts/extract_table_csv_lines.py --pdf path/to/file.pdf --pages "1" --min-cols 3 --min-rows 3

Options:
  --pdf PATH          Path to PDF file (required)
  --pages RANGE       Page range (e.g., "1,3-5") - 1-indexed (required)
  --ocr               Force OCR mode (EasyOCR) instead of PyMuPDF text extraction
  --dpi INT           DPI for OCR rendering (default: 300)
  --min-cols INT      Minimum number of columns to consider a table (default: 2)
  --min-rows INT      Minimum consecutive rows with same field count (default: 3)
  --out PATH          Output Excel path (default: Data Packages/<pdf-stem>_table.xlsx)
  --delimiter STR     Field delimiter (default: auto-detect from whitespace/tabs)
  --no-auto-rotate    Disable automatic rotation retry for landscape pages
"""

from __future__ import annotations
import argparse
import re
import sys
from pathlib import Path
from typing import List, Optional, Tuple, Dict
import csv as csv_module

# Check available PDF processing libraries
_HAVE_PYMUPDF = False
_HAVE_EASYOCR = False
_EASYOCR_READER = None

try:
    import fitz  # PyMuPDF
    _HAVE_PYMUPDF = True
except ImportError:
    pass

try:
    import easyocr
    _HAVE_EASYOCR = True
except ImportError:
    pass


def _get_easyocr_reader(langs: Optional[List[str]] = None):
    """Lazily instantiate a shared EasyOCR reader to avoid repeated heavy init."""
    global _EASYOCR_READER
    if not _HAVE_EASYOCR:
        return None
    if _EASYOCR_READER is None:
        langs = langs or ['en']
        _EASYOCR_READER = easyocr.Reader(langs, gpu=False, verbose=False)
    return _EASYOCR_READER


def parse_page_ranges(s: str) -> List[int]:
    """Parse page range string like '1,3-5,7' into list of page numbers."""
    if not s:
        return []
    pages: set[int] = set()
    for part in re.split(r"[;,\s]+", s.strip()):
        if not part:
            continue
        if "-" in part:
            a, b = part.split('-', 1)
            try:
                ai = int(re.sub(r"\D", "", a))
                bi = int(re.sub(r"\D", "", b))
                if ai > bi:
                    ai, bi = bi, ai
                for p in range(ai, bi + 1):
                    pages.add(p)
            except Exception:
                continue
        else:
            try:
                pages.add(int(re.sub(r"\D", "", part)))
            except Exception:
                continue
    return sorted(pages)


def extract_text_pymupdf(pdf_path: Path, page_num: int, rotation: int = 0) -> Optional[str]:
    """Extract text from a single page using PyMuPDF, optionally rotating to compensate for landscape scans."""
    if not _HAVE_PYMUPDF:
        return None

    import fitz  # PyMuPDF
    doc = fitz.open(str(pdf_path))
    try:
        if not (1 <= page_num <= doc.page_count):
            return ""

        page = doc.load_page(page_num - 1)  # 0-indexed

        # Apply rotation via text page matrix if requested
        if rotation:
            matrix = fitz.Matrix(1, 1).preRotate(rotation)
            text_page = page.get_textpage(matrix=matrix)
            words = text_page.extractWORDS()
        else:
            text_page = None
            words = page.get_text("words")  # Returns list of (x0, y0, x1, y1, "word", block_no, line_no, word_no)

        if not words:
            return ""

        # Group words by their Y coordinate (same row)
        # Use a tolerance to handle slight vertical misalignment
        rows_dict = {}
        y_tolerance = 3  # pixels

        for word_info in words:
            x0, y0, x1, y1, text = word_info[0], word_info[1], word_info[2], word_info[3], word_info[4]
            text = text.strip()
            if not text:
                continue

            # Find if this Y coordinate matches an existing row
            y_mid = (y0 + y1) / 2
            matched_y = None

            for existing_y in rows_dict.keys():
                if abs(y_mid - existing_y) <= y_tolerance:
                    matched_y = existing_y
                    break

            if matched_y is None:
                matched_y = y_mid
                rows_dict[matched_y] = []

            # Store word with its X position AND bounding box for accurate spacing
            rows_dict[matched_y].append((x0, x1, text))

        # Sort rows by Y coordinate (top to bottom)
        sorted_rows = sorted(rows_dict.items())

        # Build text with one row per line, preserving column spacing
        lines = []
        for y, words_in_row in sorted_rows:
            # Sort words in row by X coordinate (left to right)
            words_in_row.sort(key=lambda item: item[0])

            # Reconstruct line with spacing based on X-position gaps
            if not words_in_row:
                continue

            line_parts = []
            prev_x_end = None

            for x_start, x_end, word in words_in_row:
                if prev_x_end is not None:
                    # Calculate gap between previous word and current word
                    gap = x_start - prev_x_end
                    # If gap is large, insert multiple spaces to indicate column boundary
                    # Use smaller threshold since tables often have modest column spacing
                    if gap > 15:
                        line_parts.append("  ")  # Double space for column separator
                    elif gap > 3:
                        line_parts.append(" ")
                    # If gap is very small, words might be touching (like hyphenated)

                line_parts.append(word)
                prev_x_end = x_end  # Use actual bounding box end

            line_text = " ".join(line_parts)
            lines.append(line_text)

        return "\n".join(lines)

    finally:
        try:
            doc.close()
        except Exception:
            pass


def extract_text_easyocr(pdf_path: Path, page_num: int, dpi: int = 300, rotation: int = 0) -> Optional[str]:
    """Extract text from a single page using EasyOCR. Supports rotation to better handle landscape layouts."""
    if not _HAVE_EASYOCR:
        return None

    try:
        import fitz
    except ImportError:
        print("[WARN] PyMuPDF needed for OCR rendering (install with: pip install pymupdf)")
        return None

    try:
        from PIL import Image
        import numpy as np
    except ImportError:
        return None

    doc = fitz.open(str(pdf_path))
    try:
        if not (1 <= page_num <= doc.page_count):
            return ""

        page = doc.load_page(page_num - 1)
        pix = page.get_pixmap(dpi=max(200, min(800, dpi)))
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        if rotation:
            img = img.rotate(rotation, expand=True)
        img_np = np.array(img)

        reader = _get_easyocr_reader(['en'])
        if reader is None:
            return None
        results = reader.readtext(img_np)

        # Sort by Y coordinate (top to bottom), then X (left to right)
        boxes = []
        for bbox, text, conf in results:
            if conf < 0.3:  # Skip low confidence
                continue
            text = text.strip()
            if not text:
                continue
            ys = [pt[1] for pt in bbox]
            xs = [pt[0] for pt in bbox]
            y0, y1 = min(ys), max(ys)
            x0, x1 = min(xs), max(xs)
            y_mid = (y0 + y1) / 2.0
            boxes.append((y_mid, x0, x1, text))

        boxes.sort(key=lambda item: (item[0], item[1]))

        # Group by similar Y coordinates (same line)
        lines: List[List[Tuple[float, float, str]]] = []
        current_line: List[Tuple[float, float, str]] = []
        last_y = None
        y_threshold = 10  # pixels

        for y_mid, x0, x1, text in boxes:
            if last_y is None or abs(y_mid - last_y) <= y_threshold:
                current_line.append((x0, x1, text))
                if last_y is None:
                    last_y = y_mid
                else:
                    last_y = (last_y + y_mid) / 2  # Average
            else:
                # New line
                if current_line:
                    current_line.sort(key=lambda item: item[0])  # Sort by X
                    lines.append(current_line)
                current_line = [(x0, x1, text)]
                last_y = y_mid

        if current_line:
            current_line.sort(key=lambda item: item[0])
            lines.append(current_line)

        rendered_lines = []
        for line_boxes in lines:
            parts: List[str] = []
            prev_end = None
            for x0, x1, word in line_boxes:
                if prev_end is None:
                    parts.append(word)
                else:
                    gap = x0 - prev_end
                    if gap > 80:
                        spacer = "    "
                    elif gap > 35:
                        spacer = "  "
                    elif gap > 12:
                        spacer = " "
                    else:
                        spacer = " "
                    parts.append(spacer + word)
                prev_end = x1
            rendered_lines.append("".join(parts))

        return "\n".join(rendered_lines)

    finally:
        try:
            doc.close()
        except Exception:
            pass


def split_line_into_fields(line: str, delimiter: Optional[str] = None) -> List[str]:
    """
    Split a line into fields.

    If delimiter is provided, use it.
    Otherwise, try to auto-detect: split by multiple spaces/tabs.
    """
    line = line.strip()
    if not line:
        return []

    if delimiter:
        return [f.strip() for f in line.split(delimiter) if f.strip()]

    # Auto-detect: split by 2+ spaces or tabs
    # This handles common table formatting where columns are separated by multiple spaces
    fields = re.split(r'\s{2,}|\t+', line)
    fields = [f.strip() for f in fields if f.strip()]

    return fields


# ============================================================================
# Type Detection and Smart Column Alignment
# ============================================================================

def detect_value_type(value: str) -> str:
    """
    Detect the type of a value: 'number', 'date', or 'text'.
    """
    value = value.strip()
    if not value:
        return 'empty'

    # Check for number (including scientific notation, negative, decimals)
    # Allow numbers with units like "500" or "42.5"
    number_pattern = r'^[-+]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?(?:[eE][+-]?\d+)?(?:\s*[a-zA-Z%]+)?$'
    if re.match(number_pattern, value):
        return 'number'

    # Check for date patterns (MM/DD/YYYY, YYYY-MM-DD, DD-MM-YYYY, etc.)
    date_patterns = [
        r'^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}$',  # MM/DD/YYYY or DD-MM-YYYY
        r'^\d{4}[/-]\d{1,2}[/-]\d{1,2}$',     # YYYY-MM-DD
        r'^\d{1,2}\s+[A-Za-z]{3,9}\s+\d{2,4}$',  # 15 January 2024
    ]
    for pattern in date_patterns:
        if re.match(pattern, value):
            return 'date'

    return 'text'


def profile_column_types(rows: List[List[str]], skip_first_n: int = 1) -> List[Dict]:
    """
    Analyze column types by examining values in each column.

    Returns list of dicts with type frequencies for each column.
    skip_first_n: number of header rows to skip (default 1)
    """
    if not rows or len(rows) <= skip_first_n:
        return []

    # Determine max column count
    max_cols = max(len(row) for row in rows)

    # Initialize profiles
    profiles = []
    for col_idx in range(max_cols):
        profiles.append({
            'number': 0,
            'date': 0,
            'text': 0,
            'empty': 0,
            'values': []
        })

    # Analyze data rows (skip headers)
    for row in rows[skip_first_n:]:
        for col_idx, value in enumerate(row):
            if col_idx >= max_cols:
                break

            value_type = detect_value_type(value)
            profiles[col_idx][value_type] += 1
            profiles[col_idx]['values'].append(value)

    # Determine dominant type for each column
    for profile in profiles:
        total = sum([profile['number'], profile['date'], profile['text']])
        if total > 0:
            if profile['number'] / total > 0.5:
                profile['dominant_type'] = 'number'
            elif profile['date'] / total > 0.3:
                profile['dominant_type'] = 'date'
            else:
                profile['dominant_type'] = 'text'
        else:
            profile['dominant_type'] = 'text'

    return profiles


def calculate_type_match_score(value: str, expected_type: str, col_values: List[str]) -> float:
    """
    Calculate how well a value matches the expected column type.

    Returns score between 0.0 (no match) and 1.0 (perfect match).
    """
    if not value.strip():
        return 0.5  # Empty can fit anywhere

    actual_type = detect_value_type(value)

    # Type match
    if actual_type == expected_type:
        score = 1.0
    elif actual_type == 'empty':
        score = 0.5
    else:
        score = 0.0

    # Length similarity bonus (for text columns)
    if expected_type == 'text' and col_values:
        non_empty_values = [v for v in col_values if v.strip()]
        if non_empty_values:
            avg_len = sum(len(v) for v in non_empty_values) / len(non_empty_values)
            len_ratio = min(len(value), avg_len) / max(len(value), avg_len, 1)
            score = (score + len_ratio) / 2

    return score


def align_row_to_columns(fields: List[str], expected_cols: int,
                        column_profiles: Optional[List[Dict]] = None,
                        match_threshold: float = 0.5,
                        debug: bool = False) -> Optional[List[str]]:
    """
    Try to align a row with mismatched field count to expected columns.

    Uses type matching to snap values to correct columns.
    Returns aligned row or None if alignment fails.
    """
    if len(fields) == expected_cols:
        return fields

    if debug:
        print(f"[DEBUG] Aligning {len(fields)} fields to {expected_cols} columns")

    # If no profiles, can't do smart alignment
    if not column_profiles or len(column_profiles) != expected_cols:
        # Simple strategy: pad with empty or truncate
        if len(fields) < expected_cols:
            return fields + [''] * (expected_cols - len(fields))
        else:
            return fields[:expected_cols]

    # Smart alignment using type matching
    aligned = [''] * expected_cols
    used_fields = [False] * len(fields)

    # Try to match each field to best column
    for field_idx, field in enumerate(fields):
        if used_fields[field_idx]:
            continue

        best_col = None
        best_score = match_threshold

        for col_idx in range(expected_cols):
            if aligned[col_idx]:  # Column already filled
                continue

            profile = column_profiles[col_idx]
            score = calculate_type_match_score(
                field,
                profile['dominant_type'],
                profile['values']
            )

            if debug:
                print(f"[DEBUG]   Field '{field}' -> Col {col_idx} ({profile['dominant_type']}): score={score:.2f}")

            if score > best_score:
                best_score = score
                best_col = col_idx

        if best_col is not None:
            aligned[best_col] = field
            used_fields[field_idx] = True
            if debug:
                print(f"[DEBUG]   ✓ Matched '{field}' to column {best_col} (score={best_score:.2f})")

    # Check if alignment is successful (at least some fields matched)
    if any(aligned):
        return aligned
    else:
        return None


def is_likely_header(fields: List[str], next_fields: Optional[List[str]] = None) -> bool:
    """
    Determine if a row is likely a header row.

    Headers typically:
    - Are mostly text
    - Have different types than data rows below
    - Have descriptive words (min, max, value, name, etc.)
    """
    if not fields:
        return False

    # Check if mostly text
    text_count = sum(1 for f in fields if detect_value_type(f) == 'text')
    if text_count / len(fields) < 0.6:  # Less than 60% text
        return False

    # Check for common header keywords
    header_keywords = ['min', 'max', 'value', 'name', 'units', 'measurement',
                      'type', 'id', 'resistance', 'test', 'date', 'time']
    has_keyword = any(
        any(keyword in f.lower() for keyword in header_keywords)
        for f in fields
    )

    # If next row exists and has different types, likely header
    if next_fields and len(next_fields) == len(fields):
        type_matches = sum(
            1 for i in range(len(fields))
            if detect_value_type(fields[i]) == detect_value_type(next_fields[i])
        )
        if type_matches / len(fields) < 0.3:  # Less than 30% type match
            return True

    return has_keyword


def detect_tables_in_text(text: str, min_cols: int = 2, min_rows: int = 3,
                          delimiter: Optional[str] = None, debug: bool = False,
                          fixed_cols: Optional[int] = None, match_threshold: float = 0.5) -> List[Dict]:
    """
    Detect tables in text with smart column alignment and type-based row matching.

    Parameters:
    - fixed_cols: If set, enforce this exact column count and use smart alignment
    - match_threshold: Minimum score (0-1) for type-based column matching

    Returns list of table dictionaries with:
    - start_line: starting line number
    - end_line: ending line number
    - num_cols: number of columns
    - rows: list of field lists
    - headers: detected header rows
    """
    lines = text.split('\n')
    tables: List[Dict] = []

    if debug:
        print(f"\n[DEBUG] Extracted text ({len(lines)} lines):")
        print("=" * 80)
        for idx, line in enumerate(lines[:50], 1):  # Show first 50 lines
            print(f"{idx:3d}: {line}")
        if len(lines) > 50:
            print(f"... ({len(lines) - 50} more lines)")
        print("=" * 80)
        mode = f"fixed_cols={fixed_cols}" if fixed_cols else f"min_cols={min_cols}"
        print(f"\n[DEBUG] Detection parameters: {mode}, min_rows={min_rows}, match_threshold={match_threshold}\n")

    i = 0
    while i < len(lines):
        line = lines[i]
        fields = split_line_into_fields(line, delimiter)

        if debug and i < 30:  # Debug first 30 lines
            print(f"[DEBUG] Line {i}: {len(fields)} fields -> {fields}")

        # Skip empty or too-short lines
        if not fields or (not fixed_cols and len(fields) < min_cols):
            i += 1
            continue

        # Determine expected column count
        if fixed_cols:
            expected_cols = fixed_cols
            # Check if this line could start a table with fixed columns
            if len(fields) not in [expected_cols, expected_cols - 1, expected_cols + 1]:
                # Too different from expected, skip
                i += 1
                continue
        else:
            expected_cols = len(fields)

        # Check if this might be a header
        next_fields = split_line_into_fields(lines[i + 1], delimiter) if i + 1 < len(lines) else None
        is_header = is_likely_header(fields, next_fields)

        if debug:
            header_str = " [HEADER]" if is_header else ""
            print(f"[DEBUG] Potential table start at line {i} with {len(fields)} fields (expected {expected_cols}){header_str}")

        # Start building table
        table_rows: List[List[str]] = []
        headers: List[str] = []
        start_line = i

        # Add first row (might be header or data)
        if len(fields) == expected_cols:
            if is_header:
                headers = fields
            else:
                table_rows.append(fields)
        else:
            # Try to align
            aligned = align_row_to_columns(fields, expected_cols, None, match_threshold, debug)
            if aligned:
                if is_header:
                    headers = aligned
                else:
                    table_rows.append(aligned)
            else:
                i += 1
                continue

        j = i + 1

        # Look ahead and try to add more rows
        # After collecting a few rows, build column profiles for smart matching
        while j < len(lines):
            line_fields = split_line_into_fields(lines[j], delimiter)

            if not line_fields:  # Empty line might signal table end
                if debug:
                    print(f"[DEBUG]   Line {j}: empty, checking if table should end...")
                # Allow one empty line, but if next line also empty, break
                if j + 1 < len(lines) and not split_line_into_fields(lines[j + 1], delimiter):
                    break
                j += 1
                continue

            # Build column profiles if we have enough data rows
            column_profiles = None
            if len(table_rows) >= 3:
                column_profiles = profile_column_types([headers] + table_rows if headers else table_rows)

            # Try to add this row
            if len(line_fields) == expected_cols:
                # Check if it's a repeated header (page break continuation)
                if headers and line_fields == headers:
                    if debug:
                        print(f"[DEBUG]   Line {j}: repeated header, continuing table")
                    j += 1
                    continue

                table_rows.append(line_fields)
                j += 1

            elif fixed_cols:
                # Try smart alignment in fixed column mode
                aligned = align_row_to_columns(line_fields, expected_cols, column_profiles, match_threshold, debug)

                if aligned:
                    table_rows.append(aligned)
                    j += 1
                else:
                    # Check if this is a table-ending indicator
                    # All text or single value = likely not part of table
                    all_text = all(detect_value_type(f) == 'text' for f in line_fields)
                    single_value = len(line_fields) == 1

                    if all_text or single_value:
                        if debug:
                            print(f"[DEBUG]   Line {j}: ending table (all text or single value)")
                        break
                    else:
                        # Skip this malformed row but continue looking
                        if debug:
                            print(f"[DEBUG]   Line {j}: skipping malformed row")
                        j += 1
            else:
                # In auto-detect mode, different column count breaks table
                if debug:
                    print(f"[DEBUG]   Line {j}: different column count ({len(line_fields)} vs {expected_cols}), ending table")
                break

        # Check if we have enough rows to consider this a table
        if len(table_rows) >= min_rows:
            if debug:
                print(f"[DEBUG] ✓ Table confirmed: lines {start_line}-{j-1}, {expected_cols} cols, {len(table_rows)} data rows")
            tables.append({
                'start_line': start_line,
                'end_line': j - 1,
                'num_cols': expected_cols,
                'rows': table_rows,
                'headers': headers
            })
            i = j  # Skip past this table
        else:
            if debug:
                print(f"[DEBUG] ✗ Not enough rows: only {len(table_rows)} (need {min_rows})")
            i += 1

    return tables


def extract_tables_from_pages(pdf_path: Path, pages: List[int],
                              use_ocr: bool = False, dpi: int = 300,
                              min_cols: int = 2, min_rows: int = 3,
                              delimiter: Optional[str] = None, debug: bool = False,
                              fixed_cols: Optional[int] = None,
                              match_threshold: float = 0.5,
                              auto_rotate: bool = True) -> Dict[int, List[Dict]]:
    """
    Extract tables from specified pages.

    Returns dict mapping page_num -> list of table dicts.
    """
    results: Dict[int, List[Dict]] = {}

    # Auto-fallback to OCR if PyMuPDF not available
    actual_use_ocr = use_ocr
    if not use_ocr and not _HAVE_PYMUPDF:
        if _HAVE_EASYOCR:
            print("[INFO] PyMuPDF not available, automatically using OCR mode (EasyOCR)")
            actual_use_ocr = True
        else:
            print("[ERROR] Neither PyMuPDF nor EasyOCR is available!")
            print("  Install PyMuPDF: pip install pymupdf")
            print("  Or install EasyOCR: pip install easyocr")
            sys.exit(1)

    rotation_candidates = [0]
    if auto_rotate:
        rotation_candidates = [0, 90, 270]

    for page_num in pages:
        print(f"[INFO] Processing page {page_num}...")

        page_tables: List[Dict] = []
        used_rotation: Optional[int] = None
        extracted_any_text = False

        for rotation in rotation_candidates:
            if actual_use_ocr:
                text = extract_text_easyocr(pdf_path, page_num, dpi, rotation=rotation)
            else:
                text = extract_text_pymupdf(pdf_path, page_num, rotation=rotation)

            if text is None:
                continue
            if not text.strip():
                continue
            extracted_any_text = True

            tables = detect_tables_in_text(text, min_cols, min_rows, delimiter, debug,
                                           fixed_cols, match_threshold)
            if tables:
                page_tables = tables
                used_rotation = rotation
                break

        results[page_num] = page_tables

        if used_rotation and used_rotation != 0:
            print(f"[INFO] Page {page_num}: detected tables after rotating {used_rotation}°")

        if page_tables:
            print(f"[INFO] Found {len(page_tables)} table(s) on page {page_num}")
            for idx, table in enumerate(page_tables, 1):
                print(f"  Table {idx}: {table['num_cols']} columns, {len(table['rows'])} rows")
        else:
            if auto_rotate:
                print(f"[INFO] No tables detected on page {page_num} (after trying rotations {rotation_candidates})")
            else:
                print(f"[INFO] No tables detected on page {page_num}")
            if not extracted_any_text:
                mode_label = "OCR" if actual_use_ocr else "PyMuPDF"
                print(f"[WARN] Unable to extract readable text from page {page_num} using {mode_label}.")

    return results


def write_tables_to_excel(page_tables: Dict[int, List[Dict]], output_path: Path) -> None:
    """
    Write extracted tables to Excel.

    Each table gets its own sheet named: Page{page}_Table{idx}
    If only one table across all pages, use simpler naming.
    """
    try:
        import pandas as pd
    except ImportError as e:
        raise RuntimeError("pandas is required to write Excel files. Install with: pip install pandas openpyxl") from e

    # Choose Excel engine
    try:
        import xlsxwriter
        engine = 'xlsxwriter'
    except ImportError:
        try:
            import openpyxl
            engine = 'openpyxl'
        except ImportError as e:
            raise RuntimeError("Either xlsxwriter or openpyxl is required. Install with: pip install xlsxwriter") from e

    output_path.parent.mkdir(parents=True, exist_ok=True)

    # Collect all tables
    all_tables: List[Tuple[int, int, Dict]] = []
    for page_num in sorted(page_tables.keys()):
        for table_idx, table in enumerate(page_tables[page_num], 1):
            all_tables.append((page_num, table_idx, table))

    if not all_tables:
        print("[WARN] No tables to write")
        # Create empty file
        with pd.ExcelWriter(output_path, engine=engine) as writer:
            df = pd.DataFrame([["No tables detected"]], columns=["Result"])
            df.to_excel(writer, sheet_name="Summary", index=False)
        return

    with pd.ExcelWriter(output_path, engine=engine) as writer:
        for page_num, table_idx, table in all_tables:
            # Determine sheet name
            if len(all_tables) == 1:
                sheet_name = "Table"
            else:
                sheet_name = f"Page{page_num}_T{table_idx}"

            # Limit sheet name to 31 characters (Excel limit)
            sheet_name = sheet_name[:31]

            # Create DataFrame
            rows = table['rows']
            headers = table.get('headers', [])

            # Use detected headers or generate column names
            if headers:
                # Ensure headers are unique
                seen = {}
                unique_headers = []
                for h in headers:
                    if h in seen:
                        seen[h] += 1
                        unique_headers.append(f"{h}_{seen[h]}")
                    else:
                        seen[h] = 0
                        unique_headers.append(h)

                df = pd.DataFrame(rows, columns=unique_headers)
            else:
                # No headers detected, use generic column names
                df = pd.DataFrame(rows, columns=[f"Col_{i+1}" for i in range(table['num_cols'])])

            df.to_excel(writer, sheet_name=sheet_name, index=False)

            # Auto-adjust column widths if using xlsxwriter
            if engine == 'xlsxwriter':
                try:
                    worksheet = writer.sheets[sheet_name]
                    for i, col in enumerate(df.columns):
                        max_len = max(
                            df[col].astype(str).map(len).max() if not df.empty else 0,
                            len(str(col))
                        )
                        worksheet.set_column(i, i, min(60, max(10, max_len + 2)))
                except Exception:
                    pass

    print(f"[DONE] Wrote tables to {output_path}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extract tables from PDF pages using line-by-line CSV field detection",
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument('--pdf', required=True, help='Path to PDF file')
    parser.add_argument('--pages', required=True, help='Page range (e.g., "1,3-5") - 1-indexed')
    parser.add_argument('--ocr', action='store_true', help='Force OCR mode (EasyOCR)')
    parser.add_argument('--dpi', type=int, default=300, help='DPI for OCR (default: 300)')
    parser.add_argument('--min-cols', type=int, default=2, help='Minimum columns for table (default: 2)')
    parser.add_argument('--min-rows', type=int, default=3, help='Minimum consecutive rows (default: 3)')
    parser.add_argument('--num-cols', type=int, default=None, help='Fixed column count mode - enforce exact number of columns')
    parser.add_argument('--match-threshold', type=float, default=0.5, help='Type matching threshold 0-1 for column alignment (default: 0.5)')
    parser.add_argument('--delimiter', default=None, help='Field delimiter (default: auto-detect)')
    parser.add_argument('--out', default='', help='Output Excel path (default: Data Packages/<pdf-stem>_table.xlsx)')
    parser.add_argument('--debug', action='store_true', help='Print debug information (extracted text and field parsing)')
    parser.add_argument('--no-auto-rotate', action='store_true', help='Disable automatic landscape rotation retries')

    args = parser.parse_args()

    # Validate PDF
    pdf_path = Path(args.pdf)
    if not pdf_path.exists():
        print(f"[ERROR] PDF not found: {pdf_path}", file=sys.stderr)
        sys.exit(1)

    # Parse pages
    pages = parse_page_ranges(args.pages)
    if not pages:
        print(f"[ERROR] No valid pages specified: {args.pages}", file=sys.stderr)
        sys.exit(1)

    print(f"[INFO] Extracting tables from pages: {pages}")

    # Determine output path
    if args.out:
        output_path = Path(args.out)
    else:
        # Default to Data Packages folder
        from pathlib import Path as P
        root = P(__file__).resolve().parents[2]
        data_packages = root / "Data Packages"
        output_path = data_packages / f"{pdf_path.stem}_table.xlsx"

    # Extract tables
    page_tables = extract_tables_from_pages(
        pdf_path,
        pages,
        use_ocr=args.ocr,
        dpi=args.dpi,
        min_cols=args.min_cols,
        min_rows=args.min_rows,
        delimiter=args.delimiter,
        debug=args.debug,
        fixed_cols=args.num_cols,
        match_threshold=args.match_threshold,
        auto_rotate=(not args.no_auto_rotate)
    )

    # Write to Excel
    write_tables_to_excel(page_tables, output_path)

    print(f"[DONE] Output saved to: {output_path}")


if __name__ == '__main__':
    main()
