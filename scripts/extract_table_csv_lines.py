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


def extract_text_pymupdf(pdf_path: Path, page_num: int) -> Optional[str]:
    """Extract text from a single page using PyMuPDF, preserving row structure by Y-coordinate."""
    if not _HAVE_PYMUPDF:
        return None

    import fitz  # PyMuPDF
    doc = fitz.open(str(pdf_path))
    try:
        if not (1 <= page_num <= doc.page_count):
            return ""

        page = doc.load_page(page_num - 1)  # 0-indexed

        # Get words with their positions
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


def extract_text_easyocr(pdf_path: Path, page_num: int, dpi: int = 300) -> Optional[str]:
    """Extract text from a single page using EasyOCR. Returns None if dependencies not available."""
    if not _HAVE_EASYOCR:
        return None

    try:
        import fitz
    except ImportError:
        print("[WARN] PyMuPDF needed for OCR rendering (install with: pip install pymupdf)")
        return None

    try:
        import easyocr
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
        img_np = np.array(img)

        reader = easyocr.Reader(['en'], gpu=False, verbose=False)
        results = reader.readtext(img_np)

        # Sort by Y coordinate (top to bottom), then X (left to right)
        boxes = []
        for bbox, text, conf in results:
            if conf < 0.3:  # Skip low confidence
                continue
            ys = [pt[1] for pt in bbox]
            xs = [pt[0] for pt in bbox]
            y_avg = sum(ys) / len(ys)
            x_avg = sum(xs) / len(xs)
            boxes.append((y_avg, x_avg, text))

        boxes.sort(key=lambda item: (item[0], item[1]))

        # Group by similar Y coordinates (same line)
        lines: List[List[str]] = []
        current_line: List[Tuple[float, str]] = []
        last_y = None
        y_threshold = 10  # pixels

        for y, x, text in boxes:
            if last_y is None or abs(y - last_y) <= y_threshold:
                current_line.append((x, text))
                if last_y is None:
                    last_y = y
                else:
                    last_y = (last_y + y) / 2  # Average
            else:
                # New line
                if current_line:
                    current_line.sort(key=lambda item: item[0])  # Sort by X
                    lines.append([txt for _, txt in current_line])
                current_line = [(x, text)]
                last_y = y

        if current_line:
            current_line.sort(key=lambda item: item[0])
            lines.append([txt for _, txt in current_line])

        # Join each line with spaces
        return "\n".join(" ".join(words) for words in lines)

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


def detect_tables_in_text(text: str, min_cols: int = 2, min_rows: int = 3,
                          delimiter: Optional[str] = None, debug: bool = False) -> List[Dict]:
    """
    Detect tables in text by finding consecutive lines with the same field count.

    Returns list of table dictionaries with:
    - start_line: starting line number
    - end_line: ending line number
    - num_cols: number of columns
    - rows: list of field lists
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
        print(f"\n[DEBUG] Detection parameters: min_cols={min_cols}, min_rows={min_rows}, delimiter={delimiter!r}\n")

    i = 0
    while i < len(lines):
        line = lines[i]
        fields = split_line_into_fields(line, delimiter)

        if debug and i < 20:  # Debug first 20 lines
            print(f"[DEBUG] Line {i}: {len(fields)} fields -> {fields}")

        if not fields or len(fields) < min_cols:
            i += 1
            continue

        # Found a potential table start
        num_cols = len(fields)
        table_rows: List[List[str]] = [fields]
        start_line = i

        if debug:
            print(f"[DEBUG] Potential table start at line {i} with {num_cols} columns")

        # Look ahead for consecutive lines with same field count
        j = i + 1
        while j < len(lines):
            next_fields = split_line_into_fields(lines[j], delimiter)
            if len(next_fields) == num_cols:
                table_rows.append(next_fields)
                j += 1
            else:
                break

        # Check if we have enough rows to consider this a table
        if len(table_rows) >= min_rows:
            if debug:
                print(f"[DEBUG] ✓ Table confirmed: lines {start_line}-{j-1}, {num_cols} cols, {len(table_rows)} rows")
            tables.append({
                'start_line': start_line,
                'end_line': j - 1,
                'num_cols': num_cols,
                'rows': table_rows
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
                              delimiter: Optional[str] = None, debug: bool = False) -> Dict[int, List[Dict]]:
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

    for page_num in pages:
        print(f"[INFO] Processing page {page_num}...")

        if actual_use_ocr:
            text = extract_text_easyocr(pdf_path, page_num, dpi)
        else:
            text = extract_text_pymupdf(pdf_path, page_num)

        if text is None:
            print(f"[ERROR] Failed to extract text from page {page_num}")
            results[page_num] = []
            continue

        if not text.strip():
            print(f"[WARN] No text found on page {page_num}")
            results[page_num] = []
            continue

        tables = detect_tables_in_text(text, min_cols, min_rows, delimiter, debug)
        results[page_num] = tables

        if tables:
            print(f"[INFO] Found {len(tables)} table(s) on page {page_num}")
            for idx, table in enumerate(tables, 1):
                print(f"  Table {idx}: {table['num_cols']} columns, {len(table['rows'])} rows")
        else:
            print(f"[INFO] No tables detected on page {page_num}")

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

            # Try to detect header row (first row if it's different from others)
            if len(rows) > 1:
                # Use first row as header
                headers = rows[0]
                data_rows = rows[1:]

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

                df = pd.DataFrame(data_rows, columns=unique_headers)
            else:
                # Single row, use generic column names
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
    parser.add_argument('--delimiter', default=None, help='Field delimiter (default: auto-detect)')
    parser.add_argument('--out', default='', help='Output Excel path (default: Data Packages/<pdf-stem>_table.xlsx)')
    parser.add_argument('--debug', action='store_true', help='Print debug information (extracted text and field parsing)')

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
        root = P(__file__).resolve().parents[1]
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
        debug=args.debug
    )

    # Write to Excel
    write_tables_to_excel(page_tables, output_path)

    print(f"[DONE] Output saved to: {output_path}")


if __name__ == '__main__':
    main()
