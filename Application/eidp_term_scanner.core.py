
#!/usr/bin/env python3
# Application-consolidated build
"""
EIDP Term Scanner (Matrix + Metadata + Auto-Move)
-------------------------------------------------
- Scans PDFs in a given folder for configured "terms" within specified page ranges.
- Extracts the closest numeric value near each term occurrence.
- Serial Number (SN) is inferred from each PDF's filename; results are arranged
  as a wide matrix: rows=terms, columns=serial numbers (one per EIDP).
- Produces an Excel workbook with:
    * "results"  : Term, Pages, and one column per SN with the matched number
    * "metadata" : detailed per-term/per-file records, for auditing/debugging
- Moves scanned PDFs from the imports folder into "Scanned Docs" to avoid reprocessing.
- Falls back to CSVs if Excel writer dependencies are not available.
"""

import argparse
import csv
import json
import os
import re
import shutil
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

# ---------- Optional Excel dependencies with graceful fallback ----------
_HAVE_PANDAS = False
_HAVE_OPENPYXL_OR_XLSXWRITER = False
try:
    import pandas as pd  # Used to write Excel if available
    _HAVE_PANDAS = True
    # Check for any Excel writer engine (xlsxwriter or openpyxl)
    try:
        import xlsxwriter  # noqa: F401
        _HAVE_OPENPYXL_OR_XLSXWRITER = True
    except Exception:
        try:
            import openpyxl  # noqa: F401
            _HAVE_OPENPYXL_OR_XLSXWRITER = True
        except Exception:
            pass
except Exception:
    # If pandas isn't available, we'll write CSVs instead of XLSX
    pass

# ---------- Robust PDF text extraction / OCR imports (all optional) ----------
_HAVE_PYMUPDF = False
_HAVE_PDFMINER = False
_HAVE_PYPDF = False
_HAVE_TESSERACT = False
_HAVE_PDF2IMAGE = False

try:
    import fitz  # PyMuPDF: fast, high-fidelity text extraction
    _HAVE_PYMUPDF = True
except Exception:
    pass

try:
    from pdfminer.high_level import extract_text as pdfminer_extract_text  # pdfminer.six
    _HAVE_PDFMINER = True
except Exception:
    pass

try:
    try:
        from pypdf import PdfReader as _PdfReader  # modern fork of PyPDF2
    except Exception:
        from PyPDF2 import PdfReader as _PdfReader  # legacy fallback
    _HAVE_PYPDF = True
except Exception:
    pass

try:
    import pytesseract  # OCR engine wrapper (requires system Tesseract)
    from PIL import Image  # pillow: image container for OCR
    _HAVE_TESSERACT = True
except Exception:
    pass

try:
    from pdf2image import convert_from_path  # renders PDFs to images for OCR
    _HAVE_PDF2IMAGE = True
except Exception:
    pass


@dataclass
class TermSpec:
    """Container for a search term and the associated page constraints."""
    term: str               # The phrase/keyword to search for
    pages: List[int]        # Parsed list of 1-indexed page numbers
    pages_raw: str          # Original "Pages" string for reporting


@dataclass
class MatchResult:
    """Per-term match data for a given PDF (EIDP)."""
    pdf_file: str           # File name of the PDF scanned
    serial_number: str      # "SN XXXX" extracted from the filename
    term: str               # Term searched
    page: Optional[int]     # Page number where the closest number was found
    number: Optional[str]   # The closest numeric value found near the term
    units: Optional[str]    # Units detected adjacent to the number (e.g., lbf, sec)
    context: str            # Short snippet of nearby text (for verification)
    method: str             # Extraction pipeline used (e.g., "pymupdf > ocr")
    found: bool             # True if any number was found near the term


# Regex to detect numbers (int/float) with optional thousands separators and units
NUMBER_REGEX = re.compile(
    r"""
    (?<![A-Za-z0-9_.-])           # ensure we aren't inside a larger token on the left
    [-+]?                         # optional sign
    (?:\d{1,3}(?:,\d{3})+|\d+)    # integer with optional thousand separators OR plain digits
    (?:\.\d+)?                    # optional decimal part
    (?:\s?(?:%|ppm|ppb|ms|s|kg|g|mg|ug|lb|lbs|degC|degF|C|F|N|kN|mN|Ns|bar))?  # optional units incl. Newtons

    (?![A-Za-z0-9_.-])            # ensure we aren't inside a larger token on the right
    """,
    re.VERBOSE
)


def numeric_only(value: Optional[str]) -> Optional[str]:
    """Return just the numeric part of a matched value (e.g., "1 N" -> "1").
    Preserves sign and decimals, strips thousands separators.
    If no number is present, returns the original value unchanged.
    """
    if value is None:
        return None
    s = value.replace(" ", " ")
    import re as _re
    m = _re.search(r"[-+]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?", s)
    if not m:
        return value
    return m.group(0).replace(",", "")


def extract_units(value: Optional[str]) -> Optional[str]:
    """Extract a trailing unit token from a matched value.
    Examples: '24 lbf' -> 'lbf', '220 sec' -> 'sec', '100' -> None.
    Matches against the aerospace unit set used by NUMBER_REGEX.
    """
    if not value:
        return None
    s = value.replace("\xa0", " ").strip()
    # unit set mirrors _AERO_UNITS; keep case-insensitive matching
    unit_core = r"%|ppm|ppb|ms|s|sec|kg|g|mg|ug|lbm|lb|lbs|lbf|N|kN|mN|Ns|bar|mbar|Pa|kPa|MPa|psi|psia|psig|mm|cm|m|in|ft|K|degC|degF|C|F"
    # Look for whitespace + unit at the end of the string; use IGNORECASE flag
    m = re.search(r"\s+(" + unit_core + r")$", s, flags=re.IGNORECASE)
    if m:
        return m.group(1)
    return None

# Regex to capture serial numbers like "... SN 1234", "... SN-ABC_09", etc.
SN_REGEX = re.compile(
    r"""\bSN\W*([A-Za-z0-9][A-Za-z0-9_\-]*)""",  # capture the SN id after the "SN" prefix
    re.IGNORECASE | re.VERBOSE
)


# Extend units for aerospace contexts and override NUMBER_REGEX with a richer set.
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


def parse_page_ranges(s: str) -> List[int]:
    """
    Convert a human-friendly page range string (e.g., "5-10, 12; 15â€“18")
    into a sorted list of unique 1-indexed page numbers.

    Supported separators: comma, space, semicolon; supports en-dash and em-dash.
    """
    if not s:
        return []
    s_norm = s.strip().replace("â€“", "-").replace("â€”", "-")
    parts = re.split(r"[,\s;]+", s_norm)
    pages = set()
    for part in parts:
        if not part:
            continue
        if "-" in part:
            # Range "a-b" â†’ expand into all pages between a and b inclusive
            try:
                a, b = part.split("-", 1)
                a = int(re.sub(r"\D", "", a))
                b = int(re.sub(r"\D", "", b))
                if a > b:
                    a, b = b, a
                for p in range(a, b + 1):
                    pages.add(p)
            except Exception:
                # Ignore malformed pieces; we don't want a hard failure here
                continue
        else:
            # Single page number
            try:
                val = int(re.sub(r"\D", "", part))
                pages.add(val)
            except Exception:
                continue
    return sorted(pages)


def load_terms(input_path: Path) -> List[TermSpec]:
    """
    Load the Terms + Pages table from CSV or Excel.
    - CSV: requires column headers 'Term' and 'Pages' (case-insensitive)
    - Excel: requires openpyxl to be installed

    Returns a list of TermSpec objects (term, parsed pages list, original pages string).
    """
    ext = input_path.suffix.lower()
    terms: List[TermSpec] = []

    if ext == ".csv":
        signature = b""
        try:
            with input_path.open("rb") as f_check:
                signature = f_check.read(4)
        except OSError:
            pass

        if signature.startswith(b"PK\x03\x04"):
            print(
                "[WARN] CSV file appears to be an Excel workbook. Attempting Excel parser instead.",
                file=sys.stderr,
            )
        else:
            encodings_to_try = [
                "utf-8-sig",
                "utf-8",
                "utf-16",
                "utf-16-le",
                "utf-16-be",
                "cp1252",
                "latin-1",
            ]

            def read_csv_terms(encoding: str) -> List[TermSpec]:
                with input_path.open(newline="", encoding=encoding) as f:
                    reader = csv.DictReader(f)
                    result: List[TermSpec] = []
                    for row in reader:
                        term = None
                        pages_str = ""
                        for k, v in row.items():
                            if k and k.strip().lower() == "term":
                                term = (v or "").strip()
                            if k and k.strip().lower() == "pages":
                                pages_str = (v or "").strip()
                        if term:
                            result.append(
                                TermSpec(term=term, pages=parse_page_ranges(pages_str), pages_raw=pages_str)
                            )
                    return result

            last_error: Optional[UnicodeDecodeError] = None
            for encoding in encodings_to_try:
                try:
                    terms = read_csv_terms(encoding)
                    if encoding != "utf-8-sig":
                        print(
                            f"[WARN] CSV file decoded using fallback encoding '{encoding}'.",
                            file=sys.stderr,
                        )
                    return terms
                except UnicodeDecodeError as err:
                    last_error = err

            tried = ", ".join(encodings_to_try)
            if last_error:
                print(
                    f"[ERROR] Could not decode CSV file using encodings: {tried}. "
                    f"Last error: {last_error}",
                    file=sys.stderr,
                )
            else:
                print(f"[ERROR] Could not decode CSV file using encodings: {tried}.", file=sys.stderr)
            sys.exit(2)

    # Excel path: requires openpyxl
    try:
        import openpyxl  # type: ignore
    except Exception:
        print(
            "[ERROR] Excel file given but 'openpyxl' is not available. "
            "Install openpyxl or save your spreadsheet as CSV and re-run.",
            file=sys.stderr,
        )
        sys.exit(2)

    wb = openpyxl.load_workbook(str(input_path), data_only=True)
    ws = wb.active

    # Build a map of header name â†’ column index
    header_map: Dict[str, int] = {}
    for col_idx, cell in enumerate(ws[1], start=1):
        key = (str(cell.value) if cell.value is not None else "").strip().lower()
        if key:
            header_map[key] = col_idx

    def col_for(name: str) -> Optional[int]:
        """Return the 1-based column index for a header name, or None if absent."""
        for k, v in header_map.items():
            if k == name.lower():
                return v
        return None

    term_col = col_for("term")
    pages_col = col_for("pages")
    if not term_col:
        print("[ERROR] Could not find 'Term' header in Excel file.", file=sys.stderr)
        sys.exit(2)

    # Walk rows and collect terms
    for row in ws.iter_rows(min_row=2):
        term_val = row[term_col - 1].value if term_col else None
        pages_val = row[pages_col - 1].value if pages_col else "" if pages_col else ""
        term = (str(term_val) if term_val is not None else "").strip()
        pages_str = (str(pages_val) if pages_val is not None else "").strip()
        if term:
            terms.append(TermSpec(term=term, pages=parse_page_ranges(pages_str), pages_raw=pages_str))
    return terms


def extract_pages_text_pymupdf(pdf_path: Path, pages: Sequence[int]) -> Tuple[Dict[int, str], str]:
    """
    Extract text for specific pages using PyMuPDF if available.
    Returns a mapping {page_number: text} and a label for the method used.
    """
    out: Dict[int, str] = {}
    if not _HAVE_PYMUPDF:
        return out, "pymupdf:N/A"
    try:
        doc = fitz.open(str(pdf_path))
    except Exception as e:
        return out, f"pymupdf:open_error:{e}"
    try:
        max_page = doc.page_count
        for p in pages:
            if 1 <= p <= max_page:
                # 'text' mode preserves simple reading order
                txt = doc.load_page(p - 1).get_text("text")
                out[p] = txt or ""
        return out, "pymupdf"
    finally:
        doc.close()


def extract_pages_text_pdfminer(pdf_path: Path, pages: Sequence[int]) -> Tuple[Dict[int, str], str]:
    """
    Extract text using pdfminer.six for selected pages.
    Useful fallback if PyMuPDF fails or returns empty text (e.g., scanned pages).
    """
    out: Dict[int, str] = {}
    if not _HAVE_PDFMINER:
        return out, "pdfminer:N/A"
    for p in pages:
        try:
            text = pdfminer_extract_text(str(pdf_path), page_numbers=[p - 1])
            out[p] = text or ""
        except Exception:
            out[p] = ""
    return out, "pdfminer"


def extract_pages_text_pypdf(pdf_path: Path, pages: Sequence[int]) -> Tuple[Dict[int, str], str]:
    """
    Extract text with pypdf/PyPDF2 for selected pages.
    Light-weight fallback that works on many PDFs but may struggle with layout.
    """
    out: Dict[int, str] = {}
    if not _HAVE_PYPDF:
        return out, "pypdf:N/A"
    try:
        reader = _PdfReader(str(pdf_path))
        max_page = len(reader.pages)
        for p in pages:
            if 1 <= p <= max_page:
                try:
                    txt = reader.pages[p - 1].extract_text() or ""
                except Exception:
                    txt = ""
                out[p] = txt
    except Exception:
        pass
    return out, "pypdf"


def ocr_pages_with_pymupdf(pdf_path: Path, pages: Sequence[int]) -> Tuple[Dict[int, str], str]:
    """
    OCR selected pages using PyMuPDF to render pages to images,
    then pytesseract to extract text. Requires Tesseract and PIL.
    """
    out: Dict[int, str] = {}
    if not (_HAVE_TESSERACT and _HAVE_PYMUPDF):
        return out, "ocr_pymupdf:N/A"
    try:
        doc = fitz.open(str(pdf_path))
    except Exception as e:
        return out, f"ocr_pymupdf:open_error:{e}"
    try:
        for p in pages:
            if 1 <= p <= doc.page_count:
                page = doc.load_page(p - 1)
                pix = page.get_pixmap(dpi=300)  # 300 DPI for decent OCR fidelity
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                try:
                    text = pytesseract.image_to_string(img)
                except Exception:
                    text = ""
                out[p] = text or ""
    finally:
        doc.close()
    return out, "ocr_pymupdf"


def ocr_pages_with_pdf2image(pdf_path: Path, pages: Sequence[int]) -> Tuple[Dict[int, str], str]:
    """
    OCR selected pages by rendering them via pdf2image (Poppler required),
    then pytesseract for text recognition. Requires Tesseract and PIL.
    """
    out: Dict[int, str] = {}
    if not (_HAVE_TESSERACT and _HAVE_PDF2IMAGE):
        return out, "ocr_pdf2image:N/A"
    try:
        images = convert_from_path(str(pdf_path), dpi=300, first_page=min(pages), last_page=max(pages))
    except Exception:
        return out, "ocr_pdf2image:convert_error"
    page_list = sorted(set(pages))
    start = page_list[0]
    for idx, img in enumerate(images, start=start):
        if idx in page_list:
            try:
                text = pytesseract.image_to_string(img)
            except Exception:
                text = ""
            out[idx] = text or ""
    return out, "ocr_pdf2image"


def _normalize_text_for_search(s: str) -> str:
    """Normalize text to improve matching across table/spacing artifacts.
    - Replace non-breaking spaces with regular spaces
    - Convert en/em dashes to hyphen
    - Replace vertical bars with spaces (common in extracted tables)
    - Collapse runs of spaces/tabs while preserving newlines
    """
    if not s:
        return ""
    s = s.replace("\u00A0", " ")
    s = s.replace("–", "-").replace("—", "-")
    s = s.replace("|", " ")
    s = re.sub(r"[ \t\f\r]+", " ", s)
    return s


def extract_pages_text(pdf_path: Path, pages: Sequence[int]) -> Tuple[Dict[int, str], str]:
    """
    Try multiple extraction methods in a fixed order and fill in what we can:
      1) PyMuPDF
      2) pdfminer.six (for empty pages)
      3) pypdf/PyPDF2 (for remaining empties)
      4) OCR (as a last resort if Tesseract is available)
    Return a consolidated {page: text} mapping and a pipeline summary string.
    """
    tried = []

    # Attempt #1: PyMuPDF
    page_text, m = extract_pages_text_pymupdf(pdf_path, pages)
    tried.append(m)

    # Identify which pages are still empty after the first extractor
    empty_pages = [p for p in pages if page_text.get(p, "").strip() == ""]

    # Attempt #2: pdfminer on empty pages
    if empty_pages:
        pt2, m2 = extract_pages_text_pdfminer(pdf_path, empty_pages)
        tried.append(m2)
        for p in empty_pages:
            if (pt2.get(p) or "").strip():
                page_text[p] = pt2[p]
        empty_pages = [p for p in pages if page_text.get(p, "").strip() == ""]

    # Attempt #3: pypdf/PyPDF2 on remaining pages
    if empty_pages:
        pt3, m3 = extract_pages_text_pypdf(pdf_path, empty_pages)
        tried.append(m3)
        for p in empty_pages:
            if (pt3.get(p) or "").strip():
                page_text[p] = pt3[p]
        empty_pages = [p for p in pages if page_text.get(p, "").strip() == ""]

    # Attempt #4: OCR as final fallback
    if empty_pages and _HAVE_TESSERACT:
        if _HAVE_PYMUPDF:
            pt4, m4 = ocr_pages_with_pymupdf(pdf_path, empty_pages)
        else:
            pt4, m4 = ocr_pages_with_pdf2image(pdf_path, empty_pages)
        tried.append(m4)
        for p in empty_pages:
            if (pt4.get(p) or "").strip():
                page_text[p] = pt4[p]

    # Normalize text per page to make downstream term matching more robust
    for _p in list(page_text.keys()):
        page_text[_p] = _normalize_text_for_search(page_text.get(_p, ""))

    pipeline = " > ".join(tried)
    return page_text, pipeline


def find_closest_number_in_text(text: str, term: str, window_chars: int = 160, case_sensitive: bool = False) -> Tuple[Optional[str], Optional[str]]:
    """
    Prefer numbers on the same line to the right of the term, then left,
    then next line, previous line, else fall back to closest in a window.
    Returns (number_string or None, context_snippet or None).
    """
    if not text:
        return None, None

    src = text
    # Normalize case if needed
    hay = src if case_sensitive else src.lower()
    needle = term if case_sensitive else term.lower()

    def _simplify(t: str) -> str:
        t = t.lower()
        t = re.sub(r"\b(nominal|minimum|min|maximum|max|range|typical|average|avg|target|req(?:uirement)?)\b", " ", t)
        t = re.sub(r"\s+", " ", t).strip()
        return t

    # Find positions of the full term
    positions: List[int] = []
    start = 0
    while True:
        idx = hay.find(needle, start)
        if idx == -1:
            break
        positions.append(idx)
        start = idx + max(1, len(needle))

    # Fallback: simplified needle (e.g., "Thrust Nominal" -> "Thrust")
    if not positions:
        alt = _simplify(needle)
        if alt and alt != needle:
            start = 0
            while True:
                idx = hay.find(alt, start)
                if idx == -1:
                    break
                positions.append(idx)
                start = idx + max(1, len(alt))

    if not positions:
        return None, None

    # Pre-compute all numeric spans in the text
    nums = [(m.group(0), m.start(), m.end()) for m in NUMBER_REGEX.finditer(src)]

    def numbers_in(a: int, b: int):
        return [(n, i, j) for (n, i, j) in nums if i >= a and j <= b]

    def snippet(a: int, b: int) -> str:
        return src[max(0, a - 60): min(len(src), b + 60)].replace("\n", " ")

    best_num = None
    best_ctx = None
    best_dist = 10**9

    for pos in positions:
        # Determine line bounds
        lb = src.rfind("\n", 0, pos) + 1
        rb = src.find("\n", pos)
        if rb == -1:
            rb = len(src)
        line_nums = numbers_in(lb, rb)

        # 1) Same line, to the right
        right_side = [(n, i, j) for (n, i, j) in line_nums if i >= pos]
        if right_side:
            n, i, j = min(right_side, key=lambda t: t[1] - pos)
            return n, snippet(i, j)

        # 2) Same line, to the left
        left_side = [(n, i, j) for (n, i, j) in line_nums if j <= pos]
        if left_side:
            n, i, j = max(left_side, key=lambda t: t[2])
            return n, snippet(i, j)

        # 3) Next line
        nlb = rb + 1
        nrb = src.find("\n", nlb)
        if nrb == -1:
            nrb = len(src)
        next_nums = numbers_in(nlb, nrb)
        if next_nums:
            n, i, j = next_nums[0]
            return n, snippet(i, j)

        # 4) Previous line
        plb = src.rfind("\n", 0, lb - 1)
        if plb == -1:
            plb = 0
        else:
            plb = plb + 1
        prb = lb - 1
        prev_nums = numbers_in(plb, prb)
        if prev_nums:
            n, i, j = prev_nums[-1]
            return n, snippet(i, j)

        # 5) Fallback to closest in window
        left = max(0, pos - window_chars)
        right = min(len(src), pos + len(term) + window_chars)
        cand = [(n, i, j) for (n, i, j) in nums if i >= left and j <= right]
        if cand:
            n, i, j = min(cand, key=lambda t: min(abs(t[1]-pos), abs(t[2]-pos)))
            d = min(abs(i - pos), abs(j - pos))
            if d < best_dist:
                best_dist = d
                best_num, best_ctx = n, snippet(i, j)

    return best_num, best_ctx


def get_serial_number_from_filename(pdf_path: Path) -> str:
    """
    Extract the serial number from the PDF filename using SN_REGEX.
    Returns a string like "SN 1234". If no match is found,
    returns a fallback based on the basename (e.g., "SN_<stem>").
    """
    name = pdf_path.stem
    m = SN_REGEX.search(name)
    if m:
        return f"SN {m.group(1)}"
    m = SN_REGEX.search(pdf_path.name)
    if m:
        return f"SN {m.group(1)}"
    return f"SN_{name}"


def scan_pdf_for_term(pdf_path: Path, serial_number: str, term: str, pages: Sequence[int], window_chars: int, case_sensitive: bool) -> MatchResult:
    """
    Scan a single PDF for a single term (restricted to the provided pages).
    - Uses extract_pages_text(...) to build a map of pageâ†’text and a method pipeline string.
    - Calls find_closest_number_in_text(...) to get the nearest number and context.
    - Returns a MatchResult with page/number/context and pipeline details.
    """
    # Build text for constrained pages (or the whole doc if no pages specified)
    page_text_map, pipeline = extract_pages_text(pdf_path, pages if pages else list(range(1, 10000)))

    chosen_page = None
    chosen_number = None
    chosen_ctx = None

    # Search pages in ascending order; stop at the first page where a number is found
    for p in sorted(page_text_map.keys()):
        text = page_text_map[p]
        number, ctx = find_closest_number_in_text(text, term, window_chars=window_chars, case_sensitive=case_sensitive)
        if number:
            chosen_page = p
            chosen_number = number
            chosen_ctx = ctx or ""
            break

    if chosen_number:
        return MatchResult(
            pdf_file=pdf_path.name,
            serial_number=serial_number,
            term=term,
            page=chosen_page,
            number=chosen_number,
            units=extract_units(chosen_number),
            context=chosen_ctx or "",
            method=pipeline,
            found=True
        )
    else:
        return MatchResult(
            pdf_file=pdf_path.name,
            serial_number=serial_number,
            term=term,
            page=None,
            number=None,
            units=None,
            context="",
            method=pipeline,
            found=False
        )


def move_file_safely(src: Path, dst_folder: Path) -> Path:
    """
    Move a file into a destination folder, avoiding collisions by appending (n)
    if the filename already exists. Returns the final destination path.
    """
    dst_folder.mkdir(parents=True, exist_ok=True)
    dst = dst_folder / src.name
    if not dst.exists():
        shutil.move(str(src), str(dst))
        return dst

    # Collision handling: add numeric suffix "(1)", "(2)", ...
    stem = src.stem
    ext = src.suffix
    i = 1
    while True:
        candidate = dst_folder / f"{stem} ({i}){ext}"
        if not candidate.exists():
            shutil.move(str(src), str(candidate))
            return candidate
        i += 1


def write_outputs_excel_or_csv(
    output_xlsx: Path,
    results_matrix: Dict[str, Dict[str, Optional[str]]],
    term_order: List[str],
    term_pages_raw: Dict[str, str],
    metadata_rows: List[Dict],
    csv_fallback_prefix: Path
) -> None:
    """
    Write the "results" and "metadata" outputs to Excel if possible,
    otherwise fall back to two CSV files:
      - <prefix>.results.csv
      - <prefix>.metadata.csv
    """
    if _HAVE_PANDAS and _HAVE_OPENPYXL_OR_XLSXWRITER:
        # Build a consistent list of SN columns across all terms
        serials = set()
        for term in term_order:
            for sn in results_matrix.get(term, {}):
                serials.add(sn)
        serial_cols = sorted(serials)

        # Assemble rows for the "results" DataFrame
        rows = []
        for term in term_order:
            row = {"Term": term, "Pages": term_pages_raw.get(term, "")}
            for sn in serial_cols:
                row[sn] = results_matrix.get(term, {}).get(sn)
            rows.append(row)

        # Create DataFrames
        df_results = pd.DataFrame(rows, columns=["Term", "Pages"] + serial_cols)
        df_meta = pd.DataFrame(metadata_rows)

        # Write Excel with two sheets
        with pd.ExcelWriter(output_xlsx, engine="xlsxwriter") as writer:
            # Sheet 1: wide matrix of results
            df_results.to_excel(writer, sheet_name="results", index=False)
            # Sheet 2: detailed metadata
            df_meta.to_excel(writer, sheet_name="metadata", index=False)

            # Cosmetic improvements: freeze header rows and set reasonable column widths
            ws_res = writer.sheets["results"]
            ws_meta = writer.sheets["metadata"]
            ws_res.freeze_panes(1, 0)
            ws_meta.freeze_panes(1, 0)

            # Auto-size columns based on max content length (capped)
            for i, col in enumerate(df_results.columns):
                width = min(60, max(10, int(df_results[col].astype(str).str.len().max() if not df_results.empty else len(col)) + 2))
                ws_res.set_column(i, i, width)
            for i, col in enumerate(df_meta.columns):
                width = min(60, max(10, int(df_meta[col].astype(str).str.len().max() if not df_meta.empty else len(col)) + 2))
                ws_meta.set_column(i, i, width)

        print(f"[DONE] Excel written -> {output_xlsx}")
        return

    # ---------- CSV fallback path ----------
    results_csv = csv_fallback_prefix.with_suffix(".results.csv")
    metadata_csv = csv_fallback_prefix.with_suffix(".metadata.csv")

    # Build SN column set as above
    serials = set()
    for term in term_order:
        for sn in results_matrix.get(term, {}):
            serials.add(sn)
    serial_cols = sorted(serials)

    # Write "results" CSV
    with results_csv.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Term", "Pages"] + serial_cols)
        for term in term_order:
            row = [term, term_pages_raw.get(term, "")]
            for sn in serial_cols:
                row.append(results_matrix.get(term, {}).get(sn))
            writer.writerow(row)

    # Write "metadata" CSV
    meta_cols = ["pdf_file", "serial_number", "term", "found", "page", "number", "context", "method_pipeline"]
    with metadata_csv.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(meta_cols)
        for r in metadata_rows:
            writer.writerow([
                r.get("pdf_file"),
                r.get("serial_number"),
                r.get("term"),
                r.get("found"),
                r.get("page"),
                r.get("number"),
                r.get("context"),
                r.get("method_pipeline"),
            ])
    print(f"[DONE] CSV fallback written -> {results_csv} and {metadata_csv}")


def run_scan(
    input_path: Path,
    pdf_folder: Path,
    output_csv: Path,
    output_json: Path,
    output_xlsx: Path,
    scanned_folder: Path,
    window_chars: int,
    case_sensitive: bool
) -> None:
    """
    Orchestrate the entire scan:
      1) Load terms from input file
      2) Iterate through PDFs in the import folder
      3) For each (pdf, term) pair, collect the best match and record metadata
      4) Write progress JSON as we go (crash resilience)
      5) Move scanned PDFs to the Scanned Docs folder
      6) Produce the final Excel (or CSVs) and a flat CSV summary
    """
    # Step 1: Load the term specs
    terms = load_terms(input_path)
    if not terms:
        print("[WARN] No terms found in input. Ensure headers 'Term' and 'Pages' exist.")
        return

    # Step 2: Enumerate PDFs to scan
    pdfs = [p for p in pdf_folder.glob("*.pdf")]
    if not pdfs:
        print(f"[WARN] No PDFs found in folder: {pdf_folder}")
        return

    # Decide output location: always write artifacts into Product_Data_File/run_data/<timestamp>
    # Only the aggregate EIDP_data.csv is kept at the top level.
    from datetime import datetime
    exports_dir = Path("Product_Data_File")
    run_dir = exports_dir / "run_data" / datetime.now().strftime("%Y%m%d_%H%M%S")
    run_dir.mkdir(parents=True, exist_ok=True)

    # Reroute output paths into the run_dir regardless of CLI-provided paths.
    output_json = run_dir / "scan_results.json"
    output_csv = run_dir / "scan_results_flat.csv"
    output_xlsx = run_dir / "scan_results.xlsx"
    print(f"[INFO] Outputs will be saved under: {run_dir}")

    # Prepare structures for the wide "results" sheet and the "metadata" sheet
    term_order = [t.term for t in terms]                  # preserve input order
    term_pages_raw = {t.term: t.pages_raw for t in terms} # map term â†’ original "Pages" string
    results_matrix: Dict[str, Dict[str, Optional[str]]] = {t.term: {} for t in terms}  # term â†’ {SN â†’ number}
    metadata_rows: List[Dict] = []  # detailed records per (pdf, term)
    summary: List[Dict] = []        # JSON audit entries

    # Step 3: For each PDF, scan for each term
    for pdf_path in sorted(pdfs):
        # Derive the serial number from the filename
        serial_number = get_serial_number_from_filename(pdf_path)
        print(f"[INFO] Scanning: {pdf_path.name}  ({serial_number})")

        # Search each configured term within the allowed page ranges
        for t in terms:
            res = scan_pdf_for_term(pdf_path, serial_number, t.term, t.pages, window_chars, case_sensitive)
            # Fill the matrix cell for this (term, serial_number)
            results_matrix.setdefault(t.term, {})[serial_number] = numeric_only(res.number)

            # Build metadata record
            meta = {
                "pdf_file": res.pdf_file,
                "serial_number": res.serial_number,
                "term": res.term,
                "found": res.found,
                "page": res.page,
                "number": res.number,
                "units": res.units,
                "context": res.context,
                "method_pipeline": res.method
            }
            metadata_rows.append(meta)
            summary.append(meta)

        # Step 4: Move the scanned PDF to the "Scanned Docs" folder
        try:
            moved_to = move_file_safely(pdf_path, scanned_folder)
            print(f"[INFO] Moved scanned PDF -> {moved_to}")
        except Exception as e:
            print(f"[WARN] Could not move {pdf_path.name} to '{scanned_folder}': {e}")

        # Step 5: Persist JSON progress incrementally (so partial work isn't lost)
        try:
            with output_json.open("w", encoding="utf-8") as jf:
                json.dump(summary, jf, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"[WARN] Could not write JSON during loop: {e}")

    # Step 6: Emit Excel (or CSV fallback) and a flat CSV summary for compatibility
    csv_fallback_prefix = output_xlsx.with_suffix("")
    write_outputs_excel_or_csv(output_xlsx, results_matrix, term_order, term_pages_raw, metadata_rows, csv_fallback_prefix)

    # Legacy flat CSV (one row per (pdf, term))
    try:
        with output_csv.open("w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["pdf_file", "serial_number", "term", "found", "page", "number", "units", "context", "method_pipeline"])
            for row in summary:
                writer.writerow([
                    row["pdf_file"], row["serial_number"], row["term"], row["found"],
                    row["page"], row["number"], row.get("units"), row["context"], row["method_pipeline"]
                ])
        print(f"[DONE] Flat CSV summary -> {output_csv}")
    except Exception as e:
        print(f"[WARN] Could not write flat CSV summary: {e}")

    print(f"[DONE] Details JSON -> {output_json}")

    # Ensure metadata CSV has units column regardless of writer backend
    try:
        meta_csv = output_xlsx.with_suffix('.metadata.csv')
        with meta_csv.open('w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(["pdf_file","serial_number","term","found","page","number","units","context","method_pipeline"])
            for row in metadata_rows:
                writer.writerow([
                    row.get("pdf_file"), row.get("serial_number"), row.get("term"), row.get("found"),
                    row.get("page"), row.get("number"), row.get("units"), row.get("context"), row.get("method_pipeline")
                ])
    except Exception as e:
        print(f"[WARN] Could not rewrite metadata CSV with units: {e}")

    # --- Aggregate export: Product_Data_File/EIDP_data.csv ---
    try:
        exports_dir = Path("Product_Data_File")
        exports_dir.mkdir(parents=True, exist_ok=True)
        agg_path = exports_dir / "EIDP_data.csv"

        def _read_csv(path: Path):
            if not path.exists():
                return ["Term", "Pages"], {}
            rows = {}
            with path.open("r", encoding="utf-8", newline="") as f:
                r = csv.reader(f)
                header = next(r, [])
                for row in r:
                    if not row:
                        continue
                    term = row[0]
                    rows[term] = row
            return header, rows

        def _write_csv(path: Path, header, rows_map):
            with path.open("w", encoding="utf-8", newline="") as f:
                w = csv.writer(f)
                w.writerow(header)
                for term in sorted(rows_map.keys()):
                    w.writerow(rows_map[term])

        header, existing = _read_csv(agg_path)
        # Ensure first two columns
        if not header or header[:2] != ["Term", "Pages"]:
            header = ["Term", "Pages"] + [h for h in header if h not in ("Term", "Pages")]

        # Add new SN columns at the end in discovered order
        new_sns = []
        for term, sn_map in results_matrix.items():
            for sn in sn_map.keys():
                if sn not in header and sn not in new_sns:
                    new_sns.append(sn)
        header = header + new_sns

        # Build a lookup for column index
        col_index = {name: i for i, name in enumerate(header)}

        # Seed rows from existing
        rows_map = {}
        for term, row in existing.items():
            # Pad or trim row to header length
            out = [""] * len(header)
            for i, val in enumerate(row[:len(header)]):
                out[i] = val
            # Fill missing mandatory fields
            if not out[0]:
                out[0] = term
            rows_map[term] = out

        # Merge current results
        for term in term_order:
            if term not in rows_map:
                rows_map[term] = [""] * len(header)
                rows_map[term][col_index["Term"]] = term
                rows_map[term][col_index["Pages"]] = term_pages_raw.get(term, "")
            for sn, val in (results_matrix.get(term) or {}).items():
                if sn in col_index:
                    rows_map[term][col_index[sn]] = val or ""

        _write_csv(agg_path, header, rows_map)
        print(f"[DONE] Aggregate CSV -> {agg_path}")
    except Exception as e:
        print(f"[WARN] Could not update aggregate CSV: {e}")

    # --- Per-run snapshot note ---
    # No copy needed; all artifacts were written directly under run_dir.
    try:
        print(f"[DONE] Run saved under: {run_dir}")
    except Exception:
        pass


def main() -> None:
    """
    CLI entry point.
    Parses arguments and calls run_scan(...).
    """
    parser = argparse.ArgumentParser(
        description="Scan PDFs for terms and nearest numbers, produce a matrix by Serial Number, and move scanned PDFs."
    )
    parser.add_argument("--input", required=True, help="Path to terms file (.csv or .xlsx). Headers: Term, Pages")
    parser.add_argument("--pdf-folder", required=True, help='Folder containing PDFs to scan (e.g., "EIDP import folder")')
    parser.add_argument("--output-csv", default="scan_results_flat.csv", help="Flat CSV summary (legacy)")
    parser.add_argument("--output-json", default="scan_results.json", help="Path to write JSON details")
    parser.add_argument("--output-xlsx", default="scan_results.xlsx", help="Excel workbook with 'results' and 'metadata' sheets")
    parser.add_argument("--scanned-folder", default="Scanned Docs", help="Folder to move scanned PDFs into")
    parser.add_argument("--window-chars", type=int, default=160, help="Search window size around term (Â± chars)")
    parser.add_argument("--case-sensitive", action="store_true", help="Enable case-sensitive term matching")
    args = parser.parse_args()

    # Normalize and validate file/folder paths
    input_path = Path(args.input)
    pdf_folder = Path(args.pdf_folder)
    output_csv = Path(args.output_csv)
    output_json = Path(args.output_json)
    output_xlsx = Path(args.output_xlsx)
    scanned_folder = Path(args.scanned_folder)

    if not input_path.exists():
        print(f"[ERROR] Input file not found: {input_path}", file=sys.stderr)
        sys.exit(1)
    if not pdf_folder.exists():
        print(f"[ERROR] PDF folder not found: {pdf_folder}", file=sys.stderr)
        sys.exit(1)

    # Kick off the pipeline
    run_scan(
        input_path=input_path,
        pdf_folder=pdf_folder,
        output_csv=output_csv,
        output_json=output_json,
        output_xlsx=output_xlsx,
        scanned_folder=scanned_folder,
        window_chars=args.window_chars,
        case_sensitive=args.case_sensitive
    )


if __name__ == "__main__":
    main()

