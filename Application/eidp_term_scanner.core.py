
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
import difflib
import tempfile
import subprocess
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

# --- Output verbosity control (quiet mode) ---
# Honor QUIET=1|true|yes from environment, and optionally --quiet CLI flag (set later).
import builtins as _builtins  # noqa: E402

_QUIET = (os.environ.get("QUIET", "").strip().lower() in ("1", "true", "yes"))

def _should_emit_text(text: str, is_stderr: bool) -> bool:
    # Always allow stderr
    if is_stderr:
        return True
    # Always show critical categories
    if text.startswith("[ERROR]") or text.startswith("[WARN]") or text.startswith("[DONE]"):
        return True
    # Suppress noisy/debug categories when quiet
    if _QUIET:
        if text.startswith("[XY]"):
            return False
        if text.startswith("[CLEANUP]"):
            return False
        if text.startswith("[PROGRESS]"):
            return False
        if text.startswith("[INFO] Pre-extracted"):
            return False
        if text.startswith("[INFO]"):
            return False
    # Default: emit
    return True

_orig_print = _builtins.print

def print(*args, **kwargs):  # type: ignore[override]
    try:
        file = kwargs.get("file", sys.stdout)
    except Exception:
        file = sys.stdout
    try:
        text = str(args[0]) if args else ""
    except Exception:
        text = ""
    if _should_emit_text(text, is_stderr=(file is sys.stderr)):
        return _orig_print(*args, **kwargs)
    return None

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
_HAVE_OCRMYPDF = False
_OCRMYPDF_BIN: Optional[str] = None
_HAVE_PADDLE_OCR = False
_HAVE_EASYOCR = False

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

_HAVE_TESSERACT = False

_HAVE_PDF2IMAGE = False

_HAVE_OCRMYPDF = False

_HAVE_PADDLE_OCR = False

# EasyOCR (pure-Python OCR)
try:
    import easyocr  # type: ignore
    _HAVE_EASYOCR = True
except Exception:
    _HAVE_EASYOCR = False


def _get_ocr_mode() -> str:
    """Return OCR mode: 'fallback' (default), 'ocr_only', or 'no_ocr'.
    Accepts synonyms: 'auto'->fallback, 'none'/'off'->no_ocr, 'ocr'/'only'->ocr_only.
    """
    try:
        m = (os.environ.get('OCR_MODE', '') or '').strip().lower()
    except Exception:
        m = ''
    if m in ('ocr_only', 'ocr', 'only'):
        return 'ocr_only'
    if m in ('no_ocr', 'none', 'off', 'disabled'):
        return 'no_ocr'
    # default
    return 'fallback'


@dataclass
class TermSpec:
    """Term, page constraints, and extraction hints.

    Fields:
    - term, pages, pages_raw
    - mode: nearest | line | table(xy)
    - line/column: XY mode inputs (column may be pipe-separated alternatives)
    - anchor: line mode anchor (defaults to term)
    - field_index: 1-based index of field after anchor in line mode
    - field_split: auto | groups | tokens
    - return_type: number | string
    - range_min/max: numeric filter bounds
    - units_hint: preferred unit tokens (case-insensitive)
    - group_after/group_before: anchor lines bounding the search region vertically
    """
    term: str
    pages: List[int]
    pages_raw: str
    mode: Optional[str] = None
    line: Optional[str] = None
    column: Optional[str] = None
    anchor: Optional[str] = None
    field_index: Optional[int] = None
    field_split: Optional[str] = None
    return_type: Optional[str] = None
    range_min: Optional[float] = None
    range_max: Optional[float] = None
    units_hint: List[str] = field(default_factory=list)
    # New schema helpers
    value_format: Optional[str] = None      # Optional expected value pattern (e.g., tpl-xxxx or /TPL-\d{4}/)
    group_after: Optional[str] = None       # Optional anchor text; only consider matches appearing after this text
    group_before: Optional[str] = None      # Optional anchor text; only consider matches appearing before this text


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
    # Added attributes for reporting
    confidence: Optional[float] = None      # OCR token confidence when available (0..1), else None
    row_label: Optional[str] = None         # Row identifier (e.g., term/line label in XY/line modes)
    column_label: Optional[str] = None      # Column/header identifier in table(XY) modes
    text_source: Optional[str] = None       # 'pdf' vs 'ocr' for the selected page's text
    error_reason: Optional[str] = None      # Reason when found=False


# Regex to detect numbers (int/float) with optional thousands separators and units
NUMBER_REGEX = re.compile(
    r"""
    (?<![A-Za-z0-9_.-])                 # left boundary
    [-+]?                               # optional sign
    (?:\d{1,3}(?:,\d{3})+|\d+)        # integer (with thousands) or plain digits
    (?:\.\d+)?                        # optional decimal
    (?:[eE][+-]?\d+)?                  # optional exponent, e.g., 8E-8
    (?:\s?(?:%|ppm|ppb|ms|s|sec|kg|g|mg|ug|lbm|lb|lbs|lbf|N|kN|mN|Ns|bar|mbar|Pa|kPa|MPa|psi|psia|psig|mm|cm|m|in|ft|K|degC|degF|C|F))?
    (?![A-Za-z0-9_.-])                  # right boundary
    """,
    re.VERBOSE
)

DATE_REGEX = re.compile(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b")



def numeric_only(value: Optional[str]) -> Optional[str]:
    """Return just the numeric part of a matched value (e.g., "1 N" -> "1").

    Preserves sign and decimals, strips thousands separators.

    If no number is present, returns the original value unchanged.

    """

    if value is None:

        return None

    s = value.replace(" ", " ")

    import re as _re

    m = _re.search(r"[-+]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?(?:[eE][+-]?\d+)?", s)
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
    # Look for optional whitespace + unit at the end of the string; use IGNORECASE flag
    m = re.search(r"(?:\s*(" + unit_core + r"))$", s, flags=re.IGNORECASE)
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
    '''Convert a user-supplied page range string into a sorted list of page numbers.'''
    if not s:
        return []
    s_norm = s.strip()
    for dash in (chr(8211), chr(8212)):
        s_norm = s_norm.replace(dash, '-')
    parts = re.split(r"[,\s;]+", s_norm)
    pages: set[int] = set()
    for part in parts:
        if not part:
            continue
        if '-' in part:
            try:
                a, b = part.split('-', 1)
                a = int(re.sub(r"\D", '', a))
                b = int(re.sub(r"\D", '', b))
                if a and b:
                    lo, hi = sorted((a, b))
                    pages.update(range(lo, hi + 1))
                    continue
            except Exception:
                pass
        try:
            pages.add(int(re.sub(r"\D", '', part)))
        except Exception:
            continue
    return sorted(p for p in pages if p > 0)


def _norm_mode(s: Optional[str]) -> Optional[str]:
    if not s:
        return None
    v = s.strip().lower()
    if v in ("table", "xy", "table(xy)"):
        return "table(xy)"
    if v == "line":
        return "line"
    if v in ("nearest", "default"):
        return "nearest"
    return v

def _parse_field_index(v: Optional[str]) -> Optional[int]:
    if v is None:
        return None
    s = str(v).strip().lower()
    if not s:
        return None
    # strip suffixes like 1st, 2nd, 3rd, 4th
    s = re.sub(r"(st|nd|rd|th)$", "", s)
    try:
        n = int(s)
        if n >= 1:
            return n
    except Exception:
        pass
    return None


def _norm_field_split(s: Optional[str]) -> str:
    # Default to 'groups' so fields separated by 3+ spaces/tabs are distinct,
    # and words separated by 1Ã¢â‚¬â€œ2 spaces remain within the same field.
    if not s:
        return "groups"
    v = s.strip().lower()
    if v in ("groups", "tokens", "auto"):
        return v
    return "groups"


def _norm_return_type(s: Optional[str]) -> str:
    if not s:
        return "number"
    v = s.strip().lower()
    if v in ("string", "text"):
        return "string"
    return "number"


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
                        line = None
                        column = None
                        mode = None
                        anchor = None
                        field_index = None
                        field_split = None
                        return_type = None
                        range_min = None
                        range_max = None
                        units_hint: List[str] = []
                        value_format = None
                        group_after = None
                        group_before = None
                        for k, v in row.items():
                            if k and k.strip().lower() == "term":
                                term = (v or "").strip()
                            if k and k.strip().lower() == "pages":
                                pages_str = (v or "").strip()
                            if k and k.strip().lower() == "line":
                                line = ((v or "").strip() or None)
                            if k and k.strip().lower() == "column":
                                column = ((v or "").strip() or None)
                            if k and k.strip().lower() == "mode":
                                mode = _norm_mode(v)
                            if k and k.strip().lower() == "anchor":
                                anchor = ((v or "").strip() or None)
                            if k and k.strip().lower() == "fieldindex":
                                field_index = _parse_field_index(v)
                            if k and k.strip().lower() == "fieldsplit":
                                field_split = _norm_field_split(v)
                            if k and k.strip().lower() == "return":
                                return_type = _norm_return_type(v)
                            if k and k.strip().lower() == "range":
                                rng = (v or "").strip()
                                if rng:
                                    a, b = parse_range(rng)
                                    range_min, range_max = a, b
                            if k and k.strip().lower() == "range (min)":
                                try:
                                    range_min = float(str(v).replace(',', '')) if v not in (None, "") else range_min
                                except Exception:
                                    pass
                            if k and k.strip().lower() == "range (max)":
                                try:
                                    range_max = float(str(v).replace(',', '')) if v not in (None, "") else range_max
                                except Exception:
                                    pass
                            if k and k.strip().lower() == "units":
                                units_hint = parse_units_hint(v)
                            if k and k.strip().lower() in ("format", "value_format"):
                                value_format = ((v or "").strip() or None)
                            if k and k.strip().lower() in ("groupafter", "group_after", "group"):
                                group_after = ((v or "").strip() or None)
                            if k and k.strip().lower() in ("groupbefore", "group_before", "beforegroup", "group_before"):
                                group_before = ((v or "").strip() or None)
                        if term:
                            result.append(TermSpec(term=term,
                                                   pages=parse_page_ranges(pages_str),
                                                   pages_raw=pages_str,
                                                   mode=mode,
                                                   line=line,
                                                   column=column,
                                                   anchor=anchor,
                                                   field_index=field_index,
                                                   field_split=field_split,
                                                   return_type=return_type,
                                                   range_min=range_min,
                                                   range_max=range_max,
                                                   units_hint=units_hint,
                                                   value_format=value_format,
                                                   group_after=group_after,
                                                   group_before=group_before))
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

    # Excel path: .xls via pandas if available; otherwise .xlsx via openpyxl
    if ext == ".xls":
        if not _HAVE_PANDAS:
            print("[ERROR] .xls requires pandas (and xlrd). Save as .xlsx or .csv, or install pandas/xlrd.", file=sys.stderr)
            sys.exit(2)
        try:
            df = pd.read_excel(str(input_path))
        except Exception as e:
            print(f"[ERROR] Could not read .xls: {e}", file=sys.stderr)
            sys.exit(2)
        return _terms_from_dataframe(df)

    try:
        import openpyxl  # type: ignore
    except Exception:
        print(
            "[ERROR] Excel file given but 'openpyxl' is not available. Install openpyxl or save as CSV.",
            file=sys.stderr,
        )
        sys.exit(2)

    wb = openpyxl.load_workbook(str(input_path), data_only=True)
    ws = wb.active

    # Build a map of header name ÃƒÂ¢Ã¢â‚¬Â Ã¢â‚¬â„¢ column index
    header_map: Dict[str, int] = {}
    for col_idx, cell in enumerate(ws[1], start=1):
        key = (str(cell.value) if cell.value is not None else "").strip().lower()
        if key:
            header_map[key] = col_idx

    def col_for(name: str) -> Optional[int]:
        """Return the 1-based column index for a header name, or None if absent.
        Tolerates headers like 'line (x)' or 'column (y)'.
        """
        target = name.lower()
        # exact match
        for k, v in header_map.items():
            if k == target:
                return v
        # tolerant match stripping non-letters
        for k, v in header_map.items():
            kk = re.sub(r"[^a-z]", "", k)
            if kk == target.replace(" ", "") or kk.startswith(target.replace(" ", "")):
                return v
        return None

    term_col = col_for("term")
    pages_col = col_for("pages")
    mode_col = col_for("mode")
    line_col = col_for("line")
    column_col = col_for("column")
    anchor_col = col_for("anchor")
    fieldindex_col = col_for("fieldindex")
    fieldsplit_col = col_for("fieldsplit")
    return_col = col_for("return")
    range_col = col_for("range")
    range_min_col = col_for("range (min)")
    range_max_col = col_for("range (max)")
    units_col = col_for("units")
    fmt_col = col_for("format") or col_for("value_format")
    group_col = col_for("groupafter") or col_for("group_after") or col_for("group")
    group_before_col = col_for("groupbefore") or col_for("group_before") or col_for("beforegroup")
    if not term_col:
        print("[ERROR] Could not find 'Term' header in Excel file.", file=sys.stderr)
        sys.exit(2)

    # Walk rows and collect terms
    for row in ws.iter_rows(min_row=2):
        term_val = row[term_col - 1].value if term_col else None
        pages_val = row[pages_col - 1].value if pages_col else "" if pages_col else ""
        mode_val = row[mode_col - 1].value if mode_col else None
        line_val = row[line_col - 1].value if line_col else None
        column_val = row[column_col - 1].value if column_col else None
        anchor_val = row[anchor_col - 1].value if anchor_col else None
        fieldindex_val = row[fieldindex_col - 1].value if fieldindex_col else None
        fieldsplit_val = row[fieldsplit_col - 1].value if fieldsplit_col else None
        return_val = row[return_col - 1].value if return_col else None
        range_val = row[range_col - 1].value if range_col else None
        rmin_val = row[range_min_col - 1].value if range_min_col else None
        rmax_val = row[range_max_col - 1].value if range_max_col else None
        units_val = row[units_col - 1].value if units_col else None
        fmt_val = row[fmt_col - 1].value if fmt_col else None if fmt_col else None
        grp_val = row[group_col - 1].value if group_col else None
        grp_before_val = row[group_before_col - 1].value if group_before_col else None
        term = (str(term_val) if term_val is not None else "").strip()
        pages_str = (str(pages_val) if pages_val is not None else "").strip()
        mode = _norm_mode(str(mode_val) if mode_val is not None else None)
        line = (str(line_val).strip() if line_val is not None and str(line_val).strip() else None)
        column = (str(column_val).strip() if column_val is not None and str(column_val).strip() else None)
        anchor = (str(anchor_val).strip() if anchor_val is not None and str(anchor_val).strip() else None)
        field_index = _parse_field_index(str(fieldindex_val) if fieldindex_val is not None else None)
        field_split = _norm_field_split(str(fieldsplit_val) if fieldsplit_val is not None else None)
        return_type = _norm_return_type(str(return_val) if return_val is not None else None)
        rmin = rmax = None
        if range_val is not None and str(range_val).strip():
            rmin, rmax = parse_range(str(range_val).strip())
        if rmin_val is not None and str(rmin_val).strip():
            try:
                rmin = float(str(rmin_val).replace(',', ''))
            except Exception:
                pass
        if rmax_val is not None and str(rmax_val).strip():
            try:
                rmax = float(str(rmax_val).replace(',', ''))
            except Exception:
                pass
        units_hint = parse_units_hint(units_val)
        if term:
            terms.append(TermSpec(term=term, pages=parse_page_ranges(pages_str), pages_raw=pages_str,
                                  mode=mode, line=line, column=column, anchor=anchor,
                                  field_index=field_index, field_split=field_split, return_type=return_type,
                                  range_min=rmin, range_max=rmax, units_hint=units_hint,
                                   value_format=(str(fmt_val).strip() if fmt_val is not None and str(fmt_val).strip() else None),
                                   group_after=(str(grp_val).strip() if grp_val is not None and str(grp_val).strip() else None),
                                   group_before=(str(grp_before_val).strip() if grp_before_val is not None and str(grp_before_val).strip() else None)))
    return terms


def _terms_from_dataframe(df) -> List[TermSpec]:
    cols = {str(c).strip().lower(): c for c in df.columns}
    def get(row, key):
        col = cols.get(key)
        if col is None:
            return None
        return row.get(col)
    out: List[TermSpec] = []
    for _, row in df.iterrows():
        term = str(get(row, 'term') or '').strip()
        if not term:
            continue
        pages_str = str(get(row, 'pages') or '').strip()
        mode = _norm_mode(str(get(row, 'mode') or '').strip() or None)
        line = str(get(row, 'line') or '').strip() or None
        column = str(get(row, 'column') or '').strip() or None
        anchor = str(get(row, 'anchor') or '').strip() or None
        field_index = _parse_field_index(str(get(row, 'fieldindex') or '').strip() or None)
        field_split = _norm_field_split(str(get(row, 'fieldsplit') or '').strip() or None)
        return_type = _norm_return_type(str(get(row, 'return') or '').strip() or None)
        rng = str(get(row, 'range') or '').strip()
        rmin = rmax = None
        if rng:
            rmin, rmax = parse_range(rng)
        # Direct min/max override
        _rmin = str(get(row, 'range (min)') or '').strip()
        _rmax = str(get(row, 'range (max)') or '').strip()
        if _rmin:
            try:
                rmin = float(_rmin.replace(',', ''))
            except Exception:
                pass
        if _rmax:
            try:
                rmax = float(_rmax.replace(',', ''))
            except Exception:
                pass
        units_hint = parse_units_hint(get(row, 'units'))
        value_format = str(get(row, 'format') or get(row, 'value_format') or '').strip() or None
        group_after = str(get(row, 'groupafter') or get(row, 'group_after') or get(row, 'group') or '').strip() or None
        group_before = str(get(row, 'groupbefore') or get(row, 'group_before') or get(row, 'beforegroup') or '').strip() or None
        out.append(TermSpec(term=term, pages=parse_page_ranges(pages_str), pages_raw=pages_str,
                            mode=mode, line=line, column=column, anchor=anchor,
                            field_index=field_index, field_split=field_split, return_type=return_type,
                            range_min=rmin, range_max=rmax, units_hint=units_hint,
                            value_format=value_format, group_after=group_after, group_before=group_before))
    return out

def parse_units_hint(v) -> List[str]:
    if v is None:
        return []
    s = str(v).strip()
    if not s:
        return []
    parts = re.split(r"[|/,;]+", s)
    return [p.strip() for p in parts if p.strip()]

def parse_range(s: str) -> Tuple[Optional[float], Optional[float]]:
    s = s.strip()
    m = re.match(r"^\s*([+-]?[\d,.]+(?:[eE][+-]?\d+)?)\s*\.\.\s*([+-]?[\d,.]+(?:[eE][+-]?\d+)?)\s*$", s)
    if m:
        def to_f(x):
            try:
                return float(str(x).replace(',', ''))
            except Exception:
                return None
        return to_f(m.group(1)), to_f(m.group(2))
    return None, None


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
    error_notes: List[str] = []
    # Tunables via environment
    try:
        _dpi = int(os.environ.get('OCR_DPI', '400'))
    except Exception:
        _dpi = 400
    _dpi = max(200, min(800, _dpi))
    _tess_cfg = os.environ.get('TESSERACT_ARGS', '--psm 6')
    try:
        for p in pages:
            if 1 <= p <= doc.page_count:
                page = doc.load_page(p - 1)
                # Increase DPI to improve OCR fidelity on small text
                pix = page.get_pixmap(dpi=_dpi)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                try:
                    text = pytesseract.image_to_string(img, lang='eng', config=_tess_cfg)
                except Exception as e:
                    error_notes.append(type(e).__name__)
                    text = ""
                out[p] = text or ""
    finally:
        doc.close()
    if error_notes:
        return out, "ocr_pymupdf:error:" + ",".join(sorted(set(error_notes)))
    return out, "ocr_pymupdf"


def ocr_pages_with_pdf2image(pdf_path: Path, pages: Sequence[int]) -> Tuple[Dict[int, str], str]:
    """
    OCR selected pages by rendering them via pdf2image (Poppler required),
    then pytesseract for text recognition. Requires Tesseract and PIL.
    """
    out: Dict[int, str] = {}
    if not (_HAVE_TESSERACT and _HAVE_PDF2IMAGE):
        return out, "ocr_pdf2image:N/A"
    # Tunables via environment
    try:
        _dpi = int(os.environ.get('OCR_DPI', '400'))
    except Exception:
        _dpi = 400
    _dpi = max(200, min(800, _dpi))
    _tess_cfg = os.environ.get('TESSERACT_ARGS', '--psm 6')
    try:
        images = convert_from_path(str(pdf_path), dpi=_dpi, first_page=min(pages), last_page=max(pages))
    except Exception as e:
        return out, f"ocr_pdf2image:convert_error:{e}"
    page_list = sorted(set(pages))
    start = page_list[0]
    error_notes: List[str] = []
    for idx, img in enumerate(images, start=start):
        if idx in page_list:
            try:
                text = pytesseract.image_to_string(img, lang='eng', config=_tess_cfg)
            except Exception as e:
                error_notes.append(type(e).__name__)
                text = ""
            out[idx] = text or ""
    if error_notes:
        return out, "ocr_pdf2image:error:" + ",".join(sorted(set(error_notes)))
    return out, "ocr_pdf2image"


def ocr_pages_with_easyocr(pdf_path: Path, pages: Sequence[int]) -> Tuple[Dict[int, str], str]:
    """OCR selected pages using EasyOCR (CPU) with PyMuPDF rendering.

    Returns page->concatenated text and a pipeline label.
    - Tries env-configured languages (EASYOCR_LANGS/OCR_LANGS), then falls back to ['en'] on init errors.
    - Includes error message in pipeline if initialization ultimately fails.
    """
    out: Dict[int, str] = {}
    if not (_HAVE_EASYOCR and _HAVE_PYMUPDF):
        return out, "ocr_easyocr:N/A"
    # Reader languages from env; comma/semicolon separated
    langs_raw = (os.environ.get('EASYOCR_LANGS') or os.environ.get('OCR_LANGS') or 'en')
    langs = [s.strip() for s in re.split(r'[;,]', langs_raw) if s.strip()]
    used_langs_label = ",".join(langs or ['en'])
    reader = None
    # Suppress noisy CPU-only torch dataloader warnings about pin_memory
    try:
        import warnings as _warn
        _warn.filterwarnings(
            "ignore",
            message=r".*pin_memory.*",
            category=UserWarning,
            module=r"torch\.utils\.data\.dataloader",
        )
    except Exception:
        pass
    try:
        reader = easyocr.Reader(langs or ['en'], gpu=False, verbose=False)  # type: ignore
    except Exception as e:
        # Retry with a safe default language set to avoid env/config errors
        err_msg = f"{type(e).__name__}:{e}".replace("\n", " ")
        try:
            reader = easyocr.Reader(['en'], gpu=False, verbose=False)  # type: ignore
            used_langs_label = "en"
        except Exception as e2:
            err2 = f"{type(e2).__name__}:{e2}".replace("\n", " ")
            return out, f"ocr_easyocr:init_error:{err_msg}"

    try:
        dpi = int(os.environ.get('OCR_DPI', '600'))
    except Exception:
        dpi = 600
    dpi = max(200, min(900, dpi))

    try:
        doc = fitz.open(str(pdf_path))  # type: ignore[name-defined]
    except Exception as e:
        return out, f"ocr_easyocr:open_error:{e}"

    try:
        for p in pages:
            if 1 <= p <= doc.page_count:
                try:
                    page = doc.load_page(p - 1)
                    pix = page.get_pixmap(dpi=dpi)
                except Exception:
                    out[p] = ""
                    continue
                # Save to a temp PNG to feed reader
                try:
                    import tempfile
                    import os as _os
                    tmp_dir = tempfile.mkdtemp(prefix="easyocr_")
                    img_path = Path(tmp_dir) / f"page_{p}.png"
                    pix.save(str(img_path))
                    # Run OCR
                    try:
                        results = reader.readtext(str(img_path), detail=1)  # list of [bbox, text, conf]
                    except Exception:
                        results = []
                    # Join text lines in reading order
                    lines: List[str] = []
                    for item in results:
                        try:
                            _, t, c = item
                            if isinstance(t, str) and t.strip():
                                lines.append(t)
                        except Exception:
                            pass
                    out[p] = "\n".join(lines)
                finally:
                    try:
                        import shutil as _sh
                        _sh.rmtree(tmp_dir, ignore_errors=True)  # type: ignore
                    except Exception:
                        pass
    finally:
        try:
            doc.close()
        except Exception:
            pass

    return out, f"ocr_easyocr({used_langs_label},dpi={dpi})"


_EASYOCR_CACHE: Dict[Tuple[str, int, str, int], List[Dict[str, float]]] = {}
_EASYOCR_READER_CACHE: Dict[str, object] = {}
_PAGE_TEXT_CACHE: Dict[str, Tuple[Dict[int, str], str, int]] = {}

def _pdf_cache_key(pdf_path: Path) -> str:
    try:
        return str(pdf_path.resolve())
    except Exception:
        return str(pdf_path)

def get_pdf_page_count(pdf_path: Path) -> int:
    """Best-effort page count using PyMuPDF or pypdf."""
    if _HAVE_PYMUPDF:
        try:
            doc = fitz.open(str(pdf_path))  # type: ignore[name-defined]
            try:
                return int(getattr(doc, 'page_count', getattr(doc, 'pageCount', 0)) or 0)
            finally:
                try:
                    doc.close()
                except Exception:
                    pass
        except Exception:
            pass
    if _HAVE_PYPDF:
        try:
            reader = _PdfReader(str(pdf_path))  # type: ignore[name-defined]
            return int(len(getattr(reader, 'pages', [])))
        except Exception:
            pass
    return 0


def _update_run_registry(run_dir: Path, serial_numbers: List[str]) -> None:
    """Update a persistent Excel registry of EIDPs (serial_numbers) and their latest run date.

    - File path: Product_Data_File/run_registry.xlsx (CSV fallback if Excel writer unavailable)
    - Columns: serial_number, run_date, run_folder
    - On re-run, replaces the row for a serial number with the latest date and folder
    """
    try:
        exports_dir = Path("Product_Data_File")
        exports_dir.mkdir(parents=True, exist_ok=True)
        registry_xlsx = exports_dir / "run_registry.xlsx"
        registry_csv = exports_dir / "run_registry.csv"

        # Build rows to merge
        from datetime import datetime
        run_folder = run_dir
        run_date = None
        try:
            # Prefer timestamp parsed from folder name
            stamp = run_dir.name
            dt = datetime.strptime(stamp, "%Y%m%d_%H%M%S")
            run_date = dt.strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            run_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        new_rows = {sn: {"serial_number": sn, "run_date": run_date, "run_folder": str(run_folder)} for sn in serial_numbers}

        # If pandas + writer are available, maintain Excel; else maintain CSV
        if _HAVE_PANDAS and _HAVE_OPENPYXL_OR_XLSXWRITER:
            try:
                import pandas as _pd
                if registry_xlsx.exists():
                    try:
                        df = _pd.read_excel(registry_xlsx)
                    except Exception:
                        df = _pd.DataFrame(columns=["serial_number", "run_date", "run_folder"])
                else:
                    df = _pd.DataFrame(columns=["serial_number", "run_date", "run_folder"])
                # Index by serial_number and update
                if "serial_number" not in df.columns:
                    df = _pd.DataFrame(columns=["serial_number", "run_date", "run_folder"])
                df = df.set_index("serial_number", drop=False)
                for sn, row in new_rows.items():
                    df.loc[sn] = row
                # Sort by run_date desc for convenience (optional)
                try:
                    df_sorted = df.sort_values(by=["run_date", "serial_number"], ascending=[False, True])
                except Exception:
                    df_sorted = df
                with _pd.ExcelWriter(registry_xlsx, engine="xlsxwriter") as writer:
                    df_sorted.to_excel(writer, sheet_name="runs", index=False)
                    ws = writer.sheets["runs"]
                    ws.freeze_panes(1, 0)
                    for i, col in enumerate(df_sorted.columns):
                        try:
                            max_len = int(df_sorted[col].astype(str).map(len).max()) if not df_sorted.empty else len(col)
                        except Exception:
                            max_len = len(col)
                        ws.set_column(i, i, min(80, max(12, max_len + 2)))
                return
            except Exception:
                # Fall back to CSV path
                pass

        # CSV fallback path
        try:
            rows_map: Dict[str, Dict[str, str]] = {}
            if registry_csv.exists():
                with registry_csv.open("r", encoding="utf-8", newline="") as f:
                    r = csv.DictReader(f)
                    for row in r:
                        sn = (row.get("serial_number") or "").strip()
                        if sn:
                            rows_map[sn] = row
            for sn, row in new_rows.items():
                rows_map[sn] = row
            with registry_csv.open("w", encoding="utf-8", newline="") as f:
                w = csv.DictWriter(f, fieldnames=["serial_number", "run_date", "run_folder"])
                w.writeheader()
                for sn in sorted(rows_map.keys()):
                    w.writerow(rows_map[sn])
        except Exception:
            pass
    except Exception:
        # Never block the main run on registry updates
        pass

def _get_easyocr_reader(langs: List[str]):
    key = ",".join(langs or ['en'])
    rdr = _EASYOCR_READER_CACHE.get(key)
    if rdr is not None:
        return rdr
    try:
        # Suppress noisy CPU-only torch dataloader warnings about pin_memory
        try:
            import warnings as _warn
            _warn.filterwarnings(
                "ignore",
                message=r".*pin_memory.*",
                category=UserWarning,
                module=r"torch\.utils\.data\.dataloader",
            )
        except Exception:
            pass
        rdr = easyocr.Reader(langs or ['en'], gpu=False, verbose=False)  # type: ignore
        _EASYOCR_READER_CACHE[key] = rdr
        return rdr
    except Exception:
        # Fallback to a safe default language set
        try:
            rdr = easyocr.Reader(['en'], gpu=False, verbose=False)  # type: ignore
            _EASYOCR_READER_CACHE['en'] = rdr
            return rdr
        except Exception:
            return None

def _get_easyocr_boxes_page(pdf_path: Path, page: int, dpi: int, langs: List[str]) -> List[Dict[str, float]]:
    if not (_HAVE_EASYOCR and _HAVE_PYMUPDF):
        return []
    cache_key = (str(pdf_path), dpi, ",".join(langs or ['en']), page)
    if cache_key in _EASYOCR_CACHE:
        return _EASYOCR_CACHE[cache_key]
    reader = _get_easyocr_reader(langs)
    if reader is None:
        return []
    try:
        doc = fitz.open(str(pdf_path))  # type: ignore[name-defined]
    except Exception:
        return []
    try:
        if 1 <= page <= doc.page_count:
            try:
                pg = doc.load_page(page - 1)
                pix = pg.get_pixmap(dpi=dpi)
            except Exception:
                return []
            import tempfile, shutil
            tmp_dir = Path(tempfile.mkdtemp(prefix='easyocr_xy_'))
            img_path = tmp_dir / ('page_%d.png' % page)
            try:
                pix.save(str(img_path))
                try:
                    res = reader.readtext(str(img_path), detail=1)  # type: ignore[attr-defined]
                except Exception:
                    res = []
                items: List[Dict[str, float]] = []
                for it in res:
                    try:
                        bbox, text, conf = it
                        xs = [float(pt[0]) for pt in bbox]
                        ys = [float(pt[1]) for pt in bbox]
                        x0, y0, x1, y1 = min(xs), min(ys), max(xs), max(ys)
                        cx = (x0 + x1) / 2.0
                        cy = (y0 + y1) / 2.0
                        if isinstance(text, str) and text.strip():
                            items.append({'x0': x0, 'y0': y0, 'x1': x1, 'y1': y1, 'cx': cx, 'cy': cy, 'text': text.strip(), 'conf': float(conf) if conf is not None else 0.0})
                    except Exception:
                        pass
                _EASYOCR_CACHE[cache_key] = items
                return items
            finally:
                try:
                    shutil.rmtree(str(tmp_dir), ignore_errors=True)
                except Exception:
                    pass
        return []
    finally:
        try:
            doc.close()
        except Exception:
            pass

def _easyocr_boxes_for_pages(pdf_path: Path, pages: Sequence[int], dpi: int, langs: List[str]) -> Dict[int, List[Dict[str, float]]]:
    boxes: Dict[int, List[Dict[str, float]]] = {}
    for p in pages:
        items = _get_easyocr_boxes_page(pdf_path, p, dpi=dpi, langs=langs)
        if items:
            boxes[p] = items
    return boxes


def _fuzzy_ratio(a: str, b: str) -> float:
    a2 = re.sub(r"\s+", " ", a or '').strip().lower()
    b2 = re.sub(r"\s+", " ", b or '').strip().lower()
    return difflib.SequenceMatcher(None, a2, b2).ratio()


def _normalize_anchor_token(text: Optional[str]) -> str:
    if not text:
        return ""
    return re.sub(r"[^a-z0-9]", "", str(text).lower())


def _first_numeric(text: str) -> Optional[str]:
    nums = [m.group(0) for m in NUMBER_REGEX.finditer(text)]
    nums += [m.group(0) for m in DATE_REGEX.finditer(text)]
    return nums[0] if nums else None


def _split_fields_by_spacing(line: str) -> List[str]:
    stripped = line.strip()
    if not stripped:
        return []
    parts = [p.strip() for p in re.split(r"\s{2,}", stripped) if p.strip()]
    if len(parts) <= 1:
        parts = [p.strip() for p in re.split(r"\s+", stripped) if p.strip()]
    return parts


def _match_header_index(headers: List[str], col_alts: List[str]) -> int:
    best_idx = -1
    best_score = 0.0
    for idx, header in enumerate(headers):
        if not header:
            continue
        header_norm = _normalize_anchor_token(header)
        for alt in col_alts:
            if not alt:
                continue
            score = _fuzzy_ratio(header, alt)
            alt_norm = _normalize_anchor_token(alt)
            if alt_norm and alt_norm in header_norm:
                score = max(score, 0.99)
            if score > best_score:
                best_idx = idx
                best_score = score
    return best_idx if best_score >= 0.6 else -1


def _extract_value_from_text_table(text: str, row_text: str, column_text: str, case_sensitive: bool) -> Optional[Tuple[str, str, str]]:
    if not row_text or not column_text:
        return None
    lines = [ln for ln in text.splitlines() if ln.strip()]
    if not lines:
        return None
    col_alts = [c.strip() for c in re.split(r"[|/]", column_text) if c.strip()] or [column_text]
    row_norm_target = _normalize_anchor_token(row_text)
    for idx, line in enumerate(lines):
        headers = _split_fields_by_spacing(line)
        if not headers:
            continue
        col_idx = _match_header_index(headers, col_alts)
        if col_idx < 0:
            continue
        for row_line in lines[idx + 1:]:
            if not row_line.strip():
                break
            score = _fuzzy_ratio(row_line, row_text)
            row_norm_line = _normalize_anchor_token(row_line)
            if row_norm_target and row_norm_target in row_norm_line:
                score = max(score, 0.99)
            if score < 0.6:
                continue
            fields = _split_fields_by_spacing(row_line)
            if len(fields) <= col_idx:
                continue
            value = fields[col_idx].strip()
            if value:
                return value, row_line.strip(), headers[col_idx].strip()
        # stop scanning further down once we hit a blank separator
    return None


def _slice_text_by_groups(text: str, group_after: Optional[str], group_before: Optional[str], case_sensitive: bool) -> Tuple[str, bool, bool]:
    if not text:
        return "", not bool(group_after), False
    start_idx = 0
    after_hit = not bool(group_after)
    if group_after:
        idx = _locate_group_anchor(text, group_after, case_sensitive, after=True)
        if idx is None:
            return "", False, False
        start_idx = idx
        after_hit = True
    end_idx = len(text)
    before_hit = False
    if group_before:
        idx = _locate_group_anchor(text[start_idx:], group_before, case_sensitive, after=False)
        if idx is not None:
            end_idx = start_idx + idx
            before_hit = True
    return text[start_idx:end_idx], after_hit, before_hit


def _locate_group_anchor(text: str, anchor: Optional[str], case_sensitive: bool, *, after: bool) -> Optional[int]:
    if not anchor:
        return None
    hay = text if case_sensitive else text.lower()
    needle = anchor if case_sensitive else anchor.lower()
    idx = hay.find(needle)
    if idx >= 0:
        return idx + (len(needle) if after else 0)
    anchor_norm = _normalize_anchor_token(anchor)
    best: Optional[Tuple[float, int, int]] = None  # score, offset, seg_len
    offset = 0
    for segment in text.splitlines(keepends=True):
        raw = segment.rstrip("\r\n")
        score = _fuzzy_ratio(raw, anchor)
        seg_norm = _normalize_anchor_token(raw)
        if anchor_norm and anchor_norm in seg_norm:
            score = max(score, 0.99)
        if score >= 0.6:
            if best is None or score > best[0]:
                best = (score, offset, len(segment))
        offset += len(segment)
    if not best:
        return None
    _, start, seg_len = best
    return start + seg_len if after else start


def scan_pdf_for_term_xy_easyocr(pdf_path: Path, serial_number: str, spec: TermSpec, window_chars: int, case_sensitive: bool) -> Optional[MatchResult]:
    if not (_HAVE_EASYOCR and _HAVE_PYMUPDF):
        return None
    langs_raw = (os.environ.get('EASYOCR_LANGS') or os.environ.get('OCR_LANGS') or 'en')
    langs = [s.strip() for s in re.split(r'[;,]', langs_raw) if s.strip()]
    try:
        fuzz = float(os.environ.get('XY_FUZZ', '0.75'))
    except Exception:
        fuzz = 0.75

    row_name = (spec.line or spec.term or '').strip()
    col_raw = (spec.column or '').strip()
    col_alts = [s.strip() for s in re.split(r'[|/]', col_raw) if s.strip()] or [(spec.column or '').strip()]
    ret_type = (spec.return_type or 'number').strip().lower()
    fmt_pat = _compile_value_regex(spec.value_format or '') if spec.value_format else None

    try:
        dpi_base = int(os.environ.get('OCR_DPI', '700'))
    except Exception:
        dpi_base = 700
    dpi_candidates = [dpi_base]
    if dpi_base > 700:
        dpi_candidates.append(700)

    try:
        doc = fitz.open(str(pdf_path))  # type: ignore[name-defined]
    except Exception:
        doc = None

    pages = spec.pages if spec.pages else ([] if doc is None else list(range(1, doc.page_count + 1)))

    for dpi in dpi_candidates:
        after_found = not bool(spec.group_after)
        before_triggered = False
        for p in pages:
            if before_triggered:
                break
            items = _easyocr_boxes_for_pages(pdf_path, [p], dpi=dpi, langs=langs).get(p, [])
            if not items:
                continue

        # Optional grouping anchor: require row below this text if provided
        group_anchor_y = None
        group_upper_y = None
        if spec.group_after:
            try:
                anchor_norm = _normalize_anchor_token(spec.group_after)
                ga_thresh = max(0.45, fuzz - 0.2)
                matches: List[Tuple[Dict[str, float], float]] = []
                for it in items:
                    txt = str(it.get('text') or '')
                    score = _fuzzy_ratio(txt, spec.group_after)
                    txt_norm = _normalize_anchor_token(txt)
                    if anchor_norm and anchor_norm in txt_norm:
                        score = max(score, 0.99)
                    else:
                        try:
                            norm_ratio = difflib.SequenceMatcher(None, txt_norm, anchor_norm).ratio() if anchor_norm else 0.0
                        except Exception:
                            norm_ratio = 0.0
                        if norm_ratio >= 0.7:
                            score = max(score, norm_ratio)
                    if score >= ga_thresh:
                        matches.append((it, score))
                if matches:
                    best_score = max(m[1] for m in matches)
                    top_matches = [m for m in matches if m[1] >= best_score - 0.1]
                    group_anchor_y = min(top_matches, key=lambda t: t[0]['cy'])[0]['cy']
                    after_found = True
            except Exception:
                group_anchor_y = None
        if spec.group_before:
            try:
                anchor_norm = _normalize_anchor_token(spec.group_before)
                gb_thresh = max(0.45, fuzz - 0.2)
                matches: List[Tuple[Dict[str, float], float]] = []
                for it in items:
                    txt = str(it.get('text') or '')
                    score = _fuzzy_ratio(txt, spec.group_before)
                    txt_norm = _normalize_anchor_token(txt)
                    if anchor_norm and anchor_norm in txt_norm:
                        score = max(score, 0.99)
                    else:
                        try:
                            norm_ratio = difflib.SequenceMatcher(None, txt_norm, anchor_norm).ratio() if anchor_norm else 0.0
                        except Exception:
                            norm_ratio = 0.0
                        if norm_ratio >= 0.7:
                            score = max(score, norm_ratio)
                    if score >= gb_thresh:
                        matches.append((it, score))
                if matches and group_anchor_y is not None:
                    matches = [m for m in matches if m[0]['cy'] > group_anchor_y + 1.0] or matches
                if matches:
                    best_score = max(m[1] for m in matches)
                    top_matches = [m for m in matches if m[1] >= best_score - 0.1]
                    group_upper_y = min(top_matches, key=lambda t: t[0]['cy'])[0]['cy']
                    before_triggered = True
            except Exception:
                group_upper_y = None
        if spec.group_after and not after_found:
            continue

        # Find best row and column headers
        row_candidates = [(it, _fuzzy_ratio(it['text'], row_name)) for it in items if row_name]
        row_candidates = [t for t in row_candidates if t[1] >= fuzz]
        sandwich_eps = 0.5
        if group_anchor_y is not None:
            row_candidates = [
                t for t in row_candidates
                if float(t[0].get('y1', t[0].get('cy', 0.0))) > group_anchor_y + sandwich_eps
            ]
        if group_upper_y is not None:
            row_candidates = [
                t for t in row_candidates
                if float(t[0].get('y0', t[0].get('cy', 0.0))) < group_upper_y - sandwich_eps
            ]
        if not row_candidates:
            continue
        row_it, _ = max(row_candidates, key=lambda t: t[1])

        row_label_right = float(row_it.get('x1', row_it.get('cx', 0.0) or 0.0))

        col_cands: List[Tuple[float, float, Dict[str, float]]] = []  # (dy, -score, header_it)
        for alt in col_alts:
            for it in items:
                sc = _fuzzy_ratio(it['text'], alt)
                if sc >= fuzz and it.get('cy', 0) < row_it.get('cy', 0) and it.get('cx', 0) >= row_label_right:
                    dy = row_it['cy'] - it['cy']
                    col_cands.append((dy, -sc, it))
        if not col_cands:
            for alt in col_alts:
                for it in items:
                    sc = _fuzzy_ratio(it['text'], alt)
                    if sc >= fuzz and it.get('cy', 0) < row_it.get('cy', 0):
                        dy = row_it['cy'] - it['cy']
                        col_cands.append((dy, -sc, it))
        if not col_cands:
            for alt in col_alts:
                for it in items:
                    sc = _fuzzy_ratio(it['text'], alt)
                    if sc >= fuzz and it.get('cx', 0) >= row_label_right:
                        dy = max(0.0, row_it['cy'] - it['cy'])
                        col_cands.append((dy, -sc, it))
        if not col_cands:
            continue

        col_cands.sort(key=lambda t: (t[0], t[1]))
        _, _, hdr = col_cands[0]

        next_row_y = None
        for cand in items:
            try:
                cx_cand = cand.get('cx', 0.0)
                cy_cand = cand.get('cy', 0.0)
            except Exception:
                continue
            if cx_cand >= row_label_right:
                continue
            if cy_cand <= row_it.get('cy', 0.0):
                continue
            txt_c = str(cand.get('text') or '')
            if not txt_c or not any(ch.isalpha() for ch in txt_c):
                continue
            if next_row_y is None or cy_cand < next_row_y:
                next_row_y = cy_cand

        row_h = max(1.0, (row_it['y1'] - row_it['y0']))
        col_w = max(1.0, (hdr['x1'] - hdr['x0']))
        header_h = max(1.0, (hdr['y1'] - hdr['y0']))
        col_half_width = max(col_w, row_h * 1.2, 25.0)
        col_half_height = max(row_h * 0.6, header_h * 0.6, 8.0)
        y_min = max(hdr['y1'], row_it['cy'] - col_half_height)
        y_max = row_it['cy'] + col_half_height
        x_min = hdr['cx'] - col_half_width
        x_max = hdr['cx'] + col_half_width
        ix, iy = hdr['cx'], row_it['cy']

        candidates: List[Tuple[Tuple[int, float], Dict[str, float], str, str]] = []
        for it in items:
            if not (y_min <= it['cy'] <= y_max and x_min <= it['cx'] <= x_max and it['cx'] >= row_label_right):
                continue
            if group_upper_y is not None and not (it['cy'] < group_upper_y):
                continue
            raw_text = str(it.get('text') or '').strip()
            if not raw_text:
                continue
            if ret_type == 'string':
                if fmt_pat and not fmt_pat.search(raw_text):
                    continue
                val_text = raw_text
            else:
                val_text = _first_numeric(raw_text)
                if not val_text:
                    continue
            dx = abs(it['cx'] - ix)
            dy = abs(it['cy'] - iy)
            fmt_penalty = 0
            if fmt_pat and val_text is not None and not fmt_pat.search(val_text):
                fmt_penalty = 1
            candidates.append(((fmt_penalty, dx + dy), it, val_text, raw_text))

        if candidates:
            candidates.sort(key=lambda t: t[0])
            best_it = candidates[0][1]
            best_val = candidates[0][2]
            best_raw = candidates[0][3]
            header_text = str(hdr.get('text', '') or '')
            row_text_selected = str(row_it.get('text', '') or '')
            method_label = "easyocr:xy(dpi={})".format(dpi)
            confidence_val = float(best_it.get('conf', 0.0))
            context_snippet = "row='{}' col='{}' value='{}'".format(row_text_selected, header_text, best_raw)

            if ret_type == 'string':
                if doc:
                    try:
                        doc.close()
                    except Exception:
                        pass
                return MatchResult(
                    pdf_file=pdf_path.name,
                    serial_number=serial_number,
                    term=spec.term,
                    page=p,
                    number=best_val,
                    units=None,
                    context=context_snippet,
                    method=method_label,
                    found=True,
                    confidence=confidence_val,
                    row_label=row_text_selected,
                    column_label=header_text,
                    text_source='ocr',
                )

            numeric_candidate = best_val
            units_value = extract_units(numeric_candidate)
            numeric_clean = numeric_only(numeric_candidate)
            numeric_value = None
            if numeric_clean is not None:
                try:
                    numeric_value = float(numeric_clean.replace(',', ''))
                except Exception:
                    numeric_value = None

            range_violation = False
            if numeric_value is not None and (spec.range_min is not None or spec.range_max is not None):
                if spec.range_min is not None and numeric_value < spec.range_min:
                    range_violation = True
                if spec.range_max is not None and numeric_value > spec.range_max:
                    range_violation = True

            number_out = numeric_candidate.strip()
            if range_violation and not number_out.rstrip().endswith('(range violation)'):
                number_out = f"{number_out} (range violation)"

            if doc:
                try:
                    doc.close()
                except Exception:
                    pass
            return MatchResult(
                pdf_file=pdf_path.name,
                serial_number=serial_number,
                term=spec.term,
                page=p,
                number=number_out,
                units=units_value,
                context=context_snippet,
                method=method_label,
                found=True,
                confidence=confidence_val,
                row_label=row_text_selected,
                column_label=header_text,
                text_source='ocr',
            )

    if doc:
        try:
            doc.close()
        except Exception:
            pass
    return None
def ocr_pages_with_paddle(pdf_path: Path, pages: Sequence[int]) -> Tuple[Dict[int, str], str]:
    """OCR selected pages using PyMuPDF render + PaddleOCR (pure-Python path).

    Writes each page image to a temporary PNG and runs PaddleOCR on it.
    """
    out: Dict[int, str] = {}
    if not (_HAVE_PYMUPDF and _HAVE_PADDLE_OCR):
        return out, "ocr_paddle:N/A"
    try:
        # Avoid deprecated/unsupported args like show_log in newer releases
        ocr = PaddleOCR(use_angle_cls=True, lang='en')  # type: ignore
    except Exception as e:
        return out, f"ocr_paddle:init_error:{type(e).__name__}"
    try:
        doc = fitz.open(str(pdf_path))  # type: ignore[name-defined]
    except Exception as e:
        return out, f"ocr_paddle:open_error:{e}"
    try:
        dpi = int(os.environ.get('OCR_DPI', '400'))
    except Exception:
        dpi = 400
    dpi = max(200, min(800, dpi))
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)  # type: ignore[name-defined]
    tmp_dir = Path(tempfile.mkdtemp(prefix="paddle_ocr_"))
    try:
        for p in pages:
            if 1 <= p <= doc.page_count:
                try:
                    page = doc.load_page(p - 1)
                    pix = page.get_pixmap(matrix=mat)
                    img_path = tmp_dir / f"page_{p}.png"
                    pix.save(str(img_path))
                except Exception:
                    out[p] = ""
                    continue
                try:
                    result = ocr.ocr(str(img_path), cls=True)  # type: ignore[attr-defined]
                    lines: list[str] = []
                    for block in result or []:
                        for item in block or []:
                            try:
                                txt = item[1][0]
                                if isinstance(txt, str):
                                    lines.append(txt)
                            except Exception:
                                pass
                    out[p] = "\n".join(lines)
                except Exception:
                    out[p] = ""
    finally:
        try:
            shutil.rmtree(str(tmp_dir), ignore_errors=True)
        except Exception:
            pass
        try:
            doc.close()
        except Exception:
            pass
    return out, "ocr_paddle"


def _run_ocrmypdf_to_temp(pdf_path: Path) -> Tuple[Optional[Path], str, Optional[Path]]:
    """Run OCRmyPDF to create a temporary searchable PDF for the whole document.

    Returns (ocr_pdf_path, label, tmp_dir). Caller may remove tmp_dir unless KEEP is set.
    """
    lang = os.environ.get('OCRMYPDF_LANG', 'eng')
    try:
        optimize = int(os.environ.get('OCRMYPDF_OPTIMIZE', '1'))
    except Exception:
        optimize = 1
    keep = os.environ.get('OCRMYPDF_KEEP', '').strip().lower() in ('1', 'true', 'yes', 'keep')
    force = os.environ.get('OCRMYPDF_FORCE', '').strip().lower() in ('1','true','yes','force')

    tmp_dir = Path(tempfile.mkdtemp(prefix="ocrmypdf_"))
    out_path = tmp_dir / (pdf_path.stem + ".ocr.pdf")

    # Prefer Python API
    try:
        try:
            _ocr_fn = _ocrmypdf.ocr  # type: ignore[attr-defined]
        except Exception:
            from ocrmypdf import api as _ocr_api  # type: ignore
            _ocr_fn = _ocr_api.ocr
        _ocr_fn(
            str(pdf_path),
            str(out_path),
            language=lang,
            force_ocr=force,
            rotate_pages=True,
            deskew=True,
            optimize=optimize,
            progress_bar=False,
        )
        return out_path, f"ocrmypdf(opt={optimize})", (tmp_dir if keep else tmp_dir)
    except Exception as e:
        # Try CLI if API failed or is unavailable
        try:
            bin_path = os.environ.get('OCRMYPDF_BIN') or shutil.which('ocrmypdf') or 'ocrmypdf'
            args = [bin_path, '-l', lang]
            if force:
                args.append('--force-ocr')
            args += ['--rotate-pages', '--deskew', '--optimize', str(optimize), str(pdf_path), str(out_path)]
            proc = subprocess.run(args, capture_output=True, text=True)
            if proc.returncode == 0 and out_path.exists():
                return out_path, f"ocrmypdf(opt={optimize})", (tmp_dir if keep else tmp_dir)
            else:
                return None, f"ocrmypdf:cli_error:{proc.returncode}", (tmp_dir if keep else tmp_dir)
        except Exception as e2:
            return None, f"ocrmypdf:error:{e2}", (tmp_dir if keep else tmp_dir)


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
    s = s.replace("Ã¢â‚¬â€œ", "-").replace("Ã¢â‚¬â€", "-")
    s = s.replace("|", " ")
    s = re.sub(r"[ \t\f\r]+", " ", s)
    return s


def extract_pages_text(pdf_path: Path, pages: Sequence[int], do_ocr_fallback: bool = True, ocr_mode: Optional[str] = None) -> Tuple[Dict[int, str], str]:
    """
    Try multiple extraction methods in a fixed order and fill in what we can:
      1) PyMuPDF
      2) pdfminer.six (for empty pages)
      3) pypdf/PyPDF2 (for remaining empties)
      4) OCR (as a last resort if Tesseract is available)
    Return a consolidated {page: text} mapping and a pipeline summary string.
    """
    # Honor OCR mode: fallback (default), ocr_only, no_ocr
    mode = (ocr_mode or _get_ocr_mode())
    tried = []
    if mode == 'ocr_only':
        pt4, m4 = ocr_pages_with_easyocr(pdf_path, pages)
        # Normalize text
        for _p in list(pt4.keys()):
            pt4[_p] = _normalize_text_for_search(pt4.get(_p, ""))
        return pt4, m4
    # OCR renderer is selected internally; external knob removed
    source_pdf = pdf_path

    # Attempt #1: PyMuPDF
    page_text, m = extract_pages_text_pymupdf(source_pdf, pages)
    tried.append(m)

    # Identify which pages are still empty after the first extractor
    empty_pages = [p for p in pages if page_text.get(p, "").strip() == ""]
    # Optional override: force OCR regardless of extracted text
    _force_ocr = (os.environ.get('FORCE_OCR', '') or '').strip().lower() in ('1','true','yes','force','always')
    if _force_ocr:
        empty_pages = list(pages)

    # Attempt #2: pdfminer on empty pages
    if empty_pages:
        pt2, m2 = extract_pages_text_pdfminer(source_pdf, empty_pages)
        tried.append(m2)
        for p in empty_pages:
            if (pt2.get(p) or "").strip():
                page_text[p] = pt2[p]
        empty_pages = [p for p in pages if page_text.get(p, "").strip() == ""]
        if _force_ocr:
            empty_pages = list(pages)

    # Attempt #3: pypdf/PyPDF2 on remaining pages
    if empty_pages:
        pt3, m3 = extract_pages_text_pypdf(source_pdf, empty_pages)
        tried.append(m3)
        for p in empty_pages:
            if (pt3.get(p) or "").strip():
                page_text[p] = pt3[p]
        empty_pages = [p for p in pages if page_text.get(p, "").strip() == ""]
        if _force_ocr:
            empty_pages = list(pages)

    # Attempt #4: EasyOCR fallback for remaining empty pages (or all if forced)
    # Only performs OCR if EasyOCR is available and do_ocr_fallback=True.
    if (mode != 'no_ocr') and do_ocr_fallback and empty_pages:
        pt4, m4 = ocr_pages_with_easyocr(source_pdf, empty_pages)
        tried.append(m4)
        for p in empty_pages:
            if (pt4.get(p) or "").strip():
                page_text[p] = pt4[p]
        # No further fallback stages; recompute empties for completeness but proceed to normalize
        empty_pages = [p for p in pages if page_text.get(p, "").strip() == ""]

    # Normalize text per page to make downstream term matching more robust
    for _p in list(page_text.keys()):
        page_text[_p] = _normalize_text_for_search(page_text.get(_p, ""))

    pipeline = " > ".join(tried)
    return page_text, pipeline




def find_closest_number_in_text(text: str, term: str, window_chars: int = 160, case_sensitive: bool = False,
                                units_hint: Optional[List[str]] = None,
                                range_filter: Optional[Tuple[Optional[float], Optional[float]]] = None,
                                accept_dates: bool = True) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    '''
    Prefer numbers on the same line to the right of the term, then left,
    then next line, previous line, else fall back to closest in a window.
    Returns (number_string or None, context_snippet or None).
    '''
    if not text:
        return None, None, 'No text available'

    src = text
    hay = src if case_sensitive else src.lower()
    needle = term if case_sensitive else term.lower()

    def _simplify(t: str) -> str:
        t = t.lower()
        t = re.sub(r"\b(nominal|minimum|min|maximum|max|range|typical|average|avg|target|req(?:uirement)?)\b", " ", t)
        t = re.sub(r"\s+", " ", t).strip()
        return t

    positions: List[int] = []
    start_idx = 0
    while True:
        idx = hay.find(needle, start_idx)
        if idx == -1:
            break
        positions.append(idx)
        start_idx = idx + max(1, len(needle))

    if not positions:
        alt = _simplify(needle)
        if alt and alt != needle:
            start_idx = 0
            while True:
                idx = hay.find(alt, start_idx)
                if idx == -1:
                    break
                positions.append(idx)
                start_idx = idx + max(1, len(alt))

    if not positions:
        try:
            threshold = 0.75
            src_lines = src.splitlines()
            offset = 0
            for line in src_lines:
                hay_line = line if case_sensitive else line.lower()
                if needle and (needle[0] not in hay_line):
                    pass
                for token in re.split(r"[^A-Za-z0-9]+", hay_line):
                    if not token:
                        continue
                    if len(token) >= max(4, len(needle) - 2):
                        if difflib.SequenceMatcher(None, token, needle).ratio() >= threshold:
                            pos_local = hay_line.find(token)
                            if pos_local >= 0:
                                positions.append(offset + pos_local)
                                break
                if positions:
                    break
                offset += len(line) + 1
        except Exception:
            pass

    if not positions:
        return None, None, 'Term not located in text'

    nums = [(m.group(0), m.start(), m.end()) for m in NUMBER_REGEX.finditer(src)]
    if accept_dates:
        nums += [(m.group(0), m.start(), m.end()) for m in DATE_REGEX.finditer(src)]

    def numbers_in(a: int, b: int) -> List[Tuple[str, int, int]]:
        return [(n, i, j) for (n, i, j) in nums if i >= a and j <= b]

    def snippet(a: int, b: int) -> str:
        return src[max(0, a - 60): min(len(src), b + 60)].replace(chr(10), " ")

    def range_violation(nstr: str) -> bool:
        if not range_filter:
            return False
        lo, hi = range_filter
        try:
            raw = numeric_only(nstr)
            val = float(raw) if raw is not None else None
        except Exception:
            val = None
        if val is None:
            return False
        if lo is not None and val < lo:
            return True
        if hi is not None and val > hi:
            return True
        return False

    def units_match(nstr: str) -> bool:
        if not units_hint:
            return True
        u = extract_units(nstr)
        if not u:
            return False
        return any(u.lower() == h.lower() for h in units_hint)

    def apply_range_note(nstr: str, ok: bool) -> str:
        if ok:
            return nstr
        trimmed = nstr.rstrip()
        note = ' (range violation)'
        if trimmed.endswith(note):
            return trimmed
        return f"{trimmed}{note}"

    def prioritize(candidates: List[Tuple[str, int, int]]) -> List[Tuple[str, int, int, bool]]:
        annotated: List[Tuple[str, int, int, bool]] = []
        for n, i, j in candidates:
            annotated.append((n, i, j, not range_violation(n)))
        if not annotated:
            return []
        pref = [item for item in annotated if item[3] and units_match(item[0])]
        if pref:
            return pref
        pref = [item for item in annotated if item[3]]
        if pref:
            return pref
        pref = [item for item in annotated if units_match(item[0])]
        if pref:
            return pref
        return annotated

    best_num = None
    best_ctx = None
    best_dist = 10 ** 9
    failure_reason: Optional[str] = None

    for pos in positions:
        lb = src.rfind(chr(10), 0, pos) + 1
        rb = src.find(chr(10), pos)
        if rb == -1:
            rb = len(src)
        line_nums = numbers_in(lb, rb)
        prioritized_line = prioritize(line_nums)
        right_side = [(n, i, j, ok) for (n, i, j, ok) in prioritized_line if i >= pos]
        if right_side:
            n, i, j, range_ok = min(right_side, key=lambda t: t[1] - pos)
            return apply_range_note(n, range_ok), snippet(i, j), None

        left_side = [(n, i, j, ok) for (n, i, j, ok) in prioritized_line if j <= pos]
        if left_side:
            n, i, j, range_ok = max(left_side, key=lambda t: pos - t[2])
            return apply_range_note(n, range_ok), snippet(i, j), None

        nlb = rb + 1
        nrb = src.find(chr(10), nlb)
        if nrb == -1:
            nrb = len(src)
        next_nums = numbers_in(nlb, nrb)
        prioritized_next = prioritize(next_nums)
        if prioritized_next:
            n, i, j, range_ok = prioritized_next[0]
            return apply_range_note(n, range_ok), snippet(i, j), None

        plb = src.rfind(chr(10), 0, lb - 1) + 1
        prb = lb - 1 if lb > 0 else 0
        prev_nums = numbers_in(plb, prb)
        prioritized_prev = prioritize(prev_nums)
        if prioritized_prev:
            n, i, j, range_ok = prioritized_prev[-1]
            return apply_range_note(n, range_ok), snippet(i, j), None

        left = max(0, pos - window_chars)
        right = min(len(src), pos + len(term) + window_chars)
        cand = numbers_in(left, right)
        prioritized_window = prioritize(cand)
        for n, i, j, range_ok in prioritized_window:
            d = min(abs(i - pos), abs(j - pos))
            if d < best_dist:
                best_dist = d
                best_num = apply_range_note(n, range_ok)
                best_ctx = snippet(i, j)

    if best_num is not None:
        return best_num, best_ctx, None
    return None, None, failure_reason

def _compile_value_regex(fmt: str) -> Optional[re.Pattern]:
    if not fmt:
        return None
    s = fmt.strip()
    try:
        if len(s) >= 2 and s.startswith('/') and s.endswith('/'):
            return re.compile(s[1:-1], flags=re.IGNORECASE)
    except Exception:
        pass
    mask = []
    for ch in s:
        if ch in ('x','X'):
            mask.append('[A-Za-z0-9]')
        elif ch in ('d','D'):
            mask.append(r'\d')
        elif ch == '*':
            mask.append('.+')
        elif ch == '?':
            mask.append('.')
        else:
            mask.append(re.escape(ch))
    try:
        return re.compile(''.join(mask), flags=re.IGNORECASE)
    except Exception:
        return None





def scan_pdf_for_term_nearest(pdf_path: Path, serial_number: str, spec: TermSpec, window_chars: int, case_sensitive: bool) -> MatchResult:
    anchor_text = (spec.anchor or spec.term or '').strip()
    ret_kind = (spec.return_type or 'number').strip().lower()
    fmt_pat = _compile_value_regex(spec.value_format or '') if spec.value_format else None
    anchor_tokens = [tok for tok in re.split(r"\s+", anchor_text) if tok]
    anchor_tokens_lower = [tok.lower() for tok in anchor_tokens]
    anchor_tokens_norm = [_normalize_anchor_token(tok) for tok in anchor_tokens]
    anchor_lower_set = set(anchor_tokens_lower)
    anchor_norm_set = {tok for tok in anchor_tokens_norm if tok}

    def _is_anchor_duplicate(token: str) -> bool:
        if not token:
            return False
        token_norm = _normalize_anchor_token(token)
        if case_sensitive:
            if token in anchor_tokens:
                return True
            if token_norm and token_norm in anchor_norm_set:
                return True
            return False
        if token.lower() in anchor_lower_set:
            return True
        if token_norm and token_norm in anchor_norm_set:
            return True
        return False

    def _match_anchor(token_texts: List[str], token_norms: List[str]) -> Optional[Tuple[int, int]]:
        if not anchor_tokens:
            return None
        comparables = token_texts if case_sensitive else [t.lower() for t in token_texts]
        span = len(anchor_tokens)
        for idx in range(0, len(token_texts) - span + 1):
            ok = True
            for j in range(span):
                expected = anchor_tokens[j] if case_sensitive else anchor_tokens_lower[j]
                current = comparables[idx + j]
                if current != expected:
                    expected_norm = anchor_tokens_norm[j]
                    if not expected_norm or token_norms[idx + j] != expected_norm:
                        ok = False
                        break
            if ok:
                return idx, idx + span - 1
        return None

    def _annotate_number(raw_value: str) -> Tuple[str, Optional[str]]:
        number_out = raw_value.strip()
        numeric_clean = numeric_only(raw_value)
        numeric_val: Optional[float] = None
        if numeric_clean is not None:
            try:
                numeric_val = float(numeric_clean.replace(',', ''))
            except Exception:
                numeric_val = None
        range_violation = False
        if numeric_val is not None and (spec.range_min is not None or spec.range_max is not None):
            if spec.range_min is not None and numeric_val < spec.range_min:
                range_violation = True
            if spec.range_max is not None and numeric_val > spec.range_max:
                range_violation = True
        if range_violation and not number_out.rstrip().endswith('(range violation)'):
            number_out = f"{number_out} (range violation)"
        return number_out, extract_units(raw_value)

    def _search_with_pymupdf() -> Tuple[Optional[MatchResult], bool, Optional[str]]:
        if not _HAVE_PYMUPDF:
            return None, False, None
        if not anchor_tokens:
            return None, False, "No anchor text specified for nearest-line mode"
        try:
            doc = fitz.open(str(pdf_path))
        except Exception:
            return None, False, None
        anchor_seen = False
        failure_local: Optional[str] = None
        try:
            total_pages = getattr(doc, 'page_count', 0)
            target_pages = spec.pages if spec.pages else list(range(1, total_pages + 1))
            for p in target_pages:
                if p < 1 or (total_pages and p > total_pages):
                    continue
                try:
                    page = doc.load_page(p - 1)
                except Exception:
                    continue
                words = page.get_text("words") or []
                if not words:
                    continue
                lines_map: Dict[int, List[List[float]]] = {}
                for w in words:
                    line_id = w[6] if len(w) >= 7 else round(float(w[1]))
                    lines_map.setdefault(line_id, []).append(w)
                group_after_y: Optional[float] = None
                if spec.group_after:
                    pattern = spec.group_after if case_sensitive else spec.group_after.lower()
                    for ws in lines_map.values():
                        ordered = sorted(ws, key=lambda k: k[0])
                        line_text = " ".join(str(x[4]) for x in ordered)
                        cmp_text = line_text if case_sensitive else line_text.lower()
                        if pattern in cmp_text:
                            cy_vals = [(float(w[1]) + float(w[3])) / 2.0 for w in ordered]
                            if cy_vals:
                                candidate = max(cy_vals)
                                if group_after_y is None or candidate > group_after_y:
                                    group_after_y = candidate
                group_before_y: Optional[float] = None
                if spec.group_before:
                    pattern = spec.group_before if case_sensitive else spec.group_before.lower()
                    for ws in lines_map.values():
                        ordered = sorted(ws, key=lambda k: k[0])
                        line_text = " ".join(str(x[4]) for x in ordered)
                        cmp_text = line_text if case_sensitive else line_text.lower()
                        if pattern in cmp_text:
                            cy_vals = [(float(w[1]) + float(w[3])) / 2.0 for w in ordered]
                            if cy_vals:
                                candidate = min(cy_vals)
                                if group_before_y is None or candidate < group_before_y:
                                    group_before_y = candidate
                sorted_lines = sorted(
                    (
                        (line_id, sorted(ws, key=lambda k: k[0]))
                        for line_id, ws in lines_map.items()
                        if ws
                    ),
                    key=lambda item: min((float(w[1]) + float(w[3])) / 2.0 for w in item[1])
                )
                for _, row_words in sorted_lines:
                    token_texts = [str(w[4]) if len(w) > 4 else '' for w in row_words]
                    token_norms = [_normalize_anchor_token(t) for t in token_texts]
                    cy_vals = [(float(w[1]) + float(w[3])) / 2.0 for w in row_words if len(w) >= 4]
                    if not cy_vals:
                        continue
                    row_cy = sum(cy_vals) / len(cy_vals)
                    if group_after_y is not None and row_cy <= group_after_y:
                        continue
                    if group_before_y is not None and row_cy >= group_before_y:
                        continue
                    match_span = _match_anchor(token_texts, token_norms)
                    if not match_span:
                        continue
                    anchor_seen = True
                    start_idx, end_idx = match_span
                    tail_candidates: List[Tuple[str, List[float]]] = []
                    for w in row_words[end_idx + 1:]:
                        if len(w) < 5:
                            continue
                        token = str(w[4]).strip()
                        if not token:
                            continue
                        if _is_anchor_duplicate(token):
                            continue
                        tail_candidates.append((token, w))
                    if not tail_candidates:
                        failure_local = "No value to the right of anchor on same line"
                        continue
                    line_text = " ".join(token_texts).strip()
                    if ret_kind == 'string':
                        value_text = None
                        if fmt_pat:
                            for token, _ in tail_candidates:
                                m = fmt_pat.search(token)
                                if m:
                                    value_text = m.group(0)
                                    break
                        if value_text is None:
                            for token, _ in tail_candidates:
                                token_clean = token.strip()
                                if token_clean:
                                    value_text = token_clean
                                    break
                        if value_text is None and tail_candidates:
                            value_text = spec.column or spec.term or tail_candidates[0][0]
                        if value_text:
                            return MatchResult(
                                pdf_file=pdf_path.name,
                                serial_number=serial_number,
                                term=spec.term,
                                page=p,
                                number=value_text,
                                units=None,
                                context=line_text[:200],
                                method="pymupdf:nearest-line",
                                found=True,
                                confidence=None,
                                row_label=None,
                                column_label=None,
                                text_source="pdf",
                            ), True, None
                        failure_local = "No value to the right of anchor on same line"
                        continue
                    number_data = None
                    for token, _word in tail_candidates:
                        numeric_candidate = _first_numeric(token)
                        if not numeric_candidate:
                            continue
                        number_data = _annotate_number(numeric_candidate)
                        break
                    if number_data:
                        number_out, units_value = number_data
                        return MatchResult(
                            pdf_file=pdf_path.name,
                            serial_number=serial_number,
                            term=spec.term,
                            page=p,
                            number=number_out,
                            units=units_value,
                            context=line_text[:200],
                            method="pymupdf:nearest-line",
                            found=True,
                            confidence=None,
                            row_label=None,
                            column_label=None,
                            text_source="pdf",
                        ), True, None
                    failure_local = "No value to the right of anchor on same line"
        finally:
            try:
                doc.close()
            except Exception:
                pass
        return None, anchor_seen, failure_local

    def _search_with_pdf_text() -> Tuple[Optional[MatchResult], bool, Optional[str]]:
        if not anchor_tokens:
            return None, False, "No anchor text specified for nearest-line mode"
        try:
            page_count = get_pdf_page_count(pdf_path)
        except Exception:
            page_count = 0
        target_pages = spec.pages if spec.pages else (list(range(1, page_count + 1)) if page_count else [1])
        try:
            page_text_map, pipeline = extract_pages_text(pdf_path, target_pages, do_ocr_fallback=False)
        except Exception:
            page_text_map, pipeline = {}, "text"
        else:
            pipeline = pipeline or "text"
        after_found = not bool(spec.group_after)
        before_triggered = False
        anchor_seen = False
        failure_local: Optional[str] = None
        for p in target_pages:
            if before_triggered:
                break
            raw_text = page_text_map.get(p, '')
            if not raw_text:
                continue
            slice_after = spec.group_after if not after_found else None
            slice_before = spec.group_before if not before_triggered else None
            text_section, after_hit, before_hit = _slice_text_by_groups(raw_text, slice_after, slice_before, case_sensitive)
            if spec.group_after and not after_found:
                if not after_hit:
                    continue
                after_found = True
            else:
                after_found = True
            if before_hit:
                before_triggered = True
            if not text_section.strip():
                if before_hit:
                    break
                continue
            for line in text_section.splitlines():
                line = line.strip()
                if not line:
                    continue
                tokens = re.findall(r"\S+", line)
                if not tokens:
                    continue
                token_norms = [_normalize_anchor_token(t) for t in tokens]
                match_span = _match_anchor(tokens, token_norms)
                if not match_span:
                    continue
                anchor_seen = True
                start_idx, end_idx = match_span
                tail_tokens = []
                for tok in tokens[end_idx + 1:]:
                    if _is_anchor_duplicate(tok):
                        continue
                    tail_tokens.append(tok)
                if not tail_tokens:
                    failure_local = "No value to the right of anchor on same line"
                    continue
                if ret_kind == 'string':
                    value_text = None
                    if fmt_pat:
                        for tok in tail_tokens:
                            m = fmt_pat.search(tok)
                            if m:
                                value_text = m.group(0)
                                break
                        if value_text is None:
                            for tok in next_line_tokens:
                                m = fmt_pat.search(tok)
                                if m:
                                    value_text = m.group(0)
                                    break
                    if value_text is None:
                        for tok in tail_tokens:
                            tok_clean = tok.strip()
                            if tok_clean:
                                value_text = tok_clean
                                break
                    if value_text is None:
                        for tok in next_line_tokens:
                            tok_clean = tok.strip()
                            if tok_clean:
                                value_text = tok_clean
                                break
                    if value_text is None:
                        fallback_tokens = tail_tokens or next_line_tokens
                        if fallback_tokens:
                            value_text = spec.column or spec.term or fallback_tokens[0]
                    if value_text:
                        return MatchResult(
                            pdf_file=pdf_path.name,
                            serial_number=serial_number,
                            term=spec.term,
                            page=p,
                            number=value_text,
                            units=None,
                            context=line[:200],
                            method="text:nearest-line",
                            found=True,
                            confidence=None,
                            row_label=None,
                            column_label=None,
                            text_source="pdf",
                        ), True, None
                    failure_local = "No value to the right of anchor on same line"
                    continue
                number_data = None
                for tok in tail_tokens:
                    numeric_candidate = _first_numeric(tok)
                    if numeric_candidate:
                        number_data = _annotate_number(numeric_candidate)
                        break
                if number_data:
                    number_out, units_value = number_data
                    return MatchResult(
                        pdf_file=pdf_path.name,
                        serial_number=serial_number,
                        term=spec.term,
                        page=p,
                        number=number_out,
                        units=units_value,
                        context=line[:200],
                        method="text:nearest-line",
                        found=True,
                        confidence=None,
                        row_label=None,
                        column_label=None,
                        text_source="pdf",
                    ), True, None
                failure_local = "No value to the right of anchor on same line"
        return None, anchor_seen, failure_local

    def _group_easyocr_rows(items: List[Dict[str, float]]) -> List[Dict[str, object]]:
        rows: List[Dict[str, object]] = []
        for it in sorted(items, key=lambda d: (float(d.get('cy', 0.0)), float(d.get('cx', 0.0)))):
            cy = float(it.get('cy', 0.0))
            if not rows or abs(cy - float(rows[-1]['cy'])) > 8.0:
                rows.append({'cy': cy, 'items': [it]})
            else:
                rows[-1]['items'].append(it)
        for row in rows:
            row_items = row['items']  # type: ignore[assignment]
            row_items.sort(key=lambda d: float(d.get('cx', 0.0)))
            row['text'] = " ".join(str(d.get('text', '') or '') for d in row_items).strip()
        return rows

    def _search_with_easyocr() -> Tuple[Optional[MatchResult], bool, Optional[str]]:
        if not (_HAVE_EASYOCR and _HAVE_PYMUPDF):
            return None, False, None
        if not anchor_tokens:
            return None, False, "No anchor text specified for nearest-line mode"
        langs_raw = (os.environ.get('EASYOCR_LANGS') or os.environ.get('OCR_LANGS') or 'en')
        langs = [s.strip() for s in re.split(r'[;,]', langs_raw) if s.strip()]
        try:
            dpi_base = int(os.environ.get('OCR_DPI', '700'))
        except Exception:
            dpi_base = 700
        dpi_candidates = [dpi_base]
        if dpi_base > 700:
            dpi_candidates.append(700)
        try:
            doc = fitz.open(str(pdf_path))
            total_pages = getattr(doc, 'page_count', 0)
        except Exception:
            doc = None
            total_pages = 0
        if doc:
            try:
                doc.close()
            except Exception:
                pass
        target_pages = spec.pages if spec.pages else list(range(1, total_pages + 1))
        if not target_pages:
            return None, False, None
        anchor_seen = False
        failure_local: Optional[str] = None
        for dpi in dpi_candidates:
            boxes_by_page = _easyocr_boxes_for_pages(pdf_path, target_pages, dpi=dpi, langs=langs)
            for p in target_pages:
                items = boxes_by_page.get(p, [])
                if not items:
                    continue
                rows = _group_easyocr_rows(items)
                group_after_y: Optional[float] = None
                if spec.group_after:
                    pattern = spec.group_after if case_sensitive else spec.group_after.lower()
                    for row in rows:
                        text_row = row['text']  # type: ignore[index]
                        cmp_text = text_row if case_sensitive else text_row.lower()
                        if pattern in cmp_text:
                            if group_after_y is None or float(row['cy']) > group_after_y:
                                group_after_y = float(row['cy'])
                group_before_y: Optional[float] = None
                if spec.group_before:
                    pattern = spec.group_before if case_sensitive else spec.group_before.lower()
                    for row in rows:
                        text_row = row['text']  # type: ignore[index]
                        cmp_text = text_row if case_sensitive else text_row.lower()
                        if pattern in cmp_text:
                            if group_before_y is None or float(row['cy']) < group_before_y:
                                group_before_y = float(row['cy'])
                for row in rows:
                    row_cy = float(row['cy'])  # type: ignore[index]
                    if group_after_y is not None and row_cy <= group_after_y:
                        continue
                    if group_before_y is not None and row_cy >= group_before_y:
                        continue
                    row_items = row['items']  # type: ignore[index]
                    token_texts = [str(it.get('text', '') or '') for it in row_items]
                    token_norms = [_normalize_anchor_token(t) for t in token_texts]
                    match_span = _match_anchor(token_texts, token_norms)
                    if not match_span:
                        continue
                    anchor_seen = True
                    start_idx, end_idx = match_span
                    tail_candidates: List[Tuple[str, Dict[str, float]]] = []
                    for idx in range(end_idx + 1, len(row_items)):
                        token = str(row_items[idx].get('text', '') or '').strip()
                        if not token:
                            continue
                        if _is_anchor_duplicate(token):
                            continue
                        tail_candidates.append((token, row_items[idx]))
                    if not tail_candidates:
                        failure_local = "No value to the right of anchor on same line"
                        continue
                    row_text = row['text']  # type: ignore[index]
                    context_text = row_text[:200] if isinstance(row_text, str) else ''
                    if ret_kind == 'string':
                        value_text = None
                        if fmt_pat:
                            for token, _ in tail_candidates:
                                m = fmt_pat.search(token)
                                if m:
                                    value_text = m.group(0)
                                    break
                        if value_text is None:
                            for token, _ in tail_candidates:
                                num = _first_numeric(token)
                                if num:
                                    value_text = num
                                    break
                        if value_text is None and tail_candidates:
                            value_text = tail_candidates[0][0]
                        if value_text:
                            try:
                                conf_val = float(tail_candidates[0][1].get('conf', 0.0) or 0.0)
                            except Exception:
                                conf_val = 0.0
                            return MatchResult(
                                pdf_file=pdf_path.name,
                                serial_number=serial_number,
                                term=spec.term,
                                page=p,
                                number=value_text,
                                units=None,
                                context=context_text,
                                method=f"easyocr:nearest-line(dpi={dpi})",
                                found=True,
                                confidence=conf_val,
                                row_label=None,
                                column_label=None,
                                text_source="ocr",
                            ), True, None
                        failure_local = "No value to the right of anchor on same line"
                        continue
                    number_data = None
                    idx_used: Optional[int] = None
                    for token, item in tail_candidates:
                        numeric_candidate = _first_numeric(token)
                        if not numeric_candidate:
                            continue
                        number_data = _annotate_number(numeric_candidate)
                        idx_used = item
                        break
                    if number_data and idx_used is not None:
                        number_out, units_value = number_data
                        try:
                            conf_val = float(idx_used.get('conf', 0.0) or 0.0)
                        except Exception:
                            conf_val = 0.0
                        return MatchResult(
                            pdf_file=pdf_path.name,
                            serial_number=serial_number,
                            term=spec.term,
                            page=p,
                            number=number_out,
                            units=units_value,
                            context=context_text,
                            method=f"easyocr:nearest-line(dpi={dpi})",
                            found=True,
                            confidence=conf_val,
                            row_label=None,
                            column_label=None,
                            text_source="ocr",
                        ), True, None
                    failure_local = "No value to the right of anchor on same line"
        return None, anchor_seen, failure_local

    pdf_result, pdf_anchor_seen, pdf_failure = _search_with_pymupdf()
    if pdf_result:
        return pdf_result

    text_result, text_anchor_seen, text_failure = _search_with_pdf_text()
    if text_result:
        return text_result

    ocr_result, ocr_anchor_seen, ocr_failure = _search_with_easyocr()
    if ocr_result:
        return ocr_result

    failure_reason = pdf_failure or text_failure or ocr_failure
    if failure_reason is None:
        if anchor_tokens:
            failure_reason = "Anchor not located for nearest-line mode"
        else:
            failure_reason = "No anchor text specified for nearest-line mode"
    if ocr_anchor_seen:
        method_label = "easyocr:nearest-line"
    elif text_anchor_seen:
        method_label = "text:nearest-line"
    elif pdf_anchor_seen:
        method_label = "pymupdf:nearest-line"
    else:
        method_label = "nearest-line"
    return MatchResult(
        pdf_file=pdf_path.name,
        serial_number=serial_number,
        term=spec.term,
        page=None,
        number=None,
        units=None,
        context="",
        method=method_label,
        found=False,
        confidence=None,
        row_label=None,
        column_label=None,
        text_source=None,
        error_reason=failure_reason,
    )


def get_serial_number_from_filename(pdf_path: Path) -> str:
    """Extract the serial number from the PDF filename using SN_REGEX.
    Returns a string like "SN 1234". If no match is found, returns a fallback based on the basename."""
    name = pdf_path.stem
    m = SN_REGEX.search(name)
    if m:
        return f"SN {m.group(1)}"
    m = SN_REGEX.search(pdf_path.name)
    if m:
        return f"SN {m.group(1)}"
    return f"SN_{name}"


def scan_pdf_for_term(pdf_path: Path, serial_number: str, term: str, pages: Sequence[int], window_chars: int, case_sensitive: bool,
                      units_hint: Optional[List[str]] = None,
                      range_filter: Optional[Tuple[Optional[float], Optional[float]]] = None) -> MatchResult:
    """
    Scan a single PDF for a single term (restricted to the provided pages).
    - Uses extract_pages_text(...) to build a map of pageÃƒÂ¢Ã¢â‚¬Â Ã¢â‚¬â„¢text and a method pipeline string.
    - Calls find_closest_number_in_text(...) to get the nearest number and context.
    - Returns a MatchResult with page/number/context and pipeline details.
    """
    # Build text for constrained pages (or the whole doc if no pages specified)
    # Prefer pre-extracted cache when available to avoid re-reading per term
    key = _pdf_cache_key(pdf_path)
    if key in _PAGE_TEXT_CACHE:
        full_map, pipeline, _pc = _PAGE_TEXT_CACHE[key]
        if pages:
            page_text_map = {p: (full_map.get(p) or "") for p in pages}
        else:
            page_text_map = dict(full_map)
    else:
        page_text_map, pipeline = extract_pages_text(pdf_path, pages if pages else list(range(1, 10000)), do_ocr_fallback=False)

    chosen_page = None
    chosen_number = None
    chosen_ctx = None
    failure_reason: Optional[str] = None
    failure_reason = "No numeric value located near term"

    # Search pages in ascending order; stop at the first page where a number is found
    for p in sorted(page_text_map.keys()):
        text = page_text_map[p]
        number, ctx, reason = find_closest_number_in_text(text, term, window_chars=window_chars, case_sensitive=case_sensitive,
                                                         units_hint=units_hint, range_filter=range_filter, accept_dates=True)
        if number:
            chosen_page = p
            chosen_number = number
            chosen_ctx = ctx or ""
            break
        failure_reason = reason or failure_reason
        failure_reason = "No numeric value located near term"

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
            found=True,
            confidence=None,
            row_label=None,
            column_label=None,
            text_source="pdf"
        )
    # If not found via text extraction, OCR empty pages for this term and retry
    _mode2 = _get_ocr_mode()
    pages_list = pages if pages else list(sorted(page_text_map.keys()))
    empty_pages = [p for p in pages_list if (page_text_map.get(p, "").strip() == "")]
    if (_mode2 != 'no_ocr') and empty_pages:
        pt4, m4 = ocr_pages_with_easyocr(pdf_path, empty_pages)
        # Update cache and local view
        new_pipe = pipeline if (m4 in (pipeline or "")) else (pipeline + " > " + m4 if pipeline else m4)
        # Update full_map if available, else use page_text_map as backing
        if key in _PAGE_TEXT_CACHE:
            full_map, _, _pc = _PAGE_TEXT_CACHE[key]
        else:
            full_map, _pc = dict(page_text_map), get_pdf_page_count(pdf_path)
        for p in empty_pages:
            txt = (pt4.get(p) or "")
            if txt:
                full_map[p] = _normalize_text_for_search(txt)
        _PAGE_TEXT_CACHE[key] = (full_map, new_pipe, _pc)
        # Rebuild page_text_map for searched pages with new text
        if pages:
            page_text_map = {p: (full_map.get(p) or "") for p in pages}
        else:
            page_text_map = dict(full_map)

        # Retry search across pages
        chosen_page = None
        chosen_number = None
        chosen_ctx = None
        for p in sorted(page_text_map.keys()):
            text = page_text_map[p]
            number, ctx, reason = find_closest_number_in_text(text, term, window_chars=window_chars, case_sensitive=case_sensitive,
                                                             units_hint=units_hint, range_filter=range_filter, accept_dates=True)
            if number:
                chosen_page = p
                chosen_number = number
                chosen_ctx = ctx or ""
                break
            failure_reason = reason or failure_reason
        if chosen_number is None:
            failure_reason = "OCR fallback found no numeric value to the right"
        if chosen_number:
            return MatchResult(
                pdf_file=pdf_path.name,
                serial_number=serial_number,
                term=term,
                page=chosen_page,
                number=chosen_number,
                units=extract_units(chosen_number),
                context=chosen_ctx or "",
                method=new_pipe,
                found=True,
                confidence=None,
                row_label=None,
                column_label=None,
                text_source="ocr"
            )

        # No result after OCR fallback
        pipeline = new_pipe

    return MatchResult(
        pdf_file=pdf_path.name,
        serial_number=serial_number,
        term=term,
        page=None,
        number=None,
        units=None,
        context="",
        method=pipeline,
        found=False,
        confidence=None,
        row_label=None,
        column_label=None,
        text_source=None,
        error_reason=failure_reason
    )


def scan_pdf_for_term_xy(pdf_path: Path, serial_number: str, spec: TermSpec, window_chars: int, case_sensitive: bool) -> MatchResult:
    """Extract table intersection values using OCR geometry only."""
    row_label = spec.line or spec.term
    column_label = spec.column
    mode = _get_ocr_mode()
    if mode == 'no_ocr':
        return MatchResult(
            pdf_file=pdf_path.name,
            serial_number=serial_number,
            term=spec.term,
            page=None,
            number=None,
            units=None,
            context="",
            method="easyocr:xy",
            found=False,
            confidence=None,
            row_label=row_label,
            column_label=column_label,
            text_source=None,
            error_reason="OCR disabled for table(xy) mode"
        )
    if not _HAVE_EASYOCR:
        return MatchResult(
            pdf_file=pdf_path.name,
            serial_number=serial_number,
            term=spec.term,
            page=None,
            number=None,
            units=None,
            context="",
            method="easyocr:xy",
            found=False,
            confidence=None,
            row_label=row_label,
            column_label=column_label,
            text_source=None,
            error_reason="EasyOCR not available for table(xy) mode"
        )

    result = scan_pdf_for_term_xy_easyocr(pdf_path, serial_number, spec, window_chars, case_sensitive)
    if result is not None:
        return result

    return MatchResult(
        pdf_file=pdf_path.name,
        serial_number=serial_number,
        term=spec.term,
        page=None,
        number=None,
        units=None,
        context="",
        method="easyocr:xy",
        found=False,
        confidence=None,
        row_label=row_label,
        column_label=column_label,
        text_source='ocr',
        error_reason="OCR could not locate the table intersection"
    )
def scan_pdf_for_term_line(pdf_path: Path, serial_number: str, spec: TermSpec, window_chars: int, case_sensitive: bool) -> MatchResult:
    """Line-mode extraction: find a line containing an anchor, then pick the Nth field after it.
    Field splitting: auto (groups of 2+ spaces or tabs, else tokens), groups, tokens.
    Return type: string or number.
    """
    # 0) If PyMuPDF is available, try geometry-based grouping on the actual line using word gaps.
    if _HAVE_PYMUPDF:
        try:
            doc = fitz.open(str(pdf_path))
            pages = spec.pages if spec.pages else list(range(1, doc.page_count + 1))
            anchor = (spec.anchor or spec.term or "")
            norm = (lambda t: t) if case_sensitive else (lambda t: t.lower())
            idx = spec.field_index or 1
            after_found = not bool(spec.group_after)
            before_triggered = False
            for p in pages:
                if before_triggered:
                    break
                if p < 1 or p > doc.page_count:
                    continue
                page = doc.load_page(p - 1)
                words = page.get_text("words") or []
                # Group words by line id (w[6])
                lines_map: Dict[int, List[List[float]]] = {}
                for w in words:
                    ln = w[6] if len(w) >= 7 else round(float(w[1]))
                    lines_map.setdefault(ln, []).append(w)

                # Optional grouping anchor: require the anchor line to appear after this marker
                group_anchor_y = None
                if spec.group_after:
                    ga_thresh = 0.55
                    best_ga_score = -1.0
                    for ln_ga, ws_ga in lines_map.items():
                        sorted_ga = sorted(ws_ga, key=lambda k: k[0])
                        ga_line = " ".join(str(x[4]) for x in sorted_ga)
                        sc_ga = _fuzzy_ratio(ga_line, spec.group_after)
                        if sc_ga >= ga_thresh:
                            cy = sum(((float(w[1]) + float(w[3])) / 2.0) for w in ws_ga) / max(1, len(ws_ga))
                            if (
                                group_anchor_y is None
                                or cy > group_anchor_y
                                or (abs((group_anchor_y or 0.0) - cy) <= 0.5 and sc_ga > best_ga_score)
                            ):
                                group_anchor_y = cy
                                best_ga_score = sc_ga
                if group_anchor_y is not None:
                    after_found = True
                if spec.group_after and not after_found:
                    continue

                group_before_y = None
                before_on_page = False
                if spec.group_before:
                    gb_thresh = 0.55
                    best_gb_score = -1.0
                    for ln_gb, ws_gb in lines_map.items():
                        sorted_gb = sorted(ws_gb, key=lambda k: k[0])
                        gb_line = " ".join(str(x[4]) for x in sorted_gb)
                        sc_gb = _fuzzy_ratio(gb_line, spec.group_before)
                        if sc_gb >= gb_thresh:
                            cy = sum(((float(w[1]) + float(w[3])) / 2.0) for w in ws_gb) / max(1, len(ws_gb))
                            if (
                                group_before_y is None
                                or cy < group_before_y
                                or (abs((group_before_y or 0.0) - cy) <= 0.5 and sc_gb > best_gb_score)
                            ):
                                group_before_y = cy
                                best_gb_score = sc_gb
                    if group_before_y is not None:
                        before_on_page = True

                # find candidate lines containing the anchor
                anchor_candidates: List[Tuple[float, float, int, List[List[float]], str]] = []
                anchor_norm = norm(anchor) if anchor else ""
                anchor_thresh = 0.6
                for ln, ws in lines_map.items():
                    ws_sorted = sorted(ws, key=lambda k: k[0])
                    line_str = " ".join(str(x[4]) for x in ws_sorted)
                    if not anchor:
                        continue
                    hay = norm(line_str)
                    score = 1.0 if (anchor_norm and anchor_norm in hay) else _fuzzy_ratio(line_str, anchor)
                    if score >= anchor_thresh:
                        line_cy = sum(((float(w[1]) + float(w[3])) / 2.0) for w in ws) / max(1, len(ws))
                        if group_anchor_y is not None and line_cy <= group_anchor_y:
                            continue
                        if group_before_y is not None and line_cy >= group_before_y:
                            continue
                        anchor_candidates.append((score, line_cy, ln, ws_sorted, line_str))
                if not anchor_candidates:
                    if before_on_page:
                        before_triggered = True
                    continue
                if group_anchor_y is not None:
                    anchor_candidates.sort(key=lambda t: (t[1] - group_anchor_y, -t[0]))
                else:
                    anchor_candidates.sort(key=lambda t: (-t[0], t[1]))
                _, line_cy, target_ln, ws_sorted, line_text = anchor_candidates[0]

                # Build tail words after the anchor occurrence (to avoid counting anchor tokens)
                pos = norm(line_text).find(norm(anchor)) if anchor else 0
                # Filter words whose center is to the right of the anchor occurrence
                # Estimate anchor x by scanning characters left-to-right
                anchor_x = None
                if anchor:
                    accum = 0
                    # crude mapping: distribute line length along word widths
                    total_len = len(line_text)
                    if total_len > 0 and pos >= 0:
                        running = 0
                        for w in ws_sorted:
                            txt = str(w[4])
                            running_end = running + len(txt) + 1  # include a space
                            if running_end >= pos:
                                anchor_x = (float(w[0]) + float(w[2]))/2.0
                                break
                            running = running_end
                # Build fields by geometric gaps (concatenate words until a big gap)
                fields = []
                current = []
                prev_x1 = None
                # dynamic threshold based on median character height/word width
                gap_threshold = 18.0
                try:
                    # Estimate threshold from median word width if available
                    widths = [float(w[2]) - float(w[0]) for w in ws_sorted if (float(w[2]) - float(w[0])) > 0]
                    if widths:
                        widths.sort()
                        med = widths[len(widths)//2]
                        gap_threshold = max(12.0, min(36.0, med * 0.8))
                except Exception:
                    pass
                for w in ws_sorted:
                    cx = (float(w[0]) + float(w[2]))/2.0
                    if anchor_x is not None and cx < anchor_x:
                        continue
                    if prev_x1 is None:
                        current.append(str(w[4]))
                        prev_x1 = float(w[2])
                        continue
                    gap = float(w[0]) - prev_x1
                    if gap > gap_threshold:
                        fields.append(" ".join(current).strip())
                        current = [str(w[4])]
                    else:
                        current.append(str(w[4]))
                    prev_x1 = float(w[2])
                if current:
                    fields.append(" ".join(current).strip())
                # Remove empties and ensure we have enough fields
                fields = [f for f in fields if f]
                if len(fields) >= idx:
                    selected = fields[idx - 1].strip()
                    selected = re.sub(r"\s+", " ", selected).lstrip(":-Ã¢â‚¬â€œÃ¢â‚¬â€ ")
                    if (spec.return_type or "number").lower() == "string":
                        try:
                            doc.close()
                        except Exception:
                            pass
                        return MatchResult(
                            pdf_file=pdf_path.name,
                            serial_number=serial_number,
                            term=spec.term,
                            page=p,
                            number=selected,
                            units=None,
                            context=line_text.strip()[:200],
                            method="pymupdf:line-geom",
                            found=True,
                        )
                    # For numbers: extract from selected field
                    nums = [m.group(0) for m in NUMBER_REGEX.finditer(selected)]
                    nums += [m.group(0) for m in DATE_REGEX.finditer(selected)]
                    chosen_num = None
                    fallback_num = None
                    for n in nums:
                        range_violation = False
                        if spec.range_min is not None or spec.range_max is not None:
                            try:
                                v = float((numeric_only(n) or '').replace(',', ''))
                                if spec.range_min is not None and v < spec.range_min:
                                    range_violation = True
                                if spec.range_max is not None and v > spec.range_max:
                                    range_violation = True
                            except Exception:
                                range_violation = False
                        units_match = True
                        if spec.units_hint:
                            u = extract_units(n)
                            units_match = bool(u and any(u.lower() == h.lower() for h in spec.units_hint))
                        if chosen_num is None and not range_violation and units_match:
                            chosen_num = (n, range_violation)
                        if fallback_num is None:
                            fallback_num = (n, range_violation)
                    selected_num = chosen_num or fallback_num
                    if selected_num:
                        num_text, range_violation = selected_num
                        number_out = num_text.strip()
                        if range_violation and not number_out.rstrip().endswith('(range violation)'):
                            number_out = f"{number_out} (range violation)"
                        try:
                            doc.close()
                        except Exception:
                            pass
                        return MatchResult(
                            pdf_file=pdf_path.name,
                            serial_number=serial_number,
                            term=spec.term,
                            page=p,
                            number=number_out,
                            units=extract_units(num_text),
                            context=line_text.strip()[:200],
                            method="pymupdf:line-geom",
                            found=True,
                            confidence=None,
                            row_label=(spec.anchor or spec.term or None),
                            column_label=(spec.column or f"field_{idx}"),
                            text_source="pdf",
                        )
                if before_on_page:
                    before_triggered = True
            try:
                doc.close()
            except Exception:
                pass
        except Exception:
            # fall through to text-based approach
            pass

    # 1) Text-based approach if geometry path was unavailable or failed
    # Build text for constrained pages (or whole doc) � prefer cache
    key = _pdf_cache_key(pdf_path)
    if key in _PAGE_TEXT_CACHE:
        full_map, pipeline, _pc = _PAGE_TEXT_CACHE[key]
        if spec.pages:
            page_text_map = {p: (full_map.get(p) or "") for p in spec.pages}
        else:
            page_text_map = dict(full_map)
    else:
        page_text_map, pipeline = extract_pages_text(pdf_path, spec.pages if spec.pages else list(range(1, 10000)), do_ocr_fallback=False)
    anchor = (spec.anchor or spec.term or "")
    if not case_sensitive:
        anchor_cmp = anchor.lower()
    else:
        anchor_cmp = anchor
    split_mode = (spec.field_split or "auto").lower()
    idx = spec.field_index or 1

    def _split_fields(tail: str) -> list:
        # groups: split on 2+ spaces/tabs so single spaces remain inside a field
        fields = re.split(r"[ \t]{2,}", tail.strip()) if split_mode in ("groups", "auto") else []
        fields = [f for f in fields if f]
        if split_mode == "auto" and len(fields) <= 1:
            fields = []
        if not fields:
            fields = re.split(r"\s+", tail.strip())
            fields = [f for f in fields if f]
        return fields

    needle_after_line = spec.group_after if case_sensitive else (spec.group_after.lower() if spec.group_after else None)
    after_found_line = not bool(spec.group_after)
    before_triggered_line = False

    for p in sorted(page_text_map.keys()):
        if before_triggered_line:
            break
        text = page_text_map[p] or ""
        if not text:
            continue
        start_idx = 0
        if spec.group_after:
            idx_ga = _locate_group_anchor(text, spec.group_after, case_sensitive, after=True)
            if idx_ga is not None:
                after_found_line = True
                start_idx = idx_ga
            elif not after_found_line:
                continue
        end_idx = len(text)
        if spec.group_before:
            search_segment = text[start_idx:]
            idx_gb = _locate_group_anchor(search_segment, spec.group_before, case_sensitive, after=False)
            if idx_gb is not None:
                end_idx = start_idx + idx_gb
                before_triggered_line = True
        text_segment = text[start_idx:end_idx]
        if not text_segment.strip():
            continue
        lines = text_segment.splitlines()
        best = None  # (fields, line)
        for line in lines:
            hay = line if case_sensitive else line.lower()
            pos = hay.find(anchor_cmp) if anchor_cmp else 0
            if pos == -1:
                continue
            tail = line[pos + len(anchor):] if anchor else line
            fields = _split_fields(tail)
            if len(fields) >= idx:
                best = (fields, line)
                break  # take first satisfying occurrence on page
        if best:
            fields, line = best
            selected = fields[idx - 1].strip()
            # Normalize: collapse inner whitespace to single, strip leading punctuation like ':'
            selected = re.sub(r"\s+", " ", selected).lstrip(":-Ã¢â‚¬â€œÃ¢â‚¬â€ ")
            if (spec.return_type or "number").lower() == "string":
                return MatchResult(
                    pdf_file=pdf_path.name,
                    serial_number=serial_number,
                    term=spec.term,
                    page=p,
                    number=selected,
                    units=None,
                    context=line.strip()[:200],
                    method=f"text:line",
                    found=True,
                    confidence=None,
                    row_label=(spec.anchor or spec.term or None),
                    column_label=(spec.column or f"field_{idx}"),
                    text_source="pdf",
                )
            # Return number: search within selected field, preferring in-range values
            nums = [m.group(0) for m in NUMBER_REGEX.finditer(selected)]
            nums += [m.group(0) for m in DATE_REGEX.finditer(selected)]
            chosen_num = None
            fallback_num = None
            for n in nums:
                range_violation = False
                if spec.range_min is not None or spec.range_max is not None:
                    try:
                        v = float((numeric_only(n) or '').replace(',', ''))
                        if spec.range_min is not None and v < spec.range_min:
                            range_violation = True
                        if spec.range_max is not None and v > spec.range_max:
                            range_violation = True
                    except Exception:
                        range_violation = False
                units_match = True
                if spec.units_hint:
                    u = extract_units(n)
                    units_match = bool(u and any(u.lower()==h.lower() for h in spec.units_hint))
                if chosen_num is None and not range_violation and units_match:
                    chosen_num = (n, range_violation)
                if fallback_num is None:
                    fallback_num = (n, range_violation)
            selected_num = chosen_num or fallback_num
            if selected_num:
                num_text, range_violation = selected_num
                number_out = num_text.strip()
                if range_violation and not number_out.rstrip().endswith('(range violation)'):
                    number_out = f"{number_out} (range violation)"
                return MatchResult(
                    pdf_file=pdf_path.name,
                    serial_number=serial_number,
                    term=spec.term,
                    page=p,
                    number=number_out,
                    units=extract_units(num_text),
                    context=line.strip()[:200],
                    method=f"text:line",
                    found=True,
                    confidence=None,
                    row_label=(spec.anchor or spec.term or None),
                    column_label=(spec.column or f"field_{idx}"),
                    text_source="pdf",
                )
            # If no number matched, fall back to nearest later
    # Fallback to nearest with filters
    fallback_res = scan_pdf_for_term_nearest(pdf_path, serial_number, spec, window_chars, case_sensitive)
    if not fallback_res.found and not fallback_res.error_reason:
        fallback_res.error_reason = "No line field matched the requested index"
    return fallback_res
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
        error_cols = ["pdf_file", "serial_number", "term", "error", "method", "page", "column", "row", "group_after", "group_before"]
        if errors_rows:
            df_errors = pd.DataFrame(errors_rows, columns=error_cols)
        else:
            df_errors = pd.DataFrame(columns=error_cols)

        # Write Excel with two sheets
        with pd.ExcelWriter(output_xlsx, engine="xlsxwriter") as writer:
            # Sheet 1: wide matrix of results
            df_results.to_excel(writer, sheet_name="results", index=False)
            # Sheet 2: detailed metadata
            df_meta.to_excel(writer, sheet_name="metadata", index=False)
            # Sheet 3: failures/errors summary
            df_errors.to_excel(writer, sheet_name="errors", index=False)

            # Cosmetic improvements: freeze header rows and set reasonable column widths
            ws_res = writer.sheets["results"]
            ws_meta = writer.sheets["metadata"]
            ws_err = writer.sheets["errors"]
            ws_res.freeze_panes(1, 0)
            ws_meta.freeze_panes(1, 0)
            ws_err.freeze_panes(1, 0)

            # Auto-size columns based on max content length (capped)
            for i, col in enumerate(df_results.columns):
                width = min(60, max(10, int(df_results[col].astype(str).str.len().max() if not df_results.empty else len(col)) + 2))
                ws_res.set_column(i, i, width)
            for i, col in enumerate(df_meta.columns):
                width = min(60, max(10, int(df_meta[col].astype(str).str.len().max() if not df_meta.empty else len(col)) + 2))
                ws_meta.set_column(i, i, width)
            for i, col in enumerate(df_errors.columns):
                width = min(60, max(10, int(df_errors[col].astype(str).str.len().max() if not df_errors.empty else len(col)) + 2))
                ws_err.set_column(i, i, width)

        print(f"[DONE] Excel written -> {output_xlsx}")
        return

    # ---------- CSV fallback path ----------
    results_csv = csv_fallback_prefix.with_suffix(".results.csv")
    metadata_csv = csv_fallback_prefix.with_suffix(".metadata.csv")
    errors_csv = csv_fallback_prefix.with_suffix(".errors.csv")

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
    meta_cols = ["pdf_file", "serial_number", "term", "found", "page", "number", "context", "method_pipeline", "error_reason"]
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
                r.get("error_reason"),
            ])
    with errors_csv.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["pdf_file", "serial_number", "term", "error", "method", "page", "column", "row"])
        for r in errors_rows:
            writer.writerow([
                r.get("pdf_file"),
                r.get("serial_number"),
                r.get("term"),
                r.get("error"),
                r.get("method"),
                r.get("page"),
                r.get("column"),
                r.get("row"),
            ])
    print(f"[DONE] CSV fallback written -> {results_csv}, {metadata_csv}, {errors_csv}")


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
    # No per-PDF directory output; keep only summary JSON and flat Excel

    # Reroute output paths into the run_dir regardless of CLI-provided paths.
    output_json = run_dir / "scan_results.json"
    output_xlsx = run_dir / "scan_results_flat.xlsx"
    print(f"[INFO] Outputs will be saved under: {run_dir}")

    # Prepare structures for the wide "results" sheet and the "metadata" sheet
    term_order = [t.term for t in terms]                  # preserve input order
    term_pages_raw = {t.term: t.pages_raw for t in terms} # map term ÃƒÂ¢Ã¢â‚¬Â Ã¢â‚¬â„¢ original "Pages" string
    results_matrix: Dict[str, Dict[str, Optional[str]]] = {t.term: {} for t in terms}  # term ÃƒÂ¢Ã¢â‚¬Â Ã¢â‚¬â„¢ {SN ÃƒÂ¢Ã¢â‚¬Â Ã¢â‚¬â„¢ number}
    metadata_rows: List[Dict] = []  # detailed records per (pdf, term)
    summary: List[Dict] = []        # JSON audit entries
    errors_rows: List[Dict] = []    # rows for the errors report

    # Step 3: For each PDF, scan for each term
    for pdf_path in sorted(pdfs):
        # Derive the serial number from the filename
        serial_number = get_serial_number_from_filename(pdf_path)
        print(f"[INFO] Scanning: {pdf_path.name}  ({serial_number})")

        # No per-PDF artifact collection needed

        # Pre-extract text for all needed pages once per PDF (includes OCR fallback as configured)
        try:
            page_count = get_pdf_page_count(pdf_path)
        except Exception:
            page_count = 0
        needs_all = False
        union_pages: set[int] = set()
        for t in terms:
            if getattr(t, 'pages', None):
                for p in t.pages:
                    if isinstance(p, int) and p > 0:
                        union_pages.add(p)
            else:
                needs_all = True
        if needs_all or not union_pages:
            if page_count <= 0:
                # Fallback: let extract_pages_text deal with page bounds dynamically
                pages_for_extract = list(sorted(union_pages)) or list(range(1, 10000))
            else:
                pages_for_extract = list(range(1, page_count + 1))
        else:
            pages_for_extract = list(sorted(union_pages))

        # Honor FORCE_OCR for pre-extraction only if explicitly set; otherwise delay OCR until a term misses
        try:
            _force_ocr_pre = (os.environ.get('FORCE_OCR','') or '').strip().lower() in ('1','true','yes','force','always')
        except Exception:
            _force_ocr_pre = False
        pre_map, pre_pipe = extract_pages_text(pdf_path, pages_for_extract, do_ocr_fallback=_force_ocr_pre)
        _PAGE_TEXT_CACHE[_pdf_cache_key(pdf_path)] = (pre_map, pre_pipe, page_count)
        try:
            print(f"[INFO] Pre-extracted {len(pre_map)} page(s) via: {pre_pipe}")
        except Exception:
            pass

        # Search each configured term within the allowed page ranges
        total_terms = len(terms)
        completed = 0
        prev_pct = -1
        # Initial progress line
        try:
            print(f"[PROGRESS] Terms: 0% (0/{total_terms})")
        except Exception:
            pass

        for idx, t in enumerate(terms, start=1):
            mode = (t.mode or "").lower() if hasattr(t, 'mode') else ""
            if mode == "line":
                res = scan_pdf_for_term_line(pdf_path, serial_number, t, window_chars, case_sensitive)
            elif mode in ("table(xy)", "xy", "table") or (not mode and getattr(t, 'line', None) and getattr(t, 'column', None)):
                res = scan_pdf_for_term_xy(pdf_path, serial_number, t, window_chars, case_sensitive)
            else:
                res = scan_pdf_for_term_nearest(pdf_path, serial_number, t, window_chars, case_sensitive)
            # Normalize units/value when found
            ret_kind = (getattr(t, 'return_type', None) or 'number').strip().lower()
            if res.found and ret_kind != 'string':
                if res.number is not None:
                    note_suffix = ''
                    base_value = res.number
                    if isinstance(res.number, str):
                        trimmed = res.number.rstrip()
                        suffix = ' (range violation)'
                        if trimmed.endswith(suffix):
                            base_value = trimmed[:-len(suffix)].rstrip()
                            note_suffix = suffix
                        else:
                            base_value = res.number
                    if res.units is None and isinstance(base_value, str):
                        res.units = extract_units(base_value)
                    clean_number = numeric_only(base_value) if isinstance(base_value, str) else numeric_only(res.number)
                    if clean_number is not None:
                        res.number = clean_number + note_suffix if note_suffix else clean_number
                    elif note_suffix and isinstance(base_value, str):
                        res.number = f"{base_value}{note_suffix}"

            # Fill the matrix cell for this (term, serial_number)
            if res.found:
                if ret_kind == 'string':
                    cell_value = res.number
                else:
                    cell_value = res.number
                results_matrix.setdefault(t.term, {})[serial_number] = cell_value
            else:
                err_msg = res.error_reason or "No match found"
                results_matrix.setdefault(t.term, {})[serial_number] = f"ERROR: {err_msg}"
                errors_rows.append({
                    "pdf_file": res.pdf_file,
                    "serial_number": res.serial_number,
                    "term": res.term,
                    "error": err_msg,
                    "page": res.page,
                    "method": res.method,
                    "column": res.column_label,
                    "row": res.row_label,
                    "group_after": getattr(t, 'group_after', None),
                    "group_before": getattr(t, 'group_before', None),
                })

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
                "method_pipeline": res.method,
                "text_source": res.text_source,
                "confidence": res.confidence,
                "row_label": res.row_label,
                "column_label": res.column_label,
                # Terms schema hints for transparency/debug
                "mode": ((t.mode or ("table(xy)" if getattr(t, 'line', None) and getattr(t, 'column', None) else "nearest")) if hasattr(t, 'mode') else "nearest"),
                "pages_raw": getattr(t, 'pages_raw', ""),
                "line": getattr(t, 'line', None),
                "column": getattr(t, 'column', None),
                "range_min": getattr(t, 'range_min', None),
                "range_max": getattr(t, 'range_max', None),
                "units_hint": getattr(t, 'units_hint', None),
                "anchor": getattr(t, 'anchor', None),
                "field_index": getattr(t, 'field_index', None),
                "field_split": getattr(t, 'field_split', None),
                "return_type": getattr(t, 'return_type', None),
                "group_after": getattr(t, 'group_after', None),
                "group_before": getattr(t, 'group_before', None),
                "value_format": getattr(t, 'value_format', None),
                "error_reason": res.error_reason,
            }
            metadata_rows.append(meta)
            summary.append(meta)
            # No per-PDF accumulation

            # Update and print progress for this PDF's terms
            try:
                completed = idx
                pct = int((completed * 100) / max(1, total_terms))
                # Print at meaningful increments to avoid flooding the console
                if pct != prev_pct and (total_terms <= 20 or pct % 5 == 0 or completed == total_terms):
                    print(f"[PROGRESS] Terms: {pct}% ({completed}/{total_terms})")
                    prev_pct = pct
            except Exception:
                pass

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

        # No per-PDF JSON output

            # continue even if per-PDF write failed

            
        # Finalize per-PDF terms progress to 100%
        try:
            print(f"[PROGRESS] Terms: 100% ({total_terms}/{total_terms})")
        except Exception:
            pass

    # Output: Flat extraction table as Excel and details JSON
    # Columns for the flat (Excel/CSV) extraction table:
    # - omit verbose context
    # - include row/column labels used to derive the value
    # - include actual OCR confidence where available
    cols_display = [
        "pdf_file", "serial_number", "term",
        "row", "column",
        "found", "page",
        "number", "range_min", "range_max", "units",
        "method_pipeline", "text_source", "confidence",
        "group_after", "group_before", "error_reason",
    ]
    wrote_xlsx = False
    if _HAVE_PANDAS and _HAVE_OPENPYXL_OR_XLSXWRITER:
        try:
            import pandas as _pd
            rows_for_df = []
            for row in summary:
                row_val = row.get("row_label") or row.get("line") or ""
                col_val = row.get("column_label") or row.get("column") or ""
                rows_for_df.append({
                    "pdf_file": row.get("pdf_file"),
                    "serial_number": row.get("serial_number"),
                    "term": row.get("term"),
                    "row": row_val,
                    "column": col_val,
                    "found": row.get("found"),
                    "page": row.get("page"),
                    "number": row.get("number"),
                    "range_min": row.get("range_min"),
                    "range_max": row.get("range_max"),
                    "units": row.get("units"),
                    "method_pipeline": row.get("method_pipeline"),
                    "text_source": row.get("text_source"),
                    "confidence": row.get("confidence"),
                    "group_after": row.get("group_after") or "",
                    "group_before": row.get("group_before") or "",
                    "error_reason": row.get("error_reason") or "",
                })
            df = _pd.DataFrame(rows_for_df, columns=cols_display)
            df_errors = df[df["found"] == False].copy()
            with _pd.ExcelWriter(output_xlsx, engine="xlsxwriter") as writer:
                df.to_excel(writer, sheet_name="extraction", index=False)
                if df_errors.empty:
                    df_errors = _pd.DataFrame(columns=cols_display)
                df_errors.to_excel(writer, sheet_name="errors", index=False)
                ws = writer.sheets["extraction"]
                # Freeze header row
                ws.freeze_panes(1, 0)
                # Auto-size columns based on content
                for i, col in enumerate(df.columns):
                    try:
                        max_len = int(df[col].astype(str).map(len).max()) if not df.empty else len(col)
                    except Exception:
                        max_len = len(col)
                    ws.set_column(i, i, min(60, max(10, max_len + 2)))
                ws_err = writer.sheets["errors"]
                ws_err.freeze_panes(1, 0)
                for i, col in enumerate(df_errors.columns):
                    try:
                        max_len = int(df_errors[col].astype(str).map(len).max()) if not df_errors.empty else len(col)
                    except Exception:
                        max_len = len(col)
                    ws_err.set_column(i, i, min(60, max(10, max_len + 2)))
            wrote_xlsx = True
            print(f"[DONE] Extraction table -> {output_xlsx}")
        except Exception as e:
            print(f"[WARN] Could not write Excel extraction table: {e}")
    if not wrote_xlsx:
        # Fallback: write CSV next to intended xlsx (same basename) if Excel writer not available
        try:
            fallback_csv = output_xlsx.with_suffix(".csv")
            with fallback_csv.open("w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(cols_display)
                for row in summary:
                    row_val = row.get("row_label") or row.get("line") or ""
                    col_val = row.get("column_label") or row.get("column") or ""
                    w.writerow([
                        row.get("pdf_file"), row.get("serial_number"), row.get("term"),
                        row_val, col_val,
                        row.get("found"), row.get("page"), row.get("number"), row.get("range_min"), row.get("range_max"), row.get("units"), row.get("method_pipeline"),
                        row.get("text_source"), row.get("confidence"), row.get("group_after") or "", row.get("group_before") or "", row.get("error_reason") or ""
                    ])
            err_csv = output_xlsx.with_suffix(".errors.csv")
            with err_csv.open("w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(cols_display)
                for row in summary:
                    if row.get("found"):
                        continue
                    row_val = row.get("row_label") or row.get("line") or ""
                    col_val = row.get("column_label") or row.get("column") or ""
                    w.writerow([
                        row.get("pdf_file"), row.get("serial_number"), row.get("term"),
                        row_val, col_val,
                        row.get("found"), row.get("page"), row.get("number"), row.get("range_min"), row.get("range_max"), row.get("units"), row.get("method_pipeline"),
                        row.get("text_source"), row.get("confidence"), row.get("group_after") or "", row.get("group_before") or "", row.get("error_reason") or ""
                    ])
            print(f"[DONE] Extraction table (CSV fallback) -> {fallback_csv}; errors -> {err_csv}")
        except Exception as e:
            print(f"[WARN] Could not write extraction table fallback: {e}")

    print(f"[DONE] Details JSON -> {output_json}")

    # Remove legacy aggregate artifact if present
    try:
        agg_path = Path("Product_Data_File") / "EIDP_data.csv"
        if agg_path.exists():
            agg_path.unlink(missing_ok=True)  # type: ignore[call-arg]
            print(f"[CLEANUP] Removed legacy aggregate -> {agg_path}")
    except Exception:
        pass

    # Update the persistent run registry with all serial numbers in this run
    try:
        run_sns: List[str] = []
        # Prefer keys discovered in results_matrix
        for term, sn_map in results_matrix.items():
            for sn in sn_map.keys():
                if sn not in run_sns:
                    run_sns.append(sn)
        if run_sns:
            _update_run_registry(run_dir, run_sns)
            print("[DONE] Run registry updated (run_registry.xlsx)")
    except Exception as e:
        print(f"[WARN] Could not update run registry: {e}")

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
    parser.add_argument("--input", required=True, help="Path to terms file (.csv, .xlsx, or .xls). Headers: Term, Pages [Line, Column, Range, Units optional]")
    parser.add_argument("--pdf-folder", required=True, help='Folder containing PDFs to scan (e.g., "EIDP import folder")')
    parser.add_argument("--output-csv", default="scan_results_flat.csv", help="Flat CSV summary (legacy)")
    parser.add_argument("--output-json", default="scan_results.json", help="Path to write JSON details")
    parser.add_argument("--output-xlsx", default="scan_results.xlsx", help="Excel workbook with 'results' and 'metadata' sheets")
    parser.add_argument("--scanned-folder", default="Scanned Docs", help="Folder to move scanned PDFs into")
    parser.add_argument("--window-chars", type=int, default=160, help="Search window size around term (Ãƒâ€šÃ‚Â± chars)")
    parser.add_argument("--case-sensitive", action="store_true", help="Enable case-sensitive term matching")
    parser.add_argument("--quiet", action="store_true", help="Reduce console output (suppress progress/debug)")
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

    # Apply quiet mode if requested (affects filtered print)
    if getattr(args, "quiet", False):
        global _QUIET  # type: ignore[global-variable-not-assigned]
        _QUIET = True

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










