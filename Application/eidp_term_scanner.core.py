
#!/usr/bin/env python3
# Application-consolidated build
"""
EIDP Term Scanner (Matrix + Metadata)
-------------------------------------
- Scans PDFs in a given folder for configured "terms" within specified page ranges.
- Extracts the closest numeric value near each term occurrence.
- A serial component (data identifier) is inferred from each PDF's filename; results are arranged
  as a wide matrix: rows=terms, columns=data identifiers (one per EIDP).
- Produces an Excel workbook with:
    * "results"  : Term, Pages, and one column per data identifier with the matched number
    * "metadata" : detailed per-term/per-file records, for auditing/debugging
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
from typing import Any, Dict, List, Optional, Sequence, Tuple

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
        # Allow [PROGRESS] messages even in quiet mode (needed for UI progress tracking)
        # if text.startswith("[PROGRESS]"):
        #     return False
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


# --- Optional: table extraction helper (scripts/extract_page_tables.py) ---
def _import_tables_module():
    try:
        import importlib
        return importlib.import_module('scripts.extract_page_tables')
    except Exception:
        return None


def _filter_rows_by_anchors(rows: List[List[str]], group_after: Optional[str], group_before: Optional[str]) -> List[List[str]]:
    if not rows:
        return []
    ga = (group_after or '').strip()
    gb = (group_before or '').strip()
    if not ga and not gb:
        return rows
    start_idx = 0
    end_idx = len(rows)
    found_any = False
    if ga:
        needle = ga.lower()
        for i, r in enumerate(rows):
            line = " ".join(str(x or '') for x in r).lower()
            if needle and (needle in line):
                start_idx = i + 1
                found_any = True
                break
    if gb:
        needle = gb.lower()
        for i in range(start_idx, len(rows)):
            line = " ".join(str(x or '') for x in rows[i]).lower()
            if needle and (needle in line):
                end_idx = i
                found_any = True or found_any
                break
    if not found_any:
        # Default to full page content when anchors absent
        return rows
    return rows[start_idx:end_idx]


def _extract_full_table_rows_for_pdf(
    pdf_path: Path,
    pages: List[int],
    group_after: Optional[str],
    group_before: Optional[str],
    program_name: Optional[str],
    vehicle_number: Optional[str],
    serial_component: Optional[str],
) -> List[Dict[str, Optional[str]]]:
    mod = _import_tables_module()
    if mod is None:
        return []
    # OCR config
    try:
        ocr_mode = _get_ocr_mode()
    except Exception:
        ocr_mode = 'fallback'
    use_ocr = (ocr_mode != 'no_ocr')
    try:
        dpi = int((os.environ.get('OCR_DPI') or '600').strip())
    except Exception:
        dpi = 600
    try:
        min_conf = float((os.environ.get('EASYOCR_MIN_CONF') or '0.4').strip())
    except Exception:
        min_conf = 0.4
    try:
        langs_raw = (os.environ.get('EASYOCR_LANGS') or 'en').strip()
        import re as _re
        langs = [s.strip() for s in _re.split(r'[;,]', langs_raw) if s.strip()]
        if not langs:
            langs = ['en']
    except Exception:
        langs = ['en']

    try:
        tables_map = mod.extract_tables_for_pages(pdf_path, pages, use_ocr=bool(use_ocr), dpi=int(dpi), min_conf=float(min_conf), langs=langs, emit_tokens=False)  # type: ignore[attr-defined]
    except Exception:
        return []
    out_rows: List[Dict[str, Optional[str]]] = []
    for p in sorted(tables_map.keys()):
        spec = tables_map.get(p) or {}
        columns = spec.get('columns') or []
        rows = spec.get('rows') or []
        # Expect first column to be Section per implementation; still handle generically
        filtered = _filter_rows_by_anchors(rows, group_after, group_before)
        for r in filtered:
            row_out: Dict[str, Optional[str]] = {
                'pdf_file': pdf_path.name,
                'program_name': program_name or '',
                'vehicle_number': vehicle_number or '',
                'serial_component': serial_component or '',
                'page': str(p),
            }
            # Section is usually first column if present
            if r:
                row_out['section'] = str(r[0]) if r[0] is not None else ''
            else:
                row_out['section'] = ''
            # Remaining cells -> col_1, col_2, ...
            for i, cell in enumerate(r[1:] if r else [], start=1):
                row_out[f'col_{i}'] = str(cell) if cell is not None else ''
            out_rows.append(row_out)
    return out_rows


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
    - term_label / data_group: optional reporting metadata (not used for matching)
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
    term_label: Optional[str] = None          # Friendly label for outputs (optional)
    data_group: Optional[str] = None          # User-defined grouping bucket (optional)
    mode: Optional[str] = None
    line: Optional[str] = None
    column: Optional[str] = None
    anchor: Optional[str] = None
    field_index: Optional[int] = None
    field_split: Optional[str] = None
    return_type: Optional[str] = None
    range_min: Optional[float] = None
    range_max: Optional[float] = None
    range_min_disabled: bool = False  # True when user sets Range (min) to N/A
    range_max_disabled: bool = False  # True when user sets Range (max) to N/A
    units_hint: List[str] = field(default_factory=list)
    # New schema helpers
    value_format: Optional[str] = None      # Optional expected value pattern (e.g., tpl-xxxx or /TPL-\d{4}/)
    group_after: Optional[str] = None       # Optional anchor text; only consider matches appearing after this text
    group_before: Optional[str] = None      # Optional anchor text; only consider matches appearing before this text
    # Smart Snap mode hint (optional): 'number' | 'date' | 'time' | 'title' | 'auto'
    smart_snap_type: Optional[str] = None
    # Optional secondary label to disambiguate values within a row
    secondary_term: Optional[str] = None
    # Optional positional pick (1-based) of the Nth value to the right (smart mode only)
    smart_position: Optional[int] = None


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
    # Smart Snap debugging
    smart_snap_context: Optional[str] = None
    smart_snap_type: Optional[str] = None
    # Smart Snap extras
    smart_line_min: Optional[str] = None
    smart_line_max: Optional[str] = None
    smart_conflict: Optional[str] = None
    smart_secondary_found: Optional[bool] = None
    # Optional breakdown of numeric candidate scoring components (smart mode)
    smart_score_breakdown: Optional[Dict[str, float]] = None
    # How the value was selected in smart mode: 'smart_position' or 'smart_score'
    smart_selection_method: Optional[str] = None


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

DATE_REGEX = re.compile(r"\b(?:\d{1,2}/\d{1,2}/\d{2,4}|\d{4}-\d{2}-\d{2})\b")

DEFAULT_SMART_FORMATS = {
    "title": r"/[A-Za-z0-9\s\/_\-().]+/",
    "date": r"/(?:\d{1,2}\/\d{1,2}\/\d{2,4}|\d{4}-\d{2}-\d{2})/",
}


def _effective_value_format(spec) -> Optional[str]:
    raw = getattr(spec, "value_format", None)
    if raw:
        return raw
    smart_kind = (getattr(spec, "smart_snap_type", "") or "").strip().lower()
    return DEFAULT_SMART_FORMATS.get(smart_kind)
TIME_REGEX = re.compile(r"\b(?:(?:[01]?\d|2[0-3]):[0-5]\d(?::[0-5]\d)?\s*(?:[AP]M|[ap]m)?|(?:\d+\s*(?:ms|s|sec|mins?|minutes?|hrs?|hours?)))\b")



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


# Regex to capture legacy SN tokens like "... SN 1234", "... SN-ABC_09", etc.
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
    if v == "smart":
        return "smart"
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
    # and words separated by 1ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Å“2 spaces remain within the same field.
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


def _norm_smart_type(s: Optional[str]) -> Optional[str]:
    if not s:
        return None
    v = str(s).strip().lower()
    if not v:
        return None
    if v in ("auto", "any", "detect"):
        return "auto"
    if v in ("num", "number", "value"):
        return "number"
    if v in ("date", "dt"):
        return "date"
    if v in ("time", "tm"):
        return "time"
    if v in ("title", "text", "name", "string"):
        return "title"
    return v


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
                        smart_snap_type = None
                        secondary_term = None
                        smart_position: Optional[int] = None
                        data_group = None
                        term_label = None
                        for k, v in row.items():
                            if k and k.strip().lower() == "term":
                                term = (v or "").strip()
                            if k and k.strip().lower() in ("term_label", "term label"):
                                term_label = ((v or "").strip() or None)
                            if k and k.strip().lower() in ("data_group", "data group", "datagroup"):
                                data_group = ((v or "").strip() or None)
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
                            if k and k.strip().lower() in ("smart_snap_type", "smart", "smart_snap", "smart type"):
                                smart_snap_type = _norm_smart_type((v or "").strip())
                            if k and k.strip().lower() in ("secondary_term", "secondary", "secondary label", "secondary_label"):
                                secondary_term = ((v or "").strip() or None)
                            if k and k.strip().lower() in ("smart_position", "smart position", "smartpos"):
                                try:
                                    smart_position = int(str(v).strip()) if str(v).strip() else None
                                    if smart_position is not None and smart_position < 1:
                                        smart_position = None
                                except Exception:
                                    smart_position = None
                        if term:
                            result.append(TermSpec(term=term,
                                                   pages=parse_page_ranges(pages_str),
                                                   pages_raw=pages_str,
                                                   term_label=term_label,
                                                   data_group=data_group,
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
                                                   group_before=group_before,
                                                   smart_snap_type=smart_snap_type,
                                                   secondary_term=secondary_term,
                                                   smart_position=smart_position))
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

    # Build a map of header name ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ column index
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
    term_label_col = col_for("term_label") or col_for("term label")
    data_group_col = col_for("data_group") or col_for("data group") or col_for("datagroup")
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
    smart_type_col = col_for("smart_snap_type") or col_for("smart") or col_for("smart snap type") or col_for("smart_snap")
    secondary_col = col_for("secondary_term") or col_for("secondary") or col_for("secondary label") or col_for("secondary_label")
    smart_pos_col = col_for("smart_position") or col_for("smart position") or col_for("smartpos")
    if not term_col:
        print("[ERROR] Could not find 'Term' header in Excel file.", file=sys.stderr)
        sys.exit(2)

    # Walk rows and collect terms
    def _is_template_metadata(row_cells: Sequence[Any]) -> bool:
        try:
            hay = " ".join(str(cell.value or "").strip().lower() for cell in row_cells)
        except Exception:
            return False
        return (
            "data group" in hay
            and "term label" in hay
            and "smart snap type" in hay
        )

    for row in ws.iter_rows(min_row=2):
        try:
            row_idx = row[0].row if row and row[0] is not None else None
        except Exception:
            row_idx = None
        # Skip the explanatory second header in Smart‑Snap schema templates
        try:
            if (
                row_idx == 2
                and input_path.name.lower().endswith('terms.schema.smartsnap.xlsx')
                and _is_template_metadata(row)
            ):
                continue
        except Exception:
            pass
        term_val = row[term_col - 1].value if term_col else None
        term_label_val = row[term_label_col - 1].value if term_label_col else None
        data_group_val = row[data_group_col - 1].value if data_group_col else None
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
        smart_type_val = row[smart_type_col - 1].value if smart_type_col else None
        sec_val = row[secondary_col - 1].value if secondary_col else None
        smart_pos_val = row[smart_pos_col - 1].value if smart_pos_col else None
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
        rmin_disabled = rmax_disabled = False
        if range_val is not None and str(range_val).strip():
            rmin, rmax = parse_range(str(range_val).strip())
        if rmin_val is not None and str(rmin_val).strip():
            raw = str(rmin_val).strip()
            if _is_na_token(raw):
                rmin = None
                rmin_disabled = True
            else:
                try:
                    rmin = float(str(rmin_val).replace(',', ''))
                except Exception:
                    pass
        if rmax_val is not None and str(rmax_val).strip():
            raw = str(rmax_val).strip()
            if _is_na_token(raw):
                rmax = None
                rmax_disabled = True
            else:
                try:
                    rmax = float(str(rmax_val).replace(',', ''))
                except Exception:
                    pass
        units_hint = parse_units_hint(units_val)
        smart_snap_type = _norm_smart_type(str(smart_type_val).strip() if smart_type_val is not None else None)
        secondary_term = (str(sec_val).strip() if sec_val is not None and str(sec_val).strip() else None)
        try:
            smart_position = int(str(smart_pos_val).strip()) if smart_pos_val is not None and str(smart_pos_val).strip() else None
            if smart_position is not None and smart_position < 1:
                smart_position = None
        except Exception:
            smart_position = None
        if term:
            term_label = (str(term_label_val).strip() if term_label_val is not None and str(term_label_val).strip() else None)
            data_group = (str(data_group_val).strip() if data_group_val is not None and str(data_group_val).strip() else None)
            terms.append(TermSpec(term=term, pages=parse_page_ranges(pages_str), pages_raw=pages_str,
                                   term_label=term_label, data_group=data_group,
                                   mode=mode, line=line, column=column, anchor=anchor,
                                   field_index=field_index, field_split=field_split, return_type=return_type,
                                  range_min=rmin, range_max=rmax, units_hint=units_hint,
                                  range_min_disabled=rmin_disabled, range_max_disabled=rmax_disabled,
                                   value_format=(str(fmt_val).strip() if fmt_val is not None and str(fmt_val).strip() else None),
                                   group_after=(str(grp_val).strip() if grp_val is not None and str(grp_val).strip() else None),
                                   group_before=(str(grp_before_val).strip() if grp_before_val is not None and str(grp_before_val).strip() else None),
                                   smart_snap_type=smart_snap_type,
                                   secondary_term=secondary_term,
                                   smart_position=smart_position))
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
        term_label = str(get(row, 'term_label') or '').strip() or None
        data_group = str(get(row, 'data_group') or '').strip() or None
        rng = str(get(row, 'range') or '').strip()
        rmin = rmax = None
        rmin_disabled = rmax_disabled = False
        if rng:
            rmin, rmax = parse_range(rng)
        # Direct min/max override
        _rmin = str(get(row, 'range (min)') or '').strip()
        _rmax = str(get(row, 'range (max)') or '').strip()
        if _rmin:
            if _is_na_token(_rmin):
                rmin = None
                rmin_disabled = True
            else:
                try:
                    rmin = float(_rmin.replace(',', ''))
                except Exception:
                    pass
        if _rmax:
            if _is_na_token(_rmax):
                rmax = None
                rmax_disabled = True
            else:
                try:
                    rmax = float(_rmax.replace(',', ''))
                except Exception:
                    pass
        units_hint = parse_units_hint(get(row, 'units'))
        value_format = str(get(row, 'format') or get(row, 'value_format') or '').strip() or None
        group_after = str(get(row, 'groupafter') or get(row, 'group_after') or get(row, 'group') or '').strip() or None
        group_before = str(get(row, 'groupbefore') or get(row, 'group_before') or get(row, 'beforegroup') or '').strip() or None
        secondary_term = str(get(row, 'secondary_term') or get(row, 'secondary') or get(row, 'secondary label') or get(row, 'secondary_label') or '').strip() or None
        _smart_position_raw = str(get(row, 'smart_position') or get(row, 'smart position') or get(row, 'smartpos') or '').strip()
        try:
            smart_position = int(_smart_position_raw) if _smart_position_raw else None
            if smart_position is not None and smart_position < 1:
                smart_position = None
        except Exception:
            smart_position = None
        smart_snap_type = _norm_smart_type(str(get(row, 'smart_snap_type') or get(row, 'smart') or get(row, 'smart_snap') or '').strip() or None)
        out.append(TermSpec(term=term, pages=parse_page_ranges(pages_str), pages_raw=pages_str,
                            term_label=term_label, data_group=data_group,
                            mode=mode, line=line, column=column, anchor=anchor,
                            field_index=field_index, field_split=field_split, return_type=return_type,
                                  range_min=rmin, range_max=rmax, units_hint=units_hint,
                                  range_min_disabled=rmin_disabled, range_max_disabled=rmax_disabled,
                            value_format=value_format, group_after=group_after, group_before=group_before,
                            smart_snap_type=smart_snap_type, secondary_term=secondary_term, smart_position=smart_position))
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


def _is_na_token(value: Optional[str]) -> bool:
    """Return True if the provided cell text indicates N/A."""
    if value is None:
        return False
    text = str(value).strip().lower()
    return text in {"n/a", "na", "n.a.", "not applicable"}


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


def _update_run_registry(run_dir: Path, serial_components: List[str], serial_metadata: Optional[Dict[str, Dict[str, str]]] = None) -> None:
    """Update a persistent run registry of EIDPs (identified by serial_component) and their latest run date.

    - File path: Product_Data_File/run_registry.csv (CSV only)
    - Columns: serial_component, program_name, vehicle_number, run_date, run_folder
    - On re-run, replaces the row for a serial component with the latest date and folder
    """
    try:
        exports_dir = Path("Product_Data_File")
        exports_dir.mkdir(parents=True, exist_ok=True)
        registry_csv = exports_dir / "run_registry.csv"
        columns = [
            "serial_component",
            "program_name",
            "vehicle_number",
            "run_date",
            "run_folder",
        ]
        serial_metadata = serial_metadata or {}

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

        def meta_value(sc: str, key: str) -> str:
            try:
                val = serial_metadata.get(sc, {}).get(key, "")
            except Exception:
                val = ""
            return str(val).strip() if val is not None else ""

        new_rows = {}
        for sc in serial_components:
            sc_clean = (sc or "").strip()
            if not sc_clean:
                continue
            new_rows[sc_clean] = {
                "serial_component": sc_clean,
                "program_name": meta_value(sc_clean, "program_name"),
                "vehicle_number": meta_value(sc_clean, "vehicle_number"),
                "run_date": run_date,
                "run_folder": str(run_folder),
            }
        if not new_rows:
            return

        # CSV-only registry path
        try:
            rows_map: Dict[str, Dict[str, str]] = {}
            if registry_csv.exists():
                with registry_csv.open("r", encoding="utf-8", newline="") as f:
                    r = csv.DictReader(f)
                    for row in r:
                        sc = (row.get("serial_component") or row.get("serial_number") or "").strip()
                        if sc:
                            rows_map[sc] = {col: (row.get(col) or "").strip() for col in columns}
            for sc, row in new_rows.items():
                rows_map[sc] = {col: row.get(col, "") for col in columns}
            with registry_csv.open("w", encoding="utf-8", newline="") as f:
                w = csv.DictWriter(f, fieldnames=columns)
                w.writeheader()
                for sc in sorted(rows_map.keys()):
                    w.writerow({col: rows_map[sc].get(col, "") for col in columns})
        except Exception:
            pass
        # Best-effort: remove any legacy XLSX to avoid confusion
        try:
            legacy = exports_dir / "run_registry.xlsx"
            if legacy.exists():
                legacy.unlink()
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


def _token_bounds_and_text(token: Any) -> Tuple[float, float, str]:
    """Return (x0, x1, text) for either PyMuPDF word tuples or OCR token dicts."""
    if isinstance(token, dict):
        x0 = float(token.get('x0', 0.0))
        x1 = float(token.get('x1', 0.0))
        text = str(token.get('text') or '')
        return x0, x1, text
    try:
        x0 = float(token[0])
        x1 = float(token[2])
        text = str(token[4])
        return x0, x1, text
    except Exception:
        return 0.0, 0.0, ""


def _fields_from_items(items: Sequence[Any]) -> List[str]:
    """
    Group right-of-label tokens into horizontal \"boxes\" (fields).
    Used by Smart Position so that the Nth field corresponds to the
    Nth visual column (e.g., Requirement, Measured Value, Units).
    """
    if not items:
        return []
    # Normalize to (x0, x1, text) triples
    triples: List[Tuple[float, float, str]] = []
    for it in items:
        try:
            x0, x1, text = _token_bounds_and_text(it)
        except Exception:
            continue
        txt = text.strip()
        if not txt:
            continue
        triples.append((x0, x1, txt))
    if not triples:
        return []
    triples.sort(key=lambda t: t[0])
    fields: List[str] = []
    current: List[str] = []
    prev_right: Optional[float] = None
    # Heuristic gap threshold in text units to start a new field
    GAP = 6.0
    for x0, x1, txt in triples:
        if prev_right is None:
            current = [txt]
        else:
            gap = x0 - prev_right
            if gap > GAP:
                field = " ".join(current).strip()
                if field:
                    fields.append(field)
                current = [txt]
            else:
                current.append(txt)
        prev_right = x1
    if current:
        field = " ".join(current).strip()
        if field:
            fields.append(field)
    return fields


def _column_text_for_position(
    tokens: Sequence[Any],
    column_positions: Dict[str, float],
    header_tokens: Sequence[str],
    label_right_x: float,
    pos_n: Optional[int],
    secondary_term: Optional[str],
) -> Optional[str]:
    """Return concatenated text for the desired slot using header anchor positions."""
    if not tokens or not column_positions or not header_tokens:
        return None
    ordered: List[Tuple[str, float]] = []
    for raw_name in header_tokens:
        raw = str(raw_name).strip()
        if not raw:
            continue
        name = _normalize_anchor_token(raw)
        if not name:
            continue
        cx = column_positions.get(name)
        if cx is None:
            continue
        ordered.append((name, float(cx)))
    if not ordered:
        return None
    sec_norm = _normalize_anchor_token(secondary_term) if secondary_term else ""
    target_idx: Optional[int] = None
    if sec_norm:
        for idx, (name, _) in enumerate(ordered):
            if name == sec_norm:
                target_idx = idx
                break
    if target_idx is None and pos_n and 1 <= pos_n <= len(ordered):
        target_idx = pos_n - 1
    if target_idx is None:
        return None
    target_cx = ordered[target_idx][1]
    prev_cx = ordered[target_idx - 1][1] if target_idx > 0 else None
    next_cx = ordered[target_idx + 1][1] if target_idx + 1 < len(ordered) else None
    if prev_cx is not None:
        left = (prev_cx + target_cx) / 2.0
    else:
        span = (next_cx - target_cx) if next_cx is not None else max(target_cx - label_right_x, 12.0)
        left = target_cx - max(12.0, span)
    if next_cx is not None:
        right = (target_cx + next_cx) / 2.0
    else:
        span = (target_cx - prev_cx) if prev_cx is not None else 12.0
        right = target_cx + max(12.0, span)
    left = max(left, label_right_x - 1.0)
    window_left = left - 0.5
    window_right = right + 0.5
    pieces: List[str] = []
    for token in tokens:
        x0, x1, raw_text = _token_bounds_and_text(token)
        text = raw_text.strip()
        if not text:
            continue
        cx_token = (x0 + x1) / 2.0
        if window_left <= cx_token <= window_right:
            pieces.append(text)
    if not pieces:
        return None
    return " ".join(pieces).strip()


def _token_norm(s: str) -> str:
    return _normalize_anchor_token(s)


def _match_anchor_on_line(anchor: str, line_tokens: List[str], line_token_norms: List[str]) -> Optional[Tuple[int, int]]:
    if not anchor:
        return None
    toks = [t for t in re.split(r"\s+", anchor) if t]
    toks_norm = [_token_norm(t) for t in toks]
    if not toks:
        return None
    span = len(toks)
    for i in range(0, len(line_tokens) - span + 1):
        ok = True
        for j in range(span):
            if line_tokens[i + j] == toks[j] or (toks_norm[j] and line_token_norms[i + j] == toks_norm[j]):
                continue
            ok = False
            break
        if ok:
            return i, i + span - 1
    return None

def _anchor_tokens_present(anchor: str, text: str) -> bool:
    """
    Ensure every normalized token from the anchor exists in the target text.
    Helps prevent partial matches (single-word hits) from masquerading as full row matches.
    """
    if not anchor:
        return True
    anchor_tokens = [_normalize_anchor_token(tok) for tok in re.split(r"\s+", anchor) if _normalize_anchor_token(tok)]
    if not anchor_tokens:
        return True
    text_tokens = {_normalize_anchor_token(tok) for tok in re.split(r"\s+", text) if _normalize_anchor_token(tok)}
    if not text_tokens:
        return False
    # Require at least half of the anchor tokens (rounded up) to be present.
    # This tolerates labels that are split across multiple lines (e.g.,
    # "Serial / component" rendered as "Serial" on one line and "Component"
    # on another) while still preventing spurious single-word matches.
    present = sum(1 for tok in anchor_tokens if tok in text_tokens)
    needed = max(1, (len(anchor_tokens) + 1) // 2)
    return present >= needed


def _strip_label_tokens(value: Optional[str], label: Optional[str]) -> Optional[str]:
    """
    Remove leading/trailing occurrences of the label tokens from a value.
    Used for Smart Snap \"title\" fields so that we return just the field
    contents (e.g., 'SN42-AX' instead of 'Serial SN42-AX', or the path
    instead of 'Source Bundle <path>').
    """
    if not value or not label:
        return value
    val = str(value).strip()
    if not val:
        return value
    # Tokenize and normalize
    import re as _re
    label_tokens = [_normalize_anchor_token(tok) for tok in _re.split(r"\s+", label) if _normalize_anchor_token(tok)]
    if not label_tokens:
        return value
    tokens = [tok for tok in _re.split(r"\s+", val) if tok]
    if not tokens:
        return value
    def norm(t: str) -> str:
        return _normalize_anchor_token(t)
    # Strip matching prefix tokens
    i = 0
    while i < len(tokens) and norm(tokens[i]) in label_tokens:
        i += 1
    j = len(tokens)
    # Strip matching suffix tokens
    while j > i and norm(tokens[j - 1]) in label_tokens:
        j -= 1
    trimmed = " ".join(tokens[i:j]).strip()
    # Only return trimmed if something substantive remains
    return trimmed or value


def _strip_units_from_numeric_text(value: Optional[str]) -> Optional[str]:
    """
    When Smart Snap expects a numeric value, return only the numeric portion,
    preserving any '(range violation)' suffix if present.
    """
    if not value:
        return value
    txt = value.strip()
    suffix = ""
    rv = " (range violation)"
    if txt.endswith(rv):
        txt = txt[: -len(rv)].rstrip()
        suffix = rv
    num = numeric_only(txt)
    if num is None:
        return (txt + suffix).strip()
    if isinstance(num, str) and num != txt:
        return f"{num}{suffix}"
    return (txt + suffix).strip()


def _extract_status_from_title(value: Optional[str]) -> Optional[str]:
    """
    For status-like title fields (YES/NO/PASS/FAIL/etc.), return the most
    plausible status token from the text, e.g. 'NO' from 'NO +88*C'.
    If no status token is found, return the original value.
    """
    if not value:
        return value
    text = str(value).strip()
    if not text:
        return value
    import re as _re
    words = [w for w in _re.findall(r"[A-Za-z][A-Za-z0-9_-]*", text)]
    if not words:
        return value
    up_words = [w.upper() for w in words]
    # Multi-word statuses first (e.g., 'Not Recorded')
    bigram_statuses = {
        "NOT RECORDED",
        "NOT RUN",
    }
    for i in range(len(words) - 1):
        phrase_up = f"{up_words[i]} {up_words[i+1]}"
        if phrase_up in bigram_statuses:
            return f"{words[i]} {words[i+1]}"
    # Single-word statuses
    single_statuses = {
        "YES",
        "NO",
        "PASS",
        "FAIL",
        "MISSING",
        "INCOMPLETE",
        "CONDITIONAL",
        "OPEN",
        "CLOSED",
        "APPROVED",
    }
    for w, wu in zip(words, up_words):
        if wu in single_statuses:
            return w
    return value


def _detect_smart_type(preferred: Optional[str], text: str) -> str:
    v = (preferred or 'auto').strip().lower()
    if v in ('number','date','time','title'):
        return v
    # auto detect by presence order: date > time > number > title
    if DATE_REGEX.search(text):
        return 'date'
    if TIME_REGEX.search(text):
        return 'time'
    if NUMBER_REGEX.search(text):
        return 'number'
    return 'title'


def scan_pdf_for_term_smart(pdf_path: Path, serial_number: str, spec: TermSpec, window_chars: int, case_sensitive: bool) -> MatchResult:
    # Prefer anchor if provided, else line, else term
    row_name = (spec.anchor or spec.line or spec.term or '').strip()
    pages = spec.pages
    method_used = 'smart:pdf'
    confidence = None
    text_source = 'pdf'
    row_text_selected = None
    context_line_text = None
    value_text = None
    units_value = None
    page_hit: Optional[int] = None

    # Helper to extract for one line
    value_format_text, double_height_mode = _value_format_info(_effective_value_format(spec))
    fmt_pat = _compile_value_regex(value_format_text) if value_format_text else None
    units_hints = [str(u).strip().lower() for u in (spec.units_hint or []) if str(u).strip()]
    units_hint_set = set(units_hints)
    sec_term = (getattr(spec, 'secondary_term', None) or '').strip()
    sec_norm_global = _normalize_anchor_token(sec_term) if sec_term else ""
    try:
        _SEC_HEADER_WEIGHT = float(os.environ.get("SMART_SEC_HEADER_W", "0.7"))
    except Exception:
        _SEC_HEADER_WEIGHT = 0.7
    debug_mode = bool(os.environ.get('SMART_DEBUG') or os.environ.get('SMART_SNAP_DEBUG'))

    def extract_from_line(line_text: str, right_text: str, smart_kind: str) -> Optional[str]:
        nonlocal units_value
        target_text = right_text if right_text and right_text.strip() else line_text
        if smart_kind == 'date':
            m = DATE_REGEX.search(target_text)
            return m.group(0) if m else None
        if smart_kind == 'time':
            m = TIME_REGEX.search(target_text)
            return m.group(0) if m else None
        if smart_kind == 'number':
            matches = list(NUMBER_REGEX.finditer(right_text))
            if not matches:
                return None
            pick = None
            # Prefer matches whose units align with hints
            if units_hints:
                for m in matches:
                    ui = extract_units(m.group(0))
                    if ui and ui.strip().lower() in units_hints:
                        pick = m
                        break
            if pick is None:
                pick = matches[0]
            cand = pick.group(0)
            units_value = extract_units(cand)
            # Range check
            try:
                nclean = numeric_only(cand)
                nval = float(nclean) if nclean is not None else None
            except Exception:
                nval = None
            if nval is not None and (spec.range_min is not None or spec.range_max is not None):
                bad = False
                if spec.range_min is not None and nval < spec.range_min:
                    bad = True
                if spec.range_max is not None and nval > spec.range_max:
                    bad = True
                if bad and not cand.rstrip().endswith('(range violation)'):
                    cand = f"{cand} (range violation)"
            return cand
        # title/text fallback
        # If a value_format is specified, try to extract all matches and support positional selection
        pos_n = spec.smart_position or spec.field_index
        if fmt_pat:
            matches = list(fmt_pat.finditer(target_text))
            if matches:
                if pos_n and 1 <= pos_n <= len(matches):
                    return matches[pos_n - 1].group(0)
                return matches[0].group(0)
        # Without format, split into fields (by spacing) and support positional
        fields = _split_fields_by_spacing(target_text)
        if fields:
            if pos_n and 1 <= pos_n <= len(fields):
                return fields[pos_n - 1]
            return fields[0]
        t = right_text.strip()
        return t if t else None

    # Try PyMuPDF lines first
    if _HAVE_PYMUPDF:
        try:
            doc = fitz.open(str(pdf_path))  # type: ignore[name-defined]
        except Exception:
            doc = None
        if doc is not None:
            try:
                if not pages:
                    pages = list(range(1, doc.page_count + 1))
                best_score = 0.0
                best_info = None  # (p, line_text, right_text, val, smart_kind)
                best_components: Optional[Dict[str, float]] = None  # numeric candidate scoring breakdown
                best_selection_method: Optional[str] = None  # 'smart_position' vs 'smart_score'
                best_selection_method: Optional[str] = None  # 'smart_position' vs 'smart_score'
                pdf_best_line_only: Optional[Tuple[int, str, float]] = None  # (p, line_text, score)
                # Track first occurrences of group_after/group_before across pages
                group_after_seen = spec.group_after is None
                group_after_page: Optional[int] = None
                group_before_seen = spec.group_before is None
                group_before_page: Optional[int] = None
                # prepare optional grouping thresholds based on anchors
                def _line_anchor_score(text: str, anchor: str) -> float:
                    if not anchor:
                        return 0.0
                    sc = _fuzzy_ratio(text, anchor)
                    if _normalize_anchor_token(anchor) and _normalize_anchor_token(anchor) in _normalize_anchor_token(text):
                        sc = max(sc, 0.99)
                    return sc

                for p in pages:
                    try:
                        page = doc.load_page(p - 1)
                        words = page.get_text('words') or []
                    except Exception:
                        continue
                    # Group into row bands by Y center (robust across table blocks)
                    lines_map: Dict[int, Dict] = {}
                    for w in words:
                        try:
                            x0,y0,x1,y1,txt,*_rest = w
                        except Exception:
                            if len(w) >= 5:
                                x0,y0,x1,y1,txt = w[:5]
                            else:
                                continue
                        cy = int(round((float(y0)+float(y1))/2.0))
                        entry = lines_map.get(cy)
                        if not entry:
                            entry = {'tokens': [], 'x0': float(x0), 'y0': float(y0), 'x1': float(x1), 'y1': float(y1)}
                            lines_map[cy] = entry
                        entry['tokens'].append((float(x0),float(y0),float(x1),float(y1),str(txt)))
                        entry['x0'] = min(float(entry['x0']), float(x0))
                        entry['y0'] = min(float(entry['y0']), float(y0))
                        entry['x1'] = max(float(entry['x1']), float(x1))
                        entry['y1'] = max(float(entry['y1']), float(y1))
                    # Detect optional group bounds (y coordinates) from anchors
                    ga_y = None
                    gb_y = None
                    page_group_before_y = None
                    header_tokens: Dict[str, List[Tuple[float, float, float]]] = {'min': [], 'value': [], 'max': []}
                    for e in lines_map.values():
                        for tok in e['tokens']:
                            txt_norm = str(tok[4]).strip().lower()
                            if txt_norm in header_tokens:
                                cy_tok = (float(tok[1]) + float(tok[3])) / 2.0
                                cx_tok = (float(tok[0]) + float(tok[2])) / 2.0
                                header_tokens[txt_norm].append((cy_tok, float(e['y0']), cx_tok))
                    if spec.group_after:
                        best = None
                        for _, e in lines_map.items():
                            lt = ' '.join([t[4] for t in sorted(e['tokens'], key=lambda t: (t[1], t[0]))]).strip()
                            sc = _line_anchor_score(lt, spec.group_after)
                            if sc >= 0.6:
                                y = float(e['y1'])
                                if best is None or sc > best[0] or (abs(sc - best[0]) < 1e-6 and y < best[1]):
                                    best = (sc, y)
                        if best:
                            ga_y = best[1]
                            if not group_after_seen:
                                group_after_seen = True
                                group_after_page = p
                    if spec.group_before:
                        best = None
                        for _, e in lines_map.items():
                            lt = ' '.join([t[4] for t in sorted(e['tokens'], key=lambda t: (t[1], t[0]))]).strip()
                            sc = _line_anchor_score(lt, spec.group_before)
                            if sc >= 0.6 and (ga_y is None or float(e['y0']) > ga_y + 0.5):
                                y = float(e['y0'])
                                if best is None or sc > best[0] or (abs(sc - best[0]) < 1e-6 and y < best[1]):
                                    best = (sc, y)
                        if best:
                            gb_y = best[1]
                            page_group_before_y = best[1]
                            if not group_before_seen:
                                group_before_seen = True
                                group_before_page = p
                    # Enforce page-level group_after/group_before bounds
                    if spec.group_after and not group_after_seen:
                        # Haven't seen group_after anywhere yet (including this page); skip searching this page
                        continue
                    if spec.group_before and group_before_seen and group_before_page is not None:
                        # group_before marks the end of the search region; skip pages after it,
                        # but only when the group_before anchor lies on a page *after* the first
                        # search page. This avoids cutting off multi-page sections when the
                        # table-of-contents also contains the group_before heading on page 1.
                        try:
                            first_page = min(pages) if pages else None
                        except Exception:
                            first_page = None
                        if first_page is not None and group_before_page > first_page and p > group_before_page:
                            continue

                    # Build line texts and evaluate
                    for _, entry in sorted(lines_map.items(), key=lambda kv: (kv[1]['y0'], kv[1]['x0'])):
                        row_components: Optional[Dict[str, float]] = None
                        right_text_segment = ''
                        tokens = sorted(entry['tokens'], key=lambda t: (t[1], t[0]))
                        texts = [t[4] for t in tokens]
                        line_text = ' '.join(texts).strip()
                        if not line_text:
                            continue
                        if debug_mode:
                            try:
                                row_y = float(entry.get('y0', 0.0))
                            except Exception:
                                row_y = 0.0
                            print(f"[SMART DEBUG][PDF] row candidate page={p} y={row_y:.1f} tokens={len(tokens)} text={line_text!r}", file=sys.stderr)
                        # Apply group vertical constraints if any
                        # group_after: only rows strictly below the anchor line on the first anchor page;
                        # subsequent pages (after group_after_page) are fully within the region.
                        if ga_y is not None and group_after_page is not None and p == group_after_page:
                            if float(entry['y0']) <= ga_y + 0.5:
                                continue
                        # group_before: only rows strictly above the anchor line on the first group_before page;
                        # pages after group_before_page have already been skipped at page level.
                        if spec.group_before and group_before_page is not None and p == group_before_page:
                            if page_group_before_y is not None and float(entry['y1']) >= page_group_before_y - 0.5:
                                continue
                        raw = line_text if case_sensitive else line_text.lower()
                        needle = row_name if case_sensitive else row_name.lower()
                        score = _fuzzy_ratio(line_text, row_name) if row_name else 0.0
                        anchor_tokens_ok = _anchor_tokens_present(row_name, line_text) if row_name else True
                        min_score = 0.6
                        try:
                            ga = (spec.group_after or "").strip().lower()
                            gb = (spec.group_before or "").strip().lower()
                            if "field value" in ga and "functional acceptance snapshot" in gb:
                                # Document Profile block: allow slightly fuzzier
                                # matches because OCR often splits labels like
                                # "Serial / component" across multiple lines.
                                min_score = 0.45
                        except Exception:
                            pass
                        # containment boost
                        if anchor_tokens_ok and _normalize_anchor_token(row_name) and _normalize_anchor_token(row_name) in _normalize_anchor_token(line_text):
                            score = max(score, 0.99)
                        if row_name and score > 0.6 and anchor_tokens_ok:
                            if not pdf_best_line_only or score > pdf_best_line_only[2]:
                                pdf_best_line_only = (p, line_text, score)
                        if row_name and (score < min_score or not anchor_tokens_ok):
                            if debug_mode:
                                reason = "missing anchor tokens" if not anchor_tokens_ok else "score<0.6"
                                print(f"[SMART DEBUG][PDF] skip row {reason} page={p} score={score:.3f} text={line_text!r}", file=sys.stderr)
                            continue
                        # compute right segment and numeric candidates with coordinates
                        tok_norms = [_normalize_anchor_token(t) for t in texts]
                        anchor_span = _match_anchor_on_line(row_name, texts if case_sensitive else [t.lower() for t in texts], tok_norms)
                        label_right_x = entry['x0']
                        if anchor_span:
                            _, j = anchor_span
                            label_right_x = tokens[j][2]
                        # right-side tokens
                        right_tokens = [t for t in tokens if t[0] >= label_right_x - 1.0]
                        ordered_right_tokens = [t for t in sorted(right_tokens, key=lambda tok: (tok[0], tok[1])) if str(t[4]).strip()]
                        right_text_segment = ' '.join([t[4] for t in ordered_right_tokens]).strip() if ordered_right_tokens else ""
                        smart_kind = _detect_smart_type(spec.smart_snap_type, right_text_segment)
                        if debug_mode:
                            print(f"[SMART DEBUG][PDF] cand page={p} score={score:.3f} smart_kind={smart_kind} row={line_text!r} right={right_text_segment!r}", file=sys.stderr)

                        # Identify nearest header positions above this row
                        header_map: Dict[str, float] = {}
                        row_top = float(entry['y0'])
                        for hdr_name, positions in header_tokens.items():
                            below = [pos for pos in positions if pos[0] < row_top - 0.5]
                            if below:
                                below.sort(key=lambda t: t[0])
                                header_map[hdr_name] = below[-1][2]

                        # Track column anchors from group_after header tokens (e.g., Term/Requirement/Units)
                        column_positions: Dict[str, float] = {}
                        group_after_tokens: List[str] = []
                        if getattr(spec, 'group_after', None):
                            raw_tokens = str(spec.group_after or "").split()
                            group_after_tokens = [_normalize_anchor_token(tok) for tok in raw_tokens if tok.strip()]
                        if group_after_tokens:
                            for e in lines_map.values():
                                if float(e['y1']) >= float(entry['y0']) - 0.5:
                                    continue
                                for tok in e['tokens']:
                                    tok_norm = _normalize_anchor_token(tok[4])
                                    if tok_norm in group_after_tokens and tok_norm not in column_positions:
                                        cx_tok = (float(tok[0]) + float(tok[2])) / 2.0
                                        column_positions[tok_norm] = cx_tok
                                        if debug_mode:
                                            print(f"[SMART DEBUG][PDF] group_after_token match={tok_norm} x={cx_tok:.2f} page={p}", file=sys.stderr)
                        if debug_mode and column_positions:
                            print(f"[SMART DEBUG][PDF] column_positions page={p} {column_positions}", file=sys.stderr)

                        # Build numeric candidates from tokens (captures 500psig etc.)
                        numeric_cands = []  # list of dicts with keys: text, num_clean, units, x0,y0,x1,y1
                        for idx, t in enumerate(ordered_right_tokens):
                            raw = t[4]
                            m = NUMBER_REGEX.search(raw)
                            if not m:
                                continue
                            val_txt = m.group(0)
                            num_clean = numeric_only(val_txt)
                            units_val = extract_units(val_txt)
                            units_val_norm = units_val.lower() if units_val else None
                            unit_neighbor = bool(units_val_norm)
                            if not units_val_norm:
                                lookahead_limit = min(len(ordered_right_tokens), idx + 3)
                                for j in range(idx + 1, lookahead_limit):
                                    nxt = ordered_right_tokens[j]
                                    nxt_txt = str(nxt[4]).strip()
                                    if not nxt_txt:
                                        continue
                                    nxt_norm = nxt_txt.lower()
                                    nxt_cx = (float(nxt[0]) + float(nxt[2])) / 2.0
                                    units_x = column_positions.get('units')
                                    page_x = column_positions.get('page')
                                    if units_hint_set and nxt_norm in units_hint_set:
                                        units_val_norm = nxt_norm
                                        unit_neighbor = True
                                        break
                                    if units_x is not None:
                                        window = max(6.0, abs((page_x or (units_x + 30.0)) - units_x) * 0.2)
                                        if abs(nxt_cx - units_x) <= window:
                                            units_val_norm = nxt_norm
                                            unit_neighbor = True
                                            break
                                    # Stop once the next numeric cell begins to avoid bleeding across columns
                                    if NUMBER_REGEX.search(nxt_txt):
                                        break
                            try:
                                nval = float(num_clean) if num_clean is not None else None
                            except Exception:
                                nval = None
                            numeric_cands.append({
                                'text': val_txt,
                                'raw': raw,
                                'num_clean': num_clean,
                                'nval': nval,
                                'units': units_val_norm,
                                'unit_neighbor': unit_neighbor,
                                'cx': (float(t[0]) + float(t[2])) / 2.0,
                                'x0': t[0], 'y0': t[1], 'x1': t[2], 'y1': t[3],
                            })

                        # Populate line min/max if present
                        line_min_txt = None
                        line_max_txt = None
                        if numeric_cands:
                            vals = [c['nval'] for c in numeric_cands if c['nval'] is not None]
                            if vals:
                                try:
                                    vmin = min(vals)
                                    vmax = max(vals)
                                    for c in numeric_cands:
                                        if c['nval'] == vmin and line_min_txt is None:
                                            line_min_txt = c['text']
                                        if c['nval'] == vmax and line_max_txt is None:
                                            line_max_txt = c['text']
                                except Exception:
                                    pass

                        chosen = None
                        chosen_sec_score = 0.0
                        conflict_reason = None
                        pos_n = spec.smart_position or spec.field_index
                        # Smart Position: \"Nth box\" to the right of the term.
                        # Smart type (number/date/title) only controls how we
                        # interpret that selected box; it should not change
                        # which box is chosen.
                        has_smart_pos = getattr(spec, "smart_position", None) is not None
                        smart_pos_used = False
                        # When the secondary term maps cleanly onto a detected
                        # header token (via group_after), prefer header-based
                        # column targeting so Smart Position refers to the
                        # visual table column, not raw token order. This keeps
                        # OCR behavior aligned with PDF text even when some
                        # cells (e.g., Min='-') are missing tokens.
                        sec_norm = sec_norm_global
                        # Only use header-based column targeting when Smart Position
                        # is NOT configured. When smart_position is set, Smart
                        # Position is authoritative and should not be overridden
                        # by header alignment.
                        use_header_pos = bool(column_positions) and not has_smart_pos
                        column_text_for_pos = None
                        fields_for_pos: List[str] = []
                        if use_header_pos:
                            column_text_for_pos = _column_text_for_position(
                                ordered_right_tokens,
                                column_positions,
                                group_after_tokens,
                                label_right_x,
                                pos_n,
                                sec_term,
                            )
                        elif has_smart_pos:
                            # Fallback: build visual \"boxes\" from the
                            # right-of-label token stream.
                            fields_for_pos = _fields_from_items(ordered_right_tokens)
                        if smart_kind == 'number' and column_text_for_pos:
                            cand_match = NUMBER_REGEX.search(column_text_for_pos)
                            if cand_match:
                                cand_text = cand_match.group(0)
                                try:
                                    nval = float(numeric_only(cand_text)) if numeric_only(cand_text) is not None else None
                                except Exception:
                                    nval = None
                                units_value = extract_units(cand_text) or units_value
                                if nval is not None and (spec.range_min is not None or spec.range_max is not None):
                                    bad = False
                                    if spec.range_min is not None and nval < spec.range_min:
                                        bad = True
                                    if spec.range_max is not None and nval > spec.range_max:
                                        bad = True
                                    if bad and not cand_text.rstrip().endswith('(range violation)'):
                                        cand_text = f"{cand_text} (range violation)"
                                val = _strip_units_from_numeric_text(cand_text) or cand_text
                                if score > best_score:
                                    best_score = score
                                    best_info = (p, line_text, right_text_segment, val, smart_kind, line_min_txt, line_max_txt, None, None)
                                    best_selection_method = "smart_position"
                                    smart_pos_used = True
                                continue
                        elif column_text_for_pos and smart_kind != 'number':
                            # Header-aligned textual field (e.g., KPI status).
                            val = column_text_for_pos
                            if score > best_score:
                                best_score = score
                                best_info = (p, line_text, right_text_segment, val, smart_kind, line_min_txt, line_max_txt, None, None)
                                best_selection_method = "smart_position"
                                smart_pos_used = True
                            continue
                        elif smart_kind != 'number' and has_smart_pos and pos_n and pos_n >= 1 and fields_for_pos:
                            # Smart Position for non-numeric snaps (e.g., pick the Nth
                            # status/text field to the right of the label).
                            if pos_n <= len(fields_for_pos):
                                field_text = fields_for_pos[pos_n - 1]
                                val = field_text.strip()
                                if val and score > best_score:
                                    best_score = score
                                    best_info = (p, line_text, right_text_segment, val, smart_kind, line_min_txt, line_max_txt, None, None)
                                    best_selection_method = "smart_position"
                                    smart_pos_used = True
                                continue
                        if smart_kind == 'number' and pos_n and pos_n >= 1:
                            # Smart Position: select the Nth \"box\" (field) to the
                            # right of the term, then interpret it according to
                            # smart_snap_type. Numeric/date parsing happens
                            # *after* the field is selected.
                            if has_smart_pos and fields_for_pos:
                                if pos_n <= len(fields_for_pos):
                                    field_text = fields_for_pos[pos_n - 1]
                                    cand_text = field_text
                                    cand_match = NUMBER_REGEX.search(field_text)
                                    nval = None
                                    if cand_match:
                                        cand_text = cand_match.group(0)
                                        try:
                                            nval = float(numeric_only(cand_text)) if numeric_only(cand_text) is not None else None
                                        except Exception:
                                            nval = None
                                        units_value = extract_units(cand_text) or units_value
                                        if nval is not None and (spec.range_min is not None or spec.range_max is not None):
                                            bad = False
                                            if spec.range_min is not None and nval < spec.range_min:
                                                bad = True
                                            if spec.range_max is not None and nval > spec.range_max:
                                                bad = True
                                            if bad and not cand_text.rstrip().endswith('(range violation)'):
                                                cand_text = f"{cand_text} (range violation)"
                                    # Only accept compatible numeric values for Smart Position;
                                    # if no numeric content is present, fall back to scoring logic below.
                                    if cand_match:
                                        val = _strip_units_from_numeric_text(cand_text) or cand_text
                                        if score > best_score:
                                            best_score = score
                                            best_info = (p, line_text, right_text_segment, val, smart_kind, line_min_txt, line_max_txt, None, None)
                                            best_selection_method = "smart_position"
                                            smart_pos_used = True
                                        continue
                            elif (not has_smart_pos) and ordered_right_tokens:
                                # Legacy: field_index selects the Nth token field.
                                if pos_n <= len(ordered_right_tokens):
                                    tok = ordered_right_tokens[pos_n - 1]
                                    raw_field = str(tok[4]).strip()
                                    cand_match = NUMBER_REGEX.search(raw_field)
                                    cand_text = cand_match.group(0) if cand_match else raw_field
                                    nval = None
                                    if cand_match:
                                        try:
                                            nval = float(numeric_only(cand_text)) if numeric_only(cand_text) is not None else None
                                        except Exception:
                                            nval = None
                                        units_value = extract_units(cand_text) or units_value
                                        if nval is not None and (spec.range_min is not None or spec.range_max is not None):
                                            bad = False
                                            if spec.range_min is not None and nval < spec.range_min:
                                                bad = True
                                            if spec.range_max is not None and nval > spec.range_max:
                                                bad = True
                                            if bad and not cand_text.rstrip().endswith('(range violation)'):
                                                cand_text = f"{cand_text} (range violation)"
                                        val = _strip_units_from_numeric_text(cand_text) or cand_text
                                    else:
                                        val = raw_field
                                    if score > best_score:
                                        best_score = score
                                        best_info = (p, line_text, right_text_segment, val, smart_kind, line_min_txt, line_max_txt, None, None)
                                    continue
                        unitful_candidates = [c for c in numeric_cands if c.get('unit_neighbor')]
                        if unitful_candidates:
                            numeric_cands = unitful_candidates
                        has_units_match = bool(unitful_candidates)
                        if smart_kind == 'number' and numeric_cands and not has_smart_pos:
                            # Score candidates using middle-of-line (between min/max), units hints, range, and secondary-term header alignment.
                            # Secondary vertical sweep is ignored; only header alignment contributes.
                            def sec_score(c):
                                return 0.0

                            # Optional: header-based alignment across candidates in this row.
                            sec_header_x0 = None
                            if sec_term and sec_norm_global:
                                try:
                                    row_top = float(entry.get('y0', 0.0))
                                except Exception:
                                    row_top = float(entry['y0'])
                                header_candidates: List[Tuple[float, float, float]] = []
                                for _, ent2 in lines_map.items():
                                    try:
                                        ent2_bottom = float(ent2.get('y1', 0.0))
                                    except Exception:
                                        ent2_bottom = float(ent2['y1'])
                                    # Only consider headers that are visually above this row
                                    if ent2_bottom >= row_top - 0.5:
                                        continue
                                    for ot in ent2['tokens']:
                                        raw_txt = str(ot[4] or "")
                                        if not raw_txt.strip():
                                            continue
                                        sc = _fuzzy_ratio(raw_txt, sec_term)
                                        ot_norm = _normalize_anchor_token(raw_txt)
                                        if sec_norm_global and sec_norm_global in ot_norm:
                                            sc = max(sc, 0.99)
                                        if sc < 0.6:
                                            continue
                                        cy_tok = (float(ot[1]) + float(ot[3])) / 2.0
                                        header_candidates.append((sc, cy_tok, float(ot[0])))
                                if header_candidates:
                                    header_candidates.sort(key=lambda t: (-t[0], abs(t[1] - row_top)))
                                    sec_header_x0 = header_candidates[0][2]

                            header_alignment: Dict[int, float] = {}
                            if sec_header_x0 is not None and numeric_cands:
                                dists: List[float] = []
                                for c in numeric_cands:
                                    try:
                                        d = abs(float(c['x0']) - float(sec_header_x0))
                                    except Exception:
                                        d = abs(c['x0'] - sec_header_x0)  # type: ignore[operator]
                                    dists.append(d)
                                if dists:
                                    d_min = min(dists)
                                    d_max = max(dists)
                                    span = max(d_max - d_min, 1e-6)
                                    for c, d in zip(numeric_cands, dists):
                                        if d_max == d_min:
                                            h = 1.0
                                        else:
                                            h = max(0.0, 1.0 - (d - d_min) / span)
                                        header_alignment[id(c)] = h

                            scored = []
                            score_components: Dict[int, Dict[str, float]] = {}
                            for c in numeric_cands:
                                s = 0.0
                                comp: Dict[str, float] = {
                                    "between_min_max": 0.0,
                                    "units_hint": 0.0,
                                    "range": 0.0,
                                    "secondary_vertical": 0.0,
                                    "secondary_header": 0.0,
                                    "value_header_align": 0.0,
                                    "min_header_penalty": 0.0,
                                    "max_header_penalty": 0.0,
                                    "label_dx": 0.0,
                                }
                                # Prefer middle value between line min and max
                                in_middle = False
                                if line_min_txt is not None and line_max_txt is not None and c['nval'] is not None:
                                    try:
                                        line_min_val = float(numeric_only(line_min_txt)) if line_min_txt is not None else None
                                        line_max_val = float(numeric_only(line_max_txt)) if line_max_txt is not None else None
                                    except Exception:
                                        line_min_val = line_max_val = None
                                    if line_min_val is not None and line_max_val is not None and line_min_val < c['nval'] < line_max_val:
                                        in_middle = True
                                        s += 2.0
                                        comp["between_min_max"] += 2.0
                                    elif line_min_val is not None and c['nval'] == line_min_val:
                                        s -= 0.2
                                        comp["between_min_max"] -= 0.2
                                    elif line_max_val is not None and c['nval'] == line_max_val:
                                        s -= 0.2
                                        comp["between_min_max"] -= 0.2
                                if units_hint_set:
                                    if c.get('units') in units_hint_set:
                                        s += 0.4
                                        comp["units_hint"] += 0.4
                                    elif has_units_match:
                                        s -= 0.1
                                        comp["units_hint"] -= 0.1
                                if c['nval'] is not None and (spec.range_min is not None and spec.range_max is not None):
                                    if spec.range_min <= c['nval'] <= spec.range_max:
                                        s += 0.4
                                        comp["range"] += 0.4
                                elif c['nval'] is not None:
                                    # soft preference toward range proximity if only one bound given
                                    if spec.range_min is not None and c['nval'] >= spec.range_min:
                                        s += 0.1
                                        comp["range"] += 0.1
                                    if spec.range_max is not None and c['nval'] <= spec.range_max:
                                        s += 0.1
                                        comp["range"] += 0.1
                                sec_s = sec_score(c)
                                # vertical secondary term score ignored; header alignment only
                                # Header-based secondary term alignment (best candidate gets strongest boost)
                                hdr_align = header_alignment.get(id(c))
                                if hdr_align is not None:
                                    s += _SEC_HEADER_WEIGHT * hdr_align
                                    comp["secondary_header"] += _SEC_HEADER_WEIGHT * hdr_align
                                cx_cand = (c['x0'] + c['x1']) / 2.0
                                if 'value' in header_map:
                                    dist = abs(cx_cand - header_map['value'])
                                    delta = 0.8 * (1.0 / (1.0 + dist / 18.0))
                                    s += delta
                                    comp["value_header_align"] += delta
                                if 'min' in header_map:
                                    dist = abs(cx_cand - header_map['min'])
                                    delta = 0.4 * (1.0 / (1.0 + dist / 18.0))
                                    s -= delta
                                    comp["min_header_penalty"] -= delta
                                if 'max' in header_map:
                                    dist = abs(cx_cand - header_map['max'])
                                    delta = 0.4 * (1.0 / (1.0 + dist / 18.0))
                                    s -= delta
                                    comp["max_header_penalty"] -= delta
                                # slight preference for smaller horizontal distance from label
                                dx = max(0.0, c['x0'] - label_right_x)
                                delta_dx = 0.05 * (1.0 / (1.0 + dx/10.0))
                                s += delta_dx
                                comp["label_dx"] += delta_dx
                                if debug_mode:
                                    print(f"[SMART DEBUG][PDF] cand_score page={p} val={c['text']} units={c.get('units')} s={s:.3f}", file=sys.stderr)
                                comp["total"] = s
                                combined_sec = header_alignment.get(id(c), 0.0)
                                scored.append((s, c, combined_sec))
                                score_components[id(c)] = comp
                            scored.sort(key=lambda t: t[0], reverse=True)
                            if scored:
                                candidate_pool = scored
                                if units_hint_set and has_units_match:
                                    prioritized = [t for t in candidate_pool if t[1].get('units') in units_hint_set]
                                    if prioritized:
                                        candidate_pool = prioritized
                                top_score = candidate_pool[0][0]
                                top = [t for t in candidate_pool if t[0] >= top_score - 0.1]
                                if len(top) > 1:
                                    conflict_reason = 'multiple candidates with similar scores'
                                chosen = top[0][1]
                                chosen_sec_score = top[0][2]
                                row_components = score_components.get(id(chosen))
                                units_value = chosen.get('units') or units_value
                                # Build output value text with possible range violation annotation
                                cand = chosen['text']
                                if chosen['nval'] is not None and (spec.range_min is not None or spec.range_max is not None):
                                    bad = False
                                    if spec.range_min is not None and chosen['nval'] < spec.range_min:
                                        bad = True
                                    if spec.range_max is not None and chosen['nval'] > spec.range_max:
                                        bad = True
                                    if bad and not cand.rstrip().endswith('(range violation)'):
                                        cand = f"{cand} (range violation)"
                                val = cand
                                if smart_kind == 'number':
                                    val = _strip_units_from_numeric_text(val) or val
                                if score > best_score:
                                    best_score = score
                                    best_info = (p, line_text, right_text_segment, val, smart_kind, line_min_txt, line_max_txt, conflict_reason, (chosen_sec_score if sec_term else None))
                                    best_components = row_components
                                    if debug_mode:
                                        print(f"[SMART DEBUG][PDF] best_update page={p} score={score:.3f} val={val!r}", file=sys.stderr)
                        else:
                            # string/date/time handling via original helper; for numeric snaps only when no Smart Position is configured
                            val = None
                            if smart_kind != 'number' or not has_smart_pos:
                                val = extract_from_line(line_text, right_text_segment, smart_kind)
                                if val and smart_kind == 'number':
                                    val = _strip_units_from_numeric_text(val) or val
                            if val:
                                if score > best_score:
                                    best_score = score
                                    best_info = (p, line_text, right_text_segment, val, smart_kind, line_min_txt, line_max_txt, conflict_reason, None)
                                    best_components = row_components
                                    if debug_mode:
                                        print(f"[SMART DEBUG][PDF] best_update(direct) page={p} score={score:.3f} val={val!r}", file=sys.stderr)
            finally:
                try:
                    doc.close()
                except Exception:
                    pass
            if best_info:
                page_hit, context_line_text, right_text, value_text, smart_kind, line_min_txt, line_max_txt, conflict_reason, sec_found = best_info
                # For title/text smart snaps, strip label tokens and normalize
                # common status values so we return just the field contents
                # (e.g., 'NO' instead of 'NO +88*C').
                if smart_kind == 'title':
                    value_text = _strip_label_tokens(value_text, row_name)
                    value_text = _extract_status_from_title(value_text)
                row_text_selected = context_line_text
                confidence = best_score
                method_used = 'smart:pdf'
                text_source = 'pdf'
                sel_method = "smart_position" if getattr(spec, "smart_position", None) is not None else "smart_score"
                return MatchResult(
                    pdf_file=pdf_path.name,
                    serial_number=serial_number,
                    term=spec.term,
                    page=page_hit,
                    number=value_text,
                    units=units_value,
                    context=right_text,
                    method=method_used,
                    found=True,
                    confidence=confidence,
                    row_label=row_text_selected,
                    column_label=None,
                    text_source=text_source,
                    smart_snap_context=context_line_text,
                    smart_snap_type=smart_kind,
                    smart_line_min=line_min_txt,
                    smart_line_max=line_max_txt,
                    smart_conflict=conflict_reason,
                    smart_secondary_found=sec_found,
                    smart_score_breakdown=best_components,
                    smart_selection_method=sel_method,
                )
            if best_info is None and pdf_best_line_only is not None:
                p, context_line_text, sc = pdf_best_line_only
                return MatchResult(
                    pdf_file=pdf_path.name,
                    serial_number=serial_number,
                    term=spec.term,
                    page=p,
                    number=None,
                    units=None,
                    context="",
                    method='smart:pdf',
                    found=False,
                    confidence=sc,
                    row_label=context_line_text,
                    column_label=None,
                    text_source='pdf',
                    error_reason="Smart snap: matched row, no value",
                    smart_snap_context=context_line_text,
                    smart_snap_type=spec.smart_snap_type or 'auto',
                )

    # OCR fallback with EasyOCR boxes -> lines
    if _HAVE_EASYOCR and _HAVE_PYMUPDF:
        try:
            dpi_candidates: List[int] = []

            def _push_dpi(val: Optional[str]) -> None:
                if not val:
                    return
                try:
                    dpi_val = int(str(val).strip())
                    if dpi_val > 0 and dpi_val not in dpi_candidates:
                        dpi_candidates.append(dpi_val)
                except Exception:
                    pass

            raw_list = os.environ.get('SMART_DPI_LIST')
            if raw_list:
                for token in re.split(r"[;,]", raw_list):
                    _push_dpi(token)

            _push_dpi(os.environ.get('SMART_DPI_BASE'))
            _push_dpi(os.environ.get('OCR_DPI'))
            _push_dpi('700')

            if not dpi_candidates:
                dpi_candidates = [700]
            if debug_mode:
                print(f"[SMART DEBUG] dpi_candidates={dpi_candidates}", file=sys.stderr)

            langs_raw = (os.environ.get('EASYOCR_LANGS') or os.environ.get('OCR_LANGS') or 'en')
            langs = [s.strip() for s in re.split(r'[;,]', langs_raw) if s.strip()]
            # Open doc for page bounds
            try:
                doc = fitz.open(str(pdf_path))  # type: ignore[name-defined]
            except Exception:
                doc = None
            if not pages and doc is not None:
                pages = list(range(1, doc.page_count + 1))
            best_score = 0.0
            best_info = None
            best_components = None
            # Track first occurrences of group_after/group_before across pages for OCR path
            group_after_seen = spec.group_after is None
            group_after_page: Optional[int] = None
            group_before_seen = spec.group_before is None
            group_before_page: Optional[int] = None
            # Vertical tolerance (in OCR pixel coordinates) for grouping
            # EasyOCR boxes into logical text rows. Exposed via OCR_ROW_EPS
            # so the UI can provide a slider; default tuned for 10–14pt text.
            try:
                row_eps = float(os.environ.get("OCR_ROW_EPS", "8.0"))
            except Exception:
                row_eps = 8.0
            row_eps = max(0.5, min(50.0, row_eps))
            for dpi in dpi_candidates:
                for p in pages or []:
                    items = _get_easyocr_boxes_page(pdf_path, p, dpi=dpi, langs=langs)
                    if not items:
                        if debug_mode:
                            print(f"[SMART DEBUG] no OCR items dpi={dpi} page={p}", file=sys.stderr)
                        continue
                    # Group OCR boxes into line rows using configurable
                    # vertical tolerance so that all tokens from a visual
                    # line share the same row bucket.
                    rows: Dict[int, List[Dict[str, float]]] = {}
                    prev_cy: Optional[float] = None
                    current_key: Optional[int] = None
                    for it in sorted(items, key=lambda d: float(d.get("cy", 0.0))):
                        cy_val = float(it.get("cy", 0.0))
                        if prev_cy is None or abs(cy_val - prev_cy) > row_eps or current_key is None:
                            key = int(round(cy_val))
                            rows[key] = [it]
                            current_key = key
                        else:
                            rows[current_key].append(it)  # type: ignore[index]
                        prev_cy = cy_val
                    group_after_tokens: List[str] = []
                    if getattr(spec, 'group_after', None):
                        raw_tokens = str(spec.group_after or "").split()
                        group_after_tokens = [_normalize_anchor_token(tok) for tok in raw_tokens if tok.strip()]
                    # detect optional group bounds using anchors
                    group_anchor_y = None
                    group_upper_y = None
                    page_group_before_y = None
                    if spec.group_after:
                        matches = []
                        for it in items:
                            txt = str(it.get('text') or '')
                            sc = _fuzzy_ratio(txt, spec.group_after)
                            if _normalize_anchor_token(spec.group_after) and _normalize_anchor_token(spec.group_after) in _normalize_anchor_token(txt):
                                sc = max(sc, 0.99)
                            if sc >= 0.6:
                                matches.append((it, sc))
                        if matches:
                            anchor_best = max(m[1] for m in matches)
                            top = [m for m in matches if m[1] >= anchor_best - 0.1]
                            group_anchor_y = min(top, key=lambda t: float(t[0].get('cy',0.0)))[0].get('cy', None)
                            if not group_after_seen:
                                group_after_seen = True
                                group_after_page = p
                    if spec.group_before:
                        matches = []
                        for it in items:
                            txt = str(it.get('text') or '')
                            sc = _fuzzy_ratio(txt, spec.group_before)
                            if _normalize_anchor_token(spec.group_before) and _normalize_anchor_token(spec.group_before) in _normalize_anchor_token(txt):
                                sc = max(sc, 0.99)
                            if sc >= 0.6 and (group_anchor_y is None or float(it.get('cy',0.0)) > float(group_anchor_y) + 1.0):
                                matches.append((it, sc))
                        if matches:
                            upper_best = max(m[1] for m in matches)
                            top = [m for m in matches if m[1] >= upper_best - 0.1]
                            found_y = min(top, key=lambda t: float(t[0].get('cy',0.0)))[0].get('cy', None)
                            group_upper_y = found_y
                            page_group_before_y = found_y
                            if not group_before_seen and found_y is not None:
                                group_before_seen = True
                                group_before_page = p
                    # Enforce page-level group_after/group_before bounds
                    if spec.group_after and not group_after_seen:
                        # Haven't seen group_after anywhere yet (including this page); skip searching this page
                        continue
                    if spec.group_before and group_before_seen and group_before_page is not None:
                        # group_before marks the end of the search region; skip pages after it,
                        # but only when the group_before anchor lies on a page *after* the first
                        # search page. This avoids cutting off multi-page sections when a
                        # table-of-contents also contains the group_before heading on page 1.
                        try:
                            first_page = min(pages) if pages else None
                        except Exception:
                            first_page = None
                        if first_page is not None and group_before_page > first_page and p > group_before_page:
                            continue
                    for cy, row_items in sorted(rows.items(), key=lambda kv: kv[0]):
                        row_components: Optional[Dict[str, float]] = None
                        right_text_segment = ''
                        if debug_mode:
                            print(f"[SMART DEBUG] row candidate dpi={dpi} page={p} cy={cy} tokens={len(row_items)} text={' '.join(str(it.get('text') or '') for it in row_items)!r}", file=sys.stderr)
                        row_items.sort(key=lambda t: (t.get('y0',0.0), t.get('x0',0.0)))
                        texts = [str(it.get('text') or '') for it in row_items]
                        line_text = ' '.join(texts).strip()
                        if not line_text:
                            continue
                        # group filters
                        mean_y = sum(float(it.get('cy',0.0)) for it in row_items) / max(1,len(row_items))
                        # group_after: only rows strictly below the anchor line on the first anchor page;
                        # subsequent pages (after group_after_page) are fully within the region.
                        if group_anchor_y is not None and group_after_page is not None and p == group_after_page:
                            if mean_y <= float(group_anchor_y) + 0.5:
                                continue
                        # group_before: only rows strictly above the anchor line on the first group_before page;
                        # pages after group_before_page have already been skipped at page level.
                        if spec.group_before and group_before_page is not None and p == group_before_page:
                            if page_group_before_y is not None and mean_y >= float(page_group_before_y) - 0.5:
                                continue
                        score = _fuzzy_ratio(line_text, row_name) if row_name else 0.0
                        anchor_tokens_ok = _anchor_tokens_present(row_name, line_text) if row_name else True
                        if anchor_tokens_ok and _normalize_anchor_token(row_name) and _normalize_anchor_token(row_name) in _normalize_anchor_token(line_text):
                            score = max(score, 0.99)
                        if row_name and (score < 0.6 or not anchor_tokens_ok):
                            if debug_mode:
                                print(f"[SMART DEBUG] skip row score<0.6 dpi={dpi} page={p} score={score:.3f} text={line_text!r}", file=sys.stderr)
                            continue
                        # right-of approx via anchor match in token stream
                        tok_norms = [_normalize_anchor_token(t) for t in texts]
                        span = _match_anchor_on_line(row_name, texts if case_sensitive else [t.lower() for t in texts], tok_norms)
                        label_right_x = min((float(it.get('x0',0.0)) for it in row_items), default=0.0)
                        if span:
                            _, j = span
                            try:
                                label_right_x = float(row_items[j].get('x1', row_items[j].get('cx', 0.0)))
                            except Exception:
                                pass
                        right_items = [it for it in row_items if float(it.get('x0',0.0)) >= label_right_x - 1.0]
                        row_min_y = min((float(it.get('y0',0.0)) for it in row_items), default=0.0)
                        row_max_y = max((float(it.get('y1',0.0)) for it in row_items), default=0.0)
                        row_height = max(1.0, row_max_y - row_min_y)
                        need_augment = double_height_mode or not any(NUMBER_REGEX.search(str(it.get('text') or '')) for it in right_items)
                        if need_augment:
                            if double_height_mode:
                                top_pad = max(4.0, row_height * 0.8)
                                bottom_pad = max(8.0, row_height * 1.6)
                            else:
                                top_pad = max(1.5, min(5.0, row_height * 0.6))
                                bottom_pad = max(3.0, min(8.0, row_height * 0.9))
                            augmented: List[Dict[str, float]] = []
                            seen_ids = {id(it) for it in right_items}
                            for it in items:
                                if id(it) in seen_ids:
                                    continue
                                if float(it.get('x0',0.0)) < label_right_x - 1.5:
                                    continue
                                y0 = float(it.get('y0',0.0))
                                y1 = float(it.get('y1',0.0))
                                if y1 < row_min_y - top_pad or y0 > row_max_y + bottom_pad:
                                    continue
                                augmented.append(it)
                                seen_ids.add(id(it))
                            if augmented:
                                right_items = sorted(right_items + augmented, key=lambda t: (float(t.get('y0',0.0)), float(t.get('x0',0.0))))
                                if debug_mode:
                                    print(f"[SMART DEBUG] expanded right_items via band tolerance ({len(augmented)} extra)", file=sys.stderr)
                        ordered_right_items = [it for it in sorted(right_items, key=lambda t: (float(t.get('x0',0.0)), float(t.get('y0',0.0)))) if str(it.get('text') or '').strip()]
                        column_positions: Dict[str, float] = {}
                        if group_after_tokens:
                            for ent in rows.values():
                                if not ent:
                                    continue
                                ent_bottom = max((float(tok.get('y1', 0.0)) for tok in ent), default=0.0)
                                if ent_bottom >= row_min_y - 0.5:
                                    continue
                                for tok in ent:
                                    tok_norm = _normalize_anchor_token(str(tok.get('text') or ''))
                                    if tok_norm in group_after_tokens and tok_norm not in column_positions:
                                        cx_tok = (float(tok.get('x0', 0.0)) + float(tok.get('x1', 0.0))) / 2.0
                                        column_positions[tok_norm] = cx_tok
                        right_text_segment = ' '.join([str(it.get('text') or '') for it in right_items]).strip() if right_items else ""
                        smart_kind = _detect_smart_type(spec.smart_snap_type, right_text_segment)
                        if debug_mode:
                            print(f"[SMART DEBUG] cand dpi={dpi} page={p} score={score:.3f} smart_kind={smart_kind} row={line_text!r} right={right_text_segment!r}", file=sys.stderr)

                        # Identify nearby column headers for scoring (prefer 'value' column)
                        header_map: Dict[str, Dict[str, float]] = {}
                        for hdr_name in ('value', 'min', 'max'):
                            matches = [
                                it for it in items
                                if str(it.get('text') or '').strip().lower() == hdr_name
                                and float(it.get('cy', 0.0)) < mean_y
                                and (group_anchor_y is None or float(it.get('cy',0.0)) >= float(group_anchor_y) - 5.0)
                            ]
                            if matches:
                                header_map[hdr_name] = max(matches, key=lambda it: float(it.get('cy', 0.0)))

                        # Build numeric candidates from OCR tokens
                        numeric_cands = []
                        for it in right_items:
                            raw = str(it.get('text') or '')
                            m = NUMBER_REGEX.search(raw)
                            if not m:
                                continue
                            val_txt = m.group(0)
                            num_clean = numeric_only(val_txt)
                            try:
                                nval = float(num_clean) if num_clean is not None else None
                            except Exception:
                                nval = None
                            units_val = extract_units(val_txt)
                            numeric_cands.append({
                                'text': val_txt,
                                'raw': raw,
                                'nval': nval,
                                'units': units_val.lower() if units_val else None,
                                'x0': float(it.get('x0',0.0)), 'y0': float(it.get('y0',0.0)), 'x1': float(it.get('x1',0.0)), 'y1': float(it.get('y1',0.0)),
                            })

                        line_min_txt = None
                        line_max_txt = None
                        if numeric_cands:
                            vals = [c['nval'] for c in numeric_cands if c['nval'] is not None]
                            if vals:
                                try:
                                    vmin = min(vals)
                                    vmax = max(vals)
                                    for c in numeric_cands:
                                        if c['nval'] == vmin and line_min_txt is None:
                                            line_min_txt = c['text']
                                        if c['nval'] == vmax and line_max_txt is None:
                                            line_max_txt = c['text']
                                except Exception:
                                    pass

                        conflict_reason = None
                        chosen_sec_score = 0.0
                        pos_n = spec.smart_position or spec.field_index
                        has_smart_pos = getattr(spec, "smart_position", None) is not None
                        smart_pos_used = False
                        # Prefer header-based column targeting when the
                        # secondary term maps cleanly onto a detected header
                        # token so that Smart Position aligns with the visual
                        # table column even if some cells (e.g., Min='-') are
                        # missing OCR tokens.
                        sec_norm = sec_norm_global
                        # Only use header-based column targeting when Smart Position
                        # is NOT configured. When smart_position is set, Smart
                        # Position is authoritative and should not be overridden
                        # by header alignment.
                        use_header_pos = bool(column_positions) and not has_smart_pos
                        column_text_for_pos = None
                        fields_for_pos: List[str] = []
                        if use_header_pos:
                            column_text_for_pos = _column_text_for_position(
                                ordered_right_items,
                                column_positions,
                                group_after_tokens,
                                label_right_x,
                                pos_n,
                                sec_term,
                            )
                        elif has_smart_pos:
                            # Smart Position: treat as Nth \"box\" to the right
                            # of the term. For OCR, boxes are EasyOCR tokens.
                            fields_for_pos = _fields_from_items(ordered_right_items)
                        if smart_kind == 'number' and column_text_for_pos:
                            cand_match = NUMBER_REGEX.search(column_text_for_pos)
                            if cand_match:
                                cand_text = cand_match.group(0)
                                try:
                                    nval = float(numeric_only(cand_text)) if numeric_only(cand_text) is not None else None
                                except Exception:
                                    nval = None
                                units_value = extract_units(cand_text) or units_value
                                if nval is not None and (spec.range_min is not None or spec.range_max is not None):
                                    bad = False
                                    if spec.range_min is not None and nval < spec.range_min:
                                        bad = True
                                    if spec.range_max is not None and nval > spec.range_max:
                                        bad = True
                                    if bad and not cand_text.rstrip().endswith('(range violation)'):
                                        cand_text = f"{cand_text} (range violation)"
                                val = _strip_units_from_numeric_text(cand_text) or cand_text
                                if score > best_score:
                                    best_score = score
                                    best_info = (p, line_text, right_text_segment, val, smart_kind, line_min_txt, line_max_txt, None, None)
                                    smart_pos_used = True
                                continue
                        elif column_text_for_pos and smart_kind != 'number':
                            val = column_text_for_pos
                            if score > best_score:
                                best_score = score
                                best_info = (p, line_text, right_text_segment, val, smart_kind, line_min_txt, line_max_txt, None, None)
                                smart_pos_used = True
                            continue
                        elif smart_kind != 'number' and has_smart_pos and pos_n and pos_n >= 1 and fields_for_pos:
                            if pos_n <= len(fields_for_pos):
                                field_text = fields_for_pos[pos_n - 1]
                                val = field_text.strip()
                                if val and score > best_score:
                                    best_score = score
                                    best_info = (p, line_text, right_text_segment, val, smart_kind, line_min_txt, line_max_txt, None, None)
                                    smart_pos_used = True
                                continue
                        if smart_kind == 'number' and pos_n and pos_n >= 1 and ordered_right_items:
                            if has_smart_pos and fields_for_pos:
                                if pos_n <= len(fields_for_pos):
                                    field_text = fields_for_pos[pos_n - 1]
                                    cand_match = NUMBER_REGEX.search(field_text)
                                    if not cand_match:
                                        # Smart Position box has no numeric content; log and fall back to scoring logic below.
                                        if debug_mode:
                                            print(f"[SMART DEBUG] smart_position box non-numeric dpi={dpi} page={p} pos={pos_n} field={field_text!r}", file=sys.stderr)
                                    else:
                                        cand_text = cand_match.group(0)
                                        nval = None
                                        try:
                                            nval = float(numeric_only(cand_text)) if numeric_only(cand_text) is not None else None
                                        except Exception:
                                            nval = None
                                        units_value = extract_units(cand_text) or units_value
                                        if nval is not None and (spec.range_min is not None or spec.range_max is not None):
                                            bad = False
                                            if spec.range_min is not None and nval < spec.range_min:
                                                bad = True
                                            if spec.range_max is not None and nval > spec.range_max:
                                                bad = True
                                            if bad and not cand_text.rstrip().endswith('(range violation)'):
                                                cand_text = f"{cand_text} (range violation)"
                                        val = _strip_units_from_numeric_text(cand_text) or cand_text
                                        if score > best_score:
                                            best_score = score
                                            best_info = (p, line_text, right_text_segment, val, smart_kind, line_min_txt, line_max_txt, None, None)
                                            smart_pos_used = True
                                        continue
                            elif not has_smart_pos:
                                if pos_n <= len(ordered_right_items):
                                    tok = ordered_right_items[pos_n - 1]
                                    raw_field = str(tok.get('text') or '').strip()
                                    cand_match = NUMBER_REGEX.search(raw_field)
                                    cand_text = cand_match.group(0) if cand_match else raw_field
                                    nval = None
                                    if cand_match:
                                        try:
                                            nval = float(numeric_only(cand_text)) if numeric_only(cand_text) is not None else None
                                        except Exception:
                                            nval = None
                                        units_value = extract_units(cand_text) or units_value
                                        if nval is not None and (spec.range_min is not None or spec.range_max is not None):
                                            bad = False
                                            if spec.range_min is not None and nval < spec.range_min:
                                                bad = True
                                            if spec.range_max is not None and nval > spec.range_max:
                                                bad = True
                                            if bad and not cand_text.rstrip().endswith('(range violation)'):
                                                cand_text = f"{cand_text} (range violation)"
                                        val = _strip_units_from_numeric_text(cand_text) or cand_text
                                    else:
                                        val = raw_field
                                    if score > best_score:
                                        best_score = score
                                        best_info = (p, line_text, right_text_segment, val, smart_kind, line_min_txt, line_max_txt, None, None)
                                    continue
                        if smart_kind == 'number' and numeric_cands and not has_smart_pos:
                            # Score candidates using middle-of-line (between min/max), units hints, range, and secondary-term header alignment.
                            # Secondary vertical sweep is ignored; only header alignment contributes.
                            def sec_score(c):
                                return 0.0

                            # Optional secondary header alignment across this row's candidates
                            sec_header_x0 = None
                            if sec_term and sec_norm_global:
                                header_candidates: List[Tuple[float, float, float]] = []
                                for it in items:
                                    txt = str(it.get('text') or '')
                                    if not txt.strip():
                                        continue
                                    cy_tok = float(it.get('cy', 0.0)) if it.get('cy', None) is not None else (float(it.get('y0', 0.0)) + float(it.get('y1', 0.0))) / 2.0
                                    # Only consider tokens visually above this row
                                    if cy_tok >= row_min_y - 0.5:
                                        continue
                                    sc = _fuzzy_ratio(txt, sec_term)
                                    tok_norm = _normalize_anchor_token(txt)
                                    if sec_norm_global and sec_norm_global in tok_norm:
                                        sc = max(sc, 0.99)
                                    if sc < 0.6:
                                        continue
                                    header_candidates.append((sc, cy_tok, float(it.get('x0', 0.0))))
                                if header_candidates:
                                    header_candidates.sort(key=lambda t: (-t[0], abs(t[1] - row_min_y)))
                                    sec_header_x0 = header_candidates[0][2]

                            header_alignment: Dict[int, float] = {}
                            if sec_header_x0 is not None and numeric_cands:
                                dists: List[float] = []
                                for c in numeric_cands:
                                    dists.append(abs(float(c['x0']) - float(sec_header_x0)))
                                if dists:
                                    d_min = min(dists)
                                    d_max = max(dists)
                                    span = max(d_max - d_min, 1e-6)
                                    for c, d in zip(numeric_cands, dists):
                                        if d_max == d_min:
                                            h = 1.0
                                        else:
                                            h = max(0.0, 1.0 - (d - d_min) / span)
                                        header_alignment[id(c)] = h

                            scored = []
                            score_components: Dict[int, Dict[str, float]] = {}
                            for c in numeric_cands:
                                s = 0.0
                                comp: Dict[str, float] = {
                                    "between_min_max": 0.0,
                                    "units_hint": 0.0,
                                    "range": 0.0,
                                    "secondary_vertical": 0.0,
                                    "secondary_header": 0.0,
                                    "value_header_align": 0.0,
                                    "min_header_penalty": 0.0,
                                    "max_header_penalty": 0.0,
                                    "label_dx": 0.0,
                                }
                                # Prefer middle value between line min and max
                                if line_min_txt is not None and line_max_txt is not None and c['nval'] is not None:
                                    try:
                                        line_min_val = float(numeric_only(line_min_txt)) if line_min_txt is not None else None
                                        line_max_val = float(numeric_only(line_max_txt)) if line_max_txt is not None else None
                                    except Exception:
                                        line_min_val = line_max_val = None
                                    if line_min_val is not None and line_max_val is not None and line_min_val < c['nval'] < line_max_val:
                                        s += 2.0
                                        comp["between_min_max"] += 2.0
                                    elif line_min_val is not None and c['nval'] == line_min_val:
                                        s -= 0.2
                                        comp["between_min_max"] -= 0.2
                                    elif line_max_val is not None and c['nval'] == line_max_val:
                                        s -= 0.2
                                        comp["between_min_max"] -= 0.2
                                if units_hints and c['units'] in units_hints:
                                    s += 0.4
                                    comp["units_hint"] += 0.4
                                if c['nval'] is not None and (spec.range_min is not None and spec.range_max is not None):
                                    if spec.range_min <= c['nval'] <= spec.range_max:
                                        s += 0.4
                                        comp["range"] += 0.4
                                elif c['nval'] is not None:
                                    if spec.range_min is not None and c['nval'] >= spec.range_min:
                                        s += 0.1
                                        comp["range"] += 0.1
                                    if spec.range_max is not None and c['nval'] <= spec.range_max:
                                        s += 0.1
                                        comp["range"] += 0.1
                                sec_s = sec_score(c)
                                # vertical secondary term score ignored; header alignment only
                                hdr_align = header_alignment.get(id(c))
                                if hdr_align is not None:
                                    s += _SEC_HEADER_WEIGHT * hdr_align
                                    comp["secondary_header"] += _SEC_HEADER_WEIGHT * hdr_align
                                cand_cx = (float(c['x0']) + float(c['x1'])) / 2.0
                                if 'value' in header_map:
                                    hdr = header_map['value']
                                    hx = (float(hdr.get('x0',0.0)) + float(hdr.get('x1',0.0))) / 2.0
                                    dist = abs(cand_cx - hx)
                                    delta = 0.8 * (1.0 / (1.0 + dist / 18.0))
                                    s += delta
                                    comp["value_header_align"] += delta
                                if 'min' in header_map:
                                    hdr = header_map['min']
                                    hx = (float(hdr.get('x0',0.0)) + float(hdr.get('x1',0.0))) / 2.0
                                    dist = abs(cand_cx - hx)
                                    delta = 0.4 * (1.0 / (1.0 + dist / 18.0))
                                    s -= delta
                                    comp["min_header_penalty"] -= delta
                                if 'max' in header_map:
                                    hdr = header_map['max']
                                    hx = (float(hdr.get('x0',0.0)) + float(hdr.get('x1',0.0))) / 2.0
                                    dist = abs(cand_cx - hx)
                                    delta = 0.4 * (1.0 / (1.0 + dist / 18.0))
                                    s -= delta
                                    comp["max_header_penalty"] -= delta
                                # slight preference for smaller horizontal distance from label
                                dx = max(0.0, float(c['x0']) - label_right_x)
                                delta_dx = 0.05 * (1.0 / (1.0 + dx/10.0))
                                s += delta_dx
                                comp["label_dx"] += delta_dx
                                comp["total"] = s
                                combined_sec = header_alignment.get(id(c), 0.0)
                                scored.append((s, c, combined_sec))
                                score_components[id(c)] = comp
                            scored.sort(key=lambda t: t[0], reverse=True)
                            if scored:
                                top_score = scored[0][0]
                                top = [t for t in scored if t[0] >= top_score - 0.1]
                                if len(top) > 1:
                                    conflict_reason = 'multiple candidates with similar scores'
                                chosen = top[0][1]
                                chosen_sec_score = top[0][2]
                                row_components = score_components.get(id(chosen))
                                units_value = chosen.get('units') or units_value
                                cand = chosen['text']
                                if chosen['nval'] is not None and (spec.range_min is not None or spec.range_max is not None):
                                    bad = False
                                    if spec.range_min is not None and chosen['nval'] < spec.range_min:
                                        bad = True
                                    if spec.range_max is not None and chosen['nval'] > spec.range_max:
                                        bad = True
                                    if bad and not cand.rstrip().endswith('(range violation)'):
                                        cand = f"{cand} (range violation)"
                                val = cand
                                if smart_kind == 'number':
                                    val = _strip_units_from_numeric_text(val) or val
                                if score > best_score:
                                    best_score = score
                                    best_info = (p, line_text, right_text_segment, val, smart_kind, line_min_txt, line_max_txt, conflict_reason, (chosen_sec_score if sec_term else None))
                                    # OCR path mirrors best_components via row_components on the chosen row
                                    best_components = row_components
                                    if debug_mode:
                                        print(f"[SMART DEBUG] best_update dpi={dpi} page={p} score={score:.3f} val={val!r}", file=sys.stderr)
                                continue
                        else:
                            # Only use generic extraction for non-numeric smart snaps or when no Smart Position is configured.
                            val = None
                            if smart_kind != 'number' or not has_smart_pos:
                                val = extract_from_line(line_text, right_text_segment, smart_kind)
                                if val and smart_kind == 'number':
                                    val = _strip_units_from_numeric_text(val) or val
                            if val and score > best_score:
                                best_score = score
                                best_info = (p, line_text, right_text_segment, val, smart_kind, line_min_txt, line_max_txt, conflict_reason, None)
                                if debug_mode:
                                    print(f"[SMART DEBUG] best_update(direct) dpi={dpi} page={p} score={score:.3f} val={val!r}", file=sys.stderr)
        finally:
            try:
                if doc is not None:
                    doc.close()
            except Exception:
                pass
        if best_info:
            page_hit, context_line_text, right_text, value_text, smart_kind, line_min_txt, line_max_txt, conflict_reason, sec_found = best_info
            if smart_kind == 'title':
                value_text = _strip_label_tokens(value_text, row_name)
                value_text = _extract_status_from_title(value_text)
            sel_method = "smart_position" if getattr(spec, "smart_position", None) is not None else "smart_score"
            return MatchResult(
                pdf_file=pdf_path.name,
                serial_number=serial_number,
                term=spec.term,
                page=page_hit,
                number=value_text,
                units=units_value,
                context=right_text,
                method='smart:ocr',
                found=True,
                confidence=best_score,
                row_label=context_line_text,
                column_label=None,
                text_source='ocr',
                smart_snap_context=context_line_text,
                smart_snap_type=smart_kind,
                smart_line_min=line_min_txt,
                smart_line_max=line_max_txt,
                smart_conflict=conflict_reason,
                smart_secondary_found=sec_found,
                smart_score_breakdown=best_components,
                smart_selection_method=sel_method,
            )

    # If still not found, try to return best context line (by fuzzy score) to aid debugging
    debug_context = None
    debug_type = spec.smart_snap_type or 'auto'
    try:
        # crude last-attempt using extract_pages_text to get text and slice by group anchors
        # Use OCR fallback here so that OCR-only tables (no direct PDF text) still populate context.
        pages_try = pages or [1]
        text_map, _ = extract_pages_text(pdf_path, pages_try, do_ocr_fallback=True)
        best_score = 0.0
        for p in pages_try:
            txt = text_map.get(p, '') or ''
            for ln in txt.splitlines():
                sc = _fuzzy_ratio(ln, row_name)
                if _normalize_anchor_token(row_name) and _normalize_anchor_token(row_name) in _normalize_anchor_token(ln):
                    sc = max(sc, 0.99)
                if sc > best_score:
                    best_score = sc
                    debug_context = ln.strip()
    except Exception:
        pass

    return MatchResult(
        pdf_file=pdf_path.name,
        serial_number=serial_number,
        term=spec.term,
        page=None,
        number=None,
        units=None,
        context=debug_context or "",
        method='smart:n/a',
        found=False,
        confidence=None,
        row_label=None,
        column_label=None,
        text_source=None,
        error_reason="Smart snap: no matching row/value",
        smart_snap_context=debug_context,
        smart_snap_type=debug_type,
    )


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
    value_format_text, _ = _value_format_info(_effective_value_format(spec))
    fmt_pat = _compile_value_regex(value_format_text) if value_format_text else None

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
    s = s.replace("ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Å“", "-").replace("ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Â", "-")
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

_DOUBLE_HEIGHT_FLAG_RE = re.compile(r"\(\s*double\s+height\s*\)", re.IGNORECASE)


def _value_format_info(raw: Optional[str]) -> Tuple[Optional[str], bool]:
    """
    Returns a cleaned value_format string (suitable for regex compilation)
    and whether the original format requested double-height behavior.
    """
    if not raw:
        return None, False
    s = str(raw).strip()
    if not s:
        return None, False
    double_height = bool(_DOUBLE_HEIGHT_FLAG_RE.search(s))
    if double_height:
        s = _DOUBLE_HEIGHT_FLAG_RE.sub('', s).strip()
    return (s or None), double_height


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
    value_format_text, _ = _value_format_info(_effective_value_format(spec))
    fmt_pat = _compile_value_regex(value_format_text) if value_format_text else None
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
        """Group EasyOCR boxes into text lines using a configurable Y tolerance."""
        rows: List[Dict[str, object]] = []
        try:
            row_eps = float(os.environ.get("OCR_ROW_EPS", "8.0"))
        except Exception:
            row_eps = 8.0
        row_eps = max(0.5, min(50.0, row_eps))
        for it in sorted(items, key=lambda d: (float(d.get("cy", 0.0)), float(d.get("cx", 0.0)))):
            cy = float(it.get("cy", 0.0))
            if not rows or abs(cy - float(rows[-1]["cy"])) > row_eps:
                rows.append({"cy": cy, "items": [it]})
            else:
                rows[-1]["items"].append(it)
        for row in rows:
            row_items = row["items"]  # type: ignore[assignment]
            row_items.sort(key=lambda d: float(d.get("cx", 0.0)))
            row["text"] = " ".join(str(d.get("text", "") or "") for d in row_items).strip()
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


def derive_pdf_identity(pdf_path: Path) -> Tuple[str, str, str]:
    """Best-effort extraction of program, vehicle, and serial_component tokens from the PDF filename."""
    stem = Path(pdf_path).stem
    parts = [p.strip() for p in stem.split("_") if p.strip()]
    program_name = ""
    vehicle_number = ""
    serial_component = ""
    if len(parts) >= 3:
        program_name = parts[0]
        vehicle_number = parts[1]
        serial_component = "_".join(parts[2:])
    elif len(parts) == 2:
        program_name = parts[0]
        serial_component = parts[1]
    elif parts:
        serial_component = parts[0]
    if not serial_component:
        m = SN_REGEX.search(pdf_path.name)
        if m:
            serial_component = m.group(1)
    if not serial_component:
        serial_component = stem or pdf_path.name
    return program_name, vehicle_number, serial_component


def scan_pdf_for_term(pdf_path: Path, serial_number: str, term: str, pages: Sequence[int], window_chars: int, case_sensitive: bool,
                      units_hint: Optional[List[str]] = None,
                      range_filter: Optional[Tuple[Optional[float], Optional[float]]] = None) -> MatchResult:
    """
    Scan a single PDF for a single term (restricted to the provided pages).
    - Uses extract_pages_text(...) to build a map of pageÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢text and a method pipeline string.
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
                    selected = re.sub(r"\s+", " ", selected).lstrip(":-ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Å“ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Â ")
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
    # Build text for constrained pages (or whole doc) — prefer cache
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
            selected = re.sub(r"\s+", " ", selected).lstrip(":-ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Å“ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Â ")
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

def write_outputs_excel_or_csv(
    output_xlsx: Path,
    results_matrix: Dict[str, Dict[str, Optional[str]]],
    term_order: List[str],
    term_pages_raw: Dict[str, str],
    metadata_rows: List[Dict],
    errors_rows: List[Dict],
    csv_fallback_prefix: Path,
    tables_rows: Optional[List[Dict]] = None,
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
        display_names = {
            "extracted_value": "Extracted Value",
            "term_label": "Term Label",
            "data_group": "Data Group",
            "units_hint": "Units Hint",
            "term": "Term",
            "search_term": "Search Term",
            "program_name": "Program Name",
            "vehicle_number": "Vehicle Number",
            "serial_component": "Serial Component",
            "smart_score": "Smart Score",
            "smart_snap_type": "Smart Snap Type",
            "smart_line_min": "Smart Line Min",
            "smart_line_max": "Smart Line Max",
            "smart_conflict": "Smart Conflict",
            "smart_secondary_found": "Smart Secondary Found",
            "group_after": "Group After",
            "group_before": "Group Before",
            "error_reason": "Error Reason",
            "range_min": "Range Min",
            "range_max": "Range Max",
            "text_source": "Text Source",
            "return_type": "Return Type",
            "pages_raw": "Pages Raw",
            "smart_snap_context": "Smart Snap Context",
            "smart_position": "Smart Position",
            "secondary_term": "Secondary Term",
        }
        df_meta = df_meta.rename(columns={k: v for k, v in display_names.items() if k in df_meta.columns})
        error_cols = ["pdf_file", "program_name", "vehicle_number", "serial_component", "term", "error", "method", "page", "column", "row", "group_after", "group_before"]
        if errors_rows:
            df_errors = pd.DataFrame(errors_rows, columns=error_cols)
        else:
            df_errors = pd.DataFrame(columns=error_cols)

        # Prepare tables sheet if provided
        df_tables = None
        if tables_rows:
            # Determine max dynamic columns col_1..col_N
            max_cols = 0
            for r in tables_rows:
                for k in r.keys():
                    if isinstance(k, str) and k.startswith("col_"):
                        try:
                            idx = int(k.split("_", 1)[1])
                            if idx > max_cols:
                                max_cols = idx
                        except Exception:
                            pass
            base_cols = ["pdf_file", "program_name", "vehicle_number", "serial_component", "page", "section"]
            dyn_cols = [f"col_{i}" for i in range(1, max_cols + 1)]
            tab_cols = base_cols + dyn_cols
            df_tables = pd.DataFrame(tables_rows, columns=tab_cols)

        # Write Excel with two sheets
        with pd.ExcelWriter(output_xlsx, engine="xlsxwriter") as writer:
            # Sheet 1: Smart/standard results
            df_results.to_excel(writer, sheet_name="results", index=False)
            # Sheet 2: full table extraction (if any)
            if df_tables is not None:
                df_tables.to_excel(writer, sheet_name="tables", index=False)
            # Next sheet(s): detailed metadata and errors
            df_meta.to_excel(writer, sheet_name="metadata", index=False)
            df_errors.to_excel(writer, sheet_name="errors", index=False)

            # Cosmetic improvements: freeze header rows and set reasonable column widths
            ws_res = writer.sheets["results"]
            ws_res.freeze_panes(1, 0)
            ws_meta = writer.sheets["metadata"]
            ws_meta.freeze_panes(1, 0)
            ws_err = writer.sheets["errors"]
            ws_err.freeze_panes(1, 0)
            ws_tab = writer.sheets.get("tables") if (tables_rows) else None
            if ws_tab is not None:
                ws_tab.freeze_panes(1, 0)

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
            if ws_tab is not None and df_tables is not None:
                for i, col in enumerate(df_tables.columns):
                    width = min(60, max(10, int(df_tables[col].astype(str).str.len().max() if not df_tables.empty else len(col)) + 2))
                    ws_tab.set_column(i, i, width)

        print(f"[DONE] Excel written -> {output_xlsx}")
        return

    # ---------- CSV fallback path ----------
    results_csv = csv_fallback_prefix.with_suffix(".results.csv")
    metadata_csv = csv_fallback_prefix.with_suffix(".metadata.csv")
    errors_csv = csv_fallback_prefix.with_suffix(".errors.csv")
    tables_csv = csv_fallback_prefix.with_suffix(".tables.csv")

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
    meta_cols = [
        "pdf_file", "program_name", "vehicle_number", "serial_component",
        "term", "term_label", "data_group",
        "found", "page", "extracted_value", "units",
        "text_source", "smart_score",
        "range_min", "range_max", "units_hint",
        "return_type", "group_after", "group_before", "value_format",
        "pages_raw", "mode",
        "error_reason",
        "smart_snap_context", "smart_snap_type", "smart_line_min", "smart_line_max",
        "smart_conflict", "smart_secondary_found", "smart_position", "secondary_term"
    ]
    display_names = {
        "extracted_value": "Extracted Value",
        "term_label": "Term Label",
        "data_group": "Data Group",
        "units_hint": "Units Hint",
        "term": "Search Term",
        "program_name": "Program Name",
        "vehicle_number": "Vehicle Number",
        "serial_component": "Serial Component",
        "smart_score": "Smart Score",
        "smart_snap_type": "Smart Snap Type",
        "smart_line_min": "Smart Line Min",
        "smart_line_max": "Smart Line Max",
        "smart_conflict": "Smart Conflict",
        "smart_secondary_found": "Smart Secondary Found",
        "group_after": "Group After",
        "group_before": "Group Before",
        "error_reason": "Error Reason",
        "range_min": "Range Min",
        "range_max": "Range Max",
        "text_source": "Text Source",
        "return_type": "Return Type",
        "pages_raw": "Pages Raw",
        "smart_snap_context": "Smart Snap Context",
        "smart_position": "Smart Position",
        "secondary_term": "Secondary Term",
    }
    with metadata_csv.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow([display_names.get(col, col) for col in meta_cols])
        for r in metadata_rows:
            writer.writerow([r.get(col) for col in meta_cols])
    with errors_csv.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["pdf_file", "program_name", "vehicle_number", "serial_component", "term", "error", "method", "page", "column", "row", "group_after", "group_before"])
        for r in errors_rows:
            writer.writerow([
                r.get("pdf_file"),
                r.get("program_name"),
                r.get("vehicle_number"),
                r.get("serial_component"),
                r.get("term"),
                r.get("error"),
                r.get("method"),
                r.get("page"),
                r.get("column"),
                r.get("row"),
                r.get("group_after"),
                r.get("group_before"),
            ])
    # Write tables CSV if provided
    if tables_rows:
        try:
            # Determine schema
            max_cols = 0
            for r in tables_rows:
                for k in r.keys():
                    if isinstance(k, str) and k.startswith("col_"):
                        try:
                            idx = int(k.split("_", 1)[1])
                            if idx > max_cols:
                                max_cols = idx
                        except Exception:
                            pass
            base_cols = ["pdf_file", "program_name", "vehicle_number", "serial_component", "page", "section"]
            dyn_cols = [f"col_{i}" for i in range(1, max_cols + 1)]
            cols = base_cols + dyn_cols
            with tables_csv.open("w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(cols)
                for r in tables_rows:
                    w.writerow([r.get(c, "") for c in cols])
        except Exception as e:
            print(f"[WARN] Could not write tables CSV fallback: {e}")
    print(f"[DONE] CSV fallback written -> {results_csv}, {metadata_csv}, {errors_csv}{', ' + str(tables_csv) if tables_rows else ''}")


def run_scan(
    input_path: Path,
    pdf_folder: Path,
    output_csv: Path,
    output_json: Path,
    output_xlsx: Path,
    window_chars: int,
    case_sensitive: bool
) -> None:
    """
    Orchestrate the entire scan:
      1) Load terms from input file
      2) Iterate through PDFs in the target folder
      3) For each (pdf, term) pair, collect the best match and record metadata
      4) Write progress JSON as we go (crash resilience)
      5) Produce the final Excel (or CSVs) and a flat CSV summary
    """
    # Step 1: Load the term specs
    terms = load_terms(input_path)
    if not terms:
        print("[WARN] No terms found in input. Ensure headers 'Term' and 'Pages' exist.")
        return

    # Partition terms: keep 'full table' rows separate from standard scan rows
    def _is_full_table(t: TermSpec) -> bool:
        try:
            m = (t.mode or '').strip().lower()
            return m in ('full table', 'full_table', 'fulltable')
        except Exception:
            return False

    full_table_terms: List[TermSpec] = [t for t in terms if _is_full_table(t)]
    scan_terms: List[TermSpec] = [t for t in terms if not _is_full_table(t)]

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
    # Outputs: aggregated JSON for the run, plus per-EIDP CSV/JSON files

    # Reroute output paths into the run_dir regardless of CLI-provided paths.
    output_json = run_dir / "scan_results.json"
    output_xlsx = run_dir / "scan_results_flat.xlsx"
    print(f"[INFO] Outputs will be saved under: {run_dir}")

    # Helper for safe filename tokens (for per-EIDP outputs)
    def _safe_token(s: Optional[str]) -> str:
        try:
            t = (s or "").strip()
            # Replace Windows-invalid filename chars but preserve spaces
            t = re.sub(r"[<>:\"/\\|?*]+", "_", t)
            # Trim trailing/leading dots and spaces
            t = t.strip(" .")
            return t or "unknown"
        except Exception:
            return "unknown"

    # Prepare structures for the wide "results" sheet and the "metadata" sheet
    term_order = [t.term for t in scan_terms]                  # preserve input order
    term_pages_raw = {t.term: t.pages_raw for t in terms} # map term ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ original "Pages" string
    results_matrix: Dict[str, Dict[str, Optional[str]]] = {t.term: {} for t in terms}  # term ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ {SN ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ number}
    metadata_rows: List[Dict] = []  # detailed records per (pdf, term)
    summary: List[Dict] = []        # JSON audit entries
    errors_rows: List[Dict] = []    # rows for the errors report
    def _result_row_key(t: TermSpec) -> str:
        return (t.term_label or t.term or "").strip()

    # Override term lists to exclude 'full table' rows from the wide matrix
    term_order = [_result_row_key(t) for t in scan_terms]
    term_pages_raw = {_result_row_key(t): t.pages_raw for t in scan_terms}
    results_matrix = {_result_row_key(t): {} for t in scan_terms}
    results_priority: Dict[str, Dict[str, Tuple[int, int, float, int]]] = {
        _result_row_key(t): {} for t in scan_terms
    }
    serial_meta: Dict[str, Dict[str, str]] = {}
    tables_rows_agg: List[Dict] = []  # aggregated full-table rows across all PDFs

    # Capture global extraction tunables for debug visibility in JSON
    try:
        _xy_fuzz_debug = float(os.environ.get("XY_FUZZ", "0.75"))
    except Exception:
        _xy_fuzz_debug = None
    try:
        _ocr_row_eps_debug = float(os.environ.get("OCR_ROW_EPS", "8.0"))
    except Exception:
        _ocr_row_eps_debug = None

    # Step 3: For each PDF, scan for each term
    for pdf_path in sorted(pdfs):
        # Derive identifiers from the filename
        program_hint, vehicle_hint, serial_component = derive_pdf_identity(pdf_path)
        data_id = (serial_component or pdf_path.stem or pdf_path.name).strip()
        if not data_id:
            data_id = pdf_path.name
        info_defaults = {
            "program_name": program_hint,
            "vehicle_number": vehicle_hint,
            "serial_component": serial_component or data_id,
        }
        existing_meta = serial_meta.get(data_id, {})
        serial_meta[data_id] = {
            "program_name": existing_meta.get("program_name") or info_defaults["program_name"],
            "vehicle_number": existing_meta.get("vehicle_number") or info_defaults["vehicle_number"],
            "serial_component": existing_meta.get("serial_component") or info_defaults["serial_component"],
        }
        try:
            label = serial_meta[data_id]["serial_component"] or data_id
        except Exception:
            label = data_id
        print(f"[INFO] Scanning: {pdf_path.name}  [Data: {label}]")

        # Per-PDF accumulation for outputs
        summary_pdf: List[Dict] = []

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

        # Handle any 'full table' extractions for this PDF first
        for t in full_table_terms:
            try:
                pages_for_t = t.pages if getattr(t, 'pages', None) else (sorted(pre_map.keys()) if pre_map else (list(range(1, page_count + 1)) if page_count > 0 else [1]))
            except Exception:
                pages_for_t = sorted(pre_map.keys()) if pre_map else [1]
            ft_rows = _extract_full_table_rows_for_pdf(
                pdf_path=pdf_path,
                pages=list(pages_for_t),
                group_after=getattr(t, 'group_after', None),
                group_before=getattr(t, 'group_before', None),
                program_name=serial_meta[data_id].get("program_name"),
                vehicle_number=serial_meta[data_id].get("vehicle_number"),
                serial_component=serial_meta[data_id].get("serial_component"),
            )
            try:
                tables_rows_agg.extend(ft_rows)
            except Exception:
                pass
            # Add a concise metadata record for audit/JSON
            rows_count = len(ft_rows) if ft_rows is not None else 0
            term_label_ft = getattr(t, 'term_label', None) or getattr(t, 'term', '') or 'Full Table'
            meta_ft = {
                "pdf_file": pdf_path.name,
                "program_name": serial_meta[data_id].get("program_name"),
                "vehicle_number": serial_meta[data_id].get("vehicle_number"),
                "serial_component": serial_meta[data_id].get("serial_component") or data_id,
                "term": term_label_ft,
                "term_label": term_label_ft,
                "data_group": getattr(t, 'data_group', None) or '',
                "found": bool(rows_count > 0),
                "page": None,
                "extracted_value": f"{rows_count} rows",
                "units": None,
                "text_source": "table",
                "smart_score": None,
                "mode": "full table",
                "pages_raw": getattr(t, 'pages_raw', ''),
                "range_min": None,
                "range_max": None,
                "units_hint": None,
                "return_type": None,
                "group_after": getattr(t, 'group_after', None),
                "group_before": getattr(t, 'group_before', None),
                "value_format": getattr(t, 'value_format', None),
                "error_reason": (None if rows_count > 0 else "No table rows in selected range"),
                "smart_snap_context": None,
                "smart_snap_type": None,
                "smart_line_min": None,
                "smart_line_max": None,
                "smart_conflict": None,
                "smart_secondary_found": None,
                "smart_position": None,
                "secondary_term": getattr(t, 'secondary_term', None),
                "search_term": getattr(t, 'term', '') or 'Full Table',
                "xy_fuzz": _xy_fuzz_debug,
                "ocr_row_eps": _ocr_row_eps_debug,
            }
            metadata_rows.append(meta_ft)
            summary.append(meta_ft)
            summary_pdf.append(meta_ft)

        # Search each configured term (excluding full table rows) within the allowed page ranges
        total_terms = len(scan_terms)
        completed = 0
        found_count = 0
        prev_pct = -1
        # Initial progress line
        try:
            print(f"[PROGRESS] Terms: 0% (0/{total_terms}) | Found: 0")
        except Exception:
            pass

        for idx, t in enumerate(scan_terms, start=1):
            mode = (t.mode or "").lower() if hasattr(t, 'mode') else ""
            if mode == "line":
                res = scan_pdf_for_term_line(pdf_path, data_id, t, window_chars, case_sensitive)
            elif mode == "smart":
                res = scan_pdf_for_term_smart(pdf_path, data_id, t, window_chars, case_sensitive)
            elif mode in ("table(xy)", "xy", "table") or (not mode and getattr(t, 'line', None) and getattr(t, 'column', None)):
                res = scan_pdf_for_term_xy(pdf_path, data_id, t, window_chars, case_sensitive)
            else:
                res = scan_pdf_for_term_nearest(pdf_path, data_id, t, window_chars, case_sensitive)
            # Normalize units/value when found
            ret_kind = (getattr(t, 'return_type', None) or 'number').strip().lower()
            # Smart mode: if smart snap type is not numeric, treat as string
            try:
                if mode == 'smart' and (getattr(t, 'smart_snap_type', None) or '').strip().lower() not in ('', 'auto', 'number'):
                    ret_kind = 'string'
            except Exception:
                pass
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

            # Fill the matrix cell for this (term, serial_component)
            if res.found:
                found_count += 1
                cell_value = res.number
                # Compute a selection priority so that Smart Position
                # rows win over auxiliary rows for the same Search
                # Term / serial component, with secondary-term hits,
                # confidence, and text source as tie-breakers.
                has_smart_pos = getattr(t, "smart_position", None) is not None
                raw_sec = getattr(res, "smart_secondary_found", None)
                # Rank secondary-term hits: True/high score > unknown > explicit False/zero
                if isinstance(raw_sec, (int, float)):
                    if raw_sec >= 0.9:
                        sec_rank = 2
                    elif raw_sec <= 0.0:
                        sec_rank = 0
                    else:
                        sec_rank = 1
                else:
                    if raw_sec is True:
                        sec_rank = 2
                    elif raw_sec is None:
                        sec_rank = 1
                    else:
                        sec_rank = 0
                try:
                    conf = float(getattr(res, "confidence", 0.0) or 0.0)
                except Exception:
                    conf = 0.0
                src = (getattr(res, "text_source", None) or "").strip().lower()
                src_rank = 1 if src == "pdf" else 0
                new_priority = (
                    1 if has_smart_pos else 0,
                    sec_rank,
                    conf,
                    src_rank,
                )
                row_key = _result_row_key(t)
                prev_priority = results_priority.get(row_key, {}).get(data_id)
                # Only update the matrix if this result is strictly better
                # than any previous candidate for the same (term, serial).
                if (prev_priority is None) or (new_priority > prev_priority):
                    results_priority.setdefault(row_key, {})[data_id] = new_priority
                    results_matrix.setdefault(row_key, {})[data_id] = cell_value
            else:
                err_msg = res.error_reason or "No match found"
                # Only record an error if no successful value exists yet
                row_key = _result_row_key(t)
                cell_map = results_matrix.setdefault(row_key, {})
                if data_id not in cell_map:
                    cell_map[data_id] = f"ERROR: {err_msg}"
                pdf_stem = Path(res.pdf_file).stem if res.pdf_file else ""
                parts = [p for p in pdf_stem.split("_") if p]
                err_program = err_vehicle = err_serial_component = None
                if len(parts) >= 3:
                    err_program, err_vehicle, err_serial_component = parts[0], parts[1], "_".join(parts[2:])
                elif len(parts) == 2:
                    err_program, err_serial_component = parts[0], parts[1]
                elif len(parts) == 1:
                    err_serial_component = parts[0]
                errors_rows.append({
                    "pdf_file": res.pdf_file,
                    "program_name": err_program,
                    "vehicle_number": err_vehicle,
                    "serial_component": err_serial_component,
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
            range_min_schema = getattr(t, 'range_min', None)
            range_max_schema = getattr(t, 'range_max', None)
            range_min_disabled = getattr(t, 'range_min_disabled', False)
            range_max_disabled = getattr(t, 'range_max_disabled', False)
            smart_line_min = getattr(res, 'smart_line_min', None)
            smart_line_max = getattr(res, 'smart_line_max', None)
            units_hint_raw = getattr(t, 'units_hint', None)
            if isinstance(units_hint_raw, (list, tuple, set)):
                units_hint_display = "|".join(
                    str(u).strip() for u in units_hint_raw if str(u).strip()
                ) or None
            else:
                units_hint_display = str(units_hint_raw).strip() if units_hint_raw is not None and str(units_hint_raw).strip() else None

            if range_min_disabled:
                effective_range_min = None
            else:
                effective_range_min = range_min_schema if range_min_schema is not None else smart_line_min
            if range_max_disabled:
                effective_range_max = None
            else:
                effective_range_max = range_max_schema if range_max_schema is not None else smart_line_max

            pdf_stem = Path(res.pdf_file).stem if res.pdf_file else ""
            parts = [p for p in pdf_stem.split("_") if p]
            program_name = vehicle_number = serial_component = None
            if len(parts) >= 3:
                program_name, vehicle_number, serial_component = parts[0], parts[1], "_".join(parts[2:])
            elif len(parts) == 2:
                program_name, serial_component = parts[0], parts[1]
            elif len(parts) == 1:
                serial_component = parts[0]

            info_entry = serial_meta.setdefault(res.serial_number, {"program_name": "", "vehicle_number": "", "serial_component": ""})
            if program_name:
                info_entry["program_name"] = info_entry.get("program_name") or program_name
            if vehicle_number:
                info_entry["vehicle_number"] = info_entry.get("vehicle_number") or vehicle_number
            if serial_component:
                info_entry["serial_component"] = info_entry.get("serial_component") or serial_component

            term_label_out = (t.term_label or t.term or "").strip()
            data_group_out = (t.data_group or "").strip()

            component_value = info_entry.get("serial_component") or serial_component or res.serial_number
            meta = {
                "pdf_file": res.pdf_file,
                "program_name": info_entry.get("program_name") or program_name,
                "vehicle_number": info_entry.get("vehicle_number") or vehicle_number,
                "serial_component": component_value,
                # Expose the user-facing label as the primary Term field
                # and keep the raw search term separately for debugging.
                "term": term_label_out or res.term,
                "term_label": term_label_out,
                "data_group": data_group_out,
                "found": res.found,
                "page": res.page,
                "extracted_value": res.number,
                "units": res.units,
                "text_source": res.text_source,
                "smart_score": res.confidence,
                "mode": ((t.mode or ("table(xy)" if getattr(t, 'line', None) and getattr(t, 'column', None) else "nearest")) if hasattr(t, 'mode') else "nearest"),
                "pages_raw": getattr(t, 'pages_raw', ""),
                "range_min": effective_range_min,
                "range_max": effective_range_max,
                "units_hint": units_hint_display,
                "return_type": getattr(t, 'return_type', None),
                "group_after": getattr(t, 'group_after', None),
                "group_before": getattr(t, 'group_before', None),
                "value_format": getattr(t, 'value_format', None),
                "error_reason": res.error_reason,
                "smart_snap_context": getattr(res, 'smart_snap_context', None),
                "smart_snap_type": (getattr(t, 'smart_snap_type', None) or getattr(res, 'smart_snap_type', None)),
                "smart_line_min": smart_line_min,
                "smart_line_max": smart_line_max,
                "smart_conflict": getattr(res, 'smart_conflict', None),
                "smart_secondary_found": getattr(res, 'smart_secondary_found', None),
                "smart_score_breakdown": getattr(res, 'smart_score_breakdown', None),
                "smart_selection_method": getattr(res, 'smart_selection_method', None),
                "smart_position": getattr(t, 'smart_position', None),
                "secondary_term": getattr(t, 'secondary_term', None),
                "search_term": res.term,
                "xy_fuzz": _xy_fuzz_debug,
                "ocr_row_eps": _ocr_row_eps_debug,
            }
            metadata_rows.append(meta)
            summary.append(meta)
            summary_pdf.append(meta)
            # No per-PDF accumulation

            # Update and print progress for this PDF's terms
            try:
                completed = idx
                pct = int((completed * 100) / max(1, total_terms))
                # Print at meaningful increments to avoid flooding the console
                if pct != prev_pct and (total_terms <= 20 or pct % 5 == 0 or completed == total_terms):
                    print(f"[PROGRESS] Terms: {pct}% ({completed}/{total_terms}) | Found: {found_count}")
                    prev_pct = pct
            except Exception:
                pass

        # Step 4: Persist JSON progress incrementally (so partial work isn't lost)
        try:
            with output_json.open("w", encoding="utf-8") as jf:
                json.dump(summary, jf, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"[WARN] Could not write JSON during loop: {e}")

        # Per-PDF JSON + CSV directly in run_data folder
        try:
            safe_id = _safe_token(serial_meta.get(data_id, {}).get("serial_component") or data_id)
        except Exception:
            safe_id = _safe_token(data_id)
        per_json = run_dir / f"scan_results_{safe_id}.json"
        per_csv = run_dir / f"scan_results_flat_{safe_id}.csv"
        # Write per-PDF JSON (details + match summary metadata)
        try:
            try:
                header_w = float(os.environ.get("SMART_SEC_HEADER_W", "0.7"))
            except Exception:
                header_w = 0.7
            match_summary_row = {
                "_kind": "match_summary",
                "description": (
                    "Smart Snap scoring: row smart_score is the best fuzzy match to the row anchor; "
                    "numeric candidate ranking adds: +2.0 if value is between row min/max, "
                    "+0.4 if units match Units Hint, up to +0.4 for values within configured Range, "
                    f"+{header_w:.2f} * secondary_header_alignment for X alignment with the Secondary Term header, "
                    "plus smaller adjustments based on distance to Value/Min/Max headers and distance from the label."
                ),
                "secondary_header_weight": header_w,
                "secondary_vertical_weight": 0.0,
                "units_hint_weight": 0.4,
                "range_weight_full": 0.4,
                "xy_fuzz": _xy_fuzz_debug,
                "ocr_row_eps": _ocr_row_eps_debug,
            }
            with per_json.open("w", encoding="utf-8") as jf:
                json.dump(summary_pdf + [match_summary_row], jf, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"[WARN] Could not write per-PDF JSON for {safe_id}: {e}")
        # Write per-PDF flat CSV (header mapped to friendly names)
        try:
            cols_display = [
                "pdf_file", "program_name", "vehicle_number", "serial_component", "term_label", "data_group", "term",
                "found", "page",
                "extracted_value", "units", "units_hint",
                "range_min", "range_max",
                "text_source", "smart_score",
                "smart_snap_type", "smart_line_min", "smart_line_max",
                "smart_conflict", "smart_secondary_found",
                "group_after", "group_before", "error_reason",
            ]
            display_names = {
                "extracted_value": "Extracted Value",
                "term_label": "Term Label",
                "data_group": "Data Group",
                "units_hint": "Units Hint",
                "term": "Search Term",
                "program_name": "Program Name",
                "vehicle_number": "Vehicle Number",
                "serial_component": "Serial Component",
                "smart_score": "Smart Score",
                "smart_snap_type": "Smart Snap Type",
                "smart_line_min": "Smart Line Min",
                "smart_line_max": "Smart Line Max",
                "smart_conflict": "Smart Conflict",
                "smart_secondary_found": "Smart Secondary Found",
                "group_after": "Group After",
                "group_before": "Group Before",
                "error_reason": "Error Reason",
                "range_min": "Range Min",
                "range_max": "Range Max",
                "text_source": "Text Source",
            }
            with per_csv.open("w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow([display_names.get(col, col) for col in cols_display])
                for row in summary_pdf:
                    w.writerow([row.get(col) for col in cols_display])
            print(f"[DONE] Per-EIDP CSV -> {per_csv}")
            print(f"[DONE] Per-EIDP JSON -> {per_json}")
        except Exception as e:
            print(f"[WARN] Could not write per-PDF CSV for {safe_id}: {e}")
        # Finalize per-PDF terms progress to 100%
        try:
            print(f"[PROGRESS] Terms: 100% ({total_terms}/{total_terms}) | Found: {found_count}")
        except Exception:
            pass

    # Output: Flat extraction table as Excel and details JSON
    # Columns for the flat (Excel/CSV) extraction table:
    # - omit verbose context
    # - focus on Smart Snap scoring details
    # - include Smart Snap score (confidence) and effective range bounds
    cols_display = [
        "pdf_file", "program_name", "vehicle_number", "serial_component", "term_label", "data_group", "term",
        "found", "page",
        "extracted_value", "units", "units_hint",
        "range_min", "range_max",
        "text_source", "smart_score",
        "smart_snap_type", "smart_line_min", "smart_line_max",
        "smart_conflict", "smart_secondary_found",
        "group_after", "group_before", "error_reason",
    ]
    wrote_xlsx = True  # XLSX disabled for legacy block; using consolidated writer below
    if False:
            # (legacy block disabled)
            display_names = {
                "extracted_value": "Extracted Value",
                "term_label": "Term Label",
                "data_group": "Data Group",
                "units_hint": "Units Hint",
                "term": "Term",
                "search_term": "Search Term",
                "program_name": "Program Name",
                "vehicle_number": "Vehicle Number",
                "serial_component": "Serial Component",
                "smart_score": "Smart Score",
                "smart_snap_type": "Smart Snap Type",
                "smart_line_min": "Smart Line Min",
                "smart_line_max": "Smart Line Max",
                "smart_conflict": "Smart Conflict",
                "smart_secondary_found": "Smart Secondary Found",
                "group_after": "Group After",
                "group_before": "Group Before",
                "error_reason": "Error Reason",
                "range_min": "Range Min",
                "range_max": "Range Max",
                "text_source": "Text Source",
                "return_type": "Return Type",
                "pages_raw": "Pages Raw",
                "smart_snap_context": "Smart Snap Context",
                "smart_position": "Smart Position",
                "secondary_term": "Secondary Term",
            }
            rows_for_df = []
            for row in summary:
                rows_for_df.append({
                    "pdf_file": row.get("pdf_file"),
                    "program_name": row.get("program_name"),
                    "vehicle_number": row.get("vehicle_number"),
                    "serial_component": row.get("serial_component"),
                    "term_label": row.get("term_label"),
                    "data_group": row.get("data_group"),
                    "term": row.get("term"),
                    "found": row.get("found"),
                    "page": row.get("page"),
                    "extracted_value": row.get("extracted_value"),
                    "units": row.get("units"),
                    "units_hint": row.get("units_hint"),
                    "range_min": row.get("range_min"),
                    "range_max": row.get("range_max"),
                    "text_source": row.get("text_source"),
                    "smart_score": row.get("smart_score"),
                    "smart_snap_type": row.get("smart_snap_type"),
                    "smart_line_min": row.get("smart_line_min"),
                    "smart_line_max": row.get("smart_line_max"),
                    "smart_conflict": row.get("smart_conflict"),
                    "smart_secondary_found": row.get("smart_secondary_found"),
                    "group_after": row.get("group_after") or "",
                    "group_before": row.get("group_before") or "",
                    "error_reason": row.get("error_reason") or "",
                })
            df = _pd.DataFrame(rows_for_df, columns=cols_display)
            df = df.rename(columns={k: v for k, v in display_names.items() if k in df.columns})
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
        # except Exception as e:
        #     print(f"[WARN] Could not write Excel extraction table: {e}")
    if False:
        # Fallback: write CSV next to intended xlsx (same basename) if Excel writer not available
        try:
            fallback_csv = output_xlsx.with_suffix(".csv")
            with fallback_csv.open("w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow([display_names.get(col, col) for col in cols_display])
                for row in summary:
                    w.writerow([row.get(col) for col in cols_display])
            err_csv = output_xlsx.with_suffix(".errors.csv")
            with err_csv.open("w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(cols_display)
                for row in summary:
                    if row.get("found"):
                        continue
                    w.writerow([row.get(col) for col in cols_display])
            print(f"[DONE] Extraction table (CSV fallback) -> {fallback_csv}; errors -> {err_csv}")
        except Exception as e:
            print(f"[WARN] Could not write extraction table fallback: {e}")

    # Consolidated outputs (results + metadata + errors + tables) to one workbook
    try:
        csv_prefix = output_xlsx.with_name(output_xlsx.stem)
        write_outputs_excel_or_csv(
            output_xlsx=output_xlsx,
            results_matrix=results_matrix,
            term_order=term_order,
            term_pages_raw=term_pages_raw,
            metadata_rows=metadata_rows,
            errors_rows=errors_rows,
            csv_fallback_prefix=csv_prefix,
            tables_rows=(tables_rows_agg if tables_rows_agg else None),
        )
    except Exception as e:
        print(f"[WARN] Could not write consolidated run workbook: {e}")

    # Finalize JSON with a global match-summary metadata row explaining scoring weights
    try:
        try:
            header_w = float(os.environ.get("SMART_SEC_HEADER_W", "0.7"))
        except Exception:
            header_w = 0.7
        match_summary_row = {
            "_kind": "match_summary",
            "description": (
                "Smart Snap scoring: row smart_score is the best fuzzy match to the row anchor; "
                "numeric candidate ranking adds: +2.0 if value is between row min/max, "
                "+0.4 if units match Units Hint, up to +0.4 for values within configured Range, "
                    f"+{header_w:.2f} * secondary_header_alignment for X alignment with the Secondary Term header, "
                "plus smaller adjustments based on distance to Value/Min/Max headers and distance from the label."
            ),
            "secondary_header_weight": header_w,
            "secondary_vertical_weight": 0.0,
            "units_hint_weight": 0.4,
            "range_weight_full": 0.4,
        }
        with output_json.open("w", encoding="utf-8") as jf:
            json.dump(summary + [match_summary_row], jf, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"[WARN] Could not finalize JSON with match summary: {e}")

    print(f"[DONE] Details JSON -> {output_json}")

    # Remove legacy aggregate artifact if present
    try:
        agg_path = Path("Product_Data_File") / "EIDP_data.csv"
        if agg_path.exists():
            agg_path.unlink(missing_ok=True)  # type: ignore[call-arg]
            print(f"[CLEANUP] Removed legacy aggregate -> {agg_path}")
    except Exception:
        pass

    # Update the persistent run registry with all serial components in this run
    try:
        run_ids: List[str] = []
        # Prefer keys discovered in results_matrix
        for term, sn_map in results_matrix.items():
            for sn in sn_map.keys():
                if sn not in run_ids:
                    run_ids.append(sn)
        if run_ids:
            for sn in run_ids:
                serial_meta.setdefault(sn, {"program_name": "", "vehicle_number": "", "serial_component": sn})
            _update_run_registry(run_dir, run_ids, serial_meta)
            print("[DONE] Run registry updated (run_registry.csv)")
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
        description="Scan PDFs for terms and nearest numbers, producing a matrix by data identifier (serial component)."
    )
    parser.add_argument("--input", required=True, help="Path to terms file (.csv, .xlsx, or .xls). Headers: Term, Pages [Line, Column, Range, Units optional]")
    parser.add_argument("--pdf-folder", required=True, help='Folder containing PDFs to scan (e.g., "EIDP import folder")')
    parser.add_argument("--output-csv", default="scan_results_flat.csv", help="Flat CSV summary (legacy)")
    parser.add_argument("--output-json", default="scan_results.json", help="Path to write JSON details")
    parser.add_argument("--output-xlsx", default="scan_results.xlsx", help="Excel workbook with 'results' and 'metadata' sheets")
    parser.add_argument("--window-chars", type=int, default=160, help="Search window size around term (ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â± chars)")
    parser.add_argument("--case-sensitive", action="store_true", help="Enable case-sensitive term matching")
    parser.add_argument("--quiet", action="store_true", help="Reduce console output (suppress progress/debug)")
    args = parser.parse_args()

    # Normalize and validate file/folder paths
    input_path = Path(args.input)
    pdf_folder = Path(args.pdf_folder)
    output_csv = Path(args.output_csv)
    output_json = Path(args.output_json)
    output_xlsx = Path(args.output_xlsx)

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
        window_chars=args.window_chars,
        case_sensitive=args.case_sensitive
    )


if __name__ == "__main__":
    main()










