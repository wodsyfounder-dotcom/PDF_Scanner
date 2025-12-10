#!/usr/bin/env python3
"""
EIDP Term Scanner (Application deliverable)

Wrapper that loads the local core module and applies small runtime patches
before invoking its main(). This preserves the deliverable layout and lets us
refine 'nearest' behavior without directly editing the monolithic core file.
"""

from pathlib import Path
import sys
import importlib.util
from dataclasses import replace

HERE = Path(__file__).resolve().parent
LOCAL = HERE / "eidp_term_scanner.core.py"
ROOT = HERE.parents[0]
LEGACY = ROOT / "eidp_term_scanner.py"


def _load_core_from(path: Path):
    spec = importlib.util.spec_from_file_location("eidp_core", path)
    if spec is None or spec.loader is None:
        raise SystemExit(f"Unable to load scanner core from {path}")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)  # type: ignore[attr-defined]
    return mod


def _apply_patches(core):
    """Make 'nearest' behave as 'same-line to the right of anchor'.

    Route nearest through the existing line-mode extractor using
    field_split='groups' and field_index=1 for both strings and numbers.
    """
    if not hasattr(core, "scan_pdf_for_term_line") or not hasattr(core, "scan_pdf_for_term_nearest"):
        return

    import re as _re

    def _same_line_right_value(text: str, anchor: str, *, ret_is_string: bool, case_sensitive: bool):
        if not text:
            return None
        hay = text if case_sensitive else text.lower()
        pat = anchor if case_sensitive else anchor.lower()
        best = None
        for line in text.splitlines():
            lcmp = line if case_sensitive else line.lower()
            pos = lcmp.find(pat) if pat else 0
            if pos == -1:
                continue
            tail = line[pos + len(anchor):] if anchor else line
            # split on 2+ spaces/tabs; keep single spaces inside the field
            parts = [p for p in _re.split(r"[ \t]{2,}", tail.strip()) if p]
            if not parts:
                continue
            field = _re.sub(r"\s+", " ", parts[0]).lstrip(":- ").strip()
            if not field:
                continue
            if ret_is_string:
                return field
            # numeric: first number inside field
            num_re = getattr(core, 'NUMBER_REGEX', None)
            if num_re is None:
                num_re = _re.compile(r"[-+]?\d{1,3}(?:,\d{3})*(?:\.\d+)?(?:[eE][+-]?\d+)?")
            m = next((m for m in num_re.finditer(field)), None)
            if m:
                num_txt = m.group(0)
                units = getattr(core, 'extract_units', lambda v: None)(field)
                return num_txt, units
            # else continue scanning lines
        return None

    def patched_scan_pdf_for_term_nearest(pdf_path, serial_number, spec, window_chars, case_sensitive):
        anchor_text = (getattr(spec, 'anchor', None) or getattr(spec, 'term', None) or '').strip()
        ret_kind = (getattr(spec, 'return_type', None) or 'number').strip().lower()
        ret_is_string = (ret_kind == 'string')

        # Build text map for target pages
        try:
            page_count = core.get_pdf_page_count(pdf_path)
        except Exception:
            page_count = 0
        target_pages = getattr(spec, 'pages', None) or (list(range(1, page_count + 1)) if page_count else [1])
        try:
            page_text_map, _pipe = core.extract_pages_text(pdf_path, target_pages, do_ocr_fallback=False)
        except Exception:
            page_text_map = {}

        # Group slicing (simple): only search after first occurrence of group_after and before first group_before
        ga = getattr(spec, 'group_after', None)
        gb = getattr(spec, 'group_before', None)

        for p in target_pages:
            raw = page_text_map.get(p, '') or ''
            if not raw:
                continue
            start_idx = 0
            end_idx = len(raw)
            if ga:
                hay = raw if case_sensitive else raw.lower()
                needle = ga if case_sensitive else ga.lower()
                k = hay.find(needle)
                if k != -1:
                    start_idx = k + len(ga)
                else:
                    continue  # group_after not seen on this page
            if gb:
                hay = raw if case_sensitive else raw.lower()
                needle = gb if case_sensitive else gb.lower()
                k = hay.find(needle, start_idx)
                if k != -1:
                    end_idx = min(end_idx, k)
            segment = raw[start_idx:end_idx]
            val = _same_line_right_value(segment, anchor_text, ret_is_string=ret_is_string, case_sensitive=case_sensitive)
            if val is None:
                continue
            if ret_is_string:
                return core.MatchResult(
                    pdf_file=Path(pdf_path).name,
                    serial_number=serial_number,
                    term=getattr(spec, 'term', ''),
                    page=p,
                    number=val,
                    units=None,
                    context=segment[:200],
                    method="text:nearest-line",
                    found=True,
                    confidence=None,
                    row_label=None,
                    column_label=None,
                    text_source="pdf",
                )
            else:
                num_txt, units = val if isinstance(val, tuple) else (str(val), None)
                return core.MatchResult(
                    pdf_file=Path(pdf_path).name,
                    serial_number=serial_number,
                    term=getattr(spec, 'term', ''),
                    page=p,
                    number=num_txt,
                    units=units,
                    context=segment[:200],
                    method="text:nearest-line",
                    found=True,
                    confidence=None,
                    row_label=None,
                    column_label=None,
                    text_source="pdf",
                )

        # If nothing found on same line via text, return a clear failure (no cross-line/window search)
        return core.MatchResult(
            pdf_file=Path(pdf_path).name,
            serial_number=serial_number,
            term=getattr(spec, 'term', ''),
            page=None,
            number=None,
            units=None,
            context="",
            method="text:nearest-line",
            found=False,
            confidence=None,
            row_label=None,
            column_label=None,
            text_source="pdf",
            error_reason="No value to the right of anchor on same line",
        )

    core.scan_pdf_for_term_nearest = patched_scan_pdf_for_term_nearest  # type: ignore[attr-defined]


def main():
    if LOCAL.exists():
        core = _load_core_from(LOCAL)
        _apply_patches(core)
        if hasattr(core, "main"):
            return core.main()
        raise SystemExit("Scanner core loaded but has no main() entry point")
    elif LEGACY.exists():
        # Legacy fallback (no patching)
        sys.path.insert(0, str(ROOT))
        core = __import__("eidp_term_scanner")
        if hasattr(core, "main"):
            return core.main()
        raise SystemExit("Legacy scanner found but no main() entry point")
    else:
        raise SystemExit("No scanner core found in /Application. Expected 'eidp_term_scanner.core.py'.")


if __name__ == "__main__":
    main()
