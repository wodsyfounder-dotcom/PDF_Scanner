#!/usr/bin/env python3
"""
Scan values from EasyOCR-derived table CSV files using the selections in terms.xlsx/csv.

For table(xy) mode rows, this script:
 - Reads per-page table CSV files produced by build_table_from_tokens.py
 - Matches the row label (first column) to Term or Line, and the header to Column
 - Extracts the first numeric value in that cell, honoring optional Range and Units hints

Outputs:
 - Creates a new run folder under Product_Data_File/run_data/<timestamp>/by_pdf
   with one JSON per PDF (same shape as the main scanner by_pdf JSONs)
 - Updates Product_Data_File/EIDP_data.csv (merging with any existing file)

Usage:
  py -3.11 scripts/scan_from_tables.py \
     --terms user_inputs/terms.xlsx \
     --pdf-dir user_inputs/EIDP_Import_Docs \
     --spatial-dir Product_Data_File/easy_spatial \
     --out-dir Product_Data_File/run_data
"""
from __future__ import annotations

import argparse
import csv
import json
import os
import re
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple


def load_terms(path: Path) -> List[Dict[str, str]]:
    ext = path.suffix.lower()
    if ext == ".csv":
        with path.open(newline='', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            return [dict(row) for row in reader]
    # Try pandas for Excel first (more likely installed with scanner environment)
    try:
        import pandas as pd  # type: ignore
        df = pd.read_excel(path)
        return [dict(row) for _, row in df.iterrows()]
    except Exception:
        pass
    # Try openpyxl minimal read
    try:
        from openpyxl import load_workbook  # type: ignore
        wb = load_workbook(filename=str(path), data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return []
        headers = [str(h or '').strip() for h in rows[0]]
        out: List[Dict[str, str]] = []
        for r in rows[1:]:
            row_dict = {headers[i]: ('' if i >= len(r) or r[i] is None else str(r[i])) for i in range(len(headers))}
            out.append(row_dict)
        return out
    except Exception:
        pass
    raise SystemExit(f"[ERROR] Unable to read terms file: {path}. Install pandas or openpyxl, or provide a CSV.")


def parse_pages(s: str) -> List[int]:
    s = (s or '').strip()
    if not s:
        return []
    # Allow 1-3,5;7 etc.
    out = set()
    for part in re.split(r'[;,\s]+', s):
        part = part.strip()
        if not part:
            continue
        if '-' in part:
            a, b = part.split('-', 1)
            try:
                a = int(re.sub(r'\D', '', a))
                b = int(re.sub(r'\D', '', b))
            except Exception:
                continue
            if a > b:
                a, b = b, a
            for p in range(a, b+1):
                out.add(p)
        else:
            try:
                out.add(int(re.sub(r'\D', '', part)))
            except Exception:
                pass
    return sorted(out)


def first_number(s: str) -> Optional[str]:
    if s is None:
        return None
    unit_core = r"%|ppm|ppb|ms|s|sec|kg|g|mg|ug|lbm|lb|lbs|lbf|N|kN|mN|Ns|bar|mbar|Pa|kPa|MPa|psi|psia|psig|mm|cm|m|in|ft|K|degC|degF|C|F"
    rx = re.compile(rf"[-+]?((?:\d{{1,3}}(?:,\d{{3}})+)|\d+)(?:\.\d+)?(?:\s?(?:{unit_core}))?", re.IGNORECASE)
    m = rx.search(str(s))
    return m.group(0) if m else None


def numeric_only(s: Optional[str]) -> Optional[str]:
    if not s:
        return None
    m = re.search(r"[-+]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?(?:[eE][+-]?\d+)?", s)
    return m.group(0).replace(',', '') if m else None


def norm(s: Optional[str]) -> str:
    return re.sub(r"\s+", " ", (s or '').strip()).lower()


def collect_pdf_tables(spatial_dir: Path, pdf_stem: str) -> Dict[int, Path]:
    # Returns map page->table.(xlsx/csv) path if exists, preferring xlsx when present
    base = spatial_dir / pdf_stem.replace(' ', '_')
    out: Dict[int, Path] = {}
    if not base.exists():
        return out
    # Prefer xlsx
    for p in sorted(base.glob(f"*_page_*.table.xlsx")):
        m = re.search(r"_page_(\d+)\.table\.xlsx$", p.name)
        if not m:
            continue
        page = int(m.group(1))
        out[page] = p
    # Fallback csv
    for p in sorted(base.glob(f"*_page_*.table.csv")):
        m = re.search(r"_page_(\d+)\.table\.csv$", p.name)
        if not m:
            continue
        page = int(m.group(1))
        out.setdefault(page, p)
    return out


def read_table_any(path: Path) -> List[List[str]]:
    rows: List[List[str]] = []
    if path.suffix.lower() == '.csv':
        with path.open(newline='', encoding='utf-8') as f:
            for row in csv.reader(f):
                rows.append([c for c in row])
        return rows
    # xlsx: try pandas then openpyxl
    try:
        import pandas as pd  # type: ignore
        df = pd.read_excel(path, header=0)
        rows.append([str(c) for c in list(df.columns)])
        for _, r in df.iterrows():
            rows.append([str(v if v is not None else '') for v in list(r)])
        return rows
    except Exception:
        pass
    try:
        from openpyxl import load_workbook  # type: ignore
        wb = load_workbook(filename=str(path), data_only=True)
        ws = wb.active
        for r in ws.iter_rows(values_only=True):
            rows.append([str(v) if v is not None else '' for v in r])
        return rows
    except Exception:
        pass
    return rows


def scan_from_table(rows: List[List[str]], row_label: str, col_label_raw: str) -> Optional[str]:
    if not rows:
        return None
    headers = rows[0]
    col_alts = [s.strip() for s in re.split(r'[|/]', col_label_raw or '') if s.strip()] or [(col_label_raw or '').strip()]
    # find column index by matching any alt against header cells (case-insensitive, trim)
    col_idx = None
    for i, h in enumerate(headers):
        hh = norm(h)
        for alt in col_alts:
            if hh == norm(alt):
                col_idx = i
                break
        if col_idx is not None:
            break
    if col_idx is None:
        return None
    # find row by matching first column against the row label (Term or Line)
    target = norm(row_label)
    for r in rows[1:]:
        if not r:
            continue
        row_name = norm(r[0]) if len(r) > 0 else ''
        if row_name == target:
            cell = r[col_idx] if col_idx < len(r) else ''
            return first_number(cell)
    return None


def update_eidp_aggregate(eidp_csv: Path, results: Dict[str, Dict[str, Optional[str]]]) -> None:
    # results: term -> {SN -> number}
    existing_rows: Dict[str, Dict[str, str]] = {}
    existing_cols: List[str] = []
    if eidp_csv.exists():
        with eidp_csv.open(newline='', encoding='utf-8') as f:
            reader = csv.reader(f)
            all_rows = list(reader)
        if all_rows:
            headers = all_rows[0]
            existing_cols = headers[1:]
            for r in all_rows[1:]:
                term = r[0]
                row_map = {headers[i]: r[i] for i in range(1, min(len(headers), len(r)))}
                existing_rows[term] = row_map
    # merge
    for term, sn_map in results.items():
        row = existing_rows.get(term, {})
        for sn, val in sn_map.items():
            row[sn] = val or ''
            if sn not in existing_cols:
                existing_cols.append(sn)
        existing_rows[term] = row
    # write back
    headers = ['Term'] + existing_cols
    with eidp_csv.open('w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow(headers)
        for term in sorted(existing_rows.keys()):
            row_map = existing_rows[term]
            w.writerow([term] + [row_map.get(sn, '') for sn in existing_cols])


def main() -> None:
    ap = argparse.ArgumentParser(description="Scan values from EasyOCR table CSVs based on terms.xlsx/csv")
    ap.add_argument('--terms', required=True, help='Path to terms.xlsx or terms.csv')
    ap.add_argument('--pdf-dir', required=True, help='Folder with source PDFs')
    ap.add_argument('--spatial-dir', required=True, help='Folder with easy_spatial/<stem>/ files')
    ap.add_argument('--out-dir', required=True, help='Base run_data folder (runs will be created under here)')
    args = ap.parse_args()

    terms_raw = load_terms(Path(args.terms))
    # Build term specs
    terms: List[Dict[str, str]] = []
    for row in terms_raw:
        term = str(row.get('Term') or row.get('term') or '').strip()
        if not term:
            continue
        mode = str(row.get('Mode') or row.get('mode') or 'nearest').strip().lower()
        pages_str = str(row.get('Pages') or row.get('pages') or '').strip()
        line = str(row.get('Line') or row.get('line') or '').strip()
        column = str(row.get('Column') or row.get('column') or '').strip()
        rmin = str(row.get('Range (min)') or row.get('range (min)') or row.get('Range min') or '').strip()
        rmax = str(row.get('Range (max)') or row.get('range (max)') or row.get('Range max') or '').strip()
        terms.append({'term': term, 'mode': mode, 'pages': pages_str, 'line': line, 'column': column, 'rmin': rmin, 'rmax': rmax})

    pdf_dir = Path(args.pdf_dir)
    spatial_dir = Path(args.spatial_dir)
    out_base = Path(args.out_dir)
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    run_dir = out_base / ts
    by_pdf_dir = run_dir / 'by_pdf'
    by_pdf_dir.mkdir(parents=True, exist_ok=True)

    # Collect PDFs in folder
    pdfs = sorted(pdf_dir.glob('*.pdf'))

    # Aggregate results per term for EIDP_data.csv
    aggregate: Dict[str, Dict[str, Optional[str]]] = {}

    for pdf in pdfs:
        stem = pdf.stem
        m = re.search(r"(?i)\bSN\W*([A-Za-z0-9][A-Za-z0-9_\-]*)", stem)
        sn = f"SN {m.group(1)}" if m else f"SN {stem}"
        table_map = collect_pdf_tables(spatial_dir, stem)
        recs: List[Dict[str, object]] = []
        for t in terms:
            term = t['term']
            mode = t['mode']
            pages = parse_pages(t['pages'])
            if not pages:
                # If pages unspecified, consider any table files available
                pages = sorted(table_map.keys())
            found_val: Optional[str] = None
            found_page: Optional[int] = None
            ctx = ''
            if mode.startswith('table'):
                row_label = (t['line'] or term)
                col_label = t['column']
                for p in pages:
                    table_csv = table_map.get(p)
                    if not table_csv or not table_csv.exists():
                        continue
                    rows = read_table_any(table_csv)
                    val = scan_from_table(rows, row_label=row_label, col_label_raw=col_label)
                    if val:
                        # range filtering
                        ok = True
                        try:
                            v = float(numeric_only(val) or '')
                            if t['rmin']:
                                ok = ok and (v >= float(str(t['rmin']).replace(',', '')))
                            if t['rmax']:
                                ok = ok and (v <= float(str(t['rmax']).replace(',', '')))
                        except Exception:
                            pass
                        if ok:
                            found_val = val
                            found_page = p
                            ctx = f"row='{row_label}' col='{col_label}'"
                            break
            # other modes could be added here using lines.csv/cells.csv if needed
            recs.append({
                'pdf_file': pdf.name,
                'serial_number': sn,
                'term': term,
                'found': bool(found_val),
                'page': found_page,
                'number': found_val,
                'units': None,
                'context': ctx,
                'method_pipeline': 'easyocr:table(xy) from CSV',
                'mode': 'table(xy)' if mode.startswith('table') else mode,
                'pages_raw': t['pages'],
                'line': (t['line'] or None),
                'column': (t['column'] or None),
                'range_min': float(str(t['rmin']).replace(',', '')) if t['rmin'] else None,
                'range_max': float(str(t['rmax']).replace(',', '')) if t['rmax'] else None,
                'units_hint': [],
                'anchor': None,
                'field_index': None,
                'field_split': 'groups',
                'return_type': 'number',
            })
            # aggregate
            d = aggregate.setdefault(term, {})
            d[sn] = found_val
        # write per-pdf JSON
        out_json = by_pdf_dir / f"{pdf.stem}.json"
        out_json.write_text(json.dumps(recs, ensure_ascii=False, indent=2), encoding='utf-8')

    # Update top-level EIDP_data.csv
    eidp_csv = Path('Product_Data_File') / 'EIDP_data.csv'
    eidp_csv.parent.mkdir(parents=True, exist_ok=True)
    update_eidp_aggregate(eidp_csv, aggregate)

    print(f"[OK] Wrote run folder: {run_dir}")
    print(f"[OK] Updated: {eidp_csv}")


if __name__ == '__main__':
    main()
