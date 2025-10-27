#!/usr/bin/env python3
"""
Extract page tables from a PDF into an Excel workbook (one sheet per page).

Grid approach (position -> table):
- Get tokens with positions (PyMuPDF words; EasyOCR fallback per page).
- Normalize to [0..1] coordinates using page width/height.
- Cluster rows via baseline proximity (adaptive y tolerance from median line height).
- Detect column separators from large x-gaps aggregated across all rows (robust stats).
- Form column bands; assign tokens to the unique (row, col) cell; join cell tokens.
- Detect simple section header lines and propagate a Section column.
- Choose a header row as the row with widest column coverage; use those as headers.

Usage:
  python scripts/extract_page_tables.py --pdf path/to/file.pdf --pages "1,3-5" [--ocr] [--dpi 600] [--min-conf 0.4] [--out out.xlsx] [--emit-tokens]

Defaults write to Product_Data_File/tables/<pdf-stem>_tables.xlsx
"""
from __future__ import annotations

import argparse
from pathlib import Path
import math
import os
import re
import sys
from typing import Dict, List, Tuple, Optional

# Optional heavy deps are imported lazily where used


def parse_page_ranges(s: str) -> List[int]:
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
                for p in range(ai, bi+1):
                    pages.add(p)
            except Exception:
                continue
        else:
            try:
                pages.add(int(re.sub(r"\D", "", part)))
            except Exception:
                continue
    return sorted(pages)


_AERO_UNITS = (
    "%|ppm|ppb|ms|s|sec|kg|g|mg|ug|lb|lbm|lbf|lbs|"
    "N|kN|mN|Ns|bar|mbar|Pa|kPa|MPa|psi|psia|psig|"
    "mm|cm|m|in|ft|K|degC|degF|C|F"
)
NUMBER_REGEX = re.compile(rf"""
    (?<![A-Za-z0-9_.-])
    [-+]?(?:\d{{1,3}}(?:,\d{{3}})+|\d+)(?:\.\d+)?(?:[eE][+-]?\d+)?  # number
    (?:\s?(?:{_AERO_UNITS}))?                                          # optional unit
    (?![A-Za-z0-9_.-])
""", re.VERBOSE)


def numeric_only(v: Optional[str]) -> Optional[str]:
    if not v:
        return None
    m = re.search(r"[-+]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?(?:[eE][+-]?\d+)?", v)
    if not m:
        return v
    return m.group(0).replace(',', '')


def extract_units(v: Optional[str]) -> Optional[str]:
    if not v:
        return None
    unit_core = _AERO_UNITS
    m = re.search(r"(?:\s*(" + unit_core + r"))$", v.strip(), flags=re.IGNORECASE)
    return m.group(1) if m else None


def _words_from_pymupdf(pdf: Path, page_num: int) -> Tuple[List[Tuple[float,float,float,float,str,int]], Tuple[float,float]]:
    try:
        import fitz  # PyMuPDF
    except Exception as e:
        raise RuntimeError("PyMuPDF (fitz) is required") from e
    doc = fitz.open(str(pdf))
    try:
        if not (1 <= page_num <= doc.page_count):
            return [], (0.0, 0.0)
        page = doc.load_page(page_num - 1)
        words = page.get_text("words") or []
        out = []
        for w in words:
            try:
                x0, y0, x1, y1, text = float(w[0]), float(w[1]), float(w[2]), float(w[3]), str(w[4])
                ln = int(w[6]) if len(w) >= 7 else int(round(y0))
                if text.strip():
                    out.append((x0, y0, x1, y1, text.strip(), ln))
            except Exception:
                pass
        bbox = page.rect
        return out, (float(bbox.width), float(bbox.height))
    finally:
        try:
            doc.close()
        except Exception:
            pass


def _words_from_easyocr(pdf: Path, page_num: int, dpi: int, langs: List[str], min_conf: float) -> Tuple[List[Tuple[float,float,float,float,str,int]], Tuple[float,float]]:
    try:
        import fitz
    except Exception as e:
        return [], (0.0, 0.0)
    try:
        import easyocr  # type: ignore
    except Exception:
        return [], (0.0, 0.0)
    doc = fitz.open(str(pdf))
    try:
        if not (1 <= page_num <= doc.page_count):
            return [], (0.0, 0.0)
        page = doc.load_page(page_num - 1)
        pix = page.get_pixmap(dpi=max(200, min(800, dpi)))
        from PIL import Image
        import numpy as np  # EasyOCR expects numpy array or path
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img_np = np.array(img)
        rdr = easyocr.Reader(langs or ['en'], gpu=False, verbose=False)  # type: ignore
        res = rdr.readtext(img_np)  # [(bbox, text, conf), ...]
        out: List[Tuple[float,float,float,float,str,int]] = []
        for it in res:
            try:
                bbox, text, conf = it
                if conf is not None and conf < min_conf:
                    continue
                xs = [float(pt[0]) for pt in bbox]
                ys = [float(pt[1]) for pt in bbox]
                x0, y0, x1, y1 = min(xs), min(ys), max(xs), max(ys)
                if str(text).strip():
                    out.append((x0, y0, x1, y1, str(text).strip(), int(round(y0))))
            except Exception:
                pass
        return out, (float(pix.width), float(pix.height))
    finally:
        try:
            doc.close()
        except Exception:
            pass


def _normalize_tokens(words: List[Tuple[float,float,float,float,str,int]], page_w: float, page_h: float) -> List[Dict]:
    out = []
    pw = max(1.0, page_w or 1.0)
    ph = max(1.0, page_h or 1.0)
    for x0, y0, x1, y1, text, _ln in words:
        cx = (x0 + x1) / 2.0
        cy = (y0 + y1) / 2.0
        out.append({
            'x0': x0 / pw, 'y0': y0 / ph, 'x1': x1 / pw, 'y1': y1 / ph,
            'cx': cx / pw, 'cy': cy / ph,
            'w': max(0.0, (x1 - x0) / pw), 'h': max(0.0, (y1 - y0) / ph),
            'text': text,
        })
    return out


def _cluster_rows(tokens: List[Dict], y_tol: Optional[float] = None) -> List[List[Dict]]:
    if not tokens:
        return []
    toks = sorted(tokens, key=lambda t: t['cy'])
    # Adaptive tolerance from median token height
    import statistics as _stats
    hs = [t['h'] for t in toks if t['h'] > 0]
    med_h = _stats.median(hs) if hs else 0.012
    tol = y_tol if (y_tol is not None) else max(0.006, min(0.04, 1.5 * med_h))
    rows: List[List[Dict]] = []
    cur: List[Dict] = []
    last_y = None
    for t in toks:
        if last_y is None or abs(t['cy'] - last_y) <= tol:
            cur.append(t)
            last_y = t['cy'] if last_y is None else (last_y*0.7 + t['cy']*0.3)
        else:
            cur.sort(key=lambda x: x['cx'])
            rows.append(cur)
            cur = [t]
            last_y = t['cy']
    if cur:
        cur.sort(key=lambda x: x['cx'])
        rows.append(cur)
    return rows


def _infer_separators_grid(rows: List[List[Dict]]) -> List[float]:
    """Detect column separators from large gaps across all rows (normalized x in [0..1])."""
    gaps: List[Tuple[float, float]] = []  # (mid, width)
    for ws in rows:
        if len(ws) < 2:
            continue
        for i in range(len(ws) - 1):
            x1 = ws[i]['x1']
            x2 = ws[i+1]['x0']
            width = x2 - x1
            if width <= 0:
                continue
            mid = x1 + width/2.0
            gaps.append((mid, width))
    if not gaps:
        return []
    import statistics as _stats
    widths = [w for _, w in gaps]
    med = _stats.median(widths)
    try:
        q1 = _stats.quantiles(widths, n=4)[0]
        q3 = _stats.quantiles(widths, n=4)[2]
        iqr = max(0.0, q3 - q1)
    except Exception:
        iqr = med
    thresh = max(0.02, med + 0.5 * iqr)  # normalized threshold
    cand = [m for (m, w) in gaps if w >= thresh]
    if not cand:
        return []
    cand.sort()
    # merge nearby midpoints
    merged: List[float] = []
    tol = max(0.01, 0.5 * thresh)
    for x in cand:
        if not merged or abs(x - merged[-1]) > tol:
            merged.append(x)
        else:
            merged[-1] = (merged[-1] + x) / 2.0
    # keep separators that occur in multiple rows
    # approximate frequency by counting row crossings
    seps: List[float] = []
    for s in merged:
        count = 0
        for ws in rows:
            for i in range(len(ws) - 1):
                if ws[i]['x1'] <= s <= ws[i+1]['x0']:
                    count += 1
                    break
        if count >= max(2, int(len(rows) * 0.10)):
            seps.append(s)
    seps.sort()
    return seps


def _assign_columns_grid(rows: List[List[Dict]], seps: List[float]) -> Tuple[int, List[List[List[Dict]]]]:
    bounds = [0.0] + seps + [1.0]
    col_count = max(1, len(bounds)-1)
    table: List[List[List[Dict]]] = []  # rows -> columns -> tokens
    for ws in rows:
        cols = [[] for _ in range(col_count)]
        for t in ws:
            cx = t['cx']
            # find interval
            idx = 0
            while idx < len(bounds)-1 and not (bounds[idx] <= cx < bounds[idx+1]):
                idx += 1
            if idx >= col_count:
                idx = col_count - 1
            cols[idx].append(t)
        table.append(cols)
    return col_count, table


def _choose_header_row(table: List[List[List[Dict]]]) -> Optional[int]:
    best = None
    best_cov = -1
    for ri, cols in enumerate(table):
        cov = sum(1 for c in cols if c)
        if cov > best_cov:
            best_cov = cov
            best = ri
    return best


def _is_section_line(ws: List[Dict], page_w_norm: float) -> bool:
    if not ws:
        return False
    text = " ".join([w['text'] for w in ws])
    span = (ws[-1]['x1'] - ws[0]['x0'])
    if text.endswith(":") and len(text.split()) <= 8:
        return True
    if span / max(1.0, page_w_norm) > 0.7 and len(text.split()) <= 12:
        return True
    return False


def extract_tables_for_pages(pdf: Path, pages: List[int], use_ocr: bool, dpi: int, min_conf: float, langs: List[str], emit_tokens: bool=False) -> Dict[int, Dict]:
    """Build tables per page using block segmentation + header-anchored columns (fallback to global grid)."""
    out: Dict[int, Dict] = {}
    for p in pages:
        words, (pw, ph) = _words_from_pymupdf(pdf, p)
        if not words and use_ocr:
            words, (pw, ph) = _words_from_easyocr(pdf, p, dpi=dpi, langs=langs, min_conf=min_conf)
        tokens = _normalize_tokens(words, pw, ph)
        rows = _cluster_rows(tokens)

        # Segment by section lines
        blocks: List[Tuple[int, int, Optional[str]]] = []
        start = 0
        current_section: Optional[str] = None
        for ri, ws in enumerate(rows):
            if _is_section_line(ws, 1.0):
                if ri > start:
                    blocks.append((start, ri, current_section))
                current_section = " ".join(w['text'] for w in ws).rstrip(':').strip()
                start = ri + 1
        if start < len(rows):
            blocks.append((start, len(rows), current_section))

        page_headers: List[str] = []
        page_rows: List[List[str]] = []

        for a, b, section in blocks:
            sub_rows = rows[a:b]
            if not sub_rows:
                continue
            # Global separators as baseline
            init_seps = _infer_separators_grid(sub_rows)
            _, init_table = _assign_columns_grid(sub_rows, init_seps)
            header_idx = _choose_header_row(init_table)

            block_seps: List[float] = []
            if header_idx is not None and 0 <= header_idx < len(sub_rows):
                header_tokens = sorted(sub_rows[header_idx], key=lambda t: t['x0'])
                gaps: List[Tuple[float, float]] = []
                for i in range(len(header_tokens)-1):
                    x1 = header_tokens[i]['x1']
                    x2 = header_tokens[i+1]['x0']
                    gaps.append(((x1+x2)/2.0, x2-x1))
                if gaps:
                    import statistics as _stats
                    widths = [w for _, w in gaps]
                    med = _stats.median(widths)
                    try:
                        q1 = _stats.quantiles(widths, n=4)[0]
                        q3 = _stats.quantiles(widths, n=4)[2]
                        iqr = max(0.0, q3-q1)
                    except Exception:
                        iqr = med
                    thr = max(0.01, med + 0.5*iqr)
                    mids = [m for (m,w) in gaps if w >= thr]
                    mids.sort()
                    for m in mids:
                        if not block_seps or abs(m - block_seps[-1]) > 0.01:
                            block_seps.append(m)
                        else:
                            block_seps[-1] = (block_seps[-1] + m)/2.0
            if not block_seps:
                block_seps = init_seps

            col_count, table = _assign_columns_grid(sub_rows, block_seps)
            headers = [f"Column_{i+1}" for i in range(col_count)]
            for ri, cols in enumerate(table):
                if header_idx is not None and ri == header_idx:
                    headers = [" ".join(t['text'] for t in sorted(c, key=lambda t: t['cx'])) or f"Column_{i+1}" for i, c in enumerate(cols)]
                    continue
                cells = [" ".join(t['text'] for t in sorted(c, key=lambda t: t['cx'])) for c in cols]
                page_rows.append([section or ""] + cells)

            if not page_headers:
                page_headers = ['Section'] + headers
            else:
                if len(page_headers) - 1 < len(headers):
                    extra = len(headers) - (len(page_headers) - 1)
                    page_headers += [f"Column_{len(page_headers)+i}" for i in range(1, extra+1)]

        if not page_headers:
            page_headers = ['Section', 'Column_1']
        spec = {'columns': page_headers, 'rows': page_rows}
        if emit_tokens:
            spec['tokens'] = [{'text': t['text'], 'x0': t['x0'], 'y0': t['y0'], 'x1': t['x1'], 'y1': t['y1']} for t in tokens]
        out[p] = spec
    return out


# ----------------------------
# Grid-geometry OCR extraction
# ----------------------------
def _render_page_image(pdf: Path, page_num: int, dpi: int) -> Optional["np.ndarray"]:
    try:
        import fitz  # PyMuPDF
        import numpy as np  # type: ignore
        from PIL import Image  # type: ignore
    except Exception:
        return None
    doc = fitz.open(str(pdf))
    try:
        if not (1 <= page_num <= doc.page_count):
            return None
        page = doc.load_page(page_num - 1)
        pix = page.get_pixmap(dpi=max(200, min(900, int(dpi) if dpi else 600)))
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        return np.array(img)
    finally:
        try:
            doc.close()
        except Exception:
            pass


def _detect_table_grid(img: "np.ndarray") -> Tuple[List[int], List[int], Tuple[int, int, int, int]]:
    """
    Detect vertical and horizontal table lines using OpenCV morphology.
    Returns (x_positions, y_positions, (x0,y0,x1,y1)) for the largest grid area.
    Coordinates are pixel indices in the page image.
    """
    import cv2  # type: ignore
    import numpy as np  # type: ignore

    gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
    # Adaptive threshold is robust to lighting/scan variations
    bw = cv2.adaptiveThreshold(~gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 15, -10)

    h, w = bw.shape[:2]
    # Kernel lengths relative to image size
    v_ker_len = max(10, w // 100)
    h_ker_len = max(10, h // 100)

    vert_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, v_ker_len))
    horiz_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (h_ker_len, 1))

    vertical = cv2.erode(bw, vert_kernel, iterations=1)
    vertical = cv2.dilate(vertical, vert_kernel, iterations=1)

    horizontal = cv2.erode(bw, horiz_kernel, iterations=1)
    horizontal = cv2.dilate(horizontal, horiz_kernel, iterations=1)

    # Combine to get table grid mask and find largest component as table area
    grid = cv2.addWeighted(vertical, 0.5, horizontal, 0.5, 0.0)
    grid = cv2.threshold(grid, 0, 255, cv2.THRESH_BINARY)[1]

    contours, _ = cv2.findContours(grid, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not contours:
        return [], [], (0, 0, w, h)
    areas = [cv2.contourArea(c) for c in contours]
    idx = int(np.argmax(np.array(areas)))
    x, y, ww, hh = cv2.boundingRect(contours[idx])
    roi = (x, y, x + ww, y + hh)

    # Crop masks to ROI to improve line position detection
    v_crop = vertical[y : y + hh, x : x + ww]
    h_crop = horizontal[y : y + hh, x : x + ww]

    # Project to find line runs, then take center of each run
    v_proj = v_crop.sum(axis=0)  # across rows -> per column sum
    h_proj = h_crop.sum(axis=1)  # across cols -> per row sum

    def _peaks_from_projection(arr: "np.ndarray", thresh_ratio: float = 0.5, min_gap: int = 8) -> List[int]:
        if arr.size == 0:
            return []
        m = float(arr.max())
        if m <= 0:
            return []
        thresh = m * thresh_ratio
        peaks: List[int] = []
        i = 0
        n = int(arr.shape[0])
        while i < n:
            if arr[i] >= thresh:
                j = i
                while j < n and arr[j] >= thresh:
                    j += 1
                peaks.append((i + j - 1) // 2)
                i = j + min_gap
            else:
                i += 1
        return peaks

    xs = _peaks_from_projection(v_proj, thresh_ratio=0.45, min_gap=max(6, w // 300))
    ys = _peaks_from_projection(h_proj, thresh_ratio=0.45, min_gap=max(6, h // 300))

    # Translate peaks back to absolute page coordinates
    xs = [x + int(px) for px in xs]
    ys = [y + int(py) for py in ys]
    xs = sorted(set(xs))
    ys = sorted(set(ys))

    # Keep only grids with at least 2x2 cells
    if len(xs) < 2 or len(ys) < 2:
        return [], [], roi
    return xs, ys, roi


def _ocr_table_from_grid(img: "np.ndarray", xs: List[int], ys: List[int], langs: List[str], min_conf: float) -> List[List[str]]:
    import numpy as np  # type: ignore
    try:
        import easyocr  # type: ignore
    except Exception as e:
        raise RuntimeError("EasyOCR is required for grid OCR mode") from e

    rdr = easyocr.Reader(langs or ["en"], gpu=False, verbose=False)  # type: ignore
    rows: List[List[str]] = []
    # Use centers between lines as cells; trim a few pixels to avoid borders
    for r in range(len(ys) - 1):
        row_vals: List[str] = []
        y0, y1 = ys[r], ys[r + 1]
        for c in range(len(xs) - 1):
            x0, x1 = xs[c], xs[c + 1]
            # add a small margin inside the cell
            pad = max(1, min((x1 - x0) // 50, (y1 - y0) // 50, 5))
            cy0 = max(0, y0 + pad)
            cy1 = max(cy0, y1 - pad)
            cx0 = max(0, x0 + pad)
            cx1 = max(cx0, x1 - pad)
            if cy1 <= cy0 or cx1 <= cx0:
                row_vals.append("")
                continue
            roi = img[cy0:cy1, cx0:cx1]
            if roi.size == 0:
                row_vals.append("")
                continue
            # Run OCR
            try:
                res = rdr.readtext(roi)  # [(bbox, text, conf), ...]
                texts = [str(t).strip() for (_b, t, conf) in res if (t and str(t).strip() and (conf is None or conf >= min_conf))]
                row_vals.append(" ".join(texts))
            except Exception:
                row_vals.append("")
        rows.append(row_vals)
    return rows


def extract_tables_via_grid_ocr(pdf: Path, pages: List[int], dpi: int, langs: List[str], min_conf: float, first_row_header: bool=False, drop_empty_rows: bool=True, drop_empty_cols: bool=True) -> Dict[int, Dict]:
    """
    Extract page tables by detecting ruled grid lines, then OCR each cell.
    Produces the same schema as extract_tables_for_pages: {'columns': [...], 'rows': [...]}.
    """
    try:
        import cv2  # noqa: F401  # type: ignore
        import numpy as np  # noqa: F401  # type: ignore
    except Exception as e:
        raise RuntimeError("OpenCV (cv2) and numpy are required for --ocr-grid mode") from e

    out: Dict[int, Dict] = {}
    for p in pages:
        img = _render_page_image(pdf, p, dpi=dpi)
        if img is None:
            out[p] = {'columns': ['Column_1'], 'rows': []}
            continue
        xs, ys, _roi = _detect_table_grid(img)
        if len(xs) < 2 or len(ys) < 2:
            # No obvious grid found on this page
            out[p] = {'columns': ['Column_1'], 'rows': []}
            continue
        table = _ocr_table_from_grid(img, xs, ys, langs=langs, min_conf=min_conf)

        # Drop fully empty rows/cols if requested
        if drop_empty_rows:
            table = [row for row in table if any((cell or '').strip() for cell in row)]
        if drop_empty_cols and table:
            keep_idx = [i for i in range(len(table[0])) if any((row[i] or '').strip() for row in table)]
            if keep_idx and len(keep_idx) < len(table[0]):
                table = [[row[i] for i in keep_idx] for row in table]

        if not table:
            out[p] = {'columns': ['Column_1'], 'rows': []}
            continue

        headers: List[str] = [f"Column_{i+1}" for i in range(len(table[0]))]
        rows_out: List[List[str]] = table
        if first_row_header and len(table) >= 1:
            headers = [cell.strip() or f"Column_{i+1}" for i, cell in enumerate(table[0])]
            rows_out = table[1:]

        out[p] = {'columns': headers, 'rows': rows_out}
    return out


def write_excel(tables: Dict[int, Dict], out_path: Path) -> None:
    try:
        import pandas as pd  # type: ignore
    except Exception as e:
        raise RuntimeError("pandas required to write Excel") from e
    try:
        import xlsxwriter  # noqa: F401
        engine = 'xlsxwriter'
    except Exception:
        try:
            import openpyxl  # noqa: F401
            engine = 'openpyxl'
        except Exception as e:
            raise RuntimeError("No Excel writer (xlsxwriter/openpyxl)") from e
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_path, engine=engine) as writer:
        for p in sorted(tables.keys()):
            spec = tables[p]
            df = pd.DataFrame(spec['rows'], columns=spec['columns'])
            sheet = f"p{p:03d}"
            df.to_excel(writer, sheet_name=sheet, index=False)
            try:
                ws = writer.sheets[sheet]
                ws.freeze_panes(1, 1)
                for i, col in enumerate(df.columns):
                    try:
                        max_len = int(df[col].astype(str).map(len).max()) if not df.empty else len(col)
                    except Exception:
                        max_len = len(col)
                    if engine == 'xlsxwriter':
                        ws.set_column(i, i, min(60, max(10, max_len + 2)))
            except Exception:
                pass
            # Optional tokens sheet for debugging/audit
            if 'tokens' in spec:
                try:
                    df_tok = pd.DataFrame(spec['tokens'])
                    df_tok.to_excel(writer, sheet_name=f"p{p:03d}_tokens", index=False)
                except Exception:
                    pass


def main() -> None:
    ap = argparse.ArgumentParser(description="Extract tables from pages and write one Excel workbook (sheets per page)")
    ap.add_argument('--pdf', required=True, help='Path to PDF')
    ap.add_argument('--pages', default='', help='Page list, e.g., "1,3-5" (1-indexed). Empty=all pages')
    ap.add_argument('--ocr', action='store_true', help='Force OCR fallback when PyMuPDF words are empty (gap/grid text mode)')
    ap.add_argument('--ocr-grid', action='store_true', help='Use OCR with line/box geometry to detect table grid (OpenCV + EasyOCR)')
    ap.add_argument('--dpi', type=int, default=600, help='OCR DPI when --ocr (default 600)')
    ap.add_argument('--min-conf', type=float, default=0.4, help='Min OCR confidence [0..1] when --ocr')
    ap.add_argument('--langs', default='en', help='OCR languages csv (for EasyOCR), default en')
    ap.add_argument('--out', default='', help='Output .xlsx path (default Product_Data_File/tables/<stem>_tables.xlsx)')
    ap.add_argument('--emit-tokens', action='store_true', help='Include raw tokens sheets for debugging')
    ap.add_argument('--grid-first-row-header', action='store_true', help='When using --ocr-grid, treat first detected row as headers')
    ap.add_argument('--grid-keep-blank-rows', action='store_true', help='When using --ocr-grid, keep fully blank rows')
    ap.add_argument('--grid-keep-blank-cols', action='store_true', help='When using --ocr-grid, keep fully blank columns')
    args = ap.parse_args()

    pdf = Path(args.pdf)
    if not pdf.exists():
        print(f"[ERROR] PDF not found: {pdf}", file=sys.stderr)
        sys.exit(1)
    if args.pages.strip():
        pages = parse_page_ranges(args.pages)
    else:
        try:
            import fitz
            with fitz.open(str(pdf)) as doc:
                pages = list(range(1, int(getattr(doc,'page_count',len(doc)))+1))
        except Exception:
            pages = []
    if not pages:
        print("[ERROR] No pages selected (check --pages)", file=sys.stderr)
        sys.exit(2)

    langs = [s.strip() for s in re.split(r'[;,]', args.langs) if s.strip()]
    if bool(args.ocr_grid):
        tables = extract_tables_via_grid_ocr(
            pdf, pages,
            dpi=args.dpi,
            langs=langs,
            min_conf=args.min_conf,
            first_row_header=bool(args.grid_first_row_header),
            drop_empty_rows=not bool(args.grid_keep_blank_rows),
            drop_empty_cols=not bool(args.grid_keep_blank_cols),
        )
        # If grid fails for a page, optionally fall back to text-gap mode
        if any((not v.get('rows')) for v in tables.values()):
            try:
                fallback = extract_tables_for_pages(pdf, pages, use_ocr=bool(args.ocr), dpi=args.dpi, min_conf=args.min_conf, langs=langs, emit_tokens=bool(args.emit_tokens))
                for p in pages:
                    if not tables.get(p, {}).get('rows'):
                        tables[p] = fallback.get(p, tables.get(p))
            except Exception:
                pass
    else:
        tables = extract_tables_for_pages(pdf, pages, use_ocr=bool(args.ocr), dpi=args.dpi, min_conf=args.min_conf, langs=langs, emit_tokens=bool(args.emit_tokens))
    out = Path(args.out) if args.out else (Path('Product_Data_File')/ 'tables' / f"{pdf.stem}_tables.xlsx")
    write_excel(tables, out)
    print(f"[DONE] Wrote page tables -> {out}")


if __name__ == '__main__':
    main()

