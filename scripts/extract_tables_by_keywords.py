#!/usr/bin/env python3
"""
Keyword-driven table extractor with geometry awareness.

Highlights
- Pulls words from PDF pages (PyMuPDF first; EasyOCR fallback for image-only pages).
- Finds line clusters, anchors on provided keywords, and grows a block of lines that
  looks like a table even if page headers/footers are present.
- Tries multiple strategies per block (multi-column gap clustering and 2-col split),
  scores them, and keeps the best matches.
- Writes all detected tables to a CSV (with comment markers between tables).

Usage:
  python scripts/extract_tables_by_keywords.py --pdf "Data Packages/FakeProgram_SV1_SN1111.pdf" ^
         --keywords "program,build revision" --pages 1-3 --out extracted.csv

Dependencies: fitz (PyMuPDF), easyocr, pillow, numpy
"""

from __future__ import annotations

import argparse
import csv
import math
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Optional, Sequence, Tuple
import re

# Optional heavy deps are loaded lazily
import fitz  # PyMuPDF

_EASYOCR_READER = None


def _get_easyocr_reader(langs: Optional[List[str]] = None):
    """Lazily create EasyOCR reader (cached)."""
    global _EASYOCR_READER
    if _EASYOCR_READER is not None:
        return _EASYOCR_READER
    try:
        import easyocr  # type: ignore
    except ImportError as exc:
        raise RuntimeError("easyocr is required for OCR fallback: pip install easyocr") from exc
    langs = langs or ["en"]
    _EASYOCR_READER = easyocr.Reader(langs, gpu=False, verbose=False)
    return _EASYOCR_READER


@dataclass
class Word:
    text: str
    x0: float
    y0: float
    x1: float
    y1: float

    @property
    def mid_x(self) -> float:
        return (self.x0 + self.x1) / 2.0

    @property
    def mid_y(self) -> float:
        return (self.y0 + self.y1) / 2.0


@dataclass
class Line:
    y: float
    words: List[Word]

    @property
    def text(self) -> str:
        return " ".join(w.text for w in self.words)


@dataclass
class TableCandidate:
    rows: List[List[str]]
    num_cols: int
    strategy: str
    page: int
    anchor_text: str
    score: float


def extract_words_from_page(pdf_path: Path, page_num: int, dpi: int = 300) -> Tuple[List[Word], Tuple[float, float], str]:
    """
    Extract words with bounding boxes. Try PyMuPDF text first, then OCR if empty.
    Returns: (words, (width, height), mode_label)
    """
    doc = fitz.open(pdf_path)
    try:
        page = doc.load_page(page_num - 1)
        width, height = page.rect.width, page.rect.height
        words_raw = page.get_text("words")
        if words_raw:
            words = [
                Word(text=w[4].strip(), x0=float(w[0]), y0=float(w[1]), x1=float(w[2]), y1=float(w[3]))
                for w in words_raw
                if w[4].strip()
            ]
            return words, (width, height), "pymupdf"
        # OCR fallback
        reader = _get_easyocr_reader()
        pix = page.get_pixmap(dpi=max(150, min(600, dpi)))
        import numpy as np
        from PIL import Image

        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        img_np = np.array(img)
        results = reader.readtext(img_np)
        words: List[Word] = []
        for bbox, text, conf in results:
            text = text.strip()
            if not text:
                continue
            xs = [pt[0] for pt in bbox]
            ys = [pt[1] for pt in bbox]
            words.append(
                Word(
                    text=text,
                    x0=min(x for x in xs if isinstance(x, (int, float))),
                    y0=min(y for y in ys if isinstance(y, (int, float))),
                    x1=max(x for x in xs if isinstance(x, (int, float))),
                    y1=max(y for y in ys if isinstance(y, (int, float))),
                )
            )
        return words, (pix.width, pix.height), "easyocr"
    finally:
        doc.close()


def group_words_to_lines(words: Sequence[Word], y_tol: float = 3.0) -> List[Line]:
    """Cluster words into lines using a Y tolerance."""
    lines: List[Line] = []
    for w in sorted(words, key=lambda w: (w.mid_y, w.x0)):
        placed = False
        for line in lines:
            if abs(line.y - w.mid_y) <= y_tol:
                line.words.append(w)
                placed = True
                break
        if not placed:
            lines.append(Line(y=w.mid_y, words=[w]))
    # sort each line left-to-right and compute unified y
    for line in lines:
        line.words.sort(key=lambda w: w.x0)
        line.y = sum(word.mid_y for word in line.words) / len(line.words)
    lines.sort(key=lambda l: l.y)
    return lines


def median(values: Sequence[float]) -> float:
    if not values:
        return 0.0
    vals = sorted(values)
    mid = len(vals) // 2
    if len(vals) % 2:
        return vals[mid]
    return (vals[mid - 1] + vals[mid]) / 2.0


def contains_any(text: str, keywords: Sequence[str]) -> bool:
    import re
    low = text.lower()
    for k in keywords:
        pattern = r"\b" + re.escape(k.lower()).replace(r"\ ", r"\s+") + r"\b"
        if re.search(pattern, low):
            return True
    return False


def find_anchor_indices(lines: Sequence[Line], keywords: Sequence[str]) -> List[int]:
    return [i for i, ln in enumerate(lines) if contains_any(ln.text, keywords)]


def build_block(
    lines: Sequence[Line],
    anchor_idx: int,
    base_gap: float,
    gap_factor: float = 1.8,
    min_block_lines: int = 1,
) -> Tuple[int, int]:
    """
    Expand from anchor line up/down while line gaps stay within a multiple of the typical spacing.
    This lets us skip over page headers/footers that have large vertical gaps.
    """
    gap_limit = base_gap * gap_factor if base_gap else 20.0
    start = anchor_idx
    while start > 0:
        gap = lines[start].y - lines[start - 1].y
        if gap > gap_limit:
            break
        start -= 1
    end = anchor_idx
    while end + 1 < len(lines):
        gap = lines[end + 1].y - lines[end].y
        if gap > gap_limit:
            break
        end += 1
    # Ensure minimum block length by extending downward if needed
    while end + 1 < len(lines) and (end - start + 1) < min_block_lines:
        end += 1
    return start, end


def cluster_split_points(points: Sequence[float], tol: float = 8.0) -> List[float]:
    """Cluster nearby split positions into averaged separators."""
    if not points:
        return []
    pts = sorted(points)
    clusters: List[List[float]] = [[pts[0]]]
    for p in pts[1:]:
        if abs(p - clusters[-1][-1]) <= tol:
            clusters[-1].append(p)
        else:
            clusters.append([p])
    return [sum(c) / len(c) for c in clusters]


def table_by_gap_clustering(lines: Sequence[Line], max_cols: int = 6) -> Optional[Tuple[List[List[str]], int]]:
    """Detect column splits from large intra-line gaps."""
    gap_mids: List[float] = []
    gaps: List[float] = []
    for ln in lines:
        ws = ln.words
        for left, right in zip(ws, ws[1:]):
            gap = right.x0 - left.x1
            gaps.append(gap)
            gap_mids.append(left.x1 + gap / 2.0)
    if not gaps:
        return None
    gaps_sorted = sorted(gaps)
    thresh_idx = max(0, int(len(gaps_sorted) * 0.60) - 1)
    gap_threshold = max(10.0, gaps_sorted[thresh_idx])
    splits = cluster_split_points(
        [mid for gap, mid in zip(gaps, gap_mids) if gap >= gap_threshold]
    )
    if not splits:
        return None
    splits = sorted(splits)[: max_cols - 1]
    num_cols = len(splits) + 1

    rows: List[List[str]] = []
    for ln in lines:
        cells = ["" for _ in range(num_cols)]
        for w in ln.words:
            col = 0
            while col < len(splits) and w.mid_x > splits[col]:
                col += 1
            cells[col] = (cells[col] + " " + w.text).strip()
        if any(cells):
            rows.append(cells)
    return rows, num_cols


def table_two_column(lines: Sequence[Line]) -> Optional[Tuple[List[List[str]], int]]:
    """Simple 2-col split using dominant vertical gap or mid-page split."""
    gaps: List[float] = []
    mids: List[float] = []
    for ln in lines:
        ws = ln.words
        for l, r in zip(ws, ws[1:]):
            gaps.append(r.x0 - l.x1)
            mids.append(l.x1 + (r.x0 - l.x1) / 2.0)
    split_x = None
    if gaps:
        gaps_sorted = sorted(gaps)
        thresh = max(10.0, gaps_sorted[int(len(gaps_sorted) * 0.7)])
        big_midpoints = [m for g, m in zip(gaps, mids) if g >= thresh]
        if big_midpoints:
            split_x = median(big_midpoints)
    if split_x is None:
        # fall back to page midpoint using all words
        all_x = [w.mid_x for ln in lines for w in ln.words]
        if not all_x:
            return None
        split_x = (min(all_x) + max(all_x)) / 2.0
    rows: List[List[str]] = []
    for ln in lines:
        left_parts: List[str] = []
        right_parts: List[str] = []
        for w in ln.words:
            if w.mid_x <= split_x:
                left_parts.append(w.text)
            else:
                right_parts.append(w.text)
        if left_parts or right_parts:
            rows.append([" ".join(left_parts).strip(), " ".join(right_parts).strip()])
    return rows, 2


def table_from_keyword_columns(lines: Sequence[Line]) -> Optional[Tuple[List[List[str]], int]]:
    """
    Build columns using the X positions of known header keywords, then assign all
    words to the nearest header center. Works well for multi-row headers.
    """
    header_terms = ["term", "description", "requirement", "value", "units", "page", "notes", "quality", "log"]
    centers: List[float] = []
    for ln in lines:
        for w in ln.words:
            low = w.text.lower()
            for ht in header_terms:
                if ht in low:
                    centers.append(w.mid_x)
    centers = sorted(set(centers))
    if len(centers) < 2:
        return None
    num_cols = len(centers)
    rows: List[List[str]] = []
    for ln in lines:
        row = ["" for _ in range(num_cols)]
        for w in ln.words:
            # Nearest center
            col = min(range(num_cols), key=lambda i: abs(w.mid_x - centers[i]))
            row[col] = (row[col] + " " + w.text).strip()
        if any(cell for cell in row):
            rows.append(row)
    return rows, num_cols


def score_table(rows: Sequence[Sequence[str]], keywords: Sequence[str]) -> float:
    """Heuristic score combining keyword coverage and table size."""
    flat = " ".join(" ".join(r) for r in rows).lower()
    coverage = sum(1 for k in keywords if k.lower() in flat)
    cov_ratio = coverage / max(1, len(keywords))
    row_score = min(len(rows), 15) / 15.0
    col_consistency = 1.0 if rows and all(len(r) == len(rows[0]) for r in rows) else 0.7
    return cov_ratio * 3.0 + row_score + col_consistency


def dedupe_empty_rows(rows: List[List[str]]) -> List[List[str]]:
    cleaned = []
    for r in rows:
        if not any(cell.strip() for cell in r):
            continue
        cleaned.append([cell.strip() for cell in r])
    return cleaned


def trim_sparse_columns(rows: List[List[str]], min_fill_ratio: float = 0.15) -> Tuple[List[List[str]], int]:
    """
    Drop columns that are almost entirely empty (useful when gap detection
    over-splits and leaves empty middle columns).
    """
    if not rows:
        return rows, 0
    col_count = max(len(r) for r in rows)
    non_empty = [0] * col_count
    for r in rows:
        for idx, val in enumerate(r):
            if val.strip():
                non_empty[idx] += 1
    keep_cols = [i for i, cnt in enumerate(non_empty) if cnt / len(rows) >= min_fill_ratio]
    if len(keep_cols) < 2 or len(keep_cols) == col_count:
        return rows, col_count
    trimmed = [[row[i] if i < len(row) else "" for i in keep_cols] for row in rows]
    return trimmed, len(keep_cols)


def coalesce_orphan_values(rows: List[List[str]]) -> List[List[str]]:
    """
    If a value gets split onto its own row (empty header, value populated),
    pull it into the nearest neighbor that has a header but empty value (prev or next).
    This is common in OCR when a single line wraps.
    """
    if not rows or any(len(r) != 2 for r in rows):
        return rows
    cleaned: List[List[str]] = []
    skip_next = False
    for i, row in enumerate(rows):
        if skip_next:
            skip_next = False
            continue
        # Case: current row has header but empty value, next row is orphan value
        if i + 1 < len(rows):
            next_row = rows[i + 1]
            if row[0] and not row[1] and not next_row[0] and next_row[1]:
                row = [row[0], next_row[1]]
                skip_next = True
        # Case: current row is orphan value, try to attach to prev or next header
        if not row[0] and row[1]:
            if cleaned and cleaned[-1][0] and not cleaned[-1][1]:
                cleaned[-1][1] = row[1]
                continue
            if i + 1 < len(rows) and rows[i + 1][0] and not rows[i + 1][1]:
                rows[i + 1] = [rows[i + 1][0], row[1]]
                continue
        cleaned.append(row)
    return cleaned


def merge_header_continuations(rows: List[List[str]]) -> List[List[str]]:
    """
    If a header with a value is immediately followed by a header-only row
    (no value), merge the headers to capture split labels like "Serial / Component".
    """
    if not rows or any(len(r) != 2 for r in rows):
        return rows
    merged: List[List[str]] = []
    skip = False
    for i, row in enumerate(rows):
        if skip:
            skip = False
            continue
        if i + 1 < len(rows):
            nxt = rows[i + 1]
            if row[0] and row[1] and nxt[0] and not nxt[1]:
                row = [f"{row[0]} / {nxt[0]}".strip(), row[1]]
                skip = True
        merged.append(row)
    return merged


def merge_wrapped_rows(rows: List[List[str]]) -> List[List[str]]:
    """
    If a row has an empty first column but text in others, treat it as a continuation
    of the previous row (when that previous row has a header/term in col 0):
    append text into the same columns of the previous row.
    This reduces fragmentation from OCR line breaks within a cell.
    """
    if not rows:
        return rows
    merged: List[List[str]] = []
    for row in rows:
        non_empty_cols = [i for i, v in enumerate(row) if v.strip()]
        if merged and row and not row[0].strip() and non_empty_cols and merged[-1][0].strip():
            prev = merged[-1]
            for idx in non_empty_cols:
                if idx >= len(prev):
                    continue
                if prev[idx].strip():
                    prev[idx] = f"{prev[idx]} {row[idx]}".strip()
                else:
                    prev[idx] = row[idx].strip()
            continue
        merged.append(row)
    return merged


def collapse_data_rows(rows: List[List[str]]) -> List[List[str]]:
    """
    Merge continuation rows into the nearest data row.
    - If buffered rows appear before the data row (empty col0), attach them to the next data row.
    - If buffered rows appear after the data row, attach to the current data row.
    - If a row starts with a numeric token and the previous row has a text label in col0 but
      is missing data cells, merge numeric row into the previous one.
    """
    if not rows:
        return rows
    header_tokens = {"term", "field", "attachment", "parameter"}
    collapsed: List[List[str]] = []
    buffer: List[List[str]] = []
    num_re = r"^[0-9+\\-\\.]"
    i = 0
    while i < len(rows):
        row = rows[i]
        first = row[0].strip().lower() if row else ""
        if first and first not in header_tokens:
            # Apply any buffered continuation rows that came before this data row
            if buffer:
                for cont in buffer:
                    for idx, val in enumerate(cont):
                        if not val.strip():
                            continue
                        if idx < len(row):
                            row[idx] = (row[idx] + " " + val).strip() if row[idx].strip() else val.strip()
                buffer = []
            j = i + 1
            while j < len(rows):
                cont_row = rows[j]
                cont_first = cont_row[0].strip()
                # merge if empty first col, or numeric first col and current row still sparse
                should_merge = False
                if not cont_first:
                    should_merge = True
                elif re.match(num_re, cont_first) and sum(1 for v in row if v.strip()) < len(row):
                    should_merge = True
                if not should_merge:
                    break
                for idx, val in enumerate(cont_row):
                    if not val.strip():
                        continue
                    if idx < len(row):
                        row[idx] = (row[idx] + " " + val).strip() if row[idx].strip() else val.strip()
                j += 1
            collapsed.append(row)
            i = j
        else:
            # Header or empty first column: buffer for potential next data row
            if not first and any(v.strip() for v in row):
                buffer.append(row)
            else:
                collapsed.append(row)
            i += 1
    # Flush any remaining buffer as separate rows
    collapsed.extend(buffer)
    return collapsed


# =============================================================================
# KPI-specific normalization
# =============================================================================

_NUM_RE = re.compile(r"^[+-]?\\d+(?:\\.\\d+)?$")


def _is_numeric(token: str) -> bool:
    return bool(_NUM_RE.match(token))


def extract_kpi_rows_from_lines(lines: Sequence[Line]) -> Optional[List[List[str]]]:
    """
    Specialized parser for the Performance KPI table.
    Returns rows with columns: KPI, Min, Target, Max, Value, Source, Confidence
    using header word positions to derive column boundaries.
    """
    if not lines:
        return None
    header_order = ["KPI", "Min", "Target", "Max", "Value", "Source", "Confidence"]
    header_idx = None
    for i, ln in enumerate(lines):
        low = ln.text.lower()
        if "kpi" in low and "confidence" in low:
            header_idx = i
            break
    if header_idx is None:
        return None

    # Collect header centers from header line and the immediate next line (captures "Target")
    centers: dict[str, float] = {}
    search_lines = lines[header_idx : min(len(lines), header_idx + 2)]
    for ln in search_lines:
        for w in ln.words:
            low = w.text.lower()
            for h in header_order:
                if h.lower() in low:
                    centers[h] = w.mid_x
    if centers:
        xs = sorted(centers.values())
        if xs:
            min_x, max_x = min(xs), max(xs)
            for idx, h in enumerate(header_order):
                if h not in centers:
                    centers[h] = min_x + (max_x - min_x) * idx / max(1, len(header_order) - 1)
    if len(centers) < 3:
        return None

    ordered_centers = [centers[h] for h in header_order]
    bounds: List[float] = []
    for a, b in zip(ordered_centers, ordered_centers[1:]):
        bounds.append((a + b) / 2.0)

    def assign_col(x: float) -> int:
        for idx, b in enumerate(bounds):
            if x < b:
                return idx
        return len(header_order) - 1

    data_lines: List[Line] = []
    for ln in lines[header_idx + 1 :]:
        t = ln.text.strip().lower()
        if not t:
            continue
        if t == "target":
            continue  # secondary header line
        if t.startswith("use this table") or re.match(r"^\d+\.\s", t):
            break
        data_lines.append(ln)

    if not data_lines:
        return None

    rows: List[List[str]] = []
    current_name: Optional[str] = None
    for ln in data_lines:
        row = ["" for _ in header_order]
        tokens = [w.text for w in ln.words]
        first = tokens[0] if tokens else ""
        has_name = any(ch.isalpha() for ch in first)
        for w in ln.words:
            col = assign_col(w.mid_x)
            row[col] = (row[col] + " " + w.text).strip()
        if has_name and row[0].strip():
            current_name = row[0].strip()
        if not has_name and current_name:
            row[0] = current_name
        if any(cell.strip() for cell in row):
            rows.append(row)

    # Merge continuation rows where KPI column is empty
    merged: List[List[str]] = []
    for r in rows:
        if merged and not r[0].strip():
            prev = merged[-1]
            for i, val in enumerate(r):
                if not val.strip():
                    continue
                prev[i] = (prev[i] + " " + val).strip() if prev[i].strip() else val.strip()
        else:
            merged.append(r)

    # Drop any stray header-like rows
    merged = [r for r in merged if "kpi" not in r[0].lower()]

    # Drop empty KPI rows and merge duplicate KPI name-only rows into following row
    cleaned: List[List[str]] = []
    i = 0
    while i < len(merged):
        r = merged[i]
        value_filled = any(cell.strip() for cell in r[1:])
        if not value_filled:
            if i + 1 < len(merged) and merged[i + 1][0].strip().lower() == r[0].strip().lower():
                i += 1
                continue
        cleaned.append(r)
        i += 1

    cleaned = [r for r in cleaned if any(cell.strip() for cell in r[1:])]

    return cleaned if cleaned else None


def best_table_for_block(lines: Sequence[Line], page_num: int, anchor_text: str, keywords: Sequence[str], max_cols: int) -> Optional[TableCandidate]:
    candidates: List[TableCandidate] = []
    # Remove obvious section headers/footers from the block
    filtered = []
    for ln in lines:
        if re.match(r"^\s*\d+\.", ln.text):
            continue
        if re.match(r"^use this table", ln.text, re.IGNORECASE):
            continue
        filtered.append(ln)
    if filtered:
        lines = filtered

    # KPI-specific extraction (highest priority when keyword contains KPI)
    if any("kpi" in k.lower() for k in keywords):
        kpi_rows = extract_kpi_rows_from_lines(lines)
        if kpi_rows:
            candidates.append(
                TableCandidate(
                    rows=kpi_rows,
                    num_cols=7,
                    strategy="kpi-special",
                    page=page_num,
                    anchor_text=anchor_text,
                    score=6.0,  # strong bias to prefer this
                )
            )
    for strategy_fn, label in [
        (table_from_keyword_columns, "keyword-cols"),
        (table_by_gap_clustering, "gap-cluster"),
        (table_two_column, "two-col"),
    ]:
        if label == "gap-cluster":
            result = strategy_fn(lines, max_cols)
        else:
            result = strategy_fn(lines)
        if not result:
            continue
        rows, num_cols = result
        rows = dedupe_empty_rows(rows)
        rows, num_cols = trim_sparse_columns(rows)
        rows = coalesce_orphan_values(rows)
        rows = merge_header_continuations(rows)
        rows = merge_wrapped_rows(rows)
        rows = collapse_data_rows(rows)
        if not rows:
            continue
        if len(rows) < 2:
            continue
        # Enforce presence of critical keyword if provided
        if any(k.lower() == "kpi" for k in keywords):
            flat = " ".join(" ".join(r).lower() for r in rows)
            if "kpi" not in flat:
                continue
        # Drop low-column tables when requirements-style keywords are requested
        if any(k.lower() == "requirement" for k in keywords) and num_cols < 3:
            continue
        score = score_table(rows, keywords)
        if label == "keyword-cols":
            score += 1.5  # prefer explicit header-aligned grids
        candidates.append(
            TableCandidate(
                rows=rows,
                num_cols=num_cols,
                strategy=label,
                page=page_num,
                anchor_text=anchor_text,
                score=score,
            )
        )
    if not candidates:
        return None
    return max(candidates, key=lambda c: c.score)


def process_page(pdf_path: Path, page_num: int, keywords: Sequence[str], max_cols: int, ocr_dpi: int, debug: bool) -> List[TableCandidate]:
    # Expand expected columns for requirement-style tables
    if any(k.lower() == "requirement" for k in keywords):
        max_cols = max(max_cols, 8)
    words, (w, h), mode = extract_words_from_page(pdf_path, page_num, dpi=ocr_dpi)
    if debug:
        print(f"[DEBUG] Page {page_num}: mode={mode}, words={len(words)}, size=({w:.1f},{h:.1f})")
    if not words:
        return []
    lines = group_words_to_lines(words)
    if debug:
        for i, ln in enumerate(lines[:40]):
            print(f"[DEBUG] L{i:02d} y={ln.y:.1f} text={ln.text}")
    y_gaps = [lines[i + 1].y - lines[i].y for i in range(len(lines) - 1)]
    base_gap = median(y_gaps) if y_gaps else 12.0
    anchors = find_anchor_indices(lines, keywords)
    tables: List[TableCandidate] = []
    for idx in anchors:
        start, end = build_block(lines, idx, base_gap, min_block_lines=12)
        block_lines = lines[start : end + 1]
        best = best_table_for_block(block_lines, page_num, lines[idx].text, keywords, max_cols)
        if best:
            tables.append(best)
    return tables


def merge_cross_page(tables: List[TableCandidate], gap_page_limit: int = 1) -> List[TableCandidate]:
    """
    Naive merge: if consecutive tables have same column count and consecutive pages,
    append their rows to form a continuation.
    """
    if not tables:
        return []
    tables.sort(key=lambda t: (t.page, -t.score))
    merged: List[TableCandidate] = []
    for tbl in tables:
        if merged:
            last = merged[-1]
            # If this is another candidate on the same page, keep only the higher score
            if tbl.page == last.page and tbl.num_cols == last.num_cols:
                if tbl.score > last.score:
                    merged[-1] = tbl
                continue
            # Merge only across consecutive pages for continuations
            if tbl.page - last.page == 1 and tbl.page - last.page <= gap_page_limit and tbl.num_cols == last.num_cols:
                last.rows.extend(tbl.rows)
                last.score = max(last.score, tbl.score)
                last.anchor_text += f" | p{tbl.page}"
                continue
        merged.append(tbl)
    return merged


def write_tables_to_csv(tables: Sequence[TableCandidate], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        for idx, tbl in enumerate(tables, 1):
            writer.writerow([f"# Table {idx} (page {tbl.page}, {tbl.strategy}, score={tbl.score:.2f})"])
            for row in tbl.rows:
                writer.writerow(row)
            writer.writerow([])


def parse_page_ranges(pages_str: str, total_pages: int) -> List[int]:
    if not pages_str:
        return list(range(1, total_pages + 1))
    pages: set[int] = set()
    for part in pages_str.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            start, end = part.split("-", 1)
            try:
                a, b = int(start), int(end)
            except ValueError:
                continue
            if a > b:
                a, b = b, a
            for p in range(a, b + 1):
                pages.add(p)
        else:
            try:
                pages.add(int(part))
            except ValueError:
                continue
    return sorted(p for p in pages if 1 <= p <= total_pages)


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract tables by keywords using geometry-aware parsing")
    parser.add_argument("--pdf", required=True, help="Path to PDF file")
    parser.add_argument("--keywords", required=True, help="Comma-separated keywords to anchor on")
    parser.add_argument("--pages", default="", help="Pages to scan (e.g., '1-3,5'); default all")
    parser.add_argument("--out", default="", help="Output CSV path (default: Data Packages/<pdf-stem>_keywords.csv)")
    parser.add_argument("--max-cols", type=int, default=6, help="Maximum expected columns")
    parser.add_argument("--ocr-dpi", type=int, default=300, help="DPI for OCR fallback rendering")
    parser.add_argument("--debug", action="store_true", help="Verbose debug logging")
    parser.add_argument("--best-only", action="store_true", help="Write only the single best-matching table")
    args = parser.parse_args()

    pdf_path = Path(args.pdf)
    if not pdf_path.exists():
        raise SystemExit(f"PDF not found: {pdf_path}")
    keywords = [k.strip() for k in args.keywords.split(",") if k.strip()]
    if not keywords:
        raise SystemExit("No keywords provided")

    doc = fitz.open(pdf_path)
    total_pages = doc.page_count
    doc.close()
    pages = parse_page_ranges(args.pages, total_pages)
    if not pages:
        pages = list(range(1, total_pages + 1))

    all_candidates: List[TableCandidate] = []
    for p in pages:
        all_candidates.extend(process_page(pdf_path, p, keywords, args.max_cols, args.ocr_dpi, args.debug))

    merged = merge_cross_page(all_candidates)
    if args.best_only and merged:
        merged = [max(merged, key=lambda t: (t.score, len(t.rows), t.num_cols))]

    if not merged:
        print("[WARN] No tables matched the provided keywords")
    out_path = Path(args.out) if args.out else (pdf_path.parent / f"{pdf_path.stem}_keywords.csv")
    write_tables_to_csv(merged, out_path)
    print(f"[DONE] Wrote {len(merged)} table(s) to {out_path}")


if __name__ == "__main__":
    main()
