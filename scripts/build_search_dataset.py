#!/usr/bin/env python3
"""
Build a unified search dataset from EasyOCR spatial outputs and/or table builds.

Sources per page (preferred to minimal):
 - tokens JSON: produced by scripts/easyocr_dump_spatial.py
   <stem>_page_<n>.json
 - optional table CSV/XLSX: produced by scripts/build_table_from_tokens.py
   <stem>_page_<n>.table.csv / .xlsx

Outputs (written next to the input directory with a common prefix):
 - <prefix>.tokens.csv  : token-level records with geometry
 - <prefix>.lines.csv   : y-clustered lines with joined text and bounds
 - <prefix>.cells.csv   : reconstructed table cells (row, col, text, bounds)
 - <prefix>.search.jsonl: combined JSONL stream (type=token|line|cell)

Usage example:
  py -3.11 scripts/build_search_dataset.py \
     --dir Product_Data_File/easy_spatial/sn_4444 \
     --stem "sn 4444" \
     --out-prefix Product_Data_File/easy_spatial/sn_4444/sn_4444

Tuning:
  --row-factor 0.6            Row clustering threshold vs median token height
  --gap-mult 2.5              Gap multiplier vs median token width for cell segmentation
  --center-thresh-mult 2.0    Column center clustering threshold vs median token width
"""
from __future__ import annotations

import argparse
import csv
import json
import re
from pathlib import Path
from typing import Dict, List, Tuple, Any

Token = Dict[str, Any]


def median(values: List[float]) -> float:
    vals = sorted(values)
    n = len(vals)
    if n == 0:
        return 0.0
    mid = n // 2
    if n % 2:
        return float(vals[mid])
    return float((vals[mid - 1] + vals[mid]) / 2.0)


def load_page_tokens(json_path: Path) -> List[Token]:
    try:
        data = json.loads(json_path.read_text(encoding="utf-8"))
    except Exception as e:
        raise SystemExit(f"[ERROR] Failed reading tokens JSON {json_path}: {e}")
    if not isinstance(data, list):
        return []
    return data


def compute_stats(tokens: List[Token]) -> Tuple[float, float, float, float, float, float]:
    if not tokens:
        return 0.0, 0.0, 0.0, 0.0, 0.0, 0.0
    xs0 = [float(t["x0"]) for t in tokens]
    xs1 = [float(t["x1"]) for t in tokens]
    ys0 = [float(t["y0"]) for t in tokens]
    ys1 = [float(t["y1"]) for t in tokens]
    ws = [float(t.get("w", max(0.0, float(t["x1"]) - float(t["x0"])))) for t in tokens]
    hs = [float(t.get("h", max(0.0, float(t["y1"]) - float(t["y0"])))) for t in tokens]
    return min(xs0), max(xs1), min(ys0), max(ys1), median(ws) or 10.0, median(hs) or 12.0


def sort_tokens(tokens: List[Token]) -> List[Token]:
    return sorted(tokens, key=lambda t: (float(t["cy"]), float(t["cx"])))


def cluster_rows(tokens: List[Token], row_factor: float) -> List[List[Token]]:
    if not tokens:
        return []
    tokens_sorted = sort_tokens(tokens)
    med_h = median([float(t.get("h", 0.0)) for t in tokens if float(t.get("h", 0.0)) > 0.0]) or 12.0
    thresh = max(2.0, row_factor * med_h)
    rows: List[List[Token]] = []
    current: List[Token] = []
    current_cy = None
    for t in tokens_sorted:
        cy = float(t["cy"])
        if current_cy is None:
            current = [t]
            current_cy = cy
        else:
            if abs(cy - current_cy) <= thresh:
                current.append(t)
                current_cy = (current_cy * (len(current) - 1) + cy) / len(current)
            else:
                rows.append(sorted(current, key=lambda x: float(x["cx"])) )
                current = [t]
                current_cy = cy
    if current:
        rows.append(sorted(current, key=lambda x: float(x["cx"])) )
    return rows


def row_segments(row_tokens: List[Token], gap_mult: float, med_w: float) -> List[Tuple[float, List[Token]]]:
    # Horizontally split tokens into segments where gaps exceed threshold.
    if not row_tokens:
        return []
    segs: List[List[Token]] = []
    cur: List[Token] = [row_tokens[0]]
    for prev, t in zip(row_tokens, row_tokens[1:]):
        gap = float(t["x0"]) - float(prev["x1"])
        if gap > (gap_mult * med_w):
            segs.append(cur)
            cur = [t]
        else:
            cur.append(t)
    segs.append(cur)
    out: List[Tuple[float, List[Token]]] = []
    for seg in segs:
        center = sum(float(s["cx"]) for s in seg) / max(1, len(seg))
        out.append((center, seg))
    return out


def cluster_columns(centers: List[float], center_thresh: float) -> List[float]:
    if not centers:
        return []
    centers_sorted = sorted(centers)
    clusters: List[List[float]] = [[centers_sorted[0]]]
    for c in centers_sorted[1:]:
        if abs(c - clusters[-1][-1]) <= center_thresh:
            clusters[-1].append(c)
        else:
            clusters.append([c])
    return [sum(g)/len(g) for g in clusters]


def build_grid(rows: List[List[Token]], med_w: float, gap_mult: float, center_thresh_mult: float) -> Tuple[List[float], List[List[str]]]:
    # Determine stable column centers across rows from segments
    all_seg_centers: List[float] = []
    for row in rows:
        for c, _ in row_segments(row, gap_mult=gap_mult, med_w=med_w):
            all_seg_centers.append(c)
    if not all_seg_centers:
        return [], []
    center_thresh = max(2.0, center_thresh_mult * med_w)
    col_centers = cluster_columns(all_seg_centers, center_thresh=center_thresh)
    col_centers.sort()
    # Build text rows by assigning each row segment to nearest column
    table_rows: List[List[str]] = []
    for row in rows:
        segs = row_segments(row, gap_mult=gap_mult, med_w=med_w)
        cells = [""] * len(col_centers)
        for c, seg in segs:
            nearest_idx = min(range(len(col_centers)), key=lambda i: abs(col_centers[i] - c))
            if abs(col_centers[nearest_idx] - c) <= center_thresh:
                text = " ".join(str(s.get("text", "")).strip() for s in seg if str(s.get("text", "")).strip())
                if cells[nearest_idx]:
                    cells[nearest_idx] += " " + text
                else:
                    cells[nearest_idx] = text
        table_rows.append(cells)
    return col_centers, table_rows


def row_bounds(rows: List[List[Token]]) -> List[Tuple[float, float]]:
    spans: List[Tuple[float, float]] = []
    for row in rows:
        if not row:
            spans.append((0.0, 0.0))
            continue
        y0 = min(float(t["y0"]) for t in row)
        y1 = max(float(t["y1"]) for t in row)
        spans.append((y0, y1))
    return spans


def col_bounds(col_centers: List[float], min_x: float, max_x: float) -> List[float]:
    if not col_centers:
        return []
    b: List[float] = []
    b.append((min_x + col_centers[0]) / 2.0)
    for a, c in zip(col_centers, col_centers[1:]):
        b.append((a + c) / 2.0)
    b.append((col_centers[-1] + max_x) / 2.0)
    return b


def number_matches(s: str) -> List[str]:
    # Lightweight numeric pattern with optional unit suffix
    unit_core = r"%|ppm|ppb|ms|s|sec|kg|g|mg|ug|lbm|lb|lbs|lbf|N|kN|mN|Ns|bar|mbar|Pa|kPa|MPa|psi|psia|psig|mm|cm|m|in|ft|K|degC|degF|C|F"
    rx = re.compile(rf"[-+]?((?:\d{{1,3}}(?:,\d{{3}})+)|\d+)(?:\.\d+)?(?:\s?(?:{unit_core}))?", re.IGNORECASE)
    return [m.group(0) for m in rx.finditer(s or "")]


def build_for_dir(base_dir: Path, stem: str, out_prefix: Path, row_factor: float, gap_mult: float, center_thresh_mult: float) -> None:
    # Discover pages by tokens JSON
    page_jsons = sorted(base_dir.glob(f"{stem}_page_*.json"))
    if not page_jsons:
        raise SystemExit(f"[ERROR] No tokens JSON found at {base_dir} for stem '{stem}'")

    # Prepare outputs
    tokens_csv = out_prefix.with_suffix(".tokens.csv")
    lines_csv = out_prefix.with_suffix(".lines.csv")
    cells_csv = out_prefix.with_suffix(".cells.csv")
    jsonl_path = out_prefix.with_suffix(".search.jsonl")

    with tokens_csv.open("w", newline="", encoding="utf-8") as f_tok, \
         lines_csv.open("w", newline="", encoding="utf-8") as f_lin, \
         cells_csv.open("w", newline="", encoding="utf-8") as f_cell, \
         jsonl_path.open("w", encoding="utf-8") as f_jsonl:

        wt = csv.writer(f_tok)
        wl = csv.writer(f_lin)
        wc = csv.writer(f_cell)
        wt.writerow(["serial","page","text","conf","x0","y0","x1","y1","cx","cy"]) 
        wl.writerow(["serial","page","row_index","text","y0","y1"]) 
        wc.writerow(["serial","page","row_index","col_index","text","x0","y0","x1","y1"]) 

        # Derive serial label from stem (e.g., "sn 4444" -> "SN 4444")
        serial_label = stem
        m = re.search(r"sn\W*([A-Za-z0-9][A-Za-z0-9_\-]*)", stem, flags=re.IGNORECASE)
        if m:
            serial_label = f"SN {m.group(1)}"

        for j in page_jsons:
            page_match = re.search(r"_page_(\d+)\.json$", j.name)
            page = int(page_match.group(1)) if page_match else 1
            tokens = load_page_tokens(j)
            # Write tokens
            for t in tokens:
                wt.writerow([serial_label, page, t.get("text",""), t.get("conf", ""), t.get("x0",""), t.get("y0",""), t.get("x1",""), t.get("y1",""), t.get("cx",""), t.get("cy","")])
                rec = {
                    "type": "token",
                    "serial": serial_label,
                    "page": page,
                    "text": t.get("text",""),
                    "conf": t.get("conf", None),
                    "bbox": [t.get("x0",0), t.get("y0",0), t.get("x1",0), t.get("y1",0)],
                    "center": [t.get("cx",0), t.get("cy",0)],
                }
                nums = number_matches(str(t.get("text","")))
                if nums:
                    rec["numbers"] = nums
                f_jsonl.write(json.dumps(rec, ensure_ascii=False) + "\n")

            if not tokens:
                continue

            # Lines (row clusters)
            rows = cluster_rows(tokens, row_factor=row_factor)
            spans = row_bounds(rows)
            for idx, (row_tokens, (y0, y1)) in enumerate(zip(rows, spans), start=1):
                text = " ".join(str(x.get("text","")) for x in row_tokens if str(x.get("text","")))
                wl.writerow([serial_label, page, idx, text, y0, y1])
                rec = {
                    "type": "line",
                    "serial": serial_label,
                    "page": page,
                    "row_index": idx,
                    "text": text,
                    "y0": y0,
                    "y1": y1,
                }
                nums = number_matches(text)
                if nums:
                    rec["numbers"] = nums
                f_jsonl.write(json.dumps(rec, ensure_ascii=False) + "\n")

            # Table cells
            min_x, max_x, _min_y, _max_y, med_w, _med_h = compute_stats(tokens)
            col_centers, table_rows = build_grid(rows, med_w=med_w, gap_mult=gap_mult, center_thresh_mult=center_thresh_mult)
            if col_centers and table_rows:
                bounds = col_bounds(col_centers, min_x=min_x, max_x=max_x)
                spans = row_bounds(rows)
                for r_idx, cells in enumerate(table_rows, start=1):
                    for c_idx, text in enumerate(cells, start=1):
                        if not bounds or not spans:
                            continue
                        if c_idx-1 >= len(bounds) or c_idx >= len(bounds) or r_idx-1 >= len(spans):
                            continue
                        x0 = bounds[c_idx-1]
                        x1 = bounds[c_idx]
                        y0, y1 = spans[r_idx-1]
                        wc.writerow([serial_label, page, r_idx, c_idx, text, x0, y0, x1, y1])
                        rec = {
                            "type": "cell",
                            "serial": serial_label,
                            "page": page,
                            "row_index": r_idx,
                            "col_index": c_idx,
                            "text": text,
                            "bbox": [x0, y0, x1, y1],
                        }
                        nums = number_matches(text or "")
                        if nums:
                            rec["numbers"] = nums
                        f_jsonl.write(json.dumps(rec, ensure_ascii=False) + "\n")


def main() -> None:
    ap = argparse.ArgumentParser(description="Build unified search dataset from spatial outputs (and table builds)")
    ap.add_argument("--dir", required=True, help="Directory containing per-page tokens JSON (and optional table files)")
    ap.add_argument("--stem", required=True, help="PDF stem used in page filenames (e.g., 'sn 4444')")
    ap.add_argument("--out-prefix", required=True, help="Output prefix for dataset (without extension)")
    ap.add_argument("--row-factor", type=float, default=0.6)
    ap.add_argument("--gap-mult", type=float, default=2.5)
    ap.add_argument("--center-thresh-mult", type=float, default=2.0)
    args = ap.parse_args()

    base_dir = Path(args.dir)
    out_prefix = Path(args.out_prefix)
    build_for_dir(base_dir, args.stem, out_prefix, args.row_factor, args.gap_mult, args.center_thresh_mult)
    print(f"[OK] Wrote dataset: {out_prefix.with_suffix('.tokens.csv')} | {out_prefix.with_suffix('.lines.csv')} | {out_prefix.with_suffix('.cells.csv')} | {out_prefix.with_suffix('.search.jsonl')}")


if __name__ == "__main__":
    main()

