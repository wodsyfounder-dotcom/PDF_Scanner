#!/usr/bin/env python3
"""
Dump EasyOCR results with spatial awareness so you can inspect what OCR
"sees" and quickly reference structured tokens for downstream algorithms.

Outputs (per page):
- tokens JSON: [{x0,y0,x1,y1,cx,cy,w,h,conf,text} ...]
- tokens TSV:  id  page  cx  cy  w  h  conf  text
- rows TXT:    tokens grouped by row (y-cluster), joined left->right
- rows TSV:    row_id  page  cx  cy  text (joined)  token_count

Usage (from project root):
  py -3.11 scripts/easyocr_dump_spatial.py \
      --pdf user_inputs/EIDP_Import_Docs/sn\ 3232.pdf \
      --out-dir Product_Data_File/easy_spatial/sn_3232 \
      --dpi 800 --langs en --min-conf 0.0 --row-factor 0.6

Notes
- Pure Python (EasyOCR + PyMuPDF); no native executables required.
- Row clustering uses a y-gap threshold based on median token height.
"""
from __future__ import annotations

import argparse
import json
import statistics
from pathlib import Path
from typing import Dict, List

try:
    import fitz  # type: ignore  # PyMuPDF
except Exception as e:  # pragma: no cover
    raise SystemExit("[ERROR] PyMuPDF is required: pip install pymupdf") from e

try:
    import easyocr  # type: ignore
except Exception as e:  # pragma: no cover
    raise SystemExit("[ERROR] EasyOCR is required: pip install easyocr (and CPU torch)") from e


def ocr_page_tokens(pdf: Path, page_index: int, dpi: int, langs: List[str], min_conf: float) -> List[Dict[str, float]]:
    # Use context managers to avoid nested try/finally confusion
    with fitz.open(str(pdf)) as doc:  # type: ignore[name-defined]
        if not (0 <= page_index < doc.page_count):
            return []
        page = doc.load_page(page_index)
        pix = page.get_pixmap(dpi=dpi)
        import tempfile
        with tempfile.TemporaryDirectory(prefix="easyocr_dump_") as tmpdir:
            img_path = Path(tmpdir) / f"p{page_index+1}.png"
            pix.save(str(img_path))
            reader = easyocr.Reader(langs or ["en"], gpu=False, verbose=False)  # type: ignore
            res = reader.readtext(str(img_path), detail=1)  # list of [bbox, text, conf]
            tokens: List[Dict[str, float]] = []
            for it in res:
                try:
                    bbox, text, conf = it
                    if not isinstance(text, str) or not text.strip():
                        continue
                    if conf is None:
                        conf_val = 0.0
                    else:
                        try:
                            conf_val = float(conf)
                        except Exception:
                            conf_val = 0.0
                    if conf_val < float(min_conf):
                        continue
                    xs = [float(pt[0]) for pt in bbox]
                    ys = [float(pt[1]) for pt in bbox]
                    x0, y0, x1, y1 = min(xs), min(ys), max(xs), max(ys)
                    cx = (x0 + x1) / 2.0
                    cy = (y0 + y1) / 2.0
                    tokens.append({
                        "x0": x0, "y0": y0, "x1": x1, "y1": y1,
                        "cx": cx, "cy": cy, "w": (x1 - x0), "h": (y1 - y0),
                        "conf": conf_val, "text": text.strip(),
                    })
                except Exception:
                    # Skip any malformed OCR items defensively
                    continue
            # Sort by (cy, cx)
            tokens.sort(key=lambda t: (t["cy"], t["cx"]))
            return tokens


def cluster_rows(tokens: List[Dict[str, float]], row_factor: float) -> List[List[Dict[str, float]]]:
    if not tokens:
        return []
    heights = [t["h"] for t in tokens if t.get("h", 0) > 0]
    med_h = statistics.median(heights) if heights else 12.0
    thresh = max(2.0, row_factor * med_h)
    rows: List[List[Dict[str, float]]] = []
    current: List[Dict[str, float]] = []
    current_cy = None
    for t in tokens:
        cy = t["cy"]
        if current_cy is None:
            current = [t]
            current_cy = cy
        else:
            if abs(cy - current_cy) <= thresh:
                current.append(t)
                # update row center (running average)
                current_cy = (current_cy * (len(current) - 1) + cy) / len(current)
            else:
                rows.append(sorted(current, key=lambda x: x["cx"]))
                current = [t]
                current_cy = cy
    if current:
        rows.append(sorted(current, key=lambda x: x["cx"]))
    return rows


def write_outputs(base: Path, page: int, tokens: List[Dict[str, float]], rows: List[List[Dict[str, float]]]) -> None:
    base.parent.mkdir(parents=True, exist_ok=True)
    # JSON
    (base.with_suffix(".json")).write_text(json.dumps(tokens, ensure_ascii=False, indent=2), encoding="utf-8")
    # tokens TSV
    tsv = ["id\tpage\tcx\tcy\tw\th\tconf\ttext"]
    for i, t in enumerate(tokens, start=1):
        tsv.append("%d\t%d\t%.2f\t%.2f\t%.2f\t%.2f\t%.3f\t%s" % (
            i, page, t["cx"], t["cy"], t["w"], t["h"], t["conf"], t["text"].replace("\t"," ")
        ))
    (base.with_suffix(".tokens.tsv")).write_text("\n".join(tsv), encoding="utf-8")
    # rows TXT and TSV
    rows_txt = []
    rows_tsv = ["row_id\tpage\tcx\tcy\ttext\ttoken_count"]
    for r_id, row in enumerate(rows, start=1):
        joined = " ".join([t["text"] for t in row])
        cx = statistics.mean([t["cx"] for t in row]) if row else 0.0
        cy = statistics.mean([t["cy"] for t in row]) if row else 0.0
        rows_txt.append(joined)
        rows_tsv.append("%d\t%d\t%.2f\t%.2f\t%s\t%d" % (r_id, page, cx, cy, joined.replace("\t"," "), len(row)))
    (base.with_suffix(".rows.txt")).write_text("\n".join(rows_txt), encoding="utf-8")
    (base.with_suffix(".rows.tsv")).write_text("\n".join(rows_tsv), encoding="utf-8")


def main() -> None:
    ap = argparse.ArgumentParser(description="Dump EasyOCR tokens with spatial order and row clustering")
    ap.add_argument("--pdf", required=True, help="Path to input PDF")
    ap.add_argument("--out-dir", required=True, help="Output directory for dumps")
    ap.add_argument("--dpi", type=int, default=800, help="Render DPI (200..900)")
    ap.add_argument("--langs", default="en", help="Comma/semicolon list of languages (default en)")
    ap.add_argument("--min-conf", type=float, default=0.0, help="Min confidence to keep a token")
    ap.add_argument("--row-factor", type=float, default=0.6, help="Row cluster threshold as multiple of median token height")
    args = ap.parse_args()

    pdf = Path(args.pdf)
    out_dir = Path(args.out_dir)
    langs = [s.strip() for s in args.langs.replace(";", ",").split(",") if s.strip()]

    d = fitz.open(str(pdf))  # type: ignore[name-defined]
    try:
        for p in range(d.page_count):
            tokens = ocr_page_tokens(pdf, p, dpi=args.dpi, langs=langs, min_conf=args.min_conf)
            rows = cluster_rows(tokens, row_factor=args.row_factor)
            base = out_dir / (f"{pdf.stem}_page_{p+1}")
            write_outputs(base, p+1, tokens, rows)
        print(f"[DONE] Wrote EasyOCR spatial dumps -> {out_dir}")
    finally:
        try:
            d.close()
        except Exception:
            pass


if __name__ == "__main__":
    main()
