#!/usr/bin/env python3
"""
Utility helpers to OCR a single PDF page and export the detected table-like
content into an Excel or CSV worksheet.

Example:
  python scripts/ocr_page_to_excel.py --pdf user_inputs/EIDP_Import_Docs/sn9911.pdf --page 1 --output out.xlsx
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import List, Sequence, Tuple


ROOT = Path(__file__).resolve().parents[1]
SITE_PACKAGES = ROOT / "Lib" / "site-packages"
if SITE_PACKAGES.exists():
    sys.path.insert(0, str(SITE_PACKAGES))

import importlib.util

CORE_PATH = ROOT / "Application" / "eidp_term_scanner.core.py"
CORE_SPEC = importlib.util.spec_from_file_location("eidp_core", CORE_PATH)
if CORE_SPEC is None or CORE_SPEC.loader is None:
    raise ImportError(f"Unable to load scanner core from {CORE_PATH}")
eidp_core = importlib.util.module_from_spec(CORE_SPEC)
CORE_SPEC.loader.exec_module(eidp_core)


try:
    import pandas as pd  # type: ignore
except Exception as exc:  # pragma: no cover - dependency guard
    raise RuntimeError("pandas is required to export OCR results") from exc


def _looks_numeric(text: str) -> bool:
    text = text.strip()
    if not text:
        return False
    if any(ch.isdigit() for ch in text):
        return True
    keywords = {"min", "max", "value"}
    return text.lower() in keywords


def ocr_page_to_excel(
    pdf_path: Path,
    page_number: int,
    output_path: Path,
    *,
    dpi: int = 900,
    langs: Sequence[str] = ("en",),
    min_conf: float = 0.6,
    row_snap: float = 12.0,
    col_snap: float = 24.0,
    joiner: str = " ",
    include_confidence: bool = False,
    drop_empty_rows: bool = True,
    drop_empty_columns: bool = True,
) -> pd.DataFrame:
    """
    Run EasyOCR on a single page and export the detected tokens as a table.

    Returns the DataFrame that was written to disk.
    """
    if page_number < 1:
        raise ValueError("page_number must be >= 1")

    pdf_path = pdf_path.expanduser().resolve()
    if not pdf_path.exists():
        raise FileNotFoundError(pdf_path)

    items_map = eidp_core._easyocr_boxes_for_pages(
        pdf_path, [page_number], dpi=dpi, langs=list(langs)
    )
    items = items_map.get(page_number, [])
    filtered = [
        it
        for it in items
        if (it.get("text") or "").strip() and float(it.get("conf", 0.0)) >= min_conf
    ]
    if not filtered:
        raise RuntimeError(
            "No OCR tokens detected. Adjust --min-conf, DPI, or languages and try again."
        )

    filtered.sort(key=lambda it: (it["cy"], it["cx"]))

    rows: List[List[dict]] = []
    current: List[dict] = []
    current_center = None

    for item in filtered:
        cy = float(item["cy"])
        if not current:
            current.append(item)
            current_center = cy
            continue
        assert current_center is not None
        if abs(cy - current_center) <= row_snap:
            current.append(item)
            current_center = (current_center * (len(current) - 1) + cy) / len(current)
        else:
            rows.append(sorted(current, key=lambda it: it["cx"]))
            current = [item]
            current_center = cy
    if current:
        rows.append(sorted(current, key=lambda it: it["cx"]))

    token_rows: List[List[Tuple[str, str]]] = []
    for row in rows:
        tokens: List[Tuple[str, str]] = []
        for item in row:
            raw = (item.get("text") or "").strip()
            if not raw:
                continue
            decorated = raw
            if include_confidence:
                conf = float(item.get("conf", 0.0))
                decorated = f"{raw} ({conf:.2f})"
            tokens.append((decorated, raw))
        if drop_empty_rows and not tokens:
            continue
        token_rows.append(tokens)

    if not token_rows:
        raise RuntimeError("No OCR rows remained after filtering/blank-row removal.")

    merged_rows: List[List[Tuple[str, str]]] = []
    i = 0
    while i < len(token_rows):
        tokens = token_rows[i]
        labels = [tok for tok in tokens if not _looks_numeric(tok[1])]
        values = [tok for tok in tokens if _looks_numeric(tok[1])]
        merged = False
        if not labels and values and i + 1 < len(token_rows):
            next_tokens = token_rows[i + 1]
            next_labels = [tok for tok in next_tokens if not _looks_numeric(tok[1])]
            next_values = [tok for tok in next_tokens if _looks_numeric(tok[1])]
            if next_labels and not next_values:
                merged_rows.append(next_tokens + tokens)
                i += 2
                merged = True
        if merged:
            continue
        if labels and not values and i + 1 < len(token_rows):
            next_tokens = token_rows[i + 1]
            next_labels = [tok for tok in next_tokens if not _looks_numeric(tok[1])]
            next_values = [tok for tok in next_tokens if _looks_numeric(tok[1])]
            if next_values and not next_labels:
                merged_rows.append(tokens + next_tokens)
                i += 2
                continue
        merged_rows.append(tokens)
        i += 1
    token_rows = merged_rows

    value_cols = max(
        (sum(1 for _, raw in tokens if _looks_numeric(raw)) for tokens in token_rows),
        default=0,
    )
    label_required = any(
        any(not _looks_numeric(raw) for _, raw in tokens) for tokens in token_rows
    )
    total_cols = max(1, value_cols + (1 if label_required else 0))

    table: List[List[str]] = []
    for tokens in token_rows:
        labels = [decorated for decorated, raw in tokens if not _looks_numeric(raw)]
        values = [decorated for decorated, raw in tokens if _looks_numeric(raw)]
        if not label_required:
            values = labels + values
            labels = []

        row_cells = [""] * total_cols
        idx = 0
        if label_required:
            row_cells[0] = " ".join(labels).strip()
            idx = 1
        for value in values:
            if idx >= total_cols:
                break
            row_cells[idx] = value
            idx += 1
        if not drop_empty_rows or any(cell.strip() for cell in row_cells):
            table.append(row_cells)

    columns = [f"Column {i + 1}" for i in range(total_cols)]
    df = pd.DataFrame(table, columns=columns)
    if drop_empty_columns and not df.empty:
        keep_cols = [col for col in df.columns if df[col].astype(str).str.strip().any()]
        if keep_cols:
            df = df[keep_cols]
        else:
            raise RuntimeError("All columns were empty after OCR clustering.")

    output_path = output_path.expanduser()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if output_path.suffix.lower() == ".csv":
        df.to_csv(output_path, index=False)
    else:
        with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
            df.to_excel(writer, sheet_name=f"Page{page_number}", index=False)
            ws = writer.sheets[f"Page{page_number}"]
            for idx, column in enumerate(df.columns):
                max_len = max((len(str(val)) for val in df[column]), default=len(column))
                ws.set_column(idx, idx, min(80, max(12, max_len + 2)))

    return df


def parse_args(argv: Sequence[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="OCR a PDF page and export to Excel/CSV.")
    parser.add_argument("--pdf", required=True, type=Path, help="Input PDF path.")
    parser.add_argument("--page", required=True, type=int, help="1-based page number.")
    parser.add_argument("--output", required=True, type=Path, help="Output .xlsx or .csv path.")
    parser.add_argument("--dpi", type=int, default=900, help="Rendering DPI (default: 900).")
    parser.add_argument(
        "--langs",
        type=str,
        default="en",
        help="Comma-separated EasyOCR language codes (default: en).",
    )
    parser.add_argument(
        "--min-conf",
        type=float,
        default=0.6,
        help="Minimum confidence to keep a token (default: 0.6).",
    )
    parser.add_argument(
        "--row-snap",
        type=float,
        default=12.0,
        help="Maximum Y delta to treat tokens as the same row (default: 12).",
    )
    parser.add_argument(
        "--col-snap",
        type=float,
        default=24.0,
        help="Maximum X delta to treat tokens as the same column (unused in sequential mode; kept for compatibility).",
    )
    parser.add_argument(
        "--joiner",
        type=str,
        default=" ",
        help="Separator when multiple tokens land in the same cell (default: space).",
    )
    parser.add_argument(
        "--include-confidence",
        action="store_true",
        help="Append confidence scores to each cell value.",
    )
    parser.add_argument(
        "--keep-blank-rows",
        action="store_true",
        help="Keep blank OCR rows instead of dropping them.",
    )
    parser.add_argument(
        "--keep-blank-columns",
        action="store_true",
        help="Keep columns that are entirely blank after OCR.",
    )
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> None:
    args = parse_args(argv or sys.argv[1:])
    langs = [lang.strip() for lang in args.langs.split(",") if lang.strip()]
    ocr_page_to_excel(
        pdf_path=args.pdf,
        page_number=args.page,
        output_path=args.output,
        dpi=args.dpi,
        langs=langs,
        min_conf=args.min_conf,
        row_snap=args.row_snap,
        col_snap=args.col_snap,
        joiner=args.joiner,
        include_confidence=args.include_confidence,
        drop_empty_rows=not args.keep_blank_rows,
        drop_empty_columns=not args.keep_blank_columns,
    )


if __name__ == "__main__":  # pragma: no cover
    main()
