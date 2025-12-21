#!/usr/bin/env python3
"""
Debug JSON runner for OCR_line_geometry_solver.core.

Usage examples:
  python debug\\ocr_debug_runner.py path\\to\\file.pdf --pages 1
  python debug\\ocr_debug_runner.py path\\to\\file.pdf --pages 1,3,5 --dpi 1200

Outputs a single JSON object to stdout with page text and metadata.
"""
from __future__ import annotations

import argparse
import importlib.util
import json
import os
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Tuple


def _parse_pages(pages_raw: str) -> List[int]:
    pages: List[int] = []
    for token in (pages_raw or "").split(","):
        token = token.strip()
        if not token:
            continue
        try:
            value = int(token)
        except ValueError as exc:
            raise argparse.ArgumentTypeError(f"Invalid page '{token}', expected integer") from exc
        if value <= 0:
            raise argparse.ArgumentTypeError(f"Pages must be 1-based positive integers, got {value}")
        pages.append(value)
    if not pages:
        raise argparse.ArgumentTypeError("At least one page is required")
    return pages


def _load_core(core_path: Path):
    spec = importlib.util.spec_from_file_location("eidp_core_debug", core_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Unable to load core module from {core_path}")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)  # type: ignore[attr-defined]
    return mod


def _resolve_core_path(arg_core: str | None) -> Path:
    if arg_core:
        core = Path(arg_core)
    else:
        # Default to the line-geometry core shipped with the app.
        root = Path(__file__).resolve().parents[1]
        core = root / "EIDAT_App_Files" / "Application" / "OCR_line_geometry_solver.core.py"
    core = core.expanduser().resolve()
    if not core.exists():
        raise FileNotFoundError(f"Core module not found: {core}")
    return core


def _build_output(
    pdf_path: Path,
    core_path: Path,
    pages: List[int],
    text_map: Dict[int, str],
    pipeline: str,
    dpi: int,
    tokens_debug: Dict[int, List[Dict[str, Any]]] | None,
) -> Dict[str, object]:
    pages_obj: Dict[str, Dict[str, object]] = {}
    for p in pages:
        text = text_map.get(p, "") or ""
        lines = text.splitlines()
        page_obj: Dict[str, Any] = {
            "page": p,
            "text": text,
            "lines": lines,
        }
        if tokens_debug is not None:
            page_obj["tokens"] = tokens_debug.get(p, [])
        pages_obj[str(p)] = page_obj
    return {
        "pdf_path": str(pdf_path),
        "core_path": str(core_path),
        "pages": pages_obj,
        "requested_dpi": dpi,
        "pipeline": pipeline,
    }


def _parse_langs_from_env() -> List[str]:
    """Mirror the core's language parsing (EASYOCR_LANGS/OCR_LANGS)."""
    langs_raw = (
        os.environ.get("EASYOCR_LANGS")
        or os.environ.get("OCR_LANGS")
        or "en"
    )
    parts = [s.strip() for s in re.split(r"[;,]", langs_raw) if s.strip()]
    return parts or ["en"]


def main(argv: List[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Run OCR (line-geometry core) and emit JSON only.")
    parser.add_argument("pdf", help="Path to the PDF to OCR")
    parser.add_argument(
        "--pages",
        default="1",
        help="Comma-separated list of 1-based page numbers (default: 1)",
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=1200,
        help="Requested OCR DPI passed via OCR_DPI (default: 1200)",
    )
    parser.add_argument(
        "--core-path",
        default=None,
        help="Optional override for OCR core module path (defaults to OCR_line_geometry_solver.core.py)",
    )
    parser.add_argument(
        "--enable-backup",
        action="store_true",
        help="Enable EasyOCR backup retries for low-confidence tokens (OCR_ENABLE_BACKUP=1)",
    )
    parser.add_argument(
        "--enable-tesseract",
        action="store_true",
        help="Enable Tesseract as a secondary backup inside the line-geometry core (OCR_ENABLE_TESSERACT=1)",
    )
    args = parser.parse_args(argv)

    try:
        pdf_path = Path(args.pdf).expanduser().resolve()
        if not pdf_path.exists():
            raise FileNotFoundError(f"PDF not found: {pdf_path}")
        pages = _parse_pages(args.pages)
        if args.dpi <= 0:
            raise ValueError(f"DPI must be positive, got {args.dpi}")

        core_path = _resolve_core_path(args.core_path)

        # Suppress noisy stdout from the main app; we want pure JSON on stdout.
        os.environ.setdefault("QUIET", "1")
        os.environ["OCR_DPI"] = str(args.dpi)
        if args.enable_backup:
            os.environ["OCR_ENABLE_BACKUP"] = "1"
        if args.enable_tesseract:
            os.environ["OCR_ENABLE_TESSERACT"] = "1"

        core = _load_core(core_path)

        # Primary pipeline text using the core's standard entrypoint.
        text_map, pipeline = core.ocr_pages_with_easyocr(pdf_path, pages)  # type: ignore[attr-defined]

        # Optional per-token debug view via the overridden boxes path.
        tokens_debug: Dict[int, List[Dict[str, Any]]] = {}
        try:
            langs = _parse_langs_from_env()
            for p in pages:
                try:
                    items = core._get_easyocr_boxes_page(pdf_path, p, args.dpi, langs)  # type: ignore[attr-defined]
                except Exception:
                    items = []
                norm_items: List[Dict[str, Any]] = []
                for it in items:
                    try:
                        entry: Dict[str, Any] = {
                            "x0": float(it.get("x0", 0.0)),
                            "y0": float(it.get("y0", 0.0)),
                            "x1": float(it.get("x1", 0.0)),
                            "y1": float(it.get("y1", 0.0)),
                            "cx": float(it.get("cx", 0.0)),
                            "cy": float(it.get("cy", 0.0)),
                            "text": str(it.get("text", "")),
                            "conf": float(it.get("conf", 0.0)),
                        }
                        dbg = it.get("debug_backup")
                        if isinstance(dbg, dict):
                            entry["debug_backup"] = {
                                "text_orig": dbg.get("text_orig"),
                                "conf_orig": dbg.get("conf_orig"),
                                "easy_text": dbg.get("easy_text"),
                                "easy_conf": dbg.get("easy_conf"),
                                "tess_text": dbg.get("tess_text"),
                                "tess_conf": dbg.get("tess_conf"),
                                "final_text": dbg.get("final_text"),
                                "final_conf": dbg.get("final_conf"),
                                "used_tess": bool(dbg.get("used_tess")),
                            }
                        norm_items.append(entry)
                    except Exception:
                        continue
                tokens_debug[p] = norm_items
        except Exception:
            tokens_debug = {}

        out = _build_output(
            pdf_path=pdf_path,
            core_path=core_path,
            pages=pages,
            text_map=text_map,
            pipeline=pipeline,
            dpi=args.dpi,
            tokens_debug=tokens_debug,
        )
        json.dump(out, sys.stdout, ensure_ascii=False)
        sys.stdout.write("\n")
        return 0
    except Exception as exc:
        # Emit a JSON error envelope so callers always see JSON on stdout.
        err_obj = {
            "error": type(exc).__name__,
            "message": str(exc),
        }
        json.dump(err_obj, sys.stdout, ensure_ascii=False)
        sys.stdout.write("\n")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
