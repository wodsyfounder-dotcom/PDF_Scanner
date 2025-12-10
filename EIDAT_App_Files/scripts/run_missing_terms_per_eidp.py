from __future__ import annotations

import argparse
import json
import shutil
import sys
from pathlib import Path
from typing import Dict, List, Tuple


def _load_backend():
    """Import ui_next.backend with a stable ROOT reference."""
    here = Path(__file__).resolve()
    app_root = here.parents[1]
    if str(app_root) not in sys.path:
        sys.path.insert(0, str(app_root))
    import ui_next.backend as be  # type: ignore

    return be


def _iter_selection(selection_path: Path) -> List[Tuple[Path, str]]:
    data = json.loads(selection_path.read_text(encoding="utf-8"))
    entries: List[Tuple[Path, str]] = []
    if not isinstance(data, list):
        return entries
    for row in data:
        if not isinstance(row, dict):
            continue
        pdf = row.get("pdf") or row.get("path")
        serial = str(row.get("serial") or row.get("serial_component") or "").strip()
        if not pdf or not serial:
            continue
        try:
            p = Path(pdf)
        except Exception:
            continue
        entries.append((p, serial))
    return entries


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Run missing-terms-only extraction per EIDP.")
    parser.add_argument(
        "--terms",
        dest="terms",
        required=False,
        help="Path to the base Smart-Snap terms workbook.",
    )
    parser.add_argument(
        "--selection-json",
        dest="selection_json",
        required=True,
        help="JSON file describing selected PDFs and serials.",
    )
    args = parser.parse_args(argv)

    be = _load_backend()

    terms_path = Path(args.terms) if args.terms else be.DEFAULT_TERMS_XLSX
    if not terms_path.exists():
        print(f"[ERROR] Terms file not found: {terms_path}", file=sys.stderr, flush=True)
        return 1

    selection_path = Path(args.selection_json)
    if not selection_path.exists():
        print(f"[ERROR] Selection file not found: {selection_path}", file=sys.stderr, flush=True)
        return 1

    try:
        entries_raw = _iter_selection(selection_path)
    except Exception as exc:
        print(f"[ERROR] Unable to read selection list: {exc}", file=sys.stderr, flush=True)
        return 1

    # Deduplicate by serial; if multiple PDFs share a serial, keep the first.
    by_serial: Dict[str, Path] = {}
    for pdf_path, serial in entries_raw:
        if serial in by_serial:
            continue
        by_serial[serial] = pdf_path

    if not by_serial:
        print("[INFO] No valid PDFs/serials in selection; nothing to do.", flush=True)
        return 0

    root = be.ROOT
    stage_root = root / "user_inputs" / "Staging_Selected_Missing"
    tmp_terms = root / "user_inputs" / "terms_missing_only.xlsx"

    any_run = False

    for serial, pdf_path in sorted(by_serial.items(), key=lambda item: item[0]):
        # Compute per-EIDP missing term rows based on the current master.xlsx.
        try:
            headers, missing_rows = be._compute_missing_term_rows([serial], terms_path)  # type: ignore[attr-defined]
        except Exception as exc:  # pragma: no cover - defensive
            print(
                f"[ERROR] Serial {serial}: unable to compute missing terms: {exc}",
                file=sys.stderr,
                flush=True,
            )
            continue

        if not missing_rows:
            print(f"[INFO] Serial {serial}: no missing terms in master; skipping.", flush=True)
            continue

        # Prepare a temporary terms workbook containing only the missing rows.
        try:
            tmp_terms.parent.mkdir(parents=True, exist_ok=True)
            try:
                shutil.copy2(str(terms_path), str(tmp_terms))
            except Exception:
                # If copying fails, we'll rely on write_terms_rows to create a minimal workbook.
                pass
            be.write_terms_rows(missing_rows, path=tmp_terms, headers=headers)
        except Exception as exc:
            print(
                f"[ERROR] Serial {serial}: unable to prepare missing-terms workbook: {exc}",
                file=sys.stderr,
                flush=True,
            )
            continue

        # Stage only this PDF into a dedicated folder so the scanner processes
        # one EIDP at a time with its individualized term list.
        try:
            if stage_root.exists():
                shutil.rmtree(stage_root, ignore_errors=True)
            stage_root.mkdir(parents=True, exist_ok=True)
            shutil.copy2(str(pdf_path), str(stage_root / pdf_path.name))
        except Exception as exc:
            print(
                f"[ERROR] Serial {serial}: unable to stage PDF '{pdf_path}': {exc}",
                file=sys.stderr,
                flush=True,
            )
            continue

        print(
            f"[INFO] Starting missing-terms extraction for serial {serial} ({pdf_path.name})",
            flush=True,
        )
        try:
            proc = be.run_scanner(tmp_terms, stage_root)
        except Exception as exc:
            print(
                f"[ERROR] Serial {serial}: unable to start scanner: {exc}",
                file=sys.stderr,
                flush=True,
            )
            continue

        # Stream child's stdout into this process so the GUI can continue to
        # track progress and discover the run_data folder path.
        stream = getattr(proc, "stdout", None)
        if stream is not None:
            for line in stream:
                try:
                    sys.stdout.write(line.rstrip("\n") + "\n")
                    sys.stdout.flush()
                except Exception:
                    pass
        rc = proc.wait()
        if rc != 0:
            print(
                f"[ERROR] Serial {serial}: scanner exited with code {rc}",
                file=sys.stderr,
                flush=True,
            )
        else:
            any_run = True
            print(
                f"[INFO] Serial {serial}: missing-terms extraction complete",
                flush=True,
            )

    return 0 if any_run else 1


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
