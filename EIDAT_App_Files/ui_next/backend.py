from __future__ import annotations

import json
import math
import os
import sys
import shutil
import subprocess
from pathlib import Path
from typing import Dict, Iterable, Mapping, Optional


APP_ROOT = Path(__file__).resolve().parents[1]
ROOT = APP_ROOT.parent  # repository root that holds user data folders
EXPORTS_ROOT = ROOT / "Product_Data_File"  # default for run_data/plots
LEGACY_EXPORTS_ROOT = ROOT  # legacy run_data/plots when they were root-level
MASTER_DB_ROOT = ROOT / "Product_Data_File" / "Master_Database"  # new home for master/registry/cell state
LEGACY_MASTER_ROOT = ROOT  # previous root-based location for master/registry
LEGACY_MASTER_ROOT_2 = ROOT / "Product_Data_File"  # very old location
DEFAULT_TERMS_XLSX = ROOT / "user_inputs" / "terms.schema.smartsnap.xlsx"
DEFAULT_PLOT_TERMS_XLSX = ROOT / "user_inputs" / "plot_terms.xlsx"
DEFAULT_PROPOSED_PLOTS_JSON = ROOT / "user_inputs" / "proposed_plots.json"
# Default repository root where PDFs may live (user-organized, nested or flat)
DEFAULT_REPO_ROOT = ROOT / "Data Packages"
DEFAULT_PDF_DIR = DEFAULT_REPO_ROOT
SCANNER_ENV = ROOT / "user_inputs" / "scanner.env"
APP_ENTRY = APP_ROOT / "Application" / "eidp_term_scanner.py"
RUNS_DIR = EXPORTS_ROOT / "run_data"
LEGACY_RUNS_DIR = LEGACY_EXPORTS_ROOT / "run_data"
PLOTS_DIR = EXPORTS_ROOT / "plots"
LEGACY_PLOTS_DIR = LEGACY_EXPORTS_ROOT / "plots"
TERMS_TEMPLATE_SHEET = "Template"
TERMS_SCHEMA_COLUMNS = [
    "Data Group",
    "Term Label",
    "Smart Snap Type",
    "Term",
    "Secondary Term",
    "Pages",
    "GroupAfter",
    "GroupBefore",
    "Units",
    "Range (min)",
    "Range (max)",
    "Format",
    "Alt Search",
    "Smart Position",
    "OCR_Row_EPS",
    "DPI",
    "Mode",
    "Return",
]
TERMS_MODE_CHOICES = ["smart", "full table"]
TERMS_SMART_TYPE_CHOICES = ["", "auto", "number", "date", "time", "title"]
MASTER_XLSX = MASTER_DB_ROOT / "master.xlsx"
MASTER_CSV = MASTER_DB_ROOT / "master.csv"
LEGACY_MASTER_XLSX = LEGACY_MASTER_ROOT / "master.xlsx"
LEGACY_MASTER_CSV = LEGACY_MASTER_ROOT / "master.csv"
LEGACY2_MASTER_XLSX = LEGACY_MASTER_ROOT_2 / "master.xlsx"
LEGACY2_MASTER_CSV = LEGACY_MASTER_ROOT_2 / "master.csv"
REG_XLSX = MASTER_DB_ROOT / "run_registry.xlsx"
REG_CSV = MASTER_DB_ROOT / "run_registry.csv"
LEGACY_REG_XLSX = LEGACY_MASTER_ROOT / "run_registry.xlsx"
LEGACY_REG_CSV = LEGACY_MASTER_ROOT / "run_registry.csv"
LEGACY2_REG_XLSX = LEGACY_MASTER_ROOT_2 / "run_registry.xlsx"
LEGACY2_REG_CSV = LEGACY_MASTER_ROOT_2 / "run_registry.csv"
MASTER_BASE_COLUMNS = {"Term Label", "Data Group", "Units", "Min", "Max"}


def parse_scanner_env(path: Path = SCANNER_ENV) -> Dict[str, str]:
    env: Dict[str, str] = {}
    try:
        if not path.exists():
            return env
        for raw in path.read_text(encoding="utf-8", errors="ignore").splitlines():
            line = raw.strip()
            if not line or line.startswith("#") or line.startswith(";"):
                continue
            if "=" not in line:
                continue
            k, v = line.split("=", 1)
            k = k.strip()
            v = v.strip()
            if "#" in v:
                v = v.split("#", 1)[0].strip()
            if ";" in v:
                v = v.split(";", 1)[0].strip()
            if not k:
                continue
            if v == "":
                continue
            env[k] = v
    except Exception:
        pass
    return env


def save_scanner_env(env_map: Dict[str, str], path: Path = SCANNER_ENV) -> None:
    lines = [
        "# Scanner configuration (KEY=VALUE)",
        "# Edited via new GUI",
    ]
    order = [
        "QUIET",
        "REPO_ROOT",
        "OCR_MODE",
        "OCR_DPI",
        "EASYOCR_LANGS",
        "FORCE_OCR",
        "USE_EASYOCR_XY",
        "XY_LOG",
        "OCR_ROW_EPS",
        "VENV_DIR",
    ]
    written = set()
    for k in order:
        v = env_map.get(k)
        if v:
            lines.append(f"{k}={v}")
            written.add(k)
    for k in sorted(env_map.keys()):
        if k in written:
            continue
        v = env_map[k]
        if v:
            lines.append(f"{k}={v}")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def get_repo_root() -> Path:
    """Return repository root from scanner.env or DEFAULT_REPO_ROOT."""
    env = parse_scanner_env(SCANNER_ENV)
    val = env.get("REPO_ROOT", "").strip()
    try:
        if val:
            p = Path(val).expanduser()
            return p
    except Exception:
        pass
    return DEFAULT_REPO_ROOT


def set_repo_root(p: Path) -> None:
    env = parse_scanner_env(SCANNER_ENV)
    env["REPO_ROOT"] = str(Path(p).expanduser())
    save_scanner_env(env)


def _prefer_existing(*paths: Path) -> Path:
    """Return the first existing path, or the first element if none exist."""
    for p in paths:
        if p.exists():
            return p
    return paths[0]


def master_xlsx_for_read() -> Path:
    """Prefer the new master.xlsx location, fall back to legacy/root variants."""
    return _prefer_existing(MASTER_XLSX, LEGACY_MASTER_XLSX, LEGACY2_MASTER_XLSX)


def master_csv_for_read() -> Path:
    """Prefer the new master.csv location, fall back to legacy/root variants."""
    return _prefer_existing(MASTER_CSV, LEGACY_MASTER_CSV, LEGACY2_MASTER_CSV)


def registry_csv_for_read() -> Path:
    """Prefer the new run_registry.csv location, fall back to legacy/root variants."""
    return _prefer_existing(REG_CSV, LEGACY_REG_CSV, LEGACY2_REG_CSV)


def registry_xlsx_for_read() -> Path:
    """Prefer the new run_registry.xlsx location, fall back to legacy/root variants."""
    return _prefer_existing(REG_XLSX, LEGACY_REG_XLSX, LEGACY2_REG_XLSX)


def run_roots() -> list[Path]:
    """Return new run_data root first, with legacy Product_Data_File/run_data as fallback."""
    roots = [RUNS_DIR]
    if LEGACY_RUNS_DIR.exists():
        roots.append(LEGACY_RUNS_DIR)
    return roots


def plots_root_for_open() -> Path:
    """Prefer the new plots folder, fall back to legacy Product_Data_File/plots if present."""
    if PLOTS_DIR.exists() or not LEGACY_PLOTS_DIR.exists():
        return PLOTS_DIR
    return LEGACY_PLOTS_DIR


def _venv_python_from(path: Path) -> Path:
    if os.name == "nt":
        return path / "Scripts" / "python.exe"
    return path / "bin" / "python"


def resolve_project_python() -> str:
    env = parse_scanner_env(SCANNER_ENV)
    vdir = env.get("VENV_DIR", "").strip()
    if vdir:
        cand = _venv_python_from(Path(vdir))
        if cand.exists():
            return str(cand)
    cand = _venv_python_from(APP_ROOT / ".venv")
    if cand.exists():
        return str(cand)
    return sys.executable


def _base_env() -> Dict[str, str]:
    env = os.environ.copy()
    # Merge scanner.env values for direct Python invocations
    env.update(parse_scanner_env(SCANNER_ENV))
    # Ensure vendored site-packages are importable as fallback
    env["PYTHONPATH"] = str(APP_ROOT / "Lib" / "site-packages") + os.pathsep + env.get("PYTHONPATH", "")
    env.setdefault("QUIET", "1")
    # Force cache to always be in project root, not executable location
    # This ensures cache is consistent whether running as script or frozen exe
    if "CACHE_ROOT" not in env and "OCR_CACHE_ROOT" not in env:
        env["CACHE_ROOT"] = str(ROOT)
    return env


def spawn(cmd: Iterable[str], *, cwd: Optional[Path] = None, env: Optional[Dict[str, str]] = None) -> subprocess.Popen:
    return subprocess.Popen(
        list(cmd),
        cwd=str(cwd or ROOT),
        env=env or _base_env(),
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
    )


def run_install_full() -> subprocess.Popen:
    if not sys.platform.startswith("win"):
        raise RuntimeError("install.bat requires Windows")
    return spawn(["cmd.exe", "/c", str(ROOT / "install.bat")])


def run_scanner(terms: Path, pdf_dir: Path) -> subprocess.Popen:
    if sys.platform.startswith("win"):
        cmd = [
            "cmd.exe", "/c", str(ROOT / "run.bat"),
            "--input", str(terms),
            "--pdf-folder", str(pdf_dir),
            "--quiet",
        ]
        return spawn(cmd)
    else:
        py = resolve_project_python()
        cmd = [py, str(APP_ENTRY), "--input", str(terms), "--pdf-folder", str(pdf_dir), "--quiet"]
        return spawn(cmd)


def run_script(script_rel_path: str, *args: str) -> subprocess.Popen:
    py = resolve_project_python()
    script = APP_ROOT / script_rel_path
    if not script.exists():
        raise FileNotFoundError(f"Missing script: {script}")
    return spawn([py, str(script), *args])


def generate_terms() -> subprocess.Popen:
    return run_script("scripts/generate_terms_schema_smartsnap.py")


def compile_master() -> subprocess.Popen:
    return run_script("scripts/compile_master.py")


def compile_master_from_state() -> None:
    """
    Rebuild master.xlsx from master_cell_state.json ONLY.
    This discards all manual edits and rebuilds from extracted data.

    Unlike compile_master(), this runs synchronously in the current process.
    """
    import sys as _sys
    if str(APP_ROOT) not in _sys.path:
        _sys.path.insert(0, str(APP_ROOT))
    from scripts.compile_master import build_master_from_state, write_master

    # Build from state
    serials, rows, prog_map, sv_map, data_map = build_master_from_state()

    if not serials:
        print("[WARN] No data in cell state to compile")
        return

    # Write the master workbook
    write_master(serials, rows, program_by_sn=prog_map, sv_by_sn=sv_map, data_by_sn=data_map)


def generate_plot_terms() -> subprocess.Popen:
    return run_script("scripts/generate_plot_terms.py")


def generate_plots() -> subprocess.Popen:
    return run_script("scripts/plot_from_master.py")


def export_plots_summary() -> subprocess.Popen:
    return run_script("scripts/plots_to_excel_summary.py")


def extract_page_tables(pdf: Path, pages: str | None = None) -> subprocess.Popen:
    args = ["--pdf", str(pdf)]
    if pages:
        args += ["--pages", pages]
    return run_script("scripts/extract_page_tables.py", *args)


def extract_csv_tables(pdf: Path, pages: str, num_cols: Optional[int] = None,
                      min_cols: int = 2, min_rows: int = 3,
                      match_threshold: float = 0.5, ocr: bool = False,
                      dpi: int = 300, delimiter: Optional[str] = None,
                      output: Optional[Path] = None) -> subprocess.Popen:
    """Extract tables using line-by-line CSV detection with smart alignment."""
    args = ["--pdf", str(pdf), "--pages", pages]
    if num_cols is not None:
        args += ["--num-cols", str(num_cols)]
    if min_cols != 2:
        args += ["--min-cols", str(min_cols)]
    if min_rows != 3:
        args += ["--min-rows", str(min_rows)]
    if match_threshold != 0.5:
        args += ["--match-threshold", str(match_threshold)]
    if ocr:
        args += ["--ocr"]
    if dpi != 300:
        args += ["--dpi", str(dpi)]
    if delimiter:
        args += ["--delimiter", delimiter]
    if output:
        args += ["--out", str(output)]
    return run_script("scripts/extract_table_csv_lines.py", *args)


def extract_tables_by_keywords(pdf: Path, keywords: str, pages: str | None = None,
                               max_cols: int = 6, ocr_dpi: int = 300,
                               best_only: bool = True, output: Optional[Path] = None) -> subprocess.Popen:
    """
    Extract tables using the geometry/keyword-driven extractor.
    """
    args = ["--pdf", str(pdf), "--keywords", keywords]
    if pages:
        args += ["--pages", pages]
    if max_cols != 6:
        args += ["--max-cols", str(max_cols)]
    if ocr_dpi != 300:
        args += ["--ocr-dpi", str(ocr_dpi)]
    if best_only:
        args += ["--best-only"]
    if output:
        args += ["--out", str(output)]
    return run_script("scripts/extract_tables_by_keywords.py", *args)


def pre_ocr_merge_pdfs(paths: list[Path], out_root: Optional[Path] = None, dpi: Optional[int] = None) -> subprocess.Popen:
    """Pre-OCR selected PDFs and write merged text artifacts (no extraction)."""
    args: list[str] = []
    for p in paths:
        args += ["--pdf", str(Path(p))]
    if out_root:
        args += ["--out", str(Path(out_root))]
    if dpi is not None:
        args += ["--dpi", str(int(dpi))]
    return run_script("scripts/pre_ocr_merge.py", *args)


def run_simple_extraction(paths: list[Path], terms: Optional[Path] = None) -> subprocess.Popen:
    """Run the simple merged-text extraction pipeline on selected PDFs."""
    terms_path = Path(terms) if terms else ROOT / "user_inputs" / "terms.schema.simple.xlsx"
    args: list[str] = []
    for p in paths:
        args += ["--pdf", str(Path(p))]
    args += ["--terms", str(terms_path)]
    return run_script("scripts/simple_extraction.py", *args)


def open_path(p: Path) -> None:
    if sys.platform.startswith("win"):
        os.startfile(str(p))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(p)])
    else:
        subprocess.Popen(["xdg-open", str(p)])


def open_terms_file(path: Optional[Path] = None) -> None:
    tgt = Path(path) if path else DEFAULT_TERMS_XLSX
    open_path(tgt)


def derive_return_value(row: Mapping[str, str], existing: Optional[str] = None) -> str:
    """Infer the return type (number|string) when the column is hidden in the UI."""
    smart = (row.get("Smart Snap Type") or "").strip().lower()
    if smart in ("number", "num", "value"):
        return "number"
    if smart in ("date", "time", "title"):
        return "string"
    hint = (existing or row.get("Return") or "").strip().lower()
    if hint in ("number", "string"):
        return hint
    return "number"


def _ensure_openpyxl_loader():
    try:
        from openpyxl import load_workbook  # type: ignore
    except Exception as exc:  # pragma: no cover - depends on optional dep
        raise RuntimeError(
            "openpyxl is required to edit the Smart-Snap terms spreadsheet. "
            "Install it with `py -m pip install openpyxl` within the project environment."
        ) from exc
    return load_workbook


def read_terms_rows(path: Optional[Path] = None) -> tuple[list[str], list[dict[str, str]]]:
    """Return (headers, rows) from the Smart-Snap terms sheet starting at row 3."""
    tgt = Path(path) if path else DEFAULT_TERMS_XLSX
    if not tgt.exists():
        raise FileNotFoundError(f"Terms spreadsheet not found: {tgt}")

    # Ensure all required columns exist before reading
    ensure_terms_columns(tgt)

    load_wb = _ensure_openpyxl_loader()
    wb = load_wb(tgt)
    try:
        if TERMS_TEMPLATE_SHEET not in wb.sheetnames:
            raise RuntimeError(f"Sheet '{TERMS_TEMPLATE_SHEET}' missing in {tgt}")
        ws = wb[TERMS_TEMPLATE_SHEET]
        headers: list[str] = []
        for idx, fallback in enumerate(TERMS_SCHEMA_COLUMNS, start=1):
            raw = ws.cell(row=1, column=idx).value
            name = str(raw).strip() if raw not in (None, "") else fallback
            headers.append(name or fallback)
        rows: list[dict[str, str]] = []
        def _is_template_metadata(row_vals: tuple) -> bool:
            hay = " ".join([str(v or "").strip().lower() for v in row_vals])
            return (
                "data group" in hay
                and "term label" in hay
                and "smart snap type" in hay
            )

        for values in ws.iter_rows(min_row=2, max_col=len(headers), values_only=True):
            if _is_template_metadata(values):
                continue
            normalized: dict[str, str] = {}
            has_value = False
            for col_idx, header in enumerate(headers):
                cell_val = values[col_idx] if col_idx < len(values) else None
                if cell_val is None:
                    text = ""
                else:
                    text = str(cell_val)
                if text.strip():
                    has_value = True
                normalized[header] = text
            if has_value:
                rows.append(normalized)
        if not rows:
            rows.append({h: "" for h in headers})
        return headers, rows
    finally:
        wb.close()


def ensure_terms_columns(path: Optional[Path] = None) -> bool:
    """Ensure the terms spreadsheet has all required columns from TERMS_SCHEMA_COLUMNS.
    Returns True if columns were added, False if no changes needed."""
    tgt = Path(path) if path else DEFAULT_TERMS_XLSX
    if not tgt.exists():
        return False
    load_wb = _ensure_openpyxl_loader()
    wb = load_wb(tgt)
    try:
        if TERMS_TEMPLATE_SHEET not in wb.sheetnames:
            return False
        ws = wb[TERMS_TEMPLATE_SHEET]
        changed = False

        # Remove legacy line-mode columns that are no longer used in the Smart-Snap schema.
        deprecated = {"Line", "Column", "Anchor", "FieldIndex", "FieldSplit"}
        try:
            max_col = ws.max_column or 0
            for col_idx in range(max_col, 0, -1):
                raw = ws.cell(row=1, column=col_idx).value
                if raw is None:
                    continue
                name = str(raw).strip()
                if name in deprecated:
                    ws.delete_cols(col_idx)
                    changed = True
        except Exception:
            # Best-effort cleanup; ignore failures so we still enforce required columns.
            pass

        # Read existing headers
        existing_headers = []
        for idx in range(1, len(TERMS_SCHEMA_COLUMNS) + 1):
            raw = ws.cell(row=1, column=idx).value
            if raw is not None and str(raw).strip():
                existing_headers.append(str(raw).strip())
            else:
                break

        # Check if we need to add any columns
        missing_columns = [col for col in TERMS_SCHEMA_COLUMNS if col not in existing_headers]
        if missing_columns:
            # Add missing column headers to row 1
            start_col = len(existing_headers) + 1
            for idx, col_name in enumerate(missing_columns, start=start_col):
                ws.cell(row=1, column=idx, value=col_name)
            changed = True

        if changed:
            wb.save(tgt)
        return changed
    finally:
        wb.close()


def write_terms_rows(
    rows: list[dict[str, str]],
    path: Optional[Path] = None,
    headers: Optional[list[str]] = None,
) -> None:
    """Persist edited Smart-Snap rows back into the template sheet (rows 3+)."""
    tgt = Path(path) if path else DEFAULT_TERMS_XLSX
    if not tgt.exists():
        raise FileNotFoundError(f"Terms spreadsheet not found: {tgt}")

    # Ensure all required columns exist before writing
    ensure_terms_columns(tgt)

    load_wb = _ensure_openpyxl_loader()
    wb = load_wb(tgt)
    try:
        if TERMS_TEMPLATE_SHEET not in wb.sheetnames:
            raise RuntimeError(f"Sheet '{TERMS_TEMPLATE_SHEET}' missing in {tgt}")
        ws = wb[TERMS_TEMPLATE_SHEET]
        if headers is None:
            headers = []
            for idx, fallback in enumerate(TERMS_SCHEMA_COLUMNS, start=1):
                raw = ws.cell(row=1, column=idx).value
                name = str(raw).strip() if raw not in (None, "") else fallback
                headers.append(name or fallback)
        existing_rows = max(ws.max_row - 1, 0)
        if existing_rows > 0:
            ws.delete_rows(2, existing_rows)
        records = rows or [{h: "" for h in headers}]
        for record in records:
            ws.append([record.get(h, "") for h in headers])
        wb.save(tgt)
    finally:
        wb.close()

def _clean_master_cell(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float):
        if math.isnan(value):
            return ""
    return str(value).strip()

def _read_master_table() -> tuple[list[str], list[dict[str, str]]]:
    header: list[str] = []
    rows: list[dict[str, str]] = []
    master_path = master_xlsx_for_read()
    if not master_path.exists():
        return header, rows

    try:
        import pandas as pd  # type: ignore
        # Preserve textual sentinels such as 'N/A' instead of coercing them
        # to NaN so downstream logic can distinguish between true blanks and
        # explicit "not applicable" markers.
        df = pd.read_excel(master_path, dtype=object, keep_default_na=False)
        header = [str(col) for col in df.columns]
        raw_rows = df.fillna("").to_dict(orient="records")
        for record in raw_rows:
            cleaned = {str(k): _clean_master_cell(v) for k, v in record.items()}
            rows.append(cleaned)
        return header, rows
    except Exception:
        try:
            from openpyxl import load_workbook  # type: ignore
            wb = load_workbook(str(master_path), data_only=True)
            ws = wb.active
            first_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
            header = [(_clean_master_cell(val) or f"column_{idx}") for idx, val in enumerate(first_row)]
            for values in ws.iter_rows(min_row=2, values_only=True):
                record: dict[str, str] = {}
                for idx, val in enumerate(values):
                    if idx >= len(header):
                        continue
                    record[header[idx]] = _clean_master_cell(val)
                rows.append(record)
            wb.close()
            return header, rows
        except Exception:
            pass
    return header, rows

def _schema_value(row: Mapping[str, str], key: str) -> str:
    for variant in (key, key.lower(), key.upper()):
        if variant in row:
            val = row.get(variant)
            if val is None:
                continue
            return str(val).strip()
    return ""

def _compute_missing_term_rows(
    serials: list[str],
    terms_path: Optional[Path] = None,
) -> tuple[list[str], list[dict[str, str]]]:
    target = Path(terms_path) if terms_path else DEFAULT_TERMS_XLSX
    headers, schema_rows = read_terms_rows(target)
    normalized_serials = [s.strip() for s in serials if s and s.strip()]
    if not normalized_serials:
        return headers, []
    master_header, master_rows = _read_master_table()
    if not master_header or not master_rows:
        # No master yet: treat all schema rows as missing for the selected serials.
        return headers, list(schema_rows)
    available_serials = {col for col in master_header if col not in MASTER_BASE_COLUMNS}
    # Lookup of rows present in master keyed by (term_label, data_group)
    term_lookup: dict[tuple[str, str], dict[str, str]] = {}
    for row in master_rows:
        term_label = row.get("Term Label", "").strip()
        if not term_label:
            continue
        if term_label.lower() in ("program", "space vehicle", "data"):
            continue
        data_group = row.get("Data Group", "").strip()
        term_lookup[(term_label.lower(), data_group.lower())] = row
    missing_rows: list[dict[str, str]] = []
    for schema_row in schema_rows:
        term_label = _schema_value(schema_row, "Term Label")
        if not term_label:
            continue
        data_group = _schema_value(schema_row, "Data Group")
        key = (term_label.lower(), data_group.lower())
        master_row = term_lookup.get(key)
        needs_run = False
        for serial in normalized_serials:
            # If the serial column is absent or the row itself doesn't exist in
            # master, this schema term has never been recorded for that EIDP.
            if serial not in available_serials or master_row is None:
                needs_run = True
                break
            value = master_row.get(serial, "")
            text = str(value or "").strip()
            # Blank cells are missing; explicit 'N/A' is treated as already-attempted.
            if not text:
                needs_run = True
                break
            if text.upper() == "N/A":
                # Already attempted; do not treat as missing for this serial.
                continue
        if needs_run:
            missing_rows.append(schema_row)
    return headers, missing_rows

def count_missing_terms(serials: list[str], terms_path: Optional[Path] = None) -> int:
    _, rows = _compute_missing_term_rows(serials, terms_path)
    return len(rows)

def count_missing_terms_per_serial(
    serials: list[str],
    terms_path: Optional[Path] = None,
) -> dict[str, int]:
    """Return mapping {serial: missing_term_count} based on master.xlsx.

    A term is counted as missing for a given serial when:
      - the (Term Label, Data Group) pair doesn't exist in master at all, or
      - the row exists but the cell for that serial is blank.
    Cells containing 'N/A' are treated as already-attempted and not missing.
    """
    target = Path(terms_path) if terms_path else DEFAULT_TERMS_XLSX
    _, schema_rows = read_terms_rows(target)
    normalized_serials = [s.strip() for s in serials if s and str(s).strip()]
    if not normalized_serials:
        return {}
    counts: dict[str, int] = {s: 0 for s in normalized_serials}
    master_header, master_rows = _read_master_table()
    # If no master is present yet, treat all schema terms as missing for each serial.
    if not master_header or not master_rows:
        total_terms = 0
        for schema_row in schema_rows:
            label = _schema_value(schema_row, "Term Label")
            if label:
                total_terms += 1
        return {s: total_terms for s in normalized_serials}

    available_serials = {col for col in master_header if col not in MASTER_BASE_COLUMNS}
    # Build lookup from (term_label, data_group) -> master row (existing rows only)
    term_lookup: dict[tuple[str, str], dict[str, str]] = {}
    for row in master_rows:
        term_label = str(row.get("Term Label", "") or "").strip()
        if not term_label:
            continue
        # Skip metadata rows
        if term_label.lower() in ("program", "space vehicle", "data"):
            continue
        data_group = str(row.get("Data Group", "") or "").strip()
        term_lookup[(term_label.lower(), data_group.lower())] = row

    for schema_row in schema_rows:
        term_label = _schema_value(schema_row, "Term Label")
        if not term_label:
            continue
        data_group = _schema_value(schema_row, "Data Group")
        key = (term_label.lower(), data_group.lower())
        master_row = term_lookup.get(key)
        for serial in normalized_serials:
            # If serial column is missing or row absent, this term is missing.
            if serial not in available_serials or master_row is None:
                counts[serial] = counts.get(serial, 0) + 1
                continue
            text = str(master_row.get(serial, "") or "").strip()
            # Treat explicit N/A as "attempted" (not missing).
            if not text:
                counts[serial] = counts.get(serial, 0) + 1
            elif text.upper() == "N/A":
                # Already attempted; do not treat as missing.
                continue
    return counts

def run_missing_terms_for_selected_pdfs(
    selected: list[tuple[Path, str]],
    terms: Optional[Path] = None,
) -> subprocess.Popen:
    if not selected:
        raise RuntimeError("No data packages selected.")
    terms_path = Path(terms) if terms else DEFAULT_TERMS_XLSX
    serials = sorted({str(serial).strip() for _, serial in selected if str(serial).strip()})
    if not serials:
        raise RuntimeError("Selected data packages do not include serial identifiers.")
    # Quick pre-check using the master workbook so we can fail fast
    # if every selected EIDP already has values for all schema terms.
    try:
        total_missing = count_missing_terms(serials, terms_path)
    except Exception as exc:
        raise RuntimeError(f"Unable to inspect master workbook for missing terms: {exc}") from exc
    if total_missing <= 0:
        raise RuntimeError("No missing terms found for the selected data packages.")
    # Persist the selection for the batch helper script.
    selection_path = ROOT / "user_inputs" / "missing_terms_selection.json"
    try:
        selection_path.parent.mkdir(parents=True, exist_ok=True)
        payload: list[dict[str, str]] = []
        for pdf_path, serial in selected:
            try:
                p = Path(pdf_path)
            except Exception:
                continue
            s = str(serial).strip()
            if not s:
                continue
            payload.append({"pdf": str(p), "serial": s})
        if not payload:
            raise RuntimeError("Selected data packages do not include any valid PDF/serial pairs.")
        selection_path.write_text(json.dumps(payload), encoding="utf-8")
    except Exception as exc:
        raise RuntimeError(f"Unable to prepare missing-terms selection: {exc}") from exc
    # Delegate the per-EIDP, missing-terms-only extraction to a small helper script
    # so that the GUI can track a single process while each EIDP is processed with
    # its own tailored term list.
    return run_script(
        "scripts/run_missing_terms_per_eidp.py",
        "--terms",
        str(terms_path),
        "--selection-json",
        str(selection_path),
    )


# --- Workspace sync helpers ---

from datetime import datetime
import csv as _csv


def _parse_dt(s: str) -> Optional[datetime]:
    s = (s or "").strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S", "%m/%d/%Y %H:%M", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            continue
    return None


def _load_schema_term_keys(terms_path: Path) -> set[tuple[str, str]]:
    """Return set of (term_label, data_group) keys defined in the schema.

    Keys are lowercased to allow case-insensitive comparison. If the schema
    cannot be read, an empty set is returned and sync falls back to registry
    and file timestamps only.
    """
    keys: set[tuple[str, str]] = set()
    try:
        if not terms_path.exists():
            return keys
        suffix = terms_path.suffix.lower()
        # Excel-based schema (preferred)
        if suffix in (".xlsx", ".xlsm", ".xls"):
            try:
                import pandas as _pd  # type: ignore
                try:
                    df = _pd.read_excel(terms_path, sheet_name=TERMS_TEMPLATE_SHEET)
                except Exception:
                    df = _pd.read_excel(terms_path)
                for _, row in df.iterrows():
                    label = str(row.get("Term Label") or "").strip()
                    group = str(row.get("Data Group") or "").strip()
                    if not label and not group:
                        continue
                    keys.add((label.lower(), group.lower()))
                return keys
            except Exception:
                # Fall back to openpyxl if pandas or Excel stack isn't available
                try:
                    import openpyxl as _ox  # type: ignore
                    wb = _ox.load_workbook(str(terms_path), read_only=True, data_only=True)
                    ws = wb[TERMS_TEMPLATE_SHEET] if TERMS_TEMPLATE_SHEET in wb.sheetnames else wb.active
                    rows = list(ws.iter_rows(values_only=True))
                    if not rows:
                        return keys
                    headers = [str(v) if v is not None else "" for v in rows[0]]
                    try:
                        label_idx = headers.index("Term Label")
                    except ValueError:
                        label_idx = None
                    try:
                        group_idx = headers.index("Data Group")
                    except ValueError:
                        group_idx = None
                    if label_idx is None and group_idx is None:
                        return keys
                    for r in rows[1:]:
                        label = ""
                        group = ""
                        if label_idx is not None and label_idx < len(r) and r[label_idx] is not None:
                            label = str(r[label_idx]).strip()
                        if group_idx is not None and group_idx < len(r) and r[group_idx] is not None:
                            group = str(r[group_idx]).strip()
                        if not label and not group:
                            continue
                        keys.add((label.lower(), group.lower()))
                    return keys
                except Exception:
                    return keys
        # CSV schema (or other text-based formats)
        try:
            with terms_path.open("r", encoding="utf-8", newline="") as f:
                r = _csv.DictReader(f)
                for row in r:
                    label = str(
                        row.get("Term Label")
                        or row.get("term_label")
                        or row.get("Term")
                        or row.get("term")
                        or ""
                    ).strip()
                    group = str(row.get("Data Group") or row.get("data_group") or "").strip()
                    if not label and not group:
                        continue
                    keys.add((label.lower(), group.lower()))
        except Exception:
            return keys
    except Exception:
        return set()
    return keys


def _load_run_terms_map(reg_map: dict[str, dict[str, str]]) -> dict[str, set[tuple[str, str]]]:
    """Return mapping {serial_component: {(term_label, data_group), ...}} from scan_results.json.

    Only considers rows for the specific serial in each registry entry. Missing or
    unreadable run folders / results are silently ignored (callers can treat those
    serials as "not yet run" or out-of-sync).
    """
    out: dict[str, set[tuple[str, str]]] = {}
    for serial, info in reg_map.items():
        try:
            run_dir = Path(info.get("run_folder", ""))
        except Exception:
            continue
        if not run_dir or not run_dir.exists():
            continue
        path = run_dir / "scan_results.json"
        if not path.exists():
            continue
        try:
            with path.open("r", encoding="utf-8") as fh:
                data = json.load(fh)
        except Exception:
            continue
        if not isinstance(data, list):
            continue
        keys: set[tuple[str, str]] = set()
        for row in data:
            if not isinstance(row, dict):
                continue
            row_id = (str(row.get("serial_component") or row.get("serial_number") or "")).strip()
            if row_id and row_id != serial:
                continue
            term_label = str(row.get("term_label") or row.get("term") or "").strip()
            if not term_label:
                continue
            data_group = str(row.get("data_group") or "").strip()
            keys.add((term_label.lower(), data_group.lower()))
        if keys:
            out[serial] = keys
    return out


def _read_run_registry_map() -> dict[str, dict[str, str]]:
    """Return mapping {serial_component: {run_date, run_folder, program_name, vehicle_number}}.

    Reads run_registry.(xlsx|csv) from Product_Data_File/Master_Database, falling back to
    legacy root or Product_Data_File locations. Missing file -> empty map.
    """
    rx = registry_xlsx_for_read()
    rc = registry_csv_for_read()
    out: dict[str, dict[str, str]] = {}
    if rc.exists():
        try:
            with rc.open("r", encoding="utf-8", newline="") as f:
                r = _csv.DictReader(f)
                for row in r:
                    sc = (row.get("serial_component") or row.get("serial_number") or "").strip()
                    if not sc:
                        continue
                    out[sc] = {
                        "run_date": (row.get("run_date") or ""),
                        "run_folder": (row.get("run_folder") or ""),
                        "program_name": (row.get("program_name") or ""),
                        "vehicle_number": (row.get("vehicle_number") or ""),
                    }
        except Exception:
            pass
    elif rx.exists():
        # One-time conversion from xlsx to csv, then delete xlsx to avoid confusion
        try:
            import openpyxl as _ox  # type: ignore
            wb = _ox.load_workbook(str(rx), read_only=True, data_only=True)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True)) if ws else []
            if rows:
                headers = [str(x) if x is not None else "" for x in rows[0]]
                REG_CSV.parent.mkdir(parents=True, exist_ok=True)
                with REG_CSV.open("w", encoding="utf-8", newline="") as f:
                    w = _csv.writer(f)
                    w.writerow(headers)
                    for r in rows[1:]:
                        w.writerow([("") if c is None else str(c) for c in r])
            try:
                rx.unlink()
            except Exception:
                pass
            # Recurse to read the new CSV
            return _read_run_registry_map()
        except Exception:
            pass
    return out


def _write_run_registry_map(rows: dict[str, dict[str, str]]) -> None:
    """Write the run registry to CSV only. Remove any legacy XLSX to avoid drift."""
    reg_dir = MASTER_DB_ROOT
    reg_dir.mkdir(parents=True, exist_ok=True)
    rx = REG_XLSX
    rc = REG_CSV
    columns = ["serial_component", "program_name", "vehicle_number", "run_date", "run_folder"]
    try:
        with rc.open("w", encoding="utf-8", newline="") as f:
            w = _csv.DictWriter(f, fieldnames=columns)
            w.writeheader()
            for sc in sorted(rows.keys()):
                m = rows[sc]
                w.writerow({
                    "serial_component": sc,
                    "program_name": m.get("program_name", ""),
                    "vehicle_number": m.get("vehicle_number", ""),
                    "run_date": m.get("run_date", ""),
                    "run_folder": m.get("run_folder", ""),
                })
    except Exception:
        pass
    # Remove legacy xlsx to prevent confusion
    try:
        if rx.exists():
            rx.unlink()
        if LEGACY_REG_XLSX.exists():
            LEGACY_REG_XLSX.unlink()
        if LEGACY2_REG_XLSX.exists():
            LEGACY2_REG_XLSX.unlink()
        if LEGACY_REG_CSV.exists() and LEGACY_REG_CSV != rc:
            LEGACY_REG_CSV.unlink()
        if LEGACY2_REG_CSV.exists() and LEGACY2_REG_CSV != rc:
            LEGACY2_REG_CSV.unlink()
    except Exception:
        pass


def write_run_registry_rows(rows: list[dict[str, str]]) -> None:
    """Public helper to write a list of row dicts to the run registry.

    Expects keys including at least 'serial_component', 'run_date', 'run_folder'.
    """
    mapping: dict[str, dict[str, str]] = {}
    for r in rows:
        sc = str(r.get("serial_component", "")).strip()
        if not sc:
            continue
        mapping[sc] = {
            "serial_component": sc,
            "program_name": str(r.get("program_name", "")),
            "vehicle_number": str(r.get("vehicle_number", "")),
            "run_date": str(r.get("run_date", "")),
            "run_folder": str(r.get("run_folder", "")),
        }
    _write_run_registry_map(mapping)


def ensure_run_registry_consistent() -> dict[str, dict[str, str]]:
    """Prune invalid entries from run_registry.csv (no additions).

    Keeps rows whose run_folder exists and that contain at least one
    per-document artifact for the serial (filename contains serial and
    has .xlsx/.json/.csv extension). Removes others.
    """
    reg = _read_run_registry_map()
    if not reg:
        return {}
    cleaned: dict[str, dict[str, str]] = {}
    for sc, info in reg.items():
        try:
            run_dir = Path(info.get("run_folder", ""))
        except Exception:
            run_dir = None
        if not run_dir or not run_dir.exists():
            continue
        found = False
        try:
            # Look for files that include the serial in the name
            for fp in run_dir.iterdir():
                if not fp.is_file():
                    continue
                name = fp.name.lower()
                if sc.lower() in name and fp.suffix.lower() in (".xlsx", ".json", ".csv"):
                    found = True
                    break
        except Exception:
            pass
        if found:
            cleaned[sc] = info
    _write_run_registry_map(cleaned)
    return cleaned


def delete_registry_entries(serial_components: list[str]) -> None:
    """Delete given serials from run registry and remove their per-document run_data files.

    Does not remove whole run folders.
    """
    reg = _read_run_registry_map()
    for sc in serial_components:
        info = reg.get(sc)
        run_dir: Optional[Path] = None
        try:
            run_dir = Path(info.get("run_folder")) if info else None
        except Exception:
            run_dir = None
        if run_dir and run_dir.exists():
            try:
                for fp in run_dir.iterdir():
                    try:
                        if fp.is_file() and sc.lower() in fp.name.lower() and fp.suffix.lower() in (".xlsx", ".json", ".csv"):
                            fp.unlink(missing_ok=True)
                    except Exception:
                        pass
            except Exception:
                pass
        if sc in reg:
            try:
                del reg[sc]
            except Exception:
                pass
    _write_run_registry_map(reg)


def _derive_identity_from_name(pdf_path: Path) -> tuple[str, str, str]:
    stem = pdf_path.stem
    parts = [p.strip() for p in stem.split("_") if p.strip()]
    program_name = ""
    vehicle_number = ""
    serial_component = ""
    if len(parts) >= 3:
        program_name, vehicle_number = parts[0], parts[1]
        serial_component = "_".join(parts[2:])
    elif len(parts) == 2:
        program_name = parts[0]
        serial_component = parts[1]
    elif parts:
        serial_component = parts[0]
    if not serial_component:
        # Fallback: use full stem
        serial_component = stem
    return program_name, vehicle_number, serial_component


_LAST_SYNC_SUMMARY: dict[str, str | int] | None = None
_LAST_SYNC_DETAILS: list[dict[str, str]] | None = None


def compute_workspace_sync(repo_root: Optional[Path] = None, terms_path: Optional[Path] = None) -> tuple[dict, list[dict]]:
    """Compute workspace sync status.

    Classifies PDFs in repo_root (recursively) as new/out-of-date/up-to-date by
    comparing:
      - run_registry run_date vs PDF modification time, and
      - schema-defined (Term Label, Data Group) pairs vs the values recorded
        in the master workbook for each serial.

    A PDF/serial is considered out-of-sync for "terms" when the current schema
    defines at least one term whose value is still missing (blank) in the
    master workbook for that EIDP. Merely re-saving the schema without adding
    new terms does not mark items as out-of-sync.
    """
    root = Path(repo_root) if repo_root else DEFAULT_PDF_DIR
    terms = Path(terms_path) if terms_path else DEFAULT_TERMS_XLSX
    # Ensure registry reflects run_data contents before comparing
    reg = ensure_run_registry_consistent()
    t_mtime = datetime.fromtimestamp(terms.stat().st_mtime) if terms.exists() else None
    # Schema term keys (used only as a guard; actual missing-term detection is
    # performed against master.xlsx via count_missing_terms_per_serial).
    schema_keys = _load_schema_term_keys(terms)

    pdfs = [p for p in root.rglob("*.pdf") if p.is_file()]
    details: list[dict[str, str]] = []
    new_count = outdated_pdf = outdated_terms = up_to_date = 0
    # Cache of per-serial missing-term counts so we only evaluate against the
    # master workbook once per serial even if multiple PDFs map to the same
    # EIDP.
    missing_counts_cache: dict[str, int] = {}

    for p in sorted(pdfs):
        try:
            prog, veh, serial = _derive_identity_from_name(p)
            info = reg.get(serial)
            run_dt = _parse_dt(info.get("run_date", "") if info else "")
            pdf_dt = datetime.fromtimestamp(p.stat().st_mtime)

            reason = "up_to_date"
            # No registry entry or unusable run date -> treat as "not yet run"
            if not info or not run_dt:
                reason = "new"
                new_count += 1
            else:
                # Registry knows about this serial; check PDF freshness first
                if pdf_dt > run_dt:
                    reason = "pdf_newer"
                    outdated_pdf += 1
                else:
                    # If we have a readable schema and master workbook, ask the
                    # master whether this EIDP still has missing values for any
                    # schema-defined terms. This lets incremental or "missing
                    # terms only" scans bring EIDPs back to an up-to-date state
                    # without requiring full re-extraction.
                    needs_terms = False
                    if schema_keys:
                        if serial not in missing_counts_cache:
                            try:
                                per_serial = count_missing_terms_per_serial([serial], terms)
                            except Exception:
                                per_serial = {}
                            missing_counts_cache[serial] = int(per_serial.get(serial, 0) or 0)
                        needs_terms = missing_counts_cache.get(serial, 0) > 0
                    if needs_terms:
                        reason = "terms_newer"
                        outdated_terms += 1
                    else:
                        up_to_date += 1

            details.append({
                "pdf": str(p),
                "serial_component": serial,
                "program_name": prog or (info.get("program_name") if info else "") or "",
                "vehicle_number": veh or (info.get("vehicle_number") if info else "") or "",
                "run_date": info.get("run_date") if info else "",
                "pdf_mtime": pdf_dt.strftime("%Y-%m-%d %H:%M:%S"),
                "terms_mtime": t_mtime.strftime("%Y-%m-%d %H:%M:%S") if t_mtime else "",
                "reason": reason,
            })
        except Exception:
            continue
    total = len(pdfs)
    summary = {
        "total": total,
        "new": new_count,
        "pdf_newer": outdated_pdf,
        "terms_newer": outdated_terms,
        "up_to_date": up_to_date,
        "last_sync": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "repo_root": str(root),
    }
    global _LAST_SYNC_SUMMARY, _LAST_SYNC_DETAILS
    _LAST_SYNC_SUMMARY, _LAST_SYNC_DETAILS = summary, details
    return summary, details


def get_last_sync() -> tuple[dict, list[dict]]:
    return _LAST_SYNC_SUMMARY or {}, _LAST_SYNC_DETAILS or []


def run_selected_pdfs(paths: list[Path], terms: Optional[Path] = None) -> subprocess.Popen:
    """Stage selected PDFs into a temp folder and run the scanner only on them.

    Original repository files remain untouched. Staged copies live only for the
    duration of the run and are deleted the next time this action executes.
    """
    terms = Path(terms) if terms else DEFAULT_TERMS_XLSX
    stage = ROOT / "user_inputs" / "Staging_Selected"
    try:
        if stage.exists():
            shutil.rmtree(stage, ignore_errors=True)
        stage.mkdir(parents=True, exist_ok=True)
        for p in paths:
            try:
                pp = Path(p)
                if pp.is_file():
                    shutil.copy2(str(pp), str(stage / pp.name))
            except Exception:
                pass
    except Exception:
        pass
    return run_scanner(terms, stage)


def _gather_serials_from_run(run_dir: Path) -> dict[str, tuple[str, str]]:
    """Return mapping {serial: (program, vehicle)} discovered in a run folder."""
    found: dict[str, tuple[str, str]] = {}

    def _record(row: Mapping[str, object]) -> None:
        pdf_info = row.get("pdf_info") if isinstance(row, Mapping) else {}
        if not isinstance(pdf_info, Mapping):
            return
        serial = str(pdf_info.get("serial_component") or pdf_info.get("serial_number") or "").strip()
        if not serial or serial in found:
            return
        program = str(pdf_info.get("program_name") or "").strip()
        vehicle = str(pdf_info.get("vehicle_number") or "").strip()
        found[serial] = (program, vehicle)

    agg = run_dir / "scan_results.json"
    if agg.exists():
        try:
            data = json.loads(agg.read_text(encoding="utf-8"))
            if isinstance(data, list):
                for row in data:
                    if isinstance(row, Mapping):
                        _record(row)
        except Exception:
            pass
    if not found:
        for child in run_dir.glob("scan_results_*.json"):
            if child.name.lower() == "scan_results.json":
                continue
            try:
                data = json.loads(child.read_text(encoding="utf-8"))
            except Exception:
                continue
            rows = data if isinstance(data, list) else ([data] if isinstance(data, Mapping) else [])
            for row in rows:
                if isinstance(row, Mapping):
                    _record(row)
            if found:
                break
    return found


def rebuild_registry_from_run_data() -> dict[str, dict[str, str]]:
    """Rebuild run_registry.csv by scanning run_data folders."""
    rows: dict[str, dict[str, str]] = {}
    roots = run_roots()
    if not any(r.exists() for r in roots):
        _write_run_registry_map(rows)
        return rows
    try:
        existing = _read_run_registry_map()
    except Exception:
        existing = {}
    for root in roots:
        if not root.exists():
            continue
        for run_dir in sorted(root.iterdir()):
            if not run_dir.is_dir():
                continue
            try:
                run_dt = datetime.strptime(run_dir.name, "%Y%m%d_%H%M%S")
            except Exception:
                run_dt = datetime.fromtimestamp(run_dir.stat().st_mtime)
            display_dt = run_dt.strftime("%Y-%m-%d %H:%M:%S")
            serial_meta = _gather_serials_from_run(run_dir)
            for serial, (program, vehicle) in serial_meta.items():
                prev = rows.get(serial) or existing.get(serial)
                if prev:
                    prev_dt = _parse_dt(prev.get("run_date", ""))
                    if prev_dt and prev_dt >= run_dt:
                        rows[serial] = prev
                        continue
                rel_run = str(run_dir.relative_to(ROOT))
                rows[serial] = {
                    "serial_component": serial,
                    "program_name": program,
                    "vehicle_number": vehicle,
                    "run_date": display_dt,
                    "run_folder": rel_run,
                }
    _write_run_registry_map(rows)
    return rows


def open_last_run_folder() -> None:
    run_dirs = []
    for root in run_roots():
        if root.exists():
            run_dirs.extend([p for p in root.iterdir() if p.is_dir()])
    if not run_dirs:
        raise FileNotFoundError(f"No run_data at {RUNS_DIR}")
    latest = max(run_dirs, key=lambda p: p.stat().st_mtime, default=None)
    if not latest:
        raise FileNotFoundError("No run folders found")
    open_path(latest)

def open_run_data_root() -> None:
    RUNS_DIR.mkdir(parents=True, exist_ok=True)
    open_path(RUNS_DIR)


def open_run_registry() -> None:
    reg_xlsx = registry_xlsx_for_read()
    reg_csv = registry_csv_for_read()
    target = reg_xlsx if reg_xlsx.exists() else (reg_csv if reg_csv.exists() else None)
    if not target:
        raise FileNotFoundError("No run registry found (create by running a scan)")
    open_path(target)


# --- Run-data maintenance helpers ---

def _resolve_run_folder(value: str) -> Path:
    """Resolve a run_folder value from registry or cell state to an absolute path.

    - Absolute paths are returned as-is.
    - Bare folder names like ``20250115_103000`` are treated as children of RUNS_DIR
      (falling back to legacy run_data if present).
    - Other relative paths are treated as ROOT-relative (e.g., ``run_data/...``).
    """
    try:
        p = Path(value)
    except Exception:
        return RUNS_DIR
    if p.is_absolute():
        return p
    # Single path component -> interpret as a run_data subfolder name
    if len(p.parts) == 1:
        candidate = RUNS_DIR / p
        if candidate.exists():
            return candidate
        legacy_candidate = LEGACY_RUNS_DIR / p
        if legacy_candidate.exists():
            return legacy_candidate
        return candidate
    # Otherwise treat as ROOT-relative (e.g. "run_data/..." or legacy Product_Data_File/...)
    return ROOT / p


def clear_stale_run_data() -> tuple[int, int]:
    """Delete run_data subfolders not referenced by run_registry OR master_cell_state.json.

    Returns (deleted_count, kept_count).
    """
    deleted = 0
    kept = 0

    # Get referenced folders from run_registry
    try:
        reg = _read_run_registry_map()
    except Exception:
        reg = {}
    referenced: set[Path] = set()
    for info in (reg or {}).values():
        try:
            rf = str(info.get("run_folder", ""))
        except Exception:
            rf = ""
        if not rf:
            continue
        try:
            referenced.add(_resolve_run_folder(rf).resolve())
        except Exception:
            pass

    # Also get referenced folders from master_cell_state.json
    try:
        from scripts.master_cell_state import get_referenced_run_folders
        state_folders = get_referenced_run_folders()
        for folder_name in state_folders:
            try:
                referenced.add(_resolve_run_folder(folder_name).resolve())
            except Exception:
                pass
    except Exception:
        # If cell state module isn't available, just use registry references
        pass
    try:
        any_root = False
        for root in run_roots():
            if not root.exists():
                continue
            any_root = True
            for child in root.iterdir():
                try:
                    if not child.is_dir():
                        continue
                    rchild = child.resolve()
                    if rchild in referenced:
                        kept += 1
                        continue
                    # Remove stale folder entirely
                    shutil.rmtree(str(child), ignore_errors=True)
                    deleted += 1
                except Exception:
                    # Ignore individual folder errors
                    pass
        if not any_root:
            return (0, 0)
    except Exception:
        pass
    return (deleted, kept)


def open_master_workbook() -> None:
    xlsx = master_xlsx_for_read()
    if not xlsx.exists():
        raise FileNotFoundError("master.xlsx not found (compile first)")
    open_path(xlsx)


# Deprecated: enrichment now handled during/after runs; external script removed
def enrich_run_registry() -> subprocess.Popen:  # type: ignore[dead-code]
    raise FileNotFoundError("enrich_run_registry is no longer available")


def open_plots_folder() -> None:
    PLOTS_DIR.mkdir(parents=True, exist_ok=True)
    open_path(plots_root_for_open())


def open_plots_summary() -> None:
    target_root = plots_root_for_open()
    target = target_root / "plots_summary.xlsx"
    if not target.exists():
        raise FileNotFoundError("No plots_summary.xlsx found (export first)")
    open_path(target)


def ensure_scaffold() -> None:
    (ROOT / "user_inputs").mkdir(parents=True, exist_ok=True)
    DEFAULT_PDF_DIR.mkdir(parents=True, exist_ok=True)
    RUNS_DIR.mkdir(parents=True, exist_ok=True)
    MASTER_DB_ROOT.mkdir(parents=True, exist_ok=True)
    if not SCANNER_ENV.exists():
        save_scanner_env({"QUIET": "1"})


# --- Health checks / analysis helpers ---

def check_environment() -> subprocess.Popen:
    """Spawn a short Python check that imports key packages and reports status.

    Returns a process whose stdout can be streamed for UI feedback.
    """
    py = resolve_project_python()
    code = (
        "import sys;\n"
        "print('[INFO] Checking environment for EIDAT...');\n"
        "mods=['fitz','pandas','openpyxl','matplotlib','easyocr'];\n"
        "ok=True;\n"
        "from importlib import import_module;\n"
        "for m in mods:\n"
        "    try:\n"
        "        import_module(m); print('[OK]', m)\n"
        "    except Exception as e:\n"
        "        ok=False; print('[MISS]', m, type(e).__name__, str(e));\n"
        "print('[DONE] ok='+str(ok))\n"
    )
    return spawn([py, "-c", code])


ID_COLS = ["Term Label", "Data Group"]
Y_AXIS_COL = "Y Axis Label"


def read_plot_terms_table() -> list[dict]:
    """Read plot_terms from XLSX or CSV into list of dict rows.

    Falls back to CSV if Excel stack isn't available.
    """
    xlsx = DEFAULT_PLOT_TERMS_XLSX
    csvp = xlsx.with_suffix(".csv")
    if xlsx.exists():
        try:
            import pandas as pd  # type: ignore
            df = pd.read_excel(xlsx)
            return df.fillna("").to_dict(orient="records")
        except Exception:
            pass
    if csvp.exists():
        try:
            import csv as _csv
            with open(csvp, newline="", encoding="utf-8") as f:
                r = _csv.DictReader(f)
                return [dict(row) for row in r]
        except Exception:
            pass
    return []


def write_plot_terms_table(rows: list[dict]) -> None:
    """Write plot terms table to Excel only (user_inputs/plot_terms.xlsx)."""
    xlsx = DEFAULT_PLOT_TERMS_XLSX
    if not rows:
        rows = [{
            "Plot?": "",
            "Plot Name": "",
            "Tie To Plot": "",
            "X Axis": "SN",
            **{k: "" for k in ID_COLS},
            "Min": "",
            "Max": "",
            Y_AXIS_COL: "",
            "Series Label": "",
        }]
    keys: list[str] = list(rows[0].keys())
    xlsx.parent.mkdir(parents=True, exist_ok=True)
    import pandas as _pd  # type: ignore
    df = _pd.DataFrame(rows, columns=keys)
    # Prefer xlsxwriter, fallback to openpyxl, else fail
    try:
        import xlsxwriter  # noqa: F401
        engine = "xlsxwriter"
    except Exception:
        engine = "openpyxl"
    try:
        with _pd.ExcelWriter(xlsx, engine=engine) as writer:  # type: ignore[arg-type]
            df.to_excel(writer, sheet_name="plot_terms", index=False)
        csvp = xlsx.with_suffix(".csv")
        try:
            csvp.unlink(missing_ok=True)
        except Exception:
            pass
    except Exception as e:
        raise RuntimeError(f"Unable to write plot_terms.xlsx: {e}")


def list_plot_series_options() -> list[dict]:
    """Return catalog of available plot series derived from plot_terms."""
    rows = read_plot_terms_table()
    catalog: list[dict] = []
    for r in rows:
        name = str(r.get("Plot Name") or "").strip()
        if not name:
            continue
        entry = {
            "name": name,
            "term_label": str(r.get("Term Label") or "").strip(),
            "data_group": str(r.get("Data Group") or "").strip(),
            "units": str(r.get("Units") or "").strip(),
            "default_y_axis": str(r.get(Y_AXIS_COL) or "").strip(),
        }
        catalog.append(entry)
    catalog.sort(key=lambda item: item["name"].lower())
    return catalog


def read_proposed_plots() -> list[dict]:
    path = DEFAULT_PROPOSED_PLOTS_JSON
    if not path.exists():
        return []
    try:
        with path.open("r", encoding="utf-8") as fh:
            data = json.load(fh)
    except Exception:
        return []
    if isinstance(data, dict):
        data = data.get("plots", [])
    if not isinstance(data, list):
        return []
    cleaned: list[dict] = []
    def _bool(val: object, default: bool = True) -> bool:
        if isinstance(val, bool):
            return val
        if isinstance(val, str):
            txt = val.strip().lower()
            if txt in ("1", "true", "yes", "y", "on"):
                return True
            if txt in ("0", "false", "no", "n", "off"):
                return False
        if isinstance(val, (int, float)):
            return val != 0
        return default

    for entry in data:
        if not isinstance(entry, dict):
            continue
        name = str(entry.get("name") or "").strip()
        y_axis = str(entry.get("y_axis") or "").strip()
        x_axis = str(entry.get("x_axis") or "SN").strip() or "SN"
        series = entry.get("series") or []
        if isinstance(series, str):
            series = [series]
        if not isinstance(series, list):
            series = []
        series_names = [str(s or "").strip() for s in series if str(s or "").strip()]
        cleaned.append({
            "name": name or "Plot",
            "series": series_names,
            "y_axis": y_axis,
            "x_axis": x_axis,
            "include_min": _bool(entry.get("include_min"), True),
            "include_max": _bool(entry.get("include_max"), True),
        })
    return cleaned


def write_proposed_plots(plots: list[dict]) -> None:
    """Persist proposed plot definitions to user_inputs."""
    DEFAULT_PROPOSED_PLOTS_JSON.parent.mkdir(parents=True, exist_ok=True)
    payload: list[dict] = []
    for entry in plots:
        if not isinstance(entry, dict):
            continue
        name = str(entry.get("name") or "").strip()
        y_axis = str(entry.get("y_axis") or "").strip()
        x_axis = str(entry.get("x_axis") or "SN").strip() or "SN"
        series = entry.get("series") or []
        if isinstance(series, str):
            series = [series]
        if not isinstance(series, list):
            series = []
        series_names = [str(s or "").strip() for s in series if str(s or "").strip()]
        payload.append({
            "name": name or "Plot",
            "series": series_names,
            "y_axis": y_axis,
            "x_axis": x_axis,
            "include_min": bool(entry.get("include_min", True)),
            "include_max": bool(entry.get("include_max", True)),
        })
    DEFAULT_PROPOSED_PLOTS_JSON.write_text(json.dumps(payload, indent=2), encoding="utf-8")
