from __future__ import annotations

import json
import os
import sys
import shutil
import subprocess
from pathlib import Path
from typing import Dict, Iterable, Mapping, Optional


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_TERMS_XLSX = ROOT / "user_inputs" / "terms.schema.smartsnap.xlsx"
DEFAULT_PLOT_TERMS_XLSX = ROOT / "user_inputs" / "plot_terms.xlsx"
DEFAULT_PROPOSED_PLOTS_JSON = ROOT / "user_inputs" / "proposed_plots.json"
# Default repository root where PDFs may live (user-organized, nested or flat)
DEFAULT_REPO_ROOT = ROOT / "Data Packages"
DEFAULT_PDF_DIR = DEFAULT_REPO_ROOT
SCANNER_ENV = ROOT / "user_inputs" / "scanner.env"
APP_ENTRY = ROOT / "Application" / "eidp_term_scanner.py"
RUNS_DIR = ROOT / "Product_Data_File" / "run_data"
PLOTS_DIR = ROOT / "Product_Data_File" / "plots"
TERMS_TEMPLATE_SHEET = "Template"
TERMS_SCHEMA_COLUMNS = [
    "Data Group",
    "Term Label",
    "Term",
    "Pages",
    "Mode",
    "Line",
    "Column",
    "Anchor",
    "FieldIndex",
    "FieldSplit",
    "Return",
    "Units",
    "Range (min)",
    "Range (max)",
    "Format",
    "GroupAfter",
    "GroupBefore",
    "Smart Snap Type",
    "Secondary Term",
    "Smart Position",
]
TERMS_MODE_CHOICES = ["smart", "full table"]
TERMS_SMART_TYPE_CHOICES = ["", "auto", "number", "date", "time", "title"]


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
        "XY_FUZZ",
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
    cand = _venv_python_from(ROOT / ".venv")
    if cand.exists():
        return str(cand)
    return sys.executable


def _base_env() -> Dict[str, str]:
    env = os.environ.copy()
    # Merge scanner.env values for direct Python invocations
    env.update(parse_scanner_env(SCANNER_ENV))
    # Ensure vendored site-packages are importable as fallback
    env["PYTHONPATH"] = str(ROOT / "Lib" / "site-packages") + os.pathsep + env.get("PYTHONPATH", "")
    env.setdefault("QUIET", "1")
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
    script = ROOT / script_rel_path
    if not script.exists():
        raise FileNotFoundError(f"Missing script: {script}")
    return spawn([py, str(script), *args])


def generate_terms() -> subprocess.Popen:
    return run_script("scripts/generate_terms_schema_smartsnap.py")


def compile_master() -> subprocess.Popen:
    return run_script("scripts/compile_master.py")


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
        if not missing_columns:
            return False

        # Add missing column headers to row 1
        start_col = len(existing_headers) + 1
        for idx, col_name in enumerate(missing_columns, start=start_col):
            ws.cell(row=1, column=idx, value=col_name)

        wb.save(tgt)
        return True
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


def _read_run_registry_map() -> dict[str, dict[str, str]]:
    """Return mapping {serial_component: {run_date, run_folder, program_name, vehicle_number}}.

    Reads Product_Data_File/run_registry.(xlsx|csv). Missing file -> empty map.
    """
    reg_dir = ROOT / "Product_Data_File"
    rx = reg_dir / "run_registry.xlsx"
    rc = reg_dir / "run_registry.csv"
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
                with rc.open("w", encoding="utf-8", newline="") as f:
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
    reg_dir = ROOT / "Product_Data_File"
    reg_dir.mkdir(parents=True, exist_ok=True)
    rx = reg_dir / "run_registry.xlsx"
    rc = reg_dir / "run_registry.csv"
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
    comparing run_registry run_date and the mtime of the PDF and terms.
    """
    root = Path(repo_root) if repo_root else DEFAULT_PDF_DIR
    terms = Path(terms_path) if terms_path else DEFAULT_TERMS_XLSX
    # Ensure registry reflects run_data contents before comparing
    reg = ensure_run_registry_consistent()
    t_mtime = datetime.fromtimestamp(terms.stat().st_mtime) if terms.exists() else None

    pdfs = [p for p in root.rglob("*.pdf") if p.is_file()]
    details: list[dict[str, str]] = []
    new_count = outdated_pdf = outdated_terms = up_to_date = 0
    for p in sorted(pdfs):
        try:
            prog, veh, serial = _derive_identity_from_name(p)
            info = reg.get(serial)
            run_dt = _parse_dt(info.get("run_date", "") if info else "")
            pdf_dt = datetime.fromtimestamp(p.stat().st_mtime)
            reason = "up_to_date"
            if not run_dt:
                reason = "new"
                new_count += 1
            else:
                if pdf_dt > run_dt:
                    reason = "pdf_newer"
                    outdated_pdf += 1
                elif t_mtime and t_mtime > run_dt:
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


def rebuild_registry_from_run_data() -> dict[str, dict[str, str]]:
    """Rebuild or update run_registry.csv by scanning run_data folders.

    This leverages ensure_run_registry_consistent(), which walks run_data and
    merges/updates entries based on discovered outputs. Returns the final map.
    """
    return ensure_run_registry_consistent()


def open_last_run_folder() -> None:
    if not RUNS_DIR.exists():
        raise FileNotFoundError(f"No run_data at {RUNS_DIR}")
    latest = max((p for p in RUNS_DIR.iterdir() if p.is_dir()), key=lambda p: p.stat().st_mtime, default=None)
    if not latest:
        raise FileNotFoundError("No run folders found")
    open_path(latest)


def open_run_registry() -> None:
    reg_xlsx = ROOT / "Product_Data_File" / "run_registry.xlsx"
    reg_csv = ROOT / "Product_Data_File" / "run_registry.csv"
    target = reg_xlsx if reg_xlsx.exists() else (reg_csv if reg_csv.exists() else None)
    if not target:
        raise FileNotFoundError("No run registry found (create by running a scan)")
    open_path(target)


# --- Run-data maintenance helpers ---

def _resolve_run_folder(value: str) -> Path:
    try:
        p = Path(value)
    except Exception:
        return RUNS_DIR
    if not p.is_absolute():
        p = ROOT / p
    return p


def clear_stale_run_data() -> tuple[int, int]:
    """Delete run_data subfolders not referenced by run_registry.

    Returns (deleted_count, kept_count).
    """
    runs_root = RUNS_DIR
    deleted = 0
    kept = 0
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
    try:
        if not runs_root.exists():
            return (0, 0)
        for child in runs_root.iterdir():
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
    except Exception:
        pass
    return (deleted, kept)


def open_master_workbook() -> None:
    xlsx = ROOT / "Product_Data_File" / "master.xlsx"
    csv = ROOT / "Product_Data_File" / "master.csv"
    target = xlsx if xlsx.exists() else (csv if csv.exists() else None)
    if not target:
        raise FileNotFoundError("No master workbook found (compile first)")
    open_path(target)


# Deprecated: enrichment now handled during/after runs; external script removed
def enrich_run_registry() -> subprocess.Popen:  # type: ignore[dead-code]
    raise FileNotFoundError("enrich_run_registry is no longer available")


def open_plots_folder() -> None:
    PLOTS_DIR.mkdir(parents=True, exist_ok=True)
    open_path(PLOTS_DIR)


def open_plots_summary() -> None:
    target = PLOTS_DIR / "plots_summary.xlsx"
    if not target.exists():
        raise FileNotFoundError("No plots_summary.xlsx found (export first)")
    open_path(target)


def ensure_scaffold() -> None:
    (ROOT / "user_inputs").mkdir(parents=True, exist_ok=True)
    DEFAULT_PDF_DIR.mkdir(parents=True, exist_ok=True)
    (ROOT / "Product_Data_File" / "run_data").mkdir(parents=True, exist_ok=True)
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
        })
    DEFAULT_PROPOSED_PLOTS_JSON.write_text(json.dumps(payload, indent=2), encoding="utf-8")
