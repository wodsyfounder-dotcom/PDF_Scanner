from __future__ import annotations

import os
import sys
import shutil
import subprocess
from pathlib import Path
from typing import Dict, Iterable, Optional


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_TERMS_XLSX = ROOT / "user_inputs" / "terms.xlsx"
DEFAULT_PLOT_TERMS_XLSX = ROOT / "user_inputs" / "plot_terms.xlsx"
DEFAULT_PDF_DIR = ROOT / "user_inputs" / "EIDP_Import_Docs"
DEFAULT_SCANNED_DIR = ROOT / "user_inputs" / "Scanned_Docs"
SCANNER_ENV = ROOT / "user_inputs" / "scanner.env"
APP_ENTRY = ROOT / "Application" / "eidp_term_scanner.py"
RUNS_DIR = ROOT / "Product_Data_File" / "run_data"
PLOTS_DIR = ROOT / "Product_Data_File" / "plots"


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


def run_scanner(terms: Path, pdf_dir: Path, scanned_dir: Path) -> subprocess.Popen:
    if sys.platform.startswith("win"):
        cmd = [
            "cmd.exe", "/c", str(ROOT / "run.bat"),
            "--input", str(terms),
            "--pdf-folder", str(pdf_dir),
            "--scanned-folder", str(scanned_dir),
            "--quiet",
        ]
        return spawn(cmd)
    else:
        py = resolve_project_python()
        cmd = [py, str(APP_ENTRY), "--input", str(terms), "--pdf-folder", str(pdf_dir), "--scanned-folder", str(scanned_dir), "--quiet"]
        return spawn(cmd)


def run_script(script_rel_path: str, *args: str) -> subprocess.Popen:
    py = resolve_project_python()
    script = ROOT / script_rel_path
    if not script.exists():
        raise FileNotFoundError(f"Missing script: {script}")
    return spawn([py, str(script), *args])


def generate_terms() -> subprocess.Popen:
    return run_script("scripts/generate_terms_schema.py")


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


def open_master_workbook() -> None:
    xlsx = ROOT / "Product_Data_File" / "master.xlsx"
    csv = ROOT / "Product_Data_File" / "master.csv"
    target = xlsx if xlsx.exists() else (csv if csv.exists() else None)
    if not target:
        raise FileNotFoundError("No master workbook found (compile first)")
    open_path(target)


def enrich_run_registry() -> subprocess.Popen:
    return run_script("scripts/enrich_run_registry.py")


def open_plots_folder() -> None:
    PLOTS_DIR.mkdir(parents=True, exist_ok=True)
    open_path(PLOTS_DIR)


def open_plots_summary() -> None:
    target = ROOT / "Product_Data_File" / "plots_summary.xlsx"
    if not target.exists():
        raise FileNotFoundError("No plots_summary.xlsx found (export first)")
    open_path(target)


def ensure_scaffold() -> None:
    (ROOT / "user_inputs").mkdir(parents=True, exist_ok=True)
    DEFAULT_PDF_DIR.mkdir(parents=True, exist_ok=True)
    DEFAULT_SCANNED_DIR.mkdir(parents=True, exist_ok=True)
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


ID_COLS = ["Term", "Grouping", "Row Label", "Column Label", "Units"]


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
    """Write plot terms table back to CSV for robustness.

    We prefer not to depend on Excel writers here; the plotting script
    accepts CSV as well.
    """
    target = DEFAULT_PLOT_TERMS_XLSX.with_suffix(".csv")
    import csv as _csv
    if not rows:
        # Ensure header exists if empty write attempted
        rows = [{"Plot?": "", "Plot Name": "", "Tie To Plot": "", "X Axis": "SN", **{k: "" for k in ID_COLS}, "Min": "", "Max": "", "Series Label": ""}]
    keys: list[str] = list(rows[0].keys())
    target.parent.mkdir(parents=True, exist_ok=True)
    with open(target, "w", newline="", encoding="utf-8") as f:
        w = _csv.DictWriter(f, fieldnames=keys)
        w.writeheader()
        for row in rows:
            w.writerow({k: row.get(k, "") for k in keys})


def set_plot_flags(active_keys: set[tuple[str, str, str, str, str]]) -> None:
    """Update 'Plot?' column to 'Y' for rows whose ID tuple is in active_keys, else ''.

    ID tuple order: (Term, Grouping, Row Label, Column Label, Units)
    """
    rows = read_plot_terms_table()
    if not rows:
        return
    updated: list[dict] = []
    for r in rows:
        key = tuple(str(r.get(k, "") or "").strip() for k in ID_COLS)
        r["Plot?"] = "Y" if key in active_keys else (r.get("Plot?", "") if key not in active_keys else "")
        updated.append(r)
    write_plot_terms_table(updated)
