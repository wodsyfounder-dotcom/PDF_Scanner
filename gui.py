#!/usr/bin/env python3
"""
Simple native GUI (Tkinter) launcher for the EIDP Term Scanner.

- Pure Python standard library (Tkinter) - no external EXEs or downloads.
- Loads env from user_inputs/scanner.env.
- Lets you pick Terms file, PDFs folder, Scanned folder.
- Runs the scanner in a background thread, streams stdout/stderr to the UI.
- Provides shortcuts for terms template management and run output workbooks.

Usage:
  python gui.py
"""
from __future__ import annotations

import os
import sys
import threading
import subprocess
import shutil
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


ROOT = Path(__file__).resolve().parent
DEFAULT_TERMS_XLSX = ROOT / "user_inputs" / "terms.xlsx"
DEFAULT_PLOT_TERMS_XLSX = ROOT / "user_inputs" / "plot_terms.xlsx"
DEFAULT_PDF_DIR = ROOT / "user_inputs" / "EIDP_Import_Docs"
DEFAULT_SCANNED_DIR = ROOT / "user_inputs" / "Scanned_Docs"
SCANNER_ENV = ROOT / "user_inputs" / "scanner.env"
APP_ENTRY = ROOT / "Application" / "eidp_term_scanner.py"
RUNS_DIR = ROOT / "Product_Data_File" / "run_data"
PLOTS_DIR = ROOT / "Product_Data_File" / "plots"

def _resolve_project_python() -> str:
    """Pick the Python interpreter to run helper scripts.

    Priority:
      1) VENV_DIR from scanner.env if it contains a Python
      2) Project .venv if present
      3) Current interpreter (sys.executable)
    """
    try:
        env = parse_scanner_env(SCANNER_ENV)
    except Exception:
        env = {}
    vdir = env.get("VENV_DIR", "").strip()
    def vpy(path: Path) -> Path:
        if os.name == "nt":
            return path / "Scripts" / "python.exe"
        return path / "bin" / "python"
    if vdir:
        cand = vpy(Path(vdir))
        if cand.exists():
            return str(cand)
    # default .venv
    cand = vpy(ROOT / ".venv")
    if cand.exists():
        return str(cand)
    return sys.executable


class Tooltip:
    def __init__(self, widget, text: str):
        self.widget = widget
        self.text = text
        self.tip = None
        self.widget.bind("<Enter>", self._show)
        self.widget.bind("<Leave>", self._hide)
        self.widget.bind("<Motion>", self._move)

    def _show(self, event=None):
        if self.tip or not self.text:
            return
        x = self.widget.winfo_rootx() + 10
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        lbl = tk.Label(self.tip, text=self.text, justify=tk.LEFT,
                       background="#333", foreground="#fff",
                       relief=tk.SOLID, borderwidth=1,
                       font=("Segoe UI", 9), padx=8, pady=4)
        lbl.pack()

    def _hide(self, event=None):
        if self.tip is not None:
            try:
                self.tip.destroy()
            except Exception:
                pass
            self.tip = None

    def _move(self, event):
        if self.tip is None:
            return
        x = event.x_root + 10
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self.tip.wm_geometry(f"+{x}+{y}")


def parse_scanner_env(path: Path) -> dict[str, str]:
    env: dict[str, str] = {}
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
        # Strip inline comments and whitespace
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
    return env


class Runner:
    def __init__(self, log_cb):
        self.proc: subprocess.Popen | None = None
        self.thread: threading.Thread | None = None
        self.log_cb = log_cb
        self._stop = threading.Event()

    def start(self, terms: Path, pdf_dir: Path, scanned: Path):
        if self.proc is not None:
            return
        # Prefer running the project-provided run.bat so environment setup
        # (venv, vendored site-packages, scanner.env) matches CLI runs.
        if os.name == "nt":
            runbat = str(ROOT / "run.bat")
            cmd = [
                "cmd.exe", "/c", runbat,
                "--input", str(terms),
                "--pdf-folder", str(pdf_dir),
                "--scanned-folder", str(scanned),
                "--quiet",
            ]
        else:
            # Fallback: invoke the Python entry directly on non-Windows
            cmd = [
                sys.executable,
                str(APP_ENTRY),
                "--input", str(terms),
                "--pdf-folder", str(pdf_dir),
                "--scanned-folder", str(scanned),
                "--quiet",
            ]
        env = os.environ.copy()
        # Let run.bat parse scanner.env; still merge here for direct Python fallback
        env.update(parse_scanner_env(SCANNER_ENV))
        # Ensure vendored packages are visible for direct Python fallback
        env["PYTHONPATH"] = str(ROOT / "Lib" / "site-packages") + os.pathsep + env.get("PYTHONPATH", "")
        # Respect GUI settings via scanner.env; default to quiet if not set
        env.setdefault("QUIET", "1")

        try:
            self.proc = subprocess.Popen(
                cmd,
                cwd=str(ROOT),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
            )
        except Exception as e:
            self.log_cb(f"[ERROR] Failed to start: {e}\n")
            self.proc = None
            return

        def _pump():
            try:
                assert self.proc is not None
                for line in self.proc.stdout:  # type: ignore[arg-type]
                    if self._stop.is_set():
                        break
                    self.log_cb(line)
            except Exception as e:
                self.log_cb(f"[WARN] Reader error: {e}\n")
            finally:
                try:
                    if self.proc is not None:
                        self.proc.wait(timeout=1)
                except Exception:
                    pass
                self.log_cb("\n[INFO] Run finished.\n")
                self.proc = None

        self.thread = threading.Thread(target=_pump, daemon=True)
        self.thread.start()

    def stop(self):
        self._stop.set()
        try:
            if self.proc and self.proc.poll() is None:
                self.proc.terminate()
        except Exception:
            pass


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("EIDP Term Scanner")
        self.geometry("1040x720")
        self.minsize(920, 600)
        self.style = ttk.Style(self)
        try:
            if sys.platform.startswith("win") and "vista" in self.style.theme_names():
                self.style.theme_use("vista")
            elif "clam" in self.style.theme_names():
                self.style.theme_use("clam")
        except Exception:
            pass
        self.style.configure("Section.TLabelframe", padding=(12, 10))
        self.style.configure("Section.TLabelframe.Label", font=("Segoe UI", 11, "bold"))
        self.style.configure("Primary.TButton", padding=(12, 6), font=("Segoe UI", 10))
        self.style.configure("Secondary.TButton", padding=(12, 6), font=("Segoe UI", 10))
        self.style.configure("Danger.TButton", padding=(12, 6), font=("Segoe UI", 10))
        self.style.map("Danger.TButton", foreground=[("!disabled", "#b00020")])
        self.runner = Runner(self.append_log)
        self._run_active = False
        self._env_cache: dict[str, str] = parse_scanner_env(SCANNER_ENV)
        self._build_ui()
        self.after(500, self._poll_runner)

    def _build_ui(self):
        padx, pady = 12, 10

        # Header bar
        header = ttk.Frame(self)
        header.pack(fill=tk.X, padx=12, pady=(10, 0))
        ttk.Label(header, text="EIDP Term Scanner", font=("Segoe UI", 14, "bold")).grid(row=0, column=0, sticky=tk.W)
        self.var_interp = tk.StringVar(value=_resolve_project_python())
        ttk.Label(header, textvariable=self.var_interp).grid(row=0, column=1, sticky=tk.E, padx=(8,0))
        header.columnconfigure(0, weight=1)

        # Notebook with clear sections
        nb = ttk.Notebook(self)
        nb.pack(fill=tk.X, padx=12, pady=(8, 6))
        tab_setup = ttk.Frame(nb)
        tab_plot = ttk.Frame(nb)
        tab_out = ttk.Frame(nb)
        nb.add(tab_setup, text="Setup & Scan")
        nb.add(tab_plot, text="Plotting")
        nb.add(tab_out, text="Outputs")

        # Setup & Scan
        # Start Here (checklist)
        start = ttk.LabelFrame(tab_setup, text="Start Here", style="Section.TLabelframe")
        start.pack(fill=tk.X, padx=12, pady=10)
        ttk.Label(start, text="1) Install environment →", font=("Segoe UI", 10)).grid(row=0, column=0, sticky=tk.W, padx=(8,6), pady=4)
        btn = ttk.Button(start, text="Install", command=self._install_full, style="Primary.TButton")
        btn.grid(row=0, column=1, padx=6, pady=4, sticky=tk.W)
        Tooltip(btn, "Creates a project venv (.venv) and installs packages.")
        ttk.Label(start, text="2) Pick folders →", font=("Segoe UI", 10)).grid(row=1, column=0, sticky=tk.W, padx=(8,6), pady=4)
        btn = ttk.Button(start, text="Set Inputs", command=self._pick_pdfs, style="Secondary.TButton")
        btn.grid(row=1, column=1, padx=6, pady=4, sticky=tk.W)
        Tooltip(btn, "Choose PDFs folder (scanned dest auto-filled).")
        ttk.Label(start, text="3) Create/Open terms →", font=("Segoe UI", 10)).grid(row=2, column=0, sticky=tk.W, padx=(8,6), pady=4)
        btn = ttk.Button(start, text="Open Terms", command=self._open_terms_spreadsheet, style="Secondary.TButton")
        btn.grid(row=2, column=1, padx=6, pady=4, sticky=tk.W)
        Tooltip(btn, "Open terms.xlsx to review/edit terms and pages.")
        ttk.Label(start, text="4) Run scan →", font=("Segoe UI", 10)).grid(row=3, column=0, sticky=tk.W, padx=(8,6), pady=4)
        btn = ttk.Button(start, text="Start Scan", command=self._run, style="Primary.TButton")
        btn.grid(row=3, column=1, padx=6, pady=4, sticky=tk.W)
        Tooltip(btn, "Scan PDFs for terms and write run results.")
        ttk.Label(start, text="5) Compile master →", font=("Segoe UI", 10)).grid(row=4, column=0, sticky=tk.W, padx=(8,6), pady=4)
        btn = ttk.Button(start, text="Compile", command=self._compile_master, style="Secondary.TButton")
        btn.grid(row=4, column=1, padx=6, pady=4, sticky=tk.W)
        Tooltip(btn, "Build master.xlsx from run history.")

        lf_env = ttk.LabelFrame(tab_setup, text="1) Environment", style="Section.TLabelframe")
        lf_env.pack(fill=tk.X, padx=padx, pady=pady)
        self.btn_install = ttk.Button(lf_env, text="Install Environment", command=self._install_full, style="Primary.TButton")
        self.btn_install.grid(row=0, column=0, padx=(0,8), pady=6)
        Tooltip(self.btn_install, "Creates .venv and installs packages via install.bat.")
        ttk.Button(lf_env, text="Open scanner.env", command=self._open_scanner_env, style="Secondary.TButton").grid(row=0, column=1, padx=8, pady=6)
        self.lbl_env_status = ttk.Label(lf_env, text="Env: Unknown", foreground="#5b6b7a")
        self.lbl_env_status.grid(row=0, column=2, padx=8, pady=6, sticky=tk.W)

        lf_inputs = ttk.LabelFrame(tab_setup, text="2) Scanner Inputs", style="Section.TLabelframe")
        lf_inputs.pack(fill=tk.X, padx=padx, pady=pady)
        ttk.Label(lf_inputs, text="PDFs folder:").grid(row=0, column=0, sticky=tk.W)
        self.var_pdfs = tk.StringVar(value=str(DEFAULT_PDF_DIR))
        ttk.Entry(lf_inputs, textvariable=self.var_pdfs, width=80).grid(row=0, column=1, sticky=tk.EW, padx=(0, 6))
        ttk.Button(lf_inputs, text="Browse", command=self._pick_pdfs).grid(row=0, column=2)
        ttk.Label(lf_inputs, text="Scanned folder:").grid(row=1, column=0, sticky=tk.W)
        self.var_scanned = tk.StringVar(value=str(DEFAULT_SCANNED_DIR))
        ttk.Entry(lf_inputs, textvariable=self.var_scanned, width=80).grid(row=1, column=1, sticky=tk.EW, padx=(0, 6))
        ttk.Button(lf_inputs, text="Browse", command=self._pick_scanned).grid(row=1, column=2)
        lf_inputs.columnconfigure(1, weight=1)

        lf_terms = ttk.LabelFrame(tab_setup, text="3) Terms Spreadsheet", style="Section.TLabelframe")
        lf_terms.pack(fill=tk.X, padx=padx, pady=pady)
        self.var_terms = tk.StringVar(value=str(DEFAULT_TERMS_XLSX))
        ttk.Label(lf_terms, text="Spreadsheet path:").grid(row=0, column=0, sticky=tk.W, padx=(8, 6), pady=6)
        ttk.Entry(lf_terms, textvariable=self.var_terms, width=80).grid(row=0, column=1, sticky=tk.EW, padx=(0, 6), pady=6)
        ttk.Button(lf_terms, text="Browse", command=self._pick_terms, style="Secondary.TButton").grid(row=0, column=2, padx=(0, 6), pady=6, sticky=tk.EW)
        ttk.Button(lf_terms, text="Open Spreadsheet", command=self._open_terms_spreadsheet, style="Secondary.TButton").grid(row=0, column=3, padx=(0, 8), pady=6, sticky=tk.EW)
        self.btn_terms_create = ttk.Button(lf_terms, text="Create / Refresh", command=self._generate_terms_spreadsheet, style="Primary.TButton")
        self.btn_terms_create.grid(row=1, column=1, columnspan=3, sticky=tk.W, padx=(0, 8), pady=(0, 8))
        Tooltip(self.btn_terms_create, "Generate a terms.xlsx template under user_inputs.")
        self.lbl_terms_status = ttk.Label(lf_terms, text="Terms: Not found", foreground="#a12d2d")
        self.lbl_terms_status.grid(row=2, column=1, sticky=tk.W, padx=(0,8), pady=(0,4))

        lf_run = ttk.LabelFrame(tab_setup, text="4) Run Scan", style="Section.TLabelframe")
        lf_run.pack(fill=tk.X, padx=padx, pady=pady)
        lf_run.columnconfigure((0, 1, 2, 3), weight=1, uniform="run")
        self.btn_run = ttk.Button(lf_run, text="Start Scan", command=self._run, style="Primary.TButton")
        self.btn_run.grid(row=0, column=0, sticky=tk.EW, padx=(0, 8))
        self.btn_stop = ttk.Button(lf_run, text="Stop Scan", command=self._stop, state=tk.DISABLED, style="Danger.TButton")
        self.btn_stop.grid(row=0, column=1, sticky=tk.EW, padx=8)
        ttk.Button(lf_run, text="Settings", command=self._open_settings, style="Secondary.TButton").grid(row=0, column=2, sticky=tk.EW, padx=8)
        ttk.Button(lf_run, text="Open Last Run Folder", command=self._open_last_run, style="Secondary.TButton").grid(row=0, column=3, sticky=tk.EW, padx=(8, 0))

        # Plotting tab
        plotting = ttk.LabelFrame(tab_plot, text="Plotting", style="Section.TLabelframe")
        plotting.pack(fill=tk.X, padx=padx, pady=pady)
        plotting.columnconfigure(1, weight=1)
        self.var_plot_terms = tk.StringVar(value=str(DEFAULT_PLOT_TERMS_XLSX))
        ttk.Label(plotting, text="Plot terms file:").grid(row=0, column=0, sticky=tk.W, padx=(8,6), pady=6)
        ttk.Entry(plotting, textvariable=self.var_plot_terms, width=80).grid(row=0, column=1, sticky=tk.EW, padx=(0,6), pady=6)
        ttk.Button(plotting, text="Open", command=self._open_plot_terms, style="Secondary.TButton").grid(row=0, column=2, padx=(0,6), pady=6, sticky=tk.EW)
        self.btn_plot_terms_create = ttk.Button(plotting, text="Create / Refresh", command=self._generate_plot_terms, style="Primary.TButton")
        self.btn_plot_terms_create.grid(row=1, column=1, sticky=tk.W, padx=(0,8), pady=(0,8))
        Tooltip(self.btn_plot_terms_create, "Create plot_terms.xlsx from master.xlsx.")
        self.btn_plot_generate = ttk.Button(plotting, text="Generate Plots", command=self._generate_plots, style="Primary.TButton")
        self.btn_plot_generate.grid(row=1, column=2, sticky=tk.EW, padx=(0,6), pady=(0,8))
        Tooltip(self.btn_plot_generate, "Read plot_terms and master to produce PNG plots.")
        self.btn_plots_open = ttk.Button(plotting, text="Open Plots Folder", command=self._open_plots_folder, style="Secondary.TButton")
        self.btn_plots_open.grid(row=1, column=3, sticky=tk.EW, padx=(0,8), pady=(0,8))
        self.btn_plot_summary = ttk.Button(plotting, text="Create Plot Summary", command=self._export_plot_summary, style="Primary.TButton")
        self.btn_plot_summary.grid(row=2, column=1, sticky=tk.W, padx=(0,8), pady=(0,8))
        Tooltip(self.btn_plot_summary, "Embed generated plots into plots_summary.xlsx.")
        self.btn_open_plot_summary = ttk.Button(plotting, text="Open Plot Summary", command=self._open_plot_summary, style="Secondary.TButton")
        self.btn_open_plot_summary.grid(row=2, column=2, sticky=tk.EW, padx=(0,6), pady=(0,8))
        self.lbl_plot_cfg_status = ttk.Label(plotting, text="Plot Config: Not created", foreground="#a12d2d")
        self.lbl_plot_cfg_status.grid(row=2, column=3, sticky=tk.E, padx=(0,8))

        # Page Table Extraction
        lf_tables = ttk.LabelFrame(tab_plot, text="Page Table Extraction", style="Section.TLabelframe")
        lf_tables.pack(fill=tk.X, padx=padx, pady=pady)
        lf_tables.columnconfigure(1, weight=1)
        try:
            first_pdf = next((p for p in (DEFAULT_PDF_DIR).glob("*.pdf")), None)
            default_pdf = str(first_pdf) if first_pdf else ""
        except Exception:
            default_pdf = ""
        self.var_ext_pdf = tk.StringVar(value=default_pdf)
        self.var_ext_pages = tk.StringVar(value="1")
        ttk.Label(lf_tables, text="PDF (from EIDP_Import_Docs):").grid(row=0, column=0, sticky=tk.W, padx=(8,6), pady=6)
        ttk.Entry(lf_tables, textvariable=self.var_ext_pdf, width=80).grid(row=0, column=1, sticky=tk.EW, padx=(0,6), pady=6)
        ttk.Button(lf_tables, text="Browse", command=self._pick_ext_pdf, style="Secondary.TButton").grid(row=0, column=2, padx=(0,6), pady=6, sticky=tk.EW)
        ttk.Label(lf_tables, text="Pages (e.g., 1 or 1,3-5):").grid(row=1, column=0, sticky=tk.W, padx=(8,6), pady=6)
        ttk.Entry(lf_tables, textvariable=self.var_ext_pages, width=30).grid(row=1, column=1, sticky=tk.W, padx=(0,6), pady=6)
        self.btn_extract_tables = ttk.Button(lf_tables, text="Extract Page Tables", command=self._extract_page_tables, style="Primary.TButton")
        self.btn_extract_tables.grid(row=2, column=1, sticky=tk.W, padx=(0,8), pady=(0,8))
        Tooltip(self.btn_extract_tables, "Parse selected pages into Product_Data_File/tables/<name>_tables.xlsx")

        # Outputs tab
        outputs = ttk.LabelFrame(tab_out, text="Outputs", style="Section.TLabelframe")
        outputs.pack(fill=tk.X, padx=padx, pady=pady)
        outputs.columnconfigure((0, 1, 2), weight=1, uniform="outputs")
        self.btn_open_registry = ttk.Button(outputs, text="Open Run Registry", command=self._open_run_registry, style="Secondary.TButton")
        self.btn_open_registry.grid(row=0, column=0, sticky=tk.EW, padx=6, pady=6)
        Tooltip(self.btn_open_registry, "Open Product_Data_File/run_registry.xlsx")
        self.btn_compile_master = ttk.Button(outputs, text="Compile Master Workbook", command=self._compile_master, style="Secondary.TButton")
        self.btn_compile_master.grid(row=0, column=1, sticky=tk.EW, padx=6, pady=6)
        Tooltip(self.btn_compile_master, "Build Product_Data_File/master.xlsx from run history.")
        self.btn_open_master = ttk.Button(outputs, text="Open Master Workbook", command=self._open_master, style="Secondary.TButton")
        self.btn_open_master.grid(row=0, column=2, sticky=tk.EW, padx=6, pady=6)
        Tooltip(self.btn_open_master, "Open Product_Data_File/master.xlsx")
        self.lbl_master_status = ttk.Label(outputs, text="Master: Not built", foreground="#a12d2d")
        self.lbl_master_status.grid(row=1, column=0, columnspan=3, sticky=tk.W, padx=6)

        # Log area
        log_frame = ttk.Frame(self, padding=(12, 0, 12, 10))
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.txt = tk.Text(log_frame, wrap="word", height=24, borderwidth=1, relief=tk.SOLID, font=("Consolas", 10))
        self.txt.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        scroll_y = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.txt.yview)
        scroll_y.pack(fill=tk.Y, side=tk.RIGHT)
        self.txt.configure(state=tk.DISABLED, yscrollcommand=scroll_y.set)

        # Status bar
        self.status = tk.StringVar(value="Ready.")
        ttk.Label(self, textvariable=self.status, anchor=tk.W).pack(fill=tk.X, padx=12, pady=(0, 10))

        # Wire path changes to live refresh
        self.var_terms.trace_add('write', lambda *_: self._refresh_ui_state())
        self.var_pdfs.trace_add('write', lambda *_: self._refresh_ui_state())
        self.var_scanned.trace_add('write', lambda *_: self._refresh_ui_state())
        self.var_plot_terms.trace_add('write', lambda *_: self._refresh_ui_state())
        self.var_ext_pdf.trace_add('write', lambda *_: self._refresh_ui_state())
        self.var_ext_pages.trace_add('write', lambda *_: self._refresh_ui_state())
        self._refresh_ui_state()

    def append_log(self, s: str):
        self.txt.configure(state=tk.NORMAL)
        self.txt.insert(tk.END, s)
        self.txt.see(tk.END)
        self.txt.configure(state=tk.DISABLED)

    def _pick_terms(self):
        path = filedialog.askopenfilename(title="Select terms file", initialdir=str((DEFAULT_TERMS_XLSX).parent),
                                          filetypes=[("Terms files", "*.xlsx *.csv"), ("All", "*.*")])
        if path:
            self.var_terms.set(path)

    def _pick_pdfs(self):
        path = filedialog.askdirectory(title="Select PDFs folder", initialdir=str(DEFAULT_PDF_DIR))
        if path:
            self.var_pdfs.set(path)

    def _pick_scanned(self):
        path = filedialog.askdirectory(title="Select Scanned folder", initialdir=str(DEFAULT_SCANNED_DIR))
        if path:
            self.var_scanned.set(path)

    def _run(self):
        terms = Path(self.var_terms.get()).resolve()
        pdfs = Path(self.var_pdfs.get()).resolve()
        scanned = Path(self.var_scanned.get()).resolve()
        if not terms.exists():
            messagebox.showerror("Missing terms", f"Terms file not found:\n{terms}")
            return
        if not pdfs.exists():
            messagebox.showerror("Missing PDFs folder", f"PDFs folder not found:\n{pdfs}")
            return
        scanned.mkdir(parents=True, exist_ok=True)
        # Ensure latest env settings are saved before launching
        try:
            self._save_scanner_env(self._env_cache)
        except Exception:
            pass
        self.status.set("Running...")
        self._run_active = True
        self.btn_run.configure(state=tk.DISABLED)
        self.btn_stop.configure(state=tk.NORMAL)
        self.append_log(f"[GUI] Starting run with terms={terms}, pdfs={pdfs}\n")
        self.runner.start(terms, pdfs, scanned)

    def _stop(self):
        self.runner.stop()
        self._run_active = False
        self.status.set("Stopped (requested).")
        self.btn_run.configure(state=tk.NORMAL)
        self.btn_stop.configure(state=tk.DISABLED)

    def _open_last_run(self):
        try:
            if not RUNS_DIR.exists():
                messagebox.showinfo("No runs", f"No run_data found under:\n{RUNS_DIR}")
                return
            latest = max((p for p in RUNS_DIR.iterdir() if p.is_dir()), key=lambda p: p.stat().st_mtime, default=None)
            if not latest:
                messagebox.showinfo("No runs", "No run folders found")
                return
            self.status.set("Opening last run folder...")
            if sys.platform.startswith("win"):
                os.startfile(str(latest))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(latest)])
            else:
                subprocess.Popen(["xdg-open", str(latest)])
            self.status.set("Last run folder opened.")
        except Exception as e:
            messagebox.showwarning("Open failed", str(e))

    def _open_run_registry(self):
        try:
            reg_xlsx = ROOT / "Product_Data_File" / "run_registry.xlsx"
            reg_csv = ROOT / "Product_Data_File" / "run_registry.csv"
            target = None
            if reg_xlsx.exists():
                target = reg_xlsx
            elif reg_csv.exists():
                target = reg_csv
            else:
                messagebox.showinfo("No registry", "No run_registry.xlsx found yet. Run once to create it.")
                return
            self.status.set("Opening run registry...")
            if sys.platform.startswith("win"):
                os.startfile(str(target))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(target)])
            else:
                subprocess.Popen(["xdg-open", str(target)])
            self.status.set("Run registry opened.")
        except Exception as e:
            messagebox.showwarning("Open failed", str(e))

    def _generate_terms_spreadsheet(self):
        script = ROOT / "scripts" / "generate_terms_schema.py"
        if not script.exists():
            messagebox.showerror("Missing script", f"Could not find generator at:\n{script}")
            return
        dest_path = DEFAULT_TERMS_XLSX
        schema_path = dest_path.parent / "terms.schema.xlsx"
        self.status.set("Creating terms spreadsheet...")
        self.append_log("[GUI] Generating terms spreadsheet...\n")
        env = os.environ.copy()
        # Prefer project venv python if available
        py = _resolve_project_python()
        # Ensure vendored packages are visible as fallback
        env["PYTHONPATH"] = str(ROOT / "Lib" / "site-packages") + os.pathsep + env.get("PYTHONPATH", "")
        # Prepend venv Scripts/bin to PATH for helper tools
        try:
            scripts_dir = str(Path(py).parent)
            env["PATH"] = scripts_dir + os.pathsep + env.get("PATH", "")
        except Exception:
            pass
        try:
            proc = subprocess.Popen(
                [py, str(script)],
                cwd=str(ROOT),
                env=env,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
            )
        except Exception as exc:
            self.status.set("Ready to scan.")
            messagebox.showerror("Generation failed", str(exc))
            return

        def _pump():
            rc = 0
            try:
                stream = proc.stdout
                if stream is not None:
                    for line in stream:
                        self.append_log(line)
                rc = proc.wait()
            except Exception as exc:
                self.append_log(f"[WARN] Terms generator reader error: {exc}\n")
                rc = 1
            finally:
                if rc != 0:
                    self.status.set("Terms generation failed.")
                    messagebox.showerror("Generation failed", f"Generator exited with code {rc}")
                    return
                try:
                    if schema_path.exists():
                        try:
                            schema_path.replace(dest_path)
                        except Exception:
                            shutil.copyfile(schema_path, dest_path)
                            try:
                                schema_path.unlink()
                            except Exception:
                                pass
                    if dest_path.exists():
                        self.var_terms.set(str(dest_path))
                        self.status.set("Terms spreadsheet ready.")
                        self.append_log(f"[GUI] Terms spreadsheet available: {dest_path}\n")
                    else:
                        self.status.set("Ready to scan.")
                        messagebox.showwarning(
                            "Generation incomplete",
                            f"Generator completed but {dest_path} was not created.",
                        )
                except Exception as exc:
                    self.status.set("Ready to scan.")
                    messagebox.showerror("Post-processing failed", str(exc))

        threading.Thread(target=_pump, daemon=True).start()

    def _open_terms_spreadsheet(self):
        path = Path(self.var_terms.get()).expanduser()
        if not path.exists():
            messagebox.showinfo(
                "Terms spreadsheet missing",
                f"Terms spreadsheet not found at:\n{path}\n\nCreate it first.",
            )
            return
        try:
            self.status.set("Opening terms spreadsheet...")
            if sys.platform.startswith("win"):
                os.startfile(str(path))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
            self.status.set("Terms spreadsheet opened.")
        except Exception as e:
            messagebox.showwarning("Open failed", str(e))

    def _poll_runner(self):
        running = self.runner.proc is not None
        if running:
            self._run_active = True
        else:
            if self._run_active:
                self._run_active = False
                self.status.set("Scan finished.")
                self.btn_run.configure(state=tk.NORMAL)
                self.btn_stop.configure(state=tk.DISABLED)
        # Periodically refresh UI state (prereqs, statuses)
        try:
            self._refresh_ui_state()
        except Exception:
            pass
        self.after(750, self._poll_runner)

    # --- Settings (env knobs) ---
    def _open_settings(self):
        data = parse_scanner_env(SCANNER_ENV)
        if not data:
            data = self._env_cache.copy()
        dlg = tk.Toplevel(self)
        dlg.title("Settings")
        dlg.grab_set()
        frm = ttk.Frame(dlg, padding=8)
        frm.pack(fill=tk.BOTH, expand=True)

        # Row builder
        row = 0
        def add_row(label, widget):
            nonlocal row
            ttk.Label(frm, text=label+":").grid(row=row, column=0, sticky=tk.W, padx=(0,6), pady=4)
            widget.grid(row=row, column=1, sticky=tk.EW, pady=4)
            row += 1

        # Variables
        v_quiet = tk.IntVar(value=1 if data.get("QUIET","1").strip().lower() in ("1","true","yes") else 0)
        v_force_ocr = tk.IntVar(value=1 if data.get("FORCE_OCR","0").strip().lower() in ("1","true","yes","force","always") else 0)
        v_use_xy = tk.IntVar(value=1 if data.get("USE_EASYOCR_XY","0").strip().lower() in ("1","true","yes","on") else 0)
        v_ocr_mode = tk.StringVar(value=(data.get("OCR_MODE","fallback").strip().lower() or "fallback"))
        v_xy_log = tk.IntVar(value=1 if data.get("XY_LOG","0").strip().lower() in ("1","true","yes","on") else 0)
        v_ocr_dpi = tk.StringVar(value=data.get("OCR_DPI","600"))
        v_langs = tk.StringVar(value=data.get("EASYOCR_LANGS", data.get("OCR_LANGS", "en")))
        v_xy_fuzz = tk.StringVar(value=data.get("XY_FUZZ","0.75"))
        # bands removed: ROW_BAND, COL_TOL no longer used

        # Widgets
        add_row("Quiet logs", ttk.Checkbutton(frm, variable=v_quiet))
        add_row("Force OCR pre-extract", ttk.Checkbutton(frm, variable=v_force_ocr))
        add_row("Use EasyOCR XY", ttk.Checkbutton(frm, variable=v_use_xy))
        # OCR Mode selector
        cmb_mode = ttk.Combobox(frm, textvariable=v_ocr_mode, values=("fallback","ocr_only","no_ocr"), state="readonly", width=12)
        add_row("OCR mode", cmb_mode)
        add_row("XY debug log", ttk.Checkbutton(frm, variable=v_xy_log))
        ent_dpi = ttk.Entry(frm, textvariable=v_ocr_dpi, width=10)
        add_row("OCR DPI", ent_dpi)
        ent_langs = ttk.Entry(frm, textvariable=v_langs, width=20)
        add_row("EasyOCR langs (csv)", ent_langs)
        add_row("XY fuzz (0-1)", ttk.Entry(frm, textvariable=v_xy_fuzz, width=10))
        # Removed band/tolerance controls from UI
        frm.columnconfigure(1, weight=1)

        # Buttons
        btns = ttk.Frame(dlg)
        btns.pack(fill=tk.X, padx=8, pady=(0,8))
        def on_save():
            updates: dict[str,str|None] = {}
            updates["QUIET"] = "1" if v_quiet.get() else None
            updates["FORCE_OCR"] = "1" if v_force_ocr.get() else None
            updates["USE_EASYOCR_XY"] = "1" if v_use_xy.get() else None
            updates["XY_LOG"] = "1" if v_xy_log.get() else None
            updates["OCR_DPI"] = v_ocr_dpi.get().strip() or None
            # Prefer EASYOCR_LANGS key
            lang = v_langs.get().strip()
            updates["EASYOCR_LANGS"] = lang or None
            updates["XY_FUZZ"] = v_xy_fuzz.get().strip() or None
            updates["OCR_MODE"] = (v_ocr_mode.get().strip().lower() or None)
            # Clean obsolete keys if present in existing env
            updates["USE_OCRMYPDF"] = None
            updates["OCR_RENDERER"] = None
            try:
                # Merge with existing and persist
                merged = parse_scanner_env(SCANNER_ENV)
                for k, v in updates.items():
                    if v is None:
                        merged.pop(k, None)
                    else:
                        merged[k] = v
                self._env_cache = merged
                self._save_scanner_env(merged)
                self.status.set("Settings saved to scanner.env")
                dlg.destroy()
            except Exception as e:
                messagebox.showerror("Save failed", str(e))
        ttk.Button(btns, text="Save", command=on_save).pack(side=tk.RIGHT)
        ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side=tk.RIGHT, padx=(0,8))

    def _save_scanner_env(self, env_map: dict[str,str]) -> None:
        lines = [
            "# Scanner configuration (KEY=VALUE)",
            "# Edited via GUI Settings"
        ]
        # Order important keys first
        order = [
            "QUIET","OCR_MODE","OCR_DPI","EASYOCR_LANGS","FORCE_OCR","USE_EASYOCR_XY","XY_LOG","XY_FUZZ","VENV_DIR"
        ]
        written = set()
        for k in order:
            v = env_map.get(k)
            if v:
                lines.append(f"{k}={v}")
                written.add(k)
        # Write the rest, stable order
        for k in sorted(env_map.keys()):
            if k in written:
                continue
            v = env_map[k]
            if v:
                lines.append(f"{k}={v}")
        SCANNER_ENV.parent.mkdir(parents=True, exist_ok=True)
        SCANNER_ENV.write_text("\n".join(lines)+"\n", encoding="utf-8")

    def _compile_master(self):
        try:
            self.status.set("Compiling master...")
            # Launch compile script; stream output to log
            py = _resolve_project_python()
            cmd = [py, str(ROOT / "scripts" / "compile_master.py")]
            env = os.environ.copy()
            # Ensure vendored packages (openpyxl/xlsxwriter/pandas) are visible
            env["PYTHONPATH"] = str(ROOT / "Lib" / "site-packages") + os.pathsep + env.get("PYTHONPATH", "")
            try:
                scripts_dir = str(Path(py).parent)
                env["PATH"] = scripts_dir + os.pathsep + env.get("PATH", "")
            except Exception:
                pass
            proc = subprocess.Popen(cmd, cwd=str(ROOT), env=env, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
            self.append_log("[GUI] Compiling master...\n")
            def _pump():
                rc = 0
                try:
                    stream = proc.stdout
                    if stream is not None:
                        for line in stream:
                            self.append_log(line)
                    rc = proc.wait()
                except Exception as e:
                    self.append_log(f"[WARN] Master compile reader error: {e}\n")
                    try:
                        rc = proc.wait(timeout=1)
                    except Exception:
                        rc = 1
                finally:
                    if rc == 0:
                        self.status.set("Master compile finished.")
                    else:
                        self.status.set("Master compile failed.")
                        messagebox.showerror("Compile failed", f"compile_master.py exited with code {rc}")
            threading.Thread(target=_pump, daemon=True).start()
        except Exception as e:
            messagebox.showerror("Compile failed", str(e))

    def _open_master(self):
        try:
            xlsx = ROOT / "Product_Data_File" / "master.xlsx"
            csv = ROOT / "Product_Data_File" / "master.csv"
            target = xlsx if xlsx.exists() else (csv if csv.exists() else None)
            if not target:
                messagebox.showinfo("No master", "No master workbook found yet. Compile it first.")
                return
            self.status.set("Opening master workbook...")
            if sys.platform.startswith("win"):
                os.startfile(str(target))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(target)])
            else:
                subprocess.Popen(["xdg-open", str(target)])
            self.status.set("Master workbook opened.")
        except Exception as e:
            messagebox.showwarning("Open failed", str(e))

    # --- Setup / Install ---
    def _open_scanner_env(self):
        try:
            path = SCANNER_ENV
            path.parent.mkdir(parents=True, exist_ok=True)
            if not path.exists():
                path.write_text("# Scanner configuration (KEY=VALUE)\nQUIET=1\n", encoding="utf-8")
            if sys.platform.startswith("win"):
                os.startfile(str(path))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception as e:
            messagebox.showwarning("Open failed", str(e))

    def _install_full(self):
        # Run install.bat which sets up venv + vendored site-packages
        self.status.set("Installing (full)...")
        self.append_log("[GUI] Installing environment (full)...\n")
        if not sys.platform.startswith("win"):
            messagebox.showinfo("Windows only", "install.bat is for Windows. Use init venv on non-Windows.")
            return
        script = ROOT / "install.bat"
        if not script.exists():
            messagebox.showerror("Missing installer", f"Could not find install.bat at:\n{script}")
            return
        try:
            proc = subprocess.Popen([
                "cmd.exe", "/c", str(script)
            ], cwd=str(ROOT), stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
        except Exception as exc:
            self.status.set("Ready.")
            messagebox.showerror("Install failed", str(exc))
            return

        def _pump():
            rc = 0
            try:
                stream = proc.stdout
                if stream is not None:
                    for line in stream:
                        self.append_log(line)
                rc = proc.wait()
            except Exception as exc:
                self.append_log(f"[WARN] Installer reader error: {exc}\n")
                rc = 1
            finally:
                if rc == 0:
                    self.status.set("Install completed.")
                    # Refresh interpreter label
                    self.var_interp.set(_resolve_project_python())
                else:
                    self.status.set("Install failed.")
                    messagebox.showerror("Install failed", f"install.bat exited with code {rc}")
        threading.Thread(target=_pump, daemon=True).start()


    # --- Plotting helpers ---
    def _generate_plot_terms(self):
        script = ROOT / "scripts" / "generate_plot_terms.py"
        if not script.exists():
            messagebox.showerror("Missing script", f"Could not find generator at:\n{script}")
            return
        self.status.set("Creating plot terms...")
        self.append_log("[GUI] Generating plot terms...\n")
        env = os.environ.copy()
        py = _resolve_project_python()
        env["PYTHONPATH"] = str(ROOT / "Lib" / "site-packages") + os.pathsep + env.get("PYTHONPATH", "")
        try:
            scripts_dir = str(Path(py).parent)
            env["PATH"] = scripts_dir + os.pathsep + env.get("PATH", "")
        except Exception:
            pass
        try:
            proc = subprocess.Popen(
                [py, str(script)],
                cwd=str(ROOT),
                env=env,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
            )
        except Exception as exc:
            self.status.set("Ready to scan.")
            messagebox.showerror("Generation failed", str(exc))
            return

        def _pump():
            rc = 0
            try:
                stream = proc.stdout
                if stream is not None:
                    for line in stream:
                        self.append_log(line)
                rc = proc.wait()
            except Exception as exc:
                self.append_log(f"[WARN] Plot terms generator reader error: {exc}\n")
                rc = 1
            finally:
                if rc == 0:
                    dest_path = DEFAULT_PLOT_TERMS_XLSX if DEFAULT_PLOT_TERMS_XLSX.exists() else (DEFAULT_PLOT_TERMS_XLSX.with_suffix('.csv'))
                    if dest_path.exists():
                        self.var_plot_terms.set(str(dest_path))
                    self.status.set("Plot terms ready.")
                else:
                    self.status.set("Plot terms generation failed.")
                    messagebox.showerror("Generation failed", f"generate_plot_terms.py exited with code {rc}")
        threading.Thread(target=_pump, daemon=True).start()

    def _open_plot_terms(self):
        path = Path(self.var_plot_terms.get()).expanduser()
        if not path.exists():
            messagebox.showinfo(
                "Plot terms missing",
                f"Plot terms file not found at:\n{path}\n\nCreate it first.",
            )
            return
        try:
            self.status.set("Opening plot terms...")
            if sys.platform.startswith("win"):
                os.startfile(str(path))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
            self.status.set("Plot terms opened.")
        except Exception as e:
            messagebox.showwarning("Open failed", str(e))

    def _generate_plots(self):
        script = ROOT / "scripts" / "plot_from_master.py"
        if not script.exists():
            messagebox.showerror("Missing script", f"Could not find plotter at:\n{script}")
            return
        self.status.set("Generating plots...")
        self.append_log("[GUI] Generating plots...\n")
        env = os.environ.copy()
        py = _resolve_project_python()
        env["PYTHONPATH"] = str(ROOT / "Lib" / "site-packages") + os.pathsep + env.get("PYTHONPATH", "")
        try:
            scripts_dir = str(Path(py).parent)
            env["PATH"] = scripts_dir + os.pathsep + env.get("PATH", "")
        except Exception:
            pass
        try:
            proc = subprocess.Popen(
                [py, str(script)],
                cwd=str(ROOT),
                env=env,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
            )
        except Exception as exc:
            self.status.set("Ready to scan.")
            messagebox.showerror("Plotting failed", str(exc))
            return

        def _pump():
            rc = 0
            try:
                stream = proc.stdout
                if stream is not None:
                    for line in stream:
                        self.append_log(line)
                rc = proc.wait()
            except Exception as exc:
                self.append_log(f"[WARN] Plotter reader error: {exc}\n")
                rc = 1
            finally:
                if rc == 0:
                    self.status.set("Plots generated.")
                else:
                    self.status.set("Plotting failed.")
                    messagebox.showerror("Plotting failed", f"plot_from_master.py exited with code {rc}")
        threading.Thread(target=_pump, daemon=True).start()

    def _open_plots_folder(self):
        try:
            PLOTS_DIR.mkdir(parents=True, exist_ok=True)
            self.status.set("Opening plots folder...")
            if sys.platform.startswith("win"):
                os.startfile(str(PLOTS_DIR))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(PLOTS_DIR)])
            else:
                subprocess.Popen(["xdg-open", str(PLOTS_DIR)])
            self.status.set("Plots folder opened.")
        except Exception as e:
            messagebox.showwarning("Open failed", str(e))

    def _export_plot_summary(self):
        script = ROOT / "scripts" / "plots_to_excel_summary.py"
        if not script.exists():
            messagebox.showerror("Missing script", f"Could not find exporter at:\n{script}")
            return
        self.status.set("Creating plots summary...")
        self.append_log("[GUI] Creating plots summary...\n")
        env = os.environ.copy()
        py = _resolve_project_python()
        env["PYTHONPATH"] = str(ROOT / "Lib" / "site-packages") + os.pathsep + env.get("PYTHONPATH", "")
        try:
            scripts_dir = str(Path(py).parent)
            env["PATH"] = scripts_dir + os.pathsep + env.get("PATH", "")
        except Exception:
            pass
        try:
            proc = subprocess.Popen(
                [py, str(script)],
                cwd=str(ROOT),
                env=env,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
            )
        except Exception as exc:
            self.status.set("Ready to scan.")
            messagebox.showerror("Export failed", str(exc))
            return

        def _pump():
            rc = 0
            try:
                stream = proc.stdout
                if stream is not None:
                    for line in stream:
                        self.append_log(line)
                rc = proc.wait()
            except Exception as exc:
                self.append_log(f"[WARN] Plots summary reader error: {exc}\n")
                rc = 1
            finally:
                if rc == 0:
                    self.status.set("Plot summary created.")
                else:
                    self.status.set("Plot summary failed.")
                    messagebox.showerror("Export failed", f"plots_to_excel_summary.py exited with code {rc}")
        threading.Thread(target=_pump, daemon=True).start()

    def _open_plot_summary(self):
        try:
            target = ROOT / "Product_Data_File" / "plots_summary.xlsx"
            if not target.exists():
                messagebox.showinfo("No summary", "No plots_summary.xlsx found yet. Create it first.")
                return
            self.status.set("Opening plot summary...")
            if sys.platform.startswith("win"):
                os.startfile(str(target))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(target)])
            else:
                subprocess.Popen(["xdg-open", str(target)])
            self.status.set("Plot summary opened.")
        except Exception as e:
            messagebox.showwarning("Open failed", str(e))

    # --- Dynamic UI state ---
    def _refresh_ui_state(self):
        # Paths
        terms_path = Path(self.var_terms.get()).expanduser()
        pdfs_dir = Path(self.var_pdfs.get()).expanduser()
        scanned_dir = Path(self.var_scanned.get()).expanduser()
        master_xlsx = ROOT / "Product_Data_File" / "master.xlsx"
        master_csv = ROOT / "Product_Data_File" / "master.csv"
        registry_xlsx = ROOT / "Product_Data_File" / "run_registry.xlsx"
        registry_csv = ROOT / "Product_Data_File" / "run_registry.csv"
        plot_terms_xlsx = ROOT / "user_inputs" / "plot_terms.xlsx"
        plot_terms_csv = ROOT / "user_inputs" / "plot_terms.csv"
        plots_summary = ROOT / "Product_Data_File" / "plots_summary.xlsx"

        # Status helpers
        def exists(p: Path) -> bool:
            try:
                return p.exists()
            except Exception:
                return False

        def fmt_status(prefix: str, p: Path) -> tuple[str, str]:
            if exists(p):
                try:
                    ts = p.stat().st_mtime
                    import datetime as _dt
                    stamp = _dt.datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M")
                except Exception:
                    stamp = ""
                return f"{prefix}: OK {('('+stamp+')') if stamp else ''}", "#1b5e20"
            return f"{prefix}: Not found", "#a12d2d"

        # Environment status
        py = _resolve_project_python()
        env_ok = Path(py).exists()
        env_text = f"Env: {'OK' if env_ok else 'Not Installed'} ({py})"
        env_color = "#1b5e20" if env_ok else "#a12d2d"
        try:
            self.lbl_env_status.configure(text=env_text, foreground=env_color)
        except Exception:
            pass

        # Terms status
        try:
            t_text, t_col = fmt_status("Terms", terms_path)
            self.lbl_terms_status.configure(text=t_text, foreground=t_col)
        except Exception:
            pass

        # Master status
        master_path = master_xlsx if exists(master_xlsx) else master_csv
        try:
            m_text, m_col = fmt_status("Master", master_path)
            self.lbl_master_status.configure(text=m_text, foreground=m_col)
        except Exception:
            pass

        # Plot config status
        plot_cfg_path = plot_terms_xlsx if exists(plot_terms_xlsx) else plot_terms_csv
        try:
            p_text, p_col = fmt_status("Plot Config", plot_cfg_path)
            self.lbl_plot_cfg_status.configure(text=p_text, foreground=p_col)
        except Exception:
            pass

        # Enable/disable actions based on prereqs
        running = self.runner.proc is not None
        can_run = exists(pdfs_dir) and exists(terms_path) and not running
        self.btn_run.configure(state=(tk.NORMAL if can_run else tk.DISABLED))
        # Compile master requires some run data or registry
        has_registry = exists(registry_xlsx) or exists(registry_csv)
        has_run_data = any(d.is_dir() for d in (ROOT/"Product_Data_File"/"run_data").glob("*/")) if (ROOT/"Product_Data_File"/"run_data").exists() else False
        self.btn_compile_master.configure(state=(tk.NORMAL if (has_registry or has_run_data) else tk.DISABLED))
        # Open master enabled if present
        self.btn_open_master.configure(state=(tk.NORMAL if exists(master_path) else tk.DISABLED))
        # Open registry
        self.btn_open_registry.configure(state=(tk.NORMAL if has_registry else tk.DISABLED))
        # Plotting prereqs
        has_master = exists(master_xlsx) or exists(master_csv)
        has_plot_cfg = exists(plot_cfg_path)
        has_plots = any((PLOTS_DIR).glob("*.png")) if PLOTS_DIR.exists() else False
        self.btn_plot_terms_create.configure(state=(tk.NORMAL if has_master else tk.DISABLED))
        self.btn_plot_generate.configure(state=(tk.NORMAL if (has_master and has_plot_cfg) else tk.DISABLED))
        self.btn_plot_summary.configure(state=(tk.NORMAL if has_plots else tk.DISABLED))
        self.btn_open_plot_summary.configure(state=(tk.NORMAL if exists(plots_summary) else tk.DISABLED))
        # Extract tables prereqs
        ext_pdf_ok = exists(Path(self.var_ext_pdf.get()))
        self.btn_plots_open.configure(state=tk.NORMAL)
        try:
            self.btn_extract_tables.configure(state=(tk.NORMAL if ext_pdf_ok else tk.DISABLED))
        except Exception:
            pass


    # --- Page Table Extraction ---
    def _pick_ext_pdf(self):
        initial = str(DEFAULT_PDF_DIR)
        path = filedialog.askopenfilename(title="Select PDF", initialdir=initial, filetypes=[("PDF", "*.pdf"), ("All", "*.*")])
        if path:
            self.var_ext_pdf.set(path)

    def _extract_page_tables(self):
        script = ROOT / "scripts" / "extract_page_tables.py"
        if not script.exists():
            messagebox.showerror("Missing script", f"Could not find extractor at:\n{script}")
            return
        pdf = Path(self.var_ext_pdf.get()).expanduser()
        if not pdf.exists():
            messagebox.showerror("Missing PDF", f"Select a valid PDF under:\n{DEFAULT_PDF_DIR}")
            return
        pages = (self.var_ext_pages.get() or "").strip()
        self.status.set("Extracting page tables...")
        self.append_log(f"[GUI] Extracting tables from {pdf.name} pages='{pages}'...\n")
        env = os.environ.copy()
        py = _resolve_project_python()
        env["PYTHONPATH"] = str(ROOT / "Lib" / "site-packages") + os.pathsep + env.get("PYTHONPATH", "")
        cmd = [py, str(script), "--pdf", str(pdf)]
        if pages:
            cmd += ["--pages", pages]
        try:
            proc = subprocess.Popen(cmd, cwd=str(ROOT), env=env, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
        except Exception as exc:
            self.status.set("Ready.")
            messagebox.showerror("Extraction failed", str(exc))
            return

        def _pump():
            rc = 0
            try:
                stream = proc.stdout
                if stream is not None:
                    for line in stream:
                        self.append_log(line)
                rc = proc.wait()
            except Exception as exc:
                self.append_log(f"[WARN] Extractor reader error: {exc}\n")
                rc = 1
            finally:
                if rc == 0:
                    self.status.set("Page tables extracted.")
                else:
                    self.status.set("Table extraction failed.")
                    messagebox.showerror("Extraction failed", f"extract_page_tables.py exited with code {rc}")
        threading.Thread(target=_pump, daemon=True).start()


if __name__ == "__main__":
    App().mainloop()
