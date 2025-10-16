#!/usr/bin/env python3
"""
Simple native GUI (Tkinter) launcher for the EIDP Term Scanner.

- Pure Python standard library (Tkinter) — no external EXEs or downloads.
- Loads env from user_inputs/scanner.env
- Lets you pick Terms file, PDFs folder, Scanned folder
- Runs the scanner in a background thread, streams stdout/stderr to the UI
- Buttons to Stop run and Open last run folder

Usage:
  python gui.py
"""
from __future__ import annotations

import os
import sys
import threading
import subprocess
import queue
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


ROOT = Path(__file__).resolve().parent
DEFAULT_TERMS_XLSX = ROOT / "user_inputs" / "terms.xlsx"
DEFAULT_PDF_DIR = ROOT / "user_inputs" / "EIDP_Import_Docs"
DEFAULT_SCANNED_DIR = ROOT / "user_inputs" / "Scanned_Docs"
SCANNER_ENV = ROOT / "user_inputs" / "scanner.env"
APP_ENTRY = ROOT / "Application" / "eidp_term_scanner.py"
RUNS_DIR = ROOT / "Product_Data_File" / "run_data"


def parse_scanner_env(path: Path) -> dict[str, str]:
    env = {}
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
        if not k:
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
            ]
        else:
            # Fallback: invoke the Python entry directly on non-Windows
            cmd = [
                sys.executable,
                str(APP_ENTRY),
                "--input", str(terms),
                "--pdf-folder", str(pdf_dir),
                "--scanned-folder", str(scanned),
            ]
        env = os.environ.copy()
        # Let run.bat parse scanner.env; still merge here for direct Python fallback
        env.update(parse_scanner_env(SCANNER_ENV))
        # Ensure vendored packages are visible for direct Python fallback
        env["PYTHONPATH"] = str(ROOT / "Lib" / "site-packages") + os.pathsep + env.get("PYTHONPATH", "")

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
        self.title("EIDP Scanner GUI")
        self.geometry("900x600")
        self.runner = Runner(self.append_log)
        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 8, "pady": 6}

        frm = ttk.Frame(self)
        frm.pack(fill=tk.X, **pad)

        # Terms file
        ttk.Label(frm, text="Terms file (.xlsx/.csv):").grid(row=0, column=0, sticky=tk.W)
        self.var_terms = tk.StringVar(value=str(DEFAULT_TERMS_XLSX))
        ent_terms = ttk.Entry(frm, textvariable=self.var_terms, width=80)
        ent_terms.grid(row=0, column=1, sticky=tk.EW)
        ttk.Button(frm, text="Browse", command=self._pick_terms).grid(row=0, column=2)

        # PDFs folder
        ttk.Label(frm, text="PDFs folder:").grid(row=1, column=0, sticky=tk.W)
        self.var_pdfs = tk.StringVar(value=str(DEFAULT_PDF_DIR))
        ent_pdfs = ttk.Entry(frm, textvariable=self.var_pdfs, width=80)
        ent_pdfs.grid(row=1, column=1, sticky=tk.EW)
        ttk.Button(frm, text="Browse", command=self._pick_pdfs).grid(row=1, column=2)

        # Scanned folder
        ttk.Label(frm, text="Scanned folder:").grid(row=2, column=0, sticky=tk.W)
        self.var_scanned = tk.StringVar(value=str(DEFAULT_SCANNED_DIR))
        ent_sc = ttk.Entry(frm, textvariable=self.var_scanned, width=80)
        ent_sc.grid(row=2, column=1, sticky=tk.EW)
        ttk.Button(frm, text="Browse", command=self._pick_scanned).grid(row=2, column=2)

        frm.columnconfigure(1, weight=1)

        # Controls
        btns = ttk.Frame(self)
        btns.pack(fill=tk.X, **pad)
        self.btn_run = ttk.Button(btns, text="Run", command=self._run)
        self.btn_run.pack(side=tk.LEFT)
        self.btn_stop = ttk.Button(btns, text="Stop", command=self._stop, state=tk.DISABLED)
        self.btn_stop.pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(btns, text="Open Last Run Folder", command=self._open_last_run).pack(side=tk.LEFT, padx=(16, 0))
        ttk.Button(btns, text="Open Run Registry", command=self._open_registry).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(btns, text="Compile Master", command=self._compile_master).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(btns, text="Open Master", command=self._open_master).pack(side=tk.LEFT, padx=(8, 0))

        # Log area
        self.txt = tk.Text(self, wrap="word", height=24)
        self.txt.pack(fill=tk.BOTH, expand=True, **pad)
        self.txt.configure(state=tk.DISABLED)

        # Status bar
        self.status = tk.StringVar(value="Ready.")
        ttk.Label(self, textvariable=self.status, anchor=tk.W).pack(fill=tk.X)

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
        self.status.set("Running...")
        self.btn_run.configure(state=tk.DISABLED)
        self.btn_stop.configure(state=tk.NORMAL)
        self.append_log(f"[GUI] Starting run with terms={terms}, pdfs={pdfs}\n")
        self.runner.start(terms, pdfs, scanned)

    def _stop(self):
        self.runner.stop()
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
            if sys.platform.startswith("win"):
                os.startfile(str(latest))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(latest)])
            else:
                subprocess.Popen(["xdg-open", str(latest)])
        except Exception as e:
            messagebox.showwarning("Open failed", str(e))

    def _open_registry(self):
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
            if sys.platform.startswith("win"):
                os.startfile(str(target))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(target)])
            else:
                subprocess.Popen(["xdg-open", str(target)])
        except Exception as e:
            messagebox.showwarning("Open failed", str(e))

    def _compile_master(self):
        try:
            self.status.set("Compiling master...")
            # Launch compile script; stream output to log
            cmd = [sys.executable, str(ROOT / "scripts" / "compile_master.py")]
            env = os.environ.copy()
            # Ensure vendored packages (openpyxl/xlsxwriter/pandas) are visible
            env["PYTHONPATH"] = str(ROOT / "Lib" / "site-packages") + os.pathsep + env.get("PYTHONPATH", "")
            proc = subprocess.Popen(cmd, cwd=str(ROOT), env=env, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
            self.append_log("[GUI] Compiling master...\n")
            def _pump():
                try:
                    for line in proc.stdout:  # type: ignore[arg-type]
                        self.append_log(line)
                except Exception as e:
                    self.append_log(f"[WARN] Master compile reader error: {e}\n")
                finally:
                    try:
                        proc.wait(timeout=1)
                    except Exception:
                        pass
                    self.status.set("Master compile finished.")
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
            if sys.platform.startswith("win"):
                os.startfile(str(target))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(target)])
            else:
                subprocess.Popen(["xdg-open", str(target)])
        except Exception as e:
            messagebox.showwarning("Open failed", str(e))


if __name__ == "__main__":
    App().mainloop()
