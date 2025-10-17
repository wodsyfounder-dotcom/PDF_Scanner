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
        self.title("EIDP Scanner GUI")
        self.geometry("900x600")
        self.runner = Runner(self.append_log)
        self._env_cache: dict[str, str] = parse_scanner_env(SCANNER_ENV)
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
        ttk.Button(btns, text="Settings", command=self._open_settings).pack(side=tk.LEFT, padx=(16, 0))
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
        # Ensure latest env settings are saved before launching
        try:
            self._save_scanner_env(self._env_cache)
        except Exception:
            pass
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
