#!/usr/bin/env python3
"""
Initialize a local Python virtual environment for this project.

Usage examples:
  python scripts/init_venv.py                    # .venv, minimal deps
  python scripts/init_venv.py --dir .myvenv      # custom venv path
  python scripts/init_venv.py --full             # install full feature set (OCR)
  python scripts/init_venv.py --write-env        # write VENV_DIR to user_inputs/scanner.env

Modes:
  - minimal (default): pandas, openpyxl, xlsxwriter, matplotlib, pymupdf
  - full: minimal + opencv-python-headless, easyocr, torch+torchvision (CPU)
"""
from __future__ import annotations

import argparse
import os
import platform
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_VENV = ROOT / ".venv"
SCANNER_ENV = ROOT / "user_inputs" / "scanner.env"


def venv_python(venv_dir: Path) -> Path:
    if platform.system().lower().startswith("win"):
        return venv_dir / "Scripts" / "python.exe"
    return venv_dir / "bin" / "python"


def run(cmd: list[str], cwd: Path | None = None) -> int:
    print("[INIT]", " ".join(cmd))
    try:
        return subprocess.call(cmd, cwd=str(cwd) if cwd else None)
    except Exception as e:
        print(f"[ERROR] Command failed: {e}")
        return 1


def ensure_env_file_has_venv(venv_dir: Path) -> None:
    SCANNER_ENV.parent.mkdir(parents=True, exist_ok=True)
    lines: list[str] = []
    if SCANNER_ENV.exists():
        try:
            lines = SCANNER_ENV.read_text(encoding="utf-8", errors="ignore").splitlines()
        except Exception:
            lines = []
    # Remove existing VENV_DIR lines
    filtered = [ln for ln in lines if not ln.strip().upper().startswith("VENV_DIR=")]
    filtered.append(f"VENV_DIR={venv_dir}")
    SCANNER_ENV.write_text("\n".join(filtered) + "\n", encoding="utf-8")
    print(f"[DONE] Wrote VENV_DIR to {SCANNER_ENV}")


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--dir", dest="venv_dir", default=str(DEFAULT_VENV), help="Path to create/use venv (default: .venv)")
    p.add_argument("--full", action="store_true", help="Install full feature set (OCR)")
    p.add_argument("--write-env", action="store_true", help="Write VENV_DIR to user_inputs/scanner.env")
    args = p.parse_args()

    vdir = Path(args.venv_dir).resolve()
    # Create venv
    if not vdir.exists():
        rc = run([sys.executable, "-m", "venv", str(vdir)])
        if rc:
            print("[ERROR] Failed to create venv. Ensure Python 3 is installed and accessible.")
            return rc
    py = venv_python(vdir)
    if not py.exists():
        print(f"[ERROR] venv python not found at: {py}")
        return 1

    # Upgrade pip
    run([str(py), "-m", "pip", "install", "--upgrade", "pip"])  # ignore errors

    # Base deps for scanning + master + plotting
    base = [
        "pymupdf",
        "pandas",
        "openpyxl",
        "xlsxwriter",
        "matplotlib",
    ]
    rc = run([str(py), "-m", "pip", "install", *base])
    if rc:
        print("[WARN] Base package install reported errors; continuing.")

    if args.full:
        # OCR + helpers (CPU-only torch)
        rc_torch = run([str(py), "-m", "pip", "install", "--index-url", "https://download.pytorch.org/whl/cpu", "torch", "torchvision"])
        if rc_torch:
            print("[WARN] Torch CPU install failed/swallowed; EasyOCR may still work without it.")
        run([str(py), "-m", "pip", "install", "opencv-python-headless"])  # ignore rc
        run([str(py), "-m", "pip", "install", "easyocr"])  # ignore rc

    print(f"[READY] Virtual environment ready at: {vdir}")
    print(f"        Python: {py}")
    print("        Tip: set VENV_DIR in user_inputs/scanner.env or use --write-env.")
    if args.write_env:
        ensure_env_file_has_venv(vdir)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())

