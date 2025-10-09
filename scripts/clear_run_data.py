#!/usr/bin/env python3
"""
Clear per-run artifacts under Product_Data_File/run_data.

Usage:
  py scripts/clear_run_data.py            # prompts for confirmation
  py scripts/clear_run_data.py --yes      # no prompt
"""
from __future__ import annotations
from pathlib import Path
import shutil
import argparse
import sys


def rm_tree(path: Path) -> None:
    if not path.exists():
        return
    for child in path.iterdir():
        try:
            if child.is_dir():
                shutil.rmtree(child, ignore_errors=True)
            else:
                child.unlink(missing_ok=True)  # type: ignore[arg-type]
        except Exception:
            pass


def main() -> None:
    parser = argparse.ArgumentParser(description="Clear Product_Data_File/run_data contents")
    parser.add_argument("--yes", action="store_true", help="do not prompt for confirmation")
    args = parser.parse_args()

    root = Path(__file__).resolve().parents[1]
    run_data = root / "Product_Data_File" / "run_data"
    if not run_data.exists():
        print(f"[INFO] Nothing to clear; folder not found: {run_data}")
        return

    if not args.yes:
        print(f"This will remove all subfolders/files under:\n  {run_data}")
        try:
            resp = input("Proceed? [y/N]: ").strip().lower()
        except Exception:
            print("[ABORT] No input available.")
            sys.exit(1)
        if resp not in ("y", "yes"):
            print("[ABORT] No changes made.")
            return

    rm_tree(run_data)
    print(f"[DONE] Cleared: {run_data}")


if __name__ == "__main__":
    main()

