#!/usr/bin/env python3
"""
One-time migration script to create initial master_cell_state.json from existing data.

This script:
1. Reads the current master.xlsx (if exists)
2. Reads all scan_results.json files from run_registry
3. For each serial_component and term, finds the most recent scan result
4. Populates master_cell_state.json with this data
5. Validates that the state roughly matches current master.xlsx

Run this ONCE before deploying the new cell state system.
"""

from pathlib import Path
from typing import Dict, List, Any, Tuple, Optional
import json
import sys
from datetime import datetime

# Add parent directory to path for imports
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from scripts.master_cell_state import save_cell_state, load_cell_state
from scripts.compile_master import (
    load_registry,
    load_results_json,
    EXPORTS,
    OUT_XLSX,
    _read_existing_master,
    norm
)


def migrate_to_cell_state(dry_run: bool = False) -> None:
    """
    Create initial master_cell_state.json from existing run data.

    Args:
        dry_run: If True, only validate and don't write the state file
    """
    print("[INFO] Starting migration to cell state system...")

    # Load registry to get all runs
    print("[INFO] Loading run registry...")
    registry = load_registry()

    if not registry:
        print("[WARN] No registry entries found. Nothing to migrate.")
        return

    # Build a map of (serial_component, term) -> list of (run_folder, timestamp, value)
    # We'll keep the most recent value for each cell
    cell_history: Dict[Tuple[str, str], List[Tuple[str, datetime, Any]]] = {}

    print(f"[INFO] Processing {len(registry)} registry entries...")

    for serial_component, run_folder, metadata in registry:
        # Load scan results for this run
        rows = load_results_json(run_folder)

        if not rows:
            continue

        # Parse timestamp from run folder name
        try:
            timestamp = datetime.strptime(run_folder.name, "%Y%m%d_%H%M%S")
        except Exception:
            # Fallback to file modification time
            try:
                timestamp = datetime.fromtimestamp(run_folder.stat().st_mtime)
            except Exception:
                timestamp = datetime.now()

        # Process each row in scan results
        for row in rows:
            sn = norm(row.get("serial_component") or row.get("serial_number"))
            if not sn or sn != serial_component:
                # Skip rows that don't match this serial component
                # (there might be multiple PDFs in one run)
                continue

            # Get term label
            term_label = norm(row.get("term_label") or row.get("term"))
            if not term_label:
                continue

            # Extract value from row
            value = extract_value_from_row(row)

            # Record this historical value
            key = (serial_component, term_label)
            if key not in cell_history:
                cell_history[key] = []
            cell_history[key].append((run_folder.name, timestamp, value))

    # Now build the cell state by taking the most recent value for each cell
    print(f"[INFO] Building cell state from {len(cell_history)} unique cells...")

    state: Dict[str, Dict[str, Dict[str, Any]]] = {}

    for (serial_component, term_label), history in cell_history.items():
        # Sort by timestamp (most recent first)
        history.sort(key=lambda x: x[1], reverse=True)

        # Take the most recent
        run_folder_name, timestamp, value = history[0]

        # Add to state
        if serial_component not in state:
            state[serial_component] = {}

        state[serial_component][term_label] = {
            "value": value,
            "last_updated": timestamp.isoformat(),
            "run_folder": run_folder_name
        }

    # Validate against current master.xlsx (if exists)
    if OUT_XLSX.exists():
        print("[INFO] Validating against current master.xlsx...")
        validate_against_master(state)

    # Save the state
    if dry_run:
        print("[DRY RUN] Would save state with:")
        print(f"  - {len(state)} serial components")
        total_cells = sum(len(terms) for terms in state.values())
        print(f"  - {total_cells} total cells")
        print("[DRY RUN] Not writing to file.")
    else:
        print("[INFO] Saving cell state...")
        save_cell_state(state)
        print(f"[DONE] Migration complete!")
        print(f"  - {len(state)} serial components")
        total_cells = sum(len(terms) for terms in state.values())
        print(f"  - {total_cells} total cells")
        print(f"  - Saved to: Product_Data_File/master_cell_state.json")


def extract_value_from_row(row: Dict[str, Any]) -> Optional[str]:
    """
    Extract the cell value from a scan result row.
    Mimics the logic from compile_master.py
    """
    # Prefer any concrete extracted value
    for key in ("extracted_value", "number", "text", "string", "value"):
        if key in row:
            val = row.get(key)
            if val is None:
                continue
            text = norm(val)
            if text != "":
                return text

    # No concrete value: if the extractor ran but found nothing, mark N/A
    try:
        found_flag = row.get("found", None)
    except Exception:
        found_flag = None

    if found_flag is False:
        return "N/A"

    err = norm(row.get("error_reason"))
    if err:
        return "N/A"

    return None


def validate_against_master(state: Dict[str, Dict[str, Dict[str, Any]]]) -> None:
    """
    Validate that the cell state roughly matches current master.xlsx.
    Warns about significant discrepancies but doesn't fail.
    """
    try:
        header, rows, mtime = _read_existing_master()
    except Exception as e:
        print(f"[WARN] Could not read existing master for validation: {e}")
        return

    if not rows:
        print("[INFO] No existing master data to validate against")
        return

    # Base columns that aren't serial components
    base_columns = ["Term Label", "Data Group", "Units", "Min", "Max"]
    serial_components = [col for col in header if col not in base_columns]

    # Skip metadata rows (Program, Space Vehicle, Data)
    term_rows = [
        row for row in rows
        if norm(row.get("Term Label", "")).lower() not in ("program", "space vehicle", "data")
    ]

    # Compare values
    mismatches = 0
    matches = 0
    state_only = 0
    master_only = 0

    for row in term_rows:
        term_label = norm(row.get("Term Label"))
        if not term_label:
            continue

        for serial_component in serial_components:
            master_value = norm(row.get(serial_component, ""))
            state_value = None

            # Get value from state
            if serial_component in state and term_label in state[serial_component]:
                state_value = norm(str(state[serial_component][term_label].get("value", "")))

            # Compare
            if master_value and state_value:
                if master_value == state_value:
                    matches += 1
                else:
                    mismatches += 1
                    if mismatches <= 10:  # Only show first 10 mismatches
                        print(f"[MISMATCH] {serial_component} / {term_label}: master='{master_value}' vs state='{state_value}'")
            elif master_value and not state_value:
                master_only += 1
            elif state_value and not master_value:
                state_only += 1

    print(f"[VALIDATION] Results:")
    print(f"  - Matches: {matches}")
    print(f"  - Mismatches: {mismatches}")
    print(f"  - Only in master.xlsx: {master_only}")
    print(f"  - Only in state: {state_only}")

    if mismatches > 0:
        print(f"[WARN] Found {mismatches} mismatches - this may be expected if master.xlsx has manual edits")


def main():
    import argparse
    parser = argparse.ArgumentParser(
        description="Migrate existing data to master_cell_state.json"
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Validate and preview without writing the state file"
    )
    args = parser.parse_args()

    try:
        migrate_to_cell_state(dry_run=args.dry_run)
    except Exception as e:
        print(f"[ERROR] Migration failed: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
