"""
Master Cell State Management

This module manages the single source of truth for extracted cell values.
The state file (master_cell_state.json) is updated ONLY when the extractor runs.
It is never modified by workspace sync or other operations.

State Structure:
{
    "serial_component_id": {
        "term_name": {
            "value": "extracted_value",
            "last_updated": "2025-01-15T10:30:00",
            "run_folder": "20250115_103000",
            "ocr_dpi": 500,                 # optional, effective OCR DPI used
            "ocr_row_eps": 15.0,            # optional, effective line Y tolerance
            "smart_score": 2.03,            # optional, Smart-Snap match score
            "fuzzy_score": 0.99,            # optional, label fuzzy-match score
            "match_score": 2.03             # optional, primary term match score
        },
        ...
    },
    ...
}
"""

import json
import os
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, Optional, Set
import logging

import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

logger = logging.getLogger(__name__)

# Prefer new root-level artifacts, with legacy Product_Data_File as fallback for reads.
def _prefer_new(new_path: Path, *legacy_paths: Path) -> Path:
    if new_path.exists():
        return new_path
    for lp in legacy_paths:
        if lp.exists():
            return lp
    return new_path

# Path to the state file
STATE_FILE_PATH = Path("Product_Data_File/Master_Database/master_cell_state.json")
LEGACY_STATE_FILE_PATH = Path("master_cell_state.json")
LEGACY_STATE_FILE_PATH_2 = Path("Product_Data_File/master_cell_state.json")
MASTER_XLSX_PATH = Path("Product_Data_File/Master_Database/master.xlsx")
LEGACY_MASTER_XLSX_PATH = Path("master.xlsx")
LEGACY_MASTER_XLSX_PATH_2 = Path("Product_Data_File/master.xlsx")


def load_cell_state() -> Dict[str, Dict[str, Dict[str, Any]]]:
    """
    Load the master cell state from JSON file.

    Returns:
        Dictionary mapping serial_component -> term -> {value, last_updated, run_folder}
        Returns empty dict if file doesn't exist.
    """
    state_path = _prefer_new(STATE_FILE_PATH, LEGACY_STATE_FILE_PATH, LEGACY_STATE_FILE_PATH_2)
    if not state_path.exists():
        logger.info(f"Cell state file not found at {state_path}, returning empty state")
        return {}

    try:
        with open(state_path, 'r', encoding='utf-8') as f:
            state = json.load(f)
        logger.info(f"Loaded cell state with {len(state)} serial components")
        return state
    except json.JSONDecodeError as e:
        logger.error(f"Failed to parse cell state JSON: {e}")
        return {}
    except Exception as e:
        logger.error(f"Failed to load cell state: {e}")
        return {}


def save_cell_state(state: Dict[str, Dict[str, Dict[str, Any]]]) -> None:
    """
    Save the master cell state to JSON file atomically.

    Args:
        state: Dictionary mapping serial_component -> term -> {value, last_updated, run_folder}
    """
    state_path = STATE_FILE_PATH
    try:
        # Ensure parent directory exists
        state_path.parent.mkdir(parents=True, exist_ok=True)

        # Write to temp file first, then rename (atomic on most filesystems)
        temp_path = state_path.with_suffix('.tmp')
        with open(temp_path, 'w', encoding='utf-8') as f:
            json.dump(state, f, indent=2, ensure_ascii=False)

        # Atomic rename
        temp_path.replace(state_path)
        logger.info(f"Saved cell state with {len(state)} serial components")
        try:
            for legacy_state in (LEGACY_STATE_FILE_PATH, LEGACY_STATE_FILE_PATH_2):
                if legacy_state.exists() and legacy_state != state_path:
                    legacy_state.unlink()
        except Exception:
            pass

    except Exception as e:
        logger.error(f"Failed to save cell state: {e}")
        raise


def update_cell_state(
    serial_component: str,
    term_values: Dict[str, Any],
    run_folder: str,
    timestamp: Optional[str] = None,
    term_debug: Optional[Dict[str, Dict[str, Any]]] = None,
) -> None:
    """
    Update the cell state for a specific serial component with new extracted values.

    Args:
        serial_component: The serial component ID (e.g., "SN123_PG456")
        term_values: Dictionary mapping term names to extracted values
        run_folder: The run folder name (e.g., "20250115_103000")
        timestamp: ISO format timestamp, defaults to now
        term_debug: Optional mapping term -> debug fields
            (e.g., ocr_dpi, ocr_row_eps, smart_score, fuzzy_score, match_score)
    """
    if timestamp is None:
        timestamp = datetime.now().isoformat()

    # Load current state
    state = load_cell_state()

    # Ensure serial_component exists in state
    if serial_component not in state:
        state[serial_component] = {}

    # Update each term value (and optional debug metadata)
    for term_name, value in term_values.items():
        cell = state[serial_component].get(term_name, {})
        # Core fields used by all consumers
        cell["value"] = value
        cell["last_updated"] = timestamp
        cell["run_folder"] = run_folder
        # Optional debug metadata for JSON inspection
        if term_debug:
            debug_fields = term_debug.get(term_name) or {}
            for key, dbg_val in debug_fields.items():
                # Only persist simple JSON-serializable scalars
                if isinstance(dbg_val, (str, int, float, bool)) or dbg_val is None:
                    cell[key] = dbg_val
        state[serial_component][term_name] = cell

    # Save updated state
    save_cell_state(state)
    logger.info(f"Updated cell state for {serial_component} with {len(term_values)} terms")


def get_cell_value(serial_component: str, term: str) -> Optional[Dict[str, Any]]:
    """
    Get the state for a specific cell.

    Args:
        serial_component: The serial component ID
        term: The term name

    Returns:
        Dictionary with {value, last_updated, run_folder} or None if not found
    """
    state = load_cell_state()
    return state.get(serial_component, {}).get(term)


def get_referenced_run_folders() -> Set[str]:
    """
    Get all run folders that are referenced in the current cell state.
    Used for cache clearing - only these runs should be preserved.

    Returns:
        Set of run folder names (e.g., {"20250115_103000", "20250116_140000"})
    """
    state = load_cell_state()
    run_folders = set()

    for serial_component, terms in state.items():
        for term_name, cell_data in terms.items():
            run_folder = cell_data.get("run_folder")
            if run_folder:
                run_folders.add(run_folder)

    logger.info(f"Found {len(run_folders)} unique run folders referenced in cell state")
    return run_folders


def apply_state_to_master_incremental(serial_component: str, term_values: Dict[str, Any]) -> None:
    """
    Update ONLY the specified serial_component column in master.xlsx with the given term values.
    ALWAYS OVERRIDES existing values - newly scanned values replace old ones.

    Master.xlsx structure:
    - Row 1: Headers ["Term Label", "Data Group", "Units", "Min", "Max", "SN0000", "SN1111", ...]
    - Row 2-4: Metadata rows (Program, Space Vehicle, Data)
    - Row 5+: Term rows

    Args:
        serial_component: The serial component ID (column to update, e.g., "SN0000")
        term_values: Dictionary mapping term names to values
    """
    master_path = _prefer_new(MASTER_XLSX_PATH, LEGACY_MASTER_XLSX_PATH, LEGACY_MASTER_XLSX_PATH_2)
    if not master_path.exists():
        logger.warning(f"master.xlsx not found at {master_path}, skipping incremental update")
        return

    try:
        # Load workbook
        wb = openpyxl.load_workbook(master_path)
        ws = wb.active

        # Find header row (row 1) - get column index for this serial_component
        serial_col_idx = None
        for col_idx, cell in enumerate(ws[1], start=1):
            if cell.value and str(cell.value).strip() == serial_component:
                serial_col_idx = col_idx
                break

        if not serial_col_idx:
            logger.warning(f"Serial component '{serial_component}' not found in master.xlsx headers")
            return

        # Load schema to map term_name -> term_label
        from scripts.compile_master import _load_schema_terms
        schema_terms = _load_schema_terms()

        # Build mapping of term_label (lowercase) -> row_idx
        # Skip first 4 rows (header + 3 metadata rows), start at row 5
        term_label_to_row = {}
        for row_idx in range(5, ws.max_row + 1):
            term_label_cell = ws.cell(row=row_idx, column=1)  # Column 1 is "Term Label"
            if term_label_cell.value:
                term_label = str(term_label_cell.value).strip().lower()
                term_label_to_row[term_label] = row_idx

        # Update cells for each term
        updated_count = 0
        for term_name, value in term_values.items():
            # Get term_label from schema
            term_meta = schema_terms.get(term_name.lower(), {})
            term_label = term_meta.get("term_label", term_name)
            term_label_lower = term_label.lower()

            # Find the row for this term
            if term_label_lower in term_label_to_row:
                row_idx = term_label_to_row[term_label_lower]
                # ALWAYS OVERRIDE - set the new value
                ws.cell(row=row_idx, column=serial_col_idx, value=value)
                updated_count += 1
            else:
                logger.debug(f"Term '{term_label}' not found in master.xlsx, skipping")

        # Save workbook
        if updated_count > 0:
            MASTER_XLSX_PATH.parent.mkdir(parents=True, exist_ok=True)
            wb.save(MASTER_XLSX_PATH)
            try:
                if master_path != MASTER_XLSX_PATH and master_path.exists():
                    master_path.unlink()
            except Exception:
                pass
            logger.info(f"Updated {updated_count} cells in master.xlsx for {serial_component} (overriding existing values)")
        else:
            logger.debug(f"No cells updated for {serial_component}")

    except Exception as e:
        logger.error(f"Failed to apply incremental update to master.xlsx: {e}")
        raise


def build_dataframe_from_state() -> pd.DataFrame:
    """
    Build a pandas DataFrame from the cell state.
    Used by the "Compile New Master Workbook" function to rebuild from scratch.

    Returns:
        DataFrame with columns: serial_component, program_name, vehicle_number, term1, term2, ...
    """
    state = load_cell_state()

    if not state:
        logger.warning("Cell state is empty, returning empty DataFrame")
        return pd.DataFrame()

    # Collect all unique terms across all serial components
    all_terms = set()
    for serial_component, terms in state.items():
        all_terms.update(terms.keys())

    all_terms = sorted(all_terms)

    # Build rows
    rows = []
    for serial_component, terms in state.items():
        row = {"serial_component": serial_component}

        # Extract metadata if present (program_name, vehicle_number might be in terms or separate)
        # For now, just extract all term values
        for term_name in all_terms:
            cell_data = terms.get(term_name)
            row[term_name] = cell_data["value"] if cell_data else None

        rows.append(row)

    # Create DataFrame
    columns = ["serial_component"] + all_terms
    df = pd.DataFrame(rows, columns=columns)

    logger.info(f"Built DataFrame from state with {len(df)} rows and {len(df.columns)} columns")
    return df


def sync_cell_state_with_master() -> None:
    """
    Prune cell_state.json to only contain terms and serial components that exist in master.xlsx.

    Master.xlsx is the source of truth for which terms (rows) and serial components (columns) should exist.
    - If a user deletes a row from master.xlsx, this removes that term from all serial components
    - If a user deletes a column from master.xlsx, this removes that serial component entirely

    This ensures deleted terms and serial components don't reappear when compiling from state.
    """
    # Load master.xlsx to determine which terms and serial components should exist
    master_path = _prefer_new(MASTER_XLSX_PATH, LEGACY_MASTER_XLSX_PATH, LEGACY_MASTER_XLSX_PATH_2)
    if not master_path.exists():
        logger.warning(f"master.xlsx not found at {master_path}, cannot sync cell state")
        return

    try:
        # Read master.xlsx to get list of term_labels and serial components
        df = pd.read_excel(master_path, dtype=object, keep_default_na=False)

        # Skip first 3 metadata rows (Program, Space Vehicle, Data)
        # Term rows start at row 4 (index 3)
        if len(df) < 4:
            logger.warning("master.xlsx has fewer than 4 rows, cannot determine terms")
            return

        term_rows = df.iloc[3:]  # Skip metadata rows

        # Extract term_labels from the "Term Label" column
        if "Term Label" not in df.columns:
            logger.error("master.xlsx does not have 'Term Label' column")
            return

        allowed_term_labels = set()
        for term_label in term_rows["Term Label"]:
            if term_label and str(term_label).strip():
                allowed_term_labels.add(str(term_label).strip().lower())

        logger.info(f"Found {len(allowed_term_labels)} terms in master.xlsx")

        # Extract serial components from columns (skip first 5: Term Label, Data Group, Units, Min, Max)
        metadata_columns = {"Term Label", "Data Group", "Units", "Min", "Max"}
        allowed_serial_components = set()
        for col in df.columns:
            col_str = str(col).strip()
            if col_str and col_str not in metadata_columns:
                allowed_serial_components.add(col_str)

        logger.info(f"Found {len(allowed_serial_components)} serial components in master.xlsx")

        # Load cell state
        state = load_cell_state()
        if not state:
            logger.info("Cell state is empty, nothing to sync")
            return

        # Load schema to map term_name -> term_label
        from scripts.compile_master import _load_schema_terms
        schema_terms = _load_schema_terms()

        # Prune serial components that don't exist in master.xlsx
        serial_components_to_remove = []
        for serial_component in state.keys():
            if serial_component not in allowed_serial_components:
                serial_components_to_remove.append(serial_component)

        for serial_component in serial_components_to_remove:
            del state[serial_component]

        # Prune terms from each remaining serial component
        total_terms_pruned = 0
        for serial_component, terms in state.items():
            terms_to_remove = []

            for term_name in terms.keys():
                # Get term_label from schema
                term_meta = schema_terms.get(term_name.lower(), {})
                term_label = term_meta.get("term_label", term_name)

                # Check if this term exists in master.xlsx
                if term_label.lower() not in allowed_term_labels:
                    terms_to_remove.append(term_name)

            # Remove orphaned terms
            for term_name in terms_to_remove:
                del terms[term_name]
                total_terms_pruned += 1

        # Save pruned state
        save_cell_state(state)

        # Log results
        total_removed = len(serial_components_to_remove)
        if total_removed > 0 or total_terms_pruned > 0:
            logger.info(f"Synced cell state with master.xlsx: removed {total_removed} serial components, {total_terms_pruned} orphaned term entries")
            print(f"[INFO] Synced cell state: removed {total_removed} serial components, {total_terms_pruned} orphaned term entries")
        else:
            logger.info("Cell state already in sync with master.xlsx")
            print("[INFO] Cell state already in sync with master.xlsx")

    except Exception as e:
        logger.error(f"Failed to sync cell state with master: {e}")
        raise
