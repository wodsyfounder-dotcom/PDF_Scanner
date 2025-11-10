# EIDP Term Scanner

Scan a folder of EIDP PDFs for a list of terms (from CSV/XLSX), find the closest numeric value near each term within specified page ranges, and export a matrix (rows=terms, columns=Serial Numbers) plus detailed metadata. PDFs remain in their source location; reruns simply re-read the same repository.

---

## Install (Windows)

- No admin / no PATH changes (recommended):
  - Run `install.bat` to create a local `.venv` and install Python packages there.
  - All OCR is handled inside Python (EasyOCR). No external executables are required.
  - Optional custom venv location: `install.bat C:\\Path\\To\\PDF_Scanner.venv` or set `VENV_DIR` in `user_inputs\\scanner.env`.
  - Optional custom venv location: `install.bat C:\\Path\\To\\PDF_Scanner.venv` or set `VENV_DIR` in `user_inputs\\scanner.env`.
- System-wide (if allowed):
  - `py -m pip install --upgrade pip`
  - `py -m pip install pymupdf pdfminer.six pypdf` (minimum)
  - Optional (Excel): `py -m pip install openpyxl` (or `pandas xlsxwriter`)
  - Optional (OCR fallback): `py -m pip install --index-url https://download.pytorch.org/whl/cpu torch torchvision` then `py -m pip install easyocr`

Notes:
- If your PDFs are text-based (not scans), OCR may never be invoked. At least one text extractor (PyMuPDF or pdfminer.six) is recommended.
- Excel output requires `openpyxl` or `pandas` with `xlsxwriter`. If missing, CSV fallbacks are written instead.
- Reading `.xls` terms requires `pandas` and `xlrd`. If unavailable, save as `.xlsx` or `.csv`.

---

## Super Simple Run

- One-time setup: `install.bat` scaffolds `user_inputs` and drops a sample terms file if none exists (`user_inputs\terms.xlsx` or `.csv` with headers `Term, Pages`).
- Point the GUI/CLI at the folder that already contains your EIDP PDFs (default `Data Packages`, but any path works).
- Edit `user_inputs\terms.xlsx` or `user_inputs\terms.csv`.
- Optional: set runtime knobs in `user_inputs\scanner.env` (KEY=VALUE lines).
- Run one of:
  - Windows: double-click `run.bat`
  - Python: `py .\Application\eidp_term_scanner.py --input .\user_inputs\terms.xlsx --pdf-folder ".\Data Packages" --window-chars 400`

Minimal transfer (work PC): copy `Application\eidp_term_scanner.py`, `Application\eidp_term_scanner.core.py`, `run.bat`, `install.bat`, and `user_inputs` (or at least `user_inputs\scanner.env` and your `terms.xlsx`). Point the scanner at your existing PDF repository (e.g., `Data Packages`). Set `VENV_DIR` in `scanner.env` if you want the venv outside the repo.
Outputs are saved under `Product_Data_File\run_data\<timestamp>`.
- A persistent run registry is maintained at `Product_Data_File\run_registry.xlsx` (CSV fallback if Excel not available).
- To build a consolidated workbook across runs, use the "Compile Master" action (GUI) or run `scripts/compile_master.py` to write `Product_Data_File\master.xlsx` (CSV fallback).

---

## What You Get

- `Product_Data_File\run_data\<timestamp>\scan_results.xlsx` (if Excel writers available)
  - `results` sheet: Term, Pages, and one column per SN (Serial Number from filename)
  - `metadata` sheet: pdf_file, serial_component, term, found, page, number, units, context, method_pipeline
- CSV fallbacks (written into the same `run_data` folder if Excel libs are missing):
  - `scan_results.results.csv` and `scan_results.metadata.csv`
- `scan_results_flat.csv` and `scan_results.json` (audit) also live in the per-run folder
  - `by_pdf/` folder contains one JSON per PDF with just that fileâ€™s term results
- Persistent cross-run registry: `Product_Data_File\run_registry.xlsx` (or `run_registry.csv`).
- Optional consolidated workbook via compile: `Product_Data_File\master.xlsx` (or `master.csv`).

---

## GUI Launcher

- Start with a simple GUI: `py gui.py`
  - Choose your Terms file and PDFs folder.
  - The GUI streams scanner logs and provides shortcuts to open the last run folder, run registry, and compile/open the master workbook.
  - The GUI runs in quiet mode by default to reduce log noise.

---

## Quiet Output

- Reduce console chatter via either:
  - CLI: add `--quiet` to the scanner command
  - Env var: set `QUIET=1` in `user_inputs\scanner.env`
- Quiet mode suppresses progress/debug lines (e.g., `[PROGRESS]`, `[XY]`, setup chatter) and keeps `[ERROR]`, `[WARN]`, and `[DONE]` summaries.

---

## Packaging (Single-File Installer)

- Option A: Zip-based installer (smaller payload)
- Create a single `install.bat` using: `powershell -NoProfile -ExecutionPolicy Bypass -File scripts/make_single_installer.ps1`
  - Output: `dist\install.bat`
  - On a clean machine: copy `install.bat` to an empty folder and run it.
  - It extracts the app files and automatically runs project setup (creates a venv, installs deps, scaffolds folders, and generates `scanner.env` and `terms.schema.xlsx`).
  - Optional: pass a custom venv location: `install.bat C:\\MyVenvs\\PDF_Scanner.venv` (equivalent to calling the internal setup with that argument).

- Option B: Full-text installer (no embedded zip)
  - Generate with: `powershell -NoProfile -ExecutionPolicy Bypass -File scripts/make_full_text_installer.ps1`
  - Output: `dist\full_installer.bat`
  - This single BAT embeds file contents as Base64 and reconstructs the project via PowerShell only (no certutil/zip required).
  - Usage on a clean machine: run `full_installer.bat` (optionally pass venv path like above). It writes all files, then runs the internal setup.

---

## Features & Behavior

- Serial Number detection
  - Looks for `SN` followed by an ID in the filename (letters/digits/_/-). Output header is `SN <ID>`.
  - If no `SN` is found, uses `SN_<stem>`.
- Page-constrained search
  - Each term can specify page ranges; search is restricted to those pages.
  - If `Pages` is empty for a term, the whole document is searched.
- Closest-number extraction
  - For each occurrence, the nearest number (by characters) is selected.
  - Numbers include formats like `1,234`, `12.5`, optional units, and signs.
- Robust extraction pipeline
  - Tries PyMuPDF, pdfminer.six, pypdf.
  - Optional OCR paths:
    - OCRmyPDF (searchable PDF pre-processing) as primary or fallback (recommended for scanned tables)
    - Direct OCR via Tesseract (PyMuPDF/pdf2image renderers)
---

## Terms File Schema

- Supports `.csv`, `.xlsx`, and `.xls` (for `.xls`, pandas+xlrd required).
- Required columns:
  - `Term`: the search term name
  - `Pages`: page ranges (e.g., `5-10; 22`)
- Mode and extraction options (optional):
  - `Mode`: `nearest` | `table(xy)` | `line` (default `nearest`)

  - For `table(xy)`:
    - `Line`: row header text (data row)
    - `Column`: column header text; supports alternatives with pipe (e.g., `Value|Nominal`)

  - For `line`:
    - `Anchor`: text that must appear on the target line (defaults to `Term`)
    - `FieldIndex`: which field after the anchor to return (1-based; `2` or `2nd` etc.)
    - `FieldSplit`: `auto` | `groups` | `tokens`
      - `groups`: split by two-or-more spaces or tabs (e.g., `Ford EIDP` | `Revision A` | `Type 2`)
      - `tokens`: split by any whitespace (single words)
      - `auto`: try `groups`, fall back to `tokens`
    - `Return`: `string` | `number` (default `number`)

  - Filters (apply to `number` return types):
    - `Units`: preferred units near the number (pipe `|` list, e.g., `lbf|psi`)
    - `Range (min)`, `Range (max)`: numeric bounds; scientific notation supported (e.g., `8E-8`)
    - Legacy `Range`: `a..b` also supported; explicit min/max take precedence

Behavior
- `table(xy)`: locate column header x-position, find row containing `Line`, return the number on that row closest (by x) to the column.
  - Falls back to `nearest` if not found.
- `line`: find a line containing `Anchor` (or `Term`), split the tail into fields, pick the `FieldIndex`-th field.
  - If `Return=string` â†’ return that text; if `Return=number` â†’ extract a number from that field (respects `Units` and `Range`).
  - Falls back to `nearest` if not found.
- `nearest` (default): prefers same-line numbers (right, then left), then adjacent lines, then by distance; respects `Units`/`Range` if provided.
- Numbers support scientific notation (e.g., `8E-8`). Dates (`MM/DD/YY` or `MM/DD/YYYY`) are allowed values.

## Command-Line Options

```
--input           Path to terms file (.csv or .xlsx). Headers: Term, Pages
--pdf-folder      Folder containing PDFs to scan
--output-xlsx     Excel workbook output (ignored: artifacts always go to run_data)
--output-json     JSON details output (ignored: artifacts always go to run_data)
--output-csv      Flat CSV summary (ignored: artifacts always go to run_data)
--window-chars    Proximity window in characters around term (default: 160)
--case-sensitive  Case-sensitive term matching (off by default)
```

Examples:
```
py .\Application\eidp_term_scanner.py --input .\user_inputs\terms.xlsx --pdf-folder ".\Data Packages" --window-chars 240 --case-sensitive
```

Line mode example
- Line text: `Title: Ford EIDP   Revision A   Type 2`
- Terms row:
  - `Term=Title`, `Mode=line`, `Anchor=Title:`, `FieldIndex=2`, `FieldSplit=groups`, `Return=string`
  - Result â†’ `Revision A`

---

## Troubleshooting

- No outputs or empty results
  - Install the Python packages listed in Install (Windows).
  - If PDFs are scanned images, install Tesseract and Poppler; verify with `tesseract --version` and `pdftoppm -v`.
- Ensure PDFs exist in the folder you pass via `--pdf-folder` (default `Data Packages`).
  - Verify `terms.csv` headers are exactly `Term, Pages` and pages are 1-indexed (e.g., `5-10; 22`).
- Excel not generated
  - Install `pandas` plus `xlsxwriter` or `openpyxl`; otherwise CSVs are written.
 - OCR not triggered
  - Ensure `easyocr`, `torch` and `torchvision` are installed if you expect OCR fallback.
  - On secure/offline machines, pre-cache EasyOCR models by initializing a Reader once on a connected machine and bundling the model folder.

OCR fallback (EasyOCR)
- Handled inside Python only, no external EXEs.
- Installs: `torch` (CPU), `torchvision`, `easyocr`.

Config file
- `user_inputs\scanner.env` supports KEY=VALUE lines to set environment without editing your shell:
  - `OCR_DPI=600` (or 800) to tune EasyOCR rendering DPI.
  - `EASYOCR_LANGS=en` to control language models.
  - `VENV_DIR` to choose a custom virtual environment path; `run.bat` will create it if missing.
  - Lines starting with `#` or `;` and blank lines are ignored.

OCR tuning (env vars)
- `TESSERACT_CMD`: full path to `tesseract.exe` (if not on PATH)
- `TESSERACT_ARGS`: extra args passed to Tesseract (default `--psm 6`). Examples: `--psm 4`, `--oem 1`.
- `OCR_DPI`: rendering DPI for OCR (default `400`, range `200..800`).
- `OCR_RENDERER`: `pymupdf` or `pdf2image` to force renderer choice (defaults to PyMuPDF when available).
  
OCRmyPDF integration (env vars)
- `USE_OCRMYPDF`: `always|primary|prefer` to pre-OCR the whole doc, or any truthy value to enable fallback when pages are empty. `off|0|no|false` disables.
- `OCRMYPDF_LANG`: language (default `eng`).
- `OCRMYPDF_OPTIMIZE`: 0â€“3 (default `1`). Higher values may require external tools like `pngquant`.
- `OCRMYPDF_BIN`: path to the `ocrmypdf` executable (if Python API import isnâ€™t available).
- `OCRMYPDF_KEEP`: `1/true` to keep temporary OCR outputs.
- Serial Number column missing
  - Ensure filenames include a pattern like `SN 1234` or `SN-ABC_09`.

---

## Privacy & Safety

- Files are processed locally; no network calls are made by this script.
- Output includes nearby text snippets for verification. Remove sensitive rows/columns as needed.

---

## Changelog

- v2: Serial-number columns, Excel results/metadata sheets, per-run outputs under run_data only, plus `VENV_DIR` support and automatic venv bootstrap in `run.bat`.
- v1: CSV/JSON summary per PDF-term, multi-extractor pipeline with optional OCR.

---

## License

Internal tooling for EIDP workflows. Adapt as needed.
