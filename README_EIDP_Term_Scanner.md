# EIDP Term Scanner

Scan a folder of EIDP PDFs for a list of terms (from CSV/XLSX), find the closest numeric value near each term within specified page ranges, and export a matrix (rows=terms, columns=Serial Numbers) plus detailed metadata. Scanned PDFs are moved out so the next run only processes new files.

---

## Install (Windows)

- No admin / no PATH changes (recommended):
  - Run `install.bat` to create a local `.venv` and install Python packages there.
  - Optionally place portable tools:
    - `tools\tesseract` (folder containing `tesseract.exe`)
    - `tools\poppler\bin` (folder containing `pdftoppm.exe`)
  - The run scripts will auto-use `.venv` and extend PATH to these folders for the session.
  - Optional custom venv location: `install.bat C:\\Path\\To\\PDF_Scanner.venv` or set `VENV_DIR` in `user_inputs\\scanner.env`.
  - Optional custom venv location: `install.bat C:\\Path\\To\\PDF_Scanner.venv` or set `VENV_DIR` in `user_inputs\\scanner.env`.
- System-wide (if allowed):
  - `py -m pip install --upgrade pip`
  - `py -m pip install pymupdf pdfminer.six pypdf pandas xlsxwriter pytesseract pillow pdf2image xlrd`
- Install Tesseract OCR (needed for scanned/image-only PDFs):
  - Option A (Chocolatey, admin PowerShell): `choco install tesseract`
  - Option B: Use the Tesseract Windows installer and ensure `tesseract.exe` is on PATH
  - Verify: `tesseract --version`
- Install Poppler (renderer used by pdf2image for OCR):
  - Option A (Chocolatey, admin PowerShell): `choco install poppler`
  - Option B: Download Poppler for Windows and add its `bin` folder to PATH
  - Verify: `pdftoppm -v`

Notes:
- If your PDFs are text-based (not scans), you can skip Tesseract/Poppler. At least one text extractor (PyMuPDF or pdfminer.six) is recommended.
- Excel output requires `pandas` and either `xlsxwriter` or `openpyxl`. If missing, CSV fallbacks are written instead.
- Reading `.xls` terms requires `pandas` and `xlrd`. If unavailable, save as `.xlsx` or `.csv`.

---

## Super Simple Run

- One-time setup: install.bat scaffolds folders and drops a sample terms CSV if none exists:
  - `user_inputs\\EIDP_Import_Docs`, `user_inputs\\Scanned_Docs`
  - `user_inputs\\terms.xlsx` or `user_inputs\\terms.csv` (headers: `Term, Pages`)
- Put PDFs in `user_inputs\EIDP_Import_Docs`.
- Edit `user_inputs\terms.xlsx` or `user_inputs\terms.csv` (headers: `Term, Pages`).
- Optional: set runtime knobs in `user_inputs\scanner.env` (KEY=VALUE lines)
- Run one of:
  - Windows: double-click `run.bat`
  - Python: `py .\Application\eidp_term_scanner.py --input .\user_inputs\terms.xlsx --pdf-folder .\user_inputs\EIDP_Import_Docs --scanned-folder .\user_inputs\Scanned_Docs --window-chars 400`

Minimal transfer (work PC): copy `Application\eidp_term_scanner.py`, `Application\eidp_term_scanner.core.py`, `run.bat`, `install.bat`, and `user_inputs` (or at least `user_inputs\scanner.env` and your `terms.xlsx`). Create `user_inputs\EIDP_Import_Docs` and `user_inputs\Scanned_Docs` on first run. Set `VENV_DIR` in `scanner.env` if you want the venv outside the repo.
Outputs are saved under `Product_Data_File\run_data\<timestamp>`. The only top-level file updated is `Product_Data_File\EIDP_data.csv`.

---

## What You Get

- `Product_Data_File\run_data\<timestamp>\scan_results.xlsx` (if Excel writers available)
  - `results` sheet: Term, Pages, and one column per SN (Serial Number from filename)
  - `metadata` sheet: pdf_file, serial_number, term, found, page, number, units, context, method_pipeline
- CSV fallbacks (written into the same `run_data` folder if Excel libs are missing):
  - `scan_results.results.csv` and `scan_results.metadata.csv`
- `scan_results_flat.csv` and `scan_results.json` (audit) also live in the per-run folder
  - `by_pdf/` folder contains one JSON per PDF with just that fileâ€™s term results
- Top-level aggregate that grows over time: `Product_Data_File\EIDP_data.csv`

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
- Auto-move scanned PDFs
  - After scanning, PDFs are moved from `--pdf-folder` to `--scanned-folder`.

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
--scanned-folder  Destination for processed PDFs (default: "Scanned Docs")
--window-chars    Proximity window in characters around term (default: 160)
--case-sensitive  Case-sensitive term matching (off by default)
```

Examples:
```
py .\Application\eidp_term_scanner.py --input .\user_inputs\terms.xlsx --pdf-folder .\user_inputs\EIDP_Import_Docs --window-chars 240 --case-sensitive
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
  - Ensure PDFs exist in `user_inputs\EIDP_Import_Docs` and are not already moved to `Scanned_Docs`.
  - Verify `terms.csv` headers are exactly `Term, Pages` and pages are 1-indexed (e.g., `5-10; 22`).
- Excel not generated
  - Install `pandas` plus `xlsxwriter` or `openpyxl`; otherwise CSVs are written.
- OCR not triggered
  - Confirm Tesseract is installed and on PATH; Poppler present for `pdf2image`.
  - Optional: set `TESSERACT_CMD` env var to the full path to `tesseract.exe`.
  - If using the local venv, `run.bat` adds `.venv\Scripts` to PATH so `ocrmypdf` and Python entry points are available.

OCRmyPDF (optional, installed into venv by install.bat)
- Use the venv-local command:
  - Windows CMD: `.venv\Scripts\ocrmypdf.exe -l eng --force-ocr --rotate-pages --deskew --clean --optimize 3 --tesseract-pagesegmode 4 "user_inputs\EIDP_Import_Docs\SN 1111.pdf" "user_inputs\EIDP_Import_Docs\SN 1111.ocr.pdf"`
  - PowerShell: `.venv/Scripts/ocrmypdf.exe ...`
- Requires Tesseract; Ghostscript recommended (for cleanup/compression). Poppler not required for OCRmyPDF.

Config file
- `user_inputs\scanner.env` supports KEY=VALUE lines to set environment without editing your shell:
  - `USE_OCRMYPDF=primary` to pre-OCR all docs; or `USE_OCRMYPDF=1` to enable fallback.
  - `OCRMYPDF_FORCE=0` to avoid forcing OCR on pages that already contain text.
  - `OCRMYPDF_LANG=eng`, `OCRMYPDF_OPTIMIZE=1` for language/opt level.
  - `OCR_RENDERER=pymupdf|pdf2image`, `OCR_DPI=600`, `TESSERACT_ARGS=--psm 4` to tune direct OCR when used.
  - `TESSERACT_CMD` or `OCRMYPDF_BIN` to point to tool executables if not on PATH.
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

- v2: Serial-number columns, Excel results/metadata sheets, auto-move scanned PDFs, per-run outputs under run_data only, plus `VENV_DIR` support and automatic venv bootstrap in `run.bat`.
- v1: CSV/JSON summary per PDF-term, multi-extractor pipeline with optional OCR.

---

## License

Internal tooling for EIDP workflows. Adapt as needed.
