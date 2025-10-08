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
- System-wide (if allowed):
  - `py -m pip install --upgrade pip`
  - `py -m pip install pymupdf pdfminer.six pypdf pandas xlsxwriter pytesseract pillow pdf2image`
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

---

## Super Simple Run

- One-time setup: run `auto-scaffold.bat` (creates folders and a sample `terms.csv`).
- Put PDFs in `user_inputs\EIDP_Import_Docs`.
- Edit `user_inputs\terms.csv` (headers: `Term, Pages`).
- Run one of:
  - Windows: double-click `run.bat`
  - PowerShell: `./run.ps1`
  - Python: `py .\Application\eidp_term_scanner.py --input .\user_inputs\terms.csv --pdf-folder .\user_inputs\EIDP_Import_Docs --scanned-folder .\user_inputs\Scanned_Docs`

Outputs are saved under `Product_Data_File\run_data\<timestamp>`. The only top-level file updated is `Product_Data_File\EIDP_data.csv`.

---

## What You Get

- `Product_Data_File\run_data\<timestamp>\scan_results.xlsx` (if Excel writers available)
  - `results` sheet: Term, Pages, and one column per SN (Serial Number from filename)
  - `metadata` sheet: pdf_file, serial_number, term, found, page, number, units, context, method_pipeline
- CSV fallbacks (written into the same `run_data` folder if Excel libs are missing):
  - `scan_results.results.csv` and `scan_results.metadata.csv`
- `scan_results_flat.csv` and `scan_results.json` (audit) also live in the per-run folder
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
  - Tries PyMuPDF, pdfminer.six, pypdf; falls back to OCR via Tesseract if needed (when installed).
- Auto-move scanned PDFs
  - After scanning, PDFs are moved from `--pdf-folder` to `--scanned-folder`.

---

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
py .\Application\eidp_term_scanner.py --input .\user_inputs\terms.csv --pdf-folder .\user_inputs\EIDP_Import_Docs --window-chars 240 --case-sensitive
```

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
- Serial Number column missing
  - Ensure filenames include a pattern like `SN 1234` or `SN-ABC_09`.

---

## Privacy & Safety

- Files are processed locally; no network calls are made by this script.
- Output includes nearby text snippets for verification. Remove sensitive rows/columns as needed.

---

## Changelog

- v2: Serial-number columns, Excel results/metadata sheets, auto-move scanned PDFs, per-run outputs under run_data only.
- v1: CSV/JSON summary per PDF-term, multi-extractor pipeline with optional OCR.

---

## License

Internal tooling for EIDP workflows. Adapt as needed.
