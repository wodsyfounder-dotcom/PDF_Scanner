@echo off
setlocal EnableExtensions EnableDelayedExpansion

set "ROOT=%~dp0"

rem Default Python/venv selection (may be overridden after config load if VENV_DIR is set)
set "VENV_PY=%ROOT%.venv\Scripts\python.exe"
if exist "%VENV_PY%" (
  set "PY=%VENV_PY%"
  set "PATH=%ROOT%.venv\Scripts;%PATH%"
) else (
  set "PY=py"
  where %PY% >nul 2>nul || set "PY=python"
)

rem Make vendored packages available when not installed system-wide
set "PYTHONPATH=%ROOT%Lib\site-packages;%PYTHONPATH%"

rem Detect terms file and scaffold if missing
set "TERMS_XLSX=%ROOT%user_inputs\terms.xlsx"
set "TERMS_CSV=%ROOT%user_inputs\terms.csv"
if exist "%TERMS_XLSX%" set "TERMS=%TERMS_XLSX%" & goto has_terms
if exist "%TERMS_CSV%" set "TERMS=%TERMS_CSV%" & goto has_terms
echo [WARN] No terms file found.
if not exist "%ROOT%user_inputs\terms.schema.xlsx" (
  echo [SETUP] Creating Excel terms template (user_inputs\terms.schema.xlsx)
  "%PY%" "%ROOT%scripts\generate_terms_schema.py"
)
echo Open and edit: "%ROOT%user_inputs\terms.schema.xlsx" (save as terms.xlsx when ready)
exit /b 1

:has_terms
set "IN_DIR=%ROOT%user_inputs\EIDP_Import_Docs"
set "SCANNED=%ROOT%user_inputs\Scanned_Docs"
set "OUT_DIR=%ROOT%Product_Data_File"
set "OUT_XLSX=%OUT_DIR%\scan_results.xlsx"
set "OUT_JSON=%OUT_DIR%\scan_results.json"
set "OUT_CSV=%OUT_DIR%\scan_results_flat.csv"

rem Ensure directories exist
if not exist "%IN_DIR%"       mkdir "%IN_DIR%"
if not exist "%SCANNED%"      mkdir "%SCANNED%"
if not exist "%OUT_DIR%"      mkdir "%OUT_DIR%"
if not exist "%OUT_DIR%\run_data" mkdir "%OUT_DIR%\run_data"

rem Extend PATH with local tools if present (no admin PATH changes)
if exist "%ROOT%tools\tesseract" set "PATH=%ROOT%tools\tesseract;%PATH%"
if exist "%ROOT%tools\tesseract\bin" set "PATH=%ROOT%tools\tesseract\bin;%PATH%"
if exist "%ROOT%tools\poppler\bin" set "PATH=%ROOT%tools\poppler\bin;%PATH%"
if exist "%ROOT%tools\unpaper" set "PATH=%ROOT%tools\unpaper;%PATH%"
if exist "%ROOT%tools\pngquant" set "PATH=%ROOT%tools\pngquant;%PATH%"
if exist "%ROOT%tools\qpdf" set "PATH=%ROOT%tools\qpdf;%PATH%"

rem Common system install paths
if exist "%ProgramFiles%\qpdf\bin" set "PATH=%ProgramFiles%\qpdf\bin;%PATH%"
if exist "%ProgramFiles(x86)%\qpdf\bin" set "PATH=%ProgramFiles(x86)%\qpdf\bin;%PATH%"

rem Load optional scanner config (user_inputs\scanner.env) as KEY=VALUE pairs
set "CFG=%ROOT%user_inputs\scanner.env"
if exist "%CFG%" (
  echo [RUN] Loading config: "%CFG%"
  for /f "usebackq tokens=* delims=" %%L in ("%CFG%") do (
    set "LINE=%%L"
    if not "!LINE!"=="" if not "!LINE:~0,1!"==" " if not "!LINE:~0,1!"=="#" if not "!LINE:~0,1!"==";" if not "!LINE!"=="!LINE:=!" (
      set "%%L"
    )
  )
  echo [RUN] Config parsed.
)

rem Ensure vendored packages path is prepended even if PYTHONPATH was overridden in scanner.env
set "PYTHONPATH=%ROOT%Lib\site-packages;%PYTHONPATH%"

rem If a custom VENV_DIR was provided via env or scanner.env, prefer it and bootstrap if missing
if defined VENV_DIR (
  set "VENV_PY=%VENV_DIR%\Scripts\python.exe"
  if not exist "%VENV_PY%" (
    echo [SETUP] No venv at "%VENV_DIR%". Bootstrapping...
    call "%ROOT%install.bat" "%VENV_DIR%"
  )
  if exist "%VENV_PY%" (
    set "PY=%VENV_PY%"
    set "PATH=%VENV_DIR%\Scripts;%PATH%"
  ) else (
    echo [ERROR] Failed to create venv at "%VENV_DIR%".>&2
    exit /b 1
  )
)

rem Default to venv-local ocrmypdf if available and not overridden by config
if not defined OCRMYPDF_BIN (
  if defined VENV_DIR (
    if exist "%VENV_DIR%\Scripts\ocrmypdf.exe" set "OCRMYPDF_BIN=%VENV_DIR%\Scripts\ocrmypdf.exe"
  ) else (
    if exist "%ROOT%.venv\Scripts\ocrmypdf.exe" set "OCRMYPDF_BIN=%ROOT%.venv\Scripts\ocrmypdf.exe"
  )
)

rem Also extend PATH with common system install locations (session-only)
rem - Tesseract (default installer path)
if exist "%ProgramFiles%\Tesseract-OCR\tesseract.exe" (
  set "PATH=%ProgramFiles%\Tesseract-OCR;%PATH%"
  set "TESSERACT_CMD=%ProgramFiles%\Tesseract-OCR\tesseract.exe"
)
if exist "%ProgramFiles(x86)%\Tesseract-OCR\tesseract.exe" (
  set "PATH=%ProgramFiles(x86)%\Tesseract-OCR;%PATH%"
  set "TESSERACT_CMD=%ProgramFiles(x86)%\Tesseract-OCR\tesseract.exe"
)

rem Report Tesseract status so users know whether OCR can run (simplified)
if defined TESSERACT_CMD (
  echo [RUN] Tesseract: "%TESSERACT_CMD%"
) else (
  echo [RUN] Tesseract: not detected on default paths; proceeding.
)

rem - Poppler (various Windows builds install under poppler-*\bin)
if exist "%ProgramFiles%\poppler\bin" set "PATH=%ProgramFiles%\poppler\bin;%PATH%"
for /d %%P in ("%ProgramFiles%\poppler-*") do (
  if exist "%%P\bin" set "PATH=%%P\bin;%PATH%"
)

rem - Ghostscript (helpful for OCRmyPDF optimization)
if exist "%ProgramFiles%\gs" (
  for /d %%G in ("%ProgramFiles%\gs\gs*") do (
    if exist "%%G\bin" set "PATH=%%G\bin;%PATH%"
  )
)

echo [RUN] Python: "%PY%"
echo [RUN] Terms : "%TERMS%"  (use .xlsx/.csv)
echo [RUN] PDFs  : "%IN_DIR%"
echo [RUN] Out   : "%OUT_DIR%" (per-run outputs saved under run_data)

rem Default: on-demand OCR via core scanner (table-first is opt-in)
if not defined SCAN_FROM_TABLES set "SCAN_FROM_TABLES=0"

if "%SCAN_FROM_TABLES%"=="1" (
  rem Prebuild spatial/table for all PDFs
  if not defined SPATIAL_DPI set "SPATIAL_DPI=800"
  if not defined SPATIAL_LANGS set "SPATIAL_LANGS=en"
  if not defined SPATIAL_MIN_CONF set "SPATIAL_MIN_CONF=0.0"
  if not defined SPATIAL_ROW_FACTOR set "SPATIAL_ROW_FACTOR=0.6"
  if not defined SPATIAL_GAP_MULT set "SPATIAL_GAP_MULT=2.5"
  if not defined SPATIAL_CENTER_THRESH_MULT set "SPATIAL_CENTER_THRESH_MULT=2.0"
  if not defined SPATIAL_BUILD_TABLE set "SPATIAL_BUILD_TABLE=1"
  if not exist "%OUT_DIR%\easy_spatial" mkdir "%OUT_DIR%\easy_spatial"
  for %%F in ("%IN_DIR%\*.pdf") do (
    set "PDF_PATH=%%~fF"
    set "PDF_NAME=%%~nxF"
    set "PDF_STEM=%%~nF"
    set "PDF_STEM_SAFE=!PDF_STEM: =_!"
    set "EZ_DIR=%OUT_DIR%\easy_spatial\!PDF_STEM_SAFE!"
    if not exist "!EZ_DIR!" mkdir "!EZ_DIR!"
    echo [SPATIAL] OCR dump: "!PDF_NAME!" -> "!EZ_DIR!"
    "%PY%" "%ROOT%scripts\easyocr_dump_spatial.py" ^
      --pdf "!PDF_PATH!" ^
      --out-dir "!EZ_DIR!" ^
      --dpi "%SPATIAL_DPI%" ^
      --langs "%SPATIAL_LANGS%" ^
      --min-conf "%SPATIAL_MIN_CONF%" ^
      --row-factor "%SPATIAL_ROW_FACTOR%"
    if not "!ERRORLEVEL!"=="0" echo [WARN] easyocr_dump_spatial failed for "!PDF_NAME!" (continuing)
    if not "%SPATIAL_BUILD_TABLE%"=="0" (
      for %%J in ("!EZ_DIR!\!PDF_STEM!_page_*.json") do (
        set "PAGE_PREFIX=%%~nJ"
        set "PAGE_PREFIX_SAFE=!PAGE_PREFIX: =_!"
        "%PY%" "%ROOT%scripts\build_table_from_tokens.py" ^
          --tokens-json "%%~fJ" ^
          --out-prefix "!EZ_DIR!\!PAGE_PREFIX_SAFE!" ^
          --row-factor %SPATIAL_ROW_FACTOR% ^
          --gap-mult %SPATIAL_GAP_MULT% ^
          --center-thresh-mult %SPATIAL_CENTER_THRESH_MULT% ^
          --use-first-row-as-header
      )
    )
  )
  rem Scan from the generated tables to produce run_data and update EIDP_data.csv
  "%PY%" "%ROOT%scripts\scan_from_tables.py" ^
    --terms "%TERMS%" ^
    --pdf-dir "%IN_DIR%" ^
    --spatial-dir "%OUT_DIR%\easy_spatial" ^
    --out-dir "%OUT_DIR%\run_data"
  set "RC=%ERRORLEVEL%"
  if not "%RC%"=="0" exit /b %RC%
) else (
  rem Force OCR settings to EasyOCR-only fallback; disable OCRmyPDF
  if not defined USE_OCRMYPDF set "USE_OCRMYPDF=off"
  if not defined OCR_RENDERER set "OCR_RENDERER=easyocr"
  if not defined USE_EASYOCR_XY set "USE_EASYOCR_XY=0"

  "%PY%" "%ROOT%Application\eidp_term_scanner.py" ^
    --input "%TERMS%" ^
    --pdf-folder "%IN_DIR%" ^
    --output-xlsx "%OUT_XLSX%" ^
    --output-json "%OUT_JSON%" ^
    --output-csv "%OUT_CSV%" ^
    --scanned-folder "%SCANNED%" ^
    %*

  set "RC=%ERRORLEVEL%"
  if not "%RC%"=="0" exit /b %RC%
)

rem ---------------------------------------------------------------------------
rem Post-scan spatial build (only for PDFs with missing/empty results)
rem Controlled by scanner.env (KEY=VALUE):
rem   SPATIAL_ON_MISSING=0 (default) | 1
rem   SPATIAL_DPI=800, SPATIAL_LANGS=en, SPATIAL_MIN_CONF=0.0
rem   SPATIAL_ROW_FACTOR=0.6, SPATIAL_GAP_MULT=2.5, SPATIAL_CENTER_THRESH_MULT=2.0
rem   SPATIAL_BUILD_TABLE=1 (build per-page .table.csv/.xlsx/.svg)
rem ---------------------------------------------------------------------------
if not defined SPATIAL_ON_MISSING set "SPATIAL_ON_MISSING=0"
if not defined SPATIAL_DPI set "SPATIAL_DPI=800"
if not defined SPATIAL_LANGS set "SPATIAL_LANGS=en"
if not defined SPATIAL_MIN_CONF set "SPATIAL_MIN_CONF=0.0"
if not defined SPATIAL_ROW_FACTOR set "SPATIAL_ROW_FACTOR=0.6"
if not defined SPATIAL_GAP_MULT set "SPATIAL_GAP_MULT=2.5"
if not defined SPATIAL_CENTER_THRESH_MULT set "SPATIAL_CENTER_THRESH_MULT=2.0"
if not defined SPATIAL_BUILD_TABLE set "SPATIAL_BUILD_TABLE=1"

set "LAST_RUN="
for /f "delims=" %%D in ('dir /ad /b /o:-d "%OUT_DIR%\run_data" 2^>nul') do (
  if not defined LAST_RUN set "LAST_RUN=%%D"
)
if "%SPATIAL_ON_MISSING%"=="1" if defined LAST_RUN (
  if not exist "%OUT_DIR%\easy_spatial" mkdir "%OUT_DIR%\easy_spatial"
  echo [SPATIAL] Checking for missing results in: "%OUT_DIR%\run_data\%LAST_RUN%\by_pdf"
  "%PY%" "%ROOT%scripts\find_pdfs_needing_spatial.py" ^
    --by-pdf-dir "%OUT_DIR%\run_data\%LAST_RUN%\by_pdf" ^
    --pdf-dir "%IN_DIR%" > "%OUT_DIR%\run_data\%LAST_RUN%\_needs_spatial.txt"
  for /f "usebackq delims=" %%P in ("%OUT_DIR%\run_data\%LAST_RUN%\_needs_spatial.txt") do (
    set "PDF_PATH=%%~fP"
    set "PDF_NAME=%%~nxP"
    set "PDF_STEM=%%~nP"
    set "PDF_STEM_SAFE=!PDF_STEM: =_!"
    set "EZ_DIR=%OUT_DIR%\easy_spatial\!PDF_STEM_SAFE!"
    if not exist "!EZ_DIR!" mkdir "!EZ_DIR!"
    echo [SPATIAL] OCR dump (missing): "!PDF_NAME!" -> "!EZ_DIR!"
    "%PY%" "%ROOT%scripts\easyocr_dump_spatial.py" ^
      --pdf "!PDF_PATH!" ^
      --out-dir "!EZ_DIR!" ^
      --dpi "%SPATIAL_DPI%" ^
      --langs "%SPATIAL_LANGS%" ^
      --min-conf "%SPATIAL_MIN_CONF%" ^
      --row-factor "%SPATIAL_ROW_FACTOR%"
    if not "!ERRORLEVEL!"=="0" echo [WARN] easyocr_dump_spatial failed for "!PDF_NAME!" (continuing)

    echo [SPATIAL] Dataset build (missing): stem="!PDF_STEM!"
    "%PY%" "%ROOT%scripts\build_search_dataset.py" ^
      --dir "!EZ_DIR!" ^
      --stem "!PDF_STEM!" ^
      --out-prefix "!EZ_DIR!\!PDF_STEM_SAFE!" ^
      --row-factor %SPATIAL_ROW_FACTOR% ^
      --gap-mult %SPATIAL_GAP_MULT% ^
      --center-thresh-mult %SPATIAL_CENTER_THRESH_MULT%
    if not "!ERRORLEVEL!"=="0" echo [WARN] build_search_dataset failed for "!PDF_NAME!" (continuing)

    if not "%SPATIAL_BUILD_TABLE%"=="0" (
      for %%J in ("!EZ_DIR!\!PDF_STEM!_page_*.json") do (
        set "PAGE_PREFIX=%%~nJ"
        set "PAGE_PREFIX_SAFE=!PAGE_PREFIX: =_!"
        "%PY%" "%ROOT%scripts\build_table_from_tokens.py" ^
          --tokens-json "%%~fJ" ^
          --out-prefix "!EZ_DIR!\!PAGE_PREFIX_SAFE!" ^
          --row-factor %SPATIAL_ROW_FACTOR% ^
          --gap-mult %SPATIAL_GAP_MULT% ^
          --center-thresh-mult %SPATIAL_CENTER_THRESH_MULT% ^
          --use-first-row-as-header
      )
    )
  )
)

if defined LAST_RUN (
  echo [DONE] Run folder: "%OUT_DIR%\run_data\%LAST_RUN%"
) else (
  echo [DONE] Check: "%OUT_DIR%\run_data" for this run.
)
echo [DONE] Top-level aggregate updated (if enabled): "%OUT_DIR%\EIDP_data.csv"
endlocal & exit /b 0
