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

echo [DONE] Check: "%OUT_DIR%\run_data" for this run; top-level EIDP_data.csv was updated.
endlocal & exit /b 0
