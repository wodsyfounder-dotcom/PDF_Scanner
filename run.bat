@echo off
setlocal EnableExtensions

set "ROOT=%~dp0"
set "VENV_PY=%ROOT%.venv\Scripts\python.exe"
if exist "%VENV_PY%" (
  set "PY=%VENV_PY%"
  rem Ensure local venv scripts (ocrmypdf, etc.) are accessible on PATH
  set "PATH=%ROOT%.venv\Scripts;%PATH%"
) else (
  set "PY=py"
  where %PY% >nul 2>nul || set "PY=python"
)

set "TERMS=%ROOT%user_inputs\terms.csv"
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
