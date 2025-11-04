@echo off
setlocal EnableExtensions EnableDelayedExpansion

set "ROOT=%~dp0"

rem Prefer local venv if present
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

rem Load optional scanner config (user_inputs\scanner.env) as KEY=VALUE lines
set "CFG=%ROOT%user_inputs\scanner.env"
if exist "%CFG%" (
  if not "%QUIET%"=="1" echo [RUN] Loading config: "%CFG%"
  for /f "usebackq tokens=* delims=" %%L in ("%CFG%") do (
    set "LINE=%%L"
    if not "!LINE!"=="" if not "!LINE:~0,1!"==" " if not "!LINE:~0,1!"=="#" if not "!LINE:~0,1!"==";" if not "!LINE!"=="!LINE:=!" (
      for /f "tokens=1,* delims==" %%A in ("!LINE!") do (
        set "K=%%~A"
        set "V=%%~B"
        if defined K (
          for /f "tokens=1 delims=#;" %%C in ("!V!") do set "V=%%~C"
          set "V=!V:~0!"
          for /f "tokens=* delims= " %%D in ("!V!") do set "V=%%~D"
          if not "!V!"=="" set "!K!=!V!"
        )
      )
    )
  )
  if not "%QUIET%"=="1" echo [RUN] Config parsed.
)

rem Ensure vendored packages path is prepended even if PYTHONPATH was overridden in scanner.env
set "PYTHONPATH=%ROOT%Lib\site-packages;%PYTHONPATH%"

rem Optional venv override via scanner.env
if defined VENV_DIR (
  set "VENV_PY=%VENV_DIR%\Scripts\python.exe"
  if not exist "%VENV_PY%" (
    if not "%QUIET%"=="1" echo [SETUP] No venv at "%VENV_DIR%". Bootstrapping...
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

if not "%QUIET%"=="1" (
  echo [RUN] Python: "%PY%"
  echo [RUN] Terms : "%TERMS%"  (use .xlsx/.csv)
  echo [RUN] PDFs  : "%IN_DIR%"
  echo [RUN] Out   : "%OUT_DIR%" (per-run outputs saved under run_data)
  echo [RUN] OCR    : OCR_MODE=%OCR_MODE%, USE_EASYOCR_XY=%USE_EASYOCR_XY%
  if defined FORCE_OCR echo [RUN] OCR    : FORCE_OCR=%FORCE_OCR%
  if defined OCR_DPI echo [RUN] OCR    : OCR_DPI=%OCR_DPI%
  if defined EASYOCR_LANGS echo [RUN] OCR    : EASYOCR_LANGS=%EASYOCR_LANGS%
)

rem Default OCR policy
if not defined OCR_MODE set "OCR_MODE=fallback"

set "QUIET_FLAG="
if /I "%QUIET%"=="1" set "QUIET_FLAG=--quiet"

"%PY%" "%ROOT%Application\eidp_term_scanner.py" ^
  --input "%TERMS%" ^
  --pdf-folder "%IN_DIR%" ^
  --output-xlsx "%OUT_XLSX%" ^
  --output-json "%OUT_JSON%" ^
  --output-csv "%OUT_CSV%" ^
  --scanned-folder "%SCANNED%" ^
  %QUIET_FLAG% ^
  %*

set "RC=%ERRORLEVEL%"
if not "%RC%"=="0" exit /b %RC%

rem Locate the most recent run_data folder for convenience
set "LAST_RUN="
for /f "delims=" %%D in ('dir /ad /b /o:-d "%OUT_DIR%\run_data" 2^>nul') do (
  if not defined LAST_RUN set "LAST_RUN=%%D"
)
if defined LAST_RUN (
  echo [DONE] Run folder: "%OUT_DIR%\run_data\%LAST_RUN%"
) else (
  echo [DONE] Check: "%OUT_DIR%\run_data" for this run.
)
endlocal & exit /b 0
