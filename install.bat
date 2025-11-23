@echo off
setlocal EnableExtensions EnableDelayedExpansion

set "ROOT=%~dp0"
set "PY=py"
rem Optional first arg: custom venv directory. Defaults to %ROOT%.venv
set "VENV_DIR=%~1"
if "%VENV_DIR%"=="" set "VENV_DIR=%ROOT%.venv"
where %PY% >nul 2>nul || set "PY=python"

echo [SETUP] Creating local virtual environment: "%VENV_DIR%" ...
"%PY%" -m venv "%VENV_DIR%"
if errorlevel 1 (
  echo [ERROR] Failed to create venv. Ensure Python 3 is installed.
  endlocal & exit /b 1
)

set "VPY=%VENV_DIR%\Scripts\python.exe"
if not exist "%VPY%" (
  echo [ERROR] venv python not found: "%VPY%"
  endlocal & exit /b 1
)

echo [SETUP] Upgrading pip...
"%VPY%" -m pip install --upgrade pip
if errorlevel 1 echo [WARN] pip upgrade had warnings.

echo [SETUP] Installing required Python packages (minimal + UI)...
"%VPY%" -m pip install ^
  pymupdf ^
  pandas ^
  openpyxl ^
  XlsxWriter ^
  matplotlib ^
  opencv-python-headless ^
  PySide6
if errorlevel 1 (
  echo [ERROR] Package install failed.
  endlocal & exit /b 1
)

echo [SETUP] Installing EasyOCR + CPU torch/torchvision (OCR fallback)...
"%VPY%" -m pip install --index-url https://download.pytorch.org/whl/cpu torch torchvision
if errorlevel 1 (
  echo [ERROR] PyTorch CPU install failed.
  endlocal & exit /b 1
)
"%VPY%" -m pip install easyocr
if errorlevel 1 (
  echo [ERROR] EasyOCR install failed.
  endlocal & exit /b 1
)

rem Also vendor runtime deps into repo-local Lib\site-packages (for non-venv runs)
set "LOCAL_SITE=%ROOT%Lib\site-packages"
if not exist "%LOCAL_SITE%" mkdir "%LOCAL_SITE%"
echo [SETUP] Vendoring Python deps to Lib\site-packages (minimal):
echo         pymupdf, pandas, openpyxl, XlsxWriter, matplotlib, opencv-python-headless, PySide6, easyocr, torch, torchvision
"%VPY%" -m pip install --upgrade --no-warn-script-location --target "%LOCAL_SITE%" ^
  pymupdf pandas openpyxl XlsxWriter matplotlib opencv-python-headless PySide6 easyocr
"%VPY%" -m pip install --upgrade --no-warn-script-location --index-url https://download.pytorch.org/whl/cpu --target "%LOCAL_SITE%" ^
  torch torchvision
if errorlevel 1 (
  echo [WARN] Vendoring had warnings/failures. Non-venv runs may miss some features.
)

rem --- Scaffold expected folders and sample terms file ---
if not exist "%ROOT%user_inputs" mkdir "%ROOT%user_inputs"
if not exist "%ROOT%Data Packages" mkdir "%ROOT%Data Packages"
if not exist "%ROOT%Product_Data_File" mkdir "%ROOT%Product_Data_File"
if not exist "%ROOT%Product_Data_File\run_data" mkdir "%ROOT%Product_Data_File\run_data"
if not exist "%ROOT%user_inputs\terms.schema.smartsnap.xlsx" (
  echo [SETUP] Creating Smart-Snap template (user_inputs\terms.schema.smartsnap.xlsx)
  "%VPY%" "%ROOT%scripts\generate_terms_schema_smartsnap.py"
) else (
  echo [INFO] Smart-Snap template already present.
)

rem Create scanner.env with sensible defaults if missing
if not exist "%ROOT%user_inputs\scanner.env" (
  echo [SETUP] Creating default user_inputs\scanner.env
  >  "%ROOT%user_inputs\scanner.env" echo # Scanner configuration (KEY=VALUE)
  >> "%ROOT%user_inputs\scanner.env" echo QUIET=1
  >> "%ROOT%user_inputs\scanner.env" echo #VENV_DIR=%ROOT%.venv
  >> "%ROOT%user_inputs\scanner.env" echo #OCR_MODE=fallback   ^# fallback|ocr_only|no_ocr
  >> "%ROOT%user_inputs\scanner.env" echo #OCR_DPI=600
  >> "%ROOT%user_inputs\scanner.env" echo #EASYOCR_LANGS=en
)

echo.
echo [READY] Local environment set up.
echo   - Python venv: "%VENV_DIR%"
echo   - Vendored Python deps in Lib\site-packages for non-venv runs
echo   - OCR handled via EasyOCR (pure Python). No external OCR tools required.
echo.
echo Next steps:
echo   1) Point the app at your PDF repository (default: "Data Packages")
echo   2) Edit user_inputs\terms.xlsx (or .csv)
echo   3) Run: run.bat   or  "%VENV_DIR%\Scripts\python.exe" ui_next\qt_main.py

endlocal & exit /b 0
