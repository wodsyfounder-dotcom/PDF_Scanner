@echo off
setlocal EnableExtensions

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
  exit /b 1
)

set "VPY=%VENV_DIR%\Scripts\python.exe"
if not exist "%VPY%" (
  echo [ERROR] venv python not found: "%VPY%"
  exit /b 1
)

echo [SETUP] Upgrading pip...
"%VPY%" -m pip install --upgrade pip
if errorlevel 1 echo [WARN] pip upgrade had warnings.

echo [SETUP] Installing local Python packages...
"%VPY%" -m pip install ^
  pymupdf ^
  pdfminer.six ^
  pypdf ^
  pandas ^
  xlsxwriter ^
  openpyxl ^
  xlrd ^
  pytesseract ^
  pillow ^
  pdf2image ^
  ocrmypdf
if errorlevel 1 (
  echo [ERROR] Package install failed.
  exit /b 1
)

rem Also vendor runtime deps into repo-local Lib\site-packages (for non-venv runs)
set "LOCAL_SITE=%ROOT%Lib\site-packages"
if not exist "%LOCAL_SITE%" mkdir "%LOCAL_SITE%"
echo [SETUP] Vendoring Python deps to Lib\site-packages:
echo         pymupdf, pdfminer.six, pypdf, pillow, pytesseract, pdf2image, pandas, xlsxwriter, openpyxl, xlrd, ocrmypdf
"%VPY%" -m pip install --upgrade --no-warn-script-location --target "%LOCAL_SITE%" ^
  pymupdf pdfminer.six pypdf pillow pytesseract pdf2image pandas xlsxwriter openpyxl xlrd ocrmypdf
if errorlevel 1 (
  echo [WARN] Vendoring had warnings/failures. Non-venv runs may miss some features.
)

rem --- Scaffold expected folders and sample terms file ---
if not exist "%ROOT%user_inputs" mkdir "%ROOT%user_inputs"
if not exist "%ROOT%user_inputs\EIDP_Import_Docs" mkdir "%ROOT%user_inputs\EIDP_Import_Docs"
if not exist "%ROOT%user_inputs\Scanned_Docs" mkdir "%ROOT%user_inputs\Scanned_Docs"
if not exist "%ROOT%Product_Data_File" mkdir "%ROOT%Product_Data_File"
if not exist "%ROOT%Product_Data_File\run_data" mkdir "%ROOT%Product_Data_File\run_data"
if not exist "%ROOT%user_inputs\terms.xlsx" if not exist "%ROOT%user_inputs\terms.csv" if not exist "%ROOT%user_inputs\terms.schema.xlsx" (
  echo [SETUP] Creating Excel terms template (user_inputs\terms.schema.xlsx)
  "%VPY%" "%ROOT%scripts\generate_terms_schema.py"
)

rem Create scanner.env with sensible defaults if missing
if not exist "%ROOT%user_inputs\scanner.env" (
  echo [SETUP] Creating default user_inputs\scanner.env
  >  "%ROOT%user_inputs\scanner.env" echo # Scanner configuration (KEY=VALUE)
  >> "%ROOT%user_inputs\scanner.env" echo #USE_OCRMYPDF=primary
  >> "%ROOT%user_inputs\scanner.env" echo OCRMYPDF_FORCE=1
  >> "%ROOT%user_inputs\scanner.env" echo OCRMYPDF_LANG=eng
  >> "%ROOT%user_inputs\scanner.env" echo OCRMYPDF_OPTIMIZE=1
  >> "%ROOT%user_inputs\scanner.env" echo #OCR_DPI=600
  >> "%ROOT%user_inputs\scanner.env" echo #TESSERACT_ARGS=--psm 4
  >> "%ROOT%user_inputs\scanner.env" echo #VENV_DIR=%ROOT%.venv
  >> "%ROOT%user_inputs\scanner.env" echo #TESSERACT_CMD=C:\\Program Files\\Tesseract-OCR\\tesseract.exe
)

rem --- Check for Tesseract OCR binary (native dependency, not a Python package) ---
set "TESS_PATH_64=%ProgramFiles%\Tesseract-OCR\tesseract.exe"
set "TESS_PATH_86=%ProgramFiles(x86)%\Tesseract-OCR\tesseract.exe"
set "TESS_TOOL_1=%ROOT%tools\tesseract\tesseract.exe"
set "TESS_TOOL_2=%ROOT%tools\tesseract\bin\tesseract.exe"

if exist "%TESS_PATH_64%" (
  echo [INFO] Tesseract found: "%TESS_PATH_64%"
) else if exist "%TESS_PATH_86%" (
  echo [INFO] Tesseract found: "%TESS_PATH_86%"
) else if exist "%TESS_TOOL_1%" (
  echo [INFO] Tesseract (portable) found: "%TESS_TOOL_1%"
) else if exist "%TESS_TOOL_2%" (
  echo [INFO] Tesseract (portable) found: "%TESS_TOOL_2%"
) else (
  echo [WARN] Tesseract not detected. OCR for image-only PDFs will be unavailable.
  echo        Install via Windows installer (recommended) or place a portable build under:
  echo        "%ROOT%tools\tesseract\"  (so that "tesseract.exe" is inside this folder)
) 

rem --- Check for Ghostscript (required by OCRmyPDF) and QPDF (optional) ---
set "GS_DIR64=%ProgramFiles%\gs"
set "GS_DIR86=%ProgramFiles(x86)%\gs"
set "QPDF_BIN1=%ProgramFiles%\qpdf\bin\qpdf.exe"
set "QPDF_BIN2=%ProgramFiles(x86)%\qpdf\bin\qpdf.exe"

set "_GS_FOUND="
for /d %%G in ("%GS_DIR64%\gs*") do (
  if exist "%%~fG\bin\gswin64c.exe" set "_GS_FOUND=%%~fG\bin\gswin64c.exe"
)
if not defined _GS_FOUND (
  for /d %%G in ("%GS_DIR86%\gs*") do (
    if exist "%%~fG\bin\gswin32c.exe" set "_GS_FOUND=%%~fG\bin\gswin32c.exe"
  )
)
if defined _GS_FOUND (
  echo [INFO] Ghostscript found: "%_GS_FOUND%"
) else (
  echo [WARN] Ghostscript not detected. OCRmyPDF will fail until installed (gswin64c.exe).
)

if exist "%QPDF_BIN1%" (
  echo [INFO] QPDF found: "%QPDF_BIN1%"
) else if exist "%QPDF_BIN2%" (
  echo [INFO] QPDF found: "%QPDF_BIN2%"
) else (
  echo [INFO] QPDF not detected (optional). OCRmyPDF can still run using bundled libs.
)

echo.
echo [READY] Local environment set up.
echo   - Python venv: "%VENV_DIR%"
echo   - Vendored Python deps in Lib\site-packages for non-venv runs
echo   - Optional OCR tools (no PATH edits):
echo       Create "%ROOT%tools\tesseract" and place tesseract.exe there
echo       Create "%ROOT%tools\poppler\bin" and place pdftoppm.exe there
echo   - OCRmyPDF installed inside venv. Use: ".venv\\Scripts\\ocrmypdf.exe"
echo.
echo The run scripts will prefer .venv and extend PATH with ^"tools^" if present.
echo Next steps:
echo   1) Place PDFs under user_inputs\EIDP_Import_Docs
echo   2) Edit user_inputs\terms.csv
echo   3) Run: run.bat   (or  .\run.ps1)

endlocal & exit /b 0
