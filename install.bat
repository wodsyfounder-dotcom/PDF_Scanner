@echo off
setlocal EnableExtensions

set "ROOT=%~dp0"
set "PY=py"
where %PY% >nul 2>nul || set "PY=python"

echo [SETUP] Creating local virtual environment (.venv)...
"%PY%" -m venv "%ROOT%.venv"
if errorlevel 1 (
  echo [ERROR] Failed to create venv. Ensure Python 3 is installed.
  exit /b 1
)

set "VPY=%ROOT%.venv\Scripts\python.exe"
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

echo.
echo [READY] Local environment set up.
echo   - Python venv: "%ROOT%.venv"
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
