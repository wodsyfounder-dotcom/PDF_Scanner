@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul

rem Emit a full text snapshot of key project files into a repo-local .txt file.

set "APP_ROOT=%~dp0..\"
set "ROOT=%APP_ROOT%..\"
set "OUT=%ROOT%full_snapshot.txt"
if not "%~1"=="" set "OUT=%~1"
if exist "%OUT%" del /q "%OUT%" 2>nul
rem Keep the dependency manifest in one place for release audits.
set "CORE_PY_PACKAGES=pymupdf pandas openpyxl matplotlib opencv-python-headless PySide6"
set "OCR_PY_PACKAGES=torch torchvision easyocr"

>> "%OUT%" echo ===== PDF_Scanner Full Snapshot =====
>> "%OUT%" echo Generated: %DATE% %TIME%
>> "%OUT%" echo Root: %ROOT%
>> "%OUT%" echo.
>> "%OUT%" echo Required Python packages (ship list):
call :PRINT_PACKAGE_SECTION "Core runtime" "%CORE_PY_PACKAGES%"
call :PRINT_PACKAGE_SECTION "OCR fallback" "%OCR_PY_PACKAGES%"
>> "%OUT%" echo.

call :PRINT_FILE "%ROOT%README_EIDP_Term_Scanner.md"
call :PRINT_FILE "%ROOT%install.bat"
call :PRINT_FILE "%ROOT%run.bat"
call :PRINT_FILE "%ROOT%run_gui.bat"
rem New Qt-based UI
call :PRINT_FILE "%APP_ROOT%ui_next\qt_main.py"
call :PRINT_FILE "%APP_ROOT%ui_next\backend.py"
call :PRINT_FILE "%APP_ROOT%ui_next\requirements-ui.txt"
call :PRINT_FILE "%APP_ROOT%Application\eidp_term_scanner.py"
call :PRINT_FILE "%APP_ROOT%Application\eidp_term_scanner.core.py"
call :PRINT_FILE "%ROOT%user_inputs\scanner.env"
call :PRINT_FILE "%APP_ROOT%scripts\ocr_page_to_excel.py"

for %%F in ("%APP_ROOT%scripts\*.py") do (
  if /I not "%%~nxF"=="ocr_page_to_excel.py" (
    call :PRINT_FILE "%%~fF"
  )
)

>> "%OUT%" echo ===== END OF SNAPSHOT =====
echo Wrote snapshot: "%OUT%"
goto :EOF

:PRINT_PACKAGE_SECTION
set "_PKG_LABEL=%~1"
set "_PKG_ITEMS=%~2"
if not defined _PKG_ITEMS exit /b 0
>> "%OUT%" echo   %_PKG_LABEL%:
for %%P in (%_PKG_ITEMS%) do (
  >> "%OUT%" echo     - %%P
)
exit /b 0

:PRINT_FILE
set "_FILE=%~1"
set "_PATH=%_FILE%"
if not exist "%_PATH%" (
  rem Resolve relative to project root
  set "_PATH=%ROOT%%_FILE%"
)
>> "%OUT%" echo.
>> "%OUT%" echo ---------- BEGIN FILE: %_FILE% ----------
if exist "%_PATH%" (
  type "%_PATH%" >> "%OUT%"
) else (
  >> "%OUT%" echo [MISSING] %_FILE%
)
>> "%OUT%" echo ---------- END FILE: %_FILE% ----------
exit /b 0
