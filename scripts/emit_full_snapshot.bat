@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul

rem Emit a full text snapshot of key project files into a repo-local .txt file.

set "ROOT=%~dp0..\"
set "OUT=%ROOT%full_snapshot.txt"
if not "%~1"=="" set "OUT=%~1"
if exist "%OUT%" del /q "%OUT%" 2>nul

>> "%OUT%" echo ===== PDF_Scanner Full Snapshot =====
>> "%OUT%" echo Generated: %DATE% %TIME%
>> "%OUT%" echo Root: %ROOT%
>> "%OUT%" echo.

call :PRINT_FILE "README_EIDP_Term_Scanner.md"
call :PRINT_FILE "install.bat"
call :PRINT_FILE "run.bat"
rem New Qt-based UI
call :PRINT_FILE "ui_next\qt_main.py"
call :PRINT_FILE "ui_next\backend.py"
call :PRINT_FILE "ui_next\requirements-ui.txt"
call :PRINT_FILE "Application\eidp_term_scanner.py"
call :PRINT_FILE "Application\eidp_term_scanner.core.py"
call :PRINT_FILE "user_inputs\scanner.env"
call :PRINT_FILE "scripts\ocr_page_to_excel.py"

for %%F in ("scripts\*.py") do (
  if /I not "%%~nxF"=="ocr_page_to_excel.py" (
    call :PRINT_FILE "%%~fF"
  )
)

>> "%OUT%" echo ===== END OF SNAPSHOT =====
echo Wrote snapshot: "%OUT%"
goto :EOF

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
