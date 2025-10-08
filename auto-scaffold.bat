@echo off
setlocal EnableExtensions

set "ROOT=%~dp0"
set "PY=py"
where %PY% >nul 2>nul || set "PY=python"

if "%PY%"=="" (
  echo [ERROR] Python launcher not found. Please install Python 3 and ensure ^"py^" or ^"python^" is on PATH.
  exit /b 1
)

echo [SETUP] Creating standard folders and sample files...
"%PY%" "%ROOT%Application\setup_scaffold.py"
set "RC=%ERRORLEVEL%"
if not "%RC%"=="0" exit /b %RC%

echo [READY] Scaffold complete.
echo   - Input PDFs : "%ROOT%user_inputs\EIDP_Import_Docs"
echo   - Scanned out: "%ROOT%user_inputs\Scanned_Docs"
echo   - Terms file : "%ROOT%user_inputs\terms.csv"
echo   - Data folder: "%ROOT%Product_Data_File" (per-run outputs go under run_data)
echo.
echo Next: double-click run.bat or execute it from a terminal.

endlocal & exit /b 0

