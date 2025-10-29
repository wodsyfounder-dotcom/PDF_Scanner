@echo off
setlocal EnableExtensions EnableDelayedExpansion

set "ROOT=%~dp0"
set "PY=py"
where %PY% >nul 2>nul || set "PY=python"

rem Usage: init-venv.bat [venv_dir] [--full] [--write-env]
set "VENV_DIR=%~1"
set "ARG1=%~2"
set "ARG2=%~3"
if "%VENV_DIR%"=="" set "VENV_DIR=%ROOT%.venv"

echo [INIT] Bootstrapping venv at "%VENV_DIR%" ...
"%PY%" "%ROOT%scripts\init_venv.py" --dir "%VENV_DIR%" %ARG1% %ARG2%
set "RC=%ERRORLEVEL%"
if not "%RC%"=="0" (
  echo [ERROR] venv setup failed (rc=%RC%).
  endlocal & exit /b %RC%
)

echo [DONE] venv initialized at "%VENV_DIR%".
echo        To use it: set VENV_DIR in user_inputs\scanner.env or run via run.bat
endlocal & exit /b 0

