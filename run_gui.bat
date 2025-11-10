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

rem Load optional scanner config (user_inputs\scanner.env) as KEY=VALUE lines
set "CFG=%ROOT%user_inputs\scanner.env"
if exist "%CFG%" (
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
)

rem Optional venv override via scanner.env
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

if not "%QUIET%"=="1" (
  echo [RUN] Python: "%PY%"
)

"%PY%" "%ROOT%ui_next\qt_main.py" %*
set "RC=%ERRORLEVEL%"
endlocal & exit /b %RC%
