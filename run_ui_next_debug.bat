@echo off
setlocal EnableExtensions EnableDelayedExpansion

set "ROOT=%~dp0"
set "APP_ROOT=%ROOT%EIDAT_App_Files\"
set "LOG_FILE=%ROOT%debug_output_new.txt"

rem Prefer local venv if present (activate to ensure env vars / PATH are set)
set "VENV_PY=%APP_ROOT%.venv\Scripts\python.exe"
set "VENV_ACTIVATE=%APP_ROOT%.venv\Scripts\activate.bat"
if exist "%VENV_PY%" (
  if exist "%VENV_ACTIVATE%" call "%VENV_ACTIVATE%"
  set "PY=%VENV_PY%"
  set "PATH=%APP_ROOT%.venv\Scripts;!PATH!"
) else (
  set "PY=py"
  where %PY% >nul 2>nul || set "PY=python"
)

rem Make vendored packages and the repo available to Python
set "PYTHONPATH=%APP_ROOT%;%APP_ROOT%Lib\site-packages;%PYTHONPATH%"

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
)

rem Optional venv override via scanner.env (activate to ensure env vars / PATH are set)
if defined VENV_DIR (
  set "VENV_PY=%VENV_DIR%\Scripts\python.exe"
  set "VENV_ACTIVATE=%VENV_DIR%\Scripts\activate.bat"
  if not exist "%VENV_PY%" (
    if not "%QUIET%"=="1" echo [SETUP] No venv at "%VENV_DIR%". Bootstrapping...
    call "%ROOT%install.bat" "%VENV_DIR%"
  )
  if exist "%VENV_PY%" (
    if exist "%VENV_ACTIVATE%" call "%VENV_ACTIVATE%"
    set "PY=%VENV_PY%"
    set "PATH=%VENV_DIR%\Scripts;!PATH!"
  ) else (
    echo [ERROR] Failed to create venv at "%VENV_DIR%".>&2
    exit /b 1
  )
)

rem Force debug mode for UI Next runs
set "DEBUG_MODE=1"

if not "%QUIET%"=="1" (
  echo [RUN] Python: "%PY%"
  echo [RUN] Debug : DEBUG_MODE=1 ^(ui_next^)
  echo [RUN] Log   : "%LOG_FILE%"
)

set "PSHELL=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%PSHELL%" (
  set "PS_CMD=& { & '%PY%' -m ui_next.qt_main @args 2>&1 ^| Tee-Object -FilePath '%LOG_FILE%'; exit $LASTEXITCODE }"
  "%PSHELL%" -NoLogo -NoProfile -Command "!PS_CMD!" -- %*
) else (
  echo [WARN] PowerShell not found; running without live tee ^(log will overwrite "%LOG_FILE%"^).
  "%PY%" -m ui_next.qt_main %* > "%LOG_FILE%" 2>&1
)

set "RC=%ERRORLEVEL%"
endlocal & exit /b %RC%
