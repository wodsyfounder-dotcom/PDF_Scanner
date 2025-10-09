@echo off
setlocal EnableExtensions

powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0dev.create.ps1"

endlocal & exit /b %ERRORLEVEL%
