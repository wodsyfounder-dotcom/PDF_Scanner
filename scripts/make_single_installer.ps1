$ErrorActionPreference = 'Stop'

# Generates a single self-extracting installer BAT that recreates the
# minimal app layout (Application, run.bat, install.bat, scripts, user_inputs scaffold).
# Usage: pwsh -File scripts/make_single_installer.ps1

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path | Split-Path -Parent
$OutDir = Join-Path $Root 'dist'
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

$stamp = Get-Date -Format yyyyMMdd_HHmmss
$PayloadZip = Join-Path $OutDir ("pdf_scanner_payload_" + $stamp + ".zip")
$InstallerBat = Join-Path $OutDir ("PDF_Scanner_installer_" + $stamp + ".bat")

# Build a temporary staging folder
$Stage = Join-Path $OutDir ("stage_" + $stamp)
New-Item -ItemType Directory -Force -Path $Stage | Out-Null

# Files to include (lean runtime only)
$include = @(
  'Application/eidp_term_scanner.py',
  'Application/eidp_term_scanner.core.py',
  'run.bat',
  'install.bat',
  'scripts/generate_terms_schema.py',
  'README_EIDP_Term_Scanner.md'
)

foreach ($rel in $include) {
  $src = Join-Path $Root $rel
  if (!(Test-Path $src)) { throw "Missing required file: $rel" }
  $dest = Join-Path $Stage $rel
  New-Item -ItemType Directory -Force -Path (Split-Path -Parent $dest) | Out-Null
  Copy-Item $src $dest -Force
}

# Ensure expected folders are created at install time
New-Item -ItemType Directory -Force -Path (Join-Path $Stage 'user_inputs') | Out-Null

# Create zip payload
if (Test-Path $PayloadZip) { Remove-Item -Force $PayloadZip }
Compress-Archive -Path (Join-Path $Stage '*') -DestinationPath $PayloadZip -Force

# Base64 encode the zip
$b64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($PayloadZip))

$bat = @'
@echo off
setlocal EnableExtensions

set "SELF=%~f0"
set "ROOT=%cd%"
set "TMP_ZIP=%TEMP%\pdf_scanner_payload.zip"
set "TMP_B64=%TEMP%\pdf_scanner_payload.b64"

> "%TMP_B64%" (
echo PAYLOAD_BEGIN
'@
$bat += ($b64 -replace '(.{120})', "$1`r`n")
$bat += @'
echo PAYLOAD_END
)

for /f "usebackq tokens=1,* delims=:" %%A in ("%TMP_B64%") do (
  if "%%A"=="PAYLOAD_BEGIN" (
    break
  )
)

certutil -f -decode "%TMP_B64%" "%TMP_ZIP%" >nul 2>nul
if errorlevel 1 (
  echo [ERROR] certutil failed to decode embedded payload.>&2
  del /q "%TMP_B64%" 2>nul
  exit /b 1
)
del /q "%TMP_B64%" 2>nul

powershell -NoProfile -Command "Expand-Archive -Path '%TMP_ZIP%' -DestinationPath '%ROOT%' -Force"
if errorlevel 1 (
  echo [ERROR] Failed to extract payload.>&2
  del /q "%TMP_ZIP%" 2>nul
  exit /b 1
)
del /q "%TMP_ZIP%" 2>nul

echo [INFO] Files extracted. Run install.bat next (or set VENV_DIR in user_inputs\scanner.env and run run.bat).
endlocal & exit /b 0
'@

[IO.File]::WriteAllText($InstallerBat, $bat, [Text.UTF8Encoding]::new($false))
Write-Host "Created installer: $InstallerBat"

