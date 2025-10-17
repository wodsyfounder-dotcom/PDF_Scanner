Param()
$ErrorActionPreference = 'Stop'

# Build a single plain-text BAT that reconstructs the project files (no zip/certutil),
# then runs the project installer to create a venv and install dependencies.

$Root   = Split-Path -Parent $PSScriptRoot
$OutDir = Join-Path $Root 'dist'
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$OutBat = Join-Path $OutDir 'full_installer.bat'

# Files to embed (text/batch/python only). install.bat is renamed inside payload to install_project.bat
$include = @(
  'Application/eidp_term_scanner.py',
  'Application/eidp_term_scanner.core.py',
  'gui.py',
  'run.bat',
  'install.bat',
  'scripts/generate_terms_schema.py',
  'scripts/compile_master.py',
  'scripts/ocr_dump.py',
  'README_EIDP_Term_Scanner.md'
)

function Get-PayloadEntry($rel) {
  $src = Join-Path $Root $rel
  if (-not (Test-Path $src)) { throw "Missing file: $rel" }
  $bytes = [IO.File]::ReadAllBytes($src)
  $b64   = [Convert]::ToBase64String($bytes)
  $outRel = if ($rel -eq 'install.bat') { 'install_project.bat' } else { $rel }
  [PSCustomObject]@{ Path = $outRel; B64 = $b64 }
}

$entries = @()
foreach ($rel in $include) { $entries += (Get-PayloadEntry $rel) }

# Compose the embedded PowerShell payload that writes files and runs the installer
$ps = @'
param([Parameter(ValueFromRemainingArguments=
$true)][string[]]
$Args)
$ErrorActionPreference = 'Stop'

$Root = (Get-Location).Path
function Write-Base64File([string]
$Rel, [string]
$B64) {
  $Dest = Join-Path $Root $Rel
  New-Item -ItemType Directory -Force -Path (Split-Path -Parent $Dest) | Out-Null
  $clean = ($B64 -replace '\s','')
  [IO.File]::WriteAllBytes($Dest, [Convert]::FromBase64String($clean))
}

# Write embedded files
'@

foreach ($e in $entries) {
  # Wrap base64 in a here-string and emit a call per file
  $b64Wrapped = ($e.B64 -replace '(.{120})', "$1`n")
  $ps += "`r`nWrite-Base64File '$($e.Path)' @'`r`n$b64Wrapped`r`n'@`r`n"
}

$ps += @'
# Ensure scaffold directories exist
$dirs = @(
  'user_inputs',
  'user_inputs\EIDP_Import_Docs',
  'user_inputs\Scanned_Docs',
  'Product_Data_File',
  'Product_Data_File\run_data'
)
foreach ($d in $dirs) { New-Item -ItemType Directory -Force -Path (Join-Path $Root $d) | Out-Null }

# Run project installer (creates venv, installs deps, scaffolds defaults)
$installer = Join-Path $Root 'install_project.bat'
if (Test-Path $installer) {
  $psi = New-Object System.Diagnostics.ProcessStartInfo
  $psi.FileName = $installer
  # Quote arguments to preserve spaces when calling the BAT
  $quoted = @()
  foreach ($a in $Args) {
    if ($a -match '\s') { $quoted += '"' + ($a -replace '"','""') + '"' } else { $quoted += $a }
  }
  $psi.Arguments = [string]::Join(' ', $quoted)
  $psi.WorkingDirectory = $Root
  $psi.UseShellExecute = $false
  $psi.RedirectStandardOutput = $false
  $psi.RedirectStandardError = $false
  $p = [System.Diagnostics.Process]::Start($psi)
  $p.WaitForExit()
  exit $p.ExitCode
} else {
  Write-Error "Missing installer: $installer"
  exit 1
}
'@

# Compose the BAT wrapper that extracts the PS1 payload from within itself
$bat = @"
@echo off
setlocal EnableExtensions

set "SELF=%~f0"
set "ROOT=%cd%"
set "PAYLOAD=%TEMP%\pdf_scanner_payload_%RANDOM%.ps1"

REM Locate markers and extract embedded PowerShell to a temp file
for /f "tokens=1 delims=:" %%A in ('findstr /n /r "^:PAYLOAD_PS1" "%~f0"') do set START=%%A
for /f "tokens=1 delims=:" %%A in ('findstr /n /r "^:END_PAYLOAD_PS1" "%~f0"') do set END=%%A
if "%START%"=="" (
  echo [ERROR] Payload start marker not found.>&2
  endlocal & exit /b 1
)
if "%END%"=="" (
  echo [ERROR] Payload end marker not found.>&2
  endlocal & exit /b 1
)
set /a SKIP=%START%
set /a COUNT=%END%-%START%-1
if %COUNT% LSS 1 (
  echo [ERROR] Empty payload region.>&2
  endlocal & exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-Content -Path '%~f0' | Select-Object -Skip %SKIP% -First %COUNT% | Set-Content -Path '%PAYLOAD%' -Encoding UTF8" >nul 2>nul
if errorlevel 1 (
  echo [ERROR] Failed to extract PowerShell payload.>&2
  endlocal & exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%PAYLOAD%" %*
set "RC=%ERRORLEVEL%"
del /q "%PAYLOAD%" >nul 2>nul
endlocal & exit /b %RC%

:PAYLOAD_PS1
"@

$batEnd = @"
:END_PAYLOAD_PS1
"@

# Write final BAT with embedded PS1 (ensure newlines around payload markers)
[IO.File]::WriteAllText($OutBat, $bat + "`r`n" + $ps + "`r`n" + $batEnd, [Text.UTF8Encoding]::new($false))
Write-Host "Created: $OutBat"
