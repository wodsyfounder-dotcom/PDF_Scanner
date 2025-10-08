$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# Resolve repo root
$ROOT = Split-Path -Parent $MyInvocation.MyCommand.Path

# Choose Python launcher (prefer local venv)
$venvPy = Join-Path $ROOT '.venv\Scripts\python.exe'
if (Test-Path $venvPy) {
  $py = $venvPy
} else {
  $py = 'py'
  if (-not (Get-Command $py -ErrorAction SilentlyContinue)) { $py = 'python' }
}

# Paths
$script = Join-Path $ROOT 'Application\\eidp_term_scanner.py'
$terms  = Join-Path $ROOT 'user_inputs\\terms.csv'
$inDir  = Join-Path $ROOT 'user_inputs\\EIDP_Import_Docs'
$scanned= Join-Path $ROOT 'user_inputs\\Scanned_Docs'

$outDir = Join-Path $ROOT 'Product_Data_File'
$outXlsx= Join-Path $outDir 'scan_results.xlsx'
$outJson= Join-Path $outDir 'scan_results.json'
$outCsv = Join-Path $outDir 'scan_results_flat.csv'

# Ensure directories exist
$null = New-Item -ItemType Directory -Force -Path $inDir, $scanned, $outDir, (Join-Path $outDir 'run_data')

# Extend PATH with local tools if present (no admin PATH changes)
$tess1 = Join-Path $ROOT 'tools\tesseract'
$tess2 = Join-Path $ROOT 'tools\tesseract\bin'
$popplerBin = Join-Path $ROOT 'tools\poppler\bin'
foreach ($d in @($tess1,$tess2,$popplerBin)) { if (Test-Path $d) { $env:PATH = "$d;" + $env:PATH } }

Write-Host "[RUN] Python: $py" -ForegroundColor Cyan
Write-Host "[RUN] Terms : $terms" -ForegroundColor Cyan
Write-Host "[RUN] PDFs  : $inDir"  -ForegroundColor Cyan
Write-Host "[RUN] Out   : $outDir (per-run outputs saved under run_data)" -ForegroundColor Cyan

# Run (pass through any extra args)
& $py $script `
  --input $terms `
  --pdf-folder $inDir `
  --output-xlsx $outXlsx `
  --output-json $outJson `
  --output-csv $outCsv `
  --scanned-folder $scanned `
  @args

if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

Write-Host "[DONE] Check: $outDir\\run_data for this run; top-level EIDP_data.csv was updated." -ForegroundColor Green
