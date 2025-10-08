$ErrorActionPreference = 'Stop'

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Dist = Join-Path $Root 'dist'
New-Item -ItemType Directory -Force -Path $Dist | Out-Null

$ts = Get-Date -Format yyyyMMdd_HHmmss
$Dest = Join-Path $Dist ("PDF_Scanner_" + $ts)
New-Item -ItemType Directory -Force -Path $Dest | Out-Null

# Copy whole repo excluding dev/data folders and workspace files
$excludeDirs = @('Lib','Product_Data_File','user_inputs','dist','.git','.venv','__pycache__','.pytest_cache','.mypy_cache')
$excludeFiles = @('*.pyc','Thumbs.db','desktop.ini','*.code-workspace','dev.create.bat')

& robocopy $Root $Dest /E /COPY:DAT /XD $excludeDirs /XF $excludeFiles /NFL /NDL /NJH /NJS /NC /NS /NP | Out-Null
if ($LASTEXITCODE -ge 8) { throw "robocopy failed with rc=$LASTEXITCODE" }

# Plain text export with scaffold and file contents
$txt = Join-Path $Dest 'PLAIN_TEXT_EXPORT.txt'
'PDF_Scanner Plain Text Export' | Out-File -FilePath $txt -Encoding utf8
("Timestamp: " + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')) | Out-File -Append -FilePath $txt -Encoding utf8
'' | Out-File -Append $txt -Encoding utf8
'Scaffold (create on target):' | Out-File -Append $txt -Encoding utf8
'  user_inputs\EIDP_Import_Docs' | Out-File -Append $txt -Encoding utf8
'  user_inputs\Scanned_Docs' | Out-File -Append $txt -Encoding utf8
'  user_inputs\terms.csv   (copy from your source)' | Out-File -Append $txt -Encoding utf8
'  Product_Data_File\run_data  (created on first run)' | Out-File -Append $txt -Encoding utf8
'' | Out-File -Append $txt -Encoding utf8

Push-Location $Dest
$files = Get-ChildItem -Recurse -File -Include *.py,*.bat,*.md
foreach ($f in $files) {
  $sep = 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'
  $rel = (Resolve-Path -Relative $f.FullName)
  $sep | Out-File -Append $txt -Encoding utf8
  ('FILE START: ' + $rel) | Out-File -Append $txt -Encoding utf8
  $sep | Out-File -Append $txt -Encoding utf8
  '' | Out-File -Append $txt -Encoding utf8
  Get-Content -Raw $f.FullName | Out-File -Append $txt -Encoding utf8
  '' | Out-File -Append $txt -Encoding utf8
  $sep | Out-File -Append $txt -Encoding utf8
  ('FILE END:   ' + $rel) | Out-File -Append $txt -Encoding utf8
  $sep | Out-File -Append $txt -Encoding utf8
  '' | Out-File -Append $txt -Encoding utf8
}
Pop-Location

# Zip package (best effort)
try { Compress-Archive -Path ($Dest+'\*') -DestinationPath ($Dest+'.zip') -Force -ErrorAction Stop } catch {}

Start-Process explorer $Dest
