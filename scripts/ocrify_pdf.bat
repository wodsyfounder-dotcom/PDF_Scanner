@echo off
setlocal enabledelayedexpansion

rem -----------------------------------------------------------------------------
rem Convert a text PDF into an image-only (OCR-required) PDF using portable
rem ImageMagick bundled in the repository virtual environment.
rem Usage: scripts\ocrify_pdf.bat <input.pdf> [output.pdf] [dpi] [mode] [quality]
rem   mode: lossless (default, PNG), tiff, or jpeg
rem   quality: JPEG quality (1-100), ignored for lossless/tiff
rem -----------------------------------------------------------------------------

if "%~1"=="" (
    echo Usage: %~nx0 ^<input.pdf^> [output.pdf] [dpi] [mode] [quality]
    exit /b 1
)

set "INPUT=%~f1"
if not exist "%INPUT%" (
    echo Input file not found: %INPUT%
    exit /b 1
)

set "OUTPUT=%~2"
if "%OUTPUT%"=="" (
    set "OUTPUT=%~dpn1_scanned.pdf"
) else (
    set "OUTPUT=%~f2"
)

set "DPI=%~3"
if "%DPI%"=="" set "DPI=300"

set "MODE=%~4"
if "%MODE%"=="" set "MODE=lossless"

set "QUALITY=%~5"
if "%QUALITY%"=="" set "QUALITY=95"

for %%I in ("%MODE%") do set "MODE=%%~I"

if /I "%MODE%"=="lossless" (
    set "RASTER_EXT=png"
    set "RASTER_FLAGS=-units PixelsPerInch -density %DPI% -alpha off -define png:compression-level=0 -define png:compression-strategy=0 -define png:exclude-chunk=all"
) else if /I "%MODE%"=="tiff" (
    set "RASTER_EXT=tif"
    set "RASTER_FLAGS=-units PixelsPerInch -density %DPI% -alpha off -compress none -define tiff:rows-per-strip=0"
) else if /I "%MODE%"=="jpeg" (
    set "RASTER_EXT=jpg"
    set "RASTER_FLAGS=-units PixelsPerInch -density %DPI% -alpha off -quality %QUALITY% -sampling-factor 4:4:4 -colorspace RGB"
) else (
    echo Unknown mode "%MODE%". Use lossless, tiff, or jpeg.
    exit /b 1
)

rem Resolve repo root (script lives in /scripts)
set "SCRIPT_DIR=%~dp0"
for %%I in ("%SCRIPT_DIR%..") do set "REPO_ROOT=%%~fI"

set "MAGICK=%REPO_ROOT%\.venv\Scripts\magick.cmd"
if not exist "%MAGICK%" (
    echo Unable to locate ImageMagick shim at %MAGICK%
    echo Activate the venv or reinstall ImageMagick before running this script.
    exit /b 1
)

set "TMPDIR=%REPO_ROOT%\.tmp_ocrify"
if exist "%TMPDIR%" rd /s /q "%TMPDIR%"
md "%TMPDIR%"

set "PAGE_PATTERN=%TMPDIR%\page-%%%%02d.%RASTER_EXT%"

echo [1/3] Rasterizing "%INPUT%" at %DPI% dpi (%MODE% mode)...
call "%MAGICK%" "%INPUT%" !RASTER_FLAGS! "%PAGE_PATTERN%"
if errorlevel 1 (
    echo Image rasterization failed.
    rd /s /q "%TMPDIR%"
    exit /b 1
)

echo [2/3] Rebuilding image-only PDF -> "%OUTPUT%"...
call "%MAGICK%" "%TMPDIR%\page-*.%RASTER_EXT%" -compress Zip "%OUTPUT%"
if errorlevel 1 (
    echo PDF assembly failed.
    rd /s /q "%TMPDIR%"
    exit /b 1
)

echo [3/3] Cleaning up temp files...
rd /s /q "%TMPDIR%"

echo Done. OCR-only PDF saved to:
echo   %OUTPUT%

endlocal
