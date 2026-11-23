@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo.
echo ================================================
echo   PDFMint Studio - Optional Tool Installer
echo ================================================
echo.
echo This installs the desktop helpers used for:
echo   - OCR PDF          : Tesseract OCR
echo   - PDF to PDF/A     : Ghostscript
echo   - Legacy Office    : LibreOffice (.doc/.xls/.ppt/.odt/.ods/.odp)
echo.
echo DOCX, XLSX and PPTX already have built-in PDF fallbacks.
echo.

where winget >nul 2>&1
if errorlevel 1 (
  echo ERROR: Windows Package Manager ^(winget^) was not found.
  echo Install/update "App Installer" from Microsoft Store, then run this file again.
  pause
  exit /b 1
)

echo [1/3] Installing Tesseract OCR...
winget install -e --id UB-Mannheim.TesseractOCR --accept-source-agreements --accept-package-agreements

echo.
echo [2/3] Installing Ghostscript...
winget install -e --id ArtifexSoftware.GhostScript --accept-source-agreements --accept-package-agreements

echo.
echo [3/3] Installing LibreOffice...
winget install -e --id TheDocumentFoundation.LibreOffice --accept-source-agreements --accept-package-agreements

echo.
echo ================================================
echo Optional tools installation finished.
echo Close PDFMint's CMD window and start start_windows.bat again.
echo ================================================
echo.
pause
