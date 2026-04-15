@echo off
setlocal

cd /d "%~dp0"

echo.
echo Building one-click Windows installer...
echo -------------------------------------
echo.

echo [1/7] Checking Python...
where py >nul 2>nul
if %errorlevel%==0 (
  set "PY_CMD=py -3"
) else (
  where python >nul 2>nul
  if %errorlevel% neq 0 (
    echo Python was not found on this build PC.
    echo Install Python 3 from: https://www.python.org/downloads/windows/
    echo Enable "Add Python to PATH" during install.
    goto :end
  )
  set "PY_CMD=python"
)

echo [2/7] Locating Tesseract on build PC...
set "TESS_DIR="
if exist "%ProgramFiles%\Tesseract-OCR\tesseract.exe" set "TESS_DIR=%ProgramFiles%\Tesseract-OCR"
if not defined TESS_DIR if exist "%ProgramFiles(x86)%\Tesseract-OCR\tesseract.exe" set "TESS_DIR=%ProgramFiles(x86)%\Tesseract-OCR"
if not defined TESS_DIR (
  echo Could not auto-find Tesseract.
  echo Enter full Tesseract folder path ^(example: C:\Program Files\Tesseract-OCR^):
  set /p TESS_DIR=
)
if not exist "%TESS_DIR%\tesseract.exe" (
  echo Invalid Tesseract path: %TESS_DIR%
  goto :end
)

echo [3/7] Installing/updating build tools...
%PY_CMD% -m pip install --upgrade pip pyinstaller pillow numpy pandas pytesseract opencv-python
if %errorlevel% neq 0 (
  echo Failed to install build dependencies.
  goto :end
)

echo [4/7] Creating app icon...
%PY_CMD% create_brand_icon.py
if %errorlevel% neq 0 (
  echo Icon generation failed.
  goto :end
)

echo [5/8] Building GUI app...
%PY_CMD% -m PyInstaller --noconfirm --clean --onedir --windowed --icon assets\detailextract.ico --add-data "assets\detailextract.ico;." --add-data "assets\duck_logo.jpeg;assets" --name DetailExtractOCRApp offline_ocr_gui_windows.py
if %errorlevel% neq 0 (
  echo PyInstaller build failed.
  goto :end
)

echo [6/8] Bundling Tesseract and app folders...
if exist "dist\DetailExtractOCRApp\tesseract" rmdir /s /q "dist\DetailExtractOCRApp\tesseract"
xcopy "%TESS_DIR%" "dist\DetailExtractOCRApp\tesseract\" /e /i /y >nul
if %errorlevel% neq 0 (
  echo Failed to copy Tesseract.
  goto :end
)
if not exist "dist\DetailExtractOCRApp\input-images" mkdir "dist\DetailExtractOCRApp\input-images"
if not exist "dist\DetailExtractOCRApp\output" mkdir "dist\DetailExtractOCRApp\output"

echo [7/8] Checking Inno Setup compiler...
set "ISCC_EXE="
if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set "ISCC_EXE=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not defined ISCC_EXE if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" set "ISCC_EXE=%ProgramFiles%\Inno Setup 6\ISCC.exe"
if not defined ISCC_EXE (
  where iscc >nul 2>nul
  if %errorlevel%==0 set "ISCC_EXE=iscc"
)

if not defined ISCC_EXE (
  echo Inno Setup was not found.
  echo Install from: https://jrsoftware.org/isdl.php
  echo Then run this build file again.
  echo App folder is still ready at: %cd%\dist\DetailExtractOCRApp
  goto :end
)

echo [8/8] Building Setup.exe installer...
"%ISCC_EXE%" windows_installer.iss
if %errorlevel% neq 0 (
  echo Installer build failed.
  goto :end
)

echo.
echo Success.
echo Share this installer file:
echo %cd%\dist\DetailExtractOCR_Installer.exe
echo.
echo User flow on target PC:
echo 1. Double-click DetailExtractOCR_Installer.exe
echo 2. Click Next, Install, Finish
echo 3. Open "DetailExtract OCR" from Desktop or Start Menu
echo 4. Click Select Files, then Extract Details

:end
echo.
echo Press any key to close this window.
pause >nul
