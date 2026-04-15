@echo off
setlocal

cd /d "%~dp0"

echo.
echo Building shareable Windows OCR software (no install needed on target PC)
echo ------------------------------------------------------------------------
echo.

echo [1/6] Checking Python...
where py >nul 2>nul
if %errorlevel%==0 (
  set "PY_CMD=py -3"
) else (
  where python >nul 2>nul
  if %errorlevel% neq 0 (
    echo Python was not found on this build PC.
    echo Install Python 3 from: https://www.python.org/downloads/windows/
    echo Make sure "Add Python to PATH" is enabled.
    goto :end
  )
  set "PY_CMD=python"
)

echo [2/6] Locating Tesseract installation on build PC...
set "TESS_DIR="
if exist "%ProgramFiles%\Tesseract-OCR\tesseract.exe" set "TESS_DIR=%ProgramFiles%\Tesseract-OCR"
if not defined TESS_DIR if exist "%ProgramFiles(x86)%\Tesseract-OCR\tesseract.exe" set "TESS_DIR=%ProgramFiles(x86)%\Tesseract-OCR"
if not defined TESS_DIR (
  echo Could not auto-find Tesseract.
  echo Enter full Tesseract folder path (example: C:\Program Files\Tesseract-OCR):
  set /p TESS_DIR=
)

if not exist "%TESS_DIR%\tesseract.exe" (
  echo Invalid Tesseract path: %TESS_DIR%
  goto :end
)

if not exist "%TESS_DIR%\tessdata\eng.traineddata" (
  echo Warning: eng.traineddata was not found in "%TESS_DIR%\tessdata".
  echo OCR may fail if language data is missing.
)

echo [3/6] Installing/updating build tools...
%PY_CMD% -m pip install --upgrade pip pyinstaller
if %errorlevel% neq 0 (
  echo Failed to install PyInstaller.
  goto :end
)

echo [4/6] Building executable...
%PY_CMD% -m PyInstaller --noconfirm --clean --onedir --name OfflineOCR offline_batch_ocr_windows.py
if %errorlevel% neq 0 (
  echo Build failed.
  goto :end
)

echo [5/6] Bundling Tesseract inside the app folder...
if exist "dist\OfflineOCR\tesseract" rmdir /s /q "dist\OfflineOCR\tesseract"
xcopy "%TESS_DIR%" "dist\OfflineOCR\tesseract\" /e /i /y >nul
if %errorlevel% neq 0 (
  echo Failed to copy Tesseract into dist folder.
  goto :end
)

if not exist "dist\OfflineOCR\input-images" mkdir "dist\OfflineOCR\input-images"
if not exist "dist\OfflineOCR\output" mkdir "dist\OfflineOCR\output"

(
  echo @echo off
  echo cd /d "%%~dp0"
  echo OfflineOCR.exe
  echo echo.
  echo echo Press any key to close this window.
  echo pause ^>nul
) > "dist\OfflineOCR\Run_Offline_OCR.bat"

echo [6/6] Creating shareable zip...
powershell -NoProfile -Command "Compress-Archive -Path 'dist\\OfflineOCR\\*' -DestinationPath 'dist\\OfflineOCR_Windows_NoInstall.zip' -Force" >nul
if %errorlevel% neq 0 (
  echo Zip creation failed, but app folder is ready at:
  echo %cd%\dist\OfflineOCR
  goto :end
)

echo.
echo Build complete.
echo Share this zip file:
echo %cd%\dist\OfflineOCR_Windows_NoInstall.zip
echo.
echo On target PC: unzip and run Run_Offline_OCR.bat

:end
echo.
echo Press any key to close this window.
pause >nul
