@echo off
setlocal

cd /d "%~dp0"

echo.
echo [1/4] Checking Python...
where py >nul 2>nul
if %errorlevel%==0 (
  set "PY_CMD=py -3"
) else (
  where python >nul 2>nul
  if %errorlevel% neq 0 (
    echo Python was not found.
    echo Install Python 3 from: https://www.python.org/downloads/windows/
    echo Make sure "Add Python to PATH" is enabled during installation.
    goto :end
  )
  set "PY_CMD=python"
)

echo [2/4] Installing/updating build tool (PyInstaller)...
%PY_CMD% -m pip install --upgrade pip pyinstaller
if %errorlevel% neq 0 (
  echo Failed to install PyInstaller.
  goto :end
)

echo [3/4] Building Windows app...
%PY_CMD% -m PyInstaller --noconfirm --clean --onefile --name OfflineOCR offline_batch_ocr_windows.py
if %errorlevel% neq 0 (
  echo Build failed.
  goto :end
)

echo [4/4] Build completed.
echo EXE path: %cd%\dist\OfflineOCR.exe
echo.
echo You can copy OfflineOCR.exe to another Windows PC.
echo Tesseract still needs to be installed on that PC.

:end
echo.
echo Press any key to close this window.
pause >nul
