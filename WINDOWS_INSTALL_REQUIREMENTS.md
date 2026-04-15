# Windows Install Requirements

This file lists what is needed to build and use `DetailExtract OCR` on Windows.

## For the Windows Build PC

You need these installed before building the installer:

1. Python 3
- Download from: `https://www.python.org/downloads/windows/`
- During install, enable `Add Python to PATH`

2. Tesseract OCR
- Install in Windows
- Typical path:
  `C:\Program Files\Tesseract-OCR\tesseract.exe`

3. Inno Setup 6
- Download from: `https://jrsoftware.org/isdl.php`

## Python Packages Used During Build

The build script installs these automatically:

- `pyinstaller`
- `pillow`
- `numpy`
- `pandas`
- `pytesseract`
- `opencv-python`

## Build File

Run this on the Windows build PC:

`build_windows_installer_one_click.bat`

This creates:

`dist\DetailExtractOCR_Installer.exe`

## For the End User Windows PC

The user only needs:

1. `DetailExtractOCR_Installer.exe`
2. Double-click installer
3. Click `Next` -> `Install` -> `Finish`
4. Open `DetailExtract OCR`
5. Select images
6. Click `Extract Details`

The target Windows PC does not need separate manual installation of:

- Python
- Tesseract

because they are bundled into the installed software package.

## Output Files

The software writes output in:

`%LocalAppData%\DetailExtractOCR\output`

Main files:

- `combined_ocr.txt`
- `combined_customers.txt`
- `extraction_quality_summary.txt`
- `last_run.log`

## Supported Image Types

- `.png`
- `.jpg`
- `.jpeg`
- `.tif`
- `.tiff`
- `.bmp`
- `.webp`

## Important Notes

- OCR quality depends on image clarity
- Crowded or overlapped text may still reduce accuracy
- The latest build hides the black Tesseract console window during extraction
- The latest build uses strict 40-sequence output validation
