@echo off
chcp 65001 >nul
setlocal
rem One-time setup: creates a private Python environment in this folder (.venv) and installs the libraries.
rem Nothing is installed system-wide. Run it again any time to repair or update.
rem   install.bat          asks whether to add the optional OCR add-on
rem   install.bat ocr      installs it without asking
rem   install.bat noocr    skips it without asking
cd /d "%~dp0"

set "PY="
where py >nul 2>&1 && set "PY=py -3"
if not defined PY (where python >nul 2>&1 && set "PY=python")
if not defined PY (
  echo Python was not found. Install Python 3.11 or newer from https://www.python.org/downloads/
  echo and tick "Add python.exe to PATH" in the installer. Then run install.bat again.
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo Creating the private environment .venv ...
  %PY% -m venv .venv || (echo Could not create .venv & pause & exit /b 1)
)
echo Installing the libraries (pymupdf, pdfplumber, openpyxl) ...
".venv\Scripts\python.exe" -m pip install --upgrade pip >nul
".venv\Scripts\python.exe" -m pip install -r requirements.txt || (echo Installation failed - check the internet connection. & pause & exit /b 1)

set "OCR=%~1"
if /i not "%OCR%"=="ocr" if /i not "%OCR%"=="noocr" (
  echo.
  echo Optional: the OCR add-on reads a reserve-service certificate or NI letter that is only a picture.
  set /p "OCR=Install it? (about 100 MB) [y/N]: "
)
if /i "%OCR%"=="y" set "OCR=ocr"
if /i "%OCR%"=="ocr" (
  ".venv\Scripts\python.exe" -m pip install -r requirements-ocr.txt || echo The OCR add-on could not be installed - everything else works.
)

echo.
echo Done. Double-click start.bat to open the program.
if /i not "%~1"=="ocr" if /i not "%~1"=="noocr" pause
endlocal
