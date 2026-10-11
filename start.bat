@echo off
rem Opens the program in your browser (it runs only on this computer).
cd /d "%~dp0"
if not exist ".venv\Scripts\pythonw.exe" (
  echo The program is not installed yet. Running install.bat first ...
  call install.bat
)
start "" ".venv\Scripts\pythonw.exe" "%~dp0reserve_web.pyw"
