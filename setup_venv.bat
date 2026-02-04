@echo off
REM One-time: create venv and install requirements (ASCII only)
cd /d "%~dp0"

if not exist "venv\Scripts\python.exe" (
  echo Creating venv...
  python -m venv venv
)

echo Installing requirements...
venv\Scripts\pip install -r requirements.txt

echo Done. You can run start_printviewer.bat or watch_printviewer_test.bat now.
pause
