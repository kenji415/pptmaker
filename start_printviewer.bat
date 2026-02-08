@echo off
cd /d "C:\Users\doctor\printviewer"

if not exist "C:\Users\doctor\printviewer\venv\Scripts\python.exe" (
  echo venv not found. Run setup_venv.bat first.
  pause
  exit /b 1
)

"C:\Users\doctor\printviewer\venv\Scripts\pip.exe" install -q -r requirements.txt >> "C:\Users\doctor\printviewer\start_printviewer.log" 2>&1
"C:\Users\doctor\printviewer\venv\Scripts\python.exe" -c "import yaml" 2>nul
if errorlevel 1 (
  "C:\Users\doctor\printviewer\venv\Scripts\pip.exe" install -r requirements.txt
)

if exist "C:\tools\poppler-25.12.0\Library\bin" (
  set "POPPLER_PATH=C:\tools\poppler-25.12.0\Library\bin"
)
set "PDF_DIR=X:\算数科作業フォルダ\QS\test_pdf"

start /B "pv" "C:\Users\doctor\printviewer\venv\Scripts\python.exe" "C:\Users\doctor\printviewer\app.py"
echo Running on http://127.0.0.1:5000
timeout /t 3 /nobreak >nul
start http://127.0.0.1:5000
echo Done. You can close this window.
if not "%~1"=="nopause" pause
