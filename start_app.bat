@echo off
REM app.py を起動（venv の Python を使う）
cd /d "%~dp0"
set POPPLER_PATH=C:\tools\poppler-25.12.0\Library\bin
set PDF_DIR=X:\算数科作業フォルダ\QS\test_pdf
if exist "%~dp0venv\Scripts\python.exe" (
  "%~dp0venv\Scripts\python.exe" app.py
) else (
  python app.py
)
pause

















