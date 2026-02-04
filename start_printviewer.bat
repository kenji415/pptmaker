@echo off
REM Start printviewer (called by watch or manually)
cd /d "%~dp0"
set LOG=%~dp0start_printviewer.log

set PYEXE=
set PIPCMD=
set VENVPY=%~dp0venv\Scripts\python.exe
if exist "%VENVPY%" (
  set PYEXE=%VENVPY%
  set PIPCMD=%~dp0venv\Scripts\pip.exe
)
if not defined PYEXE (
  REM No venv: try PATH then common locations (Task Scheduler often has no PATH)
  where python >nul 2>&1
  if not errorlevel 1 (
    set PYEXE=python
    set PIPCMD=python -m pip
  ) else (
    if exist "%USERPROFILE%\anaconda3\python.exe" ( set PYEXE=%USERPROFILE%\anaconda3\python.exe & set PIPCMD=%USERPROFILE%\anaconda3\Scripts\pip.exe )
    if not defined PYEXE if exist "%USERPROFILE%\miniconda3\python.exe" ( set PYEXE=%USERPROFILE%\miniconda3\python.exe & set PIPCMD=%USERPROFILE%\miniconda3\Scripts\pip.exe )
    if not defined PYEXE if exist "C:\Python312\python.exe" ( set PYEXE=C:\Python312\python.exe & set PIPCMD=C:\Python312\Scripts\pip.exe )
    if not defined PYEXE if exist "C:\Python311\python.exe" ( set PYEXE=C:\Python311\python.exe & set PIPCMD=C:\Python311\Scripts\pip.exe )
  )
)
if not defined PYEXE (
  echo [%date% %time%] ERROR: Python not found. >> "%LOG%"
  echo ERROR: Python not found. Install Python or create venv in this folder. >> "%LOG%"
  exit /b 1
)
echo [%date% %time%] Using: %PYEXE% >> "%LOG%"

REM Ensure deps (quick if already installed)
"%PIPCMD%" install -q -r requirements.txt >> "%LOG%" 2>&1

if exist "C:\tools\poppler-25.12.0\Library\bin" (
  set POPPLER_PATH=C:\tools\poppler-25.12.0\Library\bin
)

start /B "%PYEXE%" app.py
echo [%date% %time%] Started app.py >> "%LOG%"
