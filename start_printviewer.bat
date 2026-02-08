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
  echo.
  echo ERROR: Python not found.
  echo Install Python or run setup_venv.bat first.
  echo See %LOG% for details.
  echo.
  pause
  exit /b 1
)
if not exist "%VENVPY%" (
  echo.
  echo ERROR: venv not found. Run setup_venv.bat first in this folder.
  echo.
  pause
  exit /b 1
)
echo [%date% %time%] Using: %PYEXE% >> "%LOG%"

REM Ensure deps (quick if already installed)
"%PIPCMD%" install -q -r requirements.txt >> "%LOG%" 2>&1

REM If yaml missing, install again (visible) then recheck
"%PYEXE%" -c "import yaml" 2>nul
if errorlevel 1 (
  echo Installing required packages...
  "%PIPCMD%" install -r requirements.txt
  "%PYEXE%" -c "import yaml" 2>nul
  if errorlevel 1 (
    echo.
    echo ERROR: Could not install yaml. Run in this folder:
    echo   venv\Scripts\pip install -r requirements.txt
    echo.
    pause
    exit /b 1
  )
)

if exist "C:\tools\poppler-25.12.0\Library\bin" (
  set POPPLER_PATH=C:\tools\poppler-25.12.0\Library\bin
)

REM This PC: PDF folder is on X:
set PDF_DIR=X:\算数科作業フォルダ\QS\test_pdf

start /B "" "%PYEXE%" "%~dp0app.py"
echo [%date% %time%] Started app.py >> "%LOG%"
echo.
echo Running on http://127.0.0.1:5000
echo.
timeout /t 3 /nobreak >nul
start http://127.0.0.1:5000
echo Browser opened. You can close this window; app keeps running.
echo.
if /I not "%~1"=="nopause" pause