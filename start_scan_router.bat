@echo off
cd /d "%~dp0"

set PYTHON_PATH=C:\Users\doctor\anaconda3\python.exe

if exist "%PYTHON_PATH%" (
    echo Using Anaconda Python: %PYTHON_PATH%
    "%PYTHON_PATH%" -c "import fitz; print('fitz module OK')" 2>nul
    if errorlevel 1 (
        echo Installing pymupdf...
        "%PYTHON_PATH%" -m pip install pymupdf
    )
    "%PYTHON_PATH%" scan_router.py
) else (
    echo Using system Python
    python scan_router.py
)
pause
