@echo off
REM Keep this file ASCII-only. Do not add Japanese (cmd uses CP932; Cursor saves UTF-8 = broken).
setlocal enabledelayedexpansion

echo ============================================================
echo PPT to PDF (batch)
echo ============================================================
echo.

echo Current dir: %CD%
echo.

set SCRIPT_PATH=%~dp0ppt_to_pdf_batch.py
echo Script: %SCRIPT_PATH%
echo.

if not exist "%SCRIPT_PATH%" (
    echo ERROR: Script not found: %SCRIPT_PATH%
    echo.
    pause
    exit /b 1
)

python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Please install Python.
    echo.
    pause
    exit /b 1
)

echo Python OK.
echo.

if "%~1"=="" (
    set /p TARGET_PATH="Enter folder path containing PPT files: "
) else (
    set TARGET_PATH=%~1
)

set TARGET_PATH=!TARGET_PATH:"=!

if "!TARGET_PATH!"=="" (
    echo ERROR: No path specified.
    echo.
    pause
    exit /b 1
)

echo.
echo Starting...
echo Folder: !TARGET_PATH!
echo.

python "%SCRIPT_PATH%" "!TARGET_PATH!"

set PYTHON_EXIT_CODE=%ERRORLEVEL%

if %PYTHON_EXIT_CODE% neq 0 (
    echo.
    echo ============================================================
    echo Error. Exit code: %PYTHON_EXIT_CODE%
    echo ============================================================
    echo.
    pause
    exit /b %PYTHON_EXIT_CODE%
)

echo.
echo ============================================================
echo Done.
echo ============================================================
echo.
pause
