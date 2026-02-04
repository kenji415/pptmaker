@echo off
chcp 932 >nul
setlocal enabledelayedexpansion

echo ============================================================
echo PPT File Batch PDF Converter
echo ============================================================
echo.

REM Current directory
echo Current directory: %CD%
echo.

REM Python script path
set SCRIPT_PATH=%~dp0ppt_to_pdf_batch.py
echo Python script path: %SCRIPT_PATH%
echo.

REM Check if Python script exists
if not exist "%SCRIPT_PATH%" (
    echo ERROR: Python script not found: %SCRIPT_PATH%
    echo.
    pause
    exit /b 1
)

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Please install Python.
    echo.
    pause
    exit /b 1
)

echo Python found.
echo.

REM Get target path from argument or input
if "%~1"=="" (
    set /p TARGET_PATH="Enter folder path containing PPT files: "
) else (
    set TARGET_PATH=%~1
)

REM Remove quotes from path
set TARGET_PATH=!TARGET_PATH:"=!

REM Validate path
if "!TARGET_PATH!"=="" (
    echo ERROR: Path not specified.
    echo.
    pause
    exit /b 1
)

echo.
echo Starting process...
echo Target folder: !TARGET_PATH!
echo.

REM Execute Python script
python "%SCRIPT_PATH%" "!TARGET_PATH!"

set PYTHON_EXIT_CODE=%ERRORLEVEL%

if %PYTHON_EXIT_CODE% neq 0 (
    echo.
    echo ============================================================
    echo Error occurred. Exit code: %PYTHON_EXIT_CODE%
    echo ============================================================
    echo.
    pause
    exit /b %PYTHON_EXIT_CODE%
)

echo.
echo ============================================================
echo Process completed.
echo ============================================================
echo.
pause
