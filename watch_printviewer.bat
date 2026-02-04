@echo off
REM Watch: start printviewer if not running (for Task Scheduler, every 1 min)
cd /d "%~dp0"
set LOG=%~dp0watch_printviewer.log

tasklist /FI "IMAGENAME eq python.exe" 2>nul | findstr /I "python.exe" >nul
if errorlevel 1 (
  echo [%date% %time%] Starting printviewer... >> "%LOG%"
  echo [watch] Starting printviewer...
  call "%~dp0start_printviewer.bat"
  echo [%date% %time%] Called start_printviewer.bat >> "%LOG%"
  echo [watch] Started.
) else (
  echo [%date% %time%] printviewer already running. >> "%LOG%"
  echo [watch] printviewer already running.
)
