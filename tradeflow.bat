@echo off
rem ============================================================
rem  Trade Flow Report Generator - one-word launcher (Windows)
rem
rem  Typing  tradeflow  (with the Trade Flow folder on PATH) or
rem  double-clicking this file starts the web app and opens your
rem  browser. Dependencies are installed automatically on first
rem  run. Stop the app with Ctrl+C.
rem ============================================================
cd /d "%~dp0"

call install_deps.bat
if errorlevel 1 (
    pause
    exit /b 1
)

echo Starting the Trade Flow Report Generator web app...
ping -n 3 127.0.0.1 >nul
start "" http://127.0.0.1:8000
echo If your browser did not open, go to http://127.0.0.1:8000
echo Press Ctrl+C to stop.

:run
%PY% -m uvicorn webapp.main:app --host 127.0.0.1 --port 8000
if "%errorlevel%"=="3221225786" exit /b 0
if errorlevel 1 (
    echo Server stopped unexpectedly - restarting in 3 seconds...
    ping -n 4 127.0.0.1 >nul
    goto run
)
