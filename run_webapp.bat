@echo off
rem ============================================================
rem  Trade Flow Report Generator - Web App launcher (Windows)
rem  Opens the web app at http://127.0.0.1:8000
rem  Dependencies are installed automatically on first run.
rem ============================================================
cd /d "%~dp0"

call install_deps.bat
if errorlevel 1 (
    pause
    exit /b 1
)

echo Starting the Trade Flow Report Generator web app...
echo Open http://127.0.0.1:8000 in your browser.
echo Press Ctrl+C to stop.
%PY% -m uvicorn webapp.main:app --host 127.0.0.1 --port 8000
