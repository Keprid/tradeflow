@set UNCTAD_CLIENT_ID=0d24f378-7e2d-4b34-a962-e54c3f336c1c
@set UNCTAD_API_KEY=7zun3HsqSwt238GR9/niHZ7D6N8lnwEuzrAb8Xg3h6U=



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
