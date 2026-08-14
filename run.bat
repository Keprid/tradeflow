@echo off
rem ============================================================
rem  Trade Flow Report Generator - Windows launcher
rem  Double-click this file or run it from Command Prompt.
rem  Dependencies are installed automatically on first run.
rem ============================================================
cd /d "%~dp0"

call install_deps.bat
if errorlevel 1 exit /b 1

set EXCEL_DIR=sample_data
set CONFIG=config\saudi_arabia.json
set OUTPUT=output\KENYA-SAUDI ARABIA TRADE FLOW.docx

echo.
echo Generating report...
%PY% generate_report.py --excel-dir "%EXCEL_DIR%" --config "%CONFIG%" --output "%OUTPUT%"
if errorlevel 1 (
    echo.
    echo ERROR: Generation failed. See the message above.
    pause
    exit /b 1
)

echo.
echo Done. Report saved to: %OUTPUT%
pause
