@echo off
rem ============================================================
rem  Table 1-6 Builder - Windows launcher
rem  Builds the six "Table N" workbooks (plus Figure 1 Trade
rem  Balance, derived from Tables 5 and 6) from the raw ITC
rem  source files in sourcefiles\, ready for generate_report.py.
rem  Dependencies are installed automatically on first run.
rem ============================================================
cd /d "%~dp0"

call install_deps.bat
if errorlevel 1 exit /b 1

set EXCEL_DIR=sourcefiles
set OUT_DIR=output\tables

echo.
echo Building Table 1-6 workbooks from raw ITC files...
%PY% make_tables.py --excel-dir "%EXCEL_DIR%" --out-dir "%OUT_DIR%"
if errorlevel 1 (
    echo.
    echo ERROR: Table generation failed. See the message above.
    pause
    exit /b 1
)

echo.
echo Done. The Figure 1 balance file is derived from Tables 5 and 6.
echo Run generate_report.py with --excel-dir %OUT_DIR% next.
pause
