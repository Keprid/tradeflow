@echo off
rem ============================================================
rem  Competitor Analysis - Windows launcher (standalone)
rem
rem  Runs AFTER the goods trade flow report: once the first
rem  flow tells you which products matter, download the
rem  per-product "imports from world by exporter" workbooks
rem  from ITC Trade Map into a folder, then point this script
rem  at that folder. No other Trade Flow step is required.
rem
rem  Usage:
rem    run_competitor.bat [data_folder] [out_folder]
rem
rem  Defaults: data_folder = TEKTAT, out_folder = same folder.
rem ============================================================
cd /d "%~dp0"

call install_deps.bat
if errorlevel 1 exit /b 1

set EXCEL_DIR=TEKTAT
if not "%~1"=="" set EXCEL_DIR=%~1
set OUT_DIR=%EXCEL_DIR%
if not "%~2"=="" set OUT_DIR=%~2

echo.
echo Building competitor analysis from per-product downloads in %EXCEL_DIR%...
%PY% generate_competitor_analysis.py --data "%EXCEL_DIR%" --out "%OUT_DIR%"
if errorlevel 1 (
    echo.
    echo ERROR: Competitor analysis failed. See the message above.
    pause
    exit /b 1
)

echo.
echo Done. Deliverables saved to: %OUT_DIR%
echo   COMPETITOR ANALYSIS.png   (paste-ready table picture)
echo   COMPETITOR ANALYSIS.docx  (editable Word table)
echo   COMPETITOR ANALYSIS.xlsx  (editable Excel workbook)
pause
