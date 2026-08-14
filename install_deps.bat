@echo off
rem ============================================================
rem  Trade Flow generator - dependency bootstrap (Windows)
rem
rem  Called by the launcher .bat files. Checks that Python is
rem  available and that the packages in requirements.txt are
rem  installed, installing them automatically on first run.
rem  Safe to run repeatedly: it skips the install when nothing
rem  is missing.
rem ============================================================

rem ---- locate a Python interpreter --------------------------------
where python >nul 2>nul
if not errorlevel 1 (
    set "PY=python"
    goto :found
)
where py >nul 2>nul
if not errorlevel 1 (
    set "PY=py -3"
    goto :found
)
echo.
echo ERROR: Python was not found on this computer.
echo Install Python from https://www.python.org/downloads/ and during
echo installation tick "Add python.exe to PATH", then run the launcher
echo again.
pause
exit /b 1

:found

rem ---- install any missing packages -------------------------------
%PY% -c "import openpyxl, docx, matplotlib, fastapi, uvicorn" >nul 2>nul
if not errorlevel 1 (
    echo Dependencies already installed.
    exit /b 0
)

echo.
echo Installing Python dependencies (first run only, please wait)...
%PY% -m pip install -r "%~dp0requirements.txt"
if errorlevel 1 (
    echo.
    echo ERROR: Could not install the required Python packages.
    echo If this is a locked-down work computer you may need to run
    echo the install through your company's software centre instead.
    pause
    exit /b 1
)
echo Dependencies installed.
echo.
exit /b 0
