@echo off
title MLB Daily Betting Report
echo =======================================
echo  MLB Daily Betting Report Generator
echo =======================================
echo.

REM Try python first, then python3
where python >nul 2>&1
if %errorlevel% == 0 (
    set PYTHON=python
) else (
    where python3 >nul 2>&1
    if %errorlevel% == 0 (
        set PYTHON=python3
    ) else (
        echo ERROR: Python is not installed.
        echo Please install Python from https://www.python.org/downloads/
        echo Make sure to check "Add Python to PATH" during installation.
        pause
        exit /b 1
    )
)

REM Install required packages if missing
echo Checking required packages...
%PYTHON% -m pip install pandas openpyxl xlsxwriter requests -q

echo.
echo Running daily report...
echo.

REM Run the script from the same folder as this .bat file
cd /d "%~dp0"
%PYTHON% daily_report.py

echo.
if %errorlevel% == 0 (
    echo Report generated successfully!
) else (
    echo Something went wrong. Please check the error above.
)

echo.
pause
