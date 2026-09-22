@echo off
REM ============================================================
REM  Build script for Excel File Comparator (by Asrar Ahmed Junedi)
REM  Run this on a WINDOWS machine that has Python 3.8+ installed.
REM ============================================================

echo Installing dependencies...
python -m pip install --upgrade pip
python -m pip install openpyxl pyinstaller

echo.
echo Building ExcelComparator.exe ...
pyinstaller --onefile --windowed --name "ExcelComparator" --icon=app_icon.ico excel_compare.py

echo.
echo ============================================================
echo Done! Your exe is at:  dist\ExcelComparator.exe
echo Next: run installer.iss with Inno Setup to create a proper
echo Windows installer that adds a Desktop + Start Menu shortcut.
echo ============================================================
pause
