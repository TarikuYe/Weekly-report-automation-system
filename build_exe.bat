@echo off
REM ============================================================
REM  Build script — turns gui.py into a standalone Windows app
REM  Run this ONCE. Afterwards, dist\WeeklyReportAutomator.exe
REM  is fully double-clickable and needs no terminal or Python
REM  installed on the machine you copy it to.
REM ============================================================

echo Installing/updating required packages...
pip install -r requirements.txt

echo.
echo Building WeeklyReportAutomator.exe ...
pyinstaller --noconfirm --onefile --windowed ^
    --name "WeeklyReportAutomator" ^
    --icon "icon.ico" ^
    --add-data "icon.ico;." ^
    gui.py

echo.
echo ============================================================
echo   Build complete.
echo   Your app is at:  dist\WeeklyReportAutomator.exe
echo   Copy that one .exe anywhere and double-click to run it.
echo ============================================================
pause