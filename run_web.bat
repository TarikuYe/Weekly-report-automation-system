@echo off
title Weekly Report Dashboard

echo.
echo  ============================================
echo   Weekly Report Automator -- Web Dashboard
echo  ============================================
echo.

REM Install / upgrade core dependencies (pandas 3.x required for Python 3.14)
pip install -q "pandas>=3.0" "openpyxl>=3.1.5" "flask>=3.0"

echo  Starting server...
echo  Open your browser at:  http://127.0.0.1:5000
echo.
echo  Press Ctrl+C to stop the server.
echo.

python "%~dp0app.py"

pause
