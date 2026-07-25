@echo off
:: ============================================================
:: setup_scheduler.bat
:: Registers Windows Task Scheduler jobs for the Weekly Report
:: Email Automator — auto-detected Python and project paths.
::
:: Run once as Administrator:
::   Right-click this file → "Run as administrator"
::
:: Jobs created:
::   WeeklyReport_EmailCheck   every N minutes (Mon-Fri, active window)
::   WeeklyReport_DeadlineRun  once at deadline time each week (forced)
:: ============================================================

setlocal EnableDelayedExpansion

echo.
echo ============================================================
echo  Weekly Report Automator — Task Scheduler Setup
echo ============================================================
echo.

:: ── Verify running as Administrator ──────────────────────────
net session >nul 2>&1
if errorlevel 1 (
    echo [ERROR] This script must be run as Administrator.
    echo         Right-click the file and select "Run as administrator".
    echo.
    pause
    exit /b 1
)
echo [OK] Running as Administrator.

:: ── Paths ─────────────────────────────────────────────────────
set "PYTHON=C:\Python314\python.exe"
set "SCRIPT_DIR=C:\Users\user\Desktop\weekly-report-automator"
set "RUNNER=%SCRIPT_DIR%\auto_runner.py"
set "CONFIG=%SCRIPT_DIR%\config.json"
set "LOG_DIR=%SCRIPT_DIR%\logs"

:: Verify python exists
if not exist "%PYTHON%" (
    echo [ERROR] Python not found at: %PYTHON%
    echo         Edit this file and update the PYTHON variable.
    pause
    exit /b 1
)
echo [OK] Python : %PYTHON%

:: Verify runner script
if not exist "%RUNNER%" (
    echo [ERROR] auto_runner.py not found at: %RUNNER%
    pause
    exit /b 1
)
echo [OK] Runner : %RUNNER%

:: Create logs folder if missing
if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"

:: ── Read settings from config.json via PowerShell ─────────────
echo.
echo [INFO] Reading deadline settings from config.json...

for /f "usebackq delims=" %%a in (
    `powershell -NoProfile -Command "$c=(Get-Content '%CONFIG%' -Raw | ConvertFrom-Json); $c.email.deadline_day"`
) do set "DEADLINE_DAY=%%a"

for /f "usebackq delims=" %%a in (
    `powershell -NoProfile -Command "$c=(Get-Content '%CONFIG%' -Raw | ConvertFrom-Json); $c.email.deadline_hour"`
) do set "DEADLINE_HOUR=%%a"

for /f "usebackq delims=" %%a in (
    `powershell -NoProfile -Command "$c=(Get-Content '%CONFIG%' -Raw | ConvertFrom-Json); '{0:D2}' -f [int]$c.email.deadline_minute"`
) do set "DEADLINE_MIN=%%a"

for /f "usebackq delims=" %%a in (
    `powershell -NoProfile -Command "$c=(Get-Content '%CONFIG%' -Raw | ConvertFrom-Json); $c.auto_check.interval_minutes"`
) do set "INTERVAL_MIN=%%a"

:: Defaults if parsing failed
if not defined DEADLINE_DAY   set "DEADLINE_DAY=Thursday"
if not defined DEADLINE_HOUR  set "DEADLINE_HOUR=17"
if not defined DEADLINE_MIN   set "DEADLINE_MIN=00"
if not defined INTERVAL_MIN   set "INTERVAL_MIN=30"

:: Format deadline time as HH:MM with leading zero
set "DEADLINE_TIME=0%DEADLINE_HOUR%:%DEADLINE_MIN%"
if %DEADLINE_HOUR% GEQ 10 set "DEADLINE_TIME=%DEADLINE_HOUR%:%DEADLINE_MIN%"

:: Map day name → schtasks abbreviation
set "SCHED_DAY=THU"
if /i "%DEADLINE_DAY%"=="Monday"    set "SCHED_DAY=MON"
if /i "%DEADLINE_DAY%"=="Tuesday"   set "SCHED_DAY=TUE"
if /i "%DEADLINE_DAY%"=="Wednesday" set "SCHED_DAY=WED"
if /i "%DEADLINE_DAY%"=="Thursday"  set "SCHED_DAY=THU"
if /i "%DEADLINE_DAY%"=="Friday"    set "SCHED_DAY=FRI"
if /i "%DEADLINE_DAY%"=="Saturday"  set "SCHED_DAY=SAT"
if /i "%DEADLINE_DAY%"=="Sunday"    set "SCHED_DAY=SUN"

echo [INFO] Deadline  : %DEADLINE_DAY% at %DEADLINE_TIME%
echo [INFO] Interval  : every %INTERVAL_MIN% minutes
echo.

:: ── Remove old tasks ──────────────────────────────────────────
echo [INFO] Removing old tasks (if any)...
schtasks /delete /tn "WeeklyReport_EmailCheck"  /f >nul 2>&1
schtasks /delete /tn "WeeklyReport_DeadlineRun" /f >nul 2>&1

:: ── Task 1: recurring interval check ─────────────────────────
:: Runs every INTERVAL_MIN minutes, every day.
:: auto_runner.py itself respects the active window and deadline,
:: so no day/hour restriction is needed here — it exits cleanly
:: if conditions aren't met.
echo [INFO] Creating Task 1: WeeklyReport_EmailCheck (every %INTERVAL_MIN% min)

schtasks /create ^
    /tn "WeeklyReport_EmailCheck" ^
    /tr "\"%PYTHON%\" \"%RUNNER%\" --config \"%CONFIG%\"" ^
    /sc MINUTE /mo %INTERVAL_MIN% ^
    /ru "%USERNAME%" ^
    /rl HIGHEST ^
    /f

if errorlevel 1 (
    echo [WARN] Task 1 may not have registered cleanly. Check Task Scheduler.
) else (
    echo [OK]   Task 1 created.
)

:: ── Task 2: guaranteed run at the exact deadline ──────────────
:: Uses --force --ignore-window so it always fires even if the
:: interval task somehow missed it.
echo [INFO] Creating Task 2: WeeklyReport_DeadlineRun (%DEADLINE_DAY% %DEADLINE_TIME%)

schtasks /create ^
    /tn "WeeklyReport_DeadlineRun" ^
    /tr "\"%PYTHON%\" \"%RUNNER%\" --config \"%CONFIG%\" --force --ignore-window" ^
    /sc WEEKLY /d %SCHED_DAY% /st %DEADLINE_TIME% ^
    /ru "%USERNAME%" ^
    /rl HIGHEST ^
    /f

if errorlevel 1 (
    echo [WARN] Task 2 may not have registered cleanly. Check Task Scheduler.
) else (
    echo [OK]   Task 2 created.
)

:: ── Verify both tasks exist ───────────────────────────────────
echo.
echo [INFO] Verifying registered tasks...
schtasks /query /tn "WeeklyReport_EmailCheck"  /fo LIST | findstr "Task Name\|Status\|Next Run"
schtasks /query /tn "WeeklyReport_DeadlineRun" /fo LIST | findstr "Task Name\|Status\|Next Run"

:: ── Done ─────────────────────────────────────────────────────
echo.
echo ============================================================
echo  SETUP COMPLETE
echo.
echo  Two tasks are now registered in Windows Task Scheduler:
echo.
echo  1. WeeklyReport_EmailCheck
echo     Runs every %INTERVAL_MIN% minutes continuously.
echo     Respects your active window and deadline settings.
echo.
echo  2. WeeklyReport_DeadlineRun
echo     Runs at %DEADLINE_TIME% every %DEADLINE_DAY% with --force.
echo     Guarantees delivery even if the app is closed.
echo.
echo  To manage tasks:
echo    Open Task Scheduler (search in Start menu)
echo    Look under: Task Scheduler Library
echo.
echo  To remove tasks:
echo    schtasks /delete /tn "WeeklyReport_EmailCheck"  /f
echo    schtasks /delete /tn "WeeklyReport_DeadlineRun" /f
echo.
echo  Logs are written to:
echo    %LOG_DIR%\auto_runner.log
echo ============================================================
echo.
pause
endlocal
