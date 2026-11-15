@echo off
setlocal enabledelayedexpansion
REM ============================================
REM B2B available stock (odoo_api/b2b_stock.py) - Windows entry + daily schedule
REM ISSUES [B2B-available] A, 2026-11-15. Odoo is READ-ONLY for this script.
REM
REM   B2B-Stock-Windows.bat            run once now (double-click), window stays open
REM   B2B-Stock-Windows.bat install    register a Task Scheduler job: Mon-Fri 09:00
REM   B2B-Stock-Windows.bat uninstall  remove that job
REM   B2B-Stock-Windows.bat run        what the scheduled job calls (no pause, logs to file)
REM
REM Output: output\YYYYMMDD\B2B...-YYYYMMDD.xlsx, or B2B_STOCK_OUTDIR in config.py.
REM Log:    output\odoo_api.log (appended every run).
REM Needs:  Python 3 on PATH; config.py with ODOO (read-only API key) and
REM         ODOO_EXCLUDE_VENDORS; this PC must reach the Odoo server.
REM 09:00 = after the 08:00 Tmall batch confirmation, so Tmall reservations include today.
REM The job uses StartWhenAvailable: if the PC was off/asleep at 09:00 it runs once
REM   as soon as it is back (same reason the Mac side used launchd, not cron).
REM Code update: "git pull --ff-only" (never reset): local edits are never thrown away;
REM   if the pull fails the current version runs.
REM NOTE: plain ASCII only and CRLF line endings (see .gitattributes). Do not add
REM   non-ASCII characters: cmd reads this file in the OEM code page.
REM ============================================

cd /d "%~dp0"
set "TASK=IHTCT B2B Stock"

if /i "%~1"=="install" goto :install
if /i "%~1"=="uninstall" goto :uninstall

set "MODE=%~1"
if not exist "output" mkdir "output"
set "LOG=output\odoo_api.log"

REM --- Find a usable Python launcher ---
set "PY="
py -3 --version >nul 2>&1 && set "PY=py -3"
if not defined PY ( python --version >nul 2>&1 && set "PY=python" )
if not defined PY ( python3 --version >nul 2>&1 && set "PY=python3" )
if not defined PY (
  echo Python 3 not found. Install it from https://www.python.org/downloads/ and check "Add Python to PATH".
  if /i not "%MODE%"=="run" pause
  exit /b 1
)

REM --- Update code (non-destructive) ---
where git >nul 2>&1
if !errorlevel! equ 0 if exist ".git" git pull --ff-only --quiet origin main >nul 2>&1

REM --- Virtual environment: create on first run, rebuild if broken ---
set "VPY=.venv\Scripts\python.exe"
if exist "%VPY%" "%VPY%" -c "import sys" >nul 2>&1
if exist "%VPY%" if errorlevel 1 rmdir /s /q ".venv"
if not exist "%VPY%" (
  echo First run: creating the environment, about 1-2 minutes...
  %PY% -m venv .venv || ( echo Failed to create virtual environment & exit /b 1 )
  "%VPY%" -m pip install --upgrade pip >nul 2>&1
  "%VPY%" -m pip install -r requirements.txt
)

if /i "%MODE%"=="run" (
  >>"%LOG%" echo ===== !DATE! !TIME! B2B stock start =====
  "%VPY%" odoo_api\b2b_stock.py >>"%LOG%" 2>&1
  set "RC=!errorlevel!"
  >>"%LOG%" echo ===== !DATE! !TIME! B2B stock end, exit code !RC! =====
  exit /b !RC!
)

"%VPY%" odoo_api\b2b_stock.py
if errorlevel 1 echo. & echo The script exited abnormally. Please send the error text above to the developer.
pause
exit /b 0

:install
powershell -NoProfile -ExecutionPolicy Bypass -Command "$t = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At 9:00; $a = New-ScheduledTaskAction -Execute '%~f0' -Argument 'run' -WorkingDirectory '%~dp0'; $s = New-ScheduledTaskSettingsSet -StartWhenAvailable; Register-ScheduledTask -TaskName '%TASK%' -Trigger $t -Action $a -Settings $s -Force | Out-Null"
if errorlevel 1 (
  echo Could not register the scheduled task. Run this from a normal user account and try again.
  pause
  exit /b 1
)
echo Registered "%TASK%": Monday-Friday 09:00, runs %~f0 run
echo Test it now:  schtasks /run /tn "%TASK%"   then check output\odoo_api.log
pause
exit /b 0

:uninstall
schtasks /delete /tn "%TASK%" /f
pause
exit /b 0
