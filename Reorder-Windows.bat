@echo off
setlocal enabledelayedexpansion
REM ============================================
REM Reorder Helper - Windows double-click launcher (self-updating)
REM Double-click to: 1) pull latest code  2) install deps on first run  3) launch GUI
REM Update = git fetch + reset --hard origin/main: only tracked code is overwritten;
REM   your outputs (output/), raw_data/, .venv/ are untracked/ignored and kept as-is.
REM Requires: Python 3 (check "Add Python to PATH") and Git for Windows.
REM   Missing Git or offline -> skip update, keep running the current version.
REM NOTE: all text here is plain ASCII on purpose, so this file runs correctly
REM   under ANY Windows locale / code page. Do not add non-ASCII characters.
REM ============================================

REM Go to this script's own folder
cd /d "%~dp0"

REM --- Auto-update to latest (public repo, anonymous fetch, no login needed) ---
set "OLD="
set "NEW="
where git >nul 2>&1
if !errorlevel! equ 0 (
  if exist ".git" (
    echo Checking for updates...
    for /f "delims=" %%i in ('git rev-parse HEAD 2^>nul') do set "OLD=%%i"
    git fetch --quiet origin 2>nul
    if !errorlevel! equ 0 (
      git reset --hard origin/main 2>nul
      if !errorlevel! neq 0 ( echo Update not applied, using current version. )
    ) else (
      echo Online update failed ^(maybe offline^), using current version.
    )
    for /f "delims=" %%i in ('git rev-parse HEAD 2^>nul') do set "NEW=%%i"
    if defined NEW if not "!OLD!"=="!NEW!" ( echo Updated to the latest version. )
  ) else (
    echo Note: this folder is not a git checkout, skipping auto-update.
  )
) else (
  echo Note: Git not found, skipping auto-update. Install Git for Windows to enable it.
)

REM --- Find a usable Python launcher ---
set "PY="
py -3 --version >nul 2>&1 && set "PY=py -3"
if not defined PY ( python --version >nul 2>&1 && set "PY=python" )
if not defined PY ( python3 --version >nul 2>&1 && set "PY=python3" )

if not defined PY (
  echo.
  echo Python 3 not found. Install it from: https://www.python.org/downloads/
  echo During install, be sure to check "Add Python to PATH".
  echo If typing "python" opens the Microsoft Store: Settings ^> Apps ^> Advanced app settings ^> App execution aliases
  echo   -^> turn OFF "python.exe" and "python3.exe".
  echo.
  pause
  goto :eof
)

set "VENV=.venv"
set "VPY=%VENV%\Scripts\python.exe"

REM First run needs an environment; after an update deps may change, reinstall once
set "NEED_DEP="
if not exist "%VPY%" (
  echo.
  echo First run: creating the environment ^(about 1-2 minutes, only this once^)...
  %PY% -m venv "%VENV%" || ( echo Failed to create virtual environment & pause & goto :eof )
  set "NEED_DEP=1"
)
if defined OLD if defined NEW if not "!OLD!"=="!NEW!" set "NEED_DEP=1"

if defined NEED_DEP (
  echo Installing/updating dependencies...
  "%VPY%" -m pip install --upgrade pip >nul 2>&1
  "%VPY%" -m pip install -r requirements.txt || ( echo Failed to install dependencies ^(check network/proxy^) & pause & goto :eof )
)

echo Starting Reorder Helper... ^(closing this window will close the program^)
"%VPY%" reorder_gui.py

REM Keep the window open on abnormal exit so the error stays visible
if errorlevel 1 (
  echo.
  echo The program exited abnormally. Please send the error text above to the developer.
  pause
)
endlocal
