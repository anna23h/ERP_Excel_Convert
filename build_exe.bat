@echo off
setlocal
REM ============================================
REM Build VO tool into a single-file Windows exe.
REM Keep this .bat ASCII-only (Chinese breaks on GBK consoles).
REM ============================================

REM run from this script's own folder
cd /d "%~dp0"

REM --- find a working Python launcher ---
set "PY="
py -3 --version >nul 2>&1 && set "PY=py -3"
if not defined PY ( python --version >nul 2>&1 && set "PY=python" )
if not defined PY ( python3 --version >nul 2>&1 && set "PY=python3" )

if not defined PY (
  echo.
  echo Python 3 not found.
  echo 1^) Install from https://www.python.org/downloads/
  echo 2^) During install, CHECK "Add Python to PATH"
  echo 3^) If "python" opens the Microsoft Store, turn OFF the Store alias:
  echo    Settings ^> Apps ^> Advanced app settings ^> App execution aliases
  echo    -^> turn off "python.exe" and "python3.exe"
  goto end
)
echo Using Python: %PY%
%PY% --version

echo.
echo [1/3] Upgrading pip...
%PY% -m pip install --upgrade pip

echo.
echo [2/3] Installing dependencies...
%PY% -m pip install -r requirements.txt
if errorlevel 1 goto fail

echo.
echo [3/3] Building exe...
%PY% -m PyInstaller --onefile --windowed --name VOTool --collect-all openpyxl gui.py
if errorlevel 1 goto fail

echo.
echo Done. exe is at: dist\VOTool.exe
echo You may rename VOTool.exe to a Chinese name in Explorer if you like.
goto end

:fail
echo.
echo Build failed. Read the error lines ABOVE this message.
echo Common fixes:
echo  - no internet / proxy: pip cannot download packages
echo  - rebuild WITHOUT --windowed to see GUI startup errors
echo  - if it says missing pandas submodules: add --collect-all pandas

:end
echo.
pause
endlocal
