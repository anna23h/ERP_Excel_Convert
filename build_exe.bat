@echo off
REM ============================================
REM Build VO tool into a single-file Windows exe.
REM Prerequisite: Python 3 installed (python.org, check "Add Python to PATH").
REM Usage: copy the whole project to a Windows PC, then double-click this file.
REM NOTE: keep this .bat ASCII-only (Chinese in .bat breaks on GBK consoles).
REM ============================================

echo [1/2] Installing dependencies...
python -m pip install -r requirements.txt
if errorlevel 1 goto fail

echo [2/2] Building exe...
pyinstaller --onefile --windowed --name VOTool --collect-all openpyxl gui.py
if errorlevel 1 goto fail

echo.
echo Done. exe is at: dist\VOTool.exe
echo You may rename VOTool.exe to a Chinese name in Explorer if you like.
pause
exit /b 0

:fail
echo.
echo Build failed. Tips:
echo  - rebuild WITHOUT --windowed to see the real error in the console
echo  - if it complains about missing pandas modules, append: --collect-all pandas
pause
exit /b 1
