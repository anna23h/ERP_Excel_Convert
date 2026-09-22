@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

rem 开工前拉取 . Windows 双击入口
rem
rem 与 开工前拉取.command 是同一份逻辑：这个 .bat 只负责找到 Git Bash 并把它转起来，
rem 拉取、报错、结束时的暂停全在那个脚本里，两边永远不会走岔。
rem 命名沿用仓库既有惯例：Mac双击运行.command / Windows双击运行.bat 成对。
rem 只读：只 pull，不 push、不 reset、不 checkout。加 -q 跳过结束时的暂停。
rem
rem 写这个文件踩到的两个坑，改之前先看：
rem 1. chcp 65001 必须是第一条执行的语句。否则 cmd 按 OEM 代码页读文件内容，
rem    "开工前拉取.command" 这个中文字面量会变成乱码，if exist 永远判不到。
rem 2. 不许用多行的 for %%x in ( ... ) 括号块。cmd 逐字节记录文件偏移，
rem    多字节内容 + 跨行括号块会让偏移错乱，后面整片语句被切碎乱执行
rem    ——实测报过一串 "xxx is not recognized"，看着像文件坏了。
rem    所以下面的 Git Bash 探测写成一行一条 if，别图紧凑改回括号块。

set "BASH="
if exist "%ProgramFiles%\Git\bin\bash.exe" set "BASH=%ProgramFiles%\Git\bin\bash.exe"
if not defined BASH if exist "%ProgramFiles(x86)%\Git\bin\bash.exe" set "BASH=%ProgramFiles(x86)%\Git\bin\bash.exe"
if not defined BASH if exist "%LOCALAPPDATA%\Programs\Git\bin\bash.exe" set "BASH=%LOCALAPPDATA%\Programs\Git\bin\bash.exe"
if not defined BASH for /f "delims=" %%B in ('where bash 2^>nul') do set "BASH=%%B"

if not defined BASH goto :nobash

"%BASH%" "开工前拉取.command" %*
exit /b %ERRORLEVEL%

:nobash
echo.
echo   X 找不到 Git Bash，没法跑拉取脚本。
echo.
echo     装一个 Git for Windows 即可: https://git-scm.com/download/win
echo     装在非默认位置的话，把 bash.exe 所在目录加进 PATH 也行。
echo.
pause
exit /b 1
