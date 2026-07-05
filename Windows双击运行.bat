@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
REM ============================================
REM VO 拉单工具 · Windows 双击运行（免打包，自动更新到最新版）
REM 双击即：① 自动拉取最新代码 ② 首次自动装依赖 ③ 启动图形界面。
REM 更新用 git fetch + reset --hard origin/main：只覆盖代码，
REM   你的产出(输出/)、raw_data/、.venv/ 都是未跟踪/已忽略，原样保留。
REM 前提：本机已装 Python 3（勾 Add Python to PATH）与 Git for Windows。
REM   缺 Git 或断网 → 跳过更新、用当前版本继续运行，不阻断。
REM ============================================

REM 切到本脚本所在目录
cd /d "%~dp0"

REM --- 自动更新到最新版（公开仓库，匿名可拉，无需账号密码）---
set "OLD="
set "NEW="
where git >nul 2>&1
if !errorlevel! equ 0 (
  if exist ".git" (
    echo 正在检查更新…
    for /f "delims=" %%i in ('git rev-parse HEAD 2^>nul') do set "OLD=%%i"
    git fetch --quiet origin 2>nul
    if !errorlevel! equ 0 (
      git reset --hard origin/main 2>nul
      if !errorlevel! neq 0 ( echo 更新未应用，使用当前版本。 )
    ) else (
      echo 联网更新失败（可能断网），使用当前版本。
    )
    for /f "delims=" %%i in ('git rev-parse HEAD 2^>nul') do set "NEW=%%i"
    if defined NEW if not "!OLD!"=="!NEW!" ( echo 已更新到最新版。 )
  ) else (
    echo 提示：此目录非 git 安装，跳过自动更新（用当前版本）。
  )
) else (
  echo 提示：未检测到 Git，跳过自动更新（用当前版本）。装 Git for Windows 后即可自动更新。
)

REM --- 找可用的 Python 启动器 ---
set "PY="
py -3 --version >nul 2>&1 && set "PY=py -3"
if not defined PY ( python --version >nul 2>&1 && set "PY=python" )
if not defined PY ( python3 --version >nul 2>&1 && set "PY=python3" )

if not defined PY (
  echo.
  echo 未找到 Python 3。请先安装：https://www.python.org/downloads/
  echo 安装时务必勾选 "Add Python to PATH"。
  echo 若 "python" 打开了微软商店：设置 ^> 应用 ^> 高级应用设置 ^> 应用执行别名
  echo   -^> 关闭 "python.exe" 和 "python3.exe"。
  echo.
  pause
  goto :eof
)

set "VENV=.venv"
set "VPY=%VENV%\Scripts\python.exe"

REM 首次运行需建环境；更新后(版本号变化)依赖可能变，重装一次
set "NEED_DEP="
if not exist "%VPY%" (
  echo.
  echo 首次运行：正在创建运行环境（约 1-2 分钟，仅此一次）…
  %PY% -m venv "%VENV%" || ( echo 创建虚拟环境失败 & pause & goto :eof )
  set "NEED_DEP=1"
)
if defined OLD if defined NEW if not "!OLD!"=="!NEW!" set "NEED_DEP=1"

if defined NEED_DEP (
  echo 正在安装/更新依赖…
  "%VPY%" -m pip install --upgrade pip >nul 2>&1
  "%VPY%" -m pip install -r requirements.txt || ( echo 安装依赖失败（检查网络/代理） & pause & goto :eof )
)

echo 启动 VO 拉单工具…（关闭此窗口会一并关闭程序）
"%VPY%" gui.py

REM 程序异常退出时保留窗口，便于看报错
if errorlevel 1 (
  echo.
  echo 程序异常退出，请把上面的错误信息发给开发者。
  pause
)
endlocal
