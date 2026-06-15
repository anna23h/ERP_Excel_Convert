@echo off
chcp 65001 >nul
setlocal
REM ============================================
REM VO 拉单工具 · Windows 双击运行（免打包，直跑最新代码）
REM 首次运行：自动创建本地虚拟环境 .venv 并装依赖；之后双击秒开。
REM 更新逻辑后：先 git pull，再双击本文件即可，无需重打 exe。
REM 前提：本机已装 Python 3（安装时勾选 Add Python to PATH）。
REM ============================================

REM 切到本脚本所在目录
cd /d "%~dp0"

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

if not exist "%VPY%" (
  echo.
  echo 首次运行：正在创建运行环境并安装依赖（约 1-2 分钟，仅此一次）…
  %PY% -m venv "%VENV%" || ( echo 创建虚拟环境失败 & pause & goto :eof )
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
