@echo off
chcp 65001 >nul
REM ============================================
REM 在 Windows 上把 VO 拉单工具打包成单文件 exe
REM 前置：已安装 Python 3 (python.org，安装时勾选 Add to PATH)
REM 用法：把整个项目拷到本机，双击本文件，或在命令提示符里运行
REM ============================================

echo [1/2] 安装依赖...
python -m pip install -r requirements.txt
if errorlevel 1 goto fail

echo [2/2] 打包 exe...
pyinstaller --onefile --windowed --name VO拉单工具 --collect-all openpyxl gui.py
if errorlevel 1 goto fail

echo.
echo 完成！exe 路径：dist\VO拉单工具.exe
echo 把该 exe 发给员工，双击即可使用。
pause
exit /b 0

:fail
echo.
echo 打包失败。排查：先用「不带 --windowed」重打看报错；若提示缺 pandas 模块，
echo 在命令末尾加 --collect-all pandas 再试。
pause
exit /b 1
