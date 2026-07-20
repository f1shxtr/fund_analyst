@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

set "PY_CMD="

py -3 --version >nul 2>nul
if not errorlevel 1 set "PY_CMD=py -3"

if not defined PY_CMD (
    python --version >nul 2>nul
    if not errorlevel 1 set "PY_CMD=python"
)

if not defined PY_CMD (
    echo 未检测到 Python。
    echo.
    echo 如果你拿到的是源码版，请先安装 Python 3.10 或更高版本。
    echo 如果你只是使用软件，建议让开发者提供 dist\资金流向分析工具.exe；
    echo exe 版本不需要安装 Python，也不需要安装任何库或配置镜像源。
    echo.
    pause
    exit /b 1
)

%PY_CMD% "%~dp0fund_flow_app.py"
if errorlevel 1 (
    echo.
    echo 程序启动失败。请把上面的报错截图发给开发者。
)

pause
