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
    echo 未检测到 Python，无法打包。
    echo 请先在这台开发电脑安装 Python 3.10 或更高版本。
    pause
    exit /b 1
)

%PY_CMD% -m PyInstaller --version >nul 2>nul
if errorlevel 1 (
    echo 未检测到 PyInstaller，正在尝试安装打包工具...
    %PY_CMD% -m pip install pyinstaller -i https://pypi.tuna.tsinghua.edu.cn/simple
    if errorlevel 1 (
        echo.
        echo 清华源安装失败，改用 PyPI 官方源重试...
        %PY_CMD% -m pip install pyinstaller
    )
)

%PY_CMD% -m PyInstaller --noconfirm --clean --onefile --windowed --name "资金流向分析工具" "fund_flow_app.py"
if errorlevel 1 (
    echo.
    echo 打包失败。请把上面的报错截图发给开发者。
    pause
    exit /b 1
)

echo.
echo 打包完成：
echo "%~dp0dist\资金流向分析工具.exe"
echo.
echo 把 dist 目录里的 exe 发给用户即可。用户不需要安装 Python、库或镜像源。
pause
