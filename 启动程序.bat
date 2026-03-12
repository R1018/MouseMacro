@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 正在启动 MouseMacro...
echo.
pythonw main.py
if errorlevel 1 (
    echo.
    echo 程序启动失败！
    echo 请确保已安装 Python 和所需的依赖包。
    echo 运行以下命令安装依赖：
    echo pip install -r requirements.txt
    pause
)
