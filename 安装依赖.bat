@echo off
chcp 65001 >nul
echo ========================================
echo  MouseMacro 依赖安装程序
echo ========================================
echo.

echo 检查 Python 是否已安装...
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到 Python！
    echo 请先安装 Python 3.7 或更高版本
    echo 下载地址: https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

echo [成功] Python 已安装
python --version
echo.

echo 正在安装依赖包...
echo.
pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo [错误] 依赖安装失败！
    echo 请检查网络连接或尝试手动安装：
    echo pip install pywin32 pynput
    echo.
    pause
    exit /b 1
)

echo.
echo ========================================
echo [成功] 所有依赖已安装完成！
echo ========================================
echo.
echo 现在可以运行"启动程序.bat"来启动程序
echo 或直接运行: python main.py
echo.
pause
