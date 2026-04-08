@echo off
:: 报销费用填写工具 - Windows 一键安装脚本
:: 用法: 双击运行此文件
:: 确保已安装 Python 3.8+ (python.org 下载)
:: =========================================

title 报销费用填写工具 - 安装程序
color 0A

echo.
echo ================================================
echo  报销费用填写工具 · Windows 一键安装
echo ================================================
echo.

:: ---- 检查 Python ----
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] 未找到 Python，请先从 python.org 安装 Python 3.8+
    echo.
    pause
    exit /b 1
)

for /f "tokens=*" %%i in ('python --version') do set PYVER=%%i
echo [INFO] 检测到: %PYVER%

:: ---- 检查 pip ----
where pip >nul 2>&1
if %errorlevel% neq 0 (
    echo [INFO] 使用 python -m pip
    set "PIP=python -m pip"
) else (
    set "PIP=pip"
)

:: ---- 安装依赖 ----
echo.
echo [INFO] 正在安装依赖 openpyxl Pillow ...
echo.

%PIP% install --quiet --upgrade pip
%PIP% install --quiet openpyxl>=3.1.0 Pillow>=10.0.0

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] 依赖安装失败，请尝试以管理员身份运行此脚本
    echo 或手动运行: pip install openpyxl Pillow
    echo.
    pause
    exit /b 1
)

:: ---- 验证 ----
python -c "import openpyxl, PIL" 2>nul
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] 依赖验证失败，请重启命令行后重试
    echo.
    pause
    exit /b 1
)

:: ---- 创建目录 ----
if not exist "images" mkdir images
if not exist "exports" mkdir exports
if not exist "data"    mkdir data

:: ---- 创建 config.json ----
if not exist "config.json" (
    (
        echo {
        echo     "host": "0.0.0.0",
        echo     "port": 8765,
        echo     "Cors_origins": ["*"],
        echo     "log_level": "INFO",
        echo     "max_content_length": 20971520
        echo }
    ) > config.json
    echo [INFO] 配置文件 config.json 已创建
)

echo.
echo ================================================
echo  安装完成！ ✅
echo ================================================
echo.
echo 下一步：
echo   1. 双击【云端版启动.bat】启动服务器
echo   2. 浏览器打开 http://localhost:8765
echo.
echo  云服务器部署请上传整个文件夹到服务器后运行:
echo   install.sh
echo.
pause
