@echo off
:: 报销费用填写工具 - 云服务器版 启动脚本 (Windows)
:: 双击运行，启动后用 http://localhost:8765 访问
:: 如需外网访问，改 port 为 80 或在 Nginx 反代
:: =========================================

title 报销费用填写工具 · 云服务器版
color 0A

echo.
echo ================================================
echo  报销费用填写工具 · 云服务器版
echo ================================================
echo.

:: ---- 检查依赖 ----
python -c "import openpyxl, PIL" 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] 缺少依赖，请先运行【云端部署_安装依赖.bat】
    echo.
    pause
    exit /b 1
)

:: ---- 启动 ----
echo [INFO] 正在启动服务 ...
echo.
echo  访问地址:
echo   本机:    http://localhost:8765
echo   局域网:  http://%ComputerName%:8765
echo.
echo  停止服务: Ctrl+C
echo.
echo ================================================
echo.

python server_cloud.py

pause
