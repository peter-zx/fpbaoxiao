@echo off
chcp 65001 >nul
echo ============================================
echo   报销费用填写工具 - 启动服务器
echo ============================================
echo.
echo 正在启动本地服务器...
echo 服务地址: http://localhost:8765
echo.
echo 启动后浏览器会自动打开
echo 按 Ctrl+C 可停止服务器
echo ============================================
echo.

cd /d "%~dp0%"
python server_cloud.py

pause