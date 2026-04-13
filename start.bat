@echo off
chcp 65001 >nul 2>&1
cls
echo.
echo   ================================================
echo     报销费用填写工具 v1.0
echo     Expense ^& Reimbursement Tool
echo   ================================================
echo.
echo     Author : aigc创意人竹相左边
echo     Engine : Python + xlsxwriter
echo.
echo   ------------------------------------------------
echo     启动后浏览器会自动打开
echo     按 Ctrl+C 可停止服务器
echo   ------------------------------------------------
echo.
cd /d "%~dp0%"
python server_cloud.py
pause
