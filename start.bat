@echo off
chcp 65001 >nul 2>&1
cls
echo.
echo   ================================================
echo     Baoxiao Tool v1.0
echo     Expense ^& Reimbursement Tool
echo   ================================================
echo.
echo     Author : aigc creative person
echo     Engine : xlsxwriter / xlsxwriter
echo.
echo   ------------------------------------------------
echo     Auto-opening browser...
echo     Press Ctrl+C to stop
echo   ------------------------------------------------
echo.
cd /d "%~dp0%"
python main.py
pause
