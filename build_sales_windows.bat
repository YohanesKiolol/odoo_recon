@echo off
echo ============================================================
echo   Building Standalone Windows SalesPortal.exe
echo ============================================================
call .venv\Scripts\activate.bat 2>nul
python build_sales.py
pause
