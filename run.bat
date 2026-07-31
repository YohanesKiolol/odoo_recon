@echo off
REM ── Check .venv exists ────────────────────────────────────────
if not exist ".venv\Scripts\python.exe" (
    echo [ERROR] Setup belum dijalankan.
    echo Jalankan setup.bat terlebih dahulu.
    pause
    exit /b 1
)

REM ── Launch GUI (no console window) ────────────────────────────
start "" .venv\Scripts\pythonw.exe gui.py
