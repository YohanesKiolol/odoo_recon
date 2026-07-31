@echo off
title Bank Reconciliation — Setup
echo ============================================================
echo   Bank Reconciliation Tool — First-Time Setup
echo ============================================================
echo.

REM ── Check Python ─────────────────────────────────────────────
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python tidak ditemukan.
    echo.
    echo Silahkan download dan install Python dari:
    echo   https://www.python.org/downloads/
    echo.
    echo PENTING: Centang "Add Python to PATH" saat instalasi.
    echo.
    pause
    exit /b 1
)

echo [OK] Python ditemukan:
python --version
echo.

REM ── Create virtual environment ────────────────────────────────
echo Membuat virtual environment...
python -m venv .venv
if %errorlevel% neq 0 (
    echo [ERROR] Gagal membuat virtual environment.
    pause
    exit /b 1
)

REM ── Install dependencies ──────────────────────────────────────
echo.
echo Menginstall dependencies (mungkin butuh beberapa menit)...
.venv\Scripts\pip install --upgrade pip --quiet
.venv\Scripts\pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [ERROR] Gagal install dependencies.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   Setup selesai! Sekarang bisa jalankan run.bat
echo ============================================================
echo.
pause
