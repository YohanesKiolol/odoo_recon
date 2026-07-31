@echo off
title Build EXE — Bank Reconciliation
echo ============================================================
echo   Build BankRekonsiliasi.exe dengan PyInstaller
echo ============================================================
echo.

if not exist ".venv\Scripts\python.exe" (
    echo [ERROR] Jalankan setup.bat dulu.
    pause
    exit /b 1
)

echo Menginstall PyInstaller...
.venv\Scripts\pip install pyinstaller --quiet

echo Membersihkan build sebelumnya...
if exist "dist" rmdir /s /q dist
if exist "build" rmdir /s /q build

echo.
echo Building EXE...
.venv\Scripts\pyinstaller bank_recon.spec

echo.
if exist "dist\BankRekonsiliasi.exe" (
    echo ============================================================
    echo   SUKSES!
    echo.
    echo   Kirimkan folder ini ke user lain (copy semua):
    echo     dist\BankRekonsiliasi.exe
    echo     .env
    echo     input\   ^(buat folder kosong ini^)
    echo     output\  ^(buat folder kosong ini^)
    echo.
    echo   User cukup:
    echo     1. Taruh file bank ke folder input\
    echo     2. Double-click BankRekonsiliasi.exe
    echo     3. Klik Jalankan
    echo ============================================================
    explorer dist
) else (
    echo [ERROR] Build gagal. Lihat pesan error di atas.
)
echo.
pause
