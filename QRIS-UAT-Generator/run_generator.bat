@echo off
chcp 65001 >nul
title QRIS UAT Document Generator - Bank Sahabat Sampoerna
echo ============================================================
echo   QRIS Merchant Aggregator - UAT Document Generator
echo   Bank Sahabat Sampoerna (BSS)
echo ============================================================
echo.

:: Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python tidak ditemukan!
    echo.
    echo Silakan install Python terlebih dahulu:
    echo   https://www.python.org/downloads/
    echo.
    echo PENTING: Saat install, centang "Add Python to PATH"
    echo.
    pause
    exit /b 1
)

:: Check if required libraries are installed
echo [INFO] Mengecek library yang dibutuhkan...
pip show python-docx >nul 2>&1
if %errorlevel% neq 0 (
    echo [INFO] Installing python-docx...
    pip install python-docx
)
pip show openpyxl >nul 2>&1
if %errorlevel% neq 0 (
    echo [INFO] Installing openpyxl...
    pip install openpyxl
)
echo [OK] Library siap.
echo.

:: Find Excel file
set "EXCEL_FILE="
for %%f in (*.xlsx) do (
    if "%%f" neq "" (
        set "EXCEL_FILE=%%f"
        goto :found
    )
)

:notfound
echo [ERROR] File Excel (.xlsx) tidak ditemukan di folder ini!
echo.
echo Pastikan file UAT Script Excel (.xlsx) berada di folder yang sama
echo dengan file run_generator.bat ini.
echo.
pause
exit /b 1

:found
echo [INFO] File Excel ditemukan: %EXCEL_FILE%
echo.

:: Run the generator
python generate_uat_docs.py "%EXCEL_FILE%"

echo.
echo ============================================================
echo   Tekan tombol apa saja untuk menutup...
echo ============================================================
pause >nul
