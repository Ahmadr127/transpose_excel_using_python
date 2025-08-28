@echo off
chcp 65001 >nul
title Sistem Excel Processing

echo.
echo ================================================
echo    🎯 SISTEM EXCEL PROCESSING
echo ================================================
echo.

:: Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Error: Python tidak ditemukan!
    echo Silakan install Python 3.8+ dari https://python.org
    echo.
    pause
    exit /b 1
)

:: Check Python version
for /f "tokens=2" %%i in ('python --version 2^>^&1') do set PYTHON_VERSION=%%i
echo ✅ Python %PYTHON_VERSION% ditemukan

:: Create directories if they don't exist
if not exist "uploads" (
    mkdir uploads
    echo 📁 Direktori uploads dibuat
)

if not exist "outputs" (
    mkdir outputs
    echo 📁 Direktori outputs dibuat
)

if not exist "templates" (
    mkdir templates
    echo 📁 Direktori templates dibuat
)

:: Install requirements if needed
echo.
echo 📦 Checking dependencies...
pip install -r requirements.txt >nul 2>&1
if errorlevel 1 (
    echo ❌ Error saat install dependencies
    pause
    exit /b 1
)
echo ✅ Dependencies siap

:: Run the system
echo.
echo 🚀 Menjalankan Sistem Excel Processing...
echo 🌐 Browser akan terbuka otomatis di http://localhost:5000
echo ⏹️  Tekan Ctrl+C untuk menghentikan
echo.

python run.py

echo.
echo 👋 Sistem dihentikan
pause
