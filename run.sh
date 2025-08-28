#!/bin/bash

# Sistem Excel Processing - Shell Script
# Untuk Linux, macOS, dan WSL

echo ""
echo "================================================"
echo "   🎯 SISTEM EXCEL PROCESSING"
echo "================================================"
echo ""

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    if ! command -v python &> /dev/null; then
        echo "❌ Error: Python tidak ditemukan!"
        echo "Silakan install Python 3.8+ dari https://python.org"
        echo ""
        exit 1
    else
        PYTHON_CMD="python"
    fi
else
    PYTHON_CMD="python3"
fi

# Check Python version
PYTHON_VERSION=$($PYTHON_CMD --version 2>&1)
echo "✅ $PYTHON_VERSION ditemukan"

# Create directories if they don't exist
for dir in uploads outputs templates; do
    if [ ! -d "$dir" ]; then
        mkdir -p "$dir"
        echo "📁 Direktori $dir dibuat"
    fi
done

# Check if virtual environment exists
if [ ! -d "venv" ]; then
    echo ""
    echo "🔧 Membuat virtual environment..."
    $PYTHON_CMD -m venv venv
    echo "✅ Virtual environment dibuat"
fi

# Activate virtual environment
echo ""
echo "🔧 Mengaktifkan virtual environment..."
source venv/bin/activate

# Install requirements
echo ""
echo "📦 Installing dependencies..."
pip install -r requirements.txt
if [ $? -ne 0 ]; then
    echo "❌ Error saat install dependencies"
    exit 1
fi
echo "✅ Dependencies siap"

# Run the system
echo ""
echo "🚀 Menjalankan Sistem Excel Processing..."
echo "🌐 Browser akan terbuka otomatis di http://localhost:5000"
echo "⏹️  Tekan Ctrl+C untuk menghentikan"
echo ""

$PYTHON_CMD run.py

echo ""
echo "👋 Sistem dihentikan"
