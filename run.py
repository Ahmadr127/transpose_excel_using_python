#!/usr/bin/env python3
"""
Script untuk menjalankan Sistem Excel Processing
"""

import os
import sys
import subprocess
import webbrowser
import time

def check_python_version():
    """Check Python version"""
    if sys.version_info < (3, 8):
        print("❌ Error: Python 3.8 atau lebih tinggi diperlukan")
        print(f"Versi Python saat ini: {sys.version}")
        return False
    return True

def install_requirements():
    """Install requirements jika belum ada"""
    try:
        import flask
        import pandas
        import openpyxl
        print("✅ Semua dependencies sudah terinstall")
        return True
    except ImportError:
        print("📦 Installing dependencies...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
            print("✅ Dependencies berhasil diinstall")
            return True
        except subprocess.CalledProcessError:
            print("❌ Error saat install dependencies")
            return False

def create_directories():
    """Buat direktori yang diperlukan"""
    directories = ['uploads', 'outputs', 'templates']
    for directory in directories:
        if not os.path.exists(directory):
            os.makedirs(directory)
            print(f"📁 Direktori {directory} dibuat")

def run_system():
    """Jalankan sistem"""
    print("🚀 Menjalankan Sistem Excel Processing...")
    
    # Import dan jalankan Flask app
    try:
        from app import app
        print("✅ Aplikasi berhasil dimuat")
        print("🌐 Buka browser dan akses: http://localhost:5000")
        print("⏹️  Tekan Ctrl+C untuk menghentikan")
        
        # Buka browser otomatis setelah delay
        def open_browser():
            time.sleep(2)
            webbrowser.open('http://localhost:5000')
        
        import threading
        threading.Thread(target=open_browser, daemon=True).start()
        
        # Jalankan Flask app
        app.run(debug=True, host='0.0.0.0', port=5000)
        
    except Exception as e:
        print(f"❌ Error saat menjalankan aplikasi: {e}")
        return False

def main():
    """Main function"""
    print("=" * 50)
    print("🎯 SISTEM EXCEL PROCESSING")
    print("=" * 50)
    
    # Check Python version
    if not check_python_version():
        return
    
    # Create directories
    create_directories()
    
    # Install requirements
    if not install_requirements():
        return
    
    # Run system
    run_system()

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n👋 Sistem dihentikan oleh user")
    except Exception as e:
        print(f"\n❌ Error tidak terduga: {e}")
        print("Silakan cek log error di atas")
