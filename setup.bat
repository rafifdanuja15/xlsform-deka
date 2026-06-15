@echo off
echo ============================================
echo   XLSForm Converter - Setup (Windows)
echo ============================================

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python tidak ditemukan. Install dari https://python.org
    pause
    exit /b 1
)

:: Create virtual env
if not exist "venv" (
    echo [1/4] Membuat virtual environment...
    python -m venv venv
)

:: Activate venv
echo [2/4] Mengaktifkan virtual environment...
call venv\Scripts\activate.bat

:: Install deps
echo [3/4] Menginstall dependencies...
pip install -r requirements.txt

:: Create .env if not exists
if not exist ".env" (
    echo [4/4] Membuat file .env dari template...
    copy .env.example .env
    echo.
    echo [!] PENTING: Edit file .env dan isi SUMOPOD_API_KEY Anda!
    echo     Buka .env dengan text editor dan ganti sk-xxxx dengan API key asli.
) else (
    echo [4/4] File .env sudah ada, lewati.
)

echo.
echo ============================================
echo   Setup selesai!
echo   Langkah berikutnya:
echo   1. Edit .env dan isi SUMOPOD_API_KEY
echo   2. Jalankan: python run.py
echo   3. Buka: http://localhost:5000
echo ============================================
pause
