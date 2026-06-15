#!/bin/bash
set -e

echo "============================================"
echo "  XLSForm Converter - Setup (Mac/Linux)"
echo "============================================"

# Check Python
if ! command -v python3 &> /dev/null; then
    echo "[ERROR] Python 3 tidak ditemukan. Install dari https://python.org"
    exit 1
fi

# Create virtual env
if [ ! -d "venv" ]; then
    echo "[1/4] Membuat virtual environment..."
    python3 -m venv venv
fi

# Activate
echo "[2/4] Mengaktifkan virtual environment..."
source venv/bin/activate

# Install deps
echo "[3/4] Menginstall dependencies..."
pip install -r requirements.txt

# Create .env
if [ ! -f ".env" ]; then
    echo "[4/4] Membuat file .env dari template..."
    cp .env.example .env
    echo ""
    echo "[!] PENTING: Edit file .env dan isi SUMOPOD_API_KEY Anda!"
    echo "    Buka .env dengan text editor: nano .env"
else
    echo "[4/4] File .env sudah ada, lewati."
fi

echo ""
echo "============================================"
echo "  Setup selesai!"
echo "  Langkah berikutnya:"
echo "  1. Edit .env dan isi SUMOPOD_API_KEY"
echo "  2. Jalankan: source venv/bin/activate && python run.py"
echo "  3. Buka: http://localhost:5000"
echo "============================================"
