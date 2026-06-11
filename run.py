"""
run.py — Entry point untuk development lokal
Jalankan: python run.py
"""

import os
from dotenv import load_dotenv

load_dotenv()

from backend.app import app

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_ENV", "production") == "development"
    print(f"\n🚀 XLSForm Converter berjalan di http://localhost:{port}")
    print(f"   Mode: {'development' if debug else 'production'}")
    print(f"   Model AI: {os.environ.get('SUMOPOD_MODEL', 'tidak dikonfigurasi')}\n")
    app.run(host="0.0.0.0", port=port, debug=debug)
