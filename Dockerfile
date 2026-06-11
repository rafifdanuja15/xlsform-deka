FROM python:3.11-slim

# Install system deps
# - libreoffice: konversi .doc legacy ke .docx di server
# - antiword: fallback ringan untuk .doc
# - fonts-liberation: font agar LibreOffice tidak crash saat convert
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer \
    antiword \
    fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy project
COPY . .

# Create tmp dir
RUN mkdir -p tmp_uploads

# Expose port
EXPOSE 5000

# Run with gunicorn
CMD gunicorn "backend.app:app" \
    --bind 0.0.0.0:${PORT:-5000} \
    --workers 2 \
    --timeout 300 \
    --log-level info
