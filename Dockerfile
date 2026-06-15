FROM python:3.11-slim

# Prevent Python from writing .pyc files and buffer logs
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# Install system dependencies
# - libreoffice-writer: convert .doc to .docx
# - antiword: lightweight fallback for .doc
# - fonts-liberation: avoid LibreOffice font issues
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer \
    antiword \
    fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip \
    && pip install --no-cache-dir -r requirements.txt

# Copy project files
COPY . .

# Create temporary upload directory
RUN mkdir -p tmp_uploads

# Railway injects PORT automatically
EXPOSE 5000

# Start application
CMD ["sh", "-c", "gunicorn backend.app:app --bind 0.0.0.0:${PORT:-5000} --workers 2 --timeout 300 --log-level info"]
