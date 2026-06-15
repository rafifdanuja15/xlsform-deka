FROM python:3.11-slim

# Python settings
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# Install system dependencies
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer \
    antiword \
    fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

# Working directory
WORKDIR /app

# Install Python dependencies
COPY requirements.txt .

RUN pip install --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# Copy application source
COPY . .

# Create temporary directory
RUN mkdir -p /app/tmp_uploads

# Railway injects PORT automatically
ENV PORT=5000
EXPOSE 5000

# Start Gunicorn
ENTRYPOINT ["gunicorn"]
CMD ["backend.app:app", "--bind", "0.0.0.0:5000", "--workers", "2", "--timeout", "300", "--log-level", "info"]
