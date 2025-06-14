# Gunakan base image python ringan
FROM python:3.11-slim

# Install LibreOffice & deps-nya
RUN apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-writer \
    fonts-dejavu \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Buat direktori kerja
WORKDIR /app

# Copy semua file ke image
COPY . /app

# Install dependencies dari requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# Jalankan app Flask
CMD ["python", "app.py"]
