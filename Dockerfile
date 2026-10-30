FROM python:3.11-slim

# Enable contrib and non-free repositories for Microsoft fonts
RUN sed -i 's/main/main contrib non-free non-free-firmware/g' /etc/apt/sources.list.d/debian.sources

# Install LibreOffice and true Microsoft fonts (Arial, Calibri, etc.) for perfect formatting
RUN apt-get update && apt-get install -y --no-install-recommends debconf-utils && \
    echo "ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true" | debconf-set-selections && \
    apt-get install -y --no-install-recommends \
    libreoffice \
    libreoffice-writer \
    ttf-mscorefonts-installer \
    fonts-liberation \
    fonts-dejavu \
    fonts-noto-core \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
ENV PORT=10000
CMD gunicorn app:app --bind 0.0.0.0:$PORT --workers 1 --timeout 120
