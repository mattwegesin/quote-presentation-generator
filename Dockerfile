FROM python:3.11-slim

# Accept EULA for Microsoft Fonts and install LibreOffice and fonts for PPTX/DOCX to PDF conversion
RUN echo "ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true" | debconf-set-selections && \
    apt-get update && apt-get install -y --no-install-recommends \
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

# Run gunicorn with a single worker to prevent Out-Of-Memory (OOM) on Render Free Tier
ENV PORT=10000
CMD gunicorn app:app --bind 0.0.0.0:$PORT --workers 1 --timeout 120
