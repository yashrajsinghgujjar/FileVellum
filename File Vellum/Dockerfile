FROM python:3.10-slim

RUN apt-get update && apt-get install -y \
    tesseract-ocr \
    libtesseract-dev \
    poppler-utils \
    ffmpeg \
    libsm6 \
    libxext6 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
RUN pip install --no-cache-dir gunicorn

COPY . .

EXPOSE 10000
CMD ["gunicorn", "--workers=2", "--threads=2", "--bind=0.0.0.0:10000", "app:app"]