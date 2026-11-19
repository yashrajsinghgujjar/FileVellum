FROM python:3.10-slim

ENV U2NET_HOME=/app/.u2net \
    ORT_LOG_SEVERITY_LEVEL=3 \
    OMP_NUM_THREADS=2 \
    MKL_NUM_THREADS=2

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

# Pre-download lightweight rembg model during build so runtime doesn't stall
RUN python -c "from rembg import new_session; new_session('u2netp')"

COPY . .

EXPOSE 10000
CMD ["gunicorn", "--workers=1", "--threads=4", "--timeout=120", "--bind=0.0.0.0:10000", "app:app"]