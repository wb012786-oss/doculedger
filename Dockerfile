FROM python:3.12-slim

WORKDIR /app

RUN apt-get update && apt-get install -y --no-install-recommends \
    libmupdf-dev \
    poppler-utils \
    tesseract-ocr \
    tesseract-ocr-chi-sim \
    tesseract-ocr-eng \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app/ ./app/
COPY samples/ ./samples/

RUN mkdir -p /tmp/bank_uploads && chmod 777 /tmp/bank_uploads

ENV PYTHONUNBUFFERED=1
ENV UPLOAD_DIR=/tmp/bank_uploads
ENV MAX_FILE_SIZE_MB=20
ENV MAX_FILES=5

EXPOSE 8000

CMD ["gunicorn", "app.main:app", "-k", "uvicorn.workers.UvicornWorker", "-b", "0.0.0.0:8000", "-w", "2", "--timeout", "180", "--access-logfile", "-", "--error-logfile", "-"]
