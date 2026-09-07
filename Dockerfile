FROM python:3.11-slim

# LibreOffice for .doc/.odt conversion — the same tool used throughout this project
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY extractor_core.py batch_runner.py credits.py app.py .

ENV PORT=8080
EXPOSE 8080

CMD exec gunicorn --bind 0.0.0.0:$PORT --workers 2 --timeout 300 app:app
