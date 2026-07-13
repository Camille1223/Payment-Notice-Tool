FROM python:3.12-slim

# weasyprint needs pango/cairo/gdk-pixbuf; libreoffice for higher-fidelity PDF
RUN apt-get update && apt-get install -y --no-install-recommends \
    libpango-1.0-0 libpangocairo-1.0-0 libcairo2 libgdk-pixbuf-2.0-0 \
    libffi-dev shared-mime-info \
    fonts-noto-cjk fonts-noto-cjk-extra \
    libreoffice-writer libreoffice-calc \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY app.py .

ENV PORT=8000
EXPOSE 8000

# 4 workers, 5-min timeout (PDF conversion can be slow for large batches)
CMD ["gunicorn", "--bind", "0.0.0.0:8000", "--timeout", "300", "--workers", "2", "--threads", "4", "app:app"]
