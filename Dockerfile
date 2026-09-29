FROM python:3.12-slim

# weasyprint needs pango/cairo; mammoth needs nothing extra; noto-cjk for Chinese text
RUN apt-get update && apt-get install -y --no-install-recommends \
    libpango-1.0-0 \
    libpangocairo-1.0-0 \
    libcairo2 \
    libgdk-pixbuf-2.0-0 \
    libffi-dev \
    shared-mime-info \
    fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY app.py .
COPY start.sh .
RUN chmod +x start.sh

ENV PORT=8000
EXPOSE 8000

CMD ["./start.sh"]
