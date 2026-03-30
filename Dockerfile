FROM python:3.12-slim
RUN apt-get update && apt-get install -y \
    tesseract-ocr \
    tesseract-ocr-fra \
    libgl1 \
    libglib2.0-0 \
    libpango-1.0-0 \
    libpangocairo-1.0-0 \
    libgdk-pixbuf-2.0-0 \
    libffi-dev \
    libcairo2 \
    libgobject-2.0-0 \
    libharfbuzz0b \
    fonts-dejavu-core \
    libreoffice-nogui \
    && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
ARG CACHE_BUST=v1.5.1
COPY . .
EXPOSE 8000
CMD ["python", "main.py"]
