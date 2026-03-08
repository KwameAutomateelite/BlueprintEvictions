FROM python:3.12

RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer \
    fonts-liberation fonts-dejavu-core fonts-freefont-ttf \
    && rm -rf /var/lib/apt/lists/*

ENV HOME=/tmp

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY templates/ templates/
COPY main.py .
