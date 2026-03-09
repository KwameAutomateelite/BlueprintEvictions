FROM python:3.12-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer \
    default-jre-headless \
    libxinerama1 libxrandr2 libxi6 libxtst6 libx11-xcb1 \
    libdbus-glib-1-2 libcairo2 libcups2 libglib2.0-0 \
    fonts-liberation fonts-dejavu-core fonts-freefont-ttf \
    && rm -rf /var/lib/apt/lists/*

ENV JAVA_HOME=/usr/lib/jvm/default-java

ENV HOME=/tmp

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY templates/ templates/
COPY main.py .
