FROM node:22-bookworm-slim
RUN apt-get update && apt-get install -y --no-install-recommends python3 python3-venv && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY requirements.txt .
RUN python3 -m venv /opt/parser && /opt/parser/bin/pip install --no-cache-dir -r requirements.txt
COPY --chown=node:node . .
USER node
CMD ["node", "start.mjs"]
