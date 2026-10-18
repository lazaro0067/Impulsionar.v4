# Impulsionar V4 — imagem para servidor próprio (VPS)
# (Fica dentro de hospedagem/ de propósito: assim o Railway continua funcionando como antes.)
FROM node:20-bookworm-slim

# ffmpeg: vídeos longos, transcrição de aulas e conversão para MP4
# python3/make/g++: compilam sqlite3 e bcrypt
RUN apt-get update \
 && apt-get install -y --no-install-recommends ffmpeg python3 make g++ ca-certificates curl sqlite3 \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY package*.json ./
RUN npm ci --omit=dev || npm install --omit=dev
COPY . .

ENV NODE_ENV=production \
    PORT=3000 \
    DB_PATH=/dados/database.sqlite \
    UPLOADS_PATH=/dados/uploads \
    NODE_OPTIONS=--max-old-space-size=3072

RUN mkdir -p /dados/uploads
EXPOSE 3000
HEALTHCHECK --interval=30s --timeout=10s --start-period=60s --retries=3 CMD curl -fsS http://127.0.0.1:3000/saude || exit 1
CMD ["node", "server.js"]
