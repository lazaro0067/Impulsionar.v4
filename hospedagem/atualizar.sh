#!/usr/bin/env bash
# Atualiza o sistema com a última versão do GitHub (sem perder dados).
set -e
cd /opt/impulsionar
bash hospedagem/backup.sh
git pull
docker compose up -d --build
docker image prune -f >/dev/null
echo "✅ Atualizado."
