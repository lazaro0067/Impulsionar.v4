#!/usr/bin/env bash
# Backup do banco (cópia segura mesmo com o sistema rodando). Guarda os últimos 14 dias.
set -e
cd /opt/impulsionar
mkdir -p hospedagem/backups
DATA=$(date +%Y-%m-%d_%H%M)
if [ -f dados/database.sqlite ]; then
  sqlite3 dados/database.sqlite ".backup 'hospedagem/backups/banco-$DATA.sqlite'"
  gzip -f "hospedagem/backups/banco-$DATA.sqlite"
fi
find hospedagem/backups -name 'banco-*.sqlite.gz' -mtime +14 -delete
echo "Backup $DATA ok"
