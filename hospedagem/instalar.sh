#!/usr/bin/env bash
# Instala o Impulsionar V4 num servidor Ubuntu 22.04/24.04 novinho.
# Uso (como root):  bash instalar.sh
set -e
echo "=== Impulsionar V4 — instalação do servidor ==="
read -rp "Domínio (ex.: impulsionarv4.com.br): " DOMINIO
read -rp "Endereço do repositório (ex.: https://github.com/lazaro0067/Impulsionar.v4.git): " REPO
read -rsp "Token do GitHub (só leitura; deixe vazio se o repositório for público): " TOKEN; echo

apt-get update -y && apt-get upgrade -y
apt-get install -y git curl ufw sqlite3
if ! command -v docker >/dev/null; then curl -fsSL https://get.docker.com | sh; fi

# firewall: só SSH e site
ufw allow OpenSSH && ufw allow 80 && ufw allow 443 && ufw --force enable

# memória de reserva (evita travar em picos de vídeo)
if ! swapon --show | grep -q swap; then fallocate -l 4G /swapfile && chmod 600 /swapfile && mkswap /swapfile && swapon /swapfile && echo '/swapfile none swap sw 0 0' >> /etc/fstab; fi

mkdir -p /opt && cd /opt
if [ ! -d impulsionar ]; then
  if [ -n "$TOKEN" ]; then git clone "${REPO/https:\/\//https://$TOKEN@}" impulsionar; else git clone "$REPO" impulsionar; fi
fi
cd /opt/impulsionar
mkdir -p dados/uploads hospedagem/backups

if [ ! -f .env ]; then
  cat > .env <<ENV
DOMINIO=$DOMINIO
JWT_SECRET=$(openssl rand -hex 32)
# Copie aqui as outras variáveis que estão no Railway (ANTHROPIC_API_KEY, MP_ACCESS_TOKEN, SMTP..., TWILIO..., AUTENTIQUE...)
ENV
  echo ">> Arquivo /opt/impulsionar/.env criado. Edite com: nano /opt/impulsionar/.env"
fi

docker compose up -d --build

# backup diário às 3h
( crontab -l 2>/dev/null | grep -v impulsionar-backup; echo "0 3 * * * bash /opt/impulsionar/hospedagem/backup.sh >> /var/log/impulsionar-backup.log 2>&1 # impulsionar-backup" ) | crontab -

echo
echo "✅ Pronto! Aponte o domínio $DOMINIO (registro A) para o IP deste servidor:"
curl -fsS https://api.ipify.org || true
echo
echo "Depois de 1 a 10 minutos o site abre em https://$DOMINIO com cadeado (HTTPS)."
