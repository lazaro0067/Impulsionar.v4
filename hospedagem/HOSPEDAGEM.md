# 🖥️ Servidor próprio para o Impulsionar V4

Com um servidor só seu (VPS), o sistema não divide memória com ninguém, aguenta vídeos grandes e aulas de 2 h,
e os dados ficam num disco que não some a cada atualização.

## 1. Contratar o servidor (VPS)
Configuração recomendada:
- **4 vCPU · 8 GB de RAM · 200 GB SSD/NVMe** (aulas de 2 h ocupam ~2,3 GB cada)
- Sistema **Ubuntu 24.04**
- Datacenter no **Brasil (São Paulo)** — fica mais rápido para você e seus alunos
- Exemplos para comparar: Hostinger (VPS KVM), AWS Lightsail (região São Paulo), Magalu Cloud, Locaweb.

Ao contratar, anote o **IP** e a **senha de root** (ou chave SSH).

## 2. Instalar (10 minutos)
No seu computador, abra o PowerShell e entre no servidor:
```
ssh root@IP_DO_SERVIDOR
```
Crie um **token do GitHub** (só leitura do repositório): GitHub › Settings › Developer settings ›
Fine-grained tokens › Generate new token › Repository access: só `Impulsionar.v4` › Permissions: *Contents: Read-only*.

Depois rode no servidor:
```
curl -fsSL https://raw.githubusercontent.com/lazaro0067/Impulsionar.v4/main/hospedagem/instalar.sh -o instalar.sh
bash instalar.sh
```
(Se o repositório for privado e o comando acima não baixar, copie o arquivo `hospedagem/instalar.sh` para o servidor e rode `bash instalar.sh`.)

O script instala tudo sozinho: Docker, ffmpeg, firewall, memória de reserva, HTTPS automático e backup diário.

## 3. Copiar as configurações do Railway
No Railway › seu serviço › **Variables**, copie todas as variáveis para o servidor:
```
nano /opt/impulsionar/.env
```
Cole uma por linha (`NOME=valor`), salve (Ctrl+O, Enter, Ctrl+X) e reinicie:
```
cd /opt/impulsionar && docker compose up -d
```
> Dica: use o **mesmo JWT_SECRET** do Railway para ninguém precisar logar de novo.

## 4. Trazer os dados do Railway
1. Logado como Master, baixe o banco: **https://www.impulsionarv4.com.br/api/admin/database-backup**
   (abra com o sistema logado; o arquivo `backup-impulsionar.sqlite` será baixado).
2. Envie para o servidor (no PowerShell do seu computador):
   ```
   scp backup-impulsionar.sqlite root@IP_DO_SERVIDOR:/opt/impulsionar/dados/database.sqlite
   ```
3. Reinicie: `cd /opt/impulsionar && docker compose restart app`

## 5. Apontar o domínio
No **registro.br** (ou onde o domínio foi comprado) › DNS:
- `impulsionarv4.com.br` → tipo **A** → IP do servidor
- `www` → tipo **A** → IP do servidor

Em até ~1 hora (normalmente minutos) o site abre com cadeado 🔒. Só depois desligue o Railway.

## Dia a dia
| O que | Comando (no servidor) |
|---|---|
| Atualizar para a versão nova do GitHub | `bash /opt/impulsionar/hospedagem/atualizar.sh` |
| Ver se está no ar | `curl https://SEU_DOMINIO/saude` |
| Ver os registros (logs) | `cd /opt/impulsionar && docker compose logs -f app` |
| Reiniciar | `cd /opt/impulsionar && docker compose restart app` |
| Backup manual | `bash /opt/impulsionar/hospedagem/backup.sh` |

- Backups do banco: `/opt/impulsionar/hospedagem/backups` (14 dias, todo dia às 3h).
- Vídeos e arquivos: `/opt/impulsionar/dados/uploads`.
- Se o sistema travar, ele reinicia sozinho (verificação de saúde a cada 30 s).
