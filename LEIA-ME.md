# Impulsionar V4 — Como rodar no seu computador

## ⚠️ Falta uma peça: o frontend (pasta `public/`)

Esta pasta `public/` está **vazia**. O backend (`server.js`) serve os arquivos
estáticos de dentro dela, mas o HTML/CSS/JS da tela de login que aparece no
seu print (com o erro "Erro de ligação ao servidor") **nunca foi enviado a
mim** nesta conversa — só recebi `server.js`, `package.json`,
`prisma.config.ts` e o `database.sqlite`.

Ou seja: se você já tem esse frontend rodando localmente, ele está em outra
pasta no seu computador (provavelmente uma pasta `public/` com `index.html`,
`app.js`, `style.css` etc.). **Copie essa pasta para dentro deste projeto,
substituindo esta `public/` vazia**, antes de rodar.

Se você não sabe onde esse frontend está, ou nunca chegou a existir como
arquivo separado, me avise — te ajudo a localizar ou a construir um do zero.

## O que já está pronto aqui

- `server.js` — backend Express com autenticação por token (JWT), isolamento
  de dados por empresa/executivo, dashboard com DISC/fases/alertas de PDI, e
  exports CSV protegidos contra CSV injection.
- `package.json` — lista de dependências (já inclui `dotenv` e
  `jsonwebtoken`, que o `server.js` novo usa).
- `database.sqlite` — seu banco de dados original, com duas correções:
  - Colunas que faltavam (`users.employee_id`, `mentorships.mentor_id`)
    foram adicionadas para bater com o `server.js` atual.
  - As senhas de `admin@impulsionar.com` e da sua conta
    (`lazaronettosalles0067@gmail.com`) foram resetadas para
    `Impulsionar#2026` (troque depois de logar).
- `.env.example` — modelo de variáveis de ambiente.
- `.gitignore` — evita subir `node_modules/` e `.env` para um repositório Git.

## Passo a passo para rodar

1. **Copie sua pasta `public/` (frontend) para dentro deste projeto**,
   substituindo a pasta vazia que veio aqui.

2. **Instale as dependências** (na pasta do projeto, via terminal):
   ```
   npm install
   ```

3. **Configure o `.env`**:
   - Copie `.env.example` para um novo arquivo chamado `.env`.
   - Gere um valor para `JWT_SECRET` rodando:
     ```
     node -e "console.log(require('crypto').randomBytes(48).toString('hex'))"
     ```
     e cole o resultado no `.env`.
   - Preencha `SMTP_USER`/`SMTP_PASS` se for usar o envio de e-mail (opcional
     para só testar o sistema).

4. **Inicie o servidor**:
   ```
   npm start
   ```
   Você deve ver no terminal:
   ```
   ✅ Base de dados Impulsionar V4 ativa com Notificações.
   🚀 Plataforma Completa Impulsionar V4 em http://127.0.0.1:3000
   ```

5. **Acesse no navegador**: `http://localhost:3000`

6. **Faça login** com:
   | Perfil | E-mail | Senha |
   |---|---|---|
   | Master (vê tudo) | `admin@impulsionar.com` | `Impulsionar#2026` |
   | Sua empresa ("impulsionar") | `lazaronettosalles0067@gmail.com` | `Impulsionar#2026` |

## Sobre o erro "Erro de ligação ao servidor" do seu print

Essa mensagem normalmente aparece quando o frontend não conseguiu nem
completar a chamada de rede até o backend — ou seja, o servidor Express não
estava rodando, estava rodando em outra porta, ou o frontend estava sendo
aberto direto como arquivo (`file://...`) em vez de servido pelo próprio
Express em `http://localhost:3000`. Depois de seguir os passos acima, se o
erro persistir, me mande:
- o que aparece no terminal quando você roda `npm start`;
- a URL exata que aparece na barra de endereço do navegador quando você vê o
  erro.
