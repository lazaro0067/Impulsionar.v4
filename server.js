require('dotenv').config();
const express = require('express');
const http = require('http');
const { Server: SocketIOServer } = require('socket.io');
const sqlite3 = require('sqlite3').verbose();
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const cors = require('cors');
const bcrypt = require('bcrypt');
const nodemailer = require('nodemailer');
const jwt = require('jsonwebtoken');
const multer = require('multer');
const PDFDocument = require('pdfkit');
const { MercadoPagoConfig, PreApproval, Preference, Payment, PaymentRefund } = require('mercadopago');

// Base de perguntas do checklist "DPO AMBEV" (consultoria por pilares) — vem de
// uma planilha modelo da Ambev e é conteúdo estático (não muda pela tela),
// então fica num JSON à parte em vez de virar uma tabela gigante no banco.
const DPO_AMBEV_DATA = require('./dpo_ambev_data.json');
const DPO_PILARES_ORDEM = ['gente', 'seguranca', 'planejamento', 'armazem', 'frota', 'entrega', 'gestao'];

const app = express();
const PORT = process.env.PORT || 3000;

// Pasta onde ficam os arquivos enviados pelos usuários (vídeos da Academy,
// vídeos de bio dos mentores). Fica dentro de /public para ser servida
// diretamente pelo Express como arquivo estático.
// Se UPLOADS_PATH estiver definido (ex.: uma subpasta dentro do mesmo Volume
// persistente usado pelo DB_PATH, tipo /data/uploads), os arquivos enviados
// (contratos anexados, fotos, vídeos etc.) ficam FORA da pasta do código —
// senão eles são apagados a cada novo deploy, igual acontecia com o banco.
const PASTA_UPLOADS = process.env.UPLOADS_PATH || path.join(__dirname, 'public', 'uploads');
if (!fs.existsSync(PASTA_UPLOADS)) fs.mkdirSync(PASTA_UPLOADS, { recursive: true });
console.log(`📁 Uploads salvos em: ${PASTA_UPLOADS}${process.env.UPLOADS_PATH ? ' (persistente, via UPLOADS_PATH)' : ' (⚠️ dentro da pasta do código — configure UPLOADS_PATH com um Volume para não perder arquivos a cada deploy)'}`);

const armazenamentoUpload = multer.diskStorage({
    destination: (req, file, cb) => cb(null, PASTA_UPLOADS),
    filename: (req, file, cb) => {
        const sufixo = Date.now() + '-' + Math.round(Math.random() * 1e9);
        const extensao = path.extname(file.originalname);
        cb(null, file.fieldname + '-' + sufixo + extensao);
    }
});
const upload = multer({
    storage: armazenamentoUpload,
    limits: { fileSize: 300 * 1024 * 1024 }, // 300MB — vídeos são arquivos grandes
    fileFilter: (req, file, cb) => {
        const permitidos = /video\/|image\/|application\/pdf|application\/msword|application\/vnd\.openxmlformats-officedocument\.wordprocessingml\.document/;
        if (permitidos.test(file.mimetype)) return cb(null, true);
        cb(new Error('Tipo de arquivo não permitido. Envie um vídeo, imagem, PDF ou Word.'));
    }
});

// ATENÇÃO: defina JWT_SECRET no .env em produção. Este fallback só existe
// para o servidor não quebrar em ambiente de desenvolvimento sem .env configurado.
const JWT_SECRET = process.env.JWT_SECRET || 'dev-only-troque-isto-no-env';
if (!process.env.JWT_SECRET) {
    console.warn('⚠️  JWT_SECRET não definido no .env — usando valor de desenvolvimento. NÃO use isso em produção.');
}

app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb', extended: true }));

const transporter = nodemailer.createTransport({
    host: process.env.SMTP_HOST || 'smtp.mailtrap.io',
    port: Number(process.env.SMTP_PORT) || 2525,
    auth: {
        user: process.env.SMTP_USER || 'seu_usuario_smtp',
        pass: process.env.SMTP_PASS || 'sua_senha_smtp'
    }
});

// ---------- Convites de calendário (.ics) — integração com Outlook e Gmail ----------
// Não precisa de credenciais de API do Google/Microsoft: um e-mail com um anexo
// .ics no formato iCalendar (METHOD:REQUEST) é reconhecido nativamente tanto pelo
// Gmail quanto pelo Outlook/Exchange, que oferecem "Adicionar à agenda" / aceitar
// convite automaticamente. Reenviar com o MESMO UID e um SEQUENCE maior atualiza
// o evento já existente na agenda de quem recebeu (em vez de criar um duplicado) —
// é assim que "editar e enviar de novo" atualiza a agenda de todo mundo.
function formatarDataICS(dataIso) {
    const d = new Date(dataIso);
    const pad = n => String(n).padStart(2, '0');
    return `${d.getUTCFullYear()}${pad(d.getUTCMonth() + 1)}${pad(d.getUTCDate())}T${pad(d.getUTCHours())}${pad(d.getUTCMinutes())}${pad(d.getUTCSeconds())}Z`;
}

function gerarICS({ uid, sequence, titulo, descricao, inicioIso, duracaoMinutos, organizerEmail, organizerName, attendees, status }) {
    const dtStart = formatarDataICS(inicioIso);
    const fimData = new Date(new Date(inicioIso).getTime() + (Number(duracaoMinutos) || 60) * 60000);
    const dtEnd = formatarDataICS(fimData.toISOString());
    const escapeICS = t => String(t || '').replace(/\\/g, '\\\\').replace(/;/g, '\\;').replace(/,/g, '\\,').replace(/\n/g, '\\n');
    const linhasAttendee = (attendees || [])
        .filter(a => a.email)
        .map(a => `ATTENDEE;CN=${escapeICS(a.name || a.email)};RSVP=TRUE:mailto:${a.email}`)
        .join('\r\n');
    return [
        'BEGIN:VCALENDAR',
        'PRODID:-//Impulsionar V4//Mentoria//PT-BR',
        'VERSION:2.0',
        'CALSCALE:GREGORIAN',
        `METHOD:${status === 'Cancelada' ? 'CANCEL' : 'REQUEST'}`,
        'BEGIN:VEVENT',
        `UID:${uid}`,
        `SEQUENCE:${sequence || 0}`,
        `DTSTAMP:${formatarDataICS(new Date().toISOString())}`,
        `DTSTART:${dtStart}`,
        `DTEND:${dtEnd}`,
        `SUMMARY:${escapeICS(titulo)}`,
        `DESCRIPTION:${escapeICS(descricao)}`,
        `ORGANIZER;CN=${escapeICS(organizerName || 'Impulsionar')}:mailto:${organizerEmail}`,
        linhasAttendee,
        `STATUS:${status === 'Cancelada' ? 'CANCELLED' : 'CONFIRMED'}`,
        'END:VEVENT',
        'END:VCALENDAR'
    ].filter(Boolean).join('\r\n');
}

// Monta e dispara o convite de calendário para o executivo, o mentor e os
// participantes extras (internos já resolvidos + externos por e-mail).
async function enviarConvitesMentoria(mentoria, attendees) {
    if (!attendees || !attendees.length) return;
    const organizerEmail = process.env.SMTP_FROM_EMAIL || 'no-reply@impulsionar.com';
    const ics = gerarICS({
        uid: mentoria.ics_uid,
        sequence: mentoria.ics_sequence || 0,
        titulo: `Mentoria: ${mentoria.topics || 'Sessão de Mentoria'}`,
        descricao: `Sessão de mentoria agendada na plataforma Impulsionar.\nTópicos: ${mentoria.topics || ''}`,
        inicioIso: mentoria.meeting_date,
        duracaoMinutos: mentoria.duration_minutes || 60,
        organizerEmail,
        organizerName: 'Impulsionar',
        attendees,
        status: mentoria.status
    });
    const destinatarios = attendees.filter(a => a.email).map(a => a.email);
    if (!destinatarios.length) return;
    try {
        await transporter.sendMail({
            from: process.env.SMTP_FROM || '"Impulsionar V4" <no-reply@impulsionar.com>',
            to: destinatarios.join(', '),
            subject: `${mentoria.status === 'Cancelada' ? 'Cancelado: ' : ''}Mentoria: ${mentoria.topics || 'Sessão de Mentoria'}`,
            text: `Sua mentoria foi ${mentoria.status === 'Cancelada' ? 'cancelada' : 'agendada/atualizada'}. Duração: ${mentoria.duration_minutes || 60} minutos.`,
            icalEvent: {
                filename: 'convite.ics',
                method: mentoria.status === 'Cancelada' ? 'CANCEL' : 'REQUEST',
                content: ics
            }
        });
    } catch (erroEnvio) {
        console.warn('⚠️  Não foi possível enviar o convite de calendário da mentoria (SMTP não configurado?).', erroEnvio.message);
    }
}

// ---------- Mercado Pago (assinaturas recorrentes por empresa) ----------
// Em modo sandbox, use um Access Token de teste (começa com "TEST-"), gerado
// nas credenciais de teste da sua conta Mercado Pago. Sem token configurado,
// as rotas de assinatura respondem com um erro amigável em vez de derrubar o servidor.
const MP_ACCESS_TOKEN = process.env.MP_ACCESS_TOKEN || '';
const APP_BASE_URL_ENV = process.env.APP_BASE_URL || 'http://localhost:' + (process.env.PORT || 3000);
// URL pública realmente em uso agora (banco de dados, com fallback pro .env).
// O Mercado Pago EXIGE que back_urls.success seja uma URL pública válida
// quando auto_return é usado — "http://localhost:..." é rejeitado com
// "auto_return invalid. back_url.success must be defined" assim que o link
// de pagamento é gerado em produção. Por isso essa URL também é configurável
// direto em "Meu Perfil", sem precisar editar o .env.
let appBaseUrlAtiva = APP_BASE_URL_ENV;
function urlPublicaValida(url) {
    if (!url) return false;
    try {
        const u = new URL(url);
        if (!['http:', 'https:'].includes(u.protocol)) return false;
        if (['localhost', '127.0.0.1'].includes(u.hostname)) return false;
        return true;
    } catch (e) { return false; }
}
// Monta o pedaço do body da preferência do Mercado Pago referente às URLs de
// retorno. O Mercado Pago rejeita "auto_return" quando back_urls.success não
// é uma URL pública válida (localhost não vale) — em vez de deixar isso
// estourar como erro na hora de gerar o link, aqui a gente simplesmente não
// manda back_urls/auto_return quando a URL configurada ainda não é válida; o
// link de pagamento continua funcionando normalmente, só não redireciona
// automaticamente de volta para o sistema depois do pagamento.
function montarRetornoMercadoPago() {
    if (!urlPublicaValida(appBaseUrlAtiva)) return {};
    return {
        back_urls: { success: appBaseUrlAtiva, failure: appBaseUrlAtiva, pending: appBaseUrlAtiva },
        auto_return: 'approved'
    };
}
let mpClient = null;
let mpPreApproval = null;
let mpPreference = null; // pagamento único (Checkout Pro) — usado na cobrança por vaga divulgada
let mpPayment = null;
let mpPaymentRefund = null; // estorno de um pagamento único já confirmado (ex.: vaga divulgada)
let mpAccessTokenAtivo = ''; // token realmente em uso agora (banco de dados, com fallback pro .env)

// Monta (ou desmonta) o cliente do Mercado Pago com o token informado. Chamada
// na subida do servidor (com o .env) e de novo, ao vivo, sempre que o Master
// salva/atualiza o token pela tela "Meu Perfil" — sem precisar reiniciar nada.
function configurarMercadoPago(token) {
    mpAccessTokenAtivo = (token || '').trim();
    if (mpAccessTokenAtivo) {
        mpClient = new MercadoPagoConfig({ accessToken: mpAccessTokenAtivo });
        mpPreApproval = new PreApproval(mpClient);
        mpPreference = new Preference(mpClient);
        mpPayment = new Payment(mpClient);
        mpPaymentRefund = new PaymentRefund(mpClient);
    } else {
        mpClient = null; mpPreApproval = null; mpPreference = null; mpPayment = null; mpPaymentRefund = null;
    }
}
configurarMercadoPago(MP_ACCESS_TOKEN);
if (!MP_ACCESS_TOKEN) {
    console.warn('⚠️  MP_ACCESS_TOKEN não definido no .env — configure em "Meu Perfil > Mercado Pago" dentro do sistema, ou defina no .env.');
}

// ---------- Autentique (assinatura digital de contratos) ----------
// Crie uma conta gratuita em https://autentique.com.br, gere uma API key em
// "Perfil > API" e coloque em AUTENTIQUE_API_TOKEN no .env. Sem o token
// configurado, o módulo de Contratos responde com erro amigável.
const AUTENTIQUE_API_TOKEN = process.env.AUTENTIQUE_API_TOKEN || '';
const AUTENTIQUE_GRAPHQL_URL = 'https://api.autentique.com.br/v2/graphql';
if (!AUTENTIQUE_API_TOKEN) {
    console.warn('⚠️  AUTENTIQUE_API_TOKEN não definido no .env — contratos com assinatura digital ficam desativados até configurar.');
}

// Envia um documento (PDF já pronto, em Buffer) para assinatura no Autentique,
// via GraphQL multipart request (spec do Apollo Upload). Usa o fetch nativo do
// Node (18+) e FormData/Blob nativos — não exige nenhuma lib de upload extra.
async function enviarContratoParaAutentique(titulo, pdfBuffer, signatarios, nomeArquivo) {
    const query = `
        mutation CriarDocumento($document: DocumentInput!, $signers: [SignerInput!]!, $file: Upload!) {
            createDocument(document: $document, signers: $signers, file: $file) {
                id
                name
                signatures { public_id name email link { short_link } }
            }
        }
    `;
    const variables = {
        document: { name: titulo },
        signers: signatarios.map(s => ({ email: s.email, name: s.name, action: 'SIGN' })),
        file: null
    };
    const formData = new FormData();
    formData.append('operations', JSON.stringify({ query, variables }));
    formData.append('map', JSON.stringify({ '0': ['variables.file'] }));
    formData.append('0', new Blob([pdfBuffer], { type: 'application/pdf' }), nomeArquivo || `${titulo}.pdf`);

    const resposta = await fetch(AUTENTIQUE_GRAPHQL_URL, {
        method: 'POST',
        headers: { Authorization: `Bearer ${AUTENTIQUE_API_TOKEN}` },
        body: formData
    });
    const dados = await resposta.json();
    if (dados.errors && dados.errors.length) throw new Error(dados.errors[0].message);
    return dados.data.createDocument;
}

// Consulta o status atual de um documento no Autentique (quem já assinou, e o
// link do PDF assinado quando tudo estiver concluído).
async function consultarContratoNoAutentique(documentId) {
    const query = `
        query { document(id: "${documentId}") {
            id name
            files { signed }
            signatures { public_id name email link { short_link } signed { created_at } rejected { created_at } }
        } }
    `;
    const resposta = await fetch(AUTENTIQUE_GRAPHQL_URL, {
        method: 'POST',
        headers: { Authorization: `Bearer ${AUTENTIQUE_API_TOKEN}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ query })
    });
    const dados = await resposta.json();
    if (dados.errors && dados.errors.length) throw new Error(dados.errors[0].message);
    return dados.data.document;
}

// ---------- WhatsApp (Twilio) — avisos automáticos e aprovações do Master ----------
// Crie uma conta em https://www.twilio.com, ative o "WhatsApp Sandbox" (grátis
// para testar) em Messaging > Try it out > Send a WhatsApp message, e pegue o
// Account SID, Auth Token e o número do sandbox (geralmente
// "whatsapp:+14155238886"). Cada gestor que quiser receber avisos precisa
// entrar no sandbox pelo celular dele (mandando a palavra-código pro número
// do Twilio) antes de conseguir receber mensagens.
const TWILIO_ACCOUNT_SID = process.env.TWILIO_ACCOUNT_SID || '';
const TWILIO_AUTH_TOKEN = process.env.TWILIO_AUTH_TOKEN || '';
const TWILIO_WHATSAPP_FROM = process.env.TWILIO_WHATSAPP_FROM || ''; // ex: whatsapp:+14155238886
if (!TWILIO_ACCOUNT_SID || !TWILIO_AUTH_TOKEN || !TWILIO_WHATSAPP_FROM) {
    console.warn('⚠️  Twilio não configurado no .env (TWILIO_ACCOUNT_SID / TWILIO_AUTH_TOKEN / TWILIO_WHATSAPP_FROM) — avisos por WhatsApp ficam desativados até configurar.');
}

// Envia uma mensagem de WhatsApp via Twilio. Nunca lança erro para quem chamou
// (é sempre um "extra" — se o WhatsApp falhar, o resto do sistema continua
// funcionando normalmente), só registra um aviso no console.
async function enviarWhatsApp(numeroDestino, texto) {
    if (!TWILIO_ACCOUNT_SID || !TWILIO_AUTH_TOKEN || !TWILIO_WHATSAPP_FROM || !numeroDestino) return false;
    try {
        const numeroFormatado = numeroDestino.startsWith('whatsapp:') ? numeroDestino : `whatsapp:${numeroDestino}`;
        const corpo = new URLSearchParams({ From: TWILIO_WHATSAPP_FROM, To: numeroFormatado, Body: texto });
        const auth = Buffer.from(`${TWILIO_ACCOUNT_SID}:${TWILIO_AUTH_TOKEN}`).toString('base64');
        const resposta = await fetch(`https://api.twilio.com/2010-04-01/Accounts/${TWILIO_ACCOUNT_SID}/Messages.json`, {
            method: 'POST',
            headers: { Authorization: `Basic ${auth}`, 'Content-Type': 'application/x-www-form-urlencoded' },
            body: corpo.toString()
        });
        const dados = await resposta.json();
        if (!resposta.ok) { console.warn('⚠️  Falha ao enviar WhatsApp:', dados.message || resposta.statusText); return false; }
        return true;
    } catch (e) {
        console.warn('⚠️  Erro ao enviar WhatsApp:', e.message);
        return false;
    }
}

// Manda um aviso por WhatsApp para todos os admins que tiverem telefone
// cadastrado e avisos por WhatsApp habilitados.
async function avisarGestoresPorWhatsApp(texto) {
    try {
        const gestores = await dbAll(`SELECT id, whatsapp_number FROM users WHERE role = 'admin' AND whatsapp_number IS NOT NULL AND whatsapp_number != '' AND whatsapp_notifications = 1`);
        for (const g of gestores) await enviarWhatsApp(g.whatsapp_number, texto);
    } catch (e) { console.warn('⚠️  Erro ao avisar gestores por WhatsApp:', e.message); }
}

// ---------- Agente de IA (Anthropic API) ----------
// Gere uma chave em https://console.anthropic.com e coloque em
// ANTHROPIC_API_KEY no .env. Sem ela, os recursos de IA respondem com um erro
// amigável em vez de derrubar o servidor.
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY || '';
const ANTHROPIC_MODEL = process.env.ANTHROPIC_MODEL || 'claude-3-5-haiku-20241022';
if (!ANTHROPIC_API_KEY) {
    console.warn('⚠️  ANTHROPIC_API_KEY não definido no .env — o Assistente de IA fica desativado até configurar.');
}

// Chama a API da Anthropic (Claude) com um prompt de sistema e uma mensagem do
// usuário, e devolve o texto puro da resposta.
async function perguntarIA(promptSistema, promptUsuario, maxTokens = 600) {
    if (!ANTHROPIC_API_KEY) throw new Error('Assistente de IA ainda não foi configurado no servidor (defina ANTHROPIC_API_KEY no .env).');
    const resposta = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
            'x-api-key': ANTHROPIC_API_KEY,
            'anthropic-version': '2023-06-01',
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            model: ANTHROPIC_MODEL,
            max_tokens: maxTokens,
            system: promptSistema,
            messages: [{ role: 'user', content: promptUsuario }]
        })
    });
    const dados = await resposta.json();
    if (!resposta.ok) throw new Error(dados.error?.message || 'Erro ao consultar o Assistente de IA.');
    return (dados.content || []).map(bloco => bloco.text || '').join('\n').trim();
}

// Caminho do banco: se DB_PATH estiver definido (ex.: um Volume persistente
// no Railway, tipo /data/database.sqlite), usa ele — assim o banco fica FORA
// da pasta do código e sobrevive a qualquer novo deploy/upload de arquivos.
// Sem essa variável, continua funcionando como antes (arquivo ao lado do
// server.js), só que aí ele é apagado/sobrescrito a cada novo deploy.
const DB_CAMINHO_ANTIGO = path.join(__dirname, 'database.sqlite');
const dbFile = process.env.DB_PATH || DB_CAMINHO_ANTIGO;
const pastaDoBanco = path.dirname(dbFile);
if (!fs.existsSync(pastaDoBanco)) fs.mkdirSync(pastaDoBanco, { recursive: true });

// Migração automática (só acontece uma vez, na primeira subida depois de
// configurar DB_PATH): se o banco persistente ainda não existir no Volume,
// mas existir um banco antigo dentro da pasta do código (o que veio junto
// no deploy), copia ele para dentro do Volume ANTES de abrir a conexão — sem
// isso, ligar o Volume faria o sistema começar do zero, como se todo mundo
// tivesse sido apagado.
if (process.env.DB_PATH && dbFile !== DB_CAMINHO_ANTIGO && !fs.existsSync(dbFile) && fs.existsSync(DB_CAMINHO_ANTIGO)) {
    fs.copyFileSync(DB_CAMINHO_ANTIGO, dbFile);
    console.log(`♻️  Banco existente migrado automaticamente para o Volume persistente: ${dbFile}`);
}

console.log(`📂 Usando banco de dados em: ${dbFile}${process.env.DB_PATH ? ' (persistente, via DB_PATH)' : ' (⚠️ dentro da pasta do código — configure DB_PATH com um Volume para não perder dados a cada deploy)'}`);
const db = new sqlite3.Database(dbFile, (err) => {
    if (err) {
        console.error('Erro na base de dados:', err.message);
    } else {
        console.log('✅ Base de dados Impulsionar V4 ativa com Notificações.');
        inicializarBase();
    }
});

function inicializarBase() {
    db.serialize(() => {
        db.run(`CREATE TABLE IF NOT EXISTS companies (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            cnpj TEXT UNIQUE NOT NULL,
            segment TEXT,
            phone TEXT,
            address TEXT,
            logo_url TEXT,
            document_type TEXT DEFAULT 'cnpj',
            cep TEXT,
            street TEXT,
            address_number TEXT,
            neighborhood TEXT,
            city TEXT,
            state TEXT,
            company_size TEXT
        )`);
        // Migração leve para empresas cadastradas antes destas colunas existirem.
        ['document_type TEXT DEFAULT \'cnpj\'', 'cep TEXT', 'street TEXT', 'address_number TEXT',
         'neighborhood TEXT', 'city TEXT', 'state TEXT', 'company_size TEXT',
         // enabled_modules: JSON com as chaves liberadas pelo Master conforme o
         // contrato fechado (NULL = sem restrição, libera tudo — comportamento
         // anterior, preservado para empresas já cadastradas). vaga_credito_dias:
         // saldo de dias de divulgação de vaga que o Master concede como bônus,
         // consumido automaticamente na aprovação em vez de cobrar pelo Mercado Pago.
         'enabled_modules TEXT', 'vaga_credito_dias INTEGER DEFAULT 0',
         // Data da Auditoria Oficial Ambev (o evento real de auditoria, marcado
         // pelo Master) e uma observação livre sobre esse agendamento.
         'dpo_auditoria_oficial_data TEXT', 'dpo_auditoria_oficial_nota TEXT'].forEach(coluna => {
            db.run(`ALTER TABLE companies ADD COLUMN ${coluna}`, () => {});
        });

        db.run(`CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            email TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            company_id INTEGER,
            employee_id INTEGER,
            mentor_id INTEGER,
            role TEXT DEFAULT 'client_admin',
            reset_token TEXT,
            reset_token_expires TEXT,
            FOREIGN KEY(company_id) REFERENCES companies(id),
            FOREIGN KEY(employee_id) REFERENCES employees(id),
            FOREIGN KEY(mentor_id) REFERENCES mentors(id)
        )`);
        // Migração leve para bancos criados antes destas colunas existirem
        // (ignora o erro "duplicate column" quando a coluna já existe).
        db.run(`ALTER TABLE users ADD COLUMN reset_token TEXT`, () => {});
        db.run(`ALTER TABLE users ADD COLUMN reset_token_expires TEXT`, () => {});
        db.run(`ALTER TABLE users ADD COLUMN mentor_id INTEGER`, () => {});
        db.run(`ALTER TABLE users ADD COLUMN whatsapp_number TEXT`, () => {});
        db.run(`ALTER TABLE users ADD COLUMN whatsapp_notifications INTEGER DEFAULT 0`, () => {});
        // Permissão de módulos INDIVIDUAL por acesso (client_admin) dentro da
        // empresa — antes só existia companies.enabled_modules (um valor único
        // pra todo mundo da empresa). Agora cada usuário pode ter sua própria
        // restrição; NULL = sem restrição própria, cai no padrão da empresa
        // (companies.enabled_modules), que por sua vez, se também NULL, libera tudo.
        db.run(`ALTER TABLE users ADD COLUMN enabled_modules TEXT`, () => {});

        // Dados da própria Impulsionar (usados nos contratos como CONTRATADA e
        // nas automações). Só existe uma linha, com id fixo = 1.
        db.run(`CREATE TABLE IF NOT EXISTS platform_profile (
            id INTEGER PRIMARY KEY CHECK (id = 1),
            name TEXT DEFAULT 'Impulsionar Consultoria',
            cnpj TEXT,
            phone TEXT,
            email TEXT,
            address TEXT,
            logo_url TEXT
        )`);
        db.run(`INSERT OR IGNORE INTO platform_profile (id, name) VALUES (1, 'Impulsionar Consultoria')`);

        // Painel de automação: o que roda sozinho x o que precisa de aprovação
        // do Master. Guardado como chave/valor simples para ficar fácil de
        // adicionar novas chaves no futuro sem migração de schema.
        db.run(`CREATE TABLE IF NOT EXISTS automation_settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )`);
        const CHAVES_AUTOMACAO_PADRAO = {
            auto_approve_resumes: '0',        // currículo novo já entra aprovado, sem revisão do Master
            ai_support_autopilot: '0',        // IA responde suporte sozinha, sem esperar aprovação no WhatsApp
            ai_resume_review_autopilot: '0'   // IA decide aprovação/ajuste de currículo sozinha
        };
        Object.entries(CHAVES_AUTOMACAO_PADRAO).forEach(([k, v]) => {
            db.run(`INSERT OR IGNORE INTO automation_settings (key, value) VALUES (?, ?)`, [k, v]);
        });

        // Credenciais de integrações (Mercado Pago etc.) configuráveis pelo Master
        // direto na tela "Meu Perfil", sem precisar editar o .env nem reiniciar o
        // servidor. Se não houver nada salvo aqui, o servidor usa o .env como
        // ponto de partida (comportamento anterior, preservado).
        db.run(`CREATE TABLE IF NOT EXISTS integration_settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )`);
        db.get(`SELECT value FROM integration_settings WHERE key = 'mp_access_token'`, [], (err, row) => {
            if (!err && row && row.value) {
                configurarMercadoPago(row.value);
                console.log('✅ Mercado Pago configurado a partir do banco de dados (Meu Perfil > Mercado Pago).');
            }
        });
        db.get(`SELECT value FROM integration_settings WHERE key = 'app_base_url'`, [], (err, row) => {
            if (!err && row && row.value) {
                appBaseUrlAtiva = row.value;
                console.log('✅ URL pública do sistema configurada a partir do banco de dados: ' + appBaseUrlAtiva);
            }
        });

        // Ações que a IA propôs e que estão aguardando (ou já receberam) a
        // aprovação do Master pelo WhatsApp.
        db.run(`CREATE TABLE IF NOT EXISTS ai_pending_actions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            type TEXT NOT NULL,
            candidate_user_id INTEGER,
            demand TEXT,
            proposal TEXT NOT NULL,
            suggested_status TEXT,
            status TEXT DEFAULT 'pendente',
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            resolved_at DATETIME,
            FOREIGN KEY(candidate_user_id) REFERENCES users(id)
        )`);

        db.run(`CREATE TABLE IF NOT EXISTS employees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER,
            name TEXT NOT NULL,
            role TEXT,
            email TEXT,
            phone TEXT,
            performance_level TEXT DEFAULT 'Em Desenvolvimento',
            executive_phase TEXT DEFAULT 'Fase 1: Diagnóstico',
            progress_percentage INTEGER DEFAULT 10,
            disc_profile TEXT DEFAULT 'A definir',
            photo_url TEXT,
            FOREIGN KEY(company_id) REFERENCES companies(id)
        )`);
        // current_challenge: o desafio atual do executivo nesta fase.
        // target_role: para qual vaga/posição a empresa está preparando essa pessoa.
        db.run(`ALTER TABLE employees ADD COLUMN current_challenge TEXT`, () => {});
        db.run(`ALTER TABLE employees ADD COLUMN target_role TEXT`, () => {});
        // Permissões de acesso do PRÓPRIO colaborador (login individual, role
        // 'autonomous'): JSON com os módulos liberados para ele, no mesmo
        // formato de companies.enabled_modules. NULL = sem restrição definida
        // (mantém o comportamento anterior para acessos já existentes antes
        // deste recurso). Um acesso novo criado a partir de agora começa
        // restrito só a PDI — quem cadastrou (client_admin ou Master) libera o
        // resto depois, na tela de Permissões do colaborador.
        db.run(`ALTER TABLE employees ADD COLUMN enabled_modules TEXT`, () => {});

        // Histórico de mudanças de fase do Pipeline de Desenvolvimento: registra o
        // motivo da movimentação e a frase que aparece para o colaborador, para dar
        // rastreabilidade a cada passo da jornada (não só a foto atual).
        db.run(`CREATE TABLE IF NOT EXISTS employee_phase_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id INTEGER NOT NULL,
            from_phase TEXT,
            to_phase TEXT NOT NULL,
            reason TEXT,
            message_to_employee TEXT,
            changed_by_name TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(employee_id) REFERENCES employees(id)
        )`);

        // Ações de um PDI: cada plano de desenvolvimento pode ter várias ações
        // (não iniciada / em andamento / concluída), em vez de um status único.
        db.run(`CREATE TABLE IF NOT EXISTS pdi_actions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            pd_plan_id INTEGER NOT NULL,
            description TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'Não iniciada',
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(pd_plan_id) REFERENCES pd_plans(id)
        )`);

        db.run(`CREATE TABLE IF NOT EXISTS pd_plans (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id INTEGER,
            objective TEXT NOT NULL,
            action_plan TEXT NOT NULL,
            deadline TEXT,
            status TEXT DEFAULT 'Em Andamento',
            FOREIGN KEY(employee_id) REFERENCES employees(id)
        )`);

        // Metas mensais do colaborador: até 5 por colaborador, cada uma podendo ser
        // medida em dinheiro (R$), número (unidade) ou percentual (%).
        db.run(`CREATE TABLE IF NOT EXISTS employee_goals (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id INTEGER NOT NULL,
            title TEXT NOT NULL,
            month TEXT NOT NULL,
            goal_type TEXT NOT NULL DEFAULT 'numero',
            target_value REAL DEFAULT 0,
            achieved_value REAL DEFAULT 0,
            status TEXT DEFAULT 'Em Andamento',
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(employee_id) REFERENCES employees(id)
        )`);
        db.run(`ALTER TABLE employee_goals ADD COLUMN updated_at DATETIME`, () => {});

        db.run(`CREATE TABLE IF NOT EXISTS video_lessons (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            module_name TEXT,
            video_url TEXT NOT NULL,
            duration TEXT,
            description TEXT,
            thumbnail_url TEXT,
            posted_by TEXT,
            mentor_id INTEGER,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(mentor_id) REFERENCES mentors(id)
        )`);
        db.run(`ALTER TABLE video_lessons ADD COLUMN description TEXT`, () => {});
        db.run(`ALTER TABLE video_lessons ADD COLUMN thumbnail_url TEXT`, () => {});
        db.run(`ALTER TABLE video_lessons ADD COLUMN posted_by TEXT`, () => {});
        db.run(`ALTER TABLE video_lessons ADD COLUMN mentor_id INTEGER`, () => {});
        db.run(`ALTER TABLE video_lessons ADD COLUMN created_at DATETIME`, () => {});

        // Painel social da Academy: curtidas e comentários por videoaula
        db.run(`CREATE TABLE IF NOT EXISTS video_likes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            video_id INTEGER NOT NULL,
            user_id INTEGER NOT NULL,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(video_id, user_id),
            FOREIGN KEY(video_id) REFERENCES video_lessons(id),
            FOREIGN KEY(user_id) REFERENCES users(id)
        )`);

        db.run(`CREATE TABLE IF NOT EXISTS video_comments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            video_id INTEGER NOT NULL,
            user_id INTEGER NOT NULL,
            comment TEXT NOT NULL,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(video_id) REFERENCES video_lessons(id),
            FOREIGN KEY(user_id) REFERENCES users(id)
        )`);

        db.run(`CREATE TABLE IF NOT EXISTS assessments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id INTEGER,
            leadership_competence TEXT,
            score INTEGER,
            status TEXT,
            date DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(employee_id) REFERENCES employees(id)
        )`);

        db.run(`CREATE TABLE IF NOT EXISTS mentors (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            specialty TEXT,
            email TEXT,
            available_days TEXT,
            resume TEXT,
            bio_video_url TEXT,
            photo_url TEXT
        )`);

        db.run(`CREATE TABLE IF NOT EXISTS mentorships (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id INTEGER,
            mentor_id INTEGER,
            mentor_name TEXT,
            meeting_date TEXT NOT NULL,
            topics TEXT NOT NULL,
            minutes TEXT,
            status TEXT DEFAULT 'Agendada',
            FOREIGN KEY(employee_id) REFERENCES employees(id),
            FOREIGN KEY(mentor_id) REFERENCES mentors(id)
        )`);
        // Bancos antigos têm "mentor_name" como coluna própria (NOT NULL em alguns casos);
        // mantemos ela sempre preenchida junto com mentor_id para evitar erro de gravação.
        db.run(`ALTER TABLE mentorships ADD COLUMN mentor_name TEXT`, () => {});
        // Duração, participantes extras (internos já cadastrados + externos por e-mail)
        // e controle do convite de calendário (mesmo UID + SEQUENCE crescente faz o
        // Outlook/Gmail do convidado atualizar o evento já existente em vez de duplicar).
        db.run(`ALTER TABLE mentorships ADD COLUMN duration_minutes INTEGER DEFAULT 60`, () => {});
        db.run(`ALTER TABLE mentorships ADD COLUMN participants_json TEXT DEFAULT '[]'`, () => {});
        db.run(`ALTER TABLE mentorships ADD COLUMN ics_uid TEXT`, () => {});
        db.run(`ALTER TABLE mentorships ADD COLUMN ics_sequence INTEGER DEFAULT 0`, () => {});
        // Código aleatório da sala de videochamada — usado no link público que dá
        // acesso a convidados sem login (só quem tem o link entra; o id numérico
        // sozinho não seria suficiente porque é sequencial e fácil de adivinhar).
        db.run(`ALTER TABLE mentorships ADD COLUMN room_token TEXT`, () => {});

        // Horários disponíveis de cada mentor, por data específica do calendário
        // (ex: 14/10/2026 das 09:00 às 10:00), usados para só permitir agendar
        // mentorias em janelas que o mentor realmente cadastrou naquele dia.
        db.run(`CREATE TABLE IF NOT EXISTS mentor_availability (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            mentor_id INTEGER NOT NULL,
            specific_date TEXT,
            day_of_week INTEGER,
            start_time TEXT NOT NULL,
            end_time TEXT NOT NULL,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(mentor_id) REFERENCES mentors(id)
        )`);
        db.run(`ALTER TABLE mentor_availability ADD COLUMN specific_date TEXT`, () => {});

        // Módulo de Consultoria: planos de acompanhamento separados do PDI
        // (o PDI é o plano do próprio executivo; o "plano de acompanhamento"
        // é o engajamento da consultoria com a empresa ou com o executivo,
        // com marcos/etapas próprias e datas que entram no calendário).
        db.run(`CREATE TABLE IF NOT EXISTS consulting_plans (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER,
            employee_id INTEGER,
            title TEXT NOT NULL,
            objective TEXT,
            consultant_name TEXT,
            start_date TEXT,
            end_date TEXT,
            status TEXT DEFAULT 'Em Andamento',
            FOREIGN KEY(company_id) REFERENCES companies(id),
            FOREIGN KEY(employee_id) REFERENCES employees(id)
        )`);

        db.run(`CREATE TABLE IF NOT EXISTS consulting_milestones (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            plan_id INTEGER NOT NULL,
            title TEXT NOT NULL,
            due_date TEXT,
            status TEXT DEFAULT 'Pendente',
            FOREIGN KEY(plan_id) REFERENCES consulting_plans(id)
        )`);
        // Fluxo de solicitação: quando a empresa (client_admin) pede um plano de
        // acompanhamento, ele nasce como 'Solicitado' e só o Master pode aprovar
        // (vira 'Em Andamento') ou recusar (vira 'Recusado', com motivo).
        db.run(`ALTER TABLE consulting_plans ADD COLUMN rejection_reason TEXT`, () => {});
        db.run(`ALTER TABLE consulting_plans ADD COLUMN requested_by_company INTEGER DEFAULT 0`, () => {});

        // ============================================================
        // EVOLUÇÃO "ESTILO G4": trilhas de curso, certificados, comunidade,
        // banco de currículos, eventos, planos por empresa e gamificação.
        // ============================================================

        // ---------- Identidade/rede/currículo no perfil do executivo ----------
        db.run(`ALTER TABLE employees ADD COLUMN public_bio TEXT`, () => {});
        db.run(`ALTER TABLE employees ADD COLUMN linkedin_url TEXT`, () => {});
        db.run(`ALTER TABLE employees ADD COLUMN show_in_directory INTEGER DEFAULT 0`, () => {});
        db.run(`ALTER TABLE employees ADD COLUMN resume_url TEXT`, () => {});
        db.run(`ALTER TABLE employees ADD COLUMN looking_for_opportunity INTEGER DEFAULT 0`, () => {});
        db.run(`ALTER TABLE employees ADD COLUMN desired_role TEXT`, () => {});

        // ---------- Trilhas de curso (Academy organizada em módulos sequenciais) ----------
        db.run(`CREATE TABLE IF NOT EXISTS learning_tracks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            description TEXT,
            cover_image_url TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )`);

        db.run(`CREATE TABLE IF NOT EXISTS track_lessons (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            track_id INTEGER NOT NULL,
            video_id INTEGER NOT NULL,
            order_index INTEGER DEFAULT 0,
            FOREIGN KEY(track_id) REFERENCES learning_tracks(id),
            FOREIGN KEY(video_id) REFERENCES video_lessons(id),
            UNIQUE(track_id, video_id)
        )`);

        db.run(`CREATE TABLE IF NOT EXISTS track_progress (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            track_id INTEGER NOT NULL,
            video_id INTEGER NOT NULL,
            completed_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(user_id, track_id, video_id),
            FOREIGN KEY(user_id) REFERENCES users(id),
            FOREIGN KEY(track_id) REFERENCES learning_tracks(id),
            FOREIGN KEY(video_id) REFERENCES video_lessons(id)
        )`);

        // ---------- Certificados de conclusão de trilha ----------
        db.run(`CREATE TABLE IF NOT EXISTS certificates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            track_id INTEGER NOT NULL,
            certificate_code TEXT UNIQUE,
            issued_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(user_id, track_id),
            FOREIGN KEY(user_id) REFERENCES users(id),
            FOREIGN KEY(track_id) REFERENCES learning_tracks(id)
        )`);

        // ---------- Eventos e encontros (masterclasses, presenciais, etc) ----------
        db.run(`CREATE TABLE IF NOT EXISTS events (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            description TEXT,
            event_date TEXT NOT NULL,
            event_time TEXT,
            is_online INTEGER DEFAULT 1,
            location TEXT,
            link TEXT,
            capacity INTEGER,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )`);

        db.run(`CREATE TABLE IF NOT EXISTS event_registrations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            event_id INTEGER NOT NULL,
            user_id INTEGER NOT NULL,
            registered_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(event_id, user_id),
            FOREIGN KEY(event_id) REFERENCES events(id),
            FOREIGN KEY(user_id) REFERENCES users(id)
        )`);

        // ---------- Lista de espera: quando o evento está lotado, o usuário entra
        // na fila e recebe um e-mail automático assim que uma vaga se abrir ----------
        db.run(`CREATE TABLE IF NOT EXISTS event_waitlist (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            event_id INTEGER NOT NULL,
            user_id INTEGER NOT NULL,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(event_id, user_id),
            FOREIGN KEY(event_id) REFERENCES events(id),
            FOREIGN KEY(user_id) REFERENCES users(id)
        )`);

        // ---------- Planos por empresa (limite de executivos, etc) ----------
        db.run(`CREATE TABLE IF NOT EXISTS plans (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            max_employees INTEGER,
            price_display TEXT,
            features_text TEXT
        )`);
        db.run(`ALTER TABLE companies ADD COLUMN plan_id INTEGER REFERENCES plans(id)`, () => {});

        // ---------- Assinaturas via Mercado Pago (preapproval recorrente) ----------
        db.run(`ALTER TABLE plans ADD COLUMN mp_price REAL`, () => {});
        db.run(`ALTER TABLE plans ADD COLUMN trial_days INTEGER DEFAULT 0`, () => {});
        db.run(`ALTER TABLE companies ADD COLUMN mp_preapproval_id TEXT`, () => {});
        db.run(`ALTER TABLE companies ADD COLUMN subscription_status TEXT DEFAULT 'none'`, () => {});
        db.run(`ALTER TABLE companies ADD COLUMN subscription_updated_at DATETIME`, () => {});
        db.run(`ALTER TABLE companies ADD COLUMN trial_ends_at TEXT`, () => {});
        // Plano "pendente de pagamento": quando a empresa clica pra aderir/trocar de
        // plano pela tela "Planos Disponíveis" (self-service), o plan_id só é
        // aplicado de verdade (liberando os créditos de funcionário) depois que o
        // Mercado Pago confirma o pagamento. Até lá, fica guardado aqui.
        db.run(`ALTER TABLE companies ADD COLUMN pending_plan_id INTEGER`, () => {});
        // Pedido de cancelamento: a empresa não cancela a assinatura sozinha,
        // apenas solicita — fica marcado aqui até o Master analisar e confirmar
        // (ou recusar) o cancelamento de fato.
        db.run(`ALTER TABLE companies ADD COLUMN cancellation_requested_at TEXT`, () => {});

        db.get(`SELECT COUNT(*) as total FROM plans`, [], (err, row) => {
            if (!err && row && row.total === 0) {
                db.run(`INSERT INTO plans (name, max_employees, price_display, features_text) VALUES
                    ('Starter', 5, 'R$ 1.500/mês', 'Até 5 executivos · Academy Digital · PDI · Competence Check'),
                    ('Professional', 20, 'R$ 4.500/mês', 'Até 20 executivos · tudo do Starter · Mentorias ilimitadas · Eventos'),
                    ('Enterprise', NULL, 'Sob consulta', 'Executivos ilimitados · tudo do Professional · Consultoria dedicada')
                `);
            }
        });

        // ---------- Gamificação: ledger de pontos ----------
        db.run(`CREATE TABLE IF NOT EXISTS points_ledger (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            points INTEGER NOT NULL,
            reason TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(user_id) REFERENCES users(id)
        )`);

        // ---------- Notificações in-app (sininho) ----------
        db.run(`CREATE TABLE IF NOT EXISTS notifications (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            title TEXT NOT NULL,
            message TEXT,
            link TEXT,
            is_read INTEGER DEFAULT 0,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(user_id) REFERENCES users(id)
        )`);

        // ---------- Portal público de vagas e currículos ----------
        // Candidatos são usuários com role='candidate' na tabela users já
        // existente (reaproveita todo o login/JWT/auth). Este perfil guarda os
        // dados de currículo em si.
        db.run(`CREATE TABLE IF NOT EXISTS candidate_profiles (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL UNIQUE,
            phone TEXT,
            desired_role TEXT,
            city TEXT,
            bio TEXT,
            skills TEXT,
            linkedin_url TEXT,
            resume_url TEXT,
            status TEXT DEFAULT 'pending',
            plan TEXT DEFAULT 'gratuito',
            photo_url TEXT,
            gender TEXT,
            education_level TEXT,
            languages TEXT,
            first_job INTEGER DEFAULT 0,
            experiences_json TEXT,
            desired_states TEXT,
            desired_cities TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(user_id) REFERENCES users(id)
        )`);
        // Colunas do perfil completo de currículo (auto-upgrade de bases já existentes).
        ['photo_url TEXT', 'gender TEXT', 'education_level TEXT', 'languages TEXT', 'first_job INTEGER DEFAULT 0',
         'experiences_json TEXT', 'desired_states TEXT', 'desired_cities TEXT'].forEach(coluna => {
            db.run(`ALTER TABLE candidate_profiles ADD COLUMN ${coluna}`, () => {});
        });

        // Fio de interação (chat) entre o candidato e o Master sobre o currículo dele —
        // usado para o Master pedir ajustes e o candidato responder.
        db.run(`CREATE TABLE IF NOT EXISTS candidate_messages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            candidate_user_id INTEGER NOT NULL,
            sender TEXT NOT NULL,
            message TEXT NOT NULL,
            topic TEXT DEFAULT 'curriculo',
            read_by_master INTEGER DEFAULT 0,
            read_by_candidate INTEGER DEFAULT 0,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(candidate_user_id) REFERENCES users(id)
        )`);
        db.run(`ALTER TABLE candidate_messages ADD COLUMN topic TEXT DEFAULT 'curriculo'`, () => {});

        // Planos de cobrança por vaga divulgada (dias de duração x preço),
        // configurados pelo Master em "Planos de Vaga".
        db.run(`CREATE TABLE IF NOT EXISTS vaga_plans (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            label TEXT,
            days INTEGER NOT NULL,
            price REAL NOT NULL,
            active INTEGER DEFAULT 1
        )`);

        // Taxa de sucesso (fechamento) por função — cobrada da empresa quando
        // a vaga é fechada COM contratação, além do valor de divulgação. O
        // Master cadastra a lista de funções e o valor de cada uma em
        // "Planos de Fechamento"; a empresa escolhe a função ao criar a vaga.
        db.run(`CREATE TABLE IF NOT EXISTS closing_fee_plans (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            function_label TEXT NOT NULL,
            price REAL NOT NULL,
            active INTEGER DEFAULT 1
        )`);

        // Vagas publicadas pelas empresas no portal público.
        db.run(`CREATE TABLE IF NOT EXISTS job_postings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER NOT NULL,
            title TEXT NOT NULL,
            description TEXT,
            location TEXT,
            state TEXT,
            is_remote INTEGER DEFAULT 0,
            seniority TEXT,
            salary_range TEXT,
            vaga_plan_id INTEGER,
            status TEXT DEFAULT 'pending_payment',
            mp_preference_id TEXT,
            mp_payment_id TEXT,
            checkout_url TEXT,
            published_at DATETIME,
            expires_at DATETIME,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(company_id) REFERENCES companies(id),
            FOREIGN KEY(vaga_plan_id) REFERENCES vaga_plans(id)
        )`);
        db.run(`ALTER TABLE job_postings ADD COLUMN state TEXT`, () => {});
        // Fluxo de aprovação: a vaga criada pela empresa fica 'pendente_aprovacao'
        // até o Master aprovar (só então segue para pagamento/publicação) ou
        // rejeitar (com motivo, para a empresa entender o que ajustar).
        db.run(`ALTER TABLE job_postings ADD COLUMN approved_by INTEGER`, () => {});
        db.run(`ALTER TABLE job_postings ADD COLUMN approved_at DATETIME`, () => {});
        db.run(`ALTER TABLE job_postings ADD COLUMN rejection_reason TEXT`, () => {});
        db.run(`ALTER TABLE job_postings ADD COLUMN paid_with_credit INTEGER DEFAULT 0`, () => {});
        // Detalhes completos da vaga, pedidos pela empresa no cadastro: formação,
        // idiomas e requisitos ajudam o candidato a entender se o perfil bate;
        // benefícios é opcional (nem toda empresa quer/pode informar). photo_url
        // é a imagem que aparece no topo do card de divulgação no portal.
        db.run(`ALTER TABLE job_postings ADD COLUMN education TEXT`, () => {});
        db.run(`ALTER TABLE job_postings ADD COLUMN languages TEXT`, () => {});
        db.run(`ALTER TABLE job_postings ADD COLUMN requirements TEXT`, () => {});
        db.run(`ALTER TABLE job_postings ADD COLUMN responsibilities TEXT`, () => {});
        db.run(`ALTER TABLE job_postings ADD COLUMN benefits TEXT`, () => {});
        db.run(`ALTER TABLE job_postings ADD COLUMN photo_url TEXT`, () => {});
        // "Vagas Ofertadas": fechar uma vaga (contratação concluída ou vaga
        // cancelada) exige justificativa e/ou vincular o candidato contratado;
        // exclusão passa a ser reversível (soft-delete) para existir a aba
        // "Vagas Excluídas" sem perder o histórico.
        db.run(`ALTER TABLE job_postings ADD COLUMN closed_reason TEXT`, () => {});
        db.run(`ALTER TABLE job_postings ADD COLUMN closed_application_id INTEGER`, () => {});
        db.run(`ALTER TABLE job_postings ADD COLUMN closed_at DATETIME`, () => {});
        db.run(`ALTER TABLE job_postings ADD COLUMN deleted_at DATETIME`, () => {});
        // Estorno do pagamento único da divulgação (Master decide reembolsar a
        // empresa) — mantém o pagamento original (mp_payment_id) como histórico
        // e só registra quando/quem estornou.
        db.run(`ALTER TABLE job_postings ADD COLUMN refunded_at DATETIME`, () => {});
        db.run(`ALTER TABLE job_postings ADD COLUMN refunded_by INTEGER`, () => {});
        // Taxa de fechamento (sucesso na contratação), cobrada à parte da
        // divulgação — a empresa escolhe a função no cadastro da vaga, e a
        // cobrança é gerada automaticamente quando a vaga é fechada COM
        // candidato contratado (nunca ao só cancelar/encerrar sem contratar).
        db.run(`ALTER TABLE job_postings ADD COLUMN closing_fee_plan_id INTEGER`, () => {});
        db.run(`ALTER TABLE job_postings ADD COLUMN closing_fee_status TEXT`, () => {}); // null | 'pending_payment' | 'paid'
        db.run(`ALTER TABLE job_postings ADD COLUMN closing_fee_mp_preference_id TEXT`, () => {});
        db.run(`ALTER TABLE job_postings ADD COLUMN closing_fee_payment_id TEXT`, () => {});
        db.run(`ALTER TABLE job_postings ADD COLUMN closing_fee_checkout_url TEXT`, () => {});
        db.run(`ALTER TABLE job_postings ADD COLUMN closing_fee_paid_at DATETIME`, () => {});

        db.run(`CREATE TABLE IF NOT EXISTS job_applications (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            job_posting_id INTEGER NOT NULL,
            candidate_user_id INTEGER NOT NULL,
            applied_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(job_posting_id, candidate_user_id),
            FOREIGN KEY(job_posting_id) REFERENCES job_postings(id),
            FOREIGN KEY(candidate_user_id) REFERENCES users(id)
        )`);

        // Curtidas dos candidatos nas vagas do portal — só um "like" por
        // candidato por vaga (o UNIQUE garante isso), pra medir engajamento.
        db.run(`CREATE TABLE IF NOT EXISTS job_posting_likes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            job_posting_id INTEGER NOT NULL,
            candidate_user_id INTEGER NOT NULL,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(job_posting_id, candidate_user_id),
            FOREIGN KEY(job_posting_id) REFERENCES job_postings(id),
            FOREIGN KEY(candidate_user_id) REFERENCES users(id)
        )`);

        // ---------- Contratos com assinatura digital (via Autentique) ----------
        db.run(`CREATE TABLE IF NOT EXISTS contracts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            type TEXT NOT NULL,
            company_id INTEGER,
            employee_id INTEGER,
            mentor_id INTEGER,
            title TEXT NOT NULL,
            content TEXT NOT NULL,
            status TEXT DEFAULT 'enviado',
            autentique_document_id TEXT,
            signers_json TEXT,
            signed_file_url TEXT,
            created_by INTEGER,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(company_id) REFERENCES companies(id),
            FOREIGN KEY(employee_id) REFERENCES employees(id),
            FOREIGN KEY(mentor_id) REFERENCES mentors(id),
            FOREIGN KEY(created_by) REFERENCES users(id)
        )`);

        // ---------- DPO AMBEV: consultoria de processos por pilar ----------
        // Preço de cada pilar (Master define em "DPO Ambev > Preços"). O preço
        // da "Consultoria Completa" (todos os pilares de uma vez) fica guardado
        // em integration_settings (chave dpo_full_audit_price), reaproveitando
        // a mesma tabelinha de configurações que já existe pro Mercado Pago.
        db.run(`CREATE TABLE IF NOT EXISTS dpo_pillar_prices (
            pillar_key TEXT PRIMARY KEY,
            price REAL DEFAULT 0,
            active INTEGER DEFAULT 1
        )`);
        DPO_PILARES_ORDEM.forEach(chave => {
            db.run(`INSERT OR IGNORE INTO dpo_pillar_prices (pillar_key, price, active) VALUES (?, 0, 1)`, [chave]);
        });

        // Consultores cadastrados pelo Master — a empresa escolhe um deles na
        // hora de comprar um pilar ou a Consultoria Completa.
        db.run(`CREATE TABLE IF NOT EXISTS dpo_consultants (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            email TEXT,
            phone TEXT,
            bio TEXT,
            active INTEGER DEFAULT 1,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )`);

        // Compra de um pilar avulso ou da consultoria completa por uma empresa —
        // cobrança única via Mercado Pago (Checkout Pro), no mesmo padrão já
        // usado para divulgação de vagas e taxa de fechamento.
        db.run(`CREATE TABLE IF NOT EXISTS dpo_purchases (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER NOT NULL,
            scope TEXT NOT NULL,
            pillar_key TEXT,
            price REAL,
            status TEXT DEFAULT 'pending_payment',
            mp_preference_id TEXT,
            mp_payment_id TEXT,
            checkout_url TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            paid_at DATETIME,
            FOREIGN KEY(company_id) REFERENCES companies(id)
        )`);
        // Consultor escolhido pela empresa no momento da compra.
        db.run(`ALTER TABLE dpo_purchases ADD COLUMN consultant_id INTEGER REFERENCES dpo_consultants(id)`, () => {});
        // Período de teste concedido pelo Master (sem cobrança): is_trial=1 e
        // trial_expires_at marcam a liberação temporária; passada a data, o
        // pilar deixa de contar como ativo automaticamente (ver pilaresAtivosDaEmpresa).
        ['is_trial INTEGER DEFAULT 0', 'trial_expires_at DATETIME', 'granted_by INTEGER'].forEach(coluna => {
            db.run(`ALTER TABLE dpo_purchases ADD COLUMN ${coluna}`, () => {});
        });

        // Ciclo de autoavaliação: o Master cria um novo ciclo (mensal, "quando
        // quiser") para a empresa, com os pilares que ela já comprou. É dentro
        // do ciclo que a empresa preenche as respostas daquele mês — assim dá
        // pra comparar a evolução mês a mês olhando os ciclos anteriores.
        db.run(`CREATE TABLE IF NOT EXISTS dpo_audit_cycles (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER NOT NULL,
            referencia TEXT,
            pilares TEXT NOT NULL,
            scheduled_at DATETIME,
            status TEXT DEFAULT 'agendado',
            created_by INTEGER,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            closed_at DATETIME,
            FOREIGN KEY(company_id) REFERENCES companies(id)
        )`);

        // Pontuação marcada pela empresa em cada pergunta, dentro de um ciclo.
        // question_key = "<pilar>:<numero da pergunta>", ex.: "seguranca:1.1".
        db.run(`CREATE TABLE IF NOT EXISTS dpo_answers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cycle_id INTEGER NOT NULL,
            question_key TEXT NOT NULL,
            score INTEGER,
            updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            updated_by INTEGER,
            UNIQUE(cycle_id, question_key),
            FOREIGN KEY(cycle_id) REFERENCES dpo_audit_cycles(id)
        )`);

        // Plano de ação de uma pergunta específica do ciclo, com um ou mais
        // "follows" (Follow 1, Follow 2...) de acompanhamento.
        db.run(`CREATE TABLE IF NOT EXISTS dpo_action_plans (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cycle_id INTEGER NOT NULL,
            question_key TEXT NOT NULL,
            texto TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            created_by INTEGER,
            FOREIGN KEY(cycle_id) REFERENCES dpo_audit_cycles(id)
        )`);
        // Campos adicionais do plano de ação: número da verificação (dentro da
        // lista de verificações da pergunta) a que o plano se refere, dono da
        // ação e status dela — pedidos para dar mais rastreabilidade a cada plano.
        ['verificacao_numero TEXT', 'owner TEXT', "status TEXT DEFAULT 'nao_iniciada'"].forEach(coluna => {
            db.run(`ALTER TABLE dpo_action_plans ADD COLUMN ${coluna}`, () => {});
        });
        db.run(`CREATE TABLE IF NOT EXISTS dpo_follow_ups (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            action_plan_id INTEGER NOT NULL,
            numero INTEGER,
            texto TEXT,
            data_prevista TEXT,
            status TEXT DEFAULT 'pendente',
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(action_plan_id) REFERENCES dpo_action_plans(id)
        )`);

        db.get(`SELECT COUNT(*) as total FROM vaga_plans`, [], (err, row) => {
            if (!err && row && row.total === 0) {
                db.run(`INSERT INTO vaga_plans (label, days, price, active) VALUES
                    ('Básico — 15 dias', 15, 99.90, 1),
                    ('Padrão — 30 dias', 30, 179.90, 1),
                    ('Destaque — 60 dias', 60, 299.90, 1)
                `);
            }
        });

        // Conta Master padrão. Se o banco ainda não tiver nenhum admin, cria com
        // esse login; se já existir a conta antiga (admin@impulsionar.com, com a
        // senha padrão de fábrica), ela é atualizada automaticamente para o novo
        // e-mail/senha definidos aqui — assim não fica duas contas Master soltas.
        const EMAIL_MASTER_PADRAO = 'master@impulsionarv4.com.br';
        const SENHA_MASTER_PADRAO = 'Impulsionar@v4';
        db.get(`SELECT * FROM users WHERE role = 'admin' ORDER BY id ASC LIMIT 1`, [], async (err, row) => {
            if (!row) {
                const hashedPassword = await bcrypt.hash(SENHA_MASTER_PADRAO, 10);
                db.run(`INSERT INTO users (name, email, password, company_id, role) VALUES (?, ?, ?, NULL, ?)`,
                    ['Board Executivo Master', EMAIL_MASTER_PADRAO, hashedPassword, 'admin']
                );
            } else if (row.email === 'admin@impulsionar.com') {
                const hashedPassword = await bcrypt.hash(SENHA_MASTER_PADRAO, 10);
                db.run(`UPDATE users SET email = ?, password = ? WHERE id = ?`, [EMAIL_MASTER_PADRAO, hashedPassword, row.id]);
            }
        });
    });
}

// Helpers em Promise (facilitam o dashboard, que precisa combinar várias consultas)
function dbAll(query, params = []) {
    return new Promise((resolve, reject) => {
        db.all(query, params, (err, rows) => (err ? reject(err) : resolve(rows || [])));
    });
}
function dbGet(query, params = []) {
    return new Promise((resolve, reject) => {
        db.get(query, params, (err, row) => (err ? reject(err) : resolve(row)));
    });
}

// Gamificação: registra pontos para o usuário. Nunca deixa uma falha aqui
// derrubar a ação principal (ex: concluir um PDI não pode falhar por causa
// de um erro ao pontuar), por isso sempre "engole" o próprio erro.
function darPontos(userId, pontos, motivo) {
    if (!userId) return;
    db.run(`INSERT INTO points_ledger (user_id, points, reason) VALUES (?, ?, ?)`, [userId, pontos, motivo], () => {});
}

// Cria uma notificação in-app (sininho) para um usuário. Best-effort: nunca
// derruba a requisição que a chamou, por isso engole os próprios erros.
// `link` é opcional e deve ser a chave de uma aba do menu (ex: 'mentorships'),
// usada pelo frontend para levar o usuário direto para a tela relevante.
function notificar(userId, title, message, link) {
    if (!userId) return;
    db.run(`INSERT INTO notifications (user_id, title, message, link) VALUES (?, ?, ?, ?)`,
        [userId, title, message || '', link || null], () => {});
    // Espelha a notificação por WhatsApp, se o usuário tiver número cadastrado
    // e tiver ligado os avisos por WhatsApp (best-effort, nunca atrapalha o resto).
    db.get(`SELECT whatsapp_number, whatsapp_notifications FROM users WHERE id = ?`, [userId], (err, u) => {
        if (err || !u || !u.whatsapp_number || !u.whatsapp_notifications) return;
        enviarWhatsApp(u.whatsapp_number, `🔔 *${title}*\n${message || ''}`).catch(() => {});
    });
}

// Mesma ideia, mas resolve o user_id a partir de um employee_id (útil quando
// a ação acontece sobre um registro de "employees" e precisamos avisar o
// usuário logado correspondente, que pode nem existir caso o executivo não
// tenha conta própria — nesse caso simplesmente não faz nada).
function notificarPorEmployeeId(employeeId, title, message, link) {
    if (!employeeId) return;
    db.get(`SELECT id FROM users WHERE employee_id = ?`, [employeeId], (err, user) => {
        if (!err && user) notificar(user.id, title, message, link);
    });
}

// Avisa TODOS os gestores (client_admin) cadastrados na empresa — pode haver
// mais de um acesso desde os "Acessos da Empresa" (Membros).
function notificarPorCompanyAdmins(companyId, title, message, link) {
    if (!companyId) return;
    db.all(`SELECT id FROM users WHERE company_id = ? AND role = 'client_admin'`, [companyId], (err, gestores) => {
        if (!err) gestores.forEach(g => notificar(g.id, title, message, link));
    });
}

// Quando uma vaga se abre em um evento (alguém cancela a inscrição), avisa por
// e-mail a próxima pessoa da lista de espera e a remove da fila. Best-effort:
// nunca deve derrubar a requisição que a chamou, por isso engole os próprios erros.
function notificarProximoDaListaDeEspera(eventId) {
    db.get(
        `SELECT ew.id as waitlistId, ew.user_id, u.email, u.name, ev.title, ev.event_date, ev.event_time
         FROM event_waitlist ew
         JOIN users u ON u.id = ew.user_id
         JOIN events ev ON ev.id = ew.event_id
         WHERE ew.event_id = ? ORDER BY ew.created_at ASC LIMIT 1`,
        [eventId],
        (err, proximo) => {
            if (err || !proximo) return;
            db.run(`DELETE FROM event_waitlist WHERE id = ?`, [proximo.waitlistId], () => {});
            notificar(proximo.user_id, `Vaga aberta: ${proximo.title}`, `Uma vaga abriu no evento "${proximo.title}". Garanta sua inscrição!`, 'events');
            if (!proximo.email) return;
            const mailOptions = {
                from: process.env.SMTP_FROM || '"Impulsionar V4" <no-reply@impulsionar.com>',
                to: proximo.email,
                subject: `Uma vaga abriu para "${proximo.title}"!`,
                html: `<p>Olá, ${proximo.name}!</p>
                       <p>Uma vaga acabou de abrir no evento <strong>${proximo.title}</strong>
                       (${proximo.event_date}${proximo.event_time ? ' às ' + proximo.event_time : ''}).</p>
                       <p>Entre na plataforma Impulsionar V4 e garanta sua inscrição antes que a vaga seja preenchida novamente.</p>`
            };
            transporter.sendMail(mailOptions, (mailErr) => {
                if (mailErr) console.warn('⚠️  Não foi possível enviar o e-mail de vaga de evento (SMTP não configurado?).', mailErr.message);
            });
        }
    );
}

// ============================================================
// AUTENTICAÇÃO E AUTORIZAÇÃO
// Perfis: 'admin' (Master, acesso total), 'client_admin' (gestor de
// uma corporação, restrito à própria empresa), 'autonomous'
// (executivo autônomo, restrito ao próprio registro).
// ============================================================

// Rotas que não exigem token (login e auto-cadastro de executivo autônomo)
const ROTAS_PUBLICAS = ['/api/login', '/api/register-autonomous', '/api/forgot-password', '/api/reset-password', '/api/webhooks/mercadopago', '/api/portal/register', '/api/webhooks/autentique', '/api/webhooks/whatsapp'];

// ---------- Validação de CNPJ/CPF (dígitos verificadores, algoritmo oficial) ----------
function validarCNPJ(cnpj) {
    const nums = String(cnpj || '').replace(/\D/g, '');
    if (nums.length !== 14 || /^(\d)\1{13}$/.test(nums)) return false;
    const calcularDigito = (base) => {
        let pesos = base.length === 12 ? [5,4,3,2,9,8,7,6,5,4,3,2] : [6,5,4,3,2,9,8,7,6,5,4,3,2];
        let soma = 0;
        for (let i = 0; i < base.length; i++) soma += Number(base[i]) * pesos[i];
        const resto = soma % 11;
        return resto < 2 ? 0 : 11 - resto;
    };
    const d1 = calcularDigito(nums.slice(0, 12));
    const d2 = calcularDigito(nums.slice(0, 12) + d1);
    return nums === nums.slice(0, 12) + String(d1) + String(d2);
}

function validarCPF(cpf) {
    const nums = String(cpf || '').replace(/\D/g, '');
    if (nums.length !== 11 || /^(\d)\1{10}$/.test(nums)) return false;
    const calcularDigito = (base, pesoInicial) => {
        let soma = 0;
        for (let i = 0; i < base.length; i++) soma += Number(base[i]) * (pesoInicial - i);
        const resto = (soma * 10) % 11;
        return resto === 10 ? 0 : resto;
    };
    const d1 = calcularDigito(nums.slice(0, 9), 10);
    const d2 = calcularDigito(nums.slice(0, 9) + d1, 11);
    return nums === nums.slice(0, 9) + String(d1) + String(d2);
}

function validarDocumento(tipo, valor) {
    if (tipo === 'cpf') return validarCPF(valor);
    return validarCNPJ(valor);
}

// Lê uma chave do painel de automação (retorna true/false). Se a chave ainda
// não existir na tabela (ex: banco antigo antes desta função existir), assume
// o padrão "false" (mais seguro: sempre pede aprovação do Master).
async function automacaoLigada(chave) {
    try {
        const linha = await dbGet(`SELECT value FROM automation_settings WHERE key = ?`, [chave]);
        return linha ? linha.value === '1' : false;
    } catch (e) { return false; }
}

function authenticateToken(req, res, next) {
    // A autenticação só se aplica às rotas da API. O frontend (HTML/CSS/JS em
    // /public, incluindo a própria página de login) precisa carregar livremente
    // — sem isso, nem a tela de login consegue aparecer no navegador.
    if (!req.path.startsWith('/api/')) return next();
    if (ROTAS_PUBLICAS.includes(req.path)) return next();
    // Prefixo dedicado para rotas públicas com parâmetro na URL (ex: /api/public/vagas/:id) —
    // usado pelo link de compartilhamento da vaga, que qualquer pessoa pode abrir sem login.
    if (req.path.startsWith('/api/public/')) return next();

    const authHeader = req.headers['authorization'];
    const token = authHeader && authHeader.split(' ')[1];
    if (!token) return res.status(401).json({ error: 'Token de acesso não fornecido.' });

    jwt.verify(token, JWT_SECRET, (err, payload) => {
        if (err) return res.status(403).json({ error: 'Token inválido ou expirado. Faça login novamente.' });
        req.user = payload; // { userId, role, companyId, employeeId }
        next();
    });
}

function requireRole(...roles) {
    return (req, res, next) => {
        if (!req.user || !roles.includes(req.user.role)) {
            return res.status(403).json({ error: 'Seu perfil não tem permissão para esta ação.' });
        }
        next();
    };
}

// Calcula o escopo de dados permitido para o usuário logado.
// admin -> sem restrição; client_admin -> restrito à própria empresa;
// autonomous -> restrito ao próprio registro de executivo;
// mentor -> restrito ao próprio registro de mentor (usado no Academy/painel social).
function buildScope(req) {
    if (req.user.role === 'admin') return {};
    if (req.user.role === 'client_admin') return { companyId: req.user.companyId };
    if (req.user.role === 'autonomous') return { employeeId: req.user.employeeId };
    if (req.user.role === 'mentor') return { mentorId: req.user.mentorId };
    return { deny: true };
}

// Garante que o employee_id informado no corpo/params pertence ao escopo do usuário
function ensureEmployeeAccess(getEmployeeId) {
    return async (req, res, next) => {
        if (req.user.role === 'admin') return next();
        try {
            const employeeId = getEmployeeId(req);
            if (!employeeId) return res.status(400).json({ error: 'employee_id não informado.' });

            if (req.user.role === 'autonomous') {
                if (String(employeeId) !== String(req.user.employeeId)) {
                    return res.status(403).json({ error: 'Acesso restrito aos seus próprios dados.' });
                }
                return next();
            }
            if (req.user.role === 'client_admin') {
                const row = await dbGet(`SELECT company_id FROM employees WHERE id = ?`, [employeeId]);
                if (!row || String(row.company_id) !== String(req.user.companyId)) {
                    return res.status(403).json({ error: 'Este executivo não pertence à sua corporação.' });
                }
                return next();
            }
            return res.status(403).json({ error: 'Perfil sem permissão.' });
        } catch (e) {
            return res.status(500).json({ error: 'Erro ao validar permissão de acesso.' });
        }
    };
}

// Mesma checagem, mas a partir de um registro existente (PDI, assessment, mentoria)
// identificado por :id — descobre o employee_id do próprio registro antes de validar.
function ensureRecordAccess(table) {
    return async (req, res, next) => {
        if (req.user.role === 'admin') return next();
        try {
            const row = await dbGet(`SELECT employee_id FROM ${table} WHERE id = ?`, [req.params.id]);
            if (!row) return res.status(404).json({ error: 'Registro não encontrado.' });
            return ensureEmployeeAccess(() => row.employee_id)(req, res, next);
        } catch (e) {
            return res.status(500).json({ error: 'Erro ao validar permissão de acesso.' });
        }
    };
}

app.use(authenticateToken);

// Neutraliza CSV Injection: se o valor começa com =, +, - ou @, o Excel/Sheets
// pode interpretá-lo como fórmula ao abrir o arquivo exportado. Prefixamos com
// aspas simples para forçar leitura como texto, e escapamos aspas duplas internas.
function csvSafe(value) {
    if (value === null || value === undefined) return '';
    let str = String(value);
    if (/^[=+\-@]/.test(str)) str = "'" + str;
    return str.replace(/"/g, '""');
}

// APIS DE NOTIFICAÇÕES
app.post('/api/notifications/send-email', async (req, res) => {
    const { to, subject, message } = req.body;
    if (!to || !message) return res.status(400).json({ error: 'Destinatário e mensagem são obrigatórios.' });

    try {
        let mailOptions = {
            from: '"Plataforma Impulsionar V4" <noreply@impulsionar.com>',
            to: to,
            subject: subject || 'Aviso Importante - Impulsionar V4',
            html: `<div style="font-family:sans-serif; padding:20px; border:1px solid #cbd5e1; border-radius:8px;">
                     <h2 style="color:#0b1329;">Impulsionar V4 - Notificação Executiva</h2>
                     <p style="font-size:15px; color:#334155; line-height:1.5;">${message}</p>
                     <hr style="border:none; border-top:1px solid #e2e8f0; margin:20px 0;">
                     <p style="font-size:12px; color:#64748b;">Mensagem automática do sistema corporativo.</p>
                   </div>`
        };

        transporter.sendMail(mailOptions, (error, info) => {
            if (error) {
                return res.json({ message: 'E-mail simulado com sucesso (Modo local sem SMTP ativo)!' });
            }
            res.json({ message: 'E-mail enviado com sucesso!', info: info.response });
        });
    } catch (e) {
        res.status(500).json({ error: 'Erro ao processar envio de e-mail.' });
    }
});

app.post('/api/notifications/whatsapp-link', (req, res) => {
    const { phone, message } = req.body;
    if (!phone) return res.status(400).json({ error: 'Número de telemóvel obrigatório.' });
    const cleanPhone = phone.replace(/\D/g, '');
    const encodedMessage = encodeURIComponent(message || 'Olá, atualização na Impulsionar V4.');
    const whatsappUrl = `https://wa.me/${cleanPhone}?text=${encodedMessage}`;
    res.json({ whatsappUrl });
});

app.get('/api/dashboard/stats', async (req, res) => {
    const companyId = req.query.company_id;
    const JANELA_ALERTA_DIAS = 5; // conforme especificado: PDIs vencendo em até 5 dias

    try {
        const empFiltro = companyId ? `WHERE e.company_id = ?` : '';
        const empParams = companyId ? [companyId] : [];

        const totalCompanies = companyId
            ? 1
            : ((await dbGet(`SELECT COUNT(*) as n FROM companies`)) || {}).n || 0;

        const totalEmployees = ((await dbGet(
            `SELECT COUNT(*) as n FROM employees e ${empFiltro}`, empParams
        )) || {}).n || 0;

        const totalVideos = companyId
            ? 0
            : ((await dbGet(`SELECT COUNT(*) as n FROM video_lessons`)) || {}).n || 0;

        const totalAssessments = ((await dbGet(
            `SELECT COUNT(a.id) as n FROM assessments a JOIN employees e ON a.employee_id = e.id ${empFiltro}`,
            empParams
        )) || {}).n || 0;

        // Distribuição percentual dos perfis DISC
        const discRowsRaw = await dbAll(
            `SELECT COALESCE(NULLIF(e.disc_profile, ''), 'A definir') as perfil, COUNT(*) as total
             FROM employees e ${empFiltro}
             GROUP BY perfil`,
            empParams
        );
        const discDistribution = discRowsRaw.map(r => ({
            perfil: r.perfil,
            total: r.total,
            percentual: totalEmployees > 0 ? Number(((r.total / totalEmployees) * 100).toFixed(1)) : 0
        }));

        // Avanço dos talentos pelas Fases Executivas
        const phaseRowsRaw = await dbAll(
            `SELECT COALESCE(NULLIF(e.executive_phase, ''), 'Fase 1: Diagnóstico') as fase, COUNT(*) as total
             FROM employees e ${empFiltro}
             GROUP BY fase`,
            empParams
        );
        const phaseDistribution = phaseRowsRaw.map(r => ({
            fase: r.fase,
            total: r.total,
            percentual: totalEmployees > 0 ? Number(((r.total / totalEmployees) * 100).toFixed(1)) : 0
        }));

        // Alertas de PDI: vencidos ou vencendo na janela de 5 dias, ainda não concluídos
        const pdiFiltro = companyId ? `AND e.company_id = ?` : '';
        const pdiRows = await dbAll(
            `SELECT p.id, p.objective, p.deadline, p.status,
                    e.id as employeeId, e.name as execName, e.photo_url as execPhoto,
                    COALESCE(c.name, 'Executivo Autónomo') as companyName
             FROM pd_plans p
             JOIN employees e ON p.employee_id = e.id
             LEFT JOIN companies c ON e.company_id = c.id
             WHERE p.status != 'Concluído' AND p.deadline IS NOT NULL AND p.deadline != '' ${pdiFiltro}`,
            empParams
        );

        const hoje = new Date();
        hoje.setHours(0, 0, 0, 0);

        const comDiasRestantes = pdiRows
            .map(p => {
                const prazo = new Date(p.deadline);
                if (isNaN(prazo.getTime())) return null;
                prazo.setHours(0, 0, 0, 0);
                const diasRestantes = Math.round((prazo - hoje) / (1000 * 60 * 60 * 24));
                return { ...p, diasRestantes };
            })
            .filter(p => p && p.diasRestantes <= JANELA_ALERTA_DIAS)
            .sort((a, b) => a.diasRestantes - b.diasRestantes);

        const vencidos = comDiasRestantes.filter(p => p.diasRestantes < 0);
        const vencendoEmBreve = comDiasRestantes.filter(p => p.diasRestantes >= 0);

        // Nota média geral de Competence Check (para o KPI de performance)
        const mediaGeral = await dbGet(
            `SELECT AVG(a.score) as media FROM assessments a JOIN employees e ON a.employee_id = e.id ${empFiltro}`,
            empParams
        );
        const averageScore = mediaGeral && mediaGeral.media != null ? Number(mediaGeral.media.toFixed(1)) : 0;

        // Radar de competências: nota média por competência de liderança avaliada
        const competenceRows = await dbAll(
            `SELECT a.leadership_competence as competencia, AVG(a.score) as media, COUNT(*) as total
             FROM assessments a JOIN employees e ON a.employee_id = e.id ${empFiltro}
             GROUP BY a.leadership_competence`,
            empParams
        );
        const competenceRadar = competenceRows.map(r => ({
            competencia: r.competencia,
            media: Number((r.media || 0).toFixed(1)),
            total: r.total
        }));

        // Tendência: volume e nota média de assessments nos últimos 6 meses
        const tendenciaRows = await dbAll(
            `SELECT strftime('%Y-%m', a.date) as mes, COUNT(*) as total, AVG(a.score) as media
             FROM assessments a JOIN employees e ON a.employee_id = e.id ${empFiltro}
             GROUP BY mes ORDER BY mes ASC`,
            empParams
        );
        // Preenche os últimos 6 meses mesmo quando não há dados, para o gráfico não "pular" meses
        const mapaMeses = new Map(tendenciaRows.map(r => [r.mes, r]));
        const trendAssessments = [];
        for (let i = 5; i >= 0; i--) {
            const d = new Date();
            d.setMonth(d.getMonth() - i);
            const chave = d.toISOString().slice(0, 7);
            const r = mapaMeses.get(chave);
            trendAssessments.push({
                mes: chave,
                total: r ? r.total : 0,
                mediaScore: r && r.media != null ? Number(r.media.toFixed(1)) : null
            });
        }

        // Ranking de performance: top talentos por progresso e nota média de assessment
        const rankingRows = await dbAll(
            `SELECT e.id, e.name, e.progress_percentage, e.executive_phase,
                    COALESCE(c.name, 'Executivo Autónomo') as companyName,
                    (SELECT AVG(a2.score) FROM assessments a2 WHERE a2.employee_id = e.id) as mediaAssessment
             FROM employees e LEFT JOIN companies c ON e.company_id = c.id
             ${empFiltro}
             ORDER BY e.progress_percentage DESC
             LIMIT 5`,
            empParams
        );
        const topPerformers = rankingRows.map(r => ({
            id: r.id, name: r.name, companyName: r.companyName,
            progresso: r.progress_percentage || 0,
            fase: r.executive_phase,
            mediaAssessment: r.mediaAssessment != null ? Number(r.mediaAssessment.toFixed(1)) : null
        }));

        // Mentorias realizadas no mês corrente (indicador de atividade da consultoria)
        const inicioMes = new Date(); inicioMes.setDate(1); inicioMes.setHours(0, 0, 0, 0);
        const mentFiltro = companyId ? `AND e.company_id = ?` : '';
        const mentoriasMesRow = await dbGet(
            `SELECT COUNT(*) as n FROM mentorships m JOIN employees e ON m.employee_id = e.id
             WHERE m.status = 'Realizada' AND m.meeting_date >= ? ${mentFiltro}`,
            companyId ? [inicioMes.toISOString(), companyId] : [inicioMes.toISOString()]
        );
        const mentoriasRealizadasNoMes = (mentoriasMesRow || {}).n || 0;

        // Taxa de conclusão de PDI (para o KPI de performance geral)
        const pdiTotalRow = await dbGet(
            `SELECT COUNT(*) as n FROM pd_plans p JOIN employees e ON p.employee_id = e.id ${empFiltro}`,
            empParams
        );
        const pdiConcluidoRow = await dbGet(
            `SELECT COUNT(*) as n FROM pd_plans p JOIN employees e ON p.employee_id = e.id
             WHERE p.status = 'Concluído' ${pdiFiltro}`,
            empParams
        );
        const pdiTotal = (pdiTotalRow || {}).n || 0;
        const pdiConcluidos = (pdiConcluidoRow || {}).n || 0;
        const taxaConclusaoPdi = pdiTotal > 0 ? Number(((pdiConcluidos / pdiTotal) * 100).toFixed(1)) : 0;

        res.json({
            totalCompanies,
            totalEmployees,
            totalVideos,
            totalAssessments,
            averageScore,
            taxaConclusaoPdi,
            mentoriasRealizadasNoMes,
            discDistribution,
            phaseDistribution,
            competenceRadar,
            trendAssessments,
            topPerformers,
            pdiAlerts: {
                janelaDias: JANELA_ALERTA_DIAS,
                totalAlertas: comDiasRestantes.length,
                vencidos,
                vencendoEmBreve
            }
        });
    } catch (err) {
        console.error('Erro ao montar dashboard:', err.message);
        res.status(500).json({ error: 'Erro ao carregar indicadores do dashboard.' });
    }
});

app.post('/api/login', (req, res) => {
    const { email, password } = req.body;
    db.get(`SELECT u.*, c.name as companyName, c.logo_url as companyLogoUrl, c.enabled_modules as companyEnabledModules,
                   e.enabled_modules as employeeEnabledModules
            FROM users u LEFT JOIN companies c ON u.company_id = c.id LEFT JOIN employees e ON u.employee_id = e.id
            WHERE u.email = ?`, [email], async (err, user) => {
        if (err || !user || !(await bcrypt.compare(password, user.password))) {
            return res.status(401).json({ error: 'E-mail ou senha incorretos.' });
        }

        const token = jwt.sign(
            { userId: user.id, role: user.role, companyId: user.company_id, employeeId: user.employee_id || null, mentorId: user.mentor_id || null },
            JWT_SECRET,
            { expiresIn: '8h' }
        );

        // Permissão por módulo: para um acesso de empresa (client_admin), o que
        // vale primeiro é a restrição PRÓPRIA daquele usuário (users.enabled_modules);
        // só quando ele não tem nada configurado é que cai no padrão da empresa
        // (companies.enabled_modules) — mantém compatível quem já configurava
        // por empresa antes de existir a opção por usuário.
        let enabledModules = null;
        if (user.role === 'client_admin') {
            const fonte = user.enabled_modules || user.companyEnabledModules;
            if (fonte) { try { enabledModules = JSON.parse(fonte); } catch (e) { enabledModules = null; } }
        } else if (user.role === 'autonomous' && user.employeeEnabledModules) {
            try { enabledModules = JSON.parse(user.employeeEnabledModules); } catch (e) { enabledModules = null; }
        }

        res.json({
            message: 'Login efetuado',
            token,
            user: {
                id: user.id, name: user.name, email: user.email, role: user.role,
                companyId: user.company_id, employeeId: user.employee_id || null, mentorId: user.mentor_id || null,
                companyName: user.companyName || (user.role === 'autonomous' ? 'Executivo Autónomo' : user.role === 'mentor' ? 'Mentor Impulsionar' : user.role === 'candidate' ? 'Portal de Vagas' : 'Master Global'),
                companyLogoUrl: user.companyLogoUrl || null,
                enabledModules
            }
        });
    });
});

/* ==========================================================
   RECUPERAÇÃO DE SENHA (fluxo validado por e-mail)
   ========================================================== */
app.post('/api/forgot-password', async (req, res) => {
    const { email } = req.body;
    if (!email) return res.status(400).json({ error: 'Informe o e-mail cadastrado.' });

    try {
        const user = await dbGet(`SELECT id, name, email FROM users WHERE email = ?`, [email]);
        // Resposta genérica sempre que o e-mail existir ou não, para não revelar
        // quais e-mails estão cadastrados no sistema.
        if (!user) {
            return res.json({ message: 'Se este e-mail estiver cadastrado, você receberá as instruções para redefinir a senha.' });
        }

        const token = crypto.randomBytes(32).toString('hex');
        const expira = new Date(Date.now() + 60 * 60 * 1000).toISOString(); // válido por 1h

        await new Promise((resolve, reject) => {
            db.run(`UPDATE users SET reset_token = ?, reset_token_expires = ? WHERE id = ?`, [token, expira, user.id], (err) => err ? reject(err) : resolve());
        });

        // Usa a mesma "URL Pública do Sistema" configurável em Meu Perfil (admin),
        // em vez de uma variável de ambiente separada — assim o link de
        // redefinição sempre aponta pro endereço público real, e não localhost.
        const baseUrl = appBaseUrlAtiva || process.env.APP_URL || `http://localhost:${PORT}`;
        const linkRedefinicao = `${baseUrl}/redefinir-senha.html?token=${token}`;

        const mailOptions = {
            from: '"Plataforma Impulsionar V4" <noreply@impulsionar.com>',
            to: user.email,
            subject: 'Redefinição de senha — Impulsionar V4',
            html: `<div style="font-family:sans-serif; padding:20px; border:1px solid #cbd5e1; border-radius:8px;">
                     <h2 style="color:#0b1329;">Impulsionar V4</h2>
                     <p style="font-size:15px; color:#334155;">Olá, ${user.name}. Recebemos um pedido para redefinir sua senha.</p>
                     <p style="margin:20px 0;"><a href="${linkRedefinicao}" style="background:#0b1329; color:#fff; padding:10px 18px; border-radius:8px; text-decoration:none;">Redefinir minha senha</a></p>
                     <p style="font-size:12px; color:#64748b;">Este link expira em 1 hora. Se você não pediu essa redefinição, ignore este e-mail.</p>
                   </div>`
        };

        transporter.sendMail(mailOptions, (error) => {
            if (error) {
                console.warn('⚠️  Não foi possível enviar o e-mail de redefinição (SMTP não configurado?). Link gerado:', linkRedefinicao);
            }
        });

        res.json({ message: 'Se este e-mail estiver cadastrado, você receberá as instruções para redefinir a senha.' });
    } catch (e) {
        res.status(500).json({ error: 'Erro ao processar o pedido de redefinição.' });
    }
});

app.post('/api/reset-password', async (req, res) => {
    const { token, newPassword } = req.body;
    if (!token || !newPassword) return res.status(400).json({ error: 'Token e nova senha são obrigatórios.' });

    try {
        const user = await dbGet(`SELECT id, reset_token_expires FROM users WHERE reset_token = ?`, [token]);
        if (!user) return res.status(400).json({ error: 'Link inválido ou já utilizado.' });
        if (!user.reset_token_expires || new Date(user.reset_token_expires) < new Date()) {
            return res.status(400).json({ error: 'Este link expirou. Solicite uma nova redefinição de senha.' });
        }

        const hash = await bcrypt.hash(newPassword, 10);
        await new Promise((resolve, reject) => {
            db.run(`UPDATE users SET password = ?, reset_token = NULL, reset_token_expires = NULL WHERE id = ?`, [hash, user.id], (err) => err ? reject(err) : resolve());
        });

        res.json({ message: 'Senha redefinida com sucesso! Você já pode fazer login.' });
    } catch (e) {
        res.status(500).json({ error: 'Erro ao redefinir a senha.' });
    }
});

// O LEFT JOIN usa o client_admin mais antigo (MIN(id)) como "gestor principal"
// exibido na listagem — mesmo quando a empresa tem vários acessos cadastrados
// (Membros da Empresa), isso evita duplicar a linha da empresa na tabela.
app.get('/api/companies', requireRole('admin'), (req, res) => {
    db.all(
        `SELECT c.*, u.name as adminName, u.email as adminEmail,
                (SELECT COUNT(*) FROM users WHERE company_id = c.id AND role = 'client_admin') as totalMembros
         FROM companies c
         LEFT JOIN users u ON u.id = (SELECT id FROM users WHERE company_id = c.id AND role = 'client_admin' ORDER BY id ASC LIMIT 1)`,
        [], (err, rows) => res.json(rows || [])
    );
});

// Permite que o próprio gestor da empresa (client_admin) veja e depois edite
// os dados/identidade visual da sua corporação, sem precisar do Master.
app.get('/api/companies/:id', requireRole('admin', 'client_admin'), (req, res) => {
    if (req.user.role === 'client_admin' && String(req.params.id) !== String(req.user.companyId)) {
        return res.status(403).json({ error: 'Você só pode ver os dados da sua própria corporação.' });
    }
    db.get(`SELECT * FROM companies WHERE id = ?`, [req.params.id], (err, row) => {
        if (err || !row) return res.status(404).json({ error: 'Empresa não encontrada.' });
        res.json(row);
    });
});

// Dados completos da empresa + do gestor responsável, usados só para
// pré-preencher o contrato (razão social, CNPJ, endereço completo e o
// signatário) na hora de criar — sem isso, o Master tinha que digitar tudo de
// novo manualmente e o contrato corria o risco de sair sem CNPJ/endereço
// (o que compromete a validade jurídica do documento).
app.get('/api/admin/companies/:id/contract-data', requireRole('admin'), async (req, res) => {
    try {
        const empresa = await dbGet(`SELECT * FROM companies WHERE id = ?`, [req.params.id]);
        if (!empresa) return res.status(404).json({ error: 'Empresa não encontrada.' });
        const gestor = await dbGet(`SELECT name, email FROM users WHERE company_id = ? AND role = 'client_admin' ORDER BY id ASC LIMIT 1`, [req.params.id]);
        const enderecoPartes = [
            empresa.street ? (empresa.street + (empresa.address_number ? ', ' + empresa.address_number : '')) : null,
            empresa.neighborhood || null,
            empresa.city && empresa.state ? (empresa.city + '/' + empresa.state) : (empresa.city || empresa.state || null),
            empresa.cep ? ('CEP ' + empresa.cep) : null
        ].filter(Boolean);
        const enderecoFormatado = enderecoPartes.length ? enderecoPartes.join(', ') : (empresa.address || '');
        res.json({
            name: empresa.name,
            cnpj: empresa.cnpj,
            documentType: empresa.document_type || 'cnpj',
            phone: empresa.phone || '',
            enderecoFormatado,
            adminName: gestor ? gestor.name : '',
            adminEmail: gestor ? gestor.email : ''
        });
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar os dados da empresa.' }); }
});

app.get('/api/export/companies', requireRole('admin'), (req, res) => {
    db.all(
        `SELECT c.name, c.cnpj, c.segment, c.phone, c.address, u.name as adminName, u.email as adminEmail
         FROM companies c
         LEFT JOIN users u ON u.id = (SELECT id FROM users WHERE company_id = c.id AND role = 'client_admin' ORDER BY id ASC LIMIT 1)`,
        [], (err, rows) => {
        if (err) return res.status(500).json({ error: 'Erro' });
        let csv = "Razao Social;CNPJ;Segmento;Telefone;Endereco;Gestor;E-mail Gestor\n";
        rows.forEach(r => { csv += `"${csvSafe(r.name)}";"${csvSafe(r.cnpj)}";"${csvSafe(r.segment)}";"${csvSafe(r.phone)}";"${csvSafe(r.address)}";"${csvSafe(r.adminName)}";"${csvSafe(r.adminEmail)}"\n`; });
        res.setHeader('Content-Type', 'text/csv; charset=utf-8');
        res.setHeader('Content-Disposition', 'attachment; filename=empresas_impulsionar.csv');
        res.status(200).send(Buffer.from('\uFEFF' + csv, 'utf-8'));
    });
});

app.post('/api/companies', requireRole('admin'), async (req, res) => {
    const {
        name, cnpj, document_type, segment, phone, logo_url, plan_id, adminName, adminEmail, adminPassword,
        cep, street, address_number, neighborhood, city, state, company_size
    } = req.body;
    const tipoDoc = document_type === 'cpf' ? 'cpf' : 'cnpj';
    if (!validarDocumento(tipoDoc, cnpj)) {
        return res.status(400).json({ error: tipoDoc === 'cpf' ? 'CPF inválido. Confira os números digitados.' : 'CNPJ inválido. Confira os números digitados.' });
    }
    const documentoLimpo = String(cnpj).replace(/\D/g, '');
    const enderecoCompleto = [street, address_number].filter(Boolean).join(', ') +
        (neighborhood ? ' - ' + neighborhood : '') + (city ? ', ' + city : '') + (state ? '/' + state : '') +
        (cep ? ' — CEP ' + cep : '');
    try {
        const hash = await bcrypt.hash(adminPassword, 10);
        db.run(
            `INSERT INTO companies (name, cnpj, document_type, segment, phone, address, logo_url, plan_id, cep, street, address_number, neighborhood, city, state, company_size)
             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
            [name, documentoLimpo, tipoDoc, segment, phone, enderecoCompleto, logo_url, plan_id || null, cep || '', street || '', address_number || '', neighborhood || '', city || '', (state || '').toUpperCase(), company_size || ''],
            function(err) {
                if (err) return res.status(400).json({ error: `Erro ao registrar: ${tipoDoc === 'cpf' ? 'CPF' : 'CNPJ'} já existe.` });
                const companyId = this.lastID;
                db.run(`INSERT INTO users (name, email, password, company_id, role) VALUES (?, ?, ?, ?, 'client_admin')`, [adminName, adminEmail, hash, companyId], () => {
                    res.json({ message: 'Corporação e Gestor registados com sucesso!' });
                });
            }
        );
    } catch (e) { res.status(500).json({ error: 'Erro interno.' }); }
});

app.put('/api/companies/:id', requireRole('admin', 'client_admin'), async (req, res) => {
    const { id } = req.params;
    if (req.user.role === 'client_admin' && String(id) !== String(req.user.companyId)) {
        return res.status(403).json({ error: 'Você só pode editar os dados da sua própria corporação.' });
    }
    const {
        name, cnpj, document_type, segment, phone, logo_url, plan_id, newPassword,
        cep, street, address_number, neighborhood, city, state, company_size,
        adminName, adminEmail, adminPassword, enabled_modules
    } = req.body;
    const tipoDoc = document_type === 'cpf' ? 'cpf' : 'cnpj';
    if (cnpj && !validarDocumento(tipoDoc, cnpj)) {
        return res.status(400).json({ error: tipoDoc === 'cpf' ? 'CPF inválido. Confira os números digitados.' : 'CNPJ inválido. Confira os números digitados.' });
    }
    const documentoLimpo = cnpj ? String(cnpj).replace(/\D/g, '') : cnpj;
    const enderecoCompleto = [street, address_number].filter(Boolean).join(', ') +
        (neighborhood ? ' - ' + neighborhood : '') + (city ? ', ' + city : '') + (state ? '/' + state : '') +
        (cep ? ' — CEP ' + cep : '');
    const query = req.user.role === 'admin'
        ? `UPDATE companies SET name = ?, cnpj = ?, document_type = ?, segment = ?, phone = ?, address = ?, logo_url = ?, plan_id = ?,
             cep = ?, street = ?, address_number = ?, neighborhood = ?, city = ?, state = ?, company_size = ? WHERE id = ?`
        : `UPDATE companies SET name = ?, cnpj = ?, document_type = ?, segment = ?, phone = ?, address = ?, logo_url = ?,
             cep = ?, street = ?, address_number = ?, neighborhood = ?, city = ?, state = ?, company_size = ? WHERE id = ?`;
    const camposComuns = [name, documentoLimpo, tipoDoc, segment, phone, enderecoCompleto, logo_url];
    const camposEndereco = [cep || '', street || '', address_number || '', neighborhood || '', city || '', (state || '').toUpperCase(), company_size || ''];
    const params = req.user.role === 'admin'
        ? [...camposComuns, plan_id || null, ...camposEndereco, id]
        : [...camposComuns, ...camposEndereco, id];
    db.run(query, params, async (err) => {
        if (err) return res.status(400).json({ error: 'Erro' });

        // Permissões de acesso da empresa (só o Master define, conforme o contrato
        // de consultoria fechado) — NULL/coluna vazia = sem restrição (libera tudo).
        // enabled_modules === null (checkbox desmarcado) limpa qualquer restrição
        // anterior; um array grava a lista marcada; undefined (campo não enviado,
        // ex: tela de "Minha Empresa" do cliente) deixa a coluna como está.
        if (req.user.role === 'admin' && enabled_modules !== undefined) {
            db.run(`UPDATE companies SET enabled_modules = ? WHERE id = ?`, [Array.isArray(enabled_modules) ? JSON.stringify(enabled_modules) : null, id], () => {});
        }

        // Atualiza nome/e-mail do Gestor (client_admin) junto com os dados da empresa,
        // já que "Nome do Gestor" e "E-mail do Gestor" fazem parte do mesmo cadastro.
        const atualizarGestorEDepoisSenha = () => {
            if (newPassword && newPassword.trim() !== '') {
                bcrypt.hash(newPassword, 10).then(hash => {
                    db.run(`UPDATE users SET password = ? WHERE company_id = ? AND role = 'client_admin'`, [hash, id], () => res.json({ message: 'Atualizado!' }));
                });
            } else {
                res.json({ message: 'Atualizado!' });
            }
        };

        if (adminName !== undefined || adminEmail !== undefined) {
            const camposGestor = [];
            const paramsGestor = [];
            if (adminName !== undefined) { camposGestor.push('name = ?'); paramsGestor.push(adminName); }
            if (adminEmail !== undefined && adminEmail.trim() !== '') { camposGestor.push('email = ?'); paramsGestor.push(adminEmail.trim()); }
            if (camposGestor.length) {
                paramsGestor.push(id);
                db.run(`UPDATE users SET ${camposGestor.join(', ')} WHERE company_id = ? AND role = 'client_admin'`, paramsGestor, async function(errGestor) {
                    if (errGestor) return res.status(400).json({ error: 'Não foi possível atualizar o gestor: e-mail já está em uso por outra conta.' });
                    // Nenhum Gestor existia ainda para esta empresa — se veio nome, e-mail
                    // e senha, cria o acesso dele agora em vez de só tentar atualizar.
                    if (this.changes === 0 && adminName && adminEmail && adminPassword) {
                        try {
                            const hash = await bcrypt.hash(adminPassword, 10);
                            db.run(`INSERT INTO users (name, email, password, company_id, role) VALUES (?, ?, ?, ?, 'client_admin')`,
                                [adminName, adminEmail.trim(), hash, id], (errCriar) => {
                                    if (errCriar) return res.status(400).json({ error: 'Não foi possível criar o gestor: e-mail já está em uso por outra conta.' });
                                    res.json({ message: 'Atualizado! Acesso do Gestor criado.' });
                                });
                        } catch (e) { res.status(500).json({ error: 'Erro ao criar o acesso do gestor.' }); }
                        return;
                    }
                    atualizarGestorEDepoisSenha();
                });
                return;
            }
        }
        atualizarGestorEDepoisSenha();
    });
});

// Permite que o Master acesse a conta de um Gestor (client_admin) para dar suporte,
// sem precisar da senha dele. Gera um token de login normal, igual ao /api/login.
app.post('/api/companies/:id/login-as', requireRole('admin'), async (req, res) => {
    try {
        const gestor = req.body.memberId
            ? await dbGet(
                `SELECT u.*, c.name as companyName, c.logo_url as companyLogoUrl, c.enabled_modules as companyEnabledModules FROM users u LEFT JOIN companies c ON u.company_id = c.id WHERE u.id = ? AND u.company_id = ? AND u.role = 'client_admin'`,
                [req.body.memberId, req.params.id])
            : await dbGet(
                `SELECT u.*, c.name as companyName, c.logo_url as companyLogoUrl, c.enabled_modules as companyEnabledModules FROM users u LEFT JOIN companies c ON u.company_id = c.id WHERE u.company_id = ? AND u.role = 'client_admin' ORDER BY u.id ASC LIMIT 1`,
                [req.params.id]
            );
        if (!gestor) return res.status(404).json({ error: 'Esta empresa ainda não tem um Gestor cadastrado.' });
        const token = jwt.sign(
            { userId: gestor.id, role: gestor.role, companyId: gestor.company_id, employeeId: gestor.employee_id || null, mentorId: gestor.mentor_id || null },
            JWT_SECRET,
            { expiresIn: '2h' }
        );
        let enabledModules = null;
        const fonteModulos = gestor.enabled_modules || gestor.companyEnabledModules;
        if (fonteModulos) { try { enabledModules = JSON.parse(fonteModulos); } catch (e) { enabledModules = null; } }
        res.json({
            token,
            user: {
                id: gestor.id, name: gestor.name, email: gestor.email, role: gestor.role,
                companyId: gestor.company_id, employeeId: gestor.employee_id || null, mentorId: gestor.mentor_id || null,
                companyName: gestor.companyName || 'Empresa', companyLogoUrl: gestor.companyLogoUrl || null,
                enabledModules
            }
        });
    } catch (e) { res.status(500).json({ error: 'Erro ao gerar acesso.' }); }
});

// ---------- Membros da Empresa (acessos adicionais) ----------
// Além do Gestor principal, o Master (ou o próprio Gestor) pode cadastrar
// outros membros da corporação cliente com o e-mail e senha PRÓPRIOS de cada
// um — todos entram com o mesmo nível de acesso (client_admin), escopados na
// mesma empresa, sem precisar compartilhar um único login.
app.get('/api/companies/:id/members', requireRole('admin', 'client_admin'), async (req, res) => {
    if (req.user.role === 'client_admin' && String(req.params.id) !== String(req.user.companyId)) {
        return res.status(403).json({ error: 'Você só pode ver os acessos da sua própria corporação.' });
    }
    const membros = await dbAll(`SELECT id, name, email, enabled_modules FROM users WHERE company_id = ? AND role = 'client_admin' ORDER BY id ASC`, [req.params.id]);
    res.json(membros);
});

// Permissão de módulos individual deste acesso (gestor ou membro) — some com
// o padrão da empresa (ver /api/login), então aqui só grava a restrição
// PRÓPRIA deste usuário. enabled_modules === null (nenhuma marcação) limpa a
// restrição própria e volta a valer o padrão da empresa.
app.put('/api/companies/:id/members/:memberId/permissions', requireRole('admin', 'client_admin'), async (req, res) => {
    if (req.user.role === 'client_admin' && String(req.params.id) !== String(req.user.companyId)) {
        return res.status(403).json({ error: 'Você só pode gerenciar acessos da sua própria corporação.' });
    }
    const { enabled_modules } = req.body;
    const valor = Array.isArray(enabled_modules) ? JSON.stringify(enabled_modules) : null;
    db.run(`UPDATE users SET enabled_modules = ? WHERE id = ? AND company_id = ? AND role = 'client_admin'`,
        [valor, req.params.memberId, req.params.id], function (err) {
            if (err) return res.status(400).json({ error: 'Erro ao salvar permissões.' });
            if (this.changes === 0) return res.status(404).json({ error: 'Acesso não encontrado nesta empresa.' });
            res.json({ message: 'Permissões deste acesso atualizadas!' });
        });
});

app.post('/api/companies/:id/members', requireRole('admin', 'client_admin'), async (req, res) => {
    if (req.user.role === 'client_admin' && String(req.params.id) !== String(req.user.companyId)) {
        return res.status(403).json({ error: 'Você só pode adicionar acessos na sua própria corporação.' });
    }
    const { name, email, password } = req.body;
    if (!name || !email || !password) return res.status(400).json({ error: 'Informe nome, e-mail e senha do novo acesso.' });
    if (password.length < 6) return res.status(400).json({ error: 'A senha precisa ter pelo menos 6 caracteres.' });
    try {
        const empresa = await dbGet(`SELECT id FROM companies WHERE id = ?`, [req.params.id]);
        if (!empresa) return res.status(404).json({ error: 'Empresa não encontrada.' });
        const hash = await bcrypt.hash(password, 10);
        db.run(`INSERT INTO users (name, email, password, company_id, role) VALUES (?, ?, ?, ?, 'client_admin')`,
            [name, email.trim(), hash, req.params.id], (err) => {
                if (err) return res.status(400).json({ error: 'Este e-mail já está em uso por outra conta.' });
                res.json({ message: 'Acesso criado! Este membro já pode entrar com o e-mail e senha próprios dele.' });
            });
    } catch (e) { res.status(500).json({ error: 'Erro ao criar o acesso.' }); }
});

app.delete('/api/companies/:id/members/:memberId', requireRole('admin', 'client_admin'), async (req, res) => {
    if (req.user.role === 'client_admin' && String(req.params.id) !== String(req.user.companyId)) {
        return res.status(403).json({ error: 'Você só pode remover acessos da sua própria corporação.' });
    }
    const total = await dbGet(`SELECT COUNT(*) as n FROM users WHERE company_id = ? AND role = 'client_admin'`, [req.params.id]);
    if (total && total.n <= 1) return res.status(400).json({ error: 'Esta é a única conta de acesso da empresa — cadastre outra antes de remover esta.' });
    db.run(`DELETE FROM users WHERE id = ? AND company_id = ? AND role = 'client_admin'`, [req.params.memberId, req.params.id], function(err) {
        if (err) return res.status(400).json({ error: 'Erro ao remover.' });
        if (this.changes === 0) return res.status(404).json({ error: 'Acesso não encontrado.' });
        res.json({ message: 'Acesso removido!' });
    });
});

// Crédito de dias de divulgação de vaga — o Master concede como bônus (soma ao
// saldo já existente, nunca sobrescreve) e ele é consumido automaticamente
// quando uma vaga desta empresa é aprovada, em vez de cobrar pelo Mercado Pago.
app.post('/api/companies/:id/credito', requireRole('admin'), async (req, res) => {
    const dias = Number(req.body.dias);
    if (!dias || dias <= 0) return res.status(400).json({ error: 'Informe uma quantidade de dias válida.' });
    db.run(`UPDATE companies SET vaga_credito_dias = COALESCE(vaga_credito_dias, 0) + ? WHERE id = ?`, [dias, req.params.id], function(err) {
        if (err) return res.status(400).json({ error: 'Erro ao adicionar crédito.' });
        if (this.changes === 0) return res.status(404).json({ error: 'Empresa não encontrada.' });
        dbGet(`SELECT vaga_credito_dias FROM companies WHERE id = ?`, [req.params.id]).then(c => {
            res.json({ message: `+${dias} dias de crédito adicionados!`, total: c.vaga_credito_dias });
        });
    });
});

app.delete('/api/companies/:id', requireRole('admin'), (req, res) => {
    db.run(`DELETE FROM users WHERE company_id = ?`, [req.params.id], () => {
        db.run(`DELETE FROM employees WHERE company_id = ?`, [req.params.id], () => {
            db.run(`DELETE FROM companies WHERE id = ?`, [req.params.id], () => res.json({ message: 'Eliminado!' }));
        });
    });
});

app.get('/api/employees', (req, res) => {
    const scope = buildScope(req);
    if (scope.deny) return res.status(403).json({ error: 'Perfil sem permissão.' });

    let query = `SELECT e.*, COALESCE(c.name, 'Executivo Autónomo') as companyName, c.logo_url as companyLogo,
        (SELECT COUNT(*) FROM users u WHERE u.employee_id = e.id) as hasLogin
        FROM employees e LEFT JOIN companies c ON e.company_id = c.id`;
    const conditions = [];
    const params = [];
    if (scope.companyId) { conditions.push('e.company_id = ?'); params.push(scope.companyId); }
    if (scope.employeeId) { conditions.push('e.id = ?'); params.push(scope.employeeId); }
    if (!scope.companyId && !scope.employeeId && req.query.company_id) { conditions.push('e.company_id = ?'); params.push(req.query.company_id); }
    if (conditions.length) query += ' WHERE ' + conditions.join(' AND ');
    db.all(query, params, (err, rows) => res.json(rows || []));
});

app.get('/api/export/employees', requireRole('admin', 'client_admin'), (req, res) => {
    const scope = buildScope(req);
    let query = `SELECT e.name, e.role, e.email, e.phone, COALESCE(c.name, 'Executivo Autónomo') as companyName, e.disc_profile, e.executive_phase, e.progress_percentage FROM employees e LEFT JOIN companies c ON e.company_id = c.id`;
    const conditions = [];
    const params = [];
    if (scope.companyId) { conditions.push('e.company_id = ?'); params.push(scope.companyId); }
    else if (req.query.company_id) { conditions.push('e.company_id = ?'); params.push(req.query.company_id); }
    if (conditions.length) query += ' WHERE ' + conditions.join(' AND ');
    db.all(query, params, (err, rows) => {
        if (err) return res.status(500).json({ error: 'Erro' });
        let csvContent = "Nome;Cargo;Email;Telefone;Corporacao;DISC;Fase;Progresso (%)\n";
        rows.forEach(r => { csvContent += `"${csvSafe(r.name)}";"${csvSafe(r.role)}";"${csvSafe(r.email)}";"${csvSafe(r.phone)}";"${csvSafe(r.companyName)}";"${csvSafe(r.disc_profile)}";"${csvSafe(r.executive_phase)}";"${csvSafe(r.progress_percentage || 0)}"\n`; });
        res.setHeader('Content-Type', 'text/csv; charset=utf-8');
        res.setHeader('Content-Disposition', 'attachment; filename=banco_de_talentos_impulsionar.csv');
        res.status(200).send(Buffer.from('\uFEFF' + csvContent, 'utf-8'));
    });
});

// Autocadastro público de Executivo Autónomo (Portal do Executivo Autónomo, item 3.8).
// Rota separada e pública — listada em ROTAS_PUBLICAS, não passa por autenticação.
app.post('/api/register-autonomous', async (req, res) => {
    const { name, role, email, phone, password } = req.body;
    if (!name || !email || !password) {
        return res.status(400).json({ error: 'Nome, e-mail e senha são obrigatórios.' });
    }
    db.run(`INSERT INTO employees (company_id, name, role, email, phone, performance_level, executive_phase, progress_percentage, disc_profile) VALUES (NULL, ?, ?, ?, ?, 'Em Desenvolvimento', 'Fase 1: Diagnóstico', 10, 'A definir')`,
        [name, role, email, phone], async function (err) {
            if (err) return res.status(400).json({ error: err.message });
            const employeeId = this.lastID;
            try {
                const hash = await bcrypt.hash(password, 10);
                db.run(`INSERT INTO users (name, email, password, company_id, employee_id, role) VALUES (?, ?, ?, NULL, ?, 'autonomous')`,
                    [name, email, hash, employeeId],
                    (userErr) => {
                        if (userErr) return res.status(400).json({ error: 'E-mail já cadastrado.' });
                        res.json({ message: 'Cadastro realizado com sucesso! Faça login para continuar.' });
                    }
                );
            } catch (e) {
                res.status(500).json({ error: 'Erro ao criar credenciais de acesso.' });
            }
        });
});

// ============================================================
// PORTAL PÚBLICO DE VAGAS E CURRÍCULOS
// Qualquer pessoa pode se cadastrar como candidato (role 'candidate' na
// tabela users) e montar seu currículo. Só aparece para as empresas depois
// que o Master aprova o perfil. As empresas publicam vagas pagando por um
// período (dias) configurado pelo Master em "Planos de Vaga".
// ============================================================

app.post('/api/portal/register', async (req, res) => {
    const { name, email, password, phone, desired_role, city } = req.body;
    if (!name || !email || !password) return res.status(400).json({ error: 'Nome, e-mail e senha são obrigatórios.' });
    try {
        const hash = await bcrypt.hash(password, 10);
        const aprovacaoAutomatica = await automacaoLigada('auto_approve_resumes');
        db.run(`INSERT INTO users (name, email, password, role) VALUES (?, ?, ?, 'candidate')`, [name, email, hash], function (err) {
            if (err) return res.status(400).json({ error: 'E-mail já cadastrado.' });
            const userId = this.lastID;
            db.run(`INSERT INTO candidate_profiles (user_id, phone, desired_role, city, status) VALUES (?, ?, ?, ?, ?)`,
                [userId, phone || '', desired_role || '', city || '', aprovacaoAutomatica ? 'approved' : 'pending'], (e2) => {
                    if (e2) return res.status(500).json({ error: 'Erro ao criar o perfil de candidato.' });
                    res.json({ message: aprovacaoAutomatica
                        ? 'Cadastro realizado! Seu currículo já está aprovado — complete-o para aparecer melhor posicionado para as empresas.'
                        : 'Cadastro realizado! Faça login para completar seu currículo — ele passa por uma aprovação do Master antes de aparecer para as empresas.' });
                });
        });
    } catch (e) { res.status(500).json({ error: 'Erro ao criar credenciais de acesso.' }); }
});

app.get('/api/portal/me', requireRole('candidate'), async (req, res) => {
    try {
        const perfil = await dbGet(`SELECT u.name, u.email, cp.* FROM candidate_profiles cp JOIN users u ON u.id = cp.user_id WHERE cp.user_id = ?`, [req.user.userId]);
        if (!perfil) return res.status(404).json({ error: 'Perfil não encontrado.' });
        res.json(perfil);
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar seu perfil.' }); }
});

app.put('/api/portal/me', requireRole('candidate'), (req, res) => {
    const {
        phone, desired_role, city, bio, skills, linkedin_url, resume_url,
        photo_url, gender, education_level, languages, first_job, experiences_json,
        desired_states, desired_cities
    } = req.body;
    db.run(
        `UPDATE candidate_profiles SET phone = ?, desired_role = ?, city = ?, bio = ?, skills = ?, linkedin_url = ?, resume_url = ?,
            photo_url = ?, gender = ?, education_level = ?, languages = ?, first_job = ?, experiences_json = ?,
            desired_states = ?, desired_cities = ?
         WHERE user_id = ?`,
        [
            phone || '', desired_role || '', city || '', bio || '', skills || '', linkedin_url || '', resume_url || '',
            photo_url || '', gender || '', education_level || '', languages || '', first_job ? 1 : 0,
            experiences_json || '[]', desired_states || '', desired_cities || '',
            req.user.userId
        ],
        (err) => {
            if (err) return res.status(400).json({ error: err.message });
            res.json({ message: 'Currículo atualizado!' });
        }
    );
});

// Lista as vagas ativas e ainda dentro do prazo pago — visível para o próprio candidato.
app.get('/api/portal/vagas', requireRole('candidate'), async (req, res) => {
    try {
        const perfil = await dbGet(`SELECT desired_states, desired_cities FROM candidate_profiles WHERE user_id = ?`, [req.user.userId]);
        const estadosDesejados = (perfil?.desired_states || '').split(',').map(s => s.trim()).filter(Boolean);
        const cidadesDesejadas = (perfil?.desired_cities || '').split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
        const filtrarPorRegiao = req.query.somenteMinhaRegiao === '1' && (estadosDesejados.length > 0 || cidadesDesejadas.length > 0);

        const vagas = await dbAll(
            `SELECT jp.*, c.name as companyName, c.logo_url as companyLogo,
                    (SELECT COUNT(*) FROM job_applications ja WHERE ja.job_posting_id = jp.id AND ja.candidate_user_id = ?) as jaCandidatei,
                    (SELECT COUNT(*) FROM job_posting_likes jl WHERE jl.job_posting_id = jp.id) as totalCurtidas,
                    (SELECT COUNT(*) FROM job_posting_likes jl WHERE jl.job_posting_id = jp.id AND jl.candidate_user_id = ?) as euCurti
             FROM job_postings jp JOIN companies c ON c.id = jp.company_id
             WHERE jp.status = 'active' AND jp.deleted_at IS NULL AND jp.expires_at > CURRENT_TIMESTAMP
             ORDER BY jp.published_at DESC`,
            [req.user.userId, req.user.userId]
        );
        const comFlag = vagas.map(v => {
            const bateEstado = estadosDesejados.length > 0 && estadosDesejados.includes((v.state || '').toUpperCase());
            const bateCidade = cidadesDesejadas.length > 0 && cidadesDesejadas.some(c => (v.location || '').toLowerCase().includes(c));
            return { ...v, jaCandidatei: !!v.jaCandidatei, euCurti: !!v.euCurti, minhaRegiao: !!(v.is_remote || bateEstado || bateCidade || (estadosDesejados.length === 0 && cidadesDesejadas.length === 0)) };
        });
        const resultado = filtrarPorRegiao ? comFlag.filter(v => v.minhaRegiao) : comFlag;
        res.json(resultado);
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar vagas.' }); }
});

app.post('/api/portal/vagas/:id/apply', requireRole('candidate'), async (req, res) => {
    try {
        const vaga = await dbGet(`SELECT * FROM job_postings WHERE id = ? AND status = 'active' AND deleted_at IS NULL AND expires_at > CURRENT_TIMESTAMP`, [req.params.id]);
        if (!vaga) return res.status(404).json({ error: 'Vaga não encontrada ou não está mais disponível.' });
        db.run(`INSERT INTO job_applications (job_posting_id, candidate_user_id) VALUES (?, ?)`, [req.params.id, req.user.userId], (err) => {
            if (err) return res.status(400).json({ error: 'Você já se candidatou a esta vaga.' });
            res.json({ message: 'Candidatura enviada!' });
        });
    } catch (e) { res.status(500).json({ error: 'Erro ao se candidatar.' }); }
});

// Curtir/descurtir uma vaga (toggle) — mede engajamento no portal, sem
// nenhuma etapa formal de candidatura.
app.post('/api/portal/vagas/:id/like', requireRole('candidate'), async (req, res) => {
    try {
        const jaCurtiu = await dbGet(`SELECT id FROM job_posting_likes WHERE job_posting_id = ? AND candidate_user_id = ?`, [req.params.id, req.user.userId]);
        if (jaCurtiu) {
            db.run(`DELETE FROM job_posting_likes WHERE id = ?`, [jaCurtiu.id], () => {});
        } else {
            db.run(`INSERT INTO job_posting_likes (job_posting_id, candidate_user_id) VALUES (?, ?)`, [req.params.id, req.user.userId], () => {});
        }
        const total = await dbGet(`SELECT COUNT(*) as total FROM job_posting_likes WHERE job_posting_id = ?`, [req.params.id]);
        res.json({ curtido: !jaCurtiu, totalCurtidas: total.total });
    } catch (e) { res.status(500).json({ error: 'Erro ao curtir a vaga.' }); }
});

// Página pública da vaga (link de compartilhamento) — sem login, qualquer
// pessoa pode ver os dados básicos antes de criar uma conta de candidato.
app.get('/api/public/vagas/:id', async (req, res) => {
    try {
        const vaga = await dbGet(
            `SELECT jp.id, jp.title, jp.description, jp.location, jp.state, jp.is_remote, jp.seniority, jp.salary_range,
                    jp.education, jp.languages, jp.requirements, jp.responsibilities, jp.benefits, jp.photo_url,
                    c.name as companyName, c.logo_url as companyLogo,
                    (SELECT COUNT(*) FROM job_posting_likes jl WHERE jl.job_posting_id = jp.id) as totalCurtidas
             FROM job_postings jp JOIN companies c ON c.id = jp.company_id
             WHERE jp.id = ? AND jp.status = 'active' AND jp.deleted_at IS NULL AND jp.expires_at > CURRENT_TIMESTAMP`,
            [req.params.id]
        );
        if (!vaga) return res.status(404).json({ error: 'Vaga não encontrada ou não está mais disponível.' });
        res.json(vaga);
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar a vaga.' }); }
});

app.get('/api/portal/my-applications', requireRole('candidate'), async (req, res) => {
    try {
        const lista = await dbAll(
            `SELECT ja.applied_at, jp.title, jp.location, jp.is_remote, c.name as companyName
             FROM job_applications ja JOIN job_postings jp ON jp.id = ja.job_posting_id JOIN companies c ON c.id = jp.company_id
             WHERE ja.candidate_user_id = ? ORDER BY ja.applied_at DESC`,
            [req.user.userId]
        );
        res.json(lista);
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar candidaturas.' }); }
});

// ---------- Aprovação de currículos pelo Master ----------
app.get('/api/admin/candidates', requireRole('admin'), async (req, res) => {
    try {
        const status = req.query.status || 'pending';
        let query = `SELECT u.id as userId, u.name, u.email, cp.* FROM candidate_profiles cp JOIN users u ON u.id = cp.user_id`;
        const params = [];
        if (status !== 'all') { query += ` WHERE cp.status = ?`; params.push(status); }
        query += ` ORDER BY cp.created_at DESC`;
        const lista = await dbAll(query, params);
        res.json(lista);
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar candidatos.' }); }
});

// Detalhe completo de um candidato (usado no "Ver" antes de aprovar/reprovar).
app.get('/api/admin/candidates/:userId', requireRole('admin'), async (req, res) => {
    try {
        const perfil = await dbGet(
            `SELECT u.id as userId, u.name, u.email, cp.* FROM candidate_profiles cp JOIN users u ON u.id = cp.user_id WHERE u.id = ?`,
            [req.params.userId]
        );
        if (!perfil) return res.status(404).json({ error: 'Candidato não encontrado.' });
        res.json(perfil);
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar candidato.' }); }
});

app.put('/api/admin/candidates/:userId/status', requireRole('admin'), async (req, res) => {
    const { status, feedback } = req.body;
    if (!['approved', 'rejected', 'pending', 'ajustes_solicitados'].includes(status)) return res.status(400).json({ error: 'Status inválido.' });
    try {
        await new Promise((resolve, reject) => db.run(`UPDATE candidate_profiles SET status = ? WHERE user_id = ?`, [status, req.params.userId], (err) => err ? reject(err) : resolve()));
        const rotulos = {
            approved: 'Currículo aprovado!', rejected: 'Currículo não aprovado', pending: 'Currículo em análise',
            ajustes_solicitados: 'O Master pediu ajustes no seu currículo'
        };
        const mensagensPadrao = {
            approved: 'Seu currículo já pode ser encontrado pelas empresas no portal de vagas.',
            rejected: 'Revise seu currículo e tente novamente.',
            pending: 'Seu currículo voltou para análise.',
            ajustes_solicitados: 'Confira a mensagem do Master na aba "Fale com o Master" para saber o que precisa alterar.'
        };
        notificar(req.params.userId, rotulos[status], mensagensPadrao[status], status === 'ajustes_solicitados' ? 'portalMensagens' : 'portalPerfil');
        if (feedback && feedback.trim()) {
            await new Promise((resolve, reject) => db.run(
                `INSERT INTO candidate_messages (candidate_user_id, sender, message, read_by_master) VALUES (?, 'master', ?, 1)`,
                [req.params.userId, feedback.trim()],
                (err) => err ? reject(err) : resolve()
            ));
        }
        res.json({ message: 'Status atualizado!' });
    } catch (e) { res.status(400).json({ error: e.message }); }
});

// A IA analisa o currículo do candidato e sugere uma decisão (aprovar / pedir
// ajustes / reprovar) com uma mensagem de feedback pronta. Se o autopilot de
// revisão de currículo estiver ligado, aplica direto; senão, manda pro
// WhatsApp do Master aprovar antes.
app.post('/api/admin/candidates/:userId/ai-review', requireRole('admin'), async (req, res) => {
    try {
        const c = await dbGet(
            `SELECT u.name, u.email, cp.* FROM candidate_profiles cp JOIN users u ON u.id = cp.user_id WHERE u.id = ?`,
            [req.params.userId]
        );
        if (!c) return res.status(404).json({ error: 'Candidato não encontrado.' });
        const experiencias = c.experiences_json ? JSON.parse(c.experiences_json) : [];
        const resumoCandidato = [
            `Nome: ${c.name}`,
            `Cargo desejado: ${c.desired_role || '-'}`,
            `Cidade: ${c.city || '-'}`,
            `Escolaridade: ${c.education_level || '-'}`,
            `Idiomas: ${c.languages || '-'}`,
            `Habilidades: ${c.skills || '-'}`,
            `Sobre: ${c.bio || '-'}`,
            c.first_job ? 'Primeiro emprego (sem experiências anteriores).' : `Experiências: ${experiencias.map(e => `${e.role || ''} em ${e.company || ''} (${e.period || ''})`).join('; ') || '-'}`,
            `Currículo anexado: ${c.resume_url ? 'sim' : 'não'}`
        ].join('\n');

        const resposta = await perguntarIA(
            'Você é o assistente de RH da Impulsionar, que analisa currículos cadastrados no portal público de vagas antes do Master aprovar. ' +
            'Analise os dados do candidato e responda EXATAMENTE neste formato, sem nada antes ou depois:\n' +
            'STATUS: approved ou ajustes_solicitados ou rejected\n' +
            'MENSAGEM: <mensagem curta e cordial em português para o candidato, explicando a decisão; se pedir ajustes, diga exatamente o que falta>',
            resumoCandidato
        );
        const statusMatch = resposta.match(/STATUS:\s*(approved|ajustes_solicitados|rejected)/i);
        const mensagemMatch = resposta.match(/MENSAGEM:\s*([\s\S]*)/i);
        const statusSugerido = statusMatch ? statusMatch[1].toLowerCase() : 'ajustes_solicitados';
        const mensagemSugerida = mensagemMatch ? mensagemMatch[1].trim() : resposta;

        const idAcao = await criarAcaoIA({
            type: 'ajuste_curriculo',
            candidateUserId: req.params.userId,
            demand: resumoCandidato,
            proposal: mensagemSugerida,
            suggestedStatus: statusSugerido,
            autopilotKey: 'ai_resume_review_autopilot'
        });
        const acao = await dbGet(`SELECT * FROM ai_pending_actions WHERE id = ?`, [idAcao]);
        res.json({ message: acao.status === 'aprovado' ? 'A IA analisou e já aplicou a decisão (autopilot ligado).' : 'A IA sugeriu uma decisão — aguardando aprovação do Master pelo WhatsApp.', acao });
    } catch (e) { res.status(400).json({ error: e.message }); }
});

// ---------- Interação (chat) entre o candidato e o Master sobre o currículo ----------
app.get('/api/admin/candidates/:userId/messages', requireRole('admin'), async (req, res) => {
    try {
        const lista = await dbAll(`SELECT * FROM candidate_messages WHERE candidate_user_id = ? ORDER BY created_at ASC`, [req.params.userId]);
        db.run(`UPDATE candidate_messages SET read_by_master = 1 WHERE candidate_user_id = ? AND sender = 'candidate'`, [req.params.userId], () => {});
        res.json(lista);
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar mensagens.' }); }
});

app.post('/api/admin/candidates/:userId/messages', requireRole('admin'), (req, res) => {
    const { message } = req.body;
    if (!message || !message.trim()) return res.status(400).json({ error: 'Escreva uma mensagem.' });
    db.run(`INSERT INTO candidate_messages (candidate_user_id, sender, message, read_by_master) VALUES (?, 'master', ?, 1)`,
        [req.params.userId, message.trim()], (err) => {
            if (err) return res.status(400).json({ error: err.message });
            notificar(req.params.userId, 'Nova mensagem do Master', 'Você recebeu uma mensagem sobre o seu currículo.', 'portalMensagens');
            res.json({ message: 'Enviado!' });
        });
});

app.get('/api/portal/messages', requireRole('candidate'), async (req, res) => {
    try {
        const topic = req.query.topic === 'suporte' ? 'suporte' : 'curriculo';
        const lista = await dbAll(`SELECT * FROM candidate_messages WHERE candidate_user_id = ? AND topic = ? ORDER BY created_at ASC`, [req.user.userId, topic]);
        db.run(`UPDATE candidate_messages SET read_by_candidate = 1 WHERE candidate_user_id = ? AND sender = 'master' AND topic = ?`, [req.user.userId, topic], () => {});
        res.json(lista);
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar mensagens.' }); }
});

app.post('/api/portal/messages', requireRole('candidate'), (req, res) => {
    const { message, topic } = req.body;
    const topicoFinal = topic === 'suporte' ? 'suporte' : 'curriculo';
    if (!message || !message.trim()) return res.status(400).json({ error: 'Escreva uma mensagem.' });
    db.run(`INSERT INTO candidate_messages (candidate_user_id, sender, message, topic, read_by_candidate) VALUES (?, 'candidate', ?, ?, 1)`,
        [req.user.userId, message.trim(), topicoFinal], (err) => {
            if (err) return res.status(400).json({ error: err.message });
            if (topicoFinal === 'suporte') {
                db.all(`SELECT id FROM users WHERE role = 'admin'`, [], (e2, admins) => {
                    (admins || []).forEach(a => notificar(a.id, 'Novo pedido de suporte', 'Um candidato do portal enviou uma mensagem de suporte.', 'candidatesApproval'));
                });
                // A IA tenta rascunhar uma resposta em segundo plano — se falhar
                // (sem ANTHROPIC_API_KEY, por exemplo), não afeta o candidato:
                // a mensagem dele já foi salva e o Master ainda vê e responde manualmente.
                (async () => {
                    try {
                        const proposta = await perguntarIA(
                            'Você é o assistente de suporte da Impulsionar, uma plataforma de desenvolvimento de executivos e portal de vagas. ' +
                            'Um candidato do portal mandou uma mensagem de suporte. Escreva uma resposta breve, cordial e útil em português, ' +
                            'como se fosse o Master respondendo. Se não souber resolver algo técnico, oriente a aguardar o time humano.',
                            message.trim()
                        );
                        await criarAcaoIA({
                            type: 'suporte_reply',
                            candidateUserId: req.user.userId,
                            demand: message.trim(),
                            proposal: proposta,
                            autopilotKey: 'ai_support_autopilot'
                        });
                    } catch (e) { /* IA não configurada ou indisponível — segue sem ela */ }
                })();
            }
            res.json({ message: 'Enviado!' });
        });
});

// Empresas veem os candidatos aprovados (equivalente ao Banco de Currículos interno, mas para o público externo).
app.get('/api/portal/candidates', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const { q, desiredRole } = req.query;
        let query = `SELECT u.name, u.email, cp.* FROM candidate_profiles cp JOIN users u ON u.id = cp.user_id WHERE cp.status = 'approved'`;
        const params = [];
        if (q) { query += ` AND (u.name LIKE ? OR cp.desired_role LIKE ?)`; params.push(`%${q}%`, `%${q}%`); }
        if (desiredRole) { query += ` AND cp.desired_role LIKE ?`; params.push(`%${desiredRole}%`); }
        query += ` ORDER BY cp.created_at DESC`;
        const lista = await dbAll(query, params);
        res.json(lista);
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar candidatos aprovados.' }); }
});

// ---------- Planos de Vaga (dias x preço, configurado pelo Master) ----------
app.get('/api/vaga-plans', (req, res) => {
    db.all(`SELECT * FROM vaga_plans WHERE active = 1 ORDER BY days ASC`, [], (err, rows) => res.json(rows || []));
});

app.get('/api/admin/vaga-plans', requireRole('admin'), (req, res) => {
    db.all(`SELECT * FROM vaga_plans ORDER BY days ASC`, [], (err, rows) => res.json(rows || []));
});

app.post('/api/vaga-plans', requireRole('admin'), (req, res) => {
    const { label, days, price } = req.body;
    if (!days || !price) return res.status(400).json({ error: 'Informe os dias e o preço do plano.' });
    db.run(`INSERT INTO vaga_plans (label, days, price) VALUES (?, ?, ?)`, [label || '', days, price], function (err) {
        if (err) return res.status(400).json({ error: err.message });
        res.json({ message: 'Plano de vaga criado!', id: this.lastID });
    });
});

app.put('/api/vaga-plans/:id', requireRole('admin'), (req, res) => {
    const { label, days, price, active } = req.body;
    db.run(`UPDATE vaga_plans SET label = ?, days = ?, price = ?, active = ? WHERE id = ?`,
        [label || '', days, price, active === false ? 0 : 1, req.params.id], (err) => {
            if (err) return res.status(400).json({ error: err.message });
            res.json({ message: 'Plano de vaga atualizado!' });
        });
});

app.delete('/api/vaga-plans/:id', requireRole('admin'), (req, res) => {
    db.run(`DELETE FROM vaga_plans WHERE id = ?`, [req.params.id], () => res.json({ message: 'Removido!' }));
});

// ---------- Planos de Fechamento (taxa de sucesso por função, configurado pelo Master) ----------
app.get('/api/closing-fee-plans', (req, res) => {
    db.all(`SELECT * FROM closing_fee_plans WHERE active = 1 ORDER BY price ASC`, [], (err, rows) => res.json(rows || []));
});

app.get('/api/admin/closing-fee-plans', requireRole('admin'), (req, res) => {
    db.all(`SELECT * FROM closing_fee_plans ORDER BY price ASC`, [], (err, rows) => res.json(rows || []));
});

app.post('/api/closing-fee-plans', requireRole('admin'), (req, res) => {
    const { function_label, price } = req.body;
    if (!function_label || !price) return res.status(400).json({ error: 'Informe a função e o valor de fechamento.' });
    db.run(`INSERT INTO closing_fee_plans (function_label, price) VALUES (?, ?)`, [function_label, price], function (err) {
        if (err) return res.status(400).json({ error: err.message });
        res.json({ message: 'Função de fechamento criada!', id: this.lastID });
    });
});

app.put('/api/closing-fee-plans/:id', requireRole('admin'), (req, res) => {
    const { function_label, price, active } = req.body;
    db.run(`UPDATE closing_fee_plans SET function_label = ?, price = ?, active = ? WHERE id = ?`,
        [function_label || '', price, active === false ? 0 : 1, req.params.id], (err) => {
            if (err) return res.status(400).json({ error: err.message });
            res.json({ message: 'Função de fechamento atualizada!' });
        });
});

app.delete('/api/closing-fee-plans/:id', requireRole('admin'), (req, res) => {
    db.run(`DELETE FROM closing_fee_plans WHERE id = ?`, [req.params.id], () => res.json({ message: 'Removido!' }));
});

// ---------- Vagas publicadas pelas empresas ----------
app.get('/api/job-postings', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        let query = `SELECT jp.*, vp.label as planLabel, vp.days as planDays, vp.price as planPrice, c.name as companyName, c.logo_url as companyLogo,
                     cfp.function_label as closingFeeLabel, cfp.price as closingFeePrice,
                     (SELECT COUNT(*) FROM job_applications ja WHERE ja.job_posting_id = jp.id) as totalCandidaturas,
                     (SELECT COUNT(*) FROM job_posting_likes jl WHERE jl.job_posting_id = jp.id) as totalCurtidas,
                     (SELECT u.name FROM job_applications ja2 JOIN users u ON u.id = ja2.candidate_user_id WHERE ja2.id = jp.closed_application_id) as closedApplicantName
                     FROM job_postings jp LEFT JOIN vaga_plans vp ON vp.id = jp.vaga_plan_id LEFT JOIN companies c ON c.id = jp.company_id
                     LEFT JOIN closing_fee_plans cfp ON cfp.id = jp.closing_fee_plan_id`;
        const params = [];
        if (req.user.role === 'client_admin') { query += ` WHERE jp.company_id = ?`; params.push(req.user.companyId); }
        query += ` ORDER BY jp.created_at DESC`;
        const lista = await dbAll(query, params);
        res.json(lista);
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar vagas.' }); }
});

app.post('/api/job-postings', requireRole('admin', 'client_admin'), async (req, res) => {
    const { title, description, location, state, is_remote, seniority, salary_range, vagaPlanId, closingFeePlanId, company_id,
             education, languages, requirements, responsibilities, benefits, photo_url } = req.body;
    if (!title) return res.status(400).json({ error: 'Informe o título da vaga.' });
    const companyId = req.user.role === 'client_admin' ? req.user.companyId : company_id;
    if (!companyId) return res.status(400).json({ error: 'Informe a empresa da vaga.' });
    try {
        const plano = await dbGet(`SELECT * FROM vaga_plans WHERE id = ? AND active = 1`, [vagaPlanId]);
        if (!plano) return res.status(400).json({ error: 'Selecione um plano de divulgação válido.' });
        // Taxa de fechamento é opcional — nem toda empresa/vaga precisa ter uma
        // função de cobrança por contratação vinculada.
        const planoFechamento = closingFeePlanId ? await dbGet(`SELECT id FROM closing_fee_plans WHERE id = ? AND active = 1`, [closingFeePlanId]) : null;

        // Vaga criada pela própria empresa (client_admin) sempre passa pela
        // aprovação do Master antes de ir para pagamento/publicação — é o que
        // garante que só o Master decide o que é divulgado na rede.
        if (req.user.role === 'client_admin') {
            const resultado = await new Promise((resolve, reject) => db.run(
                `INSERT INTO job_postings (company_id, title, description, location, state, is_remote, seniority, salary_range, vaga_plan_id, closing_fee_plan_id, status,
                    education, languages, requirements, responsibilities, benefits, photo_url)
                 VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'pendente_aprovacao', ?, ?, ?, ?, ?, ?)`,
                [companyId, title, description || '', location || '', (state || '').toUpperCase(), is_remote ? 1 : 0, seniority || '', salary_range || '', plano.id, planoFechamento ? planoFechamento.id : null,
                    education || '', languages || '', requirements || '', responsibilities || '', benefits || '', photo_url || ''],
                function (err) { err ? reject(err) : resolve(this.lastID); }
            ));
            db.all(`SELECT id FROM users WHERE role = 'admin'`, [], (e, admins) => {
                if (!e) admins.forEach(a => notificar(a.id, 'Nova vaga aguardando aprovação', `"${title}" — revise e aprove ou rejeite no Portal de Vagas.`, 'jobPostings'));
            });
            return res.json({ message: 'Vaga enviada para aprovação do Master! Assim que for aprovada, você recebe um aviso para pagar (ou ela é publicada direto, se a empresa tiver crédito de dias).', id: resultado, status: 'pendente_aprovacao' });
        }

        // Master criando diretamente em nome de uma empresa — mantém o fluxo
        // antigo, sem etapa de aprovação (ele já é quem aprovaria).
        if (!mpPreference) return res.status(503).json({ error: 'Mercado Pago ainda não foi configurado no servidor (defina MP_ACCESS_TOKEN no .env).' });
        const resultado = await new Promise((resolve, reject) => db.run(
            `INSERT INTO job_postings (company_id, title, description, location, state, is_remote, seniority, salary_range, vaga_plan_id, closing_fee_plan_id, status, approved_by, approved_at,
                education, languages, requirements, responsibilities, benefits, photo_url)
             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending_payment', ?, CURRENT_TIMESTAMP, ?, ?, ?, ?, ?, ?)`,
            [companyId, title, description || '', location || '', (state || '').toUpperCase(), is_remote ? 1 : 0, seniority || '', salary_range || '', plano.id, planoFechamento ? planoFechamento.id : null, req.user.userId,
                education || '', languages || '', requirements || '', responsibilities || '', benefits || '', photo_url || ''],
            function (err) { err ? reject(err) : resolve(this.lastID); }
        ));

        const preference = await mpPreference.create({
            body: {
                items: [{ title: `Divulgação de vaga: ${title} (${plano.label || plano.days + ' dias'})`, quantity: 1, unit_price: Number(plano.price), currency_id: 'BRL' }],
                external_reference: `jobposting:${resultado}`,
                ...montarRetornoMercadoPago()
            }
        });

        const checkoutUrl = preference.init_point;
        db.run(`UPDATE job_postings SET mp_preference_id = ?, checkout_url = ? WHERE id = ?`, [preference.id, checkoutUrl, resultado], () => {});
        res.json({ message: 'Vaga criada! Complete o pagamento para publicá-la.', id: resultado, initPoint: checkoutUrl });
    } catch (e) {
        console.error('Erro ao criar vaga/preferência Mercado Pago:', e.message);
        res.status(400).json({ error: 'Erro ao criar a vaga ou iniciar o pagamento no Mercado Pago.' });
    }
});

// Master aprova uma vaga pendente: se a empresa tiver crédito de dias
// suficiente, publica na hora consumindo o crédito; senão, gera o checkout do
// Mercado Pago e a empresa precisa pagar para publicar.
app.post('/api/job-postings/:id/approve', requireRole('admin'), async (req, res) => {
    try {
        const vaga = await dbGet(`SELECT jp.*, vp.days as planDays, vp.price as planPrice, vp.label as planLabel FROM job_postings jp LEFT JOIN vaga_plans vp ON vp.id = jp.vaga_plan_id WHERE jp.id = ?`, [req.params.id]);
        if (!vaga) return res.status(404).json({ error: 'Vaga não encontrada.' });
        if (vaga.status !== 'pendente_aprovacao') return res.status(400).json({ error: 'Esta vaga já foi analisada.' });
        const empresa = await dbGet(`SELECT vaga_credito_dias FROM companies WHERE id = ?`, [vaga.company_id]);
        const dias = vaga.planDays || 30;
        const creditoDisponivel = (empresa && empresa.vaga_credito_dias) || 0;

        if (creditoDisponivel >= dias) {
            await new Promise((resolve, reject) => db.run(
                `UPDATE job_postings SET status = 'active', approved_by = ?, approved_at = CURRENT_TIMESTAMP,
                    published_at = CURRENT_TIMESTAMP, expires_at = datetime(CURRENT_TIMESTAMP, '+${Number(dias)} days'), paid_with_credit = 1 WHERE id = ?`,
                [req.user.userId, req.params.id], (err) => err ? reject(err) : resolve()
            ));
            db.run(`UPDATE companies SET vaga_credito_dias = vaga_credito_dias - ? WHERE id = ?`, [dias, vaga.company_id], () => {});
            notificarPorCompanyAdmins(vaga.company_id, 'Vaga aprovada e publicada!', `"${vaga.title}" já está no ar, usando o crédito de dias da empresa.`, 'jobPostings');
            return res.json({ message: `Aprovada e publicada usando ${dias} dia(s) de crédito da empresa!` });
        }

        // Sem Mercado Pago configurado — aprova mesmo assim, só sem o link de
        // pagamento (a empresa paga depois via "Pagar / Reabrir Checkout").
        if (!mpPreference) {
            await new Promise((resolve, reject) => db.run(`UPDATE job_postings SET status = 'pending_payment', approved_by = ?, approved_at = CURRENT_TIMESTAMP WHERE id = ?`, [req.user.userId, req.params.id], (err) => err ? reject(err) : resolve()));
            notificarPorCompanyAdmins(vaga.company_id, 'Vaga aprovada!', `"${vaga.title}" foi aprovada — falta o pagamento para publicar.`, 'jobPostings');
            return res.json({ message: 'Aprovada! (Mercado Pago não está configurado no servidor, então o link de pagamento não pôde ser gerado agora — peça para a empresa usar "Pagar / Reabrir Checkout" depois de configurado.)' });
        }

        // Gerar a preferência de pagamento é uma chamada externa (Mercado Pago)
        // que pode falhar por motivos fora do nosso controle (token inválido,
        // conta não configurada, rede). Isso NÃO pode derrubar a aprovação em
        // si — a vaga já foi aprovada pelo Master, então sempre marcamos como
        // aprovada e, se o link de pagamento falhar, avisamos com o motivo real
        // em vez de travar tudo com "Erro ao aprovar a vaga".
        try {
            const preference = await mpPreference.create({
                body: {
                    items: [{ title: `Divulgação de vaga: ${vaga.title} (${vaga.planLabel || dias + ' dias'})`, quantity: 1, unit_price: Number(vaga.planPrice) || 0.01, currency_id: 'BRL' }],
                    external_reference: `jobposting:${vaga.id}`,
                    ...montarRetornoMercadoPago()
                }
            });
            const checkoutUrl = preference.init_point;
            await new Promise((resolve, reject) => db.run(
                `UPDATE job_postings SET status = 'pending_payment', approved_by = ?, approved_at = CURRENT_TIMESTAMP, mp_preference_id = ?, checkout_url = ? WHERE id = ?`,
                [req.user.userId, preference.id, checkoutUrl, vaga.id], (err) => err ? reject(err) : resolve()
            ));
            notificarPorCompanyAdmins(vaga.company_id, 'Vaga aprovada!', `"${vaga.title}" foi aprovada — complete o pagamento para publicar.`, 'jobPostings');
            res.json({ message: 'Aprovada! A empresa já pode pagar para publicar.' });
        } catch (mpErr) {
            console.error('Erro ao criar preferência Mercado Pago na aprovação da vaga:', mpErr.message);
            await new Promise((resolve, reject) => db.run(`UPDATE job_postings SET status = 'pending_payment', approved_by = ?, approved_at = CURRENT_TIMESTAMP WHERE id = ?`, [req.user.userId, req.params.id], (err) => err ? reject(err) : resolve()));
            notificarPorCompanyAdmins(vaga.company_id, 'Vaga aprovada!', `"${vaga.title}" foi aprovada — falta o pagamento para publicar.`, 'jobPostings');
            res.json({ message: `Aprovada! Porém não foi possível gerar o link de pagamento agora (${mpErr.message || 'erro no Mercado Pago'}). A empresa pode tentar de novo pelo botão "Pagar / Reabrir Checkout".` });
        }
    } catch (e) {
        console.error('Erro ao aprovar vaga:', e.message);
        res.status(400).json({ error: 'Erro ao aprovar a vaga: ' + e.message });
    }
});

// "Pagar / Reabrir Checkout": quando a vaga está pendente de pagamento mas o
// link do Mercado Pago não existe (ou expirou), gera uma preferência nova em
// vez de obrigar a empresa a excluir a vaga e recriar tudo do zero.
app.post('/api/job-postings/:id/reopen-checkout', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const vaga = await dbGet(`SELECT jp.*, vp.days as planDays, vp.price as planPrice, vp.label as planLabel FROM job_postings jp LEFT JOIN vaga_plans vp ON vp.id = jp.vaga_plan_id WHERE jp.id = ?`, [req.params.id]);
        if (!vaga) return res.status(404).json({ error: 'Vaga não encontrada.' });
        if (req.user.role === 'client_admin' && vaga.company_id !== req.user.companyId) return res.status(403).json({ error: 'Esta vaga não pertence à sua empresa.' });
        if (vaga.status !== 'pending_payment') return res.status(400).json({ error: 'Esta vaga não está aguardando pagamento.' });
        if (!mpPreference) return res.status(503).json({ error: 'Mercado Pago ainda não foi configurado no servidor.' });

        const dias = vaga.planDays || 30;
        const preference = await mpPreference.create({
            body: {
                items: [{ title: `Divulgação de vaga: ${vaga.title} (${vaga.planLabel || dias + ' dias'})`, quantity: 1, unit_price: Number(vaga.planPrice) || 0.01, currency_id: 'BRL' }],
                external_reference: `jobposting:${vaga.id}`,
                ...montarRetornoMercadoPago()
            }
        });
        const checkoutUrl = preference.init_point;
        await new Promise((resolve, reject) => db.run(
            `UPDATE job_postings SET mp_preference_id = ?, checkout_url = ? WHERE id = ?`,
            [preference.id, checkoutUrl, vaga.id], (err) => err ? reject(err) : resolve()
        ));
        res.json({ message: 'Link de pagamento gerado!', checkoutUrl });
    } catch (e) {
        console.error('Erro ao reabrir checkout da vaga:', e.message);
        res.status(400).json({ error: 'Não foi possível gerar o link de pagamento agora: ' + e.message });
    }
});

// Reabre o checkout da TAXA DE FECHAMENTO (caso o primeiro link tenha expirado
// ou a chamada ao Mercado Pago tenha falhado no momento do fechamento).
app.post('/api/job-postings/:id/reopen-closing-fee-checkout', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const vaga = await dbGet(`SELECT jp.*, cfp.function_label, cfp.price FROM job_postings jp LEFT JOIN closing_fee_plans cfp ON cfp.id = jp.closing_fee_plan_id WHERE jp.id = ?`, [req.params.id]);
        if (!vaga) return res.status(404).json({ error: 'Vaga não encontrada.' });
        if (req.user.role === 'client_admin' && vaga.company_id !== req.user.companyId) return res.status(403).json({ error: 'Esta vaga não pertence à sua empresa.' });
        if (!vaga.closing_fee_plan_id) return res.status(400).json({ error: 'Esta vaga não tem uma taxa de fechamento vinculada.' });
        if (vaga.closing_fee_status === 'paid') return res.status(400).json({ error: 'A taxa de fechamento desta vaga já foi paga.' });
        if (!mpPreference) return res.status(503).json({ error: 'Mercado Pago ainda não foi configurado no servidor.' });

        const preference = await mpPreference.create({
            body: {
                items: [{ title: `Taxa de fechamento — vaga preenchida: ${vaga.title} (${vaga.function_label || 'Função'})`, quantity: 1, unit_price: Number(vaga.price) || 0.01, currency_id: 'BRL' }],
                external_reference: `closingfee:${vaga.id}`,
                ...montarRetornoMercadoPago()
            }
        });
        const checkoutUrl = preference.init_point;
        await new Promise((resolve, reject) => db.run(
            `UPDATE job_postings SET closing_fee_status = 'pending_payment', closing_fee_mp_preference_id = ?, closing_fee_checkout_url = ? WHERE id = ?`,
            [preference.id, checkoutUrl, vaga.id], (err) => err ? reject(err) : resolve()
        ));
        res.json({ message: 'Link de pagamento da taxa de fechamento gerado!', checkoutUrl });
    } catch (e) {
        const detalhe = detalheErroMercadoPago(e);
        console.error('Erro ao reabrir checkout da taxa de fechamento:', detalhe);
        res.status(400).json({ error: `Não foi possível gerar o link de pagamento agora: ${detalhe}` });
    }
});

// Master pede ajustes em vez de aprovar/rejeitar de vez — a vaga volta para a
// empresa com a mensagem do que precisa mudar; ao editar e reenviar, ela volta
// para 'pendente_aprovacao' automaticamente (ver PUT /api/job-postings/:id).
app.post('/api/job-postings/:id/request-changes', requireRole('admin'), async (req, res) => {
    const motivo = (req.body.reason || '').trim();
    if (!motivo) return res.status(400).json({ error: 'Descreva o que precisa ser ajustado.' });
    try {
        const vaga = await dbGet(`SELECT * FROM job_postings WHERE id = ?`, [req.params.id]);
        if (!vaga) return res.status(404).json({ error: 'Vaga não encontrada.' });
        db.run(`UPDATE job_postings SET status = 'ajustes_solicitados', rejection_reason = ?, approved_by = ?, approved_at = CURRENT_TIMESTAMP WHERE id = ?`,
            [motivo, req.user.userId, req.params.id], (err) => {
                if (err) return res.status(400).json({ error: 'Erro ao solicitar ajuste.' });
                notificarPorCompanyAdmins(vaga.company_id, 'Ajustes solicitados na vaga', `"${vaga.title}": ${motivo}`, 'jobPostings');
                res.json({ message: 'Ajuste solicitado — a empresa foi avisada com a mensagem.' });
            });
    } catch (e) { res.status(500).json({ error: 'Erro ao solicitar ajuste.' }); }
});

// Editar uma vaga já criada — a própria empresa (dona da vaga) ou o Master.
// Se a empresa edita uma vaga que estava rejeitada/com ajuste solicitado, ela
// volta para 'pendente_aprovacao' automaticamente (reenvio para nova análise).
app.put('/api/job-postings/:id', requireRole('admin', 'client_admin'), async (req, res) => {
    const { title, description, location, state, is_remote, seniority, salary_range, closingFeePlanId,
             education, languages, requirements, responsibilities, benefits, photo_url } = req.body;
    if (!title) return res.status(400).json({ error: 'Informe o título da vaga.' });
    try {
        const vaga = await dbGet(`SELECT * FROM job_postings WHERE id = ?`, [req.params.id]);
        if (!vaga) return res.status(404).json({ error: 'Vaga não encontrada.' });
        if (req.user.role === 'client_admin' && vaga.company_id !== req.user.companyId) {
            return res.status(403).json({ error: 'Esta vaga não pertence à sua empresa.' });
        }
        // A função de fechamento só pode ser trocada enquanto ainda não foi
        // cobrada (senão mudaria o valor de uma cobrança já gerada/paga).
        let closingFeePlanIdFinal = vaga.closing_fee_plan_id;
        if (!vaga.closing_fee_status) {
            if (closingFeePlanId === null || closingFeePlanId === '' ) closingFeePlanIdFinal = null;
            else if (closingFeePlanId) {
                const planoFechamento = await dbGet(`SELECT id FROM closing_fee_plans WHERE id = ? AND active = 1`, [closingFeePlanId]);
                if (planoFechamento) closingFeePlanIdFinal = planoFechamento.id;
            }
        }
        const reenviarParaAnalise = req.user.role === 'client_admin' && ['rejeitada', 'ajustes_solicitados'].includes(vaga.status);
        const novoStatus = reenviarParaAnalise ? 'pendente_aprovacao' : vaga.status;
        await new Promise((resolve, reject) => db.run(
            `UPDATE job_postings SET title = ?, description = ?, location = ?, state = ?, is_remote = ?, seniority = ?, salary_range = ?, closing_fee_plan_id = ?,
                education = ?, languages = ?, requirements = ?, responsibilities = ?, benefits = ?, photo_url = ?,
                status = ?, rejection_reason = CASE WHEN ? THEN NULL ELSE rejection_reason END
             WHERE id = ?`,
            [title, description || '', location || '', (state || '').toUpperCase(), is_remote ? 1 : 0, seniority || '', salary_range || '', closingFeePlanIdFinal,
                education || '', languages || '', requirements || '', responsibilities || '', benefits || '', photo_url || vaga.photo_url || '',
                novoStatus, reenviarParaAnalise ? 1 : 0, req.params.id],
            (err) => err ? reject(err) : resolve()
        ));
        if (reenviarParaAnalise) {
            db.all(`SELECT id FROM users WHERE role = 'admin'`, [], (e, admins) => {
                if (!e) admins.forEach(a => notificar(a.id, 'Vaga reenviada para aprovação', `"${title}" foi editada e reenviada — revise novamente.`, 'jobPostings'));
            });
        }
        res.json({ message: reenviarParaAnalise ? 'Vaga atualizada e reenviada para aprovação do Master!' : 'Vaga atualizada!' });
    } catch (e) {
        console.error('Erro ao editar vaga:', e.message);
        res.status(400).json({ error: 'Erro ao editar a vaga.' });
    }
});

app.post('/api/job-postings/:id/reject', requireRole('admin'), async (req, res) => {
    const motivo = (req.body.reason || '').trim();
    if (!motivo) return res.status(400).json({ error: 'Informe o motivo da rejeição, para a empresa entender o que ajustar.' });
    try {
        const vaga = await dbGet(`SELECT * FROM job_postings WHERE id = ?`, [req.params.id]);
        if (!vaga) return res.status(404).json({ error: 'Vaga não encontrada.' });
        db.run(`UPDATE job_postings SET status = 'rejeitada', rejection_reason = ?, approved_by = ?, approved_at = CURRENT_TIMESTAMP WHERE id = ?`,
            [motivo, req.user.userId, req.params.id], (err) => {
                if (err) return res.status(400).json({ error: 'Erro ao rejeitar a vaga.' });
                notificarPorCompanyAdmins(vaga.company_id, 'Vaga não aprovada', `"${vaga.title}": ${motivo}`, 'jobPostings');
                res.json({ message: 'Vaga rejeitada — a empresa foi avisada com o motivo.' });
            });
    } catch (e) { res.status(500).json({ error: 'Erro ao rejeitar a vaga.' }); }
});

// Confirma o pagamento manualmente lendo o retorno do Checkout Pro (payment_id
// na URL de volta) — alternativa ao webhook para quem testa localmente sem
// endereço público. Sempre reconsulta o status direto na API do Mercado Pago,
// nunca confia no que vem da URL.
app.post('/api/job-postings/confirm-payment', requireRole('admin', 'client_admin'), async (req, res) => {
    const { paymentId } = req.body;
    if (!paymentId) return res.status(400).json({ error: 'Informe o paymentId.' });
    if (!mpPayment) return res.status(503).json({ error: 'Mercado Pago ainda não foi configurado no servidor.' });
    try {
        const pagamento = await mpPayment.get({ id: paymentId });
        const ref = pagamento.external_reference || '';

        if (ref.startsWith('closingfee:')) {
            const jobId = ref.split(':')[1];
            const vaga = await dbGet(`SELECT * FROM job_postings WHERE id = ?`, [jobId]);
            if (!vaga) return res.status(404).json({ error: 'Vaga não encontrada.' });
            if (req.user.role === 'client_admin' && vaga.company_id !== req.user.companyId) return res.status(403).json({ error: 'Esta vaga não pertence à sua empresa.' });
            if (pagamento.status === 'approved') {
                await new Promise((resolve, reject) => db.run(
                    `UPDATE job_postings SET closing_fee_status = 'paid', closing_fee_payment_id = ?, closing_fee_paid_at = CURRENT_TIMESTAMP WHERE id = ?`,
                    [paymentId, jobId], (err) => err ? reject(err) : resolve()
                ));
            }
            return res.json({ message: pagamento.status === 'approved' ? 'Pagamento da taxa de fechamento confirmado!' : 'Pagamento ainda não aprovado.', status: pagamento.status });
        }

        if (!ref.startsWith('jobposting:')) return res.status(400).json({ error: 'Pagamento não corresponde a uma vaga.' });
        const jobId = ref.split(':')[1];
        const vaga = await dbGet(`SELECT jp.*, vp.days as planDays FROM job_postings jp LEFT JOIN vaga_plans vp ON vp.id = jp.vaga_plan_id WHERE jp.id = ?`, [jobId]);
        if (!vaga) return res.status(404).json({ error: 'Vaga não encontrada.' });
        if (req.user.role === 'client_admin' && vaga.company_id !== req.user.companyId) return res.status(403).json({ error: 'Esta vaga não pertence à sua empresa.' });

        if (pagamento.status === 'approved') {
            const dias = vaga.planDays || 30;
            await new Promise((resolve, reject) => db.run(
                `UPDATE job_postings SET status = 'active', mp_payment_id = ?, published_at = CURRENT_TIMESTAMP, expires_at = datetime(CURRENT_TIMESTAMP, '+${Number(dias)} days') WHERE id = ?`,
                [paymentId, jobId], (err) => err ? reject(err) : resolve()
            ));
        }
        res.json({ message: pagamento.status === 'approved' ? 'Pagamento confirmado — vaga publicada!' : 'Pagamento ainda não aprovado.', status: pagamento.status });
    } catch (e) {
        console.error('Erro ao confirmar pagamento de vaga:', e.message);
        res.status(400).json({ error: 'Erro ao consultar o pagamento no Mercado Pago.' });
    }
});

// Estorna (reembolsa) o pagamento único de uma vaga divulgada. Só o Master
// decide isso (é dinheiro saindo de verdade) — nunca a própria empresa.
// Vaga paga com crédito de dias (paid_with_credit) não passou pelo Mercado
// Pago, então não há o que estornar ali.
app.post('/api/job-postings/:id/refund', requireRole('admin'), async (req, res) => {
    const { id } = req.params;
    if (!mpPaymentRefund) return res.status(503).json({ error: 'Mercado Pago ainda não foi configurado no servidor.' });
    try {
        const vaga = await dbGet(`SELECT * FROM job_postings WHERE id = ?`, [id]);
        if (!vaga) return res.status(404).json({ error: 'Vaga não encontrada.' });
        if (vaga.paid_with_credit) return res.status(400).json({ error: 'Esta vaga foi publicada com crédito de dias, não houve cobrança no Mercado Pago para estornar.' });
        if (!vaga.mp_payment_id) return res.status(400).json({ error: 'Esta vaga não tem um pagamento confirmado no Mercado Pago para estornar.' });
        if (vaga.refunded_at) return res.status(400).json({ error: 'Este pagamento já foi estornado anteriormente.' });

        await mpPaymentRefund.create({ payment_id: vaga.mp_payment_id });

        await new Promise((resolve, reject) => db.run(
            `UPDATE job_postings SET refunded_at = CURRENT_TIMESTAMP, refunded_by = ? WHERE id = ?`,
            [req.user.userId, id], (err) => err ? reject(err) : resolve()
        ));
        notificarGestoresDaEmpresa(vaga.company_id, 'Pagamento estornado', `O pagamento da vaga "${vaga.title}" foi estornado pela Impulsionar. O valor volta pelo mesmo meio usado na compra.`);
        res.json({ message: 'Pagamento estornado com sucesso! O valor volta para o cliente pelo mesmo meio usado na compra.' });
    } catch (e) {
        const detalhe = detalheErroMercadoPago(e);
        console.error('Erro ao estornar pagamento de vaga:', detalhe);
        res.status(400).json({ error: `Erro ao estornar no Mercado Pago: ${detalhe}` });
    }
});

app.get('/api/job-postings/:id/applications', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const vaga = await dbGet(`SELECT * FROM job_postings WHERE id = ?`, [req.params.id]);
        if (!vaga) return res.status(404).json({ error: 'Vaga não encontrada.' });
        if (req.user.role === 'client_admin' && vaga.company_id !== req.user.companyId) return res.status(403).json({ error: 'Esta vaga não pertence à sua empresa.' });
        const lista = await dbAll(
            `SELECT ja.id, u.name, u.email, cp.phone, cp.desired_role, cp.city, cp.bio, cp.resume_url, cp.linkedin_url, ja.applied_at
             FROM job_applications ja JOIN users u ON u.id = ja.candidate_user_id LEFT JOIN candidate_profiles cp ON cp.user_id = u.id
             WHERE ja.job_posting_id = ? ORDER BY ja.applied_at DESC`,
            [req.params.id]
        );
        res.json(lista);
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar candidaturas.' }); }
});

// Fecha a vaga (contratação concluída ou vaga cancelada). Exige justificativa
// e/ou vincular o candidato contratado — nunca os dois em branco, para que
// sempre fique registrado o motivo do fechamento.
app.post('/api/job-postings/:id/close', requireRole('admin', 'client_admin'), async (req, res) => {
    const motivo = (req.body.reason || '').trim();
    const candidatoId = req.body.hired_application_id || null;
    if (!motivo && !candidatoId) return res.status(400).json({ error: 'Informe uma justificativa ou selecione o candidato contratado para fechar a vaga.' });
    try {
        const vaga = await dbGet(`SELECT * FROM job_postings WHERE id = ?`, [req.params.id]);
        if (!vaga) return res.status(404).json({ error: 'Vaga não encontrada.' });
        if (req.user.role === 'client_admin' && vaga.company_id !== req.user.companyId) return res.status(403).json({ error: 'Esta vaga não pertence à sua empresa.' });
        if (candidatoId) {
            const candidatura = await dbGet(`SELECT id FROM job_applications WHERE id = ? AND job_posting_id = ?`, [candidatoId, req.params.id]);
            if (!candidatura) return res.status(400).json({ error: 'Candidatura informada não pertence a esta vaga.' });
        }
        await new Promise((resolve, reject) => db.run(
            `UPDATE job_postings SET status = 'fechada', closed_reason = ?, closed_application_id = ?, closed_at = CURRENT_TIMESTAMP WHERE id = ?`,
            [motivo || null, candidatoId, req.params.id],
            (err) => err ? reject(err) : resolve()
        ));

        // Taxa de fechamento (sucesso na contratação): só quando a vaga foi
        // fechada COM um candidato contratado (não em cancelamento/encerramento
        // sem contratar), quando existe uma função de fechamento vinculada, e
        // só uma vez por vaga (closing_fee_status ainda vazio).
        let cobrancaFechamento = null;
        if (candidatoId && vaga.closing_fee_plan_id && !vaga.closing_fee_status) {
            if (!mpPreference) {
                cobrancaFechamento = { aviso: 'Vaga fechada, mas a taxa de fechamento não pôde ser cobrada: Mercado Pago não está configurado no servidor.' };
            } else {
                try {
                    const planoFechamento = await dbGet(`SELECT * FROM closing_fee_plans WHERE id = ?`, [vaga.closing_fee_plan_id]);
                    const preference = await mpPreference.create({
                        body: {
                            items: [{ title: `Taxa de fechamento — vaga preenchida: ${vaga.title} (${planoFechamento?.function_label || 'Função'})`, quantity: 1, unit_price: Number(planoFechamento?.price) || 0.01, currency_id: 'BRL' }],
                            external_reference: `closingfee:${vaga.id}`,
                            ...montarRetornoMercadoPago()
                        }
                    });
                    const checkoutUrl = preference.init_point;
                    await new Promise((resolve, reject) => db.run(
                        `UPDATE job_postings SET closing_fee_status = 'pending_payment', closing_fee_mp_preference_id = ?, closing_fee_checkout_url = ? WHERE id = ?`,
                        [preference.id, checkoutUrl, vaga.id], (err) => err ? reject(err) : resolve()
                    ));
                    notificarGestoresDaEmpresa(vaga.company_id, 'Taxa de fechamento gerada', `A vaga "${vaga.title}" foi preenchida — complete o pagamento da taxa de fechamento (${planoFechamento?.function_label || ''}).`);
                    cobrancaFechamento = { initPoint: checkoutUrl, funcao: planoFechamento?.function_label, valor: planoFechamento?.price };
                } catch (mpErr) {
                    const detalhe = detalheErroMercadoPago(mpErr);
                    console.error('Erro ao gerar cobrança de taxa de fechamento:', detalhe);
                    cobrancaFechamento = { aviso: `Vaga fechada, mas não foi possível gerar a cobrança da taxa de fechamento agora (${detalhe}). Tente reabrir o checkout depois.` };
                }
            }
        }

        res.json({
            message: cobrancaFechamento?.initPoint
                ? `Vaga fechada! Taxa de fechamento (${cobrancaFechamento.funcao || ''} — R$ ${Number(cobrancaFechamento.valor || 0).toFixed(2)}) gerada — complete o pagamento no Mercado Pago.`
                : (cobrancaFechamento?.aviso || 'Vaga fechada!'),
            closingFeeInitPoint: cobrancaFechamento?.initPoint || null
        });
    } catch (e) { res.status(500).json({ error: 'Erro ao fechar a vaga.' }); }
});

app.delete('/api/job-postings/:id', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const vaga = await dbGet(`SELECT * FROM job_postings WHERE id = ?`, [req.params.id]);
        if (!vaga) return res.status(404).json({ error: 'Vaga não encontrada.' });
        if (req.user.role === 'client_admin' && vaga.company_id !== req.user.companyId) return res.status(403).json({ error: 'Esta vaga não pertence à sua empresa.' });
        // Soft-delete: mantém o histórico (aparece em "Vagas Excluídas") em vez
        // de apagar candidaturas e curtidas de verdade.
        db.run(`UPDATE job_postings SET deleted_at = CURRENT_TIMESTAMP WHERE id = ?`, [req.params.id], (err) => {
            if (err) return res.status(400).json({ error: err.message });
            res.json({ message: 'Vaga movida para Excluídas!' });
        });
    } catch (e) { res.status(500).json({ error: 'Erro ao remover vaga.' }); }
});

// Criação de executivo por um gestor de corporação (client_admin) ou pelo Master (admin).
// Cria (ou atualiza) o acesso de login do colaborador (tabela users, role
// 'autonomous'), usando o e-mail cadastrado como usuário. Reaproveitado na
// criação e na edição do executivo. Nunca derruba a operação principal —
// se a credencial falhar (ex.: e-mail já usado por outra conta), devolve
// um aviso para ser anexado à mensagem de sucesso do cadastro/edição.
async function upsertAcessoColaborador(employeeId, name, email, phone, password, companyId) {
    if (!password) return null;
    if (!email) return 'Informe um e-mail para criar o acesso do colaborador.';
    try {
        const hash = await bcrypt.hash(password, 10);
        const existente = await dbGet(`SELECT id FROM users WHERE employee_id = ?`, [employeeId]);
        if (existente) {
            await new Promise((resolve, reject) => db.run(
                `UPDATE users SET name = ?, email = ?, password = ?, company_id = ? WHERE id = ?`,
                [name, email, hash, companyId, existente.id], (err) => err ? reject(err) : resolve()
            ));
        } else {
            const conflito = await dbGet(`SELECT id FROM users WHERE email = ?`, [email]);
            if (conflito) return 'Não foi possível criar o acesso: este e-mail já está em uso por outra conta.';
            await new Promise((resolve, reject) => db.run(
                `INSERT INTO users (name, email, password, company_id, employee_id, role) VALUES (?, ?, ?, ?, ?, 'autonomous')`,
                [name, email, hash, companyId, employeeId], (err) => err ? reject(err) : resolve()
            ));
            // Primeiro login deste colaborador: começa restrito só a "PDI" (mais
            // Página Inicial/Meu Perfil, sempre visíveis) — quem cadastrou libera
            // o resto depois em "Permissões" no cartão do colaborador. Só define
            // esse padrão quando o colaborador ainda não tinha nada configurado,
            // pra não sobrescrever uma permissão já ajustada manualmente.
            const atual = await dbGet(`SELECT enabled_modules FROM employees WHERE id = ?`, [employeeId]);
            if (atual && !atual.enabled_modules) {
                await new Promise((resolve, reject) => db.run(
                    `UPDATE employees SET enabled_modules = ? WHERE id = ?`,
                    [JSON.stringify(['pdi']), employeeId], (err) => err ? reject(err) : resolve()
                ));
            }
        }
        return null;
    } catch (e) {
        return 'Não foi possível criar/atualizar o acesso de login (' + e.message + ').';
    }
}

app.post('/api/employees', requireRole('admin', 'client_admin'), async (req, res) => {
    let { company_id, name, role, email, phone, performance_level, executive_phase, progress_percentage, disc_profile, photo_url, password } = req.body;

    // client_admin só pode cadastrar executivos dentro da própria corporação
    const finalCompanyId = req.user.role === 'client_admin'
        ? req.user.companyId
        : ((!company_id || company_id === "" || company_id === "null") ? null : company_id);

    // Respeita o limite de créditos de funcionário do plano contratado pela empresa (quando houver um plano com limite definido)
    if (finalCompanyId) {
        const empresa = await dbGet(`SELECT c.*, p.max_employees as planMaxEmployees, p.name as planName FROM companies c LEFT JOIN plans p ON c.plan_id = p.id WHERE c.id = ?`, [finalCompanyId]);
        if (empresa && empresa.planMaxEmployees !== null && empresa.planMaxEmployees !== undefined) {
            const contagem = await dbGet(`SELECT COUNT(*) as total FROM employees WHERE company_id = ?`, [finalCompanyId]);
            if (contagem.total >= empresa.planMaxEmployees) {
                return res.status(400).json({ error: `Sem créditos de funcionário disponíveis: limite de ${empresa.planMaxEmployees} colaboradores do plano ${empresa.planName || ''} atingido. Faça upgrade do plano para cadastrar mais.` });
            }
        }
    }

    db.run(`INSERT INTO employees (company_id, name, role, email, phone, performance_level, executive_phase, progress_percentage, disc_profile, photo_url) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
        [finalCompanyId, name, role, email, phone, performance_level || 'Em Desenvolvimento', executive_phase || 'Fase 1: Diagnóstico', progress_percentage || 10, disc_profile || 'A definir', photo_url], async function(err) {
            if(err) return res.status(400).json({ error: err.message });
            const employeeId = this.lastID;
            const avisoAcesso = await upsertAcessoColaborador(employeeId, name, email, phone, password, finalCompanyId);
            res.json({ message: avisoAcesso ? `Executivo registado com sucesso! Porém: ${avisoAcesso}` : 'Executivo registado com sucesso!' });
        });
});

app.put('/api/employees/:id', requireRole('admin', 'client_admin'), ensureEmployeeAccess(req => req.params.id), (req, res) => {
    let { company_id, name, role, email, phone, performance_level, executive_phase, progress_percentage, disc_profile, photo_url, password } = req.body;
    const finalCompanyId = req.user.role === 'client_admin'
        ? req.user.companyId
        : ((!company_id || company_id === "" || company_id === "null") ? null : company_id);
    db.run(`UPDATE employees SET company_id = ?, name = ?, role = ?, email = ?, phone = ?, performance_level = ?, executive_phase = ?, progress_percentage = ?, disc_profile = ?, photo_url = ? WHERE id = ?`,
        [finalCompanyId, name, role, email, phone, performance_level, executive_phase, progress_percentage, disc_profile, photo_url, req.params.id], async (err) => {
            if(err) return res.status(400).json({ error: err.message });
            const avisoAcesso = await upsertAcessoColaborador(req.params.id, name, email, phone, password, finalCompanyId);
            res.json({ message: avisoAcesso ? `Atualizado com sucesso! Porém: ${avisoAcesso}` : 'Atualizado com sucesso!' });
        });
});

// Movimentação de fase no Pipeline de Desenvolvimento: separado do PUT geral
// porque aqui exigimos o motivo da mudança e a frase que o colaborador vai
// ver, além de registrar tudo no histórico (employee_phase_history) e avisar
// o colaborador (se ele tiver acesso próprio) com a frase definida.
app.put('/api/employees/:id/phase', requireRole('admin', 'client_admin'), ensureEmployeeAccess(req => req.params.id), async (req, res) => {
    const { executive_phase, progress_percentage, current_challenge, target_role, reason, message_to_employee } = req.body;
    if (!executive_phase) return res.status(400).json({ error: 'Informe a fase executiva.' });
    try {
        const atual = await dbGet(`SELECT executive_phase FROM employees WHERE id = ?`, [req.params.id]);
        if (!atual) return res.status(404).json({ error: 'Executivo não encontrado.' });
        const mudouFase = atual.executive_phase !== executive_phase;

        await new Promise((resolve, reject) => db.run(
            `UPDATE employees SET executive_phase = ?, progress_percentage = ?, current_challenge = ?, target_role = ? WHERE id = ?`,
            [executive_phase, progress_percentage || 0, current_challenge || '', target_role || '', req.params.id],
            (err) => err ? reject(err) : resolve()
        ));

        if (mudouFase) {
            const quemMudou = await dbGet(`SELECT name FROM users WHERE id = ?`, [req.user.userId]);
            await new Promise((resolve, reject) => db.run(
                `INSERT INTO employee_phase_history (employee_id, from_phase, to_phase, reason, message_to_employee, changed_by_name) VALUES (?, ?, ?, ?, ?, ?)`,
                [req.params.id, atual.executive_phase, executive_phase, reason || '', message_to_employee || '', (quemMudou && quemMudou.name) || ''],
                (err) => err ? reject(err) : resolve()
            ));
            if (message_to_employee) {
                notificarPorEmployeeId(req.params.id, 'Você avançou de fase! 🎉', message_to_employee, 'pipelineDesenvolvimento');
            }
        }
        res.json({ message: mudouFase ? 'Fase atualizada e colaborador avisado!' : 'Progresso atualizado!' });
    } catch (e) {
        res.status(500).json({ error: 'Erro ao atualizar a fase do executivo.' });
    }
});

app.get('/api/employees/:id/phase-history', ensureEmployeeAccess(req => req.params.id), (req, res) => {
    db.all(`SELECT * FROM employee_phase_history WHERE employee_id = ? ORDER BY id DESC`, [req.params.id], (err, rows) => res.json(rows || []));
});

// Permissões de acesso do PRÓPRIO colaborador (o que ele vê no menu quando
// entra com o login individual dele). null = sem restrição (vê tudo que o
// módulo da empresa libera); um array restringe só a esse colaborador,
// independente do que a empresa como um todo tem liberado.
app.put('/api/employees/:id/permissions', requireRole('admin', 'client_admin'), ensureEmployeeAccess(req => req.params.id), async (req, res) => {
    const { enabled_modules } = req.body;
    const valor = Array.isArray(enabled_modules) ? JSON.stringify(enabled_modules) : null;
    db.run(`UPDATE employees SET enabled_modules = ? WHERE id = ?`, [valor, req.params.id], function (err) {
        if (err) return res.status(400).json({ error: 'Erro ao salvar permissões.' });
        if (this.changes === 0) return res.status(404).json({ error: 'Executivo não encontrado.' });
        res.json({ message: 'Permissões de acesso atualizadas!' });
    });
});

/* ==========================================================
   METAS DO COLABORADOR (mensais, até 5 por colaborador)
   Cada meta é medida em dinheiro (R$), número (unidade) ou
   percentual (%). O gestor (client_admin) ou o Master cadastram;
   o próprio colaborador (autonomous) só visualiza e reporta o
   quanto já atingiu.
   ========================================================== */
app.get('/api/employee-goals', (req, res) => {
    const scope = buildScope(req);
    if (scope.deny) return res.status(403).json({ error: 'Perfil sem permissão.' });
    let q = `SELECT g.*, e.name as execName, e.photo_url as execPhoto FROM employee_goals g JOIN employees e ON g.employee_id = e.id`;
    const conditions = []; const params = [];
    if (scope.companyId) { conditions.push('e.company_id = ?'); params.push(scope.companyId); }
    if (scope.employeeId) { conditions.push('g.employee_id = ?'); params.push(scope.employeeId); }
    if (!scope.employeeId && req.query.employee_id) { conditions.push('g.employee_id = ?'); params.push(req.query.employee_id); }
    if (conditions.length) q += ' WHERE ' + conditions.join(' AND ');
    q += ' ORDER BY g.month DESC, g.id DESC';
    db.all(q, params, (err, rows) => res.json(rows || []));
});

// Ranking de metas: melhores do mês (mês corrente) e melhores acumulado do ano
// (ano corrente), com base no % médio atingido das metas de cada colaborador.
// Usado no Painel da Empresa para reconhecer quem está performando melhor.
app.get('/api/employee-goals/ranking', (req, res) => {
    const scope = buildScope(req);
    if (scope.deny) return res.status(403).json({ error: 'Perfil sem permissão.' });
    const companyId = scope.companyId || req.query.company_id;
    if (!companyId) return res.status(400).json({ error: 'Informe a empresa.' });
    const agora = new Date();
    const mesAtual = agora.toISOString().slice(0, 7); // YYYY-MM
    const anoAtual = String(agora.getFullYear());

    const montarRanking = (filtroMes) => new Promise((resolve, reject) => {
        const params = [companyId];
        let filtro = '';
        if (filtroMes === 'mes') { filtro = "AND g.month = ?"; params.push(mesAtual); }
        else { filtro = "AND g.month LIKE ?"; params.push(anoAtual + '-%'); }
        db.all(
            `SELECT e.id as employeeId, e.name, e.photo_url,
                AVG(CASE WHEN g.target_value > 0 THEN MIN(g.achieved_value / g.target_value, 1.5) * 100 ELSE 0 END) as percentualMedio,
                COUNT(g.id) as totalMetas
             FROM employee_goals g JOIN employees e ON e.id = g.employee_id
             WHERE e.company_id = ? ${filtro}
             GROUP BY e.id ORDER BY percentualMedio DESC LIMIT 10`,
            params, (err, rows) => err ? reject(err) : resolve((rows || []).map(r => ({ ...r, percentualMedio: Math.round(r.percentualMedio || 0) })))
        );
    });

    Promise.all([montarRanking('mes'), montarRanking('ano')])
        .then(([doMes, doAno]) => res.json({ mes: mesAtual, doMes, doAno }))
        .catch(() => res.status(500).json({ error: 'Erro ao calcular o ranking.' }));
});

// Consolidado de metas: visão única com todas as metas de todos os
// colaboradores da empresa, filtrável por mês específico ou por ano inteiro
// (nesse caso soma/agrega todos os meses daquele ano por colaborador+meta).
app.get('/api/employee-goals/consolidado', (req, res) => {
    const scope = buildScope(req);
    if (scope.deny) return res.status(403).json({ error: 'Perfil sem permissão.' });
    const companyId = scope.companyId || req.query.company_id;
    if (!companyId) return res.status(400).json({ error: 'Informe a empresa.' });
    const agora = new Date();
    const periodo = req.query.periodo === 'ano' ? 'ano' : 'mes';
    const mes = req.query.mes || agora.toISOString().slice(0, 7);
    const ano = req.query.ano || String(agora.getFullYear());

    let filtro = '', params = [companyId];
    if (periodo === 'mes') { filtro = 'AND g.month = ?'; params.push(mes); }
    else { filtro = 'AND g.month LIKE ?'; params.push(ano + '-%'); }

    db.all(
        `SELECT g.*, e.name as execName, e.photo_url as execPhoto
         FROM employee_goals g JOIN employees e ON e.id = g.employee_id
         WHERE e.company_id = ? ${filtro}
         ORDER BY e.name ASC, g.month ASC`,
        params,
        (err, linhas) => {
            if (err) return res.status(500).json({ error: 'Erro ao carregar o consolidado de metas.' });
            const metas = linhas || [];
            const totalMetas = metas.length;
            const atingidas = metas.filter(m => m.target_value > 0 && (m.achieved_value / m.target_value) >= 1).length;
            const mediaPercentual = totalMetas > 0
                ? Math.round(metas.reduce((soma, m) => soma + (m.target_value > 0 ? Math.min(m.achieved_value / m.target_value, 1.5) * 100 : 0), 0) / totalMetas)
                : 0;
            const porColaborador = {};
            metas.forEach(m => {
                const chave = m.employee_id;
                if (!porColaborador[chave]) porColaborador[chave] = { employeeId: m.employee_id, name: m.execName, photo_url: m.execPhoto, metas: [] };
                porColaborador[chave].metas.push(m);
            });
            res.json({
                periodo, mes, ano,
                resumo: { totalMetas, atingidas, naoAtingidas: totalMetas - atingidas, mediaPercentual },
                porColaborador: Object.values(porColaborador)
            });
        }
    );
});

app.post('/api/employee-goals', requireRole('admin', 'client_admin'), ensureEmployeeAccess(req => req.body.employee_id), async (req, res) => {
    const { employee_id, title, month, goal_type, target_value } = req.body;
    if (!employee_id || !title || !month) return res.status(400).json({ error: 'Preencha o colaborador, o título e o mês da meta.' });
    if (!['dinheiro', 'numero', 'percentual'].includes(goal_type)) return res.status(400).json({ error: 'Tipo de meta inválido.' });
    try {
        const contagem = await dbGet(`SELECT COUNT(*) as total FROM employee_goals WHERE employee_id = ?`, [employee_id]);
        if (contagem.total >= 5) return res.status(400).json({ error: 'Limite de 5 metas por colaborador atingido. Remova ou edite uma meta existente.' });
        db.run(`INSERT INTO employee_goals (employee_id, title, month, goal_type, target_value, achieved_value, status, updated_at) VALUES (?, ?, ?, ?, ?, 0, 'Em Andamento', CURRENT_TIMESTAMP)`,
            [employee_id, title, month, goal_type, Number(target_value) || 0], function (err) {
                if (err) return res.status(400).json({ error: err.message });
                notificarPorEmployeeId(employee_id, 'Nova meta cadastrada', `${title} (${month})`, 'employees');
                res.json({ message: 'Meta cadastrada!' });
            });
    } catch (e) { res.status(500).json({ error: 'Erro ao cadastrar meta.' }); }
});

app.put('/api/employee-goals/:id', requireRole('admin', 'client_admin'), ensureRecordAccess('employee_goals'), (req, res) => {
    const { title, month, goal_type, target_value, achieved_value, status } = req.body;
    if (!['dinheiro', 'numero', 'percentual'].includes(goal_type)) return res.status(400).json({ error: 'Tipo de meta inválido.' });
    db.run(`UPDATE employee_goals SET title = ?, month = ?, goal_type = ?, target_value = ?, achieved_value = ?, status = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?`,
        [title, month, goal_type, Number(target_value) || 0, Number(achieved_value) || 0, status || 'Em Andamento', req.params.id], (err) => {
            if (err) return res.status(400).json({ error: err.message });
            res.json({ message: 'Meta atualizada!' });
        });
});

// O próprio colaborador (autonomous) reporta o andamento da sua meta
app.put('/api/employee-goals/:id/progress', requireRole('autonomous'), ensureRecordAccess('employee_goals'), (req, res) => {
    const { achieved_value } = req.body;
    db.run(`UPDATE employee_goals SET achieved_value = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?`, [Number(achieved_value) || 0, req.params.id], (err) => {
        if (err) return res.status(400).json({ error: err.message });
        res.json({ message: 'Progresso atualizado!' });
    });
});

app.delete('/api/employee-goals/:id', requireRole('admin', 'client_admin'), ensureRecordAccess('employee_goals'), (req, res) => {
    db.run(`DELETE FROM employee_goals WHERE id = ?`, [req.params.id], () => res.json({ message: 'Removida!' }));
});

// Autoatendimento de identidade/rede/currículo — o próprio executivo (autonomous),
// o gestor da empresa (client_admin) ou o Master podem atualizar estes campos,
// sem mexer nos campos administrativos do cadastro principal.
app.put('/api/employees/:id/profile', requireRole('admin', 'client_admin', 'autonomous'), ensureEmployeeAccess(req => req.params.id), (req, res) => {
    const { public_bio, linkedin_url, show_in_directory, resume_url, looking_for_opportunity, desired_role } = req.body;
    db.run(
        `UPDATE employees SET public_bio = ?, linkedin_url = ?, show_in_directory = ?, resume_url = ?, looking_for_opportunity = ?, desired_role = ? WHERE id = ?`,
        [public_bio || '', linkedin_url || '', show_in_directory ? 1 : 0, resume_url || '', looking_for_opportunity ? 1 : 0, desired_role || '', req.params.id],
        (err) => {
            if (err) return res.status(400).json({ error: err.message });
            res.json({ message: 'Perfil atualizado!' });
        }
    );
});

app.delete('/api/employees/:id', requireRole('admin', 'client_admin'), ensureEmployeeAccess(req => req.params.id), (req, res) => {
    db.run(`DELETE FROM users WHERE employee_id = ?`, [req.params.id], () => {
        db.run(`DELETE FROM employees WHERE id = ?`, [req.params.id], () => res.json({ message: 'Removido!' }));
    });
});

app.get('/api/pdi', (req, res) => {
    const scope = buildScope(req);
    if (scope.deny) return res.status(403).json({ error: 'Perfil sem permissão.' });

    let q = `SELECT p.*, e.name as execName, e.email as execEmail, e.phone as execPhone, e.photo_url as execPhoto FROM pd_plans p JOIN employees e ON p.employee_id = e.id`;
    const conditions = [];
    const params = [];
    if (scope.companyId) { conditions.push('e.company_id = ?'); params.push(scope.companyId); }
    if (scope.employeeId) { conditions.push('p.employee_id = ?'); params.push(scope.employeeId); }
    if (!scope.employeeId && req.query.employee_id) { conditions.push('p.employee_id = ?'); params.push(req.query.employee_id); }
    if (conditions.length) q += ' WHERE ' + conditions.join(' AND ');
    db.all(q, params, (err, rows) => res.json(rows || []));
});

app.get('/api/export/pdi', requireRole('admin', 'client_admin'), (req, res) => {
    const scope = buildScope(req);
    let q = `SELECT p.*, e.name as execName FROM pd_plans p JOIN employees e ON p.employee_id = e.id`;
    const params = [];
    if (scope.companyId) { q += ' WHERE e.company_id = ?'; params.push(scope.companyId); }
    db.all(q, params, (err, rows) => {
        if (err) return res.status(500).json({ error: 'Erro' });
        let csv = "Executivo;Objetivo;Plano de Acao;Prazo;Estado\n";
        rows.forEach(r => { csv += `"${csvSafe(r.execName)}";"${csvSafe(r.objective)}";"${csvSafe(r.action_plan)}";"${csvSafe(r.deadline)}";"${csvSafe(r.status)}"\n`; });
        res.setHeader('Content-Type', 'text/csv; charset=utf-8');
        res.setHeader('Content-Disposition', 'attachment; filename=pdis_impulsionar.csv');
        res.status(200).send(Buffer.from('\uFEFF' + csv, 'utf-8'));
    });
});

app.post('/api/pdi', requireRole('admin', 'client_admin'), ensureEmployeeAccess(req => req.body.employee_id), (req, res) => {
    const { employee_id, objective, action_plan, deadline, status } = req.body;
    db.run(`INSERT INTO pd_plans (employee_id, objective, action_plan, deadline, status) VALUES (?, ?, ?, ?, ?)`, [employee_id, objective, action_plan, deadline, status || 'Em Andamento'], () => {
        notificarPorEmployeeId(employee_id, 'Novo PDI criado', objective, 'pdi');
        res.json({ message: 'PDI criado!' });
    });
});
app.put('/api/pdi/:id', requireRole('admin', 'client_admin'), ensureRecordAccess('pd_plans'), (req, res) => {
    const { objective, action_plan, deadline, status } = req.body;
    db.run(
        `UPDATE pd_plans SET objective = ?, action_plan = ?, deadline = ?, status = ? WHERE id = ?`,
        [objective, action_plan, deadline, status, req.params.id],
        (err) => {
            if (err) return res.status(400).json({ error: err.message });
            if (status === 'Concluído') {
                darPontos(req.user.userId, 50, 'PDI concluído');
                db.get(`SELECT employee_id FROM pd_plans WHERE id = ?`, [req.params.id], (e, row) => {
                    if (!e && row) notificarPorEmployeeId(row.employee_id, 'PDI concluído', objective, 'pdi');
                });
            }
            res.json({ message: 'PDI atualizado!' });
        }
    );
});
app.delete('/api/pdi/:id', requireRole('admin', 'client_admin'), ensureRecordAccess('pd_plans'), (req, res) => { db.run(`DELETE FROM pd_plans WHERE id = ?`, [req.params.id], () => res.json({ message: 'Removido!' })); });

// Ações de um PDI: quebram o plano em passos concretos (não iniciada / em
// andamento / concluída) em vez de um status único — dá a visão de "quantas
// ações estão em andamento, não iniciadas e fechadas" por colaborador.
app.get('/api/pdi/:id/actions', ensureRecordAccess('pd_plans'), (req, res) => {
    db.all(`SELECT * FROM pdi_actions WHERE pd_plan_id = ? ORDER BY id ASC`, [req.params.id], (err, rows) => res.json(rows || []));
});

app.post('/api/pdi/:id/actions', requireRole('admin', 'client_admin'), ensureRecordAccess('pd_plans'), (req, res) => {
    const { description } = req.body;
    if (!description) return res.status(400).json({ error: 'Descreva a ação.' });
    db.run(`INSERT INTO pdi_actions (pd_plan_id, description, status) VALUES (?, ?, 'Não iniciada')`, [req.params.id, description], function (err) {
        if (err) return res.status(400).json({ error: err.message });
        res.json({ message: 'Ação adicionada!', id: this.lastID });
    });
});

// Valida acesso a uma pdi_action a partir do pd_plan (e do employee) a que ela pertence.
function ensurePdiActionAccess() {
    return async (req, res, next) => {
        if (req.user.role === 'admin') return next();
        try {
            const acao = await dbGet(`SELECT pd_plan_id FROM pdi_actions WHERE id = ?`, [req.params.id]);
            if (!acao) return res.status(404).json({ error: 'Ação não encontrada.' });
            const plano = await dbGet(`SELECT employee_id FROM pd_plans WHERE id = ?`, [acao.pd_plan_id]);
            if (!plano) return res.status(404).json({ error: 'PDI não encontrado.' });
            return ensureEmployeeAccess(() => plano.employee_id)(req, res, next);
        } catch (e) { return res.status(500).json({ error: 'Erro ao validar permissão.' }); }
    };
}

app.put('/api/pdi-actions/:id', requireRole('admin', 'client_admin'), ensurePdiActionAccess(), (req, res) => {
    const { description, status } = req.body;
    db.run(`UPDATE pdi_actions SET description = ?, status = ? WHERE id = ?`, [description, status, req.params.id], (err) => {
        if (err) return res.status(400).json({ error: err.message });
        res.json({ message: 'Ação atualizada!' });
    });
});

app.delete('/api/pdi-actions/:id', requireRole('admin', 'client_admin'), ensurePdiActionAccess(), (req, res) => {
    db.run(`DELETE FROM pdi_actions WHERE id = ?`, [req.params.id], () => res.json({ message: 'Removida!' }));
});

// Upload gen\u00E9rico de arquivo (v\u00EDdeo ou imagem) \u2014 usado pela Academy e pelo
// v\u00EDdeo de bio dos mentores. Retorna a URL p\u00FAblica do arquivo salvo.
app.post('/api/upload', requireRole('admin', 'client_admin', 'mentor', 'candidate'), upload.single('file'), (req, res) => {
    if (!req.file) return res.status(400).json({ error: 'Nenhum arquivo recebido.' });
    res.json({ url: '/uploads/' + req.file.filename, originalName: req.file.originalname });
});

app.get('/api/videos', async (req, res) => {
    try {
        const rows = await dbAll(
            `SELECT v.*, m.name as mentorName,
                    (SELECT COUNT(*) FROM video_likes vl WHERE vl.video_id = v.id) as totalLikes,
                    (SELECT COUNT(*) FROM video_comments vc WHERE vc.video_id = v.id) as totalComments,
                    (SELECT COUNT(*) FROM video_likes vl2 WHERE vl2.video_id = v.id AND vl2.user_id = ?) as curtiuEu
             FROM video_lessons v LEFT JOIN mentors m ON v.mentor_id = m.id
             ORDER BY v.created_at DESC`,
            [req.user.userId]
        );
        res.json(rows.map(r => ({ ...r, curtiuEu: !!r.curtiuEu })));
    } catch (e) {
        res.status(500).json({ error: 'Erro ao carregar videoaulas.' });
    }
});

app.get('/api/export/videos', requireRole('admin'), (req, res) => {
    db.all(`SELECT * FROM video_lessons`, [], (err, rows) => {
        if (err) return res.status(500).json({ error: 'Erro' });
        let csv = "Titulo;Trilha;Duracao;Link\n";
        rows.forEach(r => { csv += `"${csvSafe(r.title)}";"${csvSafe(r.module_name)}";"${csvSafe(r.duration)}";"${csvSafe(r.video_url)}"\n`; });
        res.setHeader('Content-Type', 'text/csv; charset=utf-8');
        res.setHeader('Content-Disposition', 'attachment; filename=academy_videos_impulsionar.csv');
        res.status(200).send(Buffer.from('\uFEFF' + csv, 'utf-8'));
    });
});

app.post('/api/videos', requireRole('admin', 'client_admin', 'mentor'), async (req, res) => {
    const { title, module_name, video_url, duration, description, thumbnail_url } = req.body;
    if (!title || !video_url) return res.status(400).json({ error: 'T\u00EDtulo e v\u00EDdeo s\u00E3o obrigat\u00F3rios.' });

    let mentorId = null, postedBy = 'Consultoria Impulsionar';
    if (req.user.role === 'mentor') {
        mentorId = req.user.mentorId;
        const mentor = await dbGet(`SELECT name FROM mentors WHERE id = ?`, [mentorId]);
        postedBy = mentor ? mentor.name : 'Mentor';
    }

    db.run(
        `INSERT INTO video_lessons (title, module_name, video_url, duration, description, thumbnail_url, mentor_id, posted_by) VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
        [title, module_name, video_url, duration, description || '', thumbnail_url || '', mentorId, postedBy],
        function (err) {
            if (err) return res.status(400).json({ error: err.message });
            res.json({ message: 'Publicado!', id: this.lastID });
        }
    );
});

// Garante que s\u00F3 o admin ou o pr\u00F3prio mentor que postou pode editar/excluir o v\u00EDdeo
function ensureVideoOwnerOrAdmin() {
    return async (req, res, next) => {
        if (req.user.role === 'admin' || req.user.role === 'client_admin') return next();
        if (req.user.role !== 'mentor') return res.status(403).json({ error: 'Perfil sem permiss\u00E3o.' });
        try {
            const video = await dbGet(`SELECT mentor_id FROM video_lessons WHERE id = ?`, [req.params.id]);
            if (!video) return res.status(404).json({ error: 'Videoaula n\u00E3o encontrada.' });
            if (String(video.mentor_id) !== String(req.user.mentorId)) {
                return res.status(403).json({ error: 'Voc\u00EA s\u00F3 pode editar seus pr\u00F3prios v\u00EDdeos.' });
            }
            next();
        } catch (e) { res.status(500).json({ error: 'Erro ao validar permiss\u00E3o.' }); }
    };
}

app.put('/api/videos/:id', requireRole('admin', 'client_admin', 'mentor'), ensureVideoOwnerOrAdmin(), (req, res) => {
    const { title, module_name, video_url, duration, description, thumbnail_url } = req.body;
    db.run(
        `UPDATE video_lessons SET title = ?, module_name = ?, video_url = ?, duration = ?, description = ?, thumbnail_url = ? WHERE id = ?`,
        [title, module_name, video_url, duration, description || '', thumbnail_url || '', req.params.id],
        (err) => {
            if (err) return res.status(400).json({ error: err.message });
            res.json({ message: 'Atualizado com sucesso!' });
        }
    );
});

app.delete('/api/videos/:id', requireRole('admin', 'client_admin', 'mentor'), ensureVideoOwnerOrAdmin(), (req, res) => {
    db.run(`DELETE FROM video_likes WHERE video_id = ?`, [req.params.id], () => {
        db.run(`DELETE FROM video_comments WHERE video_id = ?`, [req.params.id], () => {
            db.run(`DELETE FROM video_lessons WHERE id = ?`, [req.params.id], () => res.json({ message: 'Removido!' }));
        });
    });
});

// ---------- Painel social: curtidas e coment\u00E1rios ----------
app.post('/api/videos/:id/like', async (req, res) => {
    try {
        const jaCurtiu = await dbGet(`SELECT id FROM video_likes WHERE video_id = ? AND user_id = ?`, [req.params.id, req.user.userId]);
        let curtiu;
        if (jaCurtiu) {
            await new Promise((resolve, reject) => db.run(`DELETE FROM video_likes WHERE id = ?`, [jaCurtiu.id], (err) => err ? reject(err) : resolve()));
            curtiu = false;
        } else {
            await new Promise((resolve, reject) => db.run(`INSERT INTO video_likes (video_id, user_id) VALUES (?, ?)`, [req.params.id, req.user.userId], (err) => err ? reject(err) : resolve()));
            curtiu = true;
        }
        const contagem = await dbGet(`SELECT COUNT(*) as total FROM video_likes WHERE video_id = ?`, [req.params.id]);
        res.json({ curtiu, totalLikes: contagem.total });
    } catch (e) {
        res.status(500).json({ error: 'Erro ao curtir v\u00EDdeo.' });
    }
});

app.get('/api/videos/:id/comments', async (req, res) => {
    try {
        const comentarios = await dbAll(
            `SELECT vc.*, u.name as userName FROM video_comments vc JOIN users u ON vc.user_id = u.id
             WHERE vc.video_id = ? ORDER BY vc.created_at ASC`,
            [req.params.id]
        );
        res.json(comentarios);
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar coment\u00E1rios.' }); }
});

app.post('/api/videos/:id/comments', (req, res) => {
    const { comment } = req.body;
    if (!comment || !comment.trim()) return res.status(400).json({ error: 'Escreva um coment\u00E1rio.' });
    db.run(
        `INSERT INTO video_comments (video_id, user_id, comment) VALUES (?, ?, ?)`,
        [req.params.id, req.user.userId, comment.trim()],
        function (err) {
            if (err) return res.status(400).json({ error: err.message });
            darPontos(req.user.userId, 5, 'Coment\u00E1rio na Academy Digital');
            res.json({ message: 'Coment\u00E1rio publicado!', id: this.lastID });
        }
    );
});

app.delete('/api/comments/:id', async (req, res) => {
    try {
        const comentario = await dbGet(`SELECT user_id FROM video_comments WHERE id = ?`, [req.params.id]);
        if (!comentario) return res.status(404).json({ error: 'Coment\u00E1rio n\u00E3o encontrado.' });
        if (req.user.role !== 'admin' && String(comentario.user_id) !== String(req.user.userId)) {
            return res.status(403).json({ error: 'Voc\u00EA s\u00F3 pode remover seus pr\u00F3prios coment\u00E1rios.' });
        }
        db.run(`DELETE FROM video_comments WHERE id = ?`, [req.params.id], () => res.json({ message: 'Coment\u00E1rio removido!' }));
    } catch (e) { res.status(500).json({ error: 'Erro ao remover coment\u00E1rio.' }); }
});

app.get('/api/assessments', (req, res) => {
    const scope = buildScope(req);
    if (scope.deny) return res.status(403).json({ error: 'Perfil sem permissão.' });

    let q = `SELECT a.*, e.name as employeeName, e.photo_url as execPhoto, c.name as companyName FROM assessments a LEFT JOIN employees e ON a.employee_id = e.id LEFT JOIN companies c ON e.company_id = c.id`;
    const conditions = [];
    const params = [];
    if (scope.companyId) { conditions.push('e.company_id = ?'); params.push(scope.companyId); }
    if (scope.employeeId) { conditions.push('a.employee_id = ?'); params.push(scope.employeeId); }
    if (!scope.employeeId && req.query.employee_id) { conditions.push('a.employee_id = ?'); params.push(req.query.employee_id); }
    if (conditions.length) q += ' WHERE ' + conditions.join(' AND ');
    db.all(q, params, (err, rows) => res.json(rows || []));
});

app.get('/api/export/assessments', requireRole('admin', 'client_admin'), (req, res) => {
    const scope = buildScope(req);
    let q = `SELECT a.*, e.name as employeeName, c.name as companyName FROM assessments a LEFT JOIN employees e ON a.employee_id = e.id LEFT JOIN companies c ON e.company_id = c.id`;
    const params = [];
    if (scope.companyId) { q += ' WHERE e.company_id = ?'; params.push(scope.companyId); }
    db.all(q, params, (err, rows) => {
        if (err) return res.status(500).json({ error: 'Erro' });
        let csv = "Executivo;Empresa;Competencia;Pontuacao;Parecer;Data\n";
        rows.forEach(r => { csv += `"${csvSafe(r.employeeName)}";"${csvSafe(r.companyName || 'Autónomo')}";"${csvSafe(r.leadership_competence)}";"${csvSafe(r.score)}";"${csvSafe(r.status)}";"${csvSafe(r.date)}"\n`; });
        res.setHeader('Content-Type', 'text/csv; charset=utf-8');
        res.setHeader('Content-Disposition', 'attachment; filename=assessments_impulsionar.csv');
        res.status(200).send(Buffer.from('\uFEFF' + csv, 'utf-8'));
    });
});

app.post('/api/assessments', requireRole('admin', 'client_admin'), ensureEmployeeAccess(req => req.body.employee_id), (req, res) => {
    const { employee_id, leadership_competence, score } = req.body;
    const s = Number(score);
    const status = s >= 80 ? 'High Potential (Aprovado)' : s >= 60 ? 'Em Desenvolvimento' : 'PDI Crítico';
    db.run(`INSERT INTO assessments (employee_id, leadership_competence, score, status) VALUES (?, ?, ?, ?)`, [employee_id, leadership_competence, s, status], () => {
        darPontos(req.user.userId, 30, 'Competence Check registrado');
        res.json({ message: 'Registado!' });
    });
});
app.put('/api/assessments/:id', requireRole('admin', 'client_admin'), ensureRecordAccess('assessments'), (req, res) => {
    const { leadership_competence, score } = req.body;
    const s = Number(score);
    const status = s >= 80 ? 'High Potential (Aprovado)' : s >= 60 ? 'Em Desenvolvimento' : 'PDI Crítico';
    db.run(
        `UPDATE assessments SET leadership_competence = ?, score = ?, status = ? WHERE id = ?`,
        [leadership_competence, s, status, req.params.id],
        (err) => {
            if (err) return res.status(400).json({ error: err.message });
            res.json({ message: 'Assessment atualizado!' });
        }
    );
});
app.delete('/api/assessments/:id', requireRole('admin', 'client_admin'), ensureRecordAccess('assessments'), (req, res) => { db.run(`DELETE FROM assessments WHERE id = ?`, [req.params.id], () => res.json({ message: 'Removido!' })); });

app.get('/api/mentors', (req, res) => { db.all(`SELECT * FROM mentors`, [], (err, rows) => res.json(rows || [])); });

// Ao cadastrar o mentor, opcionalmente já cria o login dele (role 'mentor'),
// permitindo que ele mesmo entre na plataforma e publique vídeos na Academy
// com autoridade — currículo e vídeo de bio aparecem no perfil público dele.
app.post('/api/mentors', requireRole('admin'), (req, res) => {
    const { name, specialty, email, available_days, resume, bio_video_url, photo_url, password } = req.body;
    db.run(`INSERT INTO mentors (name, specialty, email, available_days, resume, bio_video_url, photo_url) VALUES (?, ?, ?, ?, ?, ?, ?)`,
        [name, specialty, email, available_days, resume, bio_video_url, photo_url], async function (err) {
            if (err) return res.status(400).json({ error: err.message });
            const mentorId = this.lastID;
            if (email && password) {
                try {
                    const hash = await bcrypt.hash(password, 10);
                    db.run(`INSERT INTO users (name, email, password, company_id, mentor_id, role) VALUES (?, ?, ?, NULL, ?, 'mentor')`,
                        [name, email, hash, mentorId], (userErr) => {
                            if (userErr) return res.json({ message: 'Mentor cadastrado, mas não foi possível criar o login (e-mail já em uso).', id: mentorId });
                            res.json({ message: 'Mentor cadastrado com acesso próprio à plataforma!', id: mentorId });
                        }
                    );
                } catch (e) { res.json({ message: 'Mentor cadastrado, mas houve erro ao criar o login.', id: mentorId }); }
            } else {
                res.json({ message: 'Mentor cadastrado!', id: mentorId });
            }
        });
});

// ---------- Horários disponíveis do mentor (datas específicas do calendário) ----------
app.get('/api/mentors/:id/availability', (req, res) => {
    db.all(
        `SELECT * FROM mentor_availability WHERE mentor_id = ? ORDER BY specific_date ASC, start_time ASC`,
        [req.params.id],
        (err, rows) => res.json(rows || [])
    );
});

app.put('/api/mentors/:id/availability', requireRole('admin', 'mentor'), async (req, res) => {
    if (req.user.role === 'mentor' && String(req.params.id) !== String(req.user.mentorId)) {
        return res.status(403).json({ error: 'Você só pode editar seus próprios horários.' });
    }
    const { slots } = req.body;
    if (!Array.isArray(slots)) return res.status(400).json({ error: 'Lista de horários inválida.' });
    for (const s of slots) {
        if (!s.specific_date || !/^\d{4}-\d{2}-\d{2}$/.test(s.specific_date) || !s.start_time || !s.end_time) {
            return res.status(400).json({ error: 'Cada horário precisa de uma data válida, início e fim.' });
        }
    }
    try {
        await new Promise((resolve, reject) => db.run(`DELETE FROM mentor_availability WHERE mentor_id = ?`, [req.params.id], (err) => err ? reject(err) : resolve()));
        for (const s of slots) {
            const [ano, mes, dia] = s.specific_date.split('-').map(Number);
            const diaSemana = new Date(ano, mes - 1, dia).getDay();
            await new Promise((resolve, reject) => db.run(
                `INSERT INTO mentor_availability (mentor_id, specific_date, day_of_week, start_time, end_time) VALUES (?, ?, ?, ?, ?)`,
                [req.params.id, s.specific_date, diaSemana, s.start_time, s.end_time],
                (err) => err ? reject(err) : resolve()
            ));
        }
        res.json({ message: 'Horários disponíveis atualizados!' });
    } catch (e) {
        res.status(400).json({ error: 'Erro ao salvar horários: ' + e.message });
    }
});

app.put('/api/mentors/:id', requireRole('admin', 'mentor'), async (req, res) => {
    if (req.user.role === 'mentor' && String(req.params.id) !== String(req.user.mentorId)) {
        return res.status(403).json({ error: 'Você só pode editar o seu próprio perfil.' });
    }
    const { name, specialty, email, available_days, resume, bio_video_url, photo_url, password } = req.body;
    db.run(
        `UPDATE mentors SET name = ?, specialty = ?, email = ?, available_days = ?, resume = ?, bio_video_url = ?, photo_url = ? WHERE id = ?`,
        [name, specialty, email, available_days, resume, bio_video_url, photo_url, req.params.id],
        async (err) => {
            if (err) return res.status(400).json({ error: err.message });
            try {
                const usuarioVinculado = await dbGet(`SELECT id FROM users WHERE mentor_id = ?`, [req.params.id]);
                if (usuarioVinculado) {
                    if (password) {
                        const hash = await bcrypt.hash(password, 10);
                        db.run(`UPDATE users SET name = ?, email = ?, password = ? WHERE mentor_id = ?`, [name, email, hash, req.params.id], () => {});
                    } else {
                        db.run(`UPDATE users SET name = ?, email = ? WHERE mentor_id = ?`, [name, email, req.params.id], () => {});
                    }
                } else if (email && password && req.user.role === 'admin') {
                    const hash = await bcrypt.hash(password, 10);
                    db.run(`INSERT INTO users (name, email, password, company_id, mentor_id, role) VALUES (?, ?, ?, NULL, ?, 'mentor')`,
                        [name, email, hash, req.params.id], () => {});
                }
            } catch (e) { /* login opcional: falha aqui não impede a atualização do perfil */ }
            res.json({ message: 'Perfil de mentor atualizado!' });
        }
    );
});

app.delete('/api/mentors/:id', requireRole('admin'), (req, res) => {
    db.run(`DELETE FROM users WHERE mentor_id = ?`, [req.params.id], () => {
        db.run(`DELETE FROM mentor_availability WHERE mentor_id = ?`, [req.params.id], () => {
            db.run(`DELETE FROM mentors WHERE id = ?`, [req.params.id], () => res.json({ message: 'Removido!' }));
        });
    });
});

app.get('/api/mentorships', (req, res) => {
    const scope = buildScope(req);
    if (scope.deny) return res.status(403).json({ error: 'Perfil sem permissão.' });

    let q = `SELECT m.*, e.name as execName, e.email as execEmail, e.phone as execPhone, ment.name as mentorName, ment.specialty as mentorSpec FROM mentorships m JOIN employees e ON m.employee_id = e.id JOIN mentors ment ON m.mentor_id = ment.id`;
    const conditions = [];
    const params = [];
    if (scope.companyId) { conditions.push('e.company_id = ?'); params.push(scope.companyId); }
    if (scope.employeeId) { conditions.push('m.employee_id = ?'); params.push(scope.employeeId); }
    if (scope.mentorId) { conditions.push('m.mentor_id = ?'); params.push(scope.mentorId); }
    if (!scope.employeeId && req.query.employee_id) { conditions.push('m.employee_id = ?'); params.push(req.query.employee_id); }
    if (conditions.length) q += ' WHERE ' + conditions.join(' AND ');
    db.all(q, params, (err, rows) => res.json(rows || []));
});

app.get('/api/export/mentorships', requireRole('admin', 'client_admin'), (req, res) => {
    const scope = buildScope(req);
    let q = `SELECT m.*, e.name as execName, ment.name as mentorName FROM mentorships m JOIN employees e ON m.employee_id = e.id JOIN mentors ment ON m.mentor_id = ment.id`;
    const params = [];
    if (scope.companyId) { q += ' WHERE e.company_id = ?'; params.push(scope.companyId); }
    db.all(q, params, (err, rows) => {
        if (err) return res.status(500).json({ error: 'Erro' });
        let csv = "Executivo;Mentor;Data;Topicos;Atas;Estado\n";
        rows.forEach(r => { csv += `"${csvSafe(r.execName)}";"${csvSafe(r.mentorName)}";"${csvSafe(r.meeting_date)}";"${csvSafe(r.topics)}";"${csvSafe(r.minutes || '')}";"${csvSafe(r.status)}"\n`; });
        res.setHeader('Content-Type', 'text/csv; charset=utf-8');
        res.setHeader('Content-Disposition', 'attachment; filename=mentorias_impulsionar.csv');
        res.status(200).send(Buffer.from('\uFEFF' + csv, 'utf-8'));
    });
});

app.post('/api/mentorships', requireRole('admin', 'client_admin', 'autonomous'), ensureEmployeeAccess(req => req.body.employee_id), async (req, res) => {
    const { employee_id, mentor_id, meeting_date, topics, minutes, status, duration_minutes, participants } = req.body;
    if (!employee_id) return res.status(400).json({ error: 'Selecione o executivo.' });
    if (!mentor_id) return res.status(400).json({ error: 'Selecione o mentor.' });
    if (!meeting_date) return res.status(400).json({ error: 'Selecione um horário disponível para a reunião.' });
    if (!topics) return res.status(400).json({ error: 'Informe os tópicos da mentoria.' });
    const mentor = await dbGet(`SELECT name, email FROM mentors WHERE id = ?`, [mentor_id]);
    const executivo = await dbGet(`SELECT name, email FROM employees WHERE id = ?`, [employee_id]);
    const listaParticipantes = Array.isArray(participants) ? participants.filter(p => p && p.email) : [];
    const duracao = Number(duration_minutes) || 60;
    db.run(`INSERT INTO mentorships (employee_id, mentor_id, mentor_name, meeting_date, topics, minutes, status, duration_minutes, participants_json, ics_sequence) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 0)`,
        [employee_id, mentor_id, mentor ? mentor.name : '', meeting_date, topics, minutes || '', status || 'Agendada', duracao, JSON.stringify(listaParticipantes)], async function(err) {
            if(err) return res.status(400).json({ error: err.message });
            const novoId = this.lastID;
            const uid = `mentoria-${novoId}-${Date.now()}@impulsionar`;
            db.run(`UPDATE mentorships SET ics_uid = ? WHERE id = ?`, [uid, novoId], () => {});
            notificarPorEmployeeId(employee_id, 'Mentoria agendada', `Com ${mentor ? mentor.name : 'seu mentor'} em ${meeting_date}`, 'mentorships');
            db.get(`SELECT id FROM users WHERE mentor_id = ?`, [mentor_id], (e, user) => {
                if (!e && user) notificar(user.id, 'Nova mentoria agendada', `Tópicos: ${topics}`, 'mentorships');
            });
            const attendees = [
                { name: executivo ? executivo.name : '', email: executivo ? executivo.email : '' },
                { name: mentor ? mentor.name : '', email: mentor ? mentor.email : '' },
                ...listaParticipantes
            ];
            enviarConvitesMentoria({ ics_uid: uid, ics_sequence: 0, topics, meeting_date, duration_minutes: duracao, status: status || 'Agendada' }, attendees);
            res.json({ message: 'Mentoria agendada com sucesso! Convites de calendário enviados por e-mail.' });
        });
});
app.put('/api/mentorships/:id', requireRole('admin', 'client_admin', 'autonomous'), ensureRecordAccess('mentorships'), async (req, res) => {
    const { mentor_id, meeting_date, topics, minutes, status, duration_minutes, participants } = req.body;
    if (!mentor_id) return res.status(400).json({ error: 'Selecione o mentor.' });
    if (!meeting_date) return res.status(400).json({ error: 'Informe a data da reunião.' });
    const mentor = await dbGet(`SELECT name, email FROM mentors WHERE id = ?`, [mentor_id]);
    const atual = await dbGet(`SELECT * FROM mentorships WHERE id = ?`, [req.params.id]);
    const executivo = atual ? await dbGet(`SELECT name, email FROM employees WHERE id = ?`, [atual.employee_id]) : null;
    const listaParticipantes = Array.isArray(participants) ? participants.filter(p => p && p.email) : [];
    const duracao = Number(duration_minutes) || 60;
    const novoSequence = (atual && atual.ics_sequence ? atual.ics_sequence : 0) + 1;
    const uid = (atual && atual.ics_uid) || `mentoria-${req.params.id}-${Date.now()}@impulsionar`;
    db.run(
        `UPDATE mentorships SET mentor_id = ?, mentor_name = ?, meeting_date = ?, topics = ?, minutes = ?, status = ?, duration_minutes = ?, participants_json = ?, ics_uid = ?, ics_sequence = ? WHERE id = ?`,
        [mentor_id, mentor ? mentor.name : '', meeting_date, topics, minutes || '', status, duracao, JSON.stringify(listaParticipantes), uid, novoSequence, req.params.id],
        (err) => {
            if (err) return res.status(400).json({ error: err.message });
            if (status === 'Realizada') darPontos(req.user.userId, 20, 'Mentoria realizada');
            if (status === 'Realizada' || status === 'Cancelada') {
                notificarPorEmployeeId(atual ? atual.employee_id : null, `Mentoria ${status.toLowerCase()}`, `Atualização da sua mentoria com ${mentor ? mentor.name : 'o mentor'}.`, 'mentorships');
            }
            const attendees = [
                { name: executivo ? executivo.name : '', email: executivo ? executivo.email : '' },
                { name: mentor ? mentor.name : '', email: mentor ? mentor.email : '' },
                ...listaParticipantes
            ];
            enviarConvitesMentoria({ ics_uid: uid, ics_sequence: novoSequence, topics, meeting_date, duration_minutes: duracao, status }, attendees);
            res.json({ message: 'Mentoria atualizada! Convites de calendário reenviados (atualiza o evento na agenda de quem já tinha aceitado).' });
        }
    );
});
app.delete('/api/mentorships/:id', requireRole('admin', 'client_admin', 'autonomous'), ensureRecordAccess('mentorships'), (req, res) => { db.run(`DELETE FROM mentorships WHERE id = ?`, [req.params.id], () => res.json({ message: 'Removido!' })); });

// Gera (ou reaproveita) o link público da sala de videochamada desta mentoria,
// para convidar qualquer pessoa (sem precisar de login) — até 1000 pessoas na
// mesma sala. Só quem participa da mentoria pode gerar/ver o link.
app.post('/api/mentorships/:id/room-link', requireRole('admin', 'client_admin', 'mentor', 'autonomous'), async (req, res) => {
    try {
        const mentoria = await dbGet(`SELECT * FROM mentorships WHERE id = ?`, [req.params.id]);
        if (!mentoria) return res.status(404).json({ error: 'Mentoria não encontrada.' });
        const permitido =
            req.user.role === 'admin' ||
            (req.user.role === 'autonomous' && req.user.employeeId === mentoria.employee_id) ||
            (req.user.role === 'mentor' && req.user.mentorId === mentoria.mentor_id) ||
            (req.user.role === 'client_admin' && await dbGet(`SELECT id FROM employees WHERE id = ? AND company_id = ?`, [mentoria.employee_id, req.user.companyId]));
        if (!permitido) return res.status(403).json({ error: 'Você não faz parte desta mentoria.' });

        let token = mentoria.room_token;
        if (!token) {
            token = crypto.randomBytes(12).toString('hex');
            await new Promise((resolve, reject) => db.run(`UPDATE mentorships SET room_token = ? WHERE id = ?`, [token, req.params.id], (err) => err ? reject(err) : resolve()));
        }
        const baseUrl = appBaseUrlAtiva || process.env.APP_URL || `http://localhost:${PORT}`;
        res.json({ token, link: `${baseUrl.replace(/\/$/, '')}/#sala=${token}` });
    } catch (e) { res.status(500).json({ error: 'Erro ao gerar o link da sala.' }); }
});

// Consulta pública (sem login) usada pela tela de entrada do convidado, só
// para mostrar de qual reunião se trata antes de pedir o nome dele.
app.get('/api/public/sala/:token', async (req, res) => {
    const mentoria = await dbGet(`SELECT id, mentor_name, meeting_date FROM mentorships WHERE room_token = ?`, [req.params.token]);
    if (!mentoria) return res.status(404).json({ error: 'Link inválido ou expirado.' });
    res.json({ mentorName: mentoria.mentor_name || '', meeting_date: mentoria.meeting_date });
});

// ============================================================
// MÓDULO DE CONSULTORIA: planos de acompanhamento com marcos,
// separados do PDI do executivo, + calendário agregado de eventos.
// ============================================================

// Garante que o plano informado por :id está dentro do escopo do usuário
function ensurePlanAccess(getPlanId) {
    return async (req, res, next) => {
        if (req.user.role === 'admin') return next();
        try {
            const planId = getPlanId(req);
            const plano = await dbGet(`SELECT company_id, employee_id FROM consulting_plans WHERE id = ?`, [planId]);
            if (!plano) return res.status(404).json({ error: 'Plano não encontrado.' });
            if (req.user.role === 'client_admin') {
                if (String(plano.company_id) !== String(req.user.companyId)) {
                    return res.status(403).json({ error: 'Este plano não pertence à sua corporação.' });
                }
                return next();
            }
            if (req.user.role === 'autonomous') {
                if (String(plano.employee_id) !== String(req.user.employeeId)) {
                    return res.status(403).json({ error: 'Acesso restrito aos seus próprios planos.' });
                }
                return next();
            }
            return res.status(403).json({ error: 'Perfil sem permissão.' });
        } catch (e) {
            return res.status(500).json({ error: 'Erro ao validar permissão de acesso.' });
        }
    };
}

app.get('/api/consulting-plans', async (req, res) => {
    const scope = buildScope(req);
    if (scope.deny) return res.status(403).json({ error: 'Perfil sem permissão.' });

    let q = `SELECT cp.*, COALESCE(c.name, 'Corporação não vinculada') as companyName, e.name as execName
              FROM consulting_plans cp
              LEFT JOIN companies c ON cp.company_id = c.id
              LEFT JOIN employees e ON cp.employee_id = e.id`;
    const conditions = [];
    const params = [];
    if (scope.companyId) { conditions.push('cp.company_id = ?'); params.push(scope.companyId); }
    if (scope.employeeId) { conditions.push('cp.employee_id = ?'); params.push(scope.employeeId); }
    if (conditions.length) q += ' WHERE ' + conditions.join(' AND ');
    q += ' ORDER BY cp.start_date DESC';

    try {
        const planos = await dbAll(q, params);
        const ids = planos.map(p => p.id);
        let marcos = [];
        if (ids.length) {
            marcos = await dbAll(
                `SELECT * FROM consulting_milestones WHERE plan_id IN (${ids.map(() => '?').join(',')}) ORDER BY due_date ASC`,
                ids
            );
        }
        const porPlano = {};
        marcos.forEach(m => { (porPlano[m.plan_id] = porPlano[m.plan_id] || []).push(m); });
        const resultado = planos.map(p => {
            const seus = porPlano[p.id] || [];
            const concluidos = seus.filter(m => m.status === 'Concluído').length;
            return {
                ...p,
                milestones: seus,
                progresso: seus.length > 0 ? Math.round((concluidos / seus.length) * 100) : 0
            };
        });
        res.json(resultado);
    } catch (e) {
        res.status(500).json({ error: 'Erro ao carregar planos de acompanhamento.' });
    }
});

app.post('/api/consulting-plans', requireRole('admin', 'client_admin'), async (req, res) => {
    let { company_id, employee_id, title, objective, consultant_name, start_date, end_date, status } = req.body;
    if (!title) return res.status(400).json({ error: 'Título do plano é obrigatório.' });

    // A empresa (client_admin) não cria mais o plano diretamente — ela
    // SOLICITA, e o plano só entra em andamento depois que o Master aprovar.
    if (req.user.role === 'client_admin') {
        company_id = req.user.companyId;
        db.run(
            `INSERT INTO consulting_plans (company_id, employee_id, title, objective, consultant_name, start_date, end_date, status, requested_by_company)
             VALUES (?, ?, ?, ?, ?, ?, ?, 'Solicitado', 1)`,
            [company_id || null, employee_id || null, title, objective || '', consultant_name || '', start_date || null, end_date || null],
            function (err) {
                if (err) return res.status(400).json({ error: err.message });
                db.all(`SELECT id FROM users WHERE role = 'admin'`, [], (e, admins) => {
                    if (!e) admins.forEach(a => notificar(a.id, 'Novo plano de acompanhamento solicitado', `"${title}" — revise e aprove ou recuse na Consultoria.`, 'consulting'));
                });
                res.json({ message: 'Solicitação enviada! O plano entra em andamento assim que o Master aprovar.', id: this.lastID, status: 'Solicitado' });
            }
        );
        return;
    }

    // Master criando diretamente já entra em andamento (é quem aprovaria mesmo).
    db.run(
        `INSERT INTO consulting_plans (company_id, employee_id, title, objective, consultant_name, start_date, end_date, status)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
        [company_id || null, employee_id || null, title, objective || '', consultant_name || '', start_date || null, end_date || null, status || 'Em Andamento'],
        function (err) {
            if (err) return res.status(400).json({ error: err.message });
            res.json({ message: 'Plano de acompanhamento criado!', id: this.lastID });
        }
    );
});

// Master aprova a solicitação da empresa: o plano sai de 'Solicitado' e vira 'Em Andamento'.
app.post('/api/consulting-plans/:id/approve', requireRole('admin'), async (req, res) => {
    try {
        const plano = await dbGet(`SELECT * FROM consulting_plans WHERE id = ?`, [req.params.id]);
        if (!plano) return res.status(404).json({ error: 'Plano não encontrado.' });
        if (plano.status !== 'Solicitado') return res.status(400).json({ error: 'Este plano já foi analisado.' });
        db.run(`UPDATE consulting_plans SET status = 'Em Andamento', rejection_reason = NULL WHERE id = ?`, [req.params.id], (err) => {
            if (err) return res.status(400).json({ error: err.message });
            notificarPorCompanyAdmins(plano.company_id, 'Plano de acompanhamento aprovado!', `"${plano.title}" foi aprovado pelo Master e já está em andamento.`, 'consulting');
            res.json({ message: 'Plano aprovado e em andamento!' });
        });
    } catch (e) { res.status(500).json({ error: 'Erro ao aprovar o plano.' }); }
});

// Master recusa a solicitação, com um motivo obrigatório para a empresa entender.
app.post('/api/consulting-plans/:id/reject', requireRole('admin'), async (req, res) => {
    const motivo = (req.body.reason || '').trim();
    if (!motivo) return res.status(400).json({ error: 'Informe o motivo da recusa.' });
    try {
        const plano = await dbGet(`SELECT * FROM consulting_plans WHERE id = ?`, [req.params.id]);
        if (!plano) return res.status(404).json({ error: 'Plano não encontrado.' });
        db.run(`UPDATE consulting_plans SET status = 'Recusado', rejection_reason = ? WHERE id = ?`, [motivo, req.params.id], (err) => {
            if (err) return res.status(400).json({ error: err.message });
            notificarPorCompanyAdmins(plano.company_id, 'Plano de acompanhamento recusado', `"${plano.title}": ${motivo}`, 'consulting');
            res.json({ message: 'Plano recusado — a empresa foi avisada com o motivo.' });
        });
    } catch (e) { res.status(500).json({ error: 'Erro ao recusar o plano.' }); }
});

app.put('/api/consulting-plans/:id', requireRole('admin', 'client_admin'), ensurePlanAccess(req => req.params.id), async (req, res) => {
    const { title, objective, consultant_name, start_date, end_date, status } = req.body;
    try {
        if (req.user.role === 'client_admin') {
            // A empresa edita o conteúdo, mas não decide o próprio status — quem
            // aprova/recusa é sempre o Master. Se o plano tinha sido recusado, a
            // edição reenvia automaticamente para nova análise.
            const atual = await dbGet(`SELECT status FROM consulting_plans WHERE id = ?`, [req.params.id]);
            const reenviar = atual && atual.status === 'Recusado';
            const novoStatus = reenviar ? 'Solicitado' : (atual ? atual.status : 'Solicitado');
            db.run(
                `UPDATE consulting_plans SET title = ?, objective = ?, consultant_name = ?, start_date = ?, end_date = ?, status = ?, rejection_reason = CASE WHEN ? THEN NULL ELSE rejection_reason END WHERE id = ?`,
                [title, objective, consultant_name, start_date, end_date, novoStatus, reenviar ? 1 : 0, req.params.id],
                (err) => {
                    if (err) return res.status(400).json({ error: err.message });
                    if (reenviar) {
                        db.all(`SELECT id FROM users WHERE role = 'admin'`, [], (e, admins) => {
                            if (!e) admins.forEach(a => notificar(a.id, 'Plano de acompanhamento reenviado', `"${title}" foi editado e reenviado — revise novamente.`, 'consulting'));
                        });
                    }
                    res.json({ message: reenviar ? 'Plano atualizado e reenviado para aprovação do Master!' : 'Plano atualizado!' });
                }
            );
            return;
        }
        db.run(
            `UPDATE consulting_plans SET title = ?, objective = ?, consultant_name = ?, start_date = ?, end_date = ?, status = ? WHERE id = ?`,
            [title, objective, consultant_name, start_date, end_date, status, req.params.id],
            (err) => {
                if (err) return res.status(400).json({ error: err.message });
                res.json({ message: 'Plano atualizado!' });
            }
        );
    } catch (e) { res.status(500).json({ error: 'Erro ao atualizar o plano.' }); }
});

app.delete('/api/consulting-plans/:id', requireRole('admin', 'client_admin'), ensurePlanAccess(req => req.params.id), (req, res) => {
    db.run(`DELETE FROM consulting_milestones WHERE plan_id = ?`, [req.params.id], () => {
        db.run(`DELETE FROM consulting_plans WHERE id = ?`, [req.params.id], () => res.json({ message: 'Plano removido!' }));
    });
});

app.post('/api/consulting-plans/:id/milestones', requireRole('admin', 'client_admin'), ensurePlanAccess(req => req.params.id), (req, res) => {
    const { title, due_date } = req.body;
    if (!title) return res.status(400).json({ error: 'Título do marco é obrigatório.' });
    db.run(
        `INSERT INTO consulting_milestones (plan_id, title, due_date, status) VALUES (?, ?, ?, 'Pendente')`,
        [req.params.id, title, due_date || null],
        function (err) {
            if (err) return res.status(400).json({ error: err.message });
            res.json({ message: 'Marco adicionado!', id: this.lastID });
        }
    );
});

app.put('/api/milestones/:id', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const marco = await dbGet(`SELECT plan_id FROM consulting_milestones WHERE id = ?`, [req.params.id]);
        if (!marco) return res.status(404).json({ error: 'Marco não encontrado.' });
        if (req.user.role === 'client_admin') {
            const plano = await dbGet(`SELECT company_id FROM consulting_plans WHERE id = ?`, [marco.plan_id]);
            if (!plano || String(plano.company_id) !== String(req.user.companyId)) {
                return res.status(403).json({ error: 'Este marco não pertence à sua corporação.' });
            }
        }
        const { status, title, due_date } = req.body;
        db.run(
            `UPDATE consulting_milestones SET status = COALESCE(?, status), title = COALESCE(?, title), due_date = COALESCE(?, due_date) WHERE id = ?`,
            [status || null, title || null, due_date || null, req.params.id],
            (err) => {
                if (err) return res.status(400).json({ error: err.message });
                res.json({ message: 'Marco atualizado!' });
            }
        );
    } catch (e) {
        res.status(500).json({ error: 'Erro ao atualizar marco.' });
    }
});

app.delete('/api/milestones/:id', requireRole('admin', 'client_admin'), (req, res) => {
    db.run(`DELETE FROM consulting_milestones WHERE id = ?`, [req.params.id], () => res.json({ message: 'Marco removido!' }));
});

// Calendário agregado: mentorias, prazos de PDI e marcos de acompanhamento,
// todos já filtrados pelo escopo do usuário (empresa ou executivo).
app.get('/api/calendar', async (req, res) => {
    const scope = buildScope(req);
    if (scope.deny) return res.status(403).json({ error: 'Perfil sem permissão.' });

    try {
        const eventos = [];

        const condMent = [];
        const paramsMent = [];
        if (scope.companyId) { condMent.push('e.company_id = ?'); paramsMent.push(scope.companyId); }
        if (scope.employeeId) { condMent.push('m.employee_id = ?'); paramsMent.push(scope.employeeId); }
        const whereMent = condMent.length ? 'WHERE ' + condMent.join(' AND ') : '';
        const mentorias = await dbAll(
            `SELECT m.id, m.meeting_date as data, m.topics, m.status, e.name as execName, ment.name as mentorName
             FROM mentorships m JOIN employees e ON m.employee_id = e.id LEFT JOIN mentors ment ON m.mentor_id = ment.id
             ${whereMent}`,
            paramsMent
        );
        mentorias.forEach(m => eventos.push({
            tipo: 'mentoria', id: m.id, data: m.data, status: m.status,
            titulo: `Mentoria: ${m.execName}${m.mentorName ? ' com ' + m.mentorName : ''}`
        }));

        const condPdi = [];
        const paramsPdi = [];
        if (scope.companyId) { condPdi.push('e.company_id = ?'); paramsPdi.push(scope.companyId); }
        if (scope.employeeId) { condPdi.push('p.employee_id = ?'); paramsPdi.push(scope.employeeId); }
        condPdi.push(`p.deadline IS NOT NULL AND p.deadline != ''`);
        const wherePdi = 'WHERE ' + condPdi.join(' AND ');
        const pdis = await dbAll(
            `SELECT p.id, p.deadline as data, p.objective, p.status, e.name as execName
             FROM pd_plans p JOIN employees e ON p.employee_id = e.id ${wherePdi}`,
            paramsPdi
        );
        pdis.forEach(p => eventos.push({
            tipo: 'pdi', id: p.id, data: p.data, status: p.status,
            titulo: `Prazo de PDI: ${p.execName} — ${p.objective}`
        }));

        const condPlano = [];
        const paramsPlano = [];
        if (scope.companyId) { condPlano.push('cp.company_id = ?'); paramsPlano.push(scope.companyId); }
        if (scope.employeeId) { condPlano.push('cp.employee_id = ?'); paramsPlano.push(scope.employeeId); }
        const wherePlano = condPlano.length ? 'WHERE ' + condPlano.join(' AND ') : '';
        const marcos = await dbAll(
            `SELECT cm.id, cm.due_date as data, cm.title, cm.status, cp.title as planTitle
             FROM consulting_milestones cm JOIN consulting_plans cp ON cm.plan_id = cp.id
             ${wherePlano}`,
            paramsPlano
        );
        marcos.forEach(m => eventos.push({
            tipo: 'marco', id: m.id, data: m.data, status: m.status,
            titulo: `Marco (${m.planTitle}): ${m.title}`
        }));

        // Eventos e encontros da plataforma (masterclasses, presenciais) entram
        // no calendário de todos, já que não são exclusivos de uma empresa.
        const eventosPlataforma = await dbAll(`SELECT id, title, event_date FROM events`, []);
        eventosPlataforma.forEach(ev => eventos.push({
            tipo: 'evento', id: ev.id, data: ev.event_date, status: 'Agendado',
            titulo: `Evento: ${ev.title}`
        }));

        res.json(eventos.filter(e => e.data));
    } catch (e) {
        res.status(500).json({ error: 'Erro ao carregar calendário.' });
    }
});

app.get('/api/corporate-report/:companyId', requireRole('admin', 'client_admin'), (req, res) => {
    const companyId = req.params.companyId;
    if (req.user.role === 'client_admin' && String(companyId) !== String(req.user.companyId)) {
        return res.status(403).json({ error: 'Você só pode ver o relatório da sua própria corporação.' });
    }
    db.get(`SELECT * FROM companies WHERE id = ?`, [companyId], (err, company) => {
        if (!company) return res.status(404).json({ error: 'Não encontrada.' });
        db.all(`SELECT * FROM employees WHERE company_id = ?`, [companyId], (err, employees) => {
            db.all(`SELECT a.*, e.name as execName, e.photo_url as execPhoto FROM assessments a JOIN employees e ON a.employee_id = e.id WHERE e.company_id = ?`, [companyId], (err, assessments) => {
                res.json({
                    company, employees: employees || [], assessments: assessments || [],
                    summary: {
                        totalTalents: employees ? employees.length : 0,
                        totalEvaluations: assessments ? assessments.length : 0,
                        averageScore: assessments && assessments.length > 0 ? (assessments.reduce((acc, curr) => acc + curr.score, 0) / assessments.length).toFixed(1) : 0,
                        averageProgress: employees && employees.length > 0 ? Math.round(employees.reduce((acc, curr) => acc + curr.progress_percentage, 0) / employees.length) : 0
                    }
                });
            });
        });
    });
});

// ============================================================
// TRILHAS DE CURSO (Academy organizada em trilhas sequenciais,
// com progresso por usuário — modelo "estilo G4 Academy")
// ============================================================

app.get('/api/tracks', async (req, res) => {
    try {
        const trilhas = await dbAll(`SELECT * FROM learning_tracks ORDER BY id ASC`);
        const resultado = await Promise.all(trilhas.map(async (t) => {
            const totalAulas = await dbGet(`SELECT COUNT(*) as total FROM track_lessons WHERE track_id = ?`, [t.id]);
            const concluidas = await dbGet(`SELECT COUNT(*) as total FROM track_progress WHERE track_id = ? AND user_id = ?`, [t.id, req.user.userId]);
            const total = totalAulas.total || 0;
            const feitas = concluidas.total || 0;
            return { ...t, totalAulas: total, aulasConcluidas: feitas, progresso: total > 0 ? Math.round((feitas / total) * 100) : 0 };
        }));
        res.json(resultado);
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar trilhas.' }); }
});

app.get('/api/tracks/:id', async (req, res) => {
    try {
        const trilha = await dbGet(`SELECT * FROM learning_tracks WHERE id = ?`, [req.params.id]);
        if (!trilha) return res.status(404).json({ error: 'Trilha não encontrada.' });
        const aulas = await dbAll(
            `SELECT tl.id as track_lesson_id, tl.order_index, v.*,
                    (SELECT COUNT(*) FROM track_progress tp WHERE tp.track_id = tl.track_id AND tp.video_id = v.id AND tp.user_id = ?) as concluida
             FROM track_lessons tl JOIN video_lessons v ON tl.video_id = v.id
             WHERE tl.track_id = ? ORDER BY tl.order_index ASC`,
            [req.user.userId, req.params.id]
        );
        res.json({ ...trilha, aulas: aulas.map(a => ({ ...a, concluida: !!a.concluida })) });
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar trilha.' }); }
});

app.post('/api/tracks', requireRole('admin', 'client_admin'), (req, res) => {
    const { title, description, cover_image_url } = req.body;
    if (!title) return res.status(400).json({ error: 'Informe o título da trilha.' });
    db.run(`INSERT INTO learning_tracks (title, description, cover_image_url) VALUES (?, ?, ?)`,
        [title, description || '', cover_image_url || ''], function (err) {
            if (err) return res.status(400).json({ error: err.message });
            res.json({ message: 'Trilha criada!', id: this.lastID });
        });
});

app.put('/api/tracks/:id', requireRole('admin', 'client_admin'), (req, res) => {
    const { title, description, cover_image_url } = req.body;
    db.run(`UPDATE learning_tracks SET title = ?, description = ?, cover_image_url = ? WHERE id = ?`,
        [title, description || '', cover_image_url || '', req.params.id], (err) => {
            if (err) return res.status(400).json({ error: err.message });
            res.json({ message: 'Trilha atualizada!' });
        });
});

app.delete('/api/tracks/:id', requireRole('admin', 'client_admin'), (req, res) => {
    db.run(`DELETE FROM track_progress WHERE track_id = ?`, [req.params.id], () => {
        db.run(`DELETE FROM certificates WHERE track_id = ?`, [req.params.id], () => {
            db.run(`DELETE FROM track_lessons WHERE track_id = ?`, [req.params.id], () => {
                db.run(`DELETE FROM learning_tracks WHERE id = ?`, [req.params.id], () => res.json({ message: 'Removida!' }));
            });
        });
    });
});

app.post('/api/tracks/:id/lessons', requireRole('admin', 'client_admin'), (req, res) => {
    const { video_id, order_index } = req.body;
    if (!video_id) return res.status(400).json({ error: 'Selecione uma videoaula.' });
    db.run(`INSERT INTO track_lessons (track_id, video_id, order_index) VALUES (?, ?, ?)`,
        [req.params.id, video_id, order_index || 0], (err) => {
            if (err) return res.status(400).json({ error: 'Essa videoaula já está nesta trilha.' });
            res.json({ message: 'Aula adicionada à trilha!' });
        });
});

app.delete('/api/tracks/:id/lessons/:videoId', requireRole('admin', 'client_admin'), (req, res) => {
    db.run(`DELETE FROM track_lessons WHERE track_id = ? AND video_id = ?`, [req.params.id, req.params.videoId], () => res.json({ message: 'Aula removida da trilha!' }));
});

app.post('/api/tracks/:id/videos/:videoId/complete', async (req, res) => {
    try {
        await new Promise((resolve, reject) => db.run(
            `INSERT OR IGNORE INTO track_progress (user_id, track_id, video_id) VALUES (?, ?, ?)`,
            [req.user.userId, req.params.id, req.params.videoId],
            (err) => err ? reject(err) : resolve()
        ));
        darPontos(req.user.userId, 10, 'Aula concluída na trilha');
        res.json({ message: 'Aula marcada como concluída!' });
    } catch (e) { res.status(400).json({ error: 'Erro ao marcar aula como concluída.' }); }
});

// ============================================================
// CERTIFICADOS DE CONCLUSÃO
// ============================================================

app.get('/api/tracks/:id/certificate', async (req, res) => {
    try {
        const trilha = await dbGet(`SELECT * FROM learning_tracks WHERE id = ?`, [req.params.id]);
        if (!trilha) return res.status(404).json({ error: 'Trilha não encontrada.' });
        const totalAulas = await dbGet(`SELECT COUNT(*) as total FROM track_lessons WHERE track_id = ?`, [req.params.id]);
        if (!totalAulas.total) return res.status(400).json({ error: 'Esta trilha ainda não tem aulas cadastradas.' });
        const concluidas = await dbGet(`SELECT COUNT(*) as total FROM track_progress WHERE track_id = ? AND user_id = ?`, [req.params.id, req.user.userId]);
        if (concluidas.total < totalAulas.total) {
            return res.status(400).json({ error: `Conclua todas as aulas da trilha para emitir o certificado (${concluidas.total}/${totalAulas.total}).` });
        }
        let certificado = await dbGet(`SELECT * FROM certificates WHERE user_id = ? AND track_id = ?`, [req.user.userId, req.params.id]);
        if (!certificado) {
            const codigo = 'IMP4-' + crypto.randomBytes(6).toString('hex').toUpperCase();
            await new Promise((resolve, reject) => db.run(
                `INSERT INTO certificates (user_id, track_id, certificate_code) VALUES (?, ?, ?)`,
                [req.user.userId, req.params.id, codigo],
                (err) => err ? reject(err) : resolve()
            ));
            darPontos(req.user.userId, 100, 'Certificado de trilha emitido');
            notificar(req.user.userId, 'Certificado emitido!', `Seu certificado da trilha "${trilha.title}" está pronto para download.`, 'tracks');
            certificado = await dbGet(`SELECT * FROM certificates WHERE user_id = ? AND track_id = ?`, [req.user.userId, req.params.id]);
        }
        const usuario = await dbGet(`SELECT name FROM users WHERE id = ?`, [req.user.userId]);
        res.json({ ...certificado, trackTitle: trilha.title, userName: usuario ? usuario.name : '' });
    } catch (e) { res.status(500).json({ error: 'Erro ao emitir certificado.' }); }
});

app.get('/api/tracks/:id/certificate/pdf', async (req, res) => {
    try {
        const trilha = await dbGet(`SELECT * FROM learning_tracks WHERE id = ?`, [req.params.id]);
        if (!trilha) return res.status(404).json({ error: 'Trilha não encontrada.' });
        const certificado = await dbGet(`SELECT * FROM certificates WHERE user_id = ? AND track_id = ?`, [req.user.userId, req.params.id]);
        if (!certificado) return res.status(400).json({ error: 'Certificado ainda não emitido para esta trilha.' });
        const usuario = await dbGet(`SELECT name FROM users WHERE id = ?`, [req.user.userId]);
        const nomeUsuario = usuario ? usuario.name : 'Executivo(a)';
        const dataEmissao = new Date(certificado.issued_at).toLocaleDateString('pt-BR');

        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', `attachment; filename="certificado-${trilha.id}.pdf"`);

        const doc = new PDFDocument({ layout: 'landscape', size: 'A4', margin: 0 });
        doc.pipe(res);

        const largura = doc.page.width;
        const altura = doc.page.height;

        // Moldura externa
        doc.rect(0, 0, largura, altura).fill('#0b1329');
        doc.rect(24, 24, largura - 48, altura - 48).lineWidth(2).stroke('#c9a227');
        doc.rect(34, 34, largura - 68, altura - 68).lineWidth(0.75).stroke('#c9a227');

        doc.fillColor('#c9a227').font('Helvetica-Bold').fontSize(14)
            .text('IMPULSIONAR V4', 0, 70, { align: 'center' });
        doc.fillColor('#ffffff').font('Helvetica').fontSize(11)
            .text('PLATAFORMA DE LIDERANÇA E PERFORMANCE', 0, 92, { align: 'center' });

        doc.fillColor('#ffffff').font('Helvetica-Bold').fontSize(30)
            .text('Certificado de Conclusão', 0, 150, { align: 'center' });

        doc.font('Helvetica').fontSize(13).fillColor('#cbd5e1')
            .text('Certificamos que', 0, 210, { align: 'center' });

        doc.font('Helvetica-Bold').fontSize(24).fillColor('#ffffff')
            .text(nomeUsuario, 0, 235, { align: 'center' });

        doc.font('Helvetica').fontSize(13).fillColor('#cbd5e1')
            .text('concluiu integralmente a trilha de desenvolvimento', 0, 275, { align: 'center' });

        doc.font('Helvetica-Bold').fontSize(18).fillColor('#c9a227')
            .text(trilha.title, 60, 300, { align: 'center', width: largura - 120 });

        doc.font('Helvetica').fontSize(10).fillColor('#94a3b8')
            .text(`Emitido em ${dataEmissao}  ·  Código de validação: ${certificado.certificate_code}`, 0, altura - 90, { align: 'center' });

        doc.end();
    } catch (e) {
        res.status(500).json({ error: 'Erro ao gerar o PDF do certificado.' });
    }
});

app.get('/api/certificates', async (req, res) => {
    try {
        const meus = await dbAll(
            `SELECT c.*, t.title as trackTitle FROM certificates c JOIN learning_tracks t ON c.track_id = t.id WHERE c.user_id = ? ORDER BY c.issued_at DESC`,
            [req.user.userId]
        );
        res.json(meus);
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar certificados.' }); }
});

// ============================================================
// COMUNIDADE / REDE DE CONTATOS ENTRE MEMBROS
// ============================================================

app.get('/api/community', async (req, res) => {
    try {
        const busca = req.query.q ? `%${req.query.q}%` : null;
        let q = `SELECT e.id, e.name, e.role, e.photo_url, e.public_bio, e.linkedin_url, e.disc_profile, c.name as companyName
                  FROM employees e LEFT JOIN companies c ON e.company_id = c.id
                  WHERE e.show_in_directory = 1`;
        const params = [];
        if (busca) { q += ` AND (e.name LIKE ? OR e.role LIKE ? OR c.name LIKE ?)`; params.push(busca, busca, busca); }
        q += ` ORDER BY e.name ASC`;
        const membros = await dbAll(q, params);
        res.json(membros);
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar comunidade.' }); }
});

// ============================================================
// BANCO DE CURRÍCULOS (executivos buscando oportunidades)
// ============================================================

app.get('/api/talent-pool', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const { q: termo, desiredRole, discProfile, minProgress, hasResume } = req.query;
        let q = `SELECT e.*, c.name as companyName FROM employees e LEFT JOIN companies c ON e.company_id = c.id WHERE e.looking_for_opportunity = 1`;
        const params = [];
        if (req.user.role === 'client_admin') { q += ` AND e.company_id = ?`; params.push(req.user.companyId); }
        if (termo) { q += ` AND (e.name LIKE ? OR e.role LIKE ? OR e.desired_role LIKE ?)`; params.push(`%${termo}%`, `%${termo}%`, `%${termo}%`); }
        if (desiredRole) { q += ` AND e.desired_role LIKE ?`; params.push(`%${desiredRole}%`); }
        if (discProfile) { q += ` AND e.disc_profile = ?`; params.push(discProfile); }
        if (minProgress) { q += ` AND e.progress_percentage >= ?`; params.push(Number(minProgress)); }
        if (hasResume === '1') { q += ` AND e.resume_url IS NOT NULL AND e.resume_url != ''`; }
        q += ` ORDER BY e.name ASC`;
        const candidatos = await dbAll(q, params);
        res.json(candidatos);
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar banco de currículos.' }); }
});

// ============================================================
// EVENTOS E ENCONTROS
// ============================================================

app.get('/api/events', async (req, res) => {
    try {
        const eventos = await dbAll(
            `SELECT ev.*,
                    (SELECT COUNT(*) FROM event_registrations er WHERE er.event_id = ev.id) as totalInscritos,
                    (SELECT COUNT(*) FROM event_registrations er2 WHERE er2.event_id = ev.id AND er2.user_id = ?) as inscritoEu,
                    (SELECT COUNT(*) FROM event_waitlist ew WHERE ew.event_id = ev.id) as totalNaEspera,
                    (SELECT COUNT(*) FROM event_waitlist ew2 WHERE ew2.event_id = ev.id AND ew2.user_id = ?) as emEsperaEu
             FROM events ev ORDER BY ev.event_date ASC`,
            [req.user.userId, req.user.userId]
        );
        res.json(eventos.map(e => ({ ...e, inscritoEu: !!e.inscritoEu, emEsperaEu: !!e.emEsperaEu })));
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar eventos.' }); }
});

app.post('/api/events', requireRole('admin', 'client_admin'), (req, res) => {
    const { title, description, event_date, event_time, is_online, location, link, capacity } = req.body;
    if (!title || !event_date) return res.status(400).json({ error: 'Informe o título e a data do evento.' });
    db.run(`INSERT INTO events (title, description, event_date, event_time, is_online, location, link, capacity) VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
        [title, description || '', event_date, event_time || '', is_online ? 1 : 0, location || '', link || '', capacity || null], function (err) {
            if (err) return res.status(400).json({ error: err.message });
            res.json({ message: 'Evento criado!', id: this.lastID });
        });
});

app.put('/api/events/:id', requireRole('admin', 'client_admin'), (req, res) => {
    const { title, description, event_date, event_time, is_online, location, link, capacity } = req.body;
    db.run(`UPDATE events SET title = ?, description = ?, event_date = ?, event_time = ?, is_online = ?, location = ?, link = ?, capacity = ? WHERE id = ?`,
        [title, description || '', event_date, event_time || '', is_online ? 1 : 0, location || '', link || '', capacity || null, req.params.id], (err) => {
            if (err) return res.status(400).json({ error: err.message });
            res.json({ message: 'Evento atualizado!' });
        });
});

app.delete('/api/events/:id', requireRole('admin', 'client_admin'), (req, res) => {
    db.run(`DELETE FROM event_registrations WHERE event_id = ?`, [req.params.id], () => {
        db.run(`DELETE FROM event_waitlist WHERE event_id = ?`, [req.params.id], () => {
            db.run(`DELETE FROM events WHERE id = ?`, [req.params.id], () => res.json({ message: 'Removido!' }));
        });
    });
});

app.post('/api/events/:id/register', async (req, res) => {
    try {
        const jaInscrito = await dbGet(`SELECT id FROM event_registrations WHERE event_id = ? AND user_id = ?`, [req.params.id, req.user.userId]);
        if (jaInscrito) {
            await new Promise((resolve, reject) => db.run(`DELETE FROM event_registrations WHERE id = ?`, [jaInscrito.id], (err) => err ? reject(err) : resolve()));
            // Uma vaga acabou de se abrir — avisa por e-mail quem está na lista de espera.
            notificarProximoDaListaDeEspera(req.params.id);
            return res.json({ inscrito: false, emEspera: false, message: 'Inscrição cancelada.' });
        }

        const jaNaEspera = await dbGet(`SELECT id FROM event_waitlist WHERE event_id = ? AND user_id = ?`, [req.params.id, req.user.userId]);
        if (jaNaEspera) {
            await new Promise((resolve, reject) => db.run(`DELETE FROM event_waitlist WHERE id = ?`, [jaNaEspera.id], (err) => err ? reject(err) : resolve()));
            return res.json({ inscrito: false, emEspera: false, message: 'Você saiu da lista de espera.' });
        }

        const evento = await dbGet(`SELECT * FROM events WHERE id = ?`, [req.params.id]);
        if (evento && evento.capacity) {
            const contagem = await dbGet(`SELECT COUNT(*) as total FROM event_registrations WHERE event_id = ?`, [req.params.id]);
            if (contagem.total >= evento.capacity) {
                await new Promise((resolve, reject) => db.run(`INSERT INTO event_waitlist (event_id, user_id) VALUES (?, ?)`, [req.params.id, req.user.userId], (err) => err ? reject(err) : resolve()));
                return res.json({ inscrito: false, emEspera: true, message: 'Este evento está lotado. Você entrou na lista de espera e será avisado por e-mail assim que uma vaga abrir.' });
            }
        }
        await new Promise((resolve, reject) => db.run(`INSERT INTO event_registrations (event_id, user_id) VALUES (?, ?)`, [req.params.id, req.user.userId], (err) => err ? reject(err) : resolve()));
        darPontos(req.user.userId, 15, 'Inscrição em evento');
        res.json({ inscrito: true, emEspera: false, message: 'Inscrição confirmada!' });
    } catch (e) { res.status(400).json({ error: 'Erro ao processar inscrição.' }); }
});

app.get('/api/events/:id/attendees', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const lista = await dbAll(
            `SELECT u.name, u.email FROM event_registrations er JOIN users u ON er.user_id = u.id WHERE er.event_id = ?`,
            [req.params.id]
        );
        const espera = await dbAll(
            `SELECT u.name, u.email, ew.created_at FROM event_waitlist ew JOIN users u ON ew.user_id = u.id WHERE ew.event_id = ? ORDER BY ew.created_at ASC`,
            [req.params.id]
        );
        res.json({ inscritos: lista, listaDeEspera: espera });
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar inscritos.' }); }
});

// ============================================================
// MEU PERFIL (dados da própria Impulsionar — usados nos contratos e nas automações)
// ============================================================

app.get('/api/platform/profile', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const perfil = await dbGet(`SELECT * FROM platform_profile WHERE id = 1`);
        res.json(perfil || {});
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar o perfil da Impulsionar.' }); }
});

app.put('/api/platform/profile', requireRole('admin'), (req, res) => {
    const { name, cnpj, phone, email, address, logo_url } = req.body;
    db.run(
        `UPDATE platform_profile SET name = ?, cnpj = ?, phone = ?, email = ?, address = ?, logo_url = ? WHERE id = 1`,
        [name || '', cnpj || '', phone || '', email || '', address || '', logo_url || ''],
        (err) => {
            if (err) return res.status(400).json({ error: err.message });
            res.json({ message: 'Perfil da Impulsionar atualizado!' });
        }
    );
});

// Testa o SMTP configurado no .env sem precisar olhar o console do servidor —
// manda um e-mail de teste de verdade e devolve o erro exato do provedor
// (ex: senha de app inválida) direto na tela, em vez de só um aviso mudo.
app.post('/api/platform/test-email', requireRole('admin'), async (req, res) => {
    const destino = (req.body && req.body.email) || req.user.email;
    if (!destino) return res.status(400).json({ error: 'Informe um e-mail de destino para o teste.' });
    if (!process.env.SMTP_HOST || !process.env.SMTP_USER || !process.env.SMTP_PASS) {
        return res.status(400).json({ error: 'SMTP_HOST, SMTP_USER e SMTP_PASS ainda não estão definidos no .env do servidor.' });
    }
    try {
        const info = await transporter.sendMail({
            from: process.env.SMTP_FROM || `"Impulsionar V4" <${process.env.SMTP_USER}>`,
            to: destino,
            subject: 'Teste de envio — Impulsionar V4',
            html: `<p>Se você recebeu este e-mail, o SMTP configurado no <strong>.env</strong> está funcionando corretamente.</p>`
        });
        res.json({ message: `E-mail de teste enviado para ${destino}! Confira a caixa de entrada (e o spam).`, id: info.messageId });
    } catch (e) {
        res.status(400).json({ error: `Falha ao enviar: ${e.message}` });
    }
});

// ============================================================
// INTEGRAÇÃO MERCADO PAGO — configurável direto na tela "Meu Perfil", sem
// precisar editar o .env nem reiniciar o servidor. O token fica salvo no
// banco (integration_settings) e sobrepõe o MP_ACCESS_TOKEN do .env assim
// que é salvo.
// ============================================================
app.get('/api/integrations/mercadopago', requireRole('admin'), async (req, res) => {
    try {
        const row = await dbGet(`SELECT value FROM integration_settings WHERE key = 'mp_access_token'`);
        const tokenSalvo = row && row.value;
        res.json({
            configured: !!mpPreference,
            source: tokenSalvo ? 'banco' : (MP_ACCESS_TOKEN ? 'env' : 'nenhum'),
            tokenPreview: mpAccessTokenAtivo ? ('••••••••' + mpAccessTokenAtivo.slice(-6)) : null,
            isTestToken: mpAccessTokenAtivo.startsWith('TEST-')
        });
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar a configuração do Mercado Pago.' }); }
});

// Faz uma chamada real e leve na API do Mercado Pago (lista de métodos de
// pagamento) só para confirmar se o token é válido — sem custo e sem
// depender de nenhum produto/venda existir na conta.
//
// Importante: sem um timeout explícito, se o servidor não conseguir alcançar
// api.mercadopago.com (rede bloqueada, DNS, firewall do provedor de
// hospedagem), o fetch pode ficar pendurado por muito tempo. Nesse caso um
// proxy/load balancer na frente do Node (nginx, Render, Railway etc.) acaba
// cortando a requisição sozinho e devolve uma página de erro HTML — que o
// front-end não consegue entender como JSON, e é isso que aparece como
// "Erro na requisição" genérico. Com o AbortController abaixo, é o próprio
// servidor quem responde primeiro, sempre em JSON, com um motivo claro.
async function validarTokenMercadoPago(token) {
    const controlador = new AbortController();
    const timeoutId = setTimeout(() => controlador.abort(), 15000);
    let resposta;
    try {
        resposta = await fetch('https://api.mercadopago.com/v1/payment_methods', {
            headers: { Authorization: `Bearer ${token}` },
            signal: controlador.signal
        });
    } catch (e) {
        if (e.name === 'AbortError') {
            throw new Error('Tempo esgotado ao conectar com a API do Mercado Pago (15s). Verifique se este servidor tem acesso à internet para api.mercadopago.com e tente novamente.');
        }
        throw new Error(`Não foi possível conectar à API do Mercado Pago (${e.message || 'falha de rede'}). Verifique a conexão do servidor com a internet.`);
    } finally {
        clearTimeout(timeoutId);
    }
    if (resposta.status === 401 || resposta.status === 403) {
        const corpo = await resposta.json().catch(() => ({}));
        throw new Error(corpo.message || 'Token inválido ou sem autorização (verifique se copiou o Access Token certo, de Produção ou de Teste, sem espaços extras).');
    }
    if (!resposta.ok) throw new Error(`A API do Mercado Pago respondeu com erro (status ${resposta.status}).`);
    return true;
}

app.put('/api/integrations/mercadopago', requireRole('admin'), async (req, res) => {
    const token = (req.body.accessToken || '').trim();
    if (!token) return res.status(400).json({ error: 'Cole o Access Token do Mercado Pago.' });
    try {
        await validarTokenMercadoPago(token);
    } catch (e) {
        return res.status(400).json({ error: `Não foi possível validar este token: ${e.message}` });
    }
    db.run(
        `INSERT INTO integration_settings (key, value) VALUES ('mp_access_token', ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value`,
        [token],
        (err) => {
            if (err) return res.status(400).json({ error: err.message });
            configurarMercadoPago(token);
            res.json({ message: `Mercado Pago configurado e validado com sucesso! ${token.startsWith('TEST-') ? '(Este é um token de TESTE — os pagamentos gerados serão simulados no ambiente sandbox.)' : '(Token de produção — os pagamentos serão reais.)'}` });
        }
    );
});

app.post('/api/integrations/mercadopago/test', requireRole('admin'), async (req, res) => {
    if (!mpAccessTokenAtivo) return res.status(400).json({ error: 'Nenhum token configurado ainda.' });
    try {
        await validarTokenMercadoPago(mpAccessTokenAtivo);
        res.json({ message: 'Token válido! A conexão com o Mercado Pago está funcionando.' });
    } catch (e) {
        res.status(400).json({ error: e.message });
    }
});

app.delete('/api/integrations/mercadopago', requireRole('admin'), (req, res) => {
    db.run(`DELETE FROM integration_settings WHERE key = 'mp_access_token'`, [], (err) => {
        if (err) return res.status(400).json({ error: err.message });
        configurarMercadoPago(MP_ACCESS_TOKEN); // volta para o .env (se houver) ou desativa
        res.json({ message: 'Token removido do banco. ' + (MP_ACCESS_TOKEN ? 'Voltou a usar o token do .env.' : 'Mercado Pago ficou desativado.') });
    });
});

// ============================================================
// URL PÚBLICA DO SISTEMA — usada como back_urls do Mercado Pago (a página
// para onde o pagador volta depois de pagar). O Mercado Pago recusa gerar o
// link quando essa URL é "http://localhost:..." (não é pública), com o erro
// "auto_return invalid. back_url.success must be defined". Configurável
// aqui, sem precisar editar o .env nem reiniciar o servidor.
// ============================================================
app.get('/api/integrations/app-url', requireRole('admin'), async (req, res) => {
    try {
        const row = await dbGet(`SELECT value FROM integration_settings WHERE key = 'app_base_url'`);
        const salvoNoBanco = row && row.value;
        res.json({
            url: appBaseUrlAtiva,
            source: salvoNoBanco ? 'banco' : (process.env.APP_BASE_URL ? 'env' : 'padrao'),
            isPublica: urlPublicaValida(appBaseUrlAtiva)
        });
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar a URL do sistema.' }); }
});

app.put('/api/integrations/app-url', requireRole('admin'), async (req, res) => {
    const url = (req.body.url || '').trim().replace(/\/+$/, '');
    if (!url) return res.status(400).json({ error: 'Informe a URL pública do sistema.' });
    try {
        new URL(url);
    } catch (e) {
        return res.status(400).json({ error: 'URL inválida. Use o formato completo, ex.: https://seusistema.com.br' });
    }
    db.run(
        `INSERT INTO integration_settings (key, value) VALUES ('app_base_url', ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value`,
        [url],
        (err) => {
            if (err) return res.status(400).json({ error: err.message });
            appBaseUrlAtiva = url;
            const publica = urlPublicaValida(url);
            res.json({
                message: publica
                    ? 'URL do sistema salva! Os próximos links de pagamento já vão usar este endereço para o retorno automático.'
                    : 'URL salva, mas ela ainda parece ser local (localhost) — o Mercado Pago não redireciona automaticamente de volta para endereços assim. O link de pagamento continua funcionando, só sem o retorno automático até você configurar um endereço público.'
            });
        }
    );
});

app.delete('/api/integrations/app-url', requireRole('admin'), (req, res) => {
    db.run(`DELETE FROM integration_settings WHERE key = 'app_base_url'`, [], (err) => {
        if (err) return res.status(400).json({ error: err.message });
        appBaseUrlAtiva = APP_BASE_URL_ENV;
        res.json({ message: 'Removida. Voltou a usar a URL do .env (ou o padrão local).' });
    });
});

// Informa se o banco está num caminho persistente (DB_PATH, ex.: um Volume no
// Railway) ou dentro da pasta do código — nesse segundo caso, todo novo
// deploy substitui o arquivo pela versão salva no repositório, apagando o que
// foi cadastrado depois do último upload do banco no GitHub.
app.get('/api/admin/database-info', requireRole('admin'), (req, res) => {
    res.json({ persistente: !!process.env.DB_PATH, caminho: dbFile });
});

// Baixa uma cópia do arquivo .sqlite atual — útil antes de qualquer mudança
// arriscada (trocar de servidor, configurar um Volume, etc.), já que hoje o
// banco não é versionado/backupeado automaticamente em lugar nenhum.
app.get('/api/admin/database-backup', requireRole('admin'), (req, res) => {
    res.download(dbFile, 'backup-impulsionar.sqlite', (err) => {
        if (err && !res.headersSent) res.status(500).json({ error: 'Erro ao gerar o backup do banco.' });
    });
});

// ============================================================
// PAINEL DE AUTOMAÇÃO (o que roda sozinho x precisa de autorização do Master)
// ============================================================

app.get('/api/automation-settings', requireRole('admin'), async (req, res) => {
    try {
        const linhas = await dbAll(`SELECT * FROM automation_settings`);
        const objeto = {};
        linhas.forEach(l => { objeto[l.key] = l.value === '1'; });
        res.json(objeto);
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar configurações de automação.' }); }
});

app.put('/api/automation-settings', requireRole('admin'), async (req, res) => {
    try {
        const entradas = Object.entries(req.body || {});
        for (const [chave, valor] of entradas) {
            await new Promise((resolve, reject) => db.run(
                `INSERT INTO automation_settings (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value`,
                [chave, valor ? '1' : '0'],
                (err) => err ? reject(err) : resolve()
            ));
        }
        res.json({ message: 'Configurações de automação atualizadas!' });
    } catch (e) { res.status(400).json({ error: e.message }); }
});

// Dados de WhatsApp do próprio Master (número + se quer receber avisos).
app.get('/api/me/whatsapp', requireRole('admin'), async (req, res) => {
    try {
        const u = await dbGet(`SELECT whatsapp_number, whatsapp_notifications FROM users WHERE id = ?`, [req.user.userId]);
        res.json({ whatsapp_number: u?.whatsapp_number || '', whatsapp_notifications: !!u?.whatsapp_notifications });
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar dados de WhatsApp.' }); }
});

app.put('/api/me/whatsapp', requireRole('admin'), (req, res) => {
    const { whatsapp_number, whatsapp_notifications } = req.body;
    db.run(`UPDATE users SET whatsapp_number = ?, whatsapp_notifications = ? WHERE id = ?`,
        [whatsapp_number || '', whatsapp_notifications ? 1 : 0, req.user.userId],
        (err) => {
            if (err) return res.status(400).json({ error: err.message });
            res.json({ message: 'Dados de WhatsApp atualizados!' });
        });
});

// ============================================================
// ASSISTENTE DE IA (Anthropic) — sugestões e aprovação via WhatsApp
// ============================================================

// Lista as ações que a IA propôs (histórico + pendentes) — visível pro Master acompanhar.
app.get('/api/ai/pending-actions', requireRole('admin'), async (req, res) => {
    try {
        const lista = await dbAll(
            `SELECT ap.*, u.name as candidateName FROM ai_pending_actions ap LEFT JOIN users u ON u.id = ap.candidate_user_id ORDER BY ap.created_at DESC LIMIT 100`
        );
        res.json(lista);
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar ações da IA.' }); }
});

// Aprova/rejeita manualmente pela própria tela (equivalente a responder pelo WhatsApp).
app.post('/api/ai/pending-actions/:id/resolve', requireRole('admin'), async (req, res) => {
    const { decisao } = req.body; // 'aprovado' | 'rejeitado'
    if (!['aprovado', 'rejeitado'].includes(decisao)) return res.status(400).json({ error: 'Decisão inválida.' });
    try {
        await resolverAcaoIA(req.params.id, decisao);
        res.json({ message: 'Decisão aplicada!' });
    } catch (e) { res.status(400).json({ error: e.message }); }
});

// Aplica a decisão do Master sobre uma ação pendente da IA — usado tanto pela
// tela quanto pelas respostas recebidas via WhatsApp.
async function resolverAcaoIA(id, decisao) {
    const acao = await dbGet(`SELECT * FROM ai_pending_actions WHERE id = ?`, [id]);
    if (!acao) throw new Error('Ação da IA não encontrada.');
    if (acao.status !== 'pendente') throw new Error('Esta ação já foi resolvida anteriormente.');

    if (decisao === 'aprovado') {
        if (acao.type === 'suporte_reply') {
            await new Promise((resolve, reject) => db.run(
                `INSERT INTO candidate_messages (candidate_user_id, sender, message, topic, read_by_master) VALUES (?, 'master', ?, 'suporte', 1)`,
                [acao.candidate_user_id, acao.proposal],
                (err) => err ? reject(err) : resolve()
            ));
            notificar(acao.candidate_user_id, 'Nova resposta do Master', 'Você recebeu uma resposta no Suporte.', 'portalMensagens');
        } else if (acao.type === 'ajuste_curriculo') {
            await new Promise((resolve, reject) => db.run(
                `UPDATE candidate_profiles SET status = ? WHERE user_id = ?`,
                [acao.suggested_status || 'ajustes_solicitados', acao.candidate_user_id],
                (err) => err ? reject(err) : resolve()
            ));
            await new Promise((resolve, reject) => db.run(
                `INSERT INTO candidate_messages (candidate_user_id, sender, message, topic, read_by_master) VALUES (?, 'master', ?, 'curriculo', 1)`,
                [acao.candidate_user_id, acao.proposal],
                (err) => err ? reject(err) : resolve()
            ));
            notificar(acao.candidate_user_id, 'Atualização sobre seu currículo', 'O Master analisou seu currículo — confira em "Fale com o Master".', 'portalMensagens');
        }
    }
    await new Promise((resolve, reject) => db.run(
        `UPDATE ai_pending_actions SET status = ?, resolved_at = CURRENT_TIMESTAMP WHERE id = ?`,
        [decisao, id],
        (err) => err ? reject(err) : resolve()
    ));
}

// Cria uma ação pendente da IA: se o autopilot daquele tipo estiver ligado,
// aplica direto; senão, manda pro WhatsApp do(s) gestor(es) aprovar.
async function criarAcaoIA({ type, candidateUserId, demand, proposal, suggestedStatus, autopilotKey }) {
    const id = await new Promise((resolve, reject) => db.run(
        `INSERT INTO ai_pending_actions (type, candidate_user_id, demand, proposal, suggested_status) VALUES (?, ?, ?, ?, ?)`,
        [type, candidateUserId || null, demand || '', proposal, suggestedStatus || null],
        function (err) { err ? reject(err) : resolve(this.lastID); }
    ));
    const autopilotLigado = autopilotKey ? await automacaoLigada(autopilotKey) : false;
    if (autopilotLigado) {
        await resolverAcaoIA(id, 'aprovado');
    } else {
        const rotulos = { suporte_reply: 'Resposta de Suporte', ajuste_curriculo: 'Revisão de Currículo' };
        const texto = `🤖 *Impulsionar IA* — ${rotulos[type] || type}\n\n*Demanda:*\n${demand || '-'}\n\n*Sugestão da IA:*\n${proposal}\n\nResponda *APROVAR ${id}* para aplicar, ou *REJEITAR ${id}* para descartar.`;
        await avisarGestoresPorWhatsApp(texto);
    }
    return id;
}

// Recebe as respostas do gestor pelo WhatsApp (webhook do Twilio). Espera um
// texto no formato "APROVAR 12" ou "REJEITAR 12".
app.post('/api/webhooks/whatsapp', async (req, res) => {
    res.set('Content-Type', 'text/xml');
    try {
        const texto = String(req.body.Body || '').trim();
        const combinacao = texto.match(/(aprovar|rejeitar)\s*(\d+)/i);
        if (!combinacao) return res.send('<Response><Message>Não entendi. Responda "APROVAR [número]" ou "REJEITAR [número]".</Message></Response>');
        const decisao = combinacao[1].toLowerCase() === 'aprovar' ? 'aprovado' : 'rejeitado';
        const id = combinacao[2];
        await resolverAcaoIA(id, decisao);
        return res.send(`<Response><Message>✅ Ação #${id} ${decisao === 'aprovado' ? 'aprovada e aplicada' : 'rejeitada'}!</Message></Response>`);
    } catch (e) {
        return res.send(`<Response><Message>⚠️ Não consegui aplicar essa decisão: ${e.message}</Message></Response>`);
    }
});

// IA analisa os números do Painel de Engajamento e devolve um resumo com
// alertas e recomendações para o Master — sem nenhuma ação automática.
app.post('/api/ai/engagement-insights', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const dadosPainel = req.body.dados || {};
        const resposta = await perguntarIA(
            'Você é o assistente de IA da Impulsionar, uma consultoria de desenvolvimento de executivos. ' +
            'Analise os números do painel de engajamento a seguir e escreva um resumo objetivo em português, ' +
            'com no máximo 5 tópicos: o que está indo bem, o que precisa de atenção, e 1-2 recomendações práticas. ' +
            'Seja direto, sem introduções longas.',
            JSON.stringify(dadosPainel, null, 2)
        );
        res.json({ insights: resposta });
    } catch (e) { res.status(400).json({ error: e.message }); }
});

// ============================================================
// CONTRATOS COM ASSINATURA DIGITAL (via Autentique)
// ============================================================

app.get('/api/contracts', requireRole('admin', 'client_admin', 'autonomous', 'mentor'), async (req, res) => {
    try {
        let query = `SELECT c.*, co.name as companyName, e.name as employeeName, m.name as mentorName
                      FROM contracts c
                      LEFT JOIN companies co ON co.id = c.company_id
                      LEFT JOIN employees e ON e.id = c.employee_id
                      LEFT JOIN mentors m ON m.id = c.mentor_id`;
        const condicoes = [];
        const params = [];
        if (req.user.role === 'client_admin') { condicoes.push('c.company_id = ?'); params.push(req.user.companyId); }
        if (req.user.role === 'autonomous') { condicoes.push('c.employee_id = ?'); params.push(req.user.employeeId); }
        if (req.user.role === 'mentor') { condicoes.push('c.mentor_id = ?'); params.push(req.user.mentorId); }
        if (condicoes.length) query += ' WHERE ' + condicoes.join(' AND ');
        query += ' ORDER BY c.created_at DESC';
        const lista = await dbAll(query, params);
        res.json(lista.map(c => ({ ...c, signers: JSON.parse(c.signers_json || '[]') })));
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar contratos.' }); }
});

// Gera o Buffer do PDF a partir do texto do contrato (título + parágrafos) —
// usado quando o contrato é escrito/editado na própria tela (não anexado).
function gerarPdfContratoDeTexto(title, content) {
    return new Promise((resolve, reject) => {
        const doc = new PDFDocument({ margin: 60 });
        const chunks = [];
        doc.on('data', (c) => chunks.push(c));
        doc.on('end', () => resolve(Buffer.concat(chunks)));
        doc.on('error', reject);
        doc.font('Helvetica-Bold').fontSize(16).text(title, { align: 'center' });
        doc.moveDown(1.5);
        doc.font('Helvetica').fontSize(11).text(content, { align: 'justify', lineGap: 4 });
        doc.end();
    });
}

// Lê um arquivo já enviado via /api/upload (um contrato pronto — Word ou PDF —
// anexado pelo usuário só para coletar assinatura, sem gerar nada a partir de texto).
function lerArquivoAnexadoContrato(attachedFileUrl) {
    if (!attachedFileUrl || !attachedFileUrl.startsWith('/uploads/')) throw new Error('Arquivo anexado inválido.');
    const nomeArquivo = path.basename(attachedFileUrl);
    const caminho = path.join(PASTA_UPLOADS, nomeArquivo);
    if (!fs.existsSync(caminho)) throw new Error('Arquivo anexado não encontrado no servidor.');
    return fs.readFileSync(caminho);
}

app.post('/api/contracts', requireRole('admin'), async (req, res) => {
    if (!AUTENTIQUE_API_TOKEN) return res.status(503).json({ error: 'Assinatura digital ainda não foi configurada no servidor (defina AUTENTIQUE_API_TOKEN no .env).' });
    const { type, company_id, employee_id, mentor_id, title, content, signers, attachedFileUrl, attachedFileName } = req.body;
    if (!title) return res.status(400).json({ error: 'Informe o título do contrato.' });
    if (!attachedFileUrl && !content) return res.status(400).json({ error: 'Informe o conteúdo do contrato (ou anexe um arquivo pronto).' });
    if (!Array.isArray(signers) || signers.length < 1) return res.status(400).json({ error: 'Informe ao menos um signatário (nome e e-mail).' });
    if (signers.some(s => !s.name || !s.email)) return res.status(400).json({ error: 'Todo signatário precisa de nome e e-mail.' });

    const companyId = req.user.role === 'client_admin' ? req.user.companyId : (company_id || null);
    if (req.user.role === 'client_admin' && employee_id) {
        const pertence = await dbGet(`SELECT id FROM employees WHERE id = ? AND company_id = ?`, [employee_id, req.user.companyId]);
        if (!pertence) return res.status(403).json({ error: 'Este executivo não pertence à sua empresa.' });
    }

    try {
        const pdfBuffer = attachedFileUrl ? lerArquivoAnexadoContrato(attachedFileUrl) : await gerarPdfContratoDeTexto(title, content);
        const conteudoSalvo = attachedFileUrl ? `[Arquivo anexado pelo usuário: ${attachedFileName || attachedFileUrl}]` : content;
        const nomeArquivoEnvio = attachedFileUrl ? (attachedFileName || path.basename(attachedFileUrl)) : `${title}.pdf`;
        const resultado = await enviarContratoParaAutentique(title, pdfBuffer, signers, nomeArquivoEnvio);

        const novoId = await new Promise((resolve, reject) => db.run(
            `INSERT INTO contracts (type, company_id, employee_id, mentor_id, title, content, status, autentique_document_id, signers_json, created_by)
             VALUES (?, ?, ?, ?, ?, ?, 'enviado', ?, ?, ?)`,
            [type || 'empresa', companyId, employee_id || null, mentor_id || null, title, conteudoSalvo, resultado.id, JSON.stringify(resultado.signatures), req.user.userId],
            function (err) { err ? reject(err) : resolve(this.lastID); }
        ));

        // Deixa um recado na tela de login/notificações do contratante — sem
        // isso, a empresa só ficava sabendo do contrato se alguém avisasse por
        // fora, mesmo o e-mail do Autentique podendo cair no spam.
        if (companyId) notificarPorCompanyAdmins(companyId, 'Novo contrato para assinatura', `"${title}" foi enviado e está aguardando sua assinatura.`, 'contracts');

        res.json({ message: 'Contrato enviado para assinatura! Os signatários recebem um e-mail do Autentique.', signatures: resultado.signatures, id: novoId });
    } catch (e) {
        console.error('Erro ao enviar contrato para o Autentique:', e.message);
        res.status(400).json({ error: 'Erro ao gerar/enviar o contrato para assinatura: ' + e.message });
    }
});

// Edita um contrato que AINDA NÃO foi assinado — recria o documento no
// Autentique com o conteúdo atualizado e reenvia para assinatura (o link
// antigo deixa de valer na prática, já que o registro passa a apontar para
// este novo documento). Bloqueado assim que alguém já assinou.
app.put('/api/contracts/:id', requireRole('admin'), async (req, res) => {
    if (!AUTENTIQUE_API_TOKEN) return res.status(503).json({ error: 'Assinatura digital ainda não foi configurada no servidor.' });
    const { title, content, signers, attachedFileUrl, attachedFileName } = req.body;
    try {
        const contrato = await dbGet(`SELECT * FROM contracts WHERE id = ?`, [req.params.id]);
        if (!contrato) return res.status(404).json({ error: 'Contrato não encontrado.' });
        if (contrato.status === 'assinado') return res.status(400).json({ error: 'Este contrato já foi assinado e não pode mais ser editado.' });

        const tituloFinal = title || contrato.title;
        if (!attachedFileUrl && !content) return res.status(400).json({ error: 'Informe o conteúdo do contrato (ou anexe um arquivo pronto).' });
        if (!Array.isArray(signers) || signers.length < 1) return res.status(400).json({ error: 'Informe ao menos um signatário (nome e e-mail).' });
        if (signers.some(s => !s.name || !s.email)) return res.status(400).json({ error: 'Todo signatário precisa de nome e e-mail.' });

        const pdfBuffer = attachedFileUrl ? lerArquivoAnexadoContrato(attachedFileUrl) : await gerarPdfContratoDeTexto(tituloFinal, content);
        const conteudoSalvo = attachedFileUrl ? `[Arquivo anexado pelo usuário: ${attachedFileName || attachedFileUrl}]` : content;
        const nomeArquivoEnvio = attachedFileUrl ? (attachedFileName || path.basename(attachedFileUrl)) : `${tituloFinal}.pdf`;
        const resultado = await enviarContratoParaAutentique(tituloFinal, pdfBuffer, signers, nomeArquivoEnvio);

        await new Promise((resolve, reject) => db.run(
            `UPDATE contracts SET title = ?, content = ?, status = 'enviado', autentique_document_id = ?, signers_json = ?, signed_file_url = NULL WHERE id = ?`,
            [tituloFinal, conteudoSalvo, resultado.id, JSON.stringify(resultado.signatures), req.params.id],
            (err) => err ? reject(err) : resolve()
        ));

        if (contrato.company_id) notificarPorCompanyAdmins(contrato.company_id, 'Contrato atualizado — nova assinatura necessária', `"${tituloFinal}" foi atualizado e reenviado para assinatura.`, 'contracts');

        res.json({ message: 'Contrato atualizado e reenviado para assinatura!' });
    } catch (e) {
        console.error('Erro ao editar/reenviar contrato:', e.message);
        res.status(400).json({ error: 'Erro ao atualizar o contrato: ' + e.message });
    }
});

// "Reenviar" não recria o documento no Autentique (evitando duplicar o
// histórico) — só manda de novo, por e-mail, o mesmo link de assinatura que o
// Autentique já gerou na criação, só para quem ainda não assinou.
app.post('/api/contracts/:id/resend', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const contrato = await dbGet(`SELECT * FROM contracts WHERE id = ?`, [req.params.id]);
        if (!contrato) return res.status(404).json({ error: 'Contrato não encontrado.' });
        if (req.user.role === 'client_admin' && contrato.company_id !== req.user.companyId) return res.status(403).json({ error: 'Este contrato não pertence à sua empresa.' });
        if (contrato.status === 'assinado') return res.status(400).json({ error: 'Este contrato já foi assinado.' });

        const signatarios = JSON.parse(contrato.signers_json || '[]');
        const pendentes = signatarios.filter(s => !s.signed && s.email && s.link && s.link.short_link);
        if (!pendentes.length) return res.status(400).json({ error: 'Nenhum signatário pendente com link de assinatura disponível para reenviar (tente "Verificar Status Agora" primeiro).' });

        for (const s of pendentes) {
            try {
                await transporter.sendMail({
                    from: process.env.SMTP_FROM || '"Impulsionar V4" <no-reply@impulsionar.com>',
                    to: s.email,
                    subject: `Lembrete: assinatura pendente — ${contrato.title}`,
                    text: `Olá, ${s.name}!\n\nVocê ainda não assinou o documento "${contrato.title}". Acesse o link abaixo para assinar:\n${s.link.short_link}\n\nEquipe Impulsionar V4.`
                });
            } catch (erroEnvio) {
                console.warn('⚠️  Falha ao reenviar e-mail de assinatura para', s.email, erroEnvio.message);
            }
        }
        if (contrato.company_id) notificarPorCompanyAdmins(contrato.company_id, 'Assinatura pendente — lembrete', `"${contrato.title}" ainda está aguardando sua assinatura.`, 'contracts');
        res.json({ message: `Lembrete reenviado para ${pendentes.length} signatário(s) pendente(s).` });
    } catch (e) { res.status(400).json({ error: 'Erro ao reenviar o contrato.' }); }
});

// Consulta o status mais recente direto na API do Autentique (mesma lógica de
// "Verificar Status Agora" já usada nas assinaturas do Mercado Pago).
app.post('/api/contracts/:id/refresh-status', requireRole('admin', 'client_admin'), async (req, res) => {
    if (!AUTENTIQUE_API_TOKEN) return res.status(503).json({ error: 'Assinatura digital ainda não foi configurada no servidor.' });
    try {
        const contrato = await dbGet(`SELECT * FROM contracts WHERE id = ?`, [req.params.id]);
        if (!contrato) return res.status(404).json({ error: 'Contrato não encontrado.' });
        if (req.user.role === 'client_admin' && contrato.company_id !== req.user.companyId) return res.status(403).json({ error: 'Este contrato não pertence à sua empresa.' });

        const documento = await consultarContratoNoAutentique(contrato.autentique_document_id);
        const todosAssinaram = documento.signatures.every(s => s.signed);
        const algumRecusou = documento.signatures.some(s => s.rejected);
        const novoStatus = algumRecusou ? 'recusado' : (todosAssinaram ? 'assinado' : 'enviado');

        await new Promise((resolve, reject) => db.run(
            `UPDATE contracts SET status = ?, signers_json = ?, signed_file_url = ? WHERE id = ?`,
            [novoStatus, JSON.stringify(documento.signatures), documento.files ? documento.files.signed : null, req.params.id],
            (err) => err ? reject(err) : resolve()
        ));

        res.json({ message: 'Status atualizado!', status: novoStatus });
    } catch (e) {
        console.error('Erro ao consultar status do contrato no Autentique:', e.message);
        res.status(400).json({ error: 'Erro ao consultar o status no Autentique.' });
    }
});

app.delete('/api/contracts/:id', requireRole('admin'), async (req, res) => {
    try {
        const contrato = await dbGet(`SELECT * FROM contracts WHERE id = ?`, [req.params.id]);
        if (!contrato) return res.status(404).json({ error: 'Contrato não encontrado.' });
        db.run(`DELETE FROM contracts WHERE id = ?`, [req.params.id], () => res.json({ message: 'Removido!' }));
    } catch (e) { res.status(500).json({ error: 'Erro ao remover contrato.' }); }
});

// Webhook do Autentique (opcional — configurável no painel deles). Faz o
// mesmo trabalho do "Verificar Status Agora", só que automaticamente.
app.post('/api/webhooks/autentique', async (req, res) => {
    try {
        const documentId = req.body && req.body.event && req.body.event.data && req.body.event.data.object && req.body.event.data.object.id;
        if (!documentId || !AUTENTIQUE_API_TOKEN) return res.sendStatus(200);
        const contrato = await dbGet(`SELECT * FROM contracts WHERE autentique_document_id = ?`, [documentId]);
        if (!contrato) return res.sendStatus(200);
        const documento = await consultarContratoNoAutentique(documentId);
        const todosAssinaram = documento.signatures.every(s => s.signed);
        const algumRecusou = documento.signatures.some(s => s.rejected);
        const novoStatus = algumRecusou ? 'recusado' : (todosAssinaram ? 'assinado' : 'enviado');
        db.run(`UPDATE contracts SET status = ?, signers_json = ?, signed_file_url = ? WHERE id = ?`,
            [novoStatus, JSON.stringify(documento.signatures), documento.files ? documento.files.signed : null, contrato.id], () => {});
        res.sendStatus(200);
    } catch (e) {
        console.warn('⚠️  Erro ao processar webhook do Autentique:', e.message);
        res.sendStatus(200);
    }
});

// ============================================================
// PLANOS POR EMPRESA
// ============================================================

app.get('/api/plans', (req, res) => {
    db.all(`SELECT * FROM plans ORDER BY id ASC`, [], (err, rows) => res.json(rows || []));
});

app.post('/api/plans', requireRole('admin'), (req, res) => {
    const { name, max_employees, price_display, features_text, mp_price, trial_days } = req.body;
    if (!name) return res.status(400).json({ error: 'Informe o nome do plano.' });
    db.run(`INSERT INTO plans (name, max_employees, price_display, features_text, mp_price, trial_days) VALUES (?, ?, ?, ?, ?, ?)`,
        [name, max_employees || null, price_display || '', features_text || '', mp_price || null, trial_days || 0], function (err) {
            if (err) return res.status(400).json({ error: err.message });
            res.json({ message: 'Plano criado!', id: this.lastID });
        });
});

app.put('/api/plans/:id', requireRole('admin'), (req, res) => {
    const { name, max_employees, price_display, features_text, mp_price, trial_days } = req.body;
    db.run(`UPDATE plans SET name = ?, max_employees = ?, price_display = ?, features_text = ?, mp_price = ?, trial_days = ? WHERE id = ?`,
        [name, max_employees || null, price_display || '', features_text || '', mp_price || null, trial_days || 0, req.params.id], (err) => {
            if (err) return res.status(400).json({ error: err.message });
            res.json({ message: 'Plano atualizado!' });
        });
});

app.delete('/api/plans/:id', requireRole('admin'), (req, res) => {
    db.run(`UPDATE companies SET plan_id = NULL WHERE plan_id = ?`, [req.params.id], () => {
        db.run(`DELETE FROM plans WHERE id = ?`, [req.params.id], () => res.json({ message: 'Removido!' }));
    });
});

// ============================================================
// ASSINATURAS RECORRENTES (MERCADO PAGO — PREAPPROVAL)
// ============================================================

// Garante que quem chama só mexe na assinatura da própria empresa (client_admin)
// ou em qualquer empresa (admin/Master).
function ensureCompanyAccess(req, res, companyId) {
    if (req.user.role === 'client_admin' && String(companyId) !== String(req.user.companyId)) {
        res.status(403).json({ error: 'Você só pode gerenciar a assinatura da sua própria corporação.' });
        return false;
    }
    if (!['admin', 'client_admin'].includes(req.user.role)) {
        res.status(403).json({ error: 'Perfil sem permissão.' });
        return false;
    }
    return true;
}

// Busca o e-mail de contato da empresa (o do gestor/client_admin) para usar como
// payer_email na assinatura do Mercado Pago.
async function obterEmailContatoEmpresa(companyId) {
    const gestor = await dbGet(`SELECT email FROM users WHERE company_id = ? AND role = 'client_admin' ORDER BY id ASC LIMIT 1`, [companyId]);
    return gestor ? gestor.email : null;
}

// O SDK do Mercado Pago costuma jogar o motivo real do erro dentro de
// "cause" (um array de {code, description}) em vez de em e.message — sem
// isso, toda falha aparecia como o mesmo aviso genérico e ninguém conseguia
// saber o que corrigir sem acesso ao log do servidor. Essa função extrai o
// texto mais útil possível para mostrar direto na tela de quem está tentando
// assinar/trocar de plano.
function detalheErroMercadoPago(e) {
    try {
        if (Array.isArray(e?.cause) && e.cause.length) {
            return e.cause.map(c => c.description || c.message || JSON.stringify(c)).filter(Boolean).join(' | ');
        }
        if (e?.cause?.description) return e.cause.description;
        return e?.message || 'Erro desconhecido ao comunicar com o Mercado Pago.';
    } catch (err) {
        return e?.message || 'Erro desconhecido ao comunicar com o Mercado Pago.';
    }
}

// O Mercado Pago EXIGE um back_url válido (http/https, não-localhost) para
// criar uma assinatura (PreApproval) — se a "URL Pública do Sistema" (Meu
// Perfil) ainda não estiver configurada corretamente, ou tiver sido
// resetada (ex.: um redeploy que sobrescreveu o banco), a chamada à API
// falha com "Invalid value for back_url". Em vez de deixar o Mercado Pago
// devolver esse erro técnico confuso, verificamos antes e explicamos o que
// falta configurar.
function obterBackUrlAssinaturaOuErro(res) {
    if (urlPublicaValida(appBaseUrlAtiva)) return appBaseUrlAtiva;
    if (urlPublicaValida(APP_BASE_URL_ENV)) return APP_BASE_URL_ENV;
    res.status(400).json({
        error: 'A "URL Pública do Sistema" ainda não está configurada corretamente (Meu Perfil > URL Pública do Sistema). ' +
               'Ela precisa ser um endereço público começando com https:// (ex.: https://www.impulsionarv4.com.br), não localhost. ' +
               'Configure-a e tente novamente.'
    });
    return null;
}

// Notifica todos os gestores (client_admin) de uma empresa sobre mudanças na assinatura.
function notificarGestoresDaEmpresa(companyId, title, message) {
    db.all(`SELECT id FROM users WHERE company_id = ? AND role = 'client_admin'`, [companyId], (err, gestores) => {
        if (err || !gestores) return;
        gestores.forEach(g => notificar(g.id, title, message, 'minhaEmpresa'));
    });
}

app.post('/api/companies/:id/subscribe', async (req, res) => {
    const { id } = req.params;
    if (!ensureCompanyAccess(req, res, id)) return;
    if (!mpPreApproval) return res.status(503).json({ error: 'Mercado Pago ainda não foi configurado no servidor (defina MP_ACCESS_TOKEN no .env).' });
    try {
        const empresa = await dbGet(`SELECT * FROM companies WHERE id = ?`, [id]);
        if (!empresa) return res.status(404).json({ error: 'Empresa não encontrada.' });
        if (!empresa.plan_id) return res.status(400).json({ error: 'Atribua um plano a esta empresa antes de iniciar a assinatura.' });
        const plano = await dbGet(`SELECT * FROM plans WHERE id = ?`, [empresa.plan_id]);
        if (!plano || !plano.mp_price) return res.status(400).json({ error: 'Este plano ainda não tem um preço configurado para cobrança recorrente.' });
        if (empresa.subscription_status === 'authorized') return res.status(400).json({ error: 'Esta empresa já tem uma assinatura ativa.' });

        const emailContato = await obterEmailContatoEmpresa(id);
        if (!emailContato) return res.status(400).json({ error: 'Cadastre um gestor (client_admin) com e-mail para esta empresa antes de assinar.' });

        const backUrl = obterBackUrlAssinaturaOuErro(res);
        if (!backUrl) return;

        const corpo = {
            reason: `Impulsionar V4 — Plano ${plano.name}`,
            external_reference: `company:${id}`,
            payer_email: emailContato,
            back_url: backUrl,
            auto_recurring: {
                frequency: 1,
                frequency_type: 'months',
                transaction_amount: Number(plano.mp_price),
                currency_id: 'BRL'
            },
            status: 'pending'
        };
        if (plano.trial_days && Number(plano.trial_days) > 0) {
            corpo.auto_recurring.free_trial = { frequency: Number(plano.trial_days), frequency_type: 'days' };
        }

        const resultado = await mpPreApproval.create({ body: corpo });
        const dataTrial = plano.trial_days && Number(plano.trial_days) > 0
            ? new Date(Date.now() + Number(plano.trial_days) * 86400000).toISOString()
            : null;

        await new Promise((resolve, reject) => db.run(
            `UPDATE companies SET mp_preapproval_id = ?, subscription_status = ?, subscription_updated_at = CURRENT_TIMESTAMP, trial_ends_at = ? WHERE id = ?`,
            [resultado.id, resultado.status || 'pending', dataTrial, id],
            (err) => err ? reject(err) : resolve()
        ));

        res.json({ message: 'Assinatura iniciada! Complete a autorização no Mercado Pago.', initPoint: resultado.init_point, status: resultado.status });
    } catch (e) {
        const detalhe = detalheErroMercadoPago(e);
        console.error('Erro ao criar assinatura Mercado Pago:', detalhe);
        res.status(400).json({ error: `Erro ao iniciar assinatura no Mercado Pago: ${detalhe}` });
    }
});

app.get('/api/companies/:id/subscription', async (req, res) => {
    const { id } = req.params;
    if (!ensureCompanyAccess(req, res, id)) return;
    try {
        const empresa = await dbGet(`SELECT c.*, p.name as planName, p.mp_price, p.trial_days FROM companies c LEFT JOIN plans p ON p.id = c.plan_id WHERE c.id = ?`, [id]);
        if (!empresa) return res.status(404).json({ error: 'Empresa não encontrada.' });
        let pendingPlanName = null;
        if (empresa.pending_plan_id) {
            const pendente = await dbGet(`SELECT name FROM plans WHERE id = ?`, [empresa.pending_plan_id]);
            pendingPlanName = pendente ? pendente.name : null;
        }
        res.json({
            planName: empresa.planName,
            mp_price: empresa.mp_price,
            trial_days: empresa.trial_days,
            subscriptionStatus: empresa.subscription_status || 'none',
            trialEndsAt: empresa.trial_ends_at,
            hasPreapproval: !!empresa.mp_preapproval_id,
            mercadoPagoConfigurado: !!mpPreApproval,
            pendingPlanId: empresa.pending_plan_id || null,
            pendingPlanName,
            cancellationRequestedAt: empresa.cancellation_requested_at || null
        });
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar assinatura.' }); }
});

// Consulta o status da assinatura direto na API do Mercado Pago e atualiza o
// banco local. Alternativa ao webhook para quem está testando localmente sem
// um endereço público (ngrok, etc) — clique manual em vez de notificação automática.
app.post('/api/companies/:id/subscription/refresh', async (req, res) => {
    const { id } = req.params;
    if (!ensureCompanyAccess(req, res, id)) return;
    if (!mpPreApproval) return res.status(503).json({ error: 'Mercado Pago ainda não foi configurado no servidor (defina MP_ACCESS_TOKEN no .env).' });
    try {
        const empresa = await dbGet(`SELECT mp_preapproval_id, pending_plan_id FROM companies WHERE id = ?`, [id]);
        if (!empresa || !empresa.mp_preapproval_id) return res.status(400).json({ error: 'Esta empresa ainda não iniciou nenhuma assinatura.' });
        const dadosAtualizados = await mpPreApproval.get({ id: empresa.mp_preapproval_id });

        if (dadosAtualizados.status === 'authorized' && empresa.pending_plan_id) {
            // Pagamento confirmado — agora sim libera o plano (e o crédito de funcionário) escolhido.
            await new Promise((resolve, reject) => db.run(
                `UPDATE companies SET subscription_status = ?, subscription_updated_at = CURRENT_TIMESTAMP, plan_id = ?, pending_plan_id = NULL WHERE id = ?`,
                [dadosAtualizados.status, empresa.pending_plan_id, id], (err) => err ? reject(err) : resolve()
            ));
        } else {
            await new Promise((resolve, reject) => db.run(
                `UPDATE companies SET subscription_status = ?, subscription_updated_at = CURRENT_TIMESTAMP WHERE id = ?`,
                [dadosAtualizados.status, id], (err) => err ? reject(err) : resolve()
            ));
        }
        if (dadosAtualizados.status === 'authorized') notificarGestoresDaEmpresa(id, 'Assinatura ativada!', 'Sua assinatura foi confirmada no Mercado Pago.');
        res.json({ message: 'Status atualizado!', subscriptionStatus: dadosAtualizados.status });
    } catch (e) {
        console.error('Erro ao verificar status da assinatura Mercado Pago:', e.message);
        res.status(400).json({ error: 'Erro ao consultar o status no Mercado Pago.' });
    }
});

// Cancelamento de fato só o Master confirma — a empresa (client_admin) só pode
// SOLICITAR (endpoint abaixo). Isso evita que a empresa saia clicando um botão
// e derrube a cobrança recorrente sem a Impulsionar saber/negociar antes.
app.post('/api/companies/:id/subscription/cancel', requireRole('admin'), async (req, res) => {
    const { id } = req.params;
    if (!mpPreApproval) return res.status(503).json({ error: 'Mercado Pago ainda não foi configurado no servidor.' });
    try {
        const empresa = await dbGet(`SELECT mp_preapproval_id FROM companies WHERE id = ?`, [id]);
        if (!empresa || !empresa.mp_preapproval_id) return res.status(400).json({ error: 'Esta empresa não tem assinatura ativa para cancelar.' });
        await mpPreApproval.update({ id: empresa.mp_preapproval_id, body: { status: 'cancelled' } });
        await new Promise((resolve, reject) => db.run(
            `UPDATE companies SET subscription_status = 'cancelled', subscription_updated_at = CURRENT_TIMESTAMP, cancellation_requested_at = NULL WHERE id = ?`,
            [id], (err) => err ? reject(err) : resolve()
        ));
        notificarGestoresDaEmpresa(id, 'Assinatura cancelada', 'A cobrança recorrente desta empresa foi cancelada.');
        res.json({ message: 'Assinatura cancelada.' });
    } catch (e) {
        console.error('Erro ao cancelar assinatura Mercado Pago:', e.message);
        res.status(400).json({ error: 'Erro ao cancelar a assinatura no Mercado Pago.' });
    }
});

// A empresa solicita o cancelamento — não cancela na hora, só avisa o Master,
// que decide (confirma o cancelamento ou recusa o pedido e segue cobrando).
app.post('/api/companies/:id/subscription/request-cancellation', async (req, res) => {
    const { id } = req.params;
    if (!ensureCompanyAccess(req, res, id)) return;
    try {
        const empresa = await dbGet(`SELECT name, subscription_status, cancellation_requested_at FROM companies WHERE id = ?`, [id]);
        if (!empresa) return res.status(404).json({ error: 'Empresa não encontrada.' });
        if (!['authorized', 'pending', 'paused'].includes(empresa.subscription_status)) {
            return res.status(400).json({ error: 'Esta empresa não tem assinatura ativa para cancelar.' });
        }
        if (empresa.cancellation_requested_at) {
            return res.status(400).json({ error: 'Já existe um pedido de cancelamento em análise pelo Master.' });
        }
        await new Promise((resolve, reject) => db.run(
            `UPDATE companies SET cancellation_requested_at = CURRENT_TIMESTAMP WHERE id = ?`,
            [id], (err) => err ? reject(err) : resolve()
        ));
        db.all(`SELECT id FROM users WHERE role = 'admin'`, [], (e, admins) => {
            if (!e) admins.forEach(a => notificar(a.id, 'Pedido de cancelamento de plano', `${empresa.name} solicitou o cancelamento da assinatura.`, 'companies'));
        });
        res.json({ message: 'Pedido de cancelamento enviado! O Master vai analisar e confirmar em breve.' });
    } catch (e) {
        console.error('Erro ao solicitar cancelamento:', e.message);
        res.status(400).json({ error: 'Erro ao registrar o pedido de cancelamento.' });
    }
});

// Master recusa o pedido de cancelamento (a assinatura continua ativa normalmente).
app.post('/api/companies/:id/subscription/request-cancellation/dismiss', requireRole('admin'), async (req, res) => {
    const { id } = req.params;
    try {
        await new Promise((resolve, reject) => db.run(
            `UPDATE companies SET cancellation_requested_at = NULL WHERE id = ?`,
            [id], (err) => err ? reject(err) : resolve()
        ));
        notificarGestoresDaEmpresa(id, 'Pedido de cancelamento não aprovado', 'A Impulsionar entrará em contato. Sua assinatura continua ativa normalmente.');
        res.json({ message: 'Pedido de cancelamento recusado — assinatura continua ativa.' });
    } catch (e) {
        res.status(400).json({ error: 'Erro ao recusar o pedido de cancelamento.' });
    }
});

// Troca/adesão de plano pela empresa (tela "Planos Disponíveis" / self-service).
//
// Duas situações bem diferentes:
// 1) A empresa JÁ tem uma assinatura ativa e autorizada no Mercado Pago — nesse
//    caso ela já é uma cliente pagante, então um upgrade/downgrade só ajusta o
//    valor cobrado dali pra frente no preapproval existente, sem exigir nova
//    autorização, e o plan_id muda na hora.
// 2) A empresa NÃO tem assinatura ativa ainda (primeira adesão, ou assinatura
//    cancelada) — nesse caso o plan_id NÃO pode mudar na hora, porque isso
//    liberaria crédito de cadastro de funcionário de graça. Em vez disso,
//    criamos uma nova cobrança no Mercado Pago e guardamos o plano escolhido em
//    "pending_plan_id"; só quando o pagamento for confirmado (via webhook ou
//    "Verificar Status Agora") é que o plan_id de fato muda.
app.post('/api/companies/:id/subscription/change-plan', async (req, res) => {
    const { id } = req.params;
    if (!ensureCompanyAccess(req, res, id)) return;
    const { planId } = req.body;
    if (!planId) return res.status(400).json({ error: 'Informe o novo plano.' });
    try {
        const novoPlano = await dbGet(`SELECT * FROM plans WHERE id = ?`, [planId]);
        if (!novoPlano) return res.status(404).json({ error: 'Plano não encontrado.' });
        const empresa = await dbGet(`SELECT * FROM companies WHERE id = ?`, [id]);
        if (!empresa) return res.status(404).json({ error: 'Empresa não encontrada.' });

        const jaEhAssinantePagante = !!(empresa.mp_preapproval_id && empresa.subscription_status === 'authorized');

        if (jaEhAssinantePagante) {
            // Upgrade/downgrade de quem já paga: ajusta o valor cobrado e troca na hora.
            if (!mpPreApproval) return res.status(503).json({ error: 'Mercado Pago ainda não foi configurado no servidor.' });
            if (!novoPlano.mp_price) return res.status(400).json({ error: 'O novo plano ainda não tem preço configurado para cobrança recorrente. Fale com a Impulsionar.' });
            await mpPreApproval.update({
                id: empresa.mp_preapproval_id,
                body: { auto_recurring: { transaction_amount: Number(novoPlano.mp_price) } }
            });
            await new Promise((resolve, reject) => db.run(
                `UPDATE companies SET plan_id = ?, pending_plan_id = NULL WHERE id = ?`, [planId, id],
                (err) => err ? reject(err) : resolve()
            ));
            return res.json({ message: 'Plano alterado com sucesso!' });
        }

        // Primeira adesão (ou assinatura anterior cancelada): precisa pagar antes.
        if (!mpPreApproval) return res.status(503).json({ error: 'Mercado Pago ainda não foi configurado no servidor (defina MP_ACCESS_TOKEN no .env).' });
        if (!novoPlano.mp_price) return res.status(400).json({ error: 'Este plano ainda não tem preço de cobrança recorrente configurado — fale com a Impulsionar para contratar.' });

        const emailContato = await obterEmailContatoEmpresa(id);
        if (!emailContato) return res.status(400).json({ error: 'Cadastre um gestor (client_admin) com e-mail para esta empresa antes de assinar.' });

        const backUrl = obterBackUrlAssinaturaOuErro(res);
        if (!backUrl) return;

        const corpo = {
            reason: `Impulsionar V4 — Plano ${novoPlano.name}`,
            external_reference: `company:${id}`,
            payer_email: emailContato,
            back_url: backUrl,
            auto_recurring: {
                frequency: 1,
                frequency_type: 'months',
                transaction_amount: Number(novoPlano.mp_price),
                currency_id: 'BRL'
            },
            status: 'pending'
        };
        if (novoPlano.trial_days && Number(novoPlano.trial_days) > 0) {
            corpo.auto_recurring.free_trial = { frequency: Number(novoPlano.trial_days), frequency_type: 'days' };
        }

        const resultado = await mpPreApproval.create({ body: corpo });
        const dataTrial = novoPlano.trial_days && Number(novoPlano.trial_days) > 0
            ? new Date(Date.now() + Number(novoPlano.trial_days) * 86400000).toISOString()
            : null;

        await new Promise((resolve, reject) => db.run(
            `UPDATE companies SET mp_preapproval_id = ?, subscription_status = ?, subscription_updated_at = CURRENT_TIMESTAMP, trial_ends_at = ?, pending_plan_id = ? WHERE id = ?`,
            [resultado.id, resultado.status || 'pending', dataTrial, planId, id],
            (err) => err ? reject(err) : resolve()
        ));

        res.json({
            message: 'Quase lá! Complete o pagamento na aba do Mercado Pago — os créditos deste plano são liberados assim que o pagamento for confirmado.',
            initPoint: resultado.init_point,
            status: resultado.status
        });
    } catch (e) {
        const detalhe = detalheErroMercadoPago(e);
        console.error('Erro ao trocar plano da assinatura:', detalhe);
        res.status(400).json({ error: `Erro ao iniciar o pagamento no Mercado Pago: ${detalhe}` });
    }
});

// ============================================================
// DPO AMBEV — consultoria de processos por pilar (autoavaliação mensal)
// ============================================================

// Pilares que a empresa já tem liberados: todos, se ela comprou a "Consultoria
// Completa", ou só os avulsos que ela pagou individualmente.
async function pilaresAtivosDaEmpresa(companyId) {
    const todas = await dbAll(`SELECT * FROM dpo_purchases WHERE company_id = ? AND status = 'paid'`, [companyId]);
    const agora = Date.now();
    // Um "trial" com prazo vencido não conta mais como ativo (mas a linha fica
    // no banco pra histórico — não precisa apagar nem mudar o status).
    const compras = todas.filter(c => !c.is_trial || !c.trial_expires_at || new Date(c.trial_expires_at).getTime() > agora);
    if (compras.some(c => c.scope === 'completo')) return DPO_PILARES_ORDEM.slice();
    return [...new Set(compras.filter(c => c.scope === 'pilar').map(c => c.pillar_key))];
}

// Busca um ciclo garantindo que o usuário logado pode vê-lo/editá-lo (admin
// sempre pode; client_admin só o ciclo da própria empresa). Retorna null (e já
// responde o erro) quando não pode.
async function obterCicloComAcesso(req, res, cicloId) {
    const ciclo = await dbGet(`SELECT * FROM dpo_audit_cycles WHERE id = ?`, [cicloId]);
    if (!ciclo) { res.status(404).json({ error: 'Ciclo de consultoria não encontrado.' }); return null; }
    if (req.user.role === 'client_admin' && String(ciclo.company_id) !== String(req.user.companyId)) {
        res.status(403).json({ error: 'Este ciclo não pertence à sua empresa.' });
        return null;
    }
    return ciclo;
}

// Monta a lista de perguntas (dado estático) de um pilar, já mesclada com as
// respostas/planos de ação/follow-ups já salvos naquele ciclo.
async function montarPilarDoCiclo(cicloId, pilarKey) {
    const pilar = DPO_AMBEV_DATA[pilarKey];
    if (!pilar) return null;
    const respostas = await dbAll(`SELECT question_key, score FROM dpo_answers WHERE cycle_id = ? AND question_key LIKE ?`, [cicloId, `${pilarKey}:%`]);
    const mapaRespostas = {};
    respostas.forEach(r => { mapaRespostas[r.question_key] = r.score; });
    const planos = await dbAll(`SELECT * FROM dpo_action_plans WHERE cycle_id = ? AND question_key LIKE ? ORDER BY created_at ASC`, [cicloId, `${pilarKey}:%`]);
    const planoIds = planos.map(p => p.id);
    let followsPorPlano = {};
    if (planoIds.length) {
        const follows = await dbAll(`SELECT * FROM dpo_follow_ups WHERE action_plan_id IN (${planoIds.map(() => '?').join(',')}) ORDER BY numero ASC`, planoIds);
        follows.forEach(f => {
            if (!followsPorPlano[f.action_plan_id]) followsPorPlano[f.action_plan_id] = [];
            followsPorPlano[f.action_plan_id].push(f);
        });
    }
    const grupos = pilar.grupos.map(g => ({
        numero: g.numero,
        titulo: g.titulo,
        perguntas: g.perguntas.map(p => {
            const key = `${pilarKey}:${p.numero}`;
            const planosDaPergunta = planos.filter(pl => pl.question_key === key).map(pl => ({ ...pl, follows: followsPorPlano[pl.id] || [] }));
            return { ...p, questionKey: key, score: (key in mapaRespostas) ? mapaRespostas[key] : null, planosDeAcao: planosDaPergunta };
        })
    }));
    return { key: pilarKey, label: pilar.label, grupos };
}

// Catálogo de pilares + preços. Para client_admin, já vem com o que a empresa
// já comprou/paga e o que está pendente de pagamento.
app.get('/api/dpo/pillars', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const precos = await dbAll(`SELECT * FROM dpo_pillar_prices`);
        const mapaPrecos = {};
        precos.forEach(p => { mapaPrecos[p.pillar_key] = p; });
        const precoCompletoRow = await dbGet(`SELECT value FROM integration_settings WHERE key = 'dpo_full_audit_price'`);
        const precoCompleto = precoCompletoRow ? Number(precoCompletoRow.value) : 0;

        const pilares = DPO_PILARES_ORDEM.map((chave, i) => ({
            key: chave,
            numero: i + 1,
            label: DPO_AMBEV_DATA[chave].label,
            totalPerguntas: DPO_AMBEV_DATA[chave].grupos.reduce((soma, g) => soma + g.perguntas.length, 0),
            price: mapaPrecos[chave] ? Number(mapaPrecos[chave].price) : 0,
            active: mapaPrecos[chave] ? !!mapaPrecos[chave].active : true
        }));

        let meusPilares = null, compraCompleta = false, pendentes = [];
        if (req.user.role === 'client_admin') {
            const compras = await dbAll(`SELECT * FROM dpo_purchases WHERE company_id = ? ORDER BY created_at DESC`, [req.user.companyId]);
            compraCompleta = compras.some(c => c.scope === 'completo' && c.status === 'paid');
            meusPilares = compraCompleta ? DPO_PILARES_ORDEM.slice() : [...new Set(compras.filter(c => c.scope === 'pilar' && c.status === 'paid').map(c => c.pillar_key))];
            pendentes = compras.filter(c => c.status === 'pending_payment');
        }
        res.json({ pilares, precoCompleto, meusPilares, compraCompleta, pendentes });
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar os pilares do DPO Ambev.' }); }
});

app.put('/api/admin/dpo/pricing', requireRole('admin'), async (req, res) => {
    const { precos, precoCompleto } = req.body;
    try {
        if (precos && typeof precos === 'object') {
            for (const chave of Object.keys(precos)) {
                if (!DPO_PILARES_ORDEM.includes(chave)) continue;
                await new Promise((resolve, reject) => db.run(
                    `UPDATE dpo_pillar_prices SET price = ? WHERE pillar_key = ?`,
                    [Number(precos[chave]) || 0, chave], (err) => err ? reject(err) : resolve()
                ));
            }
        }
        if (precoCompleto !== undefined) {
            await new Promise((resolve, reject) => db.run(
                `INSERT INTO integration_settings (key, value) VALUES ('dpo_full_audit_price', ?)
                 ON CONFLICT(key) DO UPDATE SET value = excluded.value`,
                [String(Number(precoCompleto) || 0)], (err) => err ? reject(err) : resolve()
            ));
        }
        res.json({ message: 'Preços do DPO Ambev atualizados!' });
    } catch (e) { res.status(400).json({ error: 'Erro ao salvar os preços.' }); }
});

// Cadastro de consultores do DPO Ambev — quem faz a consultoria de verdade
// junto com a empresa. A empresa escolhe um deles na hora de comprar.
app.get('/api/admin/dpo/consultants', requireRole('admin'), async (req, res) => {
    try { res.json(await dbAll(`SELECT * FROM dpo_consultants ORDER BY name ASC`)); }
    catch (e) { res.status(500).json({ error: 'Erro ao carregar os consultores.' }); }
});

// Listagem simplificada (só ativos) usada pela empresa na hora da compra.
app.get('/api/dpo/consultants', requireRole('admin', 'client_admin'), async (req, res) => {
    try { res.json(await dbAll(`SELECT id, name, bio FROM dpo_consultants WHERE active = 1 ORDER BY name ASC`)); }
    catch (e) { res.status(500).json({ error: 'Erro ao carregar os consultores.' }); }
});

app.post('/api/admin/dpo/consultants', requireRole('admin'), async (req, res) => {
    const { name, email, phone, bio } = req.body;
    if (!name) return res.status(400).json({ error: 'Informe o nome do consultor.' });
    db.run(`INSERT INTO dpo_consultants (name, email, phone, bio) VALUES (?, ?, ?, ?)`,
        [name, email || null, phone || null, bio || null], function (err) {
            if (err) return res.status(400).json({ error: 'Erro ao cadastrar o consultor.' });
            res.json({ message: 'Consultor cadastrado!', id: this.lastID });
        });
});

app.put('/api/admin/dpo/consultants/:id', requireRole('admin'), async (req, res) => {
    const { name, email, phone, bio, active } = req.body;
    db.run(`UPDATE dpo_consultants SET name = ?, email = ?, phone = ?, bio = ?, active = ? WHERE id = ?`,
        [name, email || null, phone || null, bio || null, active === false ? 0 : 1, req.params.id], function (err) {
            if (err) return res.status(400).json({ error: 'Erro ao atualizar o consultor.' });
            if (this.changes === 0) return res.status(404).json({ error: 'Consultor não encontrado.' });
            res.json({ message: 'Consultor atualizado!' });
        });
});

app.delete('/api/admin/dpo/consultants/:id', requireRole('admin'), async (req, res) => {
    db.run(`DELETE FROM dpo_consultants WHERE id = ?`, [req.params.id], function (err) {
        if (err) return res.status(400).json({ error: 'Erro ao remover o consultor.' });
        res.json({ message: 'Consultor removido!' });
    });
});

// Master visualiza o conteúdo completo de um pilar (perguntas, verificação,
// explicação de pontos e how to check) sem depender de nenhuma empresa/ciclo —
// só para conferir o layout das perguntas.
app.get('/api/admin/dpo/pillar-layout/:key', requireRole('admin'), (req, res) => {
    const chave = req.params.key;
    if (!DPO_PILARES_ORDEM.includes(chave)) return res.status(404).json({ error: 'Pilar inválido.' });
    res.json({ key: chave, numero: DPO_PILARES_ORDEM.indexOf(chave) + 1, ...DPO_AMBEV_DATA[chave] });
});

// Master calendariza a Auditoria Oficial Ambev (o evento real de auditoria,
// diferente da autoavaliação mensal) para uma empresa específica.
app.put('/api/admin/dpo/companies/:id/auditoria-oficial', requireRole('admin'), async (req, res) => {
    const { data, nota } = req.body;
    try {
        const empresa = await dbGet(`SELECT id FROM companies WHERE id = ?`, [req.params.id]);
        if (!empresa) return res.status(404).json({ error: 'Empresa não encontrada.' });
        await new Promise((resolve, reject) => db.run(
            `UPDATE companies SET dpo_auditoria_oficial_data = ?, dpo_auditoria_oficial_nota = ? WHERE id = ?`,
            [data || null, nota || null, req.params.id], (err) => err ? reject(err) : resolve()
        ));
        res.json({ message: 'Data da Auditoria Oficial Ambev atualizada!' });
    } catch (e) { res.status(400).json({ error: 'Erro ao salvar a data da auditoria oficial.' }); }
});

// Master fecha um valor negociado (diferente do preço de tabela) com uma
// empresa específica, para um pilar avulso ou para a Consultoria Completa, e
// gera na hora o link de pagamento do Mercado Pago para enviar à empresa.
app.post('/api/admin/dpo/negotiated-purchase', requireRole('admin'), async (req, res) => {
    const { companyId, scope, pillarKey, price, consultantId } = req.body;
    if (!companyId) return res.status(400).json({ error: 'Selecione a empresa.' });
    if (!['completo', 'pilar'].includes(scope)) return res.status(400).json({ error: 'Escopo inválido.' });
    if (scope === 'pilar' && !DPO_PILARES_ORDEM.includes(pillarKey)) return res.status(400).json({ error: 'Selecione um pilar válido.' });
    const preco = Number(price);
    if (!preco || preco <= 0) return res.status(400).json({ error: 'Informe o valor negociado.' });
    if (!mpPreference) return res.status(503).json({ error: 'Mercado Pago ainda não foi configurado no servidor.' });
    try {
        const empresa = await dbGet(`SELECT id, name FROM companies WHERE id = ?`, [companyId]);
        if (!empresa) return res.status(404).json({ error: 'Empresa não encontrada.' });
        const ativos = await pilaresAtivosDaEmpresa(companyId);
        if (scope === 'completo' && ativos.length === DPO_PILARES_ORDEM.length) return res.status(400).json({ error: 'Esta empresa já tem a Consultoria Completa liberada.' });
        if (scope === 'pilar' && ativos.includes(pillarKey)) return res.status(400).json({ error: 'Esta empresa já tem este pilar liberado.' });

        const titulo = scope === 'completo'
            ? 'DPO Ambev — Consultoria Completa (valor negociado)'
            : `DPO Ambev — Pilar ${DPO_AMBEV_DATA[pillarKey].label} (valor negociado)`;

        const resultado = await new Promise((resolve, reject) => db.run(
            `INSERT INTO dpo_purchases (company_id, scope, pillar_key, price, consultant_id) VALUES (?, ?, ?, ?, ?)`,
            [companyId, scope, scope === 'pilar' ? pillarKey : null, preco, consultantId || null],
            function (err) { err ? reject(err) : resolve(this.lastID); }
        ));

        const preference = await mpPreference.create({
            body: {
                items: [{ title: titulo, quantity: 1, unit_price: preco, currency_id: 'BRL' }],
                external_reference: `dpoaudit:${resultado}`,
                ...montarRetornoMercadoPago()
            }
        });
        const checkoutUrl = preference.init_point;
        await new Promise((resolve, reject) => db.run(
            `UPDATE dpo_purchases SET mp_preference_id = ?, checkout_url = ? WHERE id = ?`,
            [preference.id, checkoutUrl, resultado], (err) => err ? reject(err) : resolve()
        ));
        notificarGestoresDaEmpresa(companyId, 'DPO Ambev — pagamento negociado', `Um valor negociado de R$ ${preco.toFixed(2)} foi definido — finalize o pagamento no seu painel do DPO Ambev.`);
        res.json({ message: 'Valor negociado criado! Envie o link de pagamento para a empresa.', id: resultado, initPoint: checkoutUrl, companyName: empresa.name });
    } catch (e) {
        const detalhe = detalheErroMercadoPago(e);
        console.error('Erro ao criar compra negociada DPO Ambev:', detalhe);
        res.status(400).json({ error: `Erro ao iniciar o pagamento no Mercado Pago: ${detalhe}` });
    }
});

// Master concede um período de teste (sem cobrança nenhuma no Mercado Pago) —
// libera o pilar (ou a Consultoria Completa) por N dias, escolhidos pelo
// Master. Passado o prazo, o pilar deixa de aparecer como ativo sozinho
// (pilaresAtivosDaEmpresa já filtra pela data), sem precisar de nenhuma ação manual.
app.post('/api/admin/dpo/grant-trial', requireRole('admin'), async (req, res) => {
    const { companyId, scope, pillarKey, days } = req.body;
    if (!companyId) return res.status(400).json({ error: 'Selecione a empresa.' });
    if (!['completo', 'pilar'].includes(scope)) return res.status(400).json({ error: 'Escopo inválido.' });
    if (scope === 'pilar' && !DPO_PILARES_ORDEM.includes(pillarKey)) return res.status(400).json({ error: 'Selecione um pilar válido.' });
    const dias = Number(days);
    if (!dias || dias <= 0) return res.status(400).json({ error: 'Informe quantos dias o período de teste vai durar.' });
    try {
        const empresa = await dbGet(`SELECT id, name FROM companies WHERE id = ?`, [companyId]);
        if (!empresa) return res.status(404).json({ error: 'Empresa não encontrada.' });
        const ativos = await pilaresAtivosDaEmpresa(companyId);
        if (scope === 'completo' && ativos.length === DPO_PILARES_ORDEM.length) return res.status(400).json({ error: 'Esta empresa já tem a Consultoria Completa liberada.' });
        if (scope === 'pilar' && ativos.includes(pillarKey)) return res.status(400).json({ error: 'Esta empresa já tem este pilar liberado.' });

        const expira = new Date(Date.now() + dias * 24 * 60 * 60 * 1000).toISOString();
        const resultado = await new Promise((resolve, reject) => db.run(
            `INSERT INTO dpo_purchases (company_id, scope, pillar_key, price, status, is_trial, trial_expires_at, granted_by, paid_at) VALUES (?, ?, ?, 0, 'paid', 1, ?, ?, CURRENT_TIMESTAMP)`,
            [companyId, scope, scope === 'pilar' ? pillarKey : null, expira, req.user.userId],
            function (err) { err ? reject(err) : resolve(this.lastID); }
        ));
        const rotuloEscopo = scope === 'completo' ? 'a Consultoria Completa (todos os pilares)' : `o pilar ${DPO_AMBEV_DATA[pillarKey].label}`;
        notificarGestoresDaEmpresa(companyId, 'DPO Ambev — período de teste liberado', `O Master liberou ${rotuloEscopo} em período de teste por ${dias} dia(s), até ${new Date(expira).toLocaleDateString('pt-BR')}.`);
        res.json({ message: `Período de teste de ${dias} dia(s) liberado para ${empresa.name}!`, id: resultado, expiresAt: expira });
    } catch (e) { res.status(400).json({ error: 'Erro ao conceder o período de teste.' }); }
});

// Empresa compra um pilar avulso ou a consultoria completa — gera cobrança
// única no Mercado Pago (Checkout Pro), igual à divulgação de vaga.
app.post('/api/dpo/purchase', requireRole('client_admin'), async (req, res) => {
    const { scope, pillarKey, consultantId } = req.body;
    if (!['completo', 'pilar'].includes(scope)) return res.status(400).json({ error: 'Escopo de compra inválido.' });
    if (scope === 'pilar' && !DPO_PILARES_ORDEM.includes(pillarKey)) return res.status(400).json({ error: 'Selecione um pilar válido.' });
    if (!consultantId) return res.status(400).json({ error: 'Escolha o consultor que vai conduzir esta consultoria.' });
    if (!mpPreference) return res.status(503).json({ error: 'Mercado Pago ainda não foi configurado no servidor.' });
    try {
        const consultor = await dbGet(`SELECT id FROM dpo_consultants WHERE id = ? AND active = 1`, [consultantId]);
        if (!consultor) return res.status(400).json({ error: 'Consultor inválido.' });
        const companyId = req.user.companyId;
        const ativos = await pilaresAtivosDaEmpresa(companyId);
        if (scope === 'completo' && ativos.length === DPO_PILARES_ORDEM.length) return res.status(400).json({ error: 'Sua empresa já tem a Consultoria Completa liberada.' });
        if (scope === 'pilar' && ativos.includes(pillarKey)) return res.status(400).json({ error: 'Sua empresa já tem este pilar liberado.' });

        // Já existe uma compra igual aguardando pagamento? Reaproveita a MESMA
        // linha (não cria outra cobrança duplicada para o mesmo pilar/escopo),
        // mas sempre gera um link novo no Mercado Pago em vez de devolver o
        // antigo — um link salvo antes pode ter ficado velho/inválido (ex:
        // apontando para o sandbox de testes em vez do checkout real).
        const pendente = scope === 'completo'
            ? await dbGet(`SELECT * FROM dpo_purchases WHERE company_id = ? AND scope = 'completo' AND status = 'pending_payment'`, [companyId])
            : await dbGet(`SELECT * FROM dpo_purchases WHERE company_id = ? AND scope = 'pilar' AND pillar_key = ? AND status = 'pending_payment'`, [companyId, pillarKey]);
        if (pendente) {
            const tituloPendente = scope === 'completo' ? 'DPO Ambev — Consultoria Completa (todos os pilares)' : `DPO Ambev — Pilar ${DPO_AMBEV_DATA[pillarKey].label}`;
            const preferencePendente = await mpPreference.create({
                body: {
                    items: [{ title: tituloPendente, quantity: 1, unit_price: Number(pendente.price), currency_id: 'BRL' }],
                    external_reference: `dpoaudit:${pendente.id}`,
                    ...montarRetornoMercadoPago()
                }
            });
            const checkoutUrlPendente = preferencePendente.init_point;
            await new Promise((resolve, reject) => db.run(
                `UPDATE dpo_purchases SET mp_preference_id = ?, checkout_url = ?, consultant_id = COALESCE(?, consultant_id) WHERE id = ?`,
                [preferencePendente.id, checkoutUrlPendente, consultantId || null, pendente.id], (err) => err ? reject(err) : resolve()
            ));
            return res.json({ message: 'Você já tem um pagamento pendente para este item — reabrindo o checkout.', id: pendente.id, initPoint: checkoutUrlPendente });
        }

        let preco, titulo;
        if (scope === 'completo') {
            const row = await dbGet(`SELECT value FROM integration_settings WHERE key = 'dpo_full_audit_price'`);
            preco = row ? Number(row.value) : 0;
            titulo = 'DPO Ambev — Consultoria Completa (todos os pilares)';
        } else {
            const row = await dbGet(`SELECT price FROM dpo_pillar_prices WHERE pillar_key = ?`, [pillarKey]);
            preco = row ? Number(row.price) : 0;
            titulo = `DPO Ambev — Pilar ${DPO_AMBEV_DATA[pillarKey].label}`;
        }
        if (!preco) return res.status(400).json({ error: 'Este item ainda não tem um preço definido pelo Master.' });

        const resultado = await new Promise((resolve, reject) => db.run(
            `INSERT INTO dpo_purchases (company_id, scope, pillar_key, price, consultant_id) VALUES (?, ?, ?, ?, ?)`,
            [companyId, scope, scope === 'pilar' ? pillarKey : null, preco, consultantId],
            function (err) { err ? reject(err) : resolve(this.lastID); }
        ));

        const preference = await mpPreference.create({
            body: {
                items: [{ title: titulo, quantity: 1, unit_price: preco, currency_id: 'BRL' }],
                external_reference: `dpoaudit:${resultado}`,
                ...montarRetornoMercadoPago()
            }
        });
        const checkoutUrl = preference.init_point;
        await new Promise((resolve, reject) => db.run(
            `UPDATE dpo_purchases SET mp_preference_id = ?, checkout_url = ? WHERE id = ?`,
            [preference.id, checkoutUrl, resultado], (err) => err ? reject(err) : resolve()
        ));
        res.json({ message: 'Compra criada! Complete o pagamento para liberar.', id: resultado, initPoint: checkoutUrl });
    } catch (e) {
        const detalhe = detalheErroMercadoPago(e);
        console.error('Erro ao criar compra DPO Ambev:', detalhe);
        res.status(400).json({ error: `Erro ao iniciar o pagamento no Mercado Pago: ${detalhe}` });
    }
});

app.post('/api/dpo/confirm-payment', requireRole('admin', 'client_admin'), async (req, res) => {
    const { paymentId } = req.body;
    if (!paymentId) return res.status(400).json({ error: 'Informe o paymentId.' });
    if (!mpPayment) return res.status(503).json({ error: 'Mercado Pago ainda não foi configurado no servidor.' });
    try {
        const pagamento = await mpPayment.get({ id: paymentId });
        const ref = pagamento.external_reference || '';
        if (!ref.startsWith('dpoaudit:')) return res.status(400).json({ error: 'Pagamento não corresponde a uma compra do DPO Ambev.' });
        const compraId = ref.split(':')[1];
        const compra = await dbGet(`SELECT * FROM dpo_purchases WHERE id = ?`, [compraId]);
        if (!compra) return res.status(404).json({ error: 'Compra não encontrada.' });
        if (req.user.role === 'client_admin' && compra.company_id !== req.user.companyId) return res.status(403).json({ error: 'Esta compra não pertence à sua empresa.' });
        if (pagamento.status === 'approved') {
            await new Promise((resolve, reject) => db.run(
                `UPDATE dpo_purchases SET status = 'paid', mp_payment_id = ?, paid_at = CURRENT_TIMESTAMP WHERE id = ?`,
                [paymentId, compraId], (err) => err ? reject(err) : resolve()
            ));
        }
        res.json({ message: pagamento.status === 'approved' ? 'Pagamento confirmado — pilar(es) liberado(s)!' : 'Pagamento ainda não aprovado.', status: pagamento.status });
    } catch (e) {
        console.error('Erro ao confirmar pagamento DPO Ambev:', e.message);
        res.status(400).json({ error: 'Erro ao consultar o pagamento no Mercado Pago.' });
    }
});

app.post('/api/dpo/reopen-checkout/:id', requireRole('client_admin'), async (req, res) => {
    try {
        const compra = await dbGet(`SELECT * FROM dpo_purchases WHERE id = ? AND company_id = ?`, [req.params.id, req.user.companyId]);
        if (!compra) return res.status(404).json({ error: 'Compra não encontrada.' });
        if (compra.status === 'paid') return res.status(400).json({ error: 'Esta compra já foi paga.' });
        if (!mpPreference) return res.status(503).json({ error: 'Mercado Pago ainda não foi configurado no servidor.' });
        const titulo = compra.scope === 'completo' ? 'DPO Ambev — Consultoria Completa (todos os pilares)' : `DPO Ambev — Pilar ${DPO_AMBEV_DATA[compra.pillar_key]?.label || compra.pillar_key}`;
        const preference = await mpPreference.create({
            body: {
                items: [{ title: titulo, quantity: 1, unit_price: Number(compra.price) || 0.01, currency_id: 'BRL' }],
                external_reference: `dpoaudit:${compra.id}`,
                ...montarRetornoMercadoPago()
            }
        });
        const checkoutUrl = preference.init_point;
        await new Promise((resolve, reject) => db.run(`UPDATE dpo_purchases SET mp_preference_id = ?, checkout_url = ? WHERE id = ?`, [preference.id, checkoutUrl, compra.id], (err) => err ? reject(err) : resolve()));
        res.json({ message: 'Novo link de pagamento gerado!', checkoutUrl });
    } catch (e) {
        const detalhe = detalheErroMercadoPago(e);
        res.status(400).json({ error: `Erro ao gerar novo link: ${detalhe}` });
    }
});

// Lista de empresas com o resumo do que já compraram, para o Master escolher
// na hora de criar um novo ciclo de consultoria.
app.get('/api/admin/dpo/companies', requireRole('admin'), async (req, res) => {
    try {
        // ?all=1 traz todas as empresas (usado na precificação negociada e na
        // agenda da Auditoria Oficial); sem o parâmetro, mantém o comportamento
        // original de só listar quem já tem algum pilar liberado (uso nos ciclos).
        const empresas = await dbAll(`SELECT id, name, dpo_auditoria_oficial_data, dpo_auditoria_oficial_nota FROM companies ORDER BY name ASC`);
        const resultado = [];
        for (const emp of empresas) {
            const ativos = await pilaresAtivosDaEmpresa(emp.id);
            if (ativos.length || req.query.all) {
                const trials = await dbAll(`SELECT scope, pillar_key, trial_expires_at FROM dpo_purchases WHERE company_id = ? AND status = 'paid' AND is_trial = 1 AND trial_expires_at > datetime('now') ORDER BY trial_expires_at ASC`, [emp.id]);
                resultado.push({
                    id: emp.id, name: emp.name, pilaresAtivos: ativos,
                    compraCompleta: ativos.length === DPO_PILARES_ORDEM.length,
                    auditoriaOficialData: emp.dpo_auditoria_oficial_data || null,
                    auditoriaOficialNota: emp.dpo_auditoria_oficial_nota || null,
                    trialsAtivos: trials.map(t => ({ scope: t.scope, pillarKey: t.pillar_key, expiraEm: t.trial_expires_at }))
                });
            }
        }
        res.json(resultado);
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar empresas.' }); }
});

// Master cria um novo ciclo (autoavaliação mensal) para a empresa, com os
// pilares que ela já tem liberados (ou um subconjunto, se preferir).
app.post('/api/admin/dpo/cycles', requireRole('admin'), async (req, res) => {
    const { company_id, referencia, pilares, scheduled_at } = req.body;
    if (!company_id) return res.status(400).json({ error: 'Selecione a empresa.' });
    try {
        const ativos = await pilaresAtivosDaEmpresa(company_id);
        if (!ativos.length) return res.status(400).json({ error: 'Esta empresa ainda não comprou nenhum pilar do DPO Ambev.' });
        const escolhidos = (Array.isArray(pilares) && pilares.length) ? pilares.filter(p => ativos.includes(p)) : ativos;
        if (!escolhidos.length) return res.status(400).json({ error: 'Nenhum dos pilares selecionados está liberado para esta empresa.' });

        const resultado = await new Promise((resolve, reject) => db.run(
            `INSERT INTO dpo_audit_cycles (company_id, referencia, pilares, scheduled_at, status, created_by) VALUES (?, ?, ?, ?, 'agendado', ?)`,
            [company_id, referencia || '', JSON.stringify(escolhidos), scheduled_at || null, req.user.userId],
            function (err) { err ? reject(err) : resolve(this.lastID); }
        ));
        notificarPorCompanyAdmins(company_id, 'Nova autoavaliação DPO Ambev', `Um novo ciclo (${referencia || ''}) foi aberto — preencha o checklist dos pilares liberados.`, 'dpoHome');
        res.json({ message: 'Ciclo de autoavaliação criado!', id: resultado });
    } catch (e) { res.status(400).json({ error: 'Erro ao criar o ciclo.' }); }
});

// A empresa acessa e responde o checklist de um pilar A QUALQUER MOMENTO,
// desde que o pilar esteja liberado (pago ou em período de teste) — não
// precisa mais esperar o Master "agendar" nada para isso. O calendário do
// Master (Auditoria Oficial Ambev) serve só para marcar a data da consultoria
// de verdade, não para travar o acesso da empresa ao autoatendimento.
// Esta rota reaproveita um ciclo aberto que já cubra o pilar ou cria um novo
// automaticamente (em 'em_andamento', sem precisar de agendamento prévio).
app.post('/api/dpo/cycles/auto-open', requireRole('client_admin'), async (req, res) => {
    const { pillarKey } = req.body;
    if (!pillarKey || !DPO_PILARES_ORDEM.includes(pillarKey)) return res.status(400).json({ error: 'Pilar inválido.' });
    try {
        const companyId = req.user.companyId;
        const ativos = await pilaresAtivosDaEmpresa(companyId);
        if (!ativos.includes(pillarKey)) return res.status(403).json({ error: 'Sua empresa ainda não tem este pilar liberado.' });

        const abertos = await dbAll(`SELECT * FROM dpo_audit_cycles WHERE company_id = ? AND status != 'concluido' ORDER BY created_at DESC`, [companyId]);
        const existente = abertos.find(c => JSON.parse(c.pilares || '[]').includes(pillarKey));
        if (existente) return res.json({ id: existente.id });

        const meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
        const agora = new Date();
        const referencia = `${meses[agora.getMonth()]}/${agora.getFullYear()}`;

        const resultado = await new Promise((resolve, reject) => db.run(
            `INSERT INTO dpo_audit_cycles (company_id, referencia, pilares, status, created_by) VALUES (?, ?, ?, 'em_andamento', ?)`,
            [companyId, referencia, JSON.stringify([pillarKey]), req.user.userId],
            function (err) { err ? reject(err) : resolve(this.lastID); }
        ));
        res.json({ id: resultado });
    } catch (e) { res.status(400).json({ error: 'Erro ao abrir o pilar.' }); }
});

app.put('/api/admin/dpo/cycles/:id', requireRole('admin'), async (req, res) => {
    const { status, scheduled_at, referencia } = req.body;
    try {
        const ciclo = await dbGet(`SELECT * FROM dpo_audit_cycles WHERE id = ?`, [req.params.id]);
        if (!ciclo) return res.status(404).json({ error: 'Ciclo não encontrado.' });
        const novoStatus = status || ciclo.status;
        await new Promise((resolve, reject) => db.run(
            `UPDATE dpo_audit_cycles SET status = ?, scheduled_at = COALESCE(?, scheduled_at), referencia = COALESCE(?, referencia), closed_at = CASE WHEN ? = 'concluido' THEN CURRENT_TIMESTAMP ELSE closed_at END WHERE id = ?`,
            [novoStatus, scheduled_at || null, referencia || null, novoStatus, req.params.id],
            (err) => err ? reject(err) : resolve()
        ));
        res.json({ message: 'Ciclo atualizado!' });
    } catch (e) { res.status(400).json({ error: 'Erro ao atualizar o ciclo.' }); }
});

// Lista os ciclos — da própria empresa (client_admin) ou, para o Master, de
// uma empresa específica (?company_id=) ou de todas (visão geral).
app.get('/api/dpo/cycles', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const companyId = req.user.role === 'client_admin' ? req.user.companyId : (req.query.company_id || null);
        const ciclos = companyId
            ? await dbAll(`SELECT dac.*, c.name as companyName FROM dpo_audit_cycles dac JOIN companies c ON c.id = dac.company_id WHERE dac.company_id = ? ORDER BY dac.created_at DESC`, [companyId])
            : await dbAll(`SELECT dac.*, c.name as companyName FROM dpo_audit_cycles dac JOIN companies c ON c.id = dac.company_id ORDER BY dac.created_at DESC`);
        const comResumo = [];
        for (const ciclo of ciclos) {
            const pilaresCiclo = JSON.parse(ciclo.pilares || '[]');
            const totalPerguntas = pilaresCiclo.reduce((soma, p) => soma + (DPO_AMBEV_DATA[p] ? DPO_AMBEV_DATA[p].grupos.reduce((s, g) => s + g.perguntas.length, 0) : 0), 0);
            const totalRespondidas = await dbGet(`SELECT COUNT(*) as total FROM dpo_answers WHERE cycle_id = ? AND score IS NOT NULL`, [ciclo.id]);
            comResumo.push({ ...ciclo, pilares: pilaresCiclo, totalPerguntas, totalRespondidas: totalRespondidas.total });
        }
        res.json(comResumo);
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar os ciclos.' }); }
});

app.get('/api/dpo/cycles/:id', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const ciclo = await obterCicloComAcesso(req, res, req.params.id);
        if (!ciclo) return;
        const empresa = await dbGet(`SELECT name FROM companies WHERE id = ?`, [ciclo.company_id]);
        const pilaresCiclo = JSON.parse(ciclo.pilares || '[]');
        const pilares = [];
        for (const chave of pilaresCiclo) {
            const montado = await montarPilarDoCiclo(ciclo.id, chave);
            if (montado) pilares.push(montado);
        }
        res.json({ ...ciclo, companyName: empresa ? empresa.name : '', pilares });
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar o ciclo.' }); }
});

// "Meus Planos DPO" — visão consolidada de TODOS os planos de ação já criados
// pela empresa, agrupados por pilar, juntando ciclos diferentes (histórico
// inteiro), para não depender de entrar ciclo por ciclo pra achar um plano.
app.get('/api/dpo/meus-planos', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const companyId = req.user.role === 'client_admin' ? req.user.companyId : (req.query.company_id || null);
        if (!companyId) return res.status(400).json({ error: 'Informe a empresa (company_id).' });
        const planos = await dbAll(`
            SELECT ap.*, dac.referencia as cicloReferencia, dac.status as cicloStatus, dac.id as cicloId
            FROM dpo_action_plans ap
            JOIN dpo_audit_cycles dac ON dac.id = ap.cycle_id
            WHERE dac.company_id = ?
            ORDER BY ap.question_key ASC, ap.created_at ASC
        `, [companyId]);
        const planoIds = planos.map(p => p.id);
        let followsPorPlano = {};
        if (planoIds.length) {
            const follows = await dbAll(`SELECT * FROM dpo_follow_ups WHERE action_plan_id IN (${planoIds.map(() => '?').join(',')}) ORDER BY numero ASC`, planoIds);
            follows.forEach(f => {
                if (!followsPorPlano[f.action_plan_id]) followsPorPlano[f.action_plan_id] = [];
                followsPorPlano[f.action_plan_id].push(f);
            });
        }
        const porPilar = {};
        planos.forEach(p => {
            const [pilarKey, numeroPergunta] = String(p.question_key).split(':');
            const pilarInfo = DPO_AMBEV_DATA[pilarKey];
            let perguntaTexto = '';
            if (pilarInfo) {
                for (const g of pilarInfo.grupos) {
                    const achou = g.perguntas.find(q => q.numero === numeroPergunta);
                    if (achou) { perguntaTexto = achou.questao; break; }
                }
            }
            if (!porPilar[pilarKey]) porPilar[pilarKey] = { key: pilarKey, label: pilarInfo ? pilarInfo.label : pilarKey, planos: [] };
            porPilar[pilarKey].planos.push({
                ...p,
                perguntaNumero: numeroPergunta,
                perguntaTexto,
                follows: followsPorPlano[p.id] || []
            });
        });
        const ordenado = DPO_PILARES_ORDEM.filter(k => porPilar[k]).map(k => porPilar[k]);
        Object.keys(porPilar).forEach(k => { if (!DPO_PILARES_ORDEM.includes(k)) ordenado.push(porPilar[k]); });
        res.json(ordenado);
    } catch (e) {
        console.error('Erro ao carregar Meus Planos DPO:', e.message);
        res.status(500).json({ error: 'Erro ao carregar os planos de ação.' });
    }
});

app.put('/api/dpo/cycles/:id/answers', requireRole('admin', 'client_admin'), async (req, res) => {
    const { questionKey, score } = req.body;
    if (!questionKey) return res.status(400).json({ error: 'Informe a pergunta.' });
    try {
        const ciclo = await obterCicloComAcesso(req, res, req.params.id);
        if (!ciclo) return;
        if (ciclo.status === 'concluido') return res.status(400).json({ error: 'Este ciclo já foi encerrado e não pode mais ser editado.' });
        const valor = (score === null || score === '') ? null : Number(score);
        await new Promise((resolve, reject) => db.run(
            `INSERT INTO dpo_answers (cycle_id, question_key, score, updated_by) VALUES (?, ?, ?, ?)
             ON CONFLICT(cycle_id, question_key) DO UPDATE SET score = excluded.score, updated_at = CURRENT_TIMESTAMP, updated_by = excluded.updated_by`,
            [req.params.id, questionKey, valor, req.user.userId], (err) => err ? reject(err) : resolve()
        ));
        if (ciclo.status === 'agendado') db.run(`UPDATE dpo_audit_cycles SET status = 'em_andamento' WHERE id = ?`, [req.params.id], () => {});
        res.json({ message: 'Resposta salva!' });
    } catch (e) { res.status(400).json({ error: 'Erro ao salvar a resposta.' }); }
});

const STATUS_PLANO_ACAO_DPO = ['nao_iniciada', 'em_andamento', 'concluida'];

app.post('/api/dpo/action-plans', requireRole('admin', 'client_admin'), async (req, res) => {
    const { cycle_id, questionKey, texto, verificacao_numero, owner, status } = req.body;
    if (!cycle_id || !questionKey || !texto) return res.status(400).json({ error: 'Preencha o plano de ação.' });
    const statusFinal = STATUS_PLANO_ACAO_DPO.includes(status) ? status : 'nao_iniciada';
    try {
        const ciclo = await obterCicloComAcesso(req, res, cycle_id);
        if (!ciclo) return;
        if (ciclo.status === 'concluido') return res.status(400).json({ error: 'Este ciclo já foi encerrado.' });
        const resultado = await new Promise((resolve, reject) => db.run(
            `INSERT INTO dpo_action_plans (cycle_id, question_key, texto, created_by, verificacao_numero, owner, status) VALUES (?, ?, ?, ?, ?, ?, ?)`,
            [cycle_id, questionKey, texto, req.user.userId, verificacao_numero || null, owner || null, statusFinal], function (err) { err ? reject(err) : resolve(this.lastID); }
        ));
        res.json({ message: 'Plano de ação criado!', id: resultado });
    } catch (e) { res.status(400).json({ error: 'Erro ao criar o plano de ação.' }); }
});

app.put('/api/dpo/action-plans/:id', requireRole('admin', 'client_admin'), async (req, res) => {
    const { texto, verificacao_numero, owner, status } = req.body;
    try {
        const plano = await dbGet(`SELECT * FROM dpo_action_plans WHERE id = ?`, [req.params.id]);
        if (!plano) return res.status(404).json({ error: 'Plano de ação não encontrado.' });
        const ciclo = await obterCicloComAcesso(req, res, plano.cycle_id);
        if (!ciclo) return;
        const statusFinal = status !== undefined ? (STATUS_PLANO_ACAO_DPO.includes(status) ? status : plano.status) : plano.status;
        db.run(`UPDATE dpo_action_plans SET texto = ?, verificacao_numero = ?, owner = ?, status = ? WHERE id = ?`,
            [texto !== undefined ? texto : plano.texto, verificacao_numero !== undefined ? verificacao_numero : plano.verificacao_numero, owner !== undefined ? owner : plano.owner, statusFinal, req.params.id],
            () => res.json({ message: 'Plano de ação atualizado!' }));
    } catch (e) { res.status(400).json({ error: 'Erro ao atualizar o plano de ação.' }); }
});

app.delete('/api/dpo/action-plans/:id', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const plano = await dbGet(`SELECT * FROM dpo_action_plans WHERE id = ?`, [req.params.id]);
        if (!plano) return res.status(404).json({ error: 'Plano de ação não encontrado.' });
        const ciclo = await obterCicloComAcesso(req, res, plano.cycle_id);
        if (!ciclo) return;
        db.run(`DELETE FROM dpo_follow_ups WHERE action_plan_id = ?`, [req.params.id], () => {});
        db.run(`DELETE FROM dpo_action_plans WHERE id = ?`, [req.params.id], () => res.json({ message: 'Plano de ação removido!' }));
    } catch (e) { res.status(400).json({ error: 'Erro ao remover o plano de ação.' }); }
});

app.post('/api/dpo/action-plans/:id/follow-ups', requireRole('admin', 'client_admin'), async (req, res) => {
    const { texto, data_prevista } = req.body;
    if (!texto) return res.status(400).json({ error: 'Descreva o follow-up.' });
    try {
        const plano = await dbGet(`SELECT * FROM dpo_action_plans WHERE id = ?`, [req.params.id]);
        if (!plano) return res.status(404).json({ error: 'Plano de ação não encontrado.' });
        const ciclo = await obterCicloComAcesso(req, res, plano.cycle_id);
        if (!ciclo) return;
        const ultimo = await dbGet(`SELECT MAX(numero) as maximo FROM dpo_follow_ups WHERE action_plan_id = ?`, [req.params.id]);
        const numero = (ultimo && ultimo.maximo) ? ultimo.maximo + 1 : 1;
        const resultado = await new Promise((resolve, reject) => db.run(
            `INSERT INTO dpo_follow_ups (action_plan_id, numero, texto, data_prevista) VALUES (?, ?, ?, ?)`,
            [req.params.id, numero, texto, data_prevista || null], function (err) { err ? reject(err) : resolve(this.lastID); }
        ));
        res.json({ message: `Follow ${numero} adicionado!`, id: resultado, numero });
    } catch (e) { res.status(400).json({ error: 'Erro ao adicionar o follow-up.' }); }
});

app.put('/api/dpo/follow-ups/:id', requireRole('admin', 'client_admin'), async (req, res) => {
    const { texto, data_prevista, status } = req.body;
    try {
        const follow = await dbGet(`SELECT * FROM dpo_follow_ups WHERE id = ?`, [req.params.id]);
        if (!follow) return res.status(404).json({ error: 'Follow-up não encontrado.' });
        const plano = await dbGet(`SELECT * FROM dpo_action_plans WHERE id = ?`, [follow.action_plan_id]);
        const ciclo = await obterCicloComAcesso(req, res, plano.cycle_id);
        if (!ciclo) return;
        db.run(`UPDATE dpo_follow_ups SET texto = COALESCE(?, texto), data_prevista = COALESCE(?, data_prevista), status = COALESCE(?, status) WHERE id = ?`,
            [texto || null, data_prevista || null, status || null, req.params.id], () => res.json({ message: 'Follow-up atualizado!' }));
    } catch (e) { res.status(400).json({ error: 'Erro ao atualizar o follow-up.' }); }
});

app.delete('/api/dpo/follow-ups/:id', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const follow = await dbGet(`SELECT * FROM dpo_follow_ups WHERE id = ?`, [req.params.id]);
        if (!follow) return res.status(404).json({ error: 'Follow-up não encontrado.' });
        const plano = await dbGet(`SELECT * FROM dpo_action_plans WHERE id = ?`, [follow.action_plan_id]);
        const ciclo = await obterCicloComAcesso(req, res, plano.cycle_id);
        if (!ciclo) return;
        db.run(`DELETE FROM dpo_follow_ups WHERE id = ?`, [req.params.id], () => res.json({ message: 'Follow-up removido!' }));
    } catch (e) { res.status(400).json({ error: 'Erro ao remover o follow-up.' }); }
});

// Webhook público do Mercado Pago — recebe notificações de mudança de status
// da assinatura (autorizada, pausada, cancelada) e sincroniza com o banco local.
app.post('/api/webhooks/mercadopago', async (req, res) => {
    try {
        const tipo = req.query.type || req.body.type;
        const dataId = (req.query['data.id']) || (req.body.data && req.body.data.id) || req.body.id;
        if (!dataId) return res.sendStatus(200);

        if (tipo === 'preapproval' && mpPreApproval) {
            const dadosAtualizados = await mpPreApproval.get({ id: dataId });
            const empresa = await dbGet(`SELECT id, pending_plan_id FROM companies WHERE mp_preapproval_id = ?`, [dataId]);
            if (dadosAtualizados.status === 'authorized' && empresa && empresa.pending_plan_id) {
                // Pagamento confirmado automaticamente pelo Mercado Pago — libera o
                // plano (e os créditos de cadastro de funcionário) escolhido.
                await new Promise((resolve) => db.run(
                    `UPDATE companies SET subscription_status = ?, subscription_updated_at = CURRENT_TIMESTAMP, plan_id = ?, pending_plan_id = NULL WHERE mp_preapproval_id = ?`,
                    [dadosAtualizados.status, empresa.pending_plan_id, dataId], () => resolve()
                ));
            } else {
                await new Promise((resolve) => db.run(
                    `UPDATE companies SET subscription_status = ?, subscription_updated_at = CURRENT_TIMESTAMP WHERE mp_preapproval_id = ?`,
                    [dadosAtualizados.status, dataId], () => resolve()
                ));
            }
            if (empresa) {
                const rotulos = { authorized: 'Assinatura ativada!', paused: 'Assinatura pausada', cancelled: 'Assinatura cancelada' };
                if (rotulos[dadosAtualizados.status]) notificarGestoresDaEmpresa(empresa.id, rotulos[dadosAtualizados.status], 'Status atualizado automaticamente pelo Mercado Pago.');
            }
        } else if (tipo === 'payment' && mpPayment) {
            // Pagamento único (Checkout Pro) — usado na cobrança por vaga divulgada.
            const pagamento = await mpPayment.get({ id: dataId });
            const ref = pagamento.external_reference || '';
            if (ref.startsWith('jobposting:') && pagamento.status === 'approved') {
                const jobId = ref.split(':')[1];
                const vaga = await dbGet(`SELECT vp.days as planDays FROM job_postings jp LEFT JOIN vaga_plans vp ON vp.id = jp.vaga_plan_id WHERE jp.id = ?`, [jobId]);
                const dias = (vaga && vaga.planDays) || 30;
                await new Promise((resolve) => db.run(
                    `UPDATE job_postings SET status = 'active', mp_payment_id = ?, published_at = CURRENT_TIMESTAMP, expires_at = datetime(CURRENT_TIMESTAMP, '+${Number(dias)} days') WHERE id = ?`,
                    [dataId, jobId], () => resolve()
                ));
            } else if (ref.startsWith('closingfee:') && pagamento.status === 'approved') {
                const jobId = ref.split(':')[1];
                const vaga = await dbGet(`SELECT company_id, title FROM job_postings WHERE id = ?`, [jobId]);
                await new Promise((resolve) => db.run(
                    `UPDATE job_postings SET closing_fee_status = 'paid', closing_fee_payment_id = ?, closing_fee_paid_at = CURRENT_TIMESTAMP WHERE id = ?`,
                    [dataId, jobId], () => resolve()
                ));
                if (vaga) notificarGestoresDaEmpresa(vaga.company_id, 'Taxa de fechamento paga', `Pagamento da taxa de fechamento da vaga "${vaga.title}" confirmado.`);
            } else if (ref.startsWith('dpoaudit:') && pagamento.status === 'approved') {
                const compraId = ref.split(':')[1];
                const compra = await dbGet(`SELECT * FROM dpo_purchases WHERE id = ?`, [compraId]);
                if (compra && compra.status !== 'paid') {
                    await new Promise((resolve) => db.run(
                        `UPDATE dpo_purchases SET status = 'paid', mp_payment_id = ?, paid_at = CURRENT_TIMESTAMP WHERE id = ?`,
                        [dataId, compraId], () => resolve()
                    ));
                    notificarGestoresDaEmpresa(compra.company_id, 'DPO Ambev liberado', 'Pagamento confirmado — o pilar/consultoria já está disponível para autoavaliação.');
                }
            }
        }
        res.sendStatus(200);
    } catch (e) {
        console.warn('⚠️  Erro ao processar webhook do Mercado Pago:', e.message);
        res.sendStatus(200); // sempre 200, senão o Mercado Pago fica reenviando indefinidamente
    }
});

// ============================================================
// DASHBOARD / RELATÓRIOS DE ENGAJAMENTO
// ============================================================

app.get('/api/engagement-dashboard', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const companyId = req.user.role === 'client_admin' ? req.user.companyId : null;
        const filtroEmployee = companyId ? ` AND e.company_id = ${Number(companyId)}` : '';
        const filtroUser = companyId ? ` AND u.company_id = ${Number(companyId)}` : '';

        const totalExecutivos = await dbGet(`SELECT COUNT(*) as total FROM employees e WHERE 1=1${filtroEmployee}`);
        const progressoMedio = await dbGet(`SELECT AVG(progress_percentage) as media FROM employees e WHERE 1=1${filtroEmployee}`);

        const pdiPorStatus = await dbAll(`SELECT p.status, COUNT(*) as total FROM pd_plans p JOIN employees e ON p.employee_id = e.id WHERE 1=1${filtroEmployee} GROUP BY p.status`);
        const pdiTotal = pdiPorStatus.reduce((soma, r) => soma + r.total, 0);
        const pdiConcluidos = (pdiPorStatus.find(r => r.status === 'Concluído') || { total: 0 }).total;

        const mentoriasPorStatus = await dbAll(`SELECT m.status, COUNT(*) as total FROM mentorships m JOIN employees e ON m.employee_id = e.id WHERE 1=1${filtroEmployee} GROUP BY m.status`);
        const mentoriasTotal = mentoriasPorStatus.reduce((soma, r) => soma + r.total, 0);
        const mentoriasRealizadas = (mentoriasPorStatus.find(r => r.status === 'Realizada') || { total: 0 }).total;

        const totalEventos = await dbGet(`SELECT COUNT(*) as total FROM events`);
        const totalInscricoes = await dbGet(`SELECT COUNT(*) as total FROM event_registrations er JOIN users u ON er.user_id = u.id WHERE 1=1${filtroUser}`);

        const usuariosAtivos = await dbGet(
            `SELECT COUNT(DISTINCT pl.user_id) as total FROM points_ledger pl JOIN users u ON pl.user_id = u.id
             WHERE pl.created_at >= datetime('now','-30 days')${filtroUser}`
        );

        const trilhas = await dbAll(`
            SELECT t.id, t.title,
                (SELECT COUNT(*) FROM track_lessons tl WHERE tl.track_id = t.id) as totalAulas,
                (SELECT COUNT(DISTINCT tp.user_id) FROM track_progress tp JOIN users u ON tp.user_id = u.id WHERE tp.track_id = t.id${filtroUser}) as alunosEngajados,
                (SELECT COUNT(*) FROM certificates c JOIN users u ON c.user_id = u.id WHERE c.track_id = t.id${filtroUser}) as certificadosEmitidos
            FROM learning_tracks t ORDER BY t.title ASC
        `);

        const topRanking = await dbAll(`
            SELECT u.name, u.role, COALESCE(SUM(pl.points), 0) as totalPontos
            FROM users u LEFT JOIN points_ledger pl ON pl.user_id = u.id
            WHERE u.role IN ('autonomous', 'client_admin', 'mentor')${filtroUser}
            GROUP BY u.id ORDER BY totalPontos DESC LIMIT 5
        `);

        res.json({
            totalExecutivos: totalExecutivos.total,
            progressoMedio: Math.round(progressoMedio.media || 0),
            pdi: { total: pdiTotal, concluidos: pdiConcluidos, taxaConclusao: pdiTotal ? Math.round((pdiConcluidos / pdiTotal) * 100) : 0 },
            mentorias: { total: mentoriasTotal, realizadas: mentoriasRealizadas },
            eventos: { totalEventos: totalEventos.total, totalInscricoes: totalInscricoes.total },
            usuariosAtivos30dias: usuariosAtivos.total,
            trilhas,
            topRanking
        });
    } catch (e) {
        console.error('Erro ao montar dashboard de engajamento:', e.message);
        res.status(500).json({ error: 'Erro ao carregar o painel de engajamento.' });
    }
});

// ============================================================
// NOTIFICAÇÕES IN-APP (SININHO)
// ============================================================

app.get('/api/notifications', async (req, res) => {
    try {
        const lista = await dbAll(`SELECT * FROM notifications WHERE user_id = ? ORDER BY created_at DESC LIMIT 30`, [req.user.userId]);
        res.json(lista);
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar notificações.' }); }
});

app.get('/api/notifications/unread-count', async (req, res) => {
    try {
        const r = await dbGet(`SELECT COUNT(*) as total FROM notifications WHERE user_id = ? AND is_read = 0`, [req.user.userId]);
        res.json({ total: r.total });
    } catch (e) { res.status(500).json({ error: 'Erro ao contar notificações.' }); }
});

app.put('/api/notifications/:id/read', (req, res) => {
    db.run(`UPDATE notifications SET is_read = 1 WHERE id = ? AND user_id = ?`, [req.params.id, req.user.userId], (err) => {
        if (err) return res.status(400).json({ error: err.message });
        res.json({ message: 'Marcada como lida.' });
    });
});

app.put('/api/notifications/read-all', (req, res) => {
    db.run(`UPDATE notifications SET is_read = 1 WHERE user_id = ? AND is_read = 0`, [req.user.userId], (err) => {
        if (err) return res.status(400).json({ error: err.message });
        res.json({ message: 'Todas marcadas como lidas.' });
    });
});

// ============================================================
// GAMIFICAÇÃO: RANKING DE PONTOS
// ============================================================

app.get('/api/ranking', async (req, res) => {
    try {
        const ranking = await dbAll(
            `SELECT u.id as userId, u.name, u.role, COALESCE(SUM(pl.points), 0) as totalPontos
             FROM users u LEFT JOIN points_ledger pl ON pl.user_id = u.id
             WHERE u.role IN ('autonomous', 'client_admin', 'mentor')
             GROUP BY u.id ORDER BY totalPontos DESC LIMIT 20`
        );
        const meuTotal = await dbGet(`SELECT COALESCE(SUM(points), 0) as total FROM points_ledger WHERE user_id = ?`, [req.user.userId]);
        res.json({ ranking, meuTotal: meuTotal.total });
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar ranking.' }); }
});

// Serve os uploads a partir de PASTA_UPLOADS (que pode estar fora da pasta do
// código, se UPLOADS_PATH estiver definido) — tem que vir ANTES do estático
// genérico de /public pra funcionar mesmo quando os dois caminhos são diferentes.
app.use('/uploads', express.static(PASTA_UPLOADS));
app.use(express.static(path.join(__dirname, 'public')));

// ============================================================
// VIDEOCHAMADA PRÓPRIA (WEBRTC) PARA AS MENTORIAS
// O servidor HTTP normal do Express passa a rodar dentro de um servidor HTTP
// explícito para que o Socket.IO possa se acoplar a ele (sinalização
// offer/answer/ICE candidate). Nenhum vídeo/áudio passa pelo servidor — ele
// só troca as mensagens de sinalização entre os dois participantes; a
// chamada em si é peer-to-peer (WebRTC), via STUN público do Google.
// ============================================================
const httpServer = http.createServer(app);
const io = new SocketIOServer(httpServer, { cors: { origin: '*' } });

// mentorshipRoomId -> Map(socketId -> { userId, name, role, convidado })
const SALAS_VIDEOCHAMADA = new Map();
// Limite só para não deixar a sala crescer sem controle — não é uma promessa de
// que 1000 câmeras simultâneas rodam bem: WebRTC em malha (cada participante
// conectado a todos os outros) funciona bem até uns 10-15 com vídeo ligado.
// Para uma audiência grande de verdade (dezenas/centenas assistindo), o jeito
// certo é os extras entrarem só de áudio/vídeo desligado, ou migrar para um
// servidor de mídia (SFU) — fica registrado aqui para uma próxima etapa.
const CAPACIDADE_MAXIMA_SALA = 1000;

function entrarNaSalaVideochamada(socket, sala, participante) {
    if (!SALAS_VIDEOCHAMADA.has(sala)) SALAS_VIDEOCHAMADA.set(sala, new Map());
    const participantes = SALAS_VIDEOCHAMADA.get(sala);
    if (participantes.size >= CAPACIDADE_MAXIMA_SALA) {
        socket.emit('sala-entrada-negada', { error: 'Esta sala já atingiu o limite de participantes.' });
        return;
    }
    const outros = Array.from(participantes.entries()).map(([id, p]) => ({ socketId: id, name: p.name }));
    participantes.set(socket.id, participante);
    socket.join(sala);
    socket.data.sala = sala;
    // O recém-chegado recebe a lista de quem já está na sala e é ele quem
    // cria a oferta WebRTC para CADA um deles (evita duas ofertas cruzadas
    // no mesmo par e permite qualquer número de participantes, não só 2).
    socket.emit('entrou-na-sala', { outros });
    socket.to(sala).emit('usuario-entrou', { socketId: socket.id, name: participante.name });
}

io.on('connection', (socket) => {
    socket.on('entrar-sala-mentoria', async ({ mentorshipId, token }) => {
        try {
            const payload = jwt.verify(token, JWT_SECRET);
            const mentoria = await dbGet(`SELECT * FROM mentorships WHERE id = ?`, [mentorshipId]);
            if (!mentoria) return socket.emit('sala-entrada-negada', { error: 'Mentoria não encontrada.' });

            const permitido =
                payload.role === 'admin' ||
                (payload.role === 'autonomous' && payload.employeeId === mentoria.employee_id) ||
                (payload.role === 'mentor' && payload.mentorId === mentoria.mentor_id) ||
                (payload.role === 'client_admin' && await dbGet(`SELECT id FROM employees WHERE id = ? AND company_id = ?`, [mentoria.employee_id, payload.companyId]));
            if (!permitido) return socket.emit('sala-entrada-negada', { error: 'Você não faz parte desta mentoria.' });

            const usuario = await dbGet(`SELECT name FROM users WHERE id = ?`, [payload.userId]);
            entrarNaSalaVideochamada(socket, `mentoria-${mentorshipId}`, { userId: payload.userId, name: usuario ? usuario.name : 'Participante', role: payload.role, convidado: false });
        } catch (e) {
            socket.emit('sala-entrada-negada', { error: 'Token inválido ou expirado.' });
        }
    });

    // Convidado externo entrando pelo link público (sem login) — só precisa do
    // código da sala (room_token, aleatório e não sequencial) e do nome dele.
    socket.on('entrar-sala-mentoria-convidado', async ({ roomToken, nome }) => {
        try {
            if (!roomToken || !nome) return socket.emit('sala-entrada-negada', { error: 'Informe seu nome.' });
            const mentoria = await dbGet(`SELECT * FROM mentorships WHERE room_token = ?`, [roomToken]);
            if (!mentoria) return socket.emit('sala-entrada-negada', { error: 'Link inválido ou expirado.' });
            entrarNaSalaVideochamada(socket, `mentoria-${mentoria.id}`, { userId: null, name: String(nome).slice(0, 60), role: 'convidado', convidado: true });
        } catch (e) {
            socket.emit('sala-entrada-negada', { error: 'Não foi possível entrar na sala.' });
        }
    });

    // Sinalização WebRTC dirigida a UM participante específico (necessário com
    // mais de 2 pessoas na sala — antes ia para a sala inteira, o que só
    // funciona quando só existem 2 participantes).
    socket.on('sinal-webrtc', ({ tipo, dados, paraSocketId }) => {
        const sala = socket.data.sala;
        if (!sala || !paraSocketId) return;
        io.to(paraSocketId).emit('sinal-webrtc', { tipo, dados, de: socket.id });
    });

    // Chat de texto dentro da videochamada — só repassa a mensagem para os
    // outros participantes da mesma sala, não fica gravado no banco.
    socket.on('mensagem-sala-mentoria', ({ texto }) => {
        const sala = socket.data.sala;
        if (!sala || !texto) return;
        const participantes = SALAS_VIDEOCHAMADA.get(sala);
        const eu = participantes && participantes.get(socket.id);
        socket.to(sala).emit('mensagem-sala-mentoria', { texto: String(texto).slice(0, 1000), nome: eu ? eu.name : 'Participante', de: socket.id });
    });

    // Avisa os outros participantes se a câmera de alguém foi ligada/desligada,
    // pra eles poderem trocar o vídeo congelado/preto por um aviso amigável
    // ("câmera desligada") em vez de ficar com um quadro estranho.
    socket.on('camera-estado', ({ ligada }) => {
        const sala = socket.data.sala;
        if (!sala) return;
        socket.to(sala).emit('camera-estado', { socketId: socket.id, ligada: !!ligada });
    });

    socket.on('sair-sala-mentoria', () => {
        const sala = socket.data.sala;
        if (sala && SALAS_VIDEOCHAMADA.has(sala)) {
            SALAS_VIDEOCHAMADA.get(sala).delete(socket.id);
            socket.to(sala).emit('usuario-saiu', { socketId: socket.id });
            socket.leave(sala);
        }
    });

    socket.on('disconnect', () => {
        const sala = socket.data.sala;
        if (sala && SALAS_VIDEOCHAMADA.has(sala)) {
            SALAS_VIDEOCHAMADA.get(sala).delete(socket.id);
            socket.to(sala).emit('usuario-saiu', { socketId: socket.id });
        }
    });
});

// Handler de erro global: garante que qualquer exceção não tratada em uma
// rota (síncrona, ou uma Promise rejeitada chegando aqui via next(err))
// sempre responda em JSON com um campo "error" — nunca a página HTML padrão
// do Express, que o front-end não consegue interpretar (e aparecia como
// "Erro na requisição" genérico sem explicar o motivo real).
app.use((err, req, res, next) => {
    console.error('Erro não tratado:', err);
    if (res.headersSent) return next(err);
    res.status(500).json({ error: 'Erro interno no servidor. Tente novamente ou verifique os logs.' });
});

// Evita que uma rejeição de Promise não tratada (ex.: um await sem try/catch
// em algum ponto) derrube o servidor inteiro — só registra no log.
process.on('unhandledRejection', (reason) => {
    console.error('Promise rejeitada sem tratamento:', reason);
});

httpServer.listen(PORT, () => console.log(`🚀 Plataforma Completa Impulsionar V4 em http://127.0.0.1:${PORT}`));