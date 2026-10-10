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
const ExcelJS = require('exceljs');

// Base de perguntas do checklist "DPO AMBEV" (consultoria por pilares) — vem de
// uma planilha modelo da Ambev e é conteúdo estático (não muda pela tela),
// então fica num JSON à parte em vez de virar uma tabela gigante no banco.
const DPO_AMBEV_DATA = require('./dpo_ambev_data.json');
const DISC_DATA = require('./disc_data.json');
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

// ---------- Envio de e-mail ----------
// O Railway bloqueia SMTP (portas 25/465/587) nos planos Free/Trial/Hobby — a conexão fica
// "pendurada" até estourar o tempo. Por isso o envio pode ir por API HTTPS (Brevo ou Resend),
// configurada na tela Meu Perfil do Master; o SMTP do .env fica como alternativa.
const SMTP_PORTA = Number(process.env.SMTP_PORT) || 2525;
const smtpTransporter = nodemailer.createTransport({
    host: process.env.SMTP_HOST || 'smtp.mailtrap.io',
    port: SMTP_PORTA,
    secure: SMTP_PORTA === 465,
    connectionTimeout: 15000, greetingTimeout: 15000, socketTimeout: 25000,
    auth: {
        user: process.env.SMTP_USER || 'seu_usuario_smtp',
        pass: process.env.SMTP_PASS || 'sua_senha_smtp'
    }
});
let EMAIL_API = { provedor: process.env.BREVO_API_KEY ? 'brevo' : process.env.RESEND_API_KEY ? 'resend' : '', chave: process.env.BREVO_API_KEY || process.env.RESEND_API_KEY || '', remetente: process.env.EMAIL_FROM || '' };
function separarRemetenteEmail(txt) {
    const t = String(txt || '').trim(), m = t.match(/^\s*"?([^"<]*)"?\s*<([^>]+)>\s*$/);
    return m ? { nome: m[1].trim(), email: m[2].trim() } : { nome: '', email: t };
}
function listaEmails(v) { return (Array.isArray(v) ? v : String(v || '').split(',')).map(x => separarRemetenteEmail(x).email).filter(Boolean); }
function anexosEmailApi(opts) {
    const l = [];
    (opts.attachments || []).forEach(a => {
        let conteudo = a.content;
        if (conteudo === undefined && a.path) { try { conteudo = fs.readFileSync(a.path); } catch (e) { return; } }
        if (conteudo === undefined) return;
        l.push({ nome: a.filename || 'anexo', base64: Buffer.isBuffer(conteudo) ? conteudo.toString('base64') : (a.encoding === 'base64' ? String(conteudo) : Buffer.from(String(conteudo)).toString('base64')) });
    });
    if (opts.icalEvent && opts.icalEvent.content) l.push({ nome: opts.icalEvent.filename || 'convite.ics', base64: Buffer.from(String(opts.icalEvent.content)).toString('base64') });
    return l;
}
async function enviarEmailPorApi(opts) {
    const rem = separarRemetenteEmail(EMAIL_API.remetente || opts.from || '');
    const doOpts = separarRemetenteEmail(opts.from || '');
    const nome = rem.nome || doOpts.nome || 'Impulsionar V4';
    const para = listaEmails(opts.to), cc = listaEmails(opts.cc), bcc = listaEmails(opts.bcc);
    if (!para.length) throw new Error('Destinatário vazio.');
    if (!rem.email) throw new Error('Configure o e-mail remetente do envio por API.');
    const anexos = anexosEmailApi(opts);
    const ctrl = new AbortController(), timer = setTimeout(() => ctrl.abort(), 20000);
    try {
        let r;
        if (EMAIL_API.provedor === 'resend') {
            r = await fetch('https://api.resend.com/emails', { method: 'POST', signal: ctrl.signal, headers: { Authorization: 'Bearer ' + EMAIL_API.chave, 'Content-Type': 'application/json' },
                body: JSON.stringify({ from: `${nome} <${rem.email}>`, to: para, cc: cc.length ? cc : undefined, bcc: bcc.length ? bcc : undefined, reply_to: opts.replyTo || undefined, subject: opts.subject || '', html: opts.html || undefined, text: opts.text || undefined, attachments: anexos.length ? anexos.map(a => ({ filename: a.nome, content: a.base64 })) : undefined }) });
        } else {
            r = await fetch('https://api.brevo.com/v3/smtp/email', { method: 'POST', signal: ctrl.signal, headers: { 'api-key': EMAIL_API.chave, 'Content-Type': 'application/json', Accept: 'application/json' },
                body: JSON.stringify({ sender: { name: nome, email: rem.email }, to: para.map(email => ({ email })), cc: cc.length ? cc.map(email => ({ email })) : undefined, bcc: bcc.length ? bcc.map(email => ({ email })) : undefined, replyTo: opts.replyTo ? { email: separarRemetenteEmail(opts.replyTo).email } : undefined, subject: opts.subject || '(sem assunto)', htmlContent: opts.html || undefined, textContent: opts.html ? undefined : (opts.text || ' '), attachment: anexos.length ? anexos.map(a => ({ name: a.nome, content: a.base64 })) : undefined }) });
        }
        const j = await r.json().catch(() => ({}));
        if (!r.ok) throw new Error((j && (j.message || j.error || j.code)) ? `${EMAIL_API.provedor === 'resend' ? 'Resend' : 'Brevo'}: ${j.message || j.error || j.code}` : `Erro ${r.status} no envio por API.`);
        return { messageId: j.messageId || j.id || '', accepted: para };
    } catch (e) {
        if (e.name === 'AbortError') throw new Error('O serviço de e-mail não respondeu em 20 segundos.');
        throw e;
    } finally { clearTimeout(timer); }
}
function explicarErroSmtp(e) {
    const m = String(e && e.message || e);
    if (/ETIMEDOUT|timeout|Greeting never received|Connection timeout|ECONNREFUSED|ENETUNREACH/i.test(m)) return new Error('O servidor não conseguiu conectar no SMTP (tempo esgotado). No Railway, o SMTP é bloqueado nos planos Free/Trial/Hobby — configure o envio por API (Brevo ou Resend) em Meu Perfil. Detalhe: ' + m);
    if (/Invalid login|535|Username and Password not accepted|EAUTH/i.test(m)) return new Error('Usuário ou senha do SMTP recusados. No Gmail é preciso usar uma "senha de app" (não a senha normal). Detalhe: ' + m);
    return e instanceof Error ? e : new Error(m);
}
// Mesmo formato do nodemailer (promise ou callback), para não mexer em quem já usa.
const transporter = {
    sendMail(opts, cb) {
        const p = (EMAIL_API.provedor && EMAIL_API.chave ? enviarEmailPorApi(opts) : smtpTransporter.sendMail(opts).catch(e => { throw explicarErroSmtp(e); }));
        if (typeof cb === 'function') { p.then(info => cb(null, info), err => cb(err)); return undefined; }
        return p;
    }
};

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
let ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY || '';
let ANTHROPIC_MODEL = process.env.ANTHROPIC_MODEL || 'claude-haiku-4-5-20251001';
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
         'dpo_auditoria_oficial_data TEXT', 'dpo_auditoria_oficial_nota TEXT',
         // Revenda que nunca foi auditada (1ª auditoria em 2026) — só ela pode
         // receber o selo "Route Basic" na régua de selos DPO 2026.
         'dpo_primeira_auditoria INTEGER DEFAULT 0',
         // Gestor principal da empresa (marcado pelo Master) — aprova cadastros
         // sensíveis como os aprovadores de CAPEX; usado em outras rotinas.
         'gestor_principal_id INTEGER'].forEach(coluna => {
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
        // Acesso criado pela própria empresa: entra sem permissões e só loga depois que o Master aprovar.
        db.run(`ALTER TABLE users ADD COLUMN aprovacao_pendente INTEGER DEFAULT 0`, () => {});
        db.run(`ALTER TABLE users ADD COLUMN ultimo_login TEXT`, () => {});
        db.run(`ALTER TABLE users ADD COLUMN qtd_logins INTEGER DEFAULT 0`, () => {});
        db.run(`ALTER TABLE users ADD COLUMN criado_em TEXT`, () => {});
        db.run(`CREATE TABLE IF NOT EXISTS vaga_creditos_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT, company_id INTEGER NOT NULL, dias INTEGER NOT NULL, motivo TEXT,
            tipo TEXT, user_id INTEGER, job_posting_id INTEGER, created_at DATETIME DEFAULT CURRENT_TIMESTAMP)`);
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
            ai_resume_review_autopilot: '0',  // IA decide aprovação/ajuste de currículo sozinha
            portal_mostrar_numeros: '0'       // mostra aos candidatos o total de vagas e empresas (ligar quando o portal tiver volume)
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
        // Chave/modelo da IA salvos pelo Master (Ministrar Treinamento > Configurações) têm prioridade sobre o .env.
        db.all(`SELECT key, value FROM integration_settings WHERE key IN ('anthropic_api_key', 'anthropic_model')`, [], (err, rows) => {
            if (err || !rows) return;
            rows.forEach(r => { if (r.key === 'anthropic_api_key' && r.value) ANTHROPIC_API_KEY = r.value; if (r.key === 'anthropic_model' && r.value) ANTHROPIC_MODEL = r.value; });
        });
        db.all(`SELECT key, value FROM integration_settings WHERE key IN ('email_api_provedor', 'email_api_chave', 'email_api_remetente')`, [], (err, rows) => {
            if (err || !rows || !rows.length) return;
            const m = Object.fromEntries(rows.map(r => [r.key, r.value]));
            if (m.email_api_provedor && m.email_api_chave) EMAIL_API = { provedor: m.email_api_provedor, chave: m.email_api_chave, remetente: m.email_api_remetente || '' };
        });
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

        // Resultados do Teste de Perfil DISC (formato "ranking por pergunta"):
        // cada linha é uma tentativa completa do próprio executivo/colaborador
        // (role 'autonomous'), com as respostas brutas (ordem escolhida em cada
        // pergunta) e os pontos já calculados por dimensão D/I/S/C.
        db.run(`CREATE TABLE IF NOT EXISTS disc_results (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id INTEGER NOT NULL,
            respostas TEXT,
            pontos_d INTEGER DEFAULT 0,
            pontos_i INTEGER DEFAULT 0,
            pontos_s INTEGER DEFAULT 0,
            pontos_c INTEGER DEFAULT 0,
            perfil_primario TEXT,
            perfil_secundario TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(employee_id) REFERENCES employees(id)
        )`);

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
        // photo_url: foto de capa do PDI. final_delivery: "o que vamos entregar
        // no final desse PDI" — o resultado esperado, combinado desde o início.
        ['photo_url TEXT', 'final_delivery TEXT'].forEach(coluna => {
            db.run(`ALTER TABLE pd_plans ADD COLUMN ${coluna}`, () => {});
        });
        // foto de evidência de cada ação do PDI (ex: print/foto do que foi feito).
        db.run(`ALTER TABLE pdi_actions ADD COLUMN foto_url TEXT`, () => {});

        // Linha do tempo de evolução do PDI: vários follow-ups ao longo do
        // acompanhamento (não um status único), cada um podendo ter uma foto —
        // é o que dá o histórico de "andamento e evolução do cliente" pedido.
        db.run(`CREATE TABLE IF NOT EXISTS pdi_updates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            pd_plan_id INTEGER NOT NULL,
            texto TEXT NOT NULL,
            foto_url TEXT,
            created_by_name TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(pd_plan_id) REFERENCES pd_plans(id)
        )`);

        // Materiais de apoio do PDI: arquivos, livros ou links indicados ao
        // executivo para ajudar de verdade no desenvolvimento dele.
        db.run(`CREATE TABLE IF NOT EXISTS pdi_materials (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            pd_plan_id INTEGER NOT NULL,
            titulo TEXT NOT NULL,
            tipo TEXT DEFAULT 'arquivo',
            url TEXT,
            nota TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(pd_plan_id) REFERENCES pd_plans(id)
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
         'experiences_json TEXT', 'desired_states TEXT', 'desired_cities TEXT', 'cep TEXT', 'neighborhood TEXT', 'state TEXT', 'lgpd_at TEXT', 'origem TEXT',
         'birth_date TEXT', 'cnh TEXT', 'pretensao_salarial TEXT', 'modalidade TEXT', 'disp_viagem INTEGER DEFAULT 0', 'disp_mudanca INTEGER DEFAULT 0', 'pcd TEXT',
         'education_json TEXT', 'courses_json TEXT', 'disponibilidade_inicio TEXT', 'curriculo_completo_em TEXT',
         'tipos_vaga TEXT', 'disc_liberado INTEGER DEFAULT 0', 'disc_liberado_em TEXT', 'disc_perfil TEXT', 'disc_json TEXT', 'disc_em TEXT'].forEach(coluna => {
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
        ['work_schedule TEXT', 'pcd INTEGER DEFAULT 0', 'contract_type TEXT', 'categoria TEXT'].forEach(c => db.run(`ALTER TABLE job_postings ADD COLUMN ${c}`, () => {}));
        // Banco de currículos da empresa: candidatos do portal guardados por função para próximas vagas.
        db.run(`CREATE TABLE IF NOT EXISTS company_talent_bank (
            id INTEGER PRIMARY KEY AUTOINCREMENT, company_id INTEGER NOT NULL, candidate_user_id INTEGER NOT NULL, funcao TEXT, observacao TEXT,
            origem_vaga_id INTEGER, user_id INTEGER, created_at DATETIME DEFAULT CURRENT_TIMESTAMP, UNIQUE(company_id, candidate_user_id))`);
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

        ['status TEXT DEFAULT \'recebida\'', 'status_em TEXT', 'entrevista_em TEXT', 'entrevista_local TEXT', 'entrevista_resposta TEXT', 'visto_empresa_em TEXT'].forEach(c => db.run(`ALTER TABLE job_applications ADD COLUMN ${c}`, () => {}));
        // Linha do tempo de cada candidatura: entrevista, recusa, mensagens da
        // empresa e respostas do candidato — tudo aparece para os dois lados.
        db.run(`CREATE TABLE IF NOT EXISTS application_events (
            id INTEGER PRIMARY KEY AUTOINCREMENT, application_id INTEGER NOT NULL, autor TEXT NOT NULL, tipo TEXT NOT NULL,
            texto TEXT, user_id INTEGER, lido_candidato INTEGER DEFAULT 0, lido_empresa INTEGER DEFAULT 0, created_at DATETIME DEFAULT CURRENT_TIMESTAMP)`);

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

        // Agenda de consultorias (data + hora início/fim) — gera convite no
        // Outlook/Teams do consultor e dos usuários da empresa (pelo e-mail).
        db.run(`CREATE TABLE IF NOT EXISTS dpo_sessoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER NOT NULL,
            consultant_id INTEGER,
            titulo TEXT NOT NULL,
            pilares TEXT,
            data TEXT NOT NULL,
            hora_inicio TEXT NOT NULL,
            hora_fim TEXT NOT NULL,
            formato TEXT DEFAULT 'teams',
            local TEXT,
            teams_link TEXT,
            participantes TEXT,
            observacao TEXT,
            status TEXT DEFAULT 'agendada',
            uid TEXT,
            sequencia INTEGER DEFAULT 0,
            graph_event_id TEXT,
            ultimo_envio TEXT,
            created_by INTEGER,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            updated_at DATETIME
        )`);

        // Empresa só SOLICITA a consultoria; quem agenda é o Master.
        db.run(`CREATE TABLE IF NOT EXISTS dpo_sessao_solicitacoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER NOT NULL,
            user_id INTEGER,
            pilares TEXT,
            data_sugerida TEXT,
            hora_sugerida TEXT,
            hora_fim_sugerida TEXT,
            formato TEXT,
            assunto TEXT,
            status TEXT DEFAULT 'pendente',
            resposta TEXT,
            sessao_id INTEGER,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            updated_at DATETIME
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
        // "data_prevista" é o prazo do PLANO em si (escolhido já na criação, sem
        // precisar abrir um follow-up separado só para ter uma data) — diferente
        // do data_prevista de cada follow-up individual, que continua existindo.
        ['verificacao_numero TEXT', 'owner TEXT', "status TEXT DEFAULT 'nao_iniciada'", 'data_prevista TEXT'].forEach(coluna => {
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
        // Foto de evidência de cada follow-up (ex: print/foto do que foi feito).
        db.run(`ALTER TABLE dpo_follow_ups ADD COLUMN foto_url TEXT`, () => {});
        db.run(`ALTER TABLE dpo_follow_ups ADD COLUMN autor TEXT`, () => {});

        // "Perguntas Bate-Papo" — pasta dentro de cada pilar onde a empresa (ou o
        // Master/consultor) registra perguntas feitas no bate-papo e as respostas,
        // sempre vinculadas a uma pergunta oficial do pilar (question_numero, ex.:
        // "1.1"). Pertence à EMPRESA + PILAR (não ao ciclo), então o histórico fica
        // guardado de um mês para o outro.
        db.run(`CREATE TABLE IF NOT EXISTS dpo_chat_questions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER NOT NULL,
            pillar_key TEXT NOT NULL,
            question_numero TEXT NOT NULL,
            pergunta TEXT NOT NULL,
            resposta TEXT,
            created_by INTEGER,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            updated_at DATETIME,
            FOREIGN KEY(company_id) REFERENCES companies(id)
        )`);
        // origem = 'impulsionar' quando veio da planilha que o Master subiu.
        db.run(`ALTER TABLE dpo_chat_questions ADD COLUMN origem TEXT`, () => {});
        // Autoavaliação MENSAL (separada da gestão): uma por empresa por mês
        // (referencia = 'AAAA-MM'), cobrindo todos os pilares liberados. Cada
        // pergunta recebe 3, 1, 0 ou N/A.
        db.run(`CREATE TABLE IF NOT EXISTS dpo_self_assessments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER NOT NULL,
            referencia TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'em_andamento',
            created_by INTEGER,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            closed_at DATETIME,
            UNIQUE(company_id, referencia)
        )`);
        // Fluxo: em_andamento (empresa editando) -> salva -> aprovada. Depois de
        // aprovada, só o Master altera (a empresa pede ajuste por chamado).
        db.run(`ALTER TABLE dpo_self_assessments ADD COLUMN approved_by INTEGER`, () => {});
        db.run(`ALTER TABLE dpo_self_assessments ADD COLUMN approved_at DATETIME`, () => {});
        db.run(`UPDATE dpo_self_assessments SET status = 'aprovada' WHERE status = 'concluida'`, () => {});
        db.run(`CREATE TABLE IF NOT EXISTS dpo_self_assessment_events (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            assessment_id INTEGER NOT NULL,
            acao TEXT NOT NULL,
            detalhe TEXT,
            user_id INTEGER,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )`);
        // Chamados (suporte): a empresa abre para o Master, com conversa,
        // anexos, prioridade, status e avaliação do atendimento.
        db.run(`CREATE TABLE IF NOT EXISTS chamados (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER NOT NULL,
            created_by INTEGER,
            categoria TEXT NOT NULL,
            assunto TEXT NOT NULL,
            prioridade TEXT NOT NULL DEFAULT 'media',
            status TEXT NOT NULL DEFAULT 'aberto',
            ref_tipo TEXT,
            ref_id INTEGER,
            nao_lido_master INTEGER NOT NULL DEFAULT 1,
            nao_lido_empresa INTEGER NOT NULL DEFAULT 0,
            avaliacao INTEGER,
            avaliacao_comentario TEXT,
            first_response_at DATETIME,
            closed_at DATETIME,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )`);
        db.run(`CREATE TABLE IF NOT EXISTS chamado_mensagens (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            chamado_id INTEGER NOT NULL,
            user_id INTEGER,
            autor_papel TEXT NOT NULL,
            texto TEXT,
            anexo_url TEXT,
            anexo_nome TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )`);
        // "Ferramentas Impulsionar": biblioteca que o Master publica por pilar
        // (e opcionalmente por pergunta) — planilhas, modelos, treinamentos, links.
        db.run(`CREATE TABLE IF NOT EXISTS dpo_ferramentas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            pillar_key TEXT NOT NULL,
            question_numero TEXT,
            tipo TEXT NOT NULL,
            titulo TEXT NOT NULL,
            descricao TEXT,
            url TEXT NOT NULL,
            original_name TEXT,
            ativo INTEGER NOT NULL DEFAULT 1,
            acessos INTEGER NOT NULL DEFAULT 0,
            created_by INTEGER,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            updated_at DATETIME
        )`);
        // "Material DPO": biblioteca do Master (por pilar + pergunta). Cada material
        // só aparece para a empresa quando o Master o disponibiliza para ela
        // (pasta "📂 Material Impulsionar" dentro da pergunta). Desmarcar = some.
        db.run(`CREATE TABLE IF NOT EXISTS dpo_material_impulsionar (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            pillar_key TEXT NOT NULL,
            question_numero TEXT NOT NULL,
            titulo TEXT NOT NULL,
            descricao TEXT,
            url TEXT NOT NULL,
            original_name TEXT,
            created_by INTEGER,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            updated_at DATETIME
        )`);
        db.run(`CREATE TABLE IF NOT EXISTS dpo_material_impulsionar_share (
            material_id INTEGER NOT NULL,
            company_id INTEGER NOT NULL,
            shared_by INTEGER,
            shared_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            acessos INTEGER NOT NULL DEFAULT 0,
            PRIMARY KEY (material_id, company_id)
        )`);
        // Ferramentas digitais por pergunta (SWOT da Gestão 1.3, PPR do Planejamento 1.1).
        db.run(`CREATE TABLE IF NOT EXISTS dpo_ferramentas_digitais (
            company_id INTEGER NOT NULL,
            chave TEXT NOT NULL,
            ano INTEGER NOT NULL,
            dados TEXT NOT NULL,
            updated_by INTEGER,
            updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY (company_id, chave, ano)
        )`);
        db.run(`CREATE TABLE IF NOT EXISTS dpo_ferramentas_digitais_arquivos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER NOT NULL,
            chave TEXT NOT NULL,
            ano INTEGER NOT NULL,
            tipo TEXT,
            url TEXT NOT NULL,
            original_name TEXT,
            comentario TEXT,
            created_by INTEGER,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )`);
        // Acompanhamentos Impulsionar liberados pelo Master por revenda e pergunta.
        db.run(`CREATE TABLE IF NOT EXISTS dpo_ronda_links (
            token TEXT PRIMARY KEY, company_id INTEGER NOT NULL, ano INTEGER NOT NULL, trimestre INTEGER NOT NULL,
            criado_por INTEGER, criado_em DATETIME DEFAULT CURRENT_TIMESTAMP
        )`);
        db.run(`CREATE TABLE IF NOT EXISTS dpo_ronda_fotos (
            id INTEGER PRIMARY KEY AUTOINCREMENT, company_id INTEGER NOT NULL, ano INTEGER NOT NULL, trimestre INTEGER NOT NULL,
            item TEXT NOT NULL, url TEXT NOT NULL, obs TEXT, autor TEXT, created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )`);
        db.run(`CREATE TABLE IF NOT EXISTS dpo_5s_links (
            token TEXT PRIMARY KEY, company_id INTEGER NOT NULL, ano INTEGER NOT NULL, criado_por INTEGER, criado_em DATETIME DEFAULT CURRENT_TIMESTAMP
        )`);
        db.run(`CREATE TABLE IF NOT EXISTS dpo_5s_resp (
            company_id INTEGER NOT NULL, ano INTEGER NOT NULL, area_id TEXT NOT NULL, mes INTEGER NOT NULL, qid TEXT NOT NULL,
            resp TEXT, foto TEXT, autor TEXT, updated_at TEXT NOT NULL, PRIMARY KEY (company_id, ano, area_id, mes, qid)
        )`);
        db.run(`CREATE TABLE IF NOT EXISTS dpo_chamado_links (token TEXT PRIMARY KEY, company_id INTEGER NOT NULL, criado_por INTEGER, criado_em DATETIME DEFAULT CURRENT_TIMESTAMP)`);
        db.run(`CREATE TABLE IF NOT EXISTS dpo_chamados_pub (id INTEGER PRIMARY KEY AUTOINCREMENT, company_id INTEGER NOT NULL, dados TEXT, foto TEXT, created_at DATETIME DEFAULT CURRENT_TIMESTAMP)`);
        db.run(`CREATE TABLE IF NOT EXISTS dpo_5s_avisos (company_id INTEGER NOT NULL, ano INTEGER NOT NULL, mes INTEGER NOT NULL, area_id TEXT NOT NULL, data TEXT, envios INTEGER DEFAULT 0, ultimo TEXT, canais TEXT, PRIMARY KEY (company_id, ano, mes, area_id))`);
        db.run(`CREATE TABLE IF NOT EXISTS dpo_vis_links (token TEXT PRIMARY KEY, company_id INTEGER NOT NULL, criado_por INTEGER, criado_em DATETIME DEFAULT CURRENT_TIMESTAMP)`);
        db.run(`CREATE TABLE IF NOT EXISTS dpo_vis_just (id INTEGER PRIMARY KEY AUTOINCREMENT, company_id INTEGER NOT NULL, data TEXT NOT NULL, matricula TEXT NOT NULL, ind_id TEXT NOT NULL, dados TEXT, status TEXT, retorno TEXT, created_at DATETIME DEFAULT CURRENT_TIMESTAMP, UNIQUE (company_id, data, matricula, ind_id))`);
        db.run(`CREATE TABLE IF NOT EXISTS dpo_vis_acessos (company_id INTEGER NOT NULL, matricula TEXT NOT NULL, data TEXT NOT NULL, qtd INTEGER DEFAULT 0, PRIMARY KEY (company_id, matricula, data))`);
        db.run(`CREATE TABLE IF NOT EXISTS dpo_ronda_notas (
            company_id INTEGER NOT NULL, ano INTEGER NOT NULL, trimestre INTEGER NOT NULL, item TEXT NOT NULL,
            nota TEXT, autor TEXT, updated_at TEXT NOT NULL, PRIMARY KEY (company_id, ano, trimestre, item)
        )`);
        db.run(`CREATE TABLE IF NOT EXISTS dpo_acomp_liberacoes (
            company_id INTEGER NOT NULL,
            question_key TEXT NOT NULL,
            liberado_por INTEGER,
            liberado_em DATETIME DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY (company_id, question_key)
        )`);
        db.run(`CREATE TABLE IF NOT EXISTS dpo_self_answers (
            assessment_id INTEGER NOT NULL,
            question_key TEXT NOT NULL,
            valor TEXT NOT NULL,
            updated_by INTEGER,
            updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY (assessment_id, question_key)
        )`);

        // Permissão por PASTA do DPO (Checklist / Perguntas Bate-Papo / Material
        // do Pilar), liberada pelo Master para cada empresa. Sem linha = padrão
        // (ver PASTAS_DPO_PADRAO: só o Checklist vem liberado).
        db.run(`CREATE TABLE IF NOT EXISTS dpo_company_folders (
            company_id INTEGER NOT NULL,
            folder_key TEXT NOT NULL,
            enabled INTEGER NOT NULL DEFAULT 0,
            updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY (company_id, folder_key)
        )`);

        // "Material do Pilar" — evidências para defender na auditoria, organizadas
        // por pergunta do pilar e, dentro dela, por ITEM da verificação.
        // Situação de cada item (pendente / em andamento / pronto p/ auditoria):
        db.run(`CREATE TABLE IF NOT EXISTS dpo_material_status (
            company_id INTEGER NOT NULL,
            pillar_key TEXT NOT NULL,
            question_numero TEXT NOT NULL,
            item_numero TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'pendente',
            observacao TEXT,
            updated_by INTEGER,
            updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY (company_id, pillar_key, question_numero, item_numero)
        )`);
        // Evidências de cada item: padrão, ata de treinamento, outra evidência
        // (arquivos enviados) ou link para sistema externo.
        db.run(`CREATE TABLE IF NOT EXISTS dpo_material_evidencias (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER NOT NULL,
            pillar_key TEXT NOT NULL,
            question_numero TEXT NOT NULL,
            item_numero TEXT NOT NULL,
            tipo TEXT NOT NULL,
            titulo TEXT,
            url TEXT NOT NULL,
            original_name TEXT,
            created_by INTEGER,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )`);
        // Check de retenção de treinamento: a empresa cria só as perguntas e o
        // sistema gera um link público (token) para as pessoas treinadas responderem.
        db.run(`CREATE TABLE IF NOT EXISTS dpo_retention_checks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER NOT NULL,
            pillar_key TEXT NOT NULL,
            question_numero TEXT NOT NULL,
            item_numero TEXT NOT NULL,
            titulo TEXT NOT NULL,
            token TEXT UNIQUE NOT NULL,
            perguntas TEXT NOT NULL,
            ativo INTEGER NOT NULL DEFAULT 1,
            created_by INTEGER,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )`);
        // Responsáveis (donos de ação) cadastrados pela empresa para os planos do DPO.
        db.run(`CREATE TABLE IF NOT EXISTS dpo_responsaveis (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER NOT NULL,
            nome TEXT NOT NULL,
            cargo TEXT,
            area TEXT,
            email TEXT,
            telefone TEXT,
            ativo INTEGER DEFAULT 1,
            created_by INTEGER,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )`);
        // "Ministrar Treinamento" (Master): treinamento com módulos (slides feitos
        // no sistema ou arquivos), público, data, rascunho e check de retenção.
        db.run(`CREATE TABLE IF NOT EXISTS treinamentos_mt (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            titulo TEXT NOT NULL,
            objetivo TEXT,
            publico TEXT DEFAULT 'todos',
            data_treinamento TEXT,
            duracao_min INTEGER,
            local TEXT,
            instrutor TEXT,
            company_id INTEGER,
            participantes INTEGER,
            status TEXT DEFAULT 'rascunho',
            modulos TEXT DEFAULT '[]',
            apresentado_em DATETIME,
            created_by INTEGER,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            updated_at DATETIME
        )`);
        db.run(`CREATE TABLE IF NOT EXISTS treinamentos_mt_checks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            treinamento_id INTEGER NOT NULL,
            titulo TEXT NOT NULL,
            token TEXT UNIQUE NOT NULL,
            perguntas TEXT NOT NULL,
            ativo INTEGER NOT NULL DEFAULT 1,
            gerado_por_ia INTEGER DEFAULT 0,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            updated_at DATETIME
        )`);
        db.run(`ALTER TABLE treinamentos_mt ADD COLUMN ata_token TEXT`, () => {});
        db.run(`ALTER TABLE treinamentos_mt ADD COLUMN ata_ativa INTEGER DEFAULT 1`, () => {});
        db.run(`CREATE TABLE IF NOT EXISTS treinamentos_mt_presencas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            treinamento_id INTEGER NOT NULL,
            nome TEXT NOT NULL,
            matricula TEXT,
            cargo TEXT,
            empresa TEXT,
            assinatura TEXT,
            ip TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )`);
        db.run(`CREATE TABLE IF NOT EXISTS treinamentos_mt_respostas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            check_id INTEGER NOT NULL,
            nome TEXT NOT NULL,
            matricula TEXT,
            respostas TEXT NOT NULL,
            acertos INTEGER,
            total_objetivas INTEGER,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )`);
        db.run(`CREATE TABLE IF NOT EXISTS dpo_retention_responses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            check_id INTEGER NOT NULL,
            nome TEXT NOT NULL,
            matricula TEXT,
            respostas TEXT NOT NULL,
            acertos INTEGER,
            total_objetivas INTEGER,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(check_id) REFERENCES dpo_retention_checks(id)
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

// ------------------------------------------------------------
// Cópia de segurança dos uploads dentro do banco: se a pasta de uploads não
// for persistente (sem Volume / UPLOADS_PATH), cada novo deploy apagava as
// fotos. Agora todo arquivo enviado (até 25 MB — fotos, PDFs, anexos) também
// é guardado no banco e, se sumir do disco, é restaurado na hora em que for
// pedido. Vídeos grandes continuam só no disco.
// ------------------------------------------------------------
const LIMITE_COPIA_UPLOAD = 25 * 1024 * 1024;
db.run(`CREATE TABLE IF NOT EXISTS arquivos_persistidos (nome TEXT PRIMARY KEY, mime TEXT, tamanho INTEGER, dados BLOB, created_at DATETIME DEFAULT CURRENT_TIMESTAMP)`);
function copiarUploadParaBanco(f) {
    try {
        if (!f || !f.filename || !f.path || !(f.size <= LIMITE_COPIA_UPLOAD)) return;
        fs.readFile(f.path, (err, dados) => {
            if (err) return;
            db.run(`INSERT OR REPLACE INTO arquivos_persistidos (nome, mime, tamanho, dados) VALUES (?, ?, ?, ?)`, [f.filename, f.mimetype || '', f.size || dados.length, dados], () => {});
        });
    } catch (e) {}
}
app.use((req, res, next) => {
    res.on('finish', () => {
        if (res.statusCode >= 400) return;
        const lista = [];
        if (req.file) lista.push(req.file);
        if (Array.isArray(req.files)) lista.push(...req.files);
        else if (req.files && typeof req.files === 'object') Object.values(req.files).forEach(v => Array.isArray(v) && lista.push(...v));
        lista.forEach(copiarUploadParaBanco);
    });
    next();
});
function restaurarUploadDoBanco(nome) {
    return new Promise(ok => {
        const limpo = path.basename(String(nome || ''));
        if (!limpo) return ok(null);
        db.get(`SELECT mime, dados FROM arquivos_persistidos WHERE nome = ?`, [limpo], (err, row) => {
            if (err || !row || !row.dados) return ok(null);
            const destino = path.join(PASTA_UPLOADS, limpo);
            fs.writeFile(destino, row.dados, () => ok({ caminho: destino, mime: row.mime, dados: row.dados }));
        });
    });
}
// Arquivo pedido em /uploads/... que não está no disco (apagado no deploy) → volta do banco.
app.get('/uploads/:nome', async (req, res, next) => {
    const nome = path.basename(req.params.nome || '');
    if (!nome || fs.existsSync(path.join(PASTA_UPLOADS, nome))) return next();
    const r = await restaurarUploadDoBanco(nome);
    if (!r) return next();
    if (r.mime) res.type(r.mime);
    res.set('Cache-Control', 'public, max-age=86400');
    res.send(r.dados);
});
// Ao subir o servidor, já guarda no banco os arquivos que ainda estão no disco e não têm cópia.
setTimeout(() => {
    try {
        fs.readdir(PASTA_UPLOADS, (err, nomes) => {
            if (err) return;
            db.all(`SELECT nome FROM arquivos_persistidos`, [], (e2, rows) => {
                const ja = new Set((rows || []).map(r => r.nome));
                (nomes || []).filter(n => !ja.has(n)).forEach(n => {
                    const caminho = path.join(PASTA_UPLOADS, n);
                    fs.stat(caminho, (e3, st) => { if (!e3 && st.isFile() && st.size <= LIMITE_COPIA_UPLOAD) copiarUploadParaBanco({ filename: n, path: caminho, size: st.size, mimetype: ({ '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.png': 'image/png', '.webp': 'image/webp', '.gif': 'image/gif', '.pdf': 'application/pdf' })[path.extname(n).toLowerCase()] || '' }); });
                });
            });
        });
    } catch (e) {}
}, 8000);

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
        if (Number(user.aprovacao_pendente) === 1) {
            return res.status(403).json({ error: 'Seu acesso foi criado e está aguardando a aprovação da Impulsionar. Você será avisado assim que for liberado.' });
        }
        db.run(`UPDATE users SET ultimo_login = CURRENT_TIMESTAMP, qtd_logins = COALESCE(qtd_logins, 0) + 1 WHERE id = ?`, [user.id], () => {});

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
    const membros = await dbAll(`SELECT id, name, email, enabled_modules, aprovacao_pendente FROM users WHERE company_id = ? AND role = 'client_admin' ORDER BY id ASC`, [req.params.id]);
    const emp = await dbGet(`SELECT gestor_principal_id FROM companies WHERE id = ?`, [req.params.id]);
    res.json(membros.map(m => ({ ...m, principal: !!(emp && emp.gestor_principal_id && Number(emp.gestor_principal_id) === Number(m.id)) })));
});

// Master marca qual acesso da empresa é o Gestor principal (aprova aprovadores de CAPEX etc.).
app.put('/api/companies/:id/gestor-principal', requireRole('admin'), async (req, res) => {
    try {
        const uid = req.body.userId ? Number(req.body.userId) : null;
        if (uid) {
            const u = await dbGet(`SELECT id FROM users WHERE id = ? AND company_id = ? AND role = 'client_admin'`, [uid, req.params.id]);
            if (!u) return res.status(400).json({ error: 'Usuário não pertence a esta empresa.' });
        }
        await new Promise((resolve, reject) => db.run(`UPDATE companies SET gestor_principal_id = ? WHERE id = ?`, [uid, req.params.id], e => e ? reject(e) : resolve()));
        res.json({ message: uid ? 'Gestor principal definido!' : 'Gestor principal removido.' });
    } catch (e) { res.status(400).json({ error: 'Erro ao definir o gestor principal.' }); }
});
async function ehGestorPrincipalDpo(companyId, userId) {
    const c = await dbGet(`SELECT gestor_principal_id FROM companies WHERE id = ?`, [companyId]);
    return !!(c && c.gestor_principal_id && Number(c.gestor_principal_id) === Number(userId));
}

// Permissão de módulos individual deste acesso (gestor ou membro) — some com
// o padrão da empresa (ver /api/login), então aqui só grava a restrição
// PRÓPRIA deste usuário. enabled_modules === null (nenhuma marcação) limpa a
// restrição própria e volta a valer o padrão da empresa.
// Só o Master libera permissões; aprovar um acesso criado pela empresa libera o login.
app.put('/api/companies/:id/members/:memberId/aprovar', requireRole('admin'), async (req, res) => {
    try {
        const u = await dbGet(`SELECT id, name FROM users WHERE id = ? AND company_id = ? AND role = 'client_admin'`, [req.params.memberId, req.params.id]);
        if (!u) return res.status(404).json({ error: 'Acesso não encontrado.' });
        await new Promise((resolve, reject) => db.run(`UPDATE users SET aprovacao_pendente = 0 WHERE id = ?`, [u.id], e => e ? reject(e) : resolve()));
        notificarPorCompanyAdmins(req.params.id, 'Acesso aprovado', `O acesso de ${u.name} foi aprovado pela Impulsionar e já pode entrar.`);
        res.json({ message: 'Acesso aprovado! Confira as permissões liberadas.' });
    } catch (e) { res.status(400).json({ error: 'Erro ao aprovar o acesso.' }); }
});
app.put('/api/companies/:id/members/:memberId/permissions', requireRole('admin'), async (req, res) => {
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
        const empresa = await dbGet(`SELECT id, name FROM companies WHERE id = ?`, [req.params.id]);
        if (!empresa) return res.status(404).json({ error: 'Empresa não encontrada.' });
        const hash = await bcrypt.hash(password, 10);
        // Criado pela empresa: sem nenhuma permissão e pendente de aprovação do Master.
        const pelaEmpresa = req.user.role !== 'admin';
        db.run(`INSERT INTO users (name, email, password, company_id, role, enabled_modules, aprovacao_pendente) VALUES (?, ?, ?, ?, 'client_admin', ?, ?)`,
            [name, email.trim(), hash, req.params.id, pelaEmpresa ? '[]' : null, pelaEmpresa ? 1 : 0], (err) => {
                if (err) return res.status(400).json({ error: 'Este e-mail já está em uso por outra conta.' });
                if (pelaEmpresa) {
                    notificarMasters('Novo acesso aguardando aprovação', `${empresa.name} cadastrou ${name} (${email.trim()}). Aprove e libere as permissões em Empresas → Acessos.`, 'companies');
                    return res.json({ message: 'Acesso criado! Ele só consegue entrar depois que a Impulsionar aprovar e liberar as permissões.' });
                }
                res.json({ message: 'Acesso criado! Este membro já pode entrar com o e-mail e senha próprios dele.' });
            });
    } catch (e) { res.status(500).json({ error: 'Erro ao criar o acesso.' }); }
});

// Edita cadastro de um acesso da empresa (nome/e-mail/senha) — senha é opcional,
// deixando em branco mantém a atual. Mesmo espírito do PUT de empresa, que já
// permite trocar nome/e-mail/senha do Gestor principal.
app.put('/api/companies/:id/members/:memberId', requireRole('admin', 'client_admin'), async (req, res) => {
    if (req.user.role === 'client_admin' && String(req.params.id) !== String(req.user.companyId)) {
        return res.status(403).json({ error: 'Você só pode editar acessos da sua própria corporação.' });
    }
    const { name, email, password } = req.body;
    if (!name || !email) return res.status(400).json({ error: 'Informe nome e e-mail.' });
    if (password && password.length < 6) return res.status(400).json({ error: 'A nova senha precisa ter pelo menos 6 caracteres.' });
    try {
        const membro = await dbGet(`SELECT id FROM users WHERE id = ? AND company_id = ? AND role = 'client_admin'`, [req.params.memberId, req.params.id]);
        if (!membro) return res.status(404).json({ error: 'Acesso não encontrado nesta empresa.' });
        if (password) {
            const hash = await bcrypt.hash(password, 10);
            db.run(`UPDATE users SET name = ?, email = ?, password = ? WHERE id = ?`, [name, email.trim(), hash, req.params.memberId], (err) => {
                if (err) return res.status(400).json({ error: 'Este e-mail já está em uso por outra conta.' });
                res.json({ message: 'Acesso atualizado!' });
            });
        } else {
            db.run(`UPDATE users SET name = ?, email = ? WHERE id = ?`, [name, email.trim(), req.params.memberId], (err) => {
                if (err) return res.status(400).json({ error: 'Este e-mail já está em uso por outra conta.' });
                res.json({ message: 'Acesso atualizado!' });
            });
        }
    } catch (e) { res.status(500).json({ error: 'Erro ao editar o acesso.' }); }
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
function registrarCreditoVaga(companyId, dias, tipo, motivo, userId, jobId) {
    db.run(`INSERT INTO vaga_creditos_log (company_id, dias, tipo, motivo, user_id, job_posting_id) VALUES (?, ?, ?, ?, ?, ?)`,
        [companyId, dias, tipo, String(motivo || '').slice(0, 300), userId || null, jobId || null], () => {});
}
// Positivo soma ao saldo; negativo retira (nunca deixa o saldo abaixo de zero).
app.post('/api/companies/:id/credito', requireRole('admin'), async (req, res) => {
    const dias = Math.trunc(Number(req.body.dias));
    if (!dias) return res.status(400).json({ error: 'Informe uma quantidade de dias válida.' });
    try {
        const emp = await dbGet(`SELECT id, name, COALESCE(vaga_credito_dias, 0) as saldo FROM companies WHERE id = ?`, [req.params.id]);
        if (!emp) return res.status(404).json({ error: 'Empresa não encontrada.' });
        const mov = dias < 0 ? -Math.min(-dias, emp.saldo) : dias;
        if (!mov) return res.status(400).json({ error: 'A empresa não tem saldo para retirar.' });
        await new Promise((ok, erro) => db.run(`UPDATE companies SET vaga_credito_dias = COALESCE(vaga_credito_dias, 0) + ? WHERE id = ?`, [mov, emp.id], e => e ? erro(e) : ok()));
        registrarCreditoVaga(emp.id, mov, mov > 0 ? 'concedido' : 'retirado', req.body.motivo, req.user.userId, null);
        if (mov > 0) notificarPorCompanyAdmins(emp.id, 'Crédito de vagas liberado! 🎁', `A Impulsionar liberou ${mov} dia(s) para você publicar vagas sem pagar.`, 'jobPostings');
        const total = emp.saldo + mov;
        res.json({ message: mov > 0 ? `+${mov} dias de crédito adicionados!` : `${-mov} dia(s) retirados do saldo.`, total });
    } catch (e) { res.status(400).json({ error: 'Erro ao ajustar crédito.' }); }
});

app.get('/api/admin/vaga-creditos', requireRole('admin'), async (req, res) => {
    try {
        const empresas = await dbAll(`SELECT c.id, c.name, c.logo_url, COALESCE(c.vaga_credito_dias, 0) as saldo,
                (SELECT COUNT(*) FROM job_postings jp WHERE jp.company_id = c.id AND jp.deleted_at IS NULL AND jp.status = 'active' AND jp.expires_at > CURRENT_TIMESTAMP) as ativas,
                (SELECT COUNT(*) FROM job_postings jp WHERE jp.company_id = c.id AND jp.deleted_at IS NULL AND jp.status = 'pending_payment') as aguardandoPagamento,
                (SELECT COALESCE(SUM(dias), 0) FROM vaga_creditos_log l WHERE l.company_id = c.id AND l.tipo IN ('concedido')) as totalConcedido,
                (SELECT COALESCE(-SUM(dias), 0) FROM vaga_creditos_log l WHERE l.company_id = c.id AND l.tipo = 'consumido') as totalConsumido
            FROM companies c ORDER BY c.name COLLATE NOCASE`);
        const log = await dbAll(`SELECT l.*, c.name as companyName, u.name as userName, jp.title as vagaTitulo FROM vaga_creditos_log l
            LEFT JOIN companies c ON c.id = l.company_id LEFT JOIN users u ON u.id = l.user_id LEFT JOIN job_postings jp ON jp.id = l.job_posting_id
            ORDER BY l.created_at DESC, l.id DESC LIMIT 200`);
        res.json({ empresas, log });
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar créditos.' }); }
});

app.get('/api/minha-empresa/credito-vagas', requireRole('client_admin'), async (req, res) => {
    try {
        const c = await dbGet(`SELECT COALESCE(vaga_credito_dias, 0) as saldo FROM companies WHERE id = ?`, [req.user.companyId]);
        res.json({ saldo: c ? c.saldo : 0 });
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar crédito.' }); }
});

async function publicarVagaSemPagamento(vaga, dias, userId) {
    await new Promise((ok, erro) => db.run(
        `UPDATE job_postings SET status = 'active', approved_by = COALESCE(approved_by, ?), approved_at = COALESCE(approved_at, CURRENT_TIMESTAMP),
            published_at = CURRENT_TIMESTAMP, expires_at = datetime(CURRENT_TIMESTAMP, '+${Number(dias)} days'), paid_with_credit = 1 WHERE id = ?`,
        [userId, vaga.id], e => e ? erro(e) : ok()));
}
// A empresa (ou o Master) usa o saldo de crédito para publicar uma vaga que
// está aguardando pagamento.
app.post('/api/job-postings/:id/usar-credito', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const vaga = await dbGet(`SELECT jp.*, vp.days as planDays FROM job_postings jp LEFT JOIN vaga_plans vp ON vp.id = jp.vaga_plan_id WHERE jp.id = ? AND jp.deleted_at IS NULL`, [req.params.id]);
        if (!vaga) return res.status(404).json({ error: 'Vaga não encontrada.' });
        if (req.user.role === 'client_admin' && vaga.company_id !== req.user.companyId) return res.status(403).json({ error: 'Esta vaga não pertence à sua empresa.' });
        if (vaga.status !== 'pending_payment') return res.status(400).json({ error: 'Esta vaga não está aguardando pagamento.' });
        const dias = vaga.planDays || 30;
        const emp = await dbGet(`SELECT COALESCE(vaga_credito_dias, 0) as saldo FROM companies WHERE id = ?`, [vaga.company_id]);
        if (!emp || emp.saldo < dias) return res.status(400).json({ error: `Crédito insuficiente: esta vaga precisa de ${dias} dia(s) e o saldo é ${emp ? emp.saldo : 0}. Peça para a Impulsionar liberar crédito.` });
        await publicarVagaSemPagamento(vaga, dias, req.user.userId);
        db.run(`UPDATE companies SET vaga_credito_dias = vaga_credito_dias - ? WHERE id = ?`, [dias, vaga.company_id], () => {});
        registrarCreditoVaga(vaga.company_id, -dias, 'consumido', `Vaga "${vaga.title}" publicada com crédito`, req.user.userId, vaga.id);
        res.json({ message: `Vaga publicada usando ${dias} dia(s) de crédito!` });
    } catch (e) { res.status(400).json({ error: 'Erro ao publicar com crédito.' }); }
});
// Só o Master: libera a vaga como cortesia, sem cobrar e sem mexer no saldo.
app.post('/api/job-postings/:id/liberar-cortesia', requireRole('admin'), async (req, res) => {
    try {
        const vaga = await dbGet(`SELECT jp.*, vp.days as planDays FROM job_postings jp LEFT JOIN vaga_plans vp ON vp.id = jp.vaga_plan_id WHERE jp.id = ? AND jp.deleted_at IS NULL`, [req.params.id]);
        if (!vaga) return res.status(404).json({ error: 'Vaga não encontrada.' });
        if (!['pending_payment', 'pendente_aprovacao'].includes(vaga.status)) return res.status(400).json({ error: 'Só dá para liberar vagas aguardando aprovação ou pagamento.' });
        const dias = Math.max(1, Math.min(365, Math.trunc(Number(req.body.dias)) || vaga.planDays || 30));
        await publicarVagaSemPagamento(vaga, dias, req.user.userId);
        registrarCreditoVaga(vaga.company_id, 0, 'cortesia', `Vaga "${vaga.title}" liberada sem pagamento por ${dias} dia(s)${req.body.motivo ? ' — ' + req.body.motivo : ''}`, req.user.userId, vaga.id);
        notificarPorCompanyAdmins(vaga.company_id, 'Vaga liberada pela Impulsionar! 🎁', `"${vaga.title}" foi publicada sem custo por ${dias} dia(s).`, 'jobPostings');
        res.json({ message: `Vaga liberada sem pagamento por ${dias} dia(s)!` });
    } catch (e) { res.status(400).json({ error: 'Erro ao liberar a vaga.' }); }
});

// ---------- Acessos dos candidatos (visão do Master) ----------
function senhaAleatoria() { return crypto.randomBytes(6).toString('base64').replace(/[^a-zA-Z0-9]/g, '').slice(0, 6) + Math.floor(10 + Math.random() * 89); }
app.get('/api/admin/candidate-accesses', requireRole('admin'), async (req, res) => {
    try {
        const lista = await dbAll(`SELECT u.id, u.name, u.email, u.ultimo_login, COALESCE(u.qtd_logins, 0) as qtd_logins,
                COALESCE(u.criado_em, cp.created_at) as criado_em, cp.phone, cp.city, cp.desired_role, cp.status, cp.origem, cp.resume_url, cp.photo_url,
                cp.experiences_json, cp.skills, cp.bio, cp.birth_date, cp.curriculo_completo_em as atualizado, cp.modalidade, cp.education_json, cp.education_level, cp.first_job, cp.curriculo_completo_em,
                (SELECT COUNT(*) FROM job_applications ja WHERE ja.candidate_user_id = u.id) as candidaturas
            FROM users u LEFT JOIN candidate_profiles cp ON cp.user_id = u.id WHERE u.role = 'candidate' ORDER BY COALESCE(u.criado_em, cp.created_at) DESC, u.id DESC`);
        res.json(lista.map(c => {
            const itens = [c.phone, c.city, c.desired_role, c.photo_url, String(c.bio || '').length >= 60 ? 1 : '', String(c.skills || '').split(',').filter(x => x.trim()).length >= 3 ? 1 : '', c.birth_date, c.modalidade, c.education_level, c.education_json && c.education_json !== '[]' ? 1 : '', c.first_job || (c.experiences_json && c.experiences_json !== '[]') ? 1 : ''];
            const completo = Math.round(100 * itens.filter(Boolean).length / itens.length);
            const { experiences_json, skills, bio, birth_date, modalidade, education_json, education_level, first_job, ...resto } = c;
            return { ...resto, completo };
        }));
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar acessos.' }); }
});
app.post('/api/admin/candidate-accesses/pedir-atualizacao', requireRole('admin'), async (req, res) => {
    try {
        const l = await dbAll(`SELECT u.id FROM users u LEFT JOIN candidate_profiles cp ON cp.user_id = u.id WHERE u.role = 'candidate' AND cp.curriculo_completo_em IS NULL`);
        l.forEach(u => notificar(u.id, 'Atualize o seu currículo 📝', 'O Portal de Vagas ganhou o currículo padrão. Complete experiências, atividades e formação para poder se candidatar às vagas.', 'portalPerfil'));
        res.json({ message: l.length ? `Aviso enviado para ${l.length} candidato(s) com currículo incompleto.` : 'Todos os candidatos já estão com o currículo completo.' });
    } catch (e) { res.status(500).json({ error: 'Erro ao enviar avisos.' }); }
});
app.post('/api/admin/candidate-accesses', requireRole('admin'), async (req, res) => {
    const b = req.body || {};
    const name = String(b.name || '').trim(), email = String(b.email || '').trim().toLowerCase();
    if (!name || !/^\S+@\S+\.\S+$/.test(email)) return res.status(400).json({ error: 'Informe nome e um e-mail válido.' });
    const senha = String(b.password || '').trim() || senhaAleatoria();
    if (senha.length < 6) return res.status(400).json({ error: 'A senha precisa ter pelo menos 6 caracteres.' });
    try {
        const hash = await bcrypt.hash(senha, 10);
        const userId = await new Promise((ok, erro) => db.run(`INSERT INTO users (name, email, password, role, criado_em) VALUES (?, ?, ?, 'candidate', CURRENT_TIMESTAMP)`, [name, email, hash], function (e) { e ? erro(e) : ok(this.lastID); }))
            .catch(() => { throw new Error('Este e-mail já tem acesso.'); });
        const city = String(b.city || '').trim();
        await new Promise((ok, erro) => db.run(`INSERT INTO candidate_profiles (user_id, phone, desired_role, city, state, status, origem) VALUES (?, ?, ?, ?, ?, ?, 'master')`,
            [userId, String(b.phone || '').trim(), String(b.desired_role || '').trim(), city, city.includes('/') ? city.split('/')[1].trim().toUpperCase() : '', b.aprovado === false ? 'pending' : 'approved'], e => e ? erro(e) : ok()));
        const base = appBaseUrlAtiva || process.env.APP_URL || `http://localhost:${PORT}`;
        let emailEnviado = false;
        if (b.enviarEmail) {
            try {
                await transporter.sendMail({ from: process.env.SMTP_FROM || process.env.SMTP_USER || EMAIL_API.remetente, to: email, subject: 'Seu acesso ao Portal de Vagas Impulsionar',
                    html: `<p>Olá, ${name.split(' ')[0]}!</p><p>A Impulsionar criou o seu acesso ao Portal de Vagas.</p><p><b>Link:</b> <a href="${base}/#entrar=${encodeURIComponent(email)}">${base}</a><br><b>E-mail:</b> ${email}<br><b>Senha:</b> ${senha}</p><p>Recomendamos trocar a senha no primeiro acesso e completar o seu currículo.</p>` });
                emailEnviado = true;
            } catch (e) {}
        }
        res.json({ message: 'Acesso criado!', id: userId, senha, link: `${base}/#entrar=${encodeURIComponent(email)}`, emailEnviado });
    } catch (e) { res.status(400).json({ error: e.message || 'Erro ao criar acesso.' }); }
});
app.put('/api/admin/candidate-accesses/:id', requireRole('admin'), async (req, res) => {
    const b = req.body || {};
    const name = String(b.name || '').trim(), email = String(b.email || '').trim().toLowerCase();
    if (!name || !/^\S+@\S+\.\S+$/.test(email)) return res.status(400).json({ error: 'Informe nome e um e-mail válido.' });
    try {
        const u = await dbGet(`SELECT id FROM users WHERE id = ? AND role = 'candidate'`, [req.params.id]);
        if (!u) return res.status(404).json({ error: 'Candidato não encontrado.' });
        await new Promise((ok, erro) => db.run(`UPDATE users SET name = ?, email = ? WHERE id = ?`, [name, email, u.id], e => e ? erro(new Error('Este e-mail já é usado por outro acesso.')) : ok()));
        const city = String(b.city || '').trim();
        db.run(`INSERT OR IGNORE INTO candidate_profiles (user_id) VALUES (?)`, [u.id], () => {
            db.run(`UPDATE candidate_profiles SET phone = ?, desired_role = ?, city = ?, state = CASE WHEN ? <> '' THEN ? ELSE state END WHERE user_id = ?`,
                [String(b.phone || '').trim(), String(b.desired_role || '').trim(), city, city.includes('/') ? 'x' : '', city.includes('/') ? city.split('/')[1].trim().toUpperCase() : '', u.id], () => res.json({ message: 'Acesso atualizado!' }));
        });
    } catch (e) { res.status(400).json({ error: e.message || 'Erro ao atualizar.' }); }
});
app.post('/api/admin/candidate-accesses/:id/reset-password', requireRole('admin'), async (req, res) => {
    try {
        const u = await dbGet(`SELECT id, email FROM users WHERE id = ? AND role = 'candidate'`, [req.params.id]);
        if (!u) return res.status(404).json({ error: 'Candidato não encontrado.' });
        const senha = senhaAleatoria();
        const hash = await bcrypt.hash(senha, 10);
        await new Promise((ok, erro) => db.run(`UPDATE users SET password = ? WHERE id = ?`, [hash, u.id], e => e ? erro(e) : ok()));
        const base = appBaseUrlAtiva || process.env.APP_URL || `http://localhost:${PORT}`;
        res.json({ message: 'Nova senha gerada!', senha, link: `${base}/#entrar=${encodeURIComponent(u.email)}` });
    } catch (e) { res.status(400).json({ error: 'Erro ao gerar senha.' }); }
});
app.delete('/api/admin/candidate-accesses/:id', requireRole('admin'), async (req, res) => {
    try {
        const u = await dbGet(`SELECT id FROM users WHERE id = ? AND role = 'candidate'`, [req.params.id]);
        if (!u) return res.status(404).json({ error: 'Candidato não encontrado.' });
        for (const sql of [`DELETE FROM job_applications WHERE candidate_user_id = ?`, `DELETE FROM job_posting_likes WHERE candidate_user_id = ?`,
            `DELETE FROM candidate_messages WHERE candidate_user_id = ?`, `DELETE FROM candidate_profiles WHERE user_id = ?`, `DELETE FROM users WHERE id = ?`]) {
            await new Promise(ok => db.run(sql, [u.id], () => ok()));
        }
        res.json({ message: 'Acesso excluído.' });
    } catch (e) { res.status(400).json({ error: 'Erro ao excluir.' }); }
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

// Lista oficial de municípios (IBGE) para o cadastro de candidatos escolher a cidade certa.
// Busca uma vez no IBGE e guarda em memória e em arquivo (sobrevive a reinícios).
let CIDADES_IBGE = null, CIDADES_IBGE_PROMESSA = null;
const ARQ_CIDADES_IBGE = path.join(__dirname, 'cidades-ibge.json');
function semAcentoCidade(s) { return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/\s+/g, ' ').trim(); }
async function carregarCidadesIbge() {
    if (CIDADES_IBGE) return CIDADES_IBGE;
    if (CIDADES_IBGE_PROMESSA) return CIDADES_IBGE_PROMESSA;
    CIDADES_IBGE_PROMESSA = (async () => {
        try { const l = JSON.parse(fs.readFileSync(ARQ_CIDADES_IBGE, 'utf8')); if (Array.isArray(l) && l.length > 5000) { CIDADES_IBGE = l; return l; } } catch (e) {}
        try {
            const r = await fetch('https://servicodados.ibge.gov.br/api/v1/localidades/municipios?view=nivelado');
            const j = await r.json();
            const l = j.map(m => [m['municipio-nome'], m['UF-sigla']]).filter(x => x[0] && x[1]).sort((a, b) => a[0].localeCompare(b[0], 'pt-BR'));
            if (l.length > 5000) { CIDADES_IBGE = l; try { fs.writeFileSync(ARQ_CIDADES_IBGE, JSON.stringify(l)); } catch (e) {} }
            return CIDADES_IBGE;
        } catch (e) { console.warn('⚠️  Não consegui baixar a lista de cidades do IBGE:', e.message); return null; }
        finally { CIDADES_IBGE_PROMESSA = null; }
    })();
    return CIDADES_IBGE_PROMESSA;
}
app.get('/api/public/cidades', async (req, res) => {
    const l = await carregarCidadesIbge();
    if (!l) return res.status(503).json({ error: 'Lista de cidades indisponível no momento.' });
    res.set('Cache-Control', 'public, max-age=86400');
    res.json(l);
});
app.get('/api/public/cep/:cep', async (req, res) => {
    const cep = String(req.params.cep || '').replace(/\D/g, '');
    if (cep.length !== 8) return res.status(400).json({ error: 'CEP inválido.' });
    for (const url of [`https://viacep.com.br/ws/${cep}/json/`, `https://brasilapi.com.br/api/cep/v1/${cep}`]) {
        try {
            const r = await fetch(url); if (!r.ok) continue; const j = await r.json(); if (j.erro) continue;
            return res.json({ cep, cidade: j.localidade || j.city || '', uf: j.uf || j.state || '', bairro: j.bairro || j.neighborhood || '', rua: j.logradouro || j.street || '' });
        } catch (e) {}
    }
    res.status(404).json({ error: 'CEP não encontrado.' });
});

app.post('/api/portal/register', async (req, res) => {
    const { name, email, password, phone, desired_role } = req.body;
    let { city } = req.body;
    if (!name || !email || !password) return res.status(400).json({ error: 'Nome, e-mail e senha são obrigatórios.' });
    if (String(password).length < 6) return res.status(400).json({ error: 'A senha precisa ter pelo menos 6 caracteres.' });
    try {
        // Cidade sempre no formato oficial "Cidade/UF" (lista do IBGE), para não entrar cidade escrita errada.
        if (city) {
            const lista = await carregarCidadesIbge();
            if (lista) {
                const [nomeCid, ufCid] = String(city).split('/').map(x => x.trim());
                const achada = lista.find(([n, uf]) => semAcentoCidade(n) === semAcentoCidade(nomeCid) && (!ufCid || uf.toUpperCase() === ufCid.toUpperCase()));
                if (!achada) return res.status(400).json({ error: 'Escolha a sua cidade na lista (ou preencha o CEP).' });
                city = `${achada[0]}/${achada[1]}`;
            }
        }
        const hash = await bcrypt.hash(password, 10);
        const aprovacaoAutomatica = await automacaoLigada('auto_approve_resumes');
        db.run(`INSERT INTO users (name, email, password, role) VALUES (?, ?, ?, 'candidate')`, [String(name).trim(), String(email).trim().toLowerCase(), hash], function (err) {
            if (err) return res.status(400).json({ error: 'Este e-mail já está cadastrado. Faça login ou use "Esqueci minha senha".' });
            const userId = this.lastID;
            const b = req.body;
            db.run(`UPDATE users SET criado_em = CURRENT_TIMESTAMP WHERE id = ?`, [userId], () => {});
            db.all(`SELECT id FROM users WHERE role = 'admin'`, [], (eA, admins) => {
                if (!eA) (admins || []).forEach(ad => notificar(ad.id, 'Novo candidato pelo link', `${String(name).trim()}${desired_role ? ' — ' + desired_role : ''}${city ? ' (' + city + ')' : ''}`, 'acessosCandidatos'));
            });
            db.run(`INSERT INTO candidate_profiles (user_id, phone, desired_role, city, status, cep, neighborhood, state, first_job, lgpd_at, origem) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
                [userId, phone || '', desired_role || '', city || '', aprovacaoAutomatica ? 'approved' : 'pending', String(b.cep || '').replace(/\D/g, '').slice(0, 8), String(b.neighborhood || '').slice(0, 120), city ? city.split('/')[1] || '' : '', b.first_job ? 1 : 0, b.lgpd ? new Date().toISOString() : null, String(b.origem || '').slice(0, 40)], (e2) => {
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

// Campos obrigatórios do currículo padrão — a mesma regra vale na tela (que
// destaca o que falta) e aqui no servidor (para ninguém salvar incompleto).
function jsonListaCurriculo(v) { try { const l = JSON.parse(v || '[]'); return Array.isArray(l) ? l : []; } catch (e) { return []; } }
function pendenciasCurriculo(b) {
    const falta = [];
    const txt = k => String(b[k] || '').trim();
    if (!txt('name')) falta.push('Nome completo');
    if (txt('phone').replace(/\D/g, '').length < 10) falta.push('Telefone/WhatsApp com DDD');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(txt('birth_date'))) falta.push('Data de nascimento');
    if (!txt('city')) falta.push('Cidade atual');
    if (!txt('desired_role')) falta.push('Cargo desejado');
    if (!txt('modalidade')) falta.push('Modelo de trabalho');
    if (!txt('tipos_vaga')) falta.push('Tipo de vaga que busca');
    if (txt('bio').length < 60) falta.push('Resumo profissional (mín. 60 caracteres)');
    if (!txt('education_level')) falta.push('Escolaridade');
    const formacoes = jsonListaCurriculo(b.education_json).filter(f => f && (f.institution || f.course));
    if (!formacoes.length) falta.push('Formação (ao menos uma)');
    else if (formacoes.some(f => !f.institution || !f.level || (!f.end && !f.current))) falta.push('Formação: nível, instituição e conclusão');
    if (!b.first_job) {
        const exps = jsonListaCurriculo(b.experiences_json).filter(e => e && (e.role || e.company));
        if (!exps.length) falta.push('Experiência profissional (ou marque primeiro emprego)');
        else if (exps.some(e => !e.role || !e.company || !e.start || (!e.end && !e.current))) falta.push('Experiências: cargo, empresa, início e fim');
        else if (exps.some(e => String(e.activities || '').split('\n').filter(x => x.trim()).length < 2)) falta.push('Experiências: informe ao menos 2 atividades realizadas em cada uma');
        else if (exps.some(e => e.end && e.start && e.end < e.start)) falta.push('Experiências: data de fim antes do início');
    }
    if (String(b.skills || '').split(',').map(x => x.trim()).filter(Boolean).length < 3) falta.push('Habilidades (mín. 3)');
    if (!txt('desired_cities') && !txt('desired_states')) falta.push('Onde deseja trabalhar (ao menos uma cidade ou estado)');
    if (!txt('photo_url')) falta.push('Foto de perfil');
    return falta;
}
app.put('/api/portal/me', requireRole('candidate'), async (req, res) => {
    const b = req.body || {};
    const falta = pendenciasCurriculo(b);
    if (falta.length) return res.status(400).json({ error: 'Complete os campos obrigatórios: ' + falta.join(', ') + '.', falta });
    const t = (k, n = 2000) => String(b[k] == null ? '' : b[k]).trim().slice(0, n);
    try {
        await new Promise((ok, erro) => db.run(`UPDATE users SET name = ? WHERE id = ?`, [t('name', 120), req.user.userId], e => e ? erro(e) : ok()));
        await new Promise((ok, erro) => db.run(
            `UPDATE candidate_profiles SET phone = ?, desired_role = ?, city = ?, state = ?, bio = ?, skills = ?, linkedin_url = ?, resume_url = ?,
                photo_url = ?, gender = ?, education_level = ?, languages = ?, first_job = ?, experiences_json = ?,
                desired_states = ?, desired_cities = ?, birth_date = ?, cnh = ?, pretensao_salarial = ?, modalidade = ?, disp_viagem = ?, disp_mudanca = ?,
                pcd = ?, education_json = ?, courses_json = ?, disponibilidade_inicio = ?, cep = ?, neighborhood = ?, tipos_vaga = ?,
                curriculo_completo_em = COALESCE(curriculo_completo_em, CURRENT_TIMESTAMP)
             WHERE user_id = ?`,
            [t('phone', 30), t('desired_role', 120), t('city', 120), t('city').includes('/') ? t('city').split('/').pop().trim().toUpperCase().slice(0, 2) : '', t('bio', 3000), t('skills', 1500), t('linkedin_url', 300), t('resume_url', 500),
                t('photo_url', 500), t('gender', 40), t('education_level', 80), t('languages', 600), b.first_job ? 1 : 0, JSON.stringify(jsonListaCurriculo(b.experiences_json)).slice(0, 30000),
                t('desired_states', 300), t('desired_cities', 1500), t('birth_date', 10), t('cnh', 10), t('pretensao_salarial', 60), t('modalidade', 40), b.disp_viagem ? 1 : 0, b.disp_mudanca ? 1 : 0,
                t('pcd', 200), JSON.stringify(jsonListaCurriculo(b.education_json)).slice(0, 15000), JSON.stringify(jsonListaCurriculo(b.courses_json)).slice(0, 15000), t('disponibilidade_inicio', 40),
                t('cep', 9).replace(/\D/g, ''), t('neighborhood', 120), String(b.tipos_vaga || '').split(',').map(x => x.trim()).filter(x => TIPOS_VAGA.includes(x)).join(', '), req.user.userId],
            e => e ? erro(e) : ok()));
        res.json({ message: 'Currículo atualizado!' });
    } catch (e) { res.status(400).json({ error: 'Erro ao salvar o currículo.' }); }
});

// Lista as vagas ativas e ainda dentro do prazo pago — visível para o próprio candidato.
app.get('/api/portal/vagas', requireRole('candidate'), async (req, res) => {
    try {
        const perfil = await dbGet(`SELECT desired_states, desired_cities, tipos_vaga FROM candidate_profiles WHERE user_id = ?`, [req.user.userId]);
        const tiposCand = String(perfil?.tipos_vaga || '').split(',').map(x => x.trim()).filter(Boolean);
        const estadosDesejados = (perfil?.desired_states || '').split(',').map(s => s.trim()).filter(Boolean);
        const cidadesDesejadas = (perfil?.desired_cities || '').split(',').map(s => s.split('/')[0].trim().toLowerCase()).filter(Boolean);
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
            const categorias = categoriasDaVaga(v);
            return { ...v, categorias, paraVoce: tiposCand.length ? categorias.some(c => tiposCand.includes(c)) : false, jaCandidatei: !!v.jaCandidatei, euCurti: !!v.euCurti, minhaRegiao: !!(v.is_remote || bateEstado || bateCidade || (estadosDesejados.length === 0 && cidadesDesejadas.length === 0)) };
        });
        const resultado = filtrarPorRegiao ? comFlag.filter(v => v.minhaRegiao) : comFlag;
        res.json(resultado);
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar vagas.' }); }
});

app.post('/api/portal/vagas/:id/apply', requireRole('candidate'), async (req, res) => {
    try {
        const vaga = await dbGet(`SELECT * FROM job_postings WHERE id = ? AND status = 'active' AND deleted_at IS NULL AND expires_at > CURRENT_TIMESTAMP`, [req.params.id]);
        if (!vaga) return res.status(404).json({ error: 'Vaga não encontrada ou não está mais disponível.' });
        const perfilCand = await dbGet(`SELECT curriculo_completo_em FROM candidate_profiles WHERE user_id = ?`, [req.user.userId]);
        if (!perfilCand || !perfilCand.curriculo_completo_em) return res.status(400).json({ error: 'Complete o seu currículo antes de se candidatar — as empresas precisam ver suas experiências e atividades.', completarCurriculo: true });
        db.run(`INSERT INTO job_applications (job_posting_id, candidate_user_id) VALUES (?, ?)`, [req.params.id, req.user.userId], (err) => {
            if (err) return res.status(400).json({ error: 'Você já se candidatou a esta vaga.' });
            notificarPorCompanyAdmins(vaga.company_id, 'Nova candidatura 📄', `Novo currículo para "${vaga.title}". Abra em Vagas Ofertadas → Ver currículos.`, 'jobPostings');
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
                    jp.education, jp.languages, jp.requirements, jp.responsibilities, jp.benefits, jp.photo_url, jp.work_schedule, jp.pcd, jp.contract_type, jp.published_at,
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

app.get('/api/portal/config', async (req, res) => {
    res.json({ mostrarNumeros: await automacaoLigada('portal_mostrar_numeros') });
});
// Banco de currículos da empresa (candidatos do portal guardados por função)
app.post('/api/applications/:id/banco', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const { a, erro } = await candidaturaComAcesso(req, req.params.id); if (erro) return res.status(erro[0]).json({ error: erro[1] });
        const funcao = String(req.body.funcao || '').trim().slice(0, 120);
        if (!funcao) return res.status(400).json({ error: 'Escolha a função para guardar o currículo.' });
        await new Promise((ok, er) => db.run(`INSERT INTO company_talent_bank (company_id, candidate_user_id, funcao, observacao, origem_vaga_id, user_id) VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(company_id, candidate_user_id) DO UPDATE SET funcao = excluded.funcao, observacao = excluded.observacao`,
            [a.company_id, a.candidate_user_id, funcao, String(req.body.observacao || '').slice(0, 500), a.job_posting_id, req.user.userId], e => e ? er(e) : ok()));
        res.json({ message: `Currículo guardado no Banco de Currículos como "${funcao}".` });
    } catch (e) { res.status(400).json({ error: 'Erro ao guardar no banco.' }); }
});
app.get('/api/banco-curriculos-portal', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const params = []; let filtro = '';
        if (req.user.role === 'client_admin') { filtro = 'WHERE b.company_id = ?'; params.push(req.user.companyId); }
        const lista = await dbAll(`SELECT cp.*, u.id as userId, u.name, u.email, b.id as bancoId, b.funcao, b.observacao, b.created_at as guardado_em, b.company_id as bancoEmpresa,
                c.name as companyName, jp.title as vagaOrigem
            FROM company_talent_bank b JOIN users u ON u.id = b.candidate_user_id LEFT JOIN candidate_profiles cp ON cp.user_id = u.id
            LEFT JOIN companies c ON c.id = b.company_id LEFT JOIN job_postings jp ON jp.id = b.origem_vaga_id ${filtro} ORDER BY b.funcao COLLATE NOCASE, u.name COLLATE NOCASE`, params);
        res.json(lista);
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar o banco.' }); }
});
app.put('/api/banco-curriculos-portal/:id', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const b = await dbGet(`SELECT * FROM company_talent_bank WHERE id = ?`, [req.params.id]);
        if (!b || (req.user.role === 'client_admin' && b.company_id !== req.user.companyId)) return res.status(404).json({ error: 'Registro não encontrado.' });
        db.run(`UPDATE company_talent_bank SET funcao = ?, observacao = ? WHERE id = ?`, [String(req.body.funcao || b.funcao).slice(0, 120), String(req.body.observacao ?? b.observacao ?? '').slice(0, 500), b.id], () => res.json({ message: 'Atualizado!' }));
    } catch (e) { res.status(400).json({ error: 'Erro ao atualizar.' }); }
});
app.delete('/api/banco-curriculos-portal/:id', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const b = await dbGet(`SELECT * FROM company_talent_bank WHERE id = ?`, [req.params.id]);
        if (!b || (req.user.role === 'client_admin' && b.company_id !== req.user.companyId)) return res.status(404).json({ error: 'Registro não encontrado.' });
        db.run(`DELETE FROM company_talent_bank WHERE id = ?`, [b.id], () => res.json({ message: 'Removido do banco.' }));
    } catch (e) { res.status(400).json({ error: 'Erro ao remover.' }); }
});
app.get('/api/portal/my-applications', requireRole('candidate'), async (req, res) => {
    try {
        const lista = await dbAll(
            `SELECT ja.id, ja.applied_at, COALESCE(ja.status, 'recebida') as status, ja.status_em, ja.entrevista_em, ja.entrevista_local, ja.entrevista_resposta, ja.visto_empresa_em,
                    jp.id as vagaId, jp.title, jp.location, jp.state, jp.is_remote, c.name as companyName, c.logo_url as companyLogo
             FROM job_applications ja JOIN job_postings jp ON jp.id = ja.job_posting_id JOIN companies c ON c.id = jp.company_id
             WHERE ja.candidate_user_id = ? ORDER BY ja.applied_at DESC`,
            [req.user.userId]
        );
        const ids = lista.map(l => l.id);
        const eventos = ids.length ? await dbAll(`SELECT id, application_id, autor, tipo, texto, created_at, lido_candidato FROM application_events WHERE application_id IN (${ids.map(() => '?').join(',')}) ORDER BY created_at ASC, id ASC`, ids) : [];
        res.json(lista.map(l => ({ ...l, eventos: eventos.filter(e => e.application_id === l.id), naoLidas: eventos.filter(e => e.application_id === l.id && e.autor === 'empresa' && !e.lido_candidato).length })));
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar candidaturas.' }); }
});
app.post('/api/portal/my-applications/lidas', requireRole('candidate'), (req, res) => {
    db.run(`UPDATE application_events SET lido_candidato = 1 WHERE autor = 'empresa' AND application_id IN (SELECT id FROM job_applications WHERE candidate_user_id = ?)`, [req.user.userId], () => res.json({ ok: true }));
});

// ---------- Gestão das candidaturas (empresa dona da vaga ou Master) ----------
async function candidaturaComAcesso(req, appId) {
    const a = await dbGet(`SELECT ja.*, jp.company_id, jp.title as vagaTitulo, u.name as candNome, u.email as candEmail, c.name as empresaNome
        FROM job_applications ja JOIN job_postings jp ON jp.id = ja.job_posting_id JOIN users u ON u.id = ja.candidate_user_id JOIN companies c ON c.id = jp.company_id WHERE ja.id = ?`, [appId]);
    if (!a) return { erro: [404, 'Candidatura não encontrada.'] };
    if (req.user.role === 'client_admin' && a.company_id !== req.user.companyId) return { erro: [403, 'Esta candidatura não é de uma vaga da sua empresa.'] };
    if (req.user.role === 'candidate' && a.candidate_user_id !== req.user.userId) return { erro: [403, 'Candidatura de outra pessoa.'] };
    return { a };
}
function registrarEventoCandidatura(appId, autor, tipo, texto, userId) {
    return new Promise(ok => db.run(`INSERT INTO application_events (application_id, autor, tipo, texto, user_id, lido_candidato, lido_empresa) VALUES (?, ?, ?, ?, ?, ?, ?)`,
        [appId, autor, tipo, String(texto || '').slice(0, 3000), userId || null, autor === 'candidato' ? 1 : 0, autor === 'empresa' ? 1 : 0], () => ok()));
}
function avisarCandidatoPorEmail(a, assunto, html) {
    try { transporter.sendMail({ from: process.env.SMTP_FROM || process.env.SMTP_USER || EMAIL_API.remetente, to: a.candEmail, subject: assunto, html }).catch(() => {}); } catch (e) {}
}
app.get('/api/applications/:id/eventos', requireRole('admin', 'client_admin', 'candidate'), async (req, res) => {
    try {
        const { a, erro } = await candidaturaComAcesso(req, req.params.id); if (erro) return res.status(erro[0]).json({ error: erro[1] });
        const eventos = await dbAll(`SELECT id, autor, tipo, texto, created_at FROM application_events WHERE application_id = ? ORDER BY created_at ASC, id ASC`, [a.id]);
        if (req.user.role === 'candidate') db.run(`UPDATE application_events SET lido_candidato = 1 WHERE application_id = ? AND autor = 'empresa'`, [a.id], () => {});
        else db.run(`UPDATE application_events SET lido_empresa = 1 WHERE application_id = ? AND autor = 'candidato'`, [a.id], () => {});
        res.json({ status: a.status || 'recebida', entrevista_em: a.entrevista_em, entrevista_local: a.entrevista_local, entrevista_resposta: a.entrevista_resposta, eventos });
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar o histórico.' }); }
});
// A empresa abriu o currículo: marca "em análise" (uma vez só) e o candidato fica sabendo.
app.post('/api/applications/:id/visto', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const { a, erro } = await candidaturaComAcesso(req, req.params.id); if (erro) return res.status(erro[0]).json({ error: erro[1] });
        if (!a.visto_empresa_em) {
            db.run(`UPDATE job_applications SET visto_empresa_em = CURRENT_TIMESTAMP, status = CASE WHEN COALESCE(status, 'recebida') = 'recebida' THEN 'em_analise' ELSE status END, status_em = CASE WHEN COALESCE(status, 'recebida') = 'recebida' THEN CURRENT_TIMESTAMP ELSE status_em END WHERE id = ?`, [a.id], () => {});
            if ((a.status || 'recebida') === 'recebida') {
                await registrarEventoCandidatura(a.id, 'sistema', 'status', `${a.empresaNome} está analisando o seu currículo.`, req.user.userId);
                notificar(a.candidate_user_id, 'Seu currículo está em análise 👀', `${a.empresaNome} abriu o seu currículo para a vaga "${a.vagaTitulo}".`, 'portalCandidaturas');
            }
        }
        res.json({ ok: true });
    } catch (e) { res.status(500).json({ error: 'Erro.' }); }
});
app.post('/api/applications/:id/acao', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const { a, erro } = await candidaturaComAcesso(req, req.params.id); if (erro) return res.status(erro[0]).json({ error: erro[1] });
        const b = req.body || {}, tipo = String(b.tipo || ''), msg = String(b.mensagem || '').trim();
        const upd = (sql, params) => new Promise((ok, er) => db.run(sql, params, e => e ? er(e) : ok()));
        if (tipo === 'entrevista') {
            if (!/^\d{4}-\d{2}-\d{2}$/.test(b.data || '') || !/^\d{2}:\d{2}$/.test(b.hora || '')) return res.status(400).json({ error: 'Escolha a data e a hora da entrevista.' });
            const quando = `${b.data}T${b.hora}`, local = String(b.local || '').trim().slice(0, 300);
            if (!local) return res.status(400).json({ error: 'Informe o local ou o link da entrevista.' });
            await upd(`UPDATE job_applications SET status = 'entrevista', status_em = CURRENT_TIMESTAMP, entrevista_em = ?, entrevista_local = ?, entrevista_resposta = NULL, visto_empresa_em = COALESCE(visto_empresa_em, CURRENT_TIMESTAMP) WHERE id = ?`, [quando, local, a.id]);
            const dataBr = b.data.split('-').reverse().join('/');
            await registrarEventoCandidatura(a.id, 'empresa', 'entrevista', `📅 Entrevista marcada para ${dataBr} às ${b.hora}\n📍 ${local}${msg ? '\n\n' + msg : ''}`, req.user.userId);
            notificar(a.candidate_user_id, 'Entrevista marcada! 📅', `${a.empresaNome} marcou uma entrevista para "${a.vagaTitulo}" em ${dataBr} às ${b.hora}.`, 'portalCandidaturas');
            avisarCandidatoPorEmail(a, `Entrevista marcada — ${a.vagaTitulo}`, `<p>Olá, ${String(a.candNome).split(' ')[0]}!</p><p><b>${a.empresaNome}</b> marcou uma entrevista com você para a vaga <b>${a.vagaTitulo}</b>.</p><p>📅 <b>${dataBr} às ${b.hora}</b><br>📍 ${local}</p>${msg ? `<p>${msg.replace(/</g, '&lt;').replace(/\n/g, '<br>')}</p>` : ''}<p>Confirme sua presença no Portal de Vagas Impulsionar, em "Minhas Candidaturas".</p>`);
            return res.json({ message: 'Entrevista marcada! O candidato foi avisado.' });
        }
        if (tipo === 'recusa') {
            const texto = msg || 'Agradecemos o seu interesse. Analisamos o seu currículo e, neste momento, seguiremos com outros candidatos para esta vaga. Deixaremos o seu currículo salvo para próximas oportunidades.';
            await upd(`UPDATE job_applications SET status = 'recusada', status_em = CURRENT_TIMESTAMP, visto_empresa_em = COALESCE(visto_empresa_em, CURRENT_TIMESTAMP) WHERE id = ?`, [a.id]);
            await registrarEventoCandidatura(a.id, 'empresa', 'recusa', texto, req.user.userId);
            notificar(a.candidate_user_id, 'Atualização da sua candidatura', `${a.empresaNome} respondeu sobre a vaga "${a.vagaTitulo}".`, 'portalCandidaturas');
            avisarCandidatoPorEmail(a, `Retorno da candidatura — ${a.vagaTitulo}`, `<p>Olá, ${String(a.candNome).split(' ')[0]}!</p><p>${texto.replace(/</g, '&lt;').replace(/\n/g, '<br>')}</p><p>— ${a.empresaNome}</p>`);
            return res.json({ message: 'Candidatura recusada. O candidato recebeu a mensagem.' });
        }
        if (tipo === 'aprovado') {
            await upd(`UPDATE job_applications SET status = 'aprovado', status_em = CURRENT_TIMESTAMP WHERE id = ?`, [a.id]);
            await registrarEventoCandidatura(a.id, 'empresa', 'status', `🎉 Parabéns! Você foi aprovado(a) no processo seletivo.${msg ? '\n\n' + msg : ''}`, req.user.userId);
            notificar(a.candidate_user_id, 'Você foi aprovado(a)! 🎉', `${a.empresaNome} aprovou você na vaga "${a.vagaTitulo}".`, 'portalCandidaturas');
            return res.json({ message: 'Candidato marcado como aprovado e avisado.' });
        }
        if (tipo === 'reabrir') {
            await upd(`UPDATE job_applications SET status = 'em_analise', status_em = CURRENT_TIMESTAMP WHERE id = ?`, [a.id]);
            await registrarEventoCandidatura(a.id, 'sistema', 'status', 'Sua candidatura voltou para análise.', req.user.userId);
            return res.json({ message: 'Candidatura voltou para análise.' });
        }
        if (tipo === 'mensagem') {
            if (!msg) return res.status(400).json({ error: 'Escreva a mensagem.' });
            await registrarEventoCandidatura(a.id, 'empresa', 'mensagem', msg, req.user.userId);
            notificar(a.candidate_user_id, `Mensagem de ${a.empresaNome} 💬`, msg.slice(0, 140), 'portalCandidaturas');
            return res.json({ message: 'Mensagem enviada ao candidato.' });
        }
        res.status(400).json({ error: 'Ação inválida.' });
    } catch (e) { res.status(400).json({ error: 'Erro ao registrar a ação.' }); }
});
// Candidato responde: mensagem livre ou confirmação/remarcação da entrevista.
app.post('/api/portal/applications/:id/responder', requireRole('candidate'), async (req, res) => {
    try {
        const { a, erro } = await candidaturaComAcesso(req, req.params.id); if (erro) return res.status(erro[0]).json({ error: erro[1] });
        const b = req.body || {}, msg = String(b.mensagem || '').trim().slice(0, 2000);
        let texto = msg;
        if (b.entrevista === 'confirmado' || b.entrevista === 'remarcar') {
            if (a.status !== 'entrevista') return res.status(400).json({ error: 'Não há entrevista marcada.' });
            db.run(`UPDATE job_applications SET entrevista_resposta = ? WHERE id = ?`, [b.entrevista, a.id], () => {});
            texto = (b.entrevista === 'confirmado' ? '✅ Presença confirmada na entrevista.' : '🔁 Preciso remarcar a entrevista.') + (msg ? '\n\n' + msg : '');
        }
        if (!texto) return res.status(400).json({ error: 'Escreva a mensagem.' });
        await registrarEventoCandidatura(a.id, 'candidato', b.entrevista ? 'resposta' : 'mensagem', texto, req.user.userId);
        notificarPorCompanyAdmins(a.company_id, `${a.candNome} respondeu 💬`, `Vaga "${a.vagaTitulo}": ${texto.slice(0, 120)}`, 'jobPostings');
        res.json({ message: 'Enviado para a empresa!' });
    } catch (e) { res.status(400).json({ error: 'Erro ao enviar.' }); }
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
app.get('/api/portal/candidates', requireRole('admin'), async (req, res) => {
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
        let query = `SELECT jp.*, vp.label as planLabel, vp.days as planDays, vp.price as planPrice, c.name as companyName, c.logo_url as companyLogo, COALESCE(c.vaga_credito_dias, 0) as companyCredito,
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

// Tipos de vaga usados para direcionar as vagas ao perfil do candidato.
const TIPOS_VAGA = ['Operacional', 'Administrativo', 'Técnico', 'Comercial', 'Liderança', 'Estágio / Aprendiz'];
function categoriasDaVaga(v) {
    if (v.categoria && TIPOS_VAGA.includes(v.categoria)) return [v.categoria];
    const t = ' ' + String((v.title || '') + ' ' + (v.seniority || '')).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase() + ' ';
    const c = [];
    if (/lideranca|coordenacao|gerencia|diretoria|gerente|supervis|coorden|lider|encarreg|diretor|gestor|chefe|head /.test(t)) c.push('Liderança');
    if (/estagi|aprendiz|trainee/.test(t)) c.push('Estágio / Aprendiz');
    if (/vend|comercial|promotor|represent|televend|key account|executivo de contas/.test(t)) c.push('Comercial');
    if (/tecnic|mecanic|eletric|manutenc|engenhe|desenvolv|programad| ti |sistemas|dados|soldador/.test(t)) c.push('Técnico');
    if (/administr|financ|contab|fiscal| rh |recursos humanos|pessoal|analista|assistente|recep|secret|faturist|compras|juridic/.test(t)) c.push('Administrativo');
    if (/operacional|motorista|ajudante|conferent|operador|empilhadeira|estoqu|armaz|separador|repositor|entregador|auxiliar|expedi|almox|vigilante|porteiro|limpeza|produc|logistic/.test(t) && !c.includes('Liderança')) c.push('Operacional');
    if (!c.length) c.push('Operacional');
    return c;
}
function salvarCamposExtrasVaga(id, b) {
    db.run(`UPDATE job_postings SET work_schedule = ?, pcd = ?, contract_type = ?, categoria = ? WHERE id = ?`,
        [String(b.work_schedule || '').trim().slice(0, 150), b.pcd ? 1 : 0, String(b.contract_type || '').trim().slice(0, 40), TIPOS_VAGA.includes(b.categoria) ? b.categoria : null, id], () => {});
}
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
            salvarCamposExtrasVaga(resultado, req.body);
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

        salvarCamposExtrasVaga(resultado, req.body);
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
            registrarCreditoVaga(vaga.company_id, -dias, 'consumido', `Vaga "${vaga.title}" aprovada e publicada com crédito`, req.user.userId, vaga.id);
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
        salvarCamposExtrasVaga(req.params.id, req.body);
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
            `SELECT cp.*, u.id as userId, u.name, u.email, ja.id, ja.applied_at, COALESCE(ja.status, 'recebida') as appStatus, ja.status_em, ja.entrevista_em, ja.entrevista_local, ja.entrevista_resposta, ja.visto_empresa_em,
                    (SELECT COUNT(*) FROM application_events ev WHERE ev.application_id = ja.id AND ev.autor = 'candidato' AND ev.lido_empresa = 0) as naoLidas
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

/* ==========================================================
   TESTE DE PERFIL DISC (ranking por pergunta) — o próprio
   executivo/colaborador (role 'autonomous') responde, liberado pelo
   Master/empresa via o módulo 'discTest' em employees.enabled_modules.
   ========================================================== */
const ROTULOS_PERFIL_DISC = { D: 'Dominância', I: 'Influência', S: 'Estabilidade', C: 'Conformidade' };

app.get('/api/disc-test/questions', requireRole('autonomous', 'admin', 'candidate'), (req, res) => {
    res.json(DISC_DATA);
});

// Master: envia (libera) o Teste DISC para um funcionário específico ou para
// todos os funcionários de uma empresa. Só mexe em quem já tem uma restrição
// de módulos definida (array) — quem já tem acesso completo (null) já
// enxerga o módulo sem precisar de envio.
app.post('/api/disc-test/enviar', requireRole('admin'), async (req, res) => {
    const { employee_id, company_id } = req.body;
    if (!employee_id && !company_id) return res.status(400).json({ error: 'Informe um funcionário ou uma empresa.' });
    try {
        const alvos = employee_id
            ? await dbAll(`SELECT id, enabled_modules FROM employees WHERE id = ?`, [employee_id])
            : await dbAll(`SELECT id, enabled_modules FROM employees WHERE company_id = ?`, [company_id]);
        if (!alvos.length) return res.status(404).json({ error: 'Nenhum executivo encontrado.' });
        for (const emp of alvos) {
            let modulos;
            try { modulos = emp.enabled_modules ? JSON.parse(emp.enabled_modules) : null; } catch (e) { modulos = null; }
            if (Array.isArray(modulos) && !modulos.includes('discTest')) {
                modulos.push('discTest');
                await new Promise((resolve, reject) => db.run(`UPDATE employees SET enabled_modules = ? WHERE id = ?`, [JSON.stringify(modulos), emp.id], (err) => err ? reject(err) : resolve()));
            }
            notificarPorEmployeeId(emp.id, 'Teste de Perfil DISC liberado', 'Você tem um novo Teste de Perfil DISC disponível para responder.', 'discTeste');
        }
        res.json({ message: `Teste DISC enviado para ${alvos.length} colaborador(es)!` });
    } catch (e) { res.status(400).json({ error: 'Erro ao enviar o teste DISC.' }); }
});

// Master: lista os resultados já respondidos, para acompanhar quem já fez o teste.
app.get('/api/disc-test/resultados', requireRole('admin'), async (req, res) => {
    const resultados = await dbAll(
        `SELECT r.id, r.employee_id, r.perfil_primario, r.perfil_secundario, r.created_at,
                e.name as employeeName, c.name as companyName
         FROM disc_results r
         JOIN employees e ON e.id = r.employee_id
         LEFT JOIN companies c ON c.id = e.company_id
         ORDER BY r.created_at DESC LIMIT 100`
    );
    const cands = await dbAll(`SELECT u.name as employeeName, 'Candidato do Portal' as companyName, cp.disc_perfil, cp.disc_em as created_at, cp.disc_liberado, cp.disc_liberado_em
        FROM candidate_profiles cp JOIN users u ON u.id = cp.user_id WHERE cp.disc_perfil IS NOT NULL AND cp.disc_perfil <> '' ORDER BY cp.disc_em DESC LIMIT 100`).catch(() => []);
    const pendentes = await dbAll(`SELECT u.id, u.name, cp.disc_liberado_em FROM candidate_profiles cp JOIN users u ON u.id = cp.user_id WHERE cp.disc_liberado = 1 ORDER BY cp.disc_liberado_em DESC`).catch(() => []);
    const lista = [...resultados, ...cands.map(c => ({ ...c, perfil_primario: c.disc_perfil.split('/')[0], perfil_secundario: c.disc_perfil.split('/')[1], candidato: true }))]
        .sort((a, b) => String(b.created_at || '').localeCompare(String(a.created_at || '')));
    if (req.query.comPendentes === '1') return res.json({ resultados: lista, pendentesCandidatos: pendentes });
    res.json(lista);
});

function calcularDiscServidor(respostas) {
    if (!Array.isArray(respostas) || respostas.length !== DISC_DATA.length) return { erro: 'É preciso responder todas as perguntas do teste.' };
    const pontos = { D: 0, I: 0, S: 0, C: 0 }, pesos = [4, 3, 2, 1];
    for (const resp of respostas) {
        const ordem = Array.isArray(resp.ordem) ? resp.ordem : [];
        if (ordem.length !== 4 || new Set(ordem).size !== 4) return { erro: 'Cada pergunta precisa ranquear as 4 frases, sem repetição.' };
        ordem.forEach((dim, i) => { if (pontos[dim] !== undefined) pontos[dim] += pesos[i] || 0; });
    }
    const max = DISC_DATA.length * 4;
    const percentuais = { D: Math.round(pontos.D / max * 100), I: Math.round(pontos.I / max * 100), S: Math.round(pontos.S / max * 100), C: Math.round(pontos.C / max * 100) };
    const ranking = ['D', 'I', 'S', 'C'].sort((a, b) => pontos[b] - pontos[a]);
    return { pontos, percentuais, primario: ranking[0], secundario: ranking[1] };
}
function resultadoDiscCandidato(cp) {
    if (!cp || !cp.disc_perfil) return null;
    let j = {}; try { j = JSON.parse(cp.disc_json || '{}'); } catch (e) {}
    const [p1, p2] = String(cp.disc_perfil).split('/');
    return { percentuais: j.percentuais || {}, perfil_primario: p1, perfil_secundario: p2, perfilPrimarioLabel: ROTULOS_PERFIL_DISC[p1], perfilSecundarioLabel: ROTULOS_PERFIL_DISC[p2], created_at: cp.disc_em, candidato: true };
}
// Candidatos do Portal de Vagas: o Master libera o teste (uso único); o resultado vai para o currículo padrão.
app.post('/api/disc-test/enviar-candidatos', requireRole('admin'), async (req, res) => {
    try {
        const { user_id, todos } = req.body || {};
        if (!user_id && !todos) return res.status(400).json({ error: 'Escolha o candidato (ou "todos").' });
        const alvos = user_id ? await dbAll(`SELECT u.id FROM users u WHERE u.id = ? AND u.role = 'candidate'`, [user_id])
            : await dbAll(`SELECT u.id FROM users u JOIN candidate_profiles cp ON cp.user_id = u.id WHERE u.role = 'candidate' AND (cp.disc_perfil IS NULL OR cp.disc_perfil = '')`);
        if (!alvos.length) return res.status(404).json({ error: todos ? 'Todos os candidatos já fizeram o teste.' : 'Candidato não encontrado.' });
        for (const a of alvos) {
            await new Promise(ok => db.run(`INSERT OR IGNORE INTO candidate_profiles (user_id) VALUES (?)`, [a.id], () => ok()));
            await new Promise(ok => db.run(`UPDATE candidate_profiles SET disc_liberado = 1, disc_liberado_em = CURRENT_TIMESTAMP WHERE user_id = ?`, [a.id], () => ok()));
            notificar(a.id, 'Teste de Perfil DISC liberado 🧠', 'Faça o teste (leva uns 10 minutos). O resultado aparece no seu currículo para as empresas.', 'discTeste');
        }
        res.json({ message: `Teste DISC enviado para ${alvos.length} candidato(s)!` });
    } catch (e) { res.status(400).json({ error: 'Erro ao enviar o teste DISC.' }); }
});
app.get('/api/disc-test/meu-resultado', requireRole('autonomous', 'candidate'), async (req, res, next) => {
    if (req.user.role !== 'candidate') return next();
    const cp = await dbGet(`SELECT disc_perfil, disc_json, disc_em, disc_liberado FROM candidate_profiles WHERE user_id = ?`, [req.user.userId]);
    const r = resultadoDiscCandidato(cp);
    res.json(r ? { ...r, liberado: !!(cp && cp.disc_liberado) } : (cp && cp.disc_liberado ? null : { naoLiberado: true }));
});
app.post('/api/disc-test/submit', requireRole('autonomous', 'candidate'), async (req, res, next) => {
    if (req.user.role !== 'candidate') return next();
    try {
        const cp = await dbGet(`SELECT disc_liberado FROM candidate_profiles WHERE user_id = ?`, [req.user.userId]);
        if (!cp || !cp.disc_liberado) return res.status(403).json({ error: 'O Teste DISC ainda não foi liberado para você.' });
        const r = calcularDiscServidor(req.body.respostas); if (r.erro) return res.status(400).json({ error: r.erro });
        await new Promise((ok, er) => db.run(`UPDATE candidate_profiles SET disc_perfil = ?, disc_json = ?, disc_em = CURRENT_TIMESTAMP, disc_liberado = 0 WHERE user_id = ?`,
            [`${r.primario}/${r.secundario}`, JSON.stringify({ pontos: r.pontos, percentuais: r.percentuais }), req.user.userId], e => e ? er(e) : ok()));
        res.json({ message: 'Teste DISC concluído!', percentuais: r.percentuais, perfil_primario: r.primario, perfil_secundario: r.secundario,
            perfilPrimarioLabel: ROTULOS_PERFIL_DISC[r.primario], perfilSecundarioLabel: ROTULOS_PERFIL_DISC[r.secundario], created_at: new Date().toISOString(), candidato: true });
    } catch (e) { res.status(400).json({ error: 'Erro ao salvar o teste.' }); }
});
app.get('/api/disc-test/meu-resultado', requireRole('autonomous', 'candidate'), async (req, res) => {
    const resultado = await dbGet(
        `SELECT * FROM disc_results WHERE employee_id = ? ORDER BY id DESC LIMIT 1`,
        [req.user.employeeId]
    );
    if (!resultado) return res.json(null);
    try { resultado.respostas = JSON.parse(resultado.respostas || '[]'); } catch (e) { resultado.respostas = []; }
    const maximoPorDimensao = DISC_DATA.length * 4;
    resultado.percentuais = {
        D: Math.round((resultado.pontos_d / maximoPorDimensao) * 100),
        I: Math.round((resultado.pontos_i / maximoPorDimensao) * 100),
        S: Math.round((resultado.pontos_s / maximoPorDimensao) * 100),
        C: Math.round((resultado.pontos_c / maximoPorDimensao) * 100)
    };
    resultado.perfilPrimarioLabel = ROTULOS_PERFIL_DISC[resultado.perfil_primario];
    resultado.perfilSecundarioLabel = ROTULOS_PERFIL_DISC[resultado.perfil_secundario];
    res.json(resultado);
});

app.post('/api/disc-test/submit', requireRole('autonomous', 'candidate'), async (req, res) => {
    const { respostas } = req.body; // [{ perguntaId, ordem: ['D','I','S','C'] em ordem do que MAIS combina para o que MENOS combina }]
    if (!Array.isArray(respostas) || respostas.length !== DISC_DATA.length) {
        return res.status(400).json({ error: 'É preciso responder todas as perguntas do teste.' });
    }
    const pontos = { D: 0, I: 0, S: 0, C: 0 };
    const pesos = [4, 3, 2, 1]; // 1º lugar (mais combina) = 4 pontos, ... 4º lugar (menos combina) = 1 ponto
    for (const resp of respostas) {
        const ordem = Array.isArray(resp.ordem) ? resp.ordem : [];
        if (ordem.length !== 4 || new Set(ordem).size !== 4) {
            return res.status(400).json({ error: 'Cada pergunta precisa ranquear as 4 frases, sem repetição.' });
        }
        ordem.forEach((dim, i) => {
            if (pontos[dim] === undefined) return;
            pontos[dim] += pesos[i] || 0;
        });
    }
    const maximoPorDimensao = DISC_DATA.length * 4; // 24 perguntas * 4 pontos no melhor caso
    const percentuais = {
        D: Math.round((pontos.D / maximoPorDimensao) * 100),
        I: Math.round((pontos.I / maximoPorDimensao) * 100),
        S: Math.round((pontos.S / maximoPorDimensao) * 100),
        C: Math.round((pontos.C / maximoPorDimensao) * 100)
    };
    const ranking = ['D', 'I', 'S', 'C'].sort((a, b) => pontos[b] - pontos[a]);
    const perfilPrimario = ranking[0];
    const perfilSecundario = ranking[1];

    db.run(
        `INSERT INTO disc_results (employee_id, respostas, pontos_d, pontos_i, pontos_s, pontos_c, perfil_primario, perfil_secundario)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
        [req.user.employeeId, JSON.stringify(respostas), pontos.D, pontos.I, pontos.S, pontos.C, perfilPrimario, perfilSecundario],
        async function (err) {
            if (err) return res.status(400).json({ error: err.message });
            const perfilTexto = `${perfilPrimario}/${perfilSecundario}`;
            db.run(`UPDATE employees SET disc_profile = ? WHERE id = ?`, [perfilTexto, req.user.employeeId], () => {});
            // Uso único: quem tem o módulo por uma restrição explícita (array)
            // perde o acesso de volta ao concluir — só volta a ficar disponível
            // se o Master/empresa "enviar" (liberar) de novo.
            try {
                const emp = await dbGet(`SELECT enabled_modules FROM employees WHERE id = ?`, [req.user.employeeId]);
                if (emp) {
                    let modulos;
                    try { modulos = emp.enabled_modules ? JSON.parse(emp.enabled_modules) : null; } catch (e) { modulos = null; }
                    if (Array.isArray(modulos) && modulos.includes('discTest')) {
                        db.run(`UPDATE employees SET enabled_modules = ? WHERE id = ?`,
                            [JSON.stringify(modulos.filter(m => m !== 'discTest')), req.user.employeeId], () => {});
                    }
                }
            } catch (e) { /* não bloqueia a resposta do teste por causa disso */ }
            res.json({
                message: 'Teste DISC concluído!',
                id: this.lastID,
                employee_id: req.user.employeeId,
                pontos_d: pontos.D, pontos_i: pontos.I, pontos_s: pontos.S, pontos_c: pontos.C,
                percentuais,
                perfil_primario: perfilPrimario, perfil_secundario: perfilSecundario,
                perfilPrimarioLabel: ROTULOS_PERFIL_DISC[perfilPrimario],
                perfilSecundarioLabel: ROTULOS_PERFIL_DISC[perfilSecundario],
                created_at: new Date().toISOString()
            });
        }
    );
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
    const { employee_id, objective, action_plan, deadline, status, photo_url, final_delivery } = req.body;
    db.run(`INSERT INTO pd_plans (employee_id, objective, action_plan, deadline, status, photo_url, final_delivery) VALUES (?, ?, ?, ?, ?, ?, ?)`,
        [employee_id, objective, action_plan, deadline, status || 'Em Andamento', photo_url || null, final_delivery || null], () => {
        notificarPorEmployeeId(employee_id, 'Novo PDI criado', objective, 'pdi');
        res.json({ message: 'PDI criado!' });
    });
});
app.put('/api/pdi/:id', requireRole('admin', 'client_admin'), ensureRecordAccess('pd_plans'), async (req, res) => {
    const { objective, action_plan, deadline, status, photo_url, final_delivery } = req.body;
    try {
        const atual = await dbGet(`SELECT * FROM pd_plans WHERE id = ?`, [req.params.id]);
        if (!atual) return res.status(404).json({ error: 'PDI não encontrado.' });
        db.run(
            `UPDATE pd_plans SET objective = ?, action_plan = ?, deadline = ?, status = ?, photo_url = ?, final_delivery = ? WHERE id = ?`,
            [
                objective !== undefined ? objective : atual.objective,
                action_plan !== undefined ? action_plan : atual.action_plan,
                deadline !== undefined ? deadline : atual.deadline,
                status !== undefined ? status : atual.status,
                photo_url !== undefined ? photo_url : atual.photo_url,
                final_delivery !== undefined ? final_delivery : atual.final_delivery,
                req.params.id
            ],
            (err) => {
                if (err) return res.status(400).json({ error: err.message });
                if (status === 'Concluído' && atual.status !== 'Concluído') {
                    darPontos(req.user.userId, 50, 'PDI concluído');
                    notificarPorEmployeeId(atual.employee_id, 'PDI concluído', objective || atual.objective, 'pdi');
                }
                res.json({ message: 'PDI atualizado!' });
            }
        );
    } catch (e) { res.status(400).json({ error: 'Erro ao atualizar o PDI.' }); }
});
app.delete('/api/pdi/:id', requireRole('admin', 'client_admin'), ensureRecordAccess('pd_plans'), (req, res) => {
    db.run(`DELETE FROM pdi_updates WHERE pd_plan_id = ?`, [req.params.id], () => {});
    db.run(`DELETE FROM pdi_materials WHERE pd_plan_id = ?`, [req.params.id], () => {});
    db.run(`DELETE FROM pdi_actions WHERE pd_plan_id = ?`, [req.params.id], () => {});
    db.run(`DELETE FROM pd_plans WHERE id = ?`, [req.params.id], () => res.json({ message: 'Removido!' }));
});

// Ações de um PDI: quebram o plano em passos concretos (não iniciada / em
// andamento / concluída) em vez de um status único — dá a visão de "quantas
// ações estão em andamento, não iniciadas e fechadas" por colaborador.
app.get('/api/pdi/:id/actions', ensureRecordAccess('pd_plans'), (req, res) => {
    db.all(`SELECT * FROM pdi_actions WHERE pd_plan_id = ? ORDER BY id ASC`, [req.params.id], (err, rows) => res.json(rows || []));
});

app.post('/api/pdi/:id/actions', requireRole('admin', 'client_admin'), ensureRecordAccess('pd_plans'), (req, res) => {
    const { description, foto_url } = req.body;
    if (!description) return res.status(400).json({ error: 'Descreva a ação.' });
    db.run(`INSERT INTO pdi_actions (pd_plan_id, description, status, foto_url) VALUES (?, ?, 'Não iniciada', ?)`, [req.params.id, description, foto_url || null], function (err) {
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

app.put('/api/pdi-actions/:id', requireRole('admin', 'client_admin'), ensurePdiActionAccess(), async (req, res) => {
    const { description, status, foto_url } = req.body;
    try {
        const atual = await dbGet(`SELECT * FROM pdi_actions WHERE id = ?`, [req.params.id]);
        if (!atual) return res.status(404).json({ error: 'Ação não encontrada.' });
        db.run(`UPDATE pdi_actions SET description = ?, status = ?, foto_url = ? WHERE id = ?`,
            [
                description !== undefined ? description : atual.description,
                status !== undefined ? status : atual.status,
                foto_url !== undefined ? foto_url : atual.foto_url,
                req.params.id
            ],
            (err) => {
                if (err) return res.status(400).json({ error: err.message });
                res.json({ message: 'Ação atualizada!' });
            });
    } catch (e) { res.status(400).json({ error: 'Erro ao atualizar a ação.' }); }
});

app.delete('/api/pdi-actions/:id', requireRole('admin', 'client_admin'), ensurePdiActionAccess(), (req, res) => {
    db.run(`DELETE FROM pdi_actions WHERE id = ?`, [req.params.id], () => res.json({ message: 'Removida!' }));
});

// Valida acesso a um registro (pdi_updates ou pdi_materials) a partir do
// pd_plan_id que ele carrega, no mesmo espírito de ensurePdiActionAccess.
function ensurePdiRelatedAccess(table) {
    return async (req, res, next) => {
        if (req.user.role === 'admin') return next();
        try {
            const row = await dbGet(`SELECT pd_plan_id FROM ${table} WHERE id = ?`, [req.params.id]);
            if (!row) return res.status(404).json({ error: 'Registro não encontrado.' });
            const plano = await dbGet(`SELECT employee_id FROM pd_plans WHERE id = ?`, [row.pd_plan_id]);
            if (!plano) return res.status(404).json({ error: 'PDI não encontrado.' });
            return ensureEmployeeAccess(() => plano.employee_id)(req, res, next);
        } catch (e) { return res.status(500).json({ error: 'Erro ao validar permissão.' }); }
    };
}

// Evolução do PDI: vários follow-ups ao longo do acompanhamento, cada um
// podendo trazer uma foto de evidência — histórico do andamento do cliente.
app.get('/api/pdi/:id/updates', ensureRecordAccess('pd_plans'), (req, res) => {
    db.all(`SELECT * FROM pdi_updates WHERE pd_plan_id = ? ORDER BY id DESC`, [req.params.id], (err, rows) => res.json(rows || []));
});

app.post('/api/pdi/:id/updates', requireRole('admin', 'client_admin'), ensureRecordAccess('pd_plans'), async (req, res) => {
    const { texto, foto_url } = req.body;
    if (!texto) return res.status(400).json({ error: 'Descreva a evolução/atualização.' });
    try {
        const plano = await dbGet(`SELECT employee_id FROM pd_plans WHERE id = ?`, [req.params.id]);
        const autor = await dbGet(`SELECT name FROM users WHERE id = ?`, [req.user.userId]);
        db.run(`INSERT INTO pdi_updates (pd_plan_id, texto, foto_url, created_by_name) VALUES (?, ?, ?, ?)`,
            [req.params.id, texto, foto_url || null, (autor && autor.name) || null], function (err) {
                if (err) return res.status(400).json({ error: err.message });
                if (plano) notificarPorEmployeeId(plano.employee_id, 'Atualização no seu PDI', texto, 'pdi');
                res.json({ message: 'Evolução registrada!', id: this.lastID });
            });
    } catch (e) { res.status(400).json({ error: 'Erro ao registrar a evolução.' }); }
});

app.delete('/api/pdi-updates/:id', requireRole('admin', 'client_admin'), ensurePdiRelatedAccess('pdi_updates'), (req, res) => {
    db.run(`DELETE FROM pdi_updates WHERE id = ?`, [req.params.id], () => res.json({ message: 'Removida!' }));
});

// Materiais de apoio do PDI: arquivos, livros e links indicados ao executivo.
app.get('/api/pdi/:id/materials', ensureRecordAccess('pd_plans'), (req, res) => {
    db.all(`SELECT * FROM pdi_materials WHERE pd_plan_id = ? ORDER BY id DESC`, [req.params.id], (err, rows) => res.json(rows || []));
});

app.post('/api/pdi/:id/materials', requireRole('admin', 'client_admin'), ensureRecordAccess('pd_plans'), (req, res) => {
    const { titulo, tipo, url, nota } = req.body;
    if (!titulo) return res.status(400).json({ error: 'Dê um título ao material.' });
    db.run(`INSERT INTO pdi_materials (pd_plan_id, titulo, tipo, url, nota) VALUES (?, ?, ?, ?, ?)`,
        [req.params.id, titulo, tipo || 'arquivo', url || null, nota || null], function (err) {
            if (err) return res.status(400).json({ error: err.message });
            res.json({ message: 'Material adicionado!', id: this.lastID });
        });
});

app.delete('/api/pdi-materials/:id', requireRole('admin', 'client_admin'), ensurePdiRelatedAccess('pdi_materials'), (req, res) => {
    db.run(`DELETE FROM pdi_materials WHERE id = ?`, [req.params.id], () => res.json({ message: 'Removido!' }));
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
    const viaApi = !!(EMAIL_API.provedor && EMAIL_API.chave);
    if (!viaApi && (!process.env.SMTP_HOST || !process.env.SMTP_USER || !process.env.SMTP_PASS)) {
        return res.status(400).json({ error: 'Nenhum envio configurado: cadastre o envio por API (Brevo ou Resend) aqui em Meu Perfil, ou defina SMTP_HOST, SMTP_USER e SMTP_PASS no servidor.' });
    }
    try {
        const info = await transporter.sendMail({
            from: process.env.SMTP_FROM || `"Impulsionar V4" <${process.env.SMTP_USER || EMAIL_API.remetente}>`,
            to: destino,
            subject: 'Teste de envio — Impulsionar V4',
            html: `<p>Se você recebeu este e-mail, o SMTP configurado no <strong>.env</strong> está funcionando corretamente.</p>`
        });
        res.json({ message: `E-mail de teste enviado para ${destino} via ${viaApi ? (EMAIL_API.provedor === 'resend' ? 'Resend (API)' : 'Brevo (API)') : 'SMTP'}! Confira a caixa de entrada (e o spam).`, id: info.messageId });
    } catch (e) {
        res.status(400).json({ error: `Falha ao enviar: ${e.message}` });
    }
});

// Configuração do envio de e-mail por API (Brevo / Resend) — feita pelo Master na tela Meu Perfil.
app.get('/api/platform/email-api', requireRole('admin'), (req, res) => {
    res.json({ provedor: EMAIL_API.provedor || '', remetente: EMAIL_API.remetente || '', chavePreview: EMAIL_API.chave ? '••••••••' + EMAIL_API.chave.slice(-5) : null, smtpConfigurado: !!(process.env.SMTP_HOST && process.env.SMTP_USER && process.env.SMTP_PASS), smtpHost: process.env.SMTP_HOST || '', smtpPorta: SMTP_PORTA });
});
app.put('/api/platform/email-api', requireRole('admin'), async (req, res) => {
    try {
        const provedor = String(req.body.provedor || '').toLowerCase();
        if (!provedor) {
            await new Promise(r => db.run(`DELETE FROM integration_settings WHERE key IN ('email_api_provedor', 'email_api_chave', 'email_api_remetente')`, [], () => r()));
            EMAIL_API = { provedor: '', chave: '', remetente: '' };
            return res.json({ message: 'Envio por API desligado — os e-mails voltam a usar o SMTP.' });
        }
        if (!['brevo', 'resend'].includes(provedor)) return res.status(400).json({ error: 'Provedor inválido.' });
        const chave = String(req.body.chave || '').trim() || (EMAIL_API.provedor === provedor ? EMAIL_API.chave : '');
        const remetente = String(req.body.remetente || '').trim();
        if (!chave) return res.status(400).json({ error: 'Informe a chave de API.' });
        if (!/@/.test(remetente)) return res.status(400).json({ error: 'Informe o e-mail remetente (ex.: Impulsionar <contato@seudominio.com.br>).' });
        const salvar = (k, v) => new Promise((ok, erro) => db.run(`INSERT INTO integration_settings (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value`, [k, v], e => e ? erro(e) : ok()));
        await salvar('email_api_provedor', provedor); await salvar('email_api_chave', chave); await salvar('email_api_remetente', remetente);
        EMAIL_API = { provedor, chave, remetente };
        res.json({ message: 'Envio por API salvo! Use "Testar Envio" para confirmar.' });
    } catch (e) { res.status(400).json({ error: 'Erro ao salvar a configuração de e-mail.' }); }
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

// ======================================================================
// DPO — AGENDA DE CONSULTORIAS com convite no Outlook / Teams
//  • Sempre: e-mail com convite de calendário (.ics) para o consultor e para
//    os usuários da empresa — o Outlook/Teams mostra "Aceitar" e coloca na agenda.
//  • Se a integração Microsoft 365 estiver configurada (Azure: tenant, client id,
//    secret e a caixa organizadora), o evento é criado direto no Outlook pelo
//    Microsoft Graph com reunião do Teams gerada automaticamente; a própria
//    Microsoft envia os convites e as atualizações/cancelamentos.
// ======================================================================
const MS365_CHAVES = ['ms365_tenant', 'ms365_client_id', 'ms365_client_secret', 'ms365_organizador'];
async function lerConfigMs365() {
    const rows = await dbAll(`SELECT key, value FROM integration_settings WHERE key IN (${MS365_CHAVES.map(() => '?').join(',')})`, MS365_CHAVES);
    const m = Object.fromEntries(rows.map(r => [r.key, r.value]));
    return {
        tenant: m.ms365_tenant || process.env.MS365_TENANT_ID || '',
        clientId: m.ms365_client_id || process.env.MS365_CLIENT_ID || '',
        secret: m.ms365_client_secret || process.env.MS365_CLIENT_SECRET || '',
        organizador: m.ms365_organizador || process.env.MS365_ORGANIZER || ''
    };
}
function ms365Ativo(c) { return !!(c && c.tenant && c.clientId && c.secret && c.organizador); }
let CACHE_TOKEN_MS365 = { token: '', expira: 0, chave: '' };
async function tokenMs365(c) {
    const chave = c.tenant + '|' + c.clientId;
    if (CACHE_TOKEN_MS365.token && CACHE_TOKEN_MS365.chave === chave && Date.now() < CACHE_TOKEN_MS365.expira) return CACHE_TOKEN_MS365.token;
    const corpo = new URLSearchParams({ client_id: c.clientId, client_secret: c.secret, scope: 'https://graph.microsoft.com/.default', grant_type: 'client_credentials' });
    const r = await fetch(`https://login.microsoftonline.com/${encodeURIComponent(c.tenant)}/oauth2/v2.0/token`, { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: corpo });
    const j = await r.json().catch(() => ({}));
    if (!r.ok || !j.access_token) throw new Error('Microsoft 365: não foi possível autenticar (' + (j.error_description || j.error || r.status) + '). Confira Tenant, Client ID e Secret.');
    CACHE_TOKEN_MS365 = { token: j.access_token, expira: Date.now() + ((j.expires_in || 3600) - 120) * 1000, chave };
    return j.access_token;
}
async function graphMs365(c, metodo, caminho, corpo) {
    const token = await tokenMs365(c);
    const r = await fetch('https://graph.microsoft.com/v1.0' + caminho, { method: metodo, headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json', Prefer: 'outlook.timezone="America/Sao_Paulo"' }, body: corpo ? JSON.stringify(corpo) : undefined });
    if (r.status === 204 || r.status === 202) return {};
    const j = await r.json().catch(() => ({}));
    if (!r.ok) throw new Error('Microsoft 365: ' + ((j.error && j.error.message) || ('erro ' + r.status)) + (r.status === 403 ? ' — dê ao app a permissão de aplicativo "Calendars.ReadWrite" com consentimento do administrador.' : ''));
    return j;
}

const HORA_RE = /^([01]\d|2[0-3]):[0-5]\d$/;
// Brasil (America/Sao_Paulo) está fixo em UTC-3 desde 2019.
function dataHoraUtcDpo(data, hora) { return new Date(`${data}T${hora}:00-03:00`); }
function icsData(d) { return d.toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, ''); }
function icsTexto(t) { return String(t || '').replace(/\\/g, '\\\\').replace(/;/g, '\;').replace(/,/g, '\\,').replace(/\r?\n/g, '\\n'); }
function icsDobrar(linha) { const out = []; let l = linha; while (l.length > 73) { out.push(l.slice(0, 73)); l = ' ' + l.slice(73); } out.push(l); return out.join('\r\n'); }

async function montarSessaoDpo(id) {
    const s = await dbGet(`SELECT s.*, c.name as companyName, k.name as consultorNome, k.email as consultorEmail, k.phone as consultorTelefone
        FROM dpo_sessoes s LEFT JOIN companies c ON c.id = s.company_id LEFT JOIN dpo_consultants k ON k.id = s.consultant_id WHERE s.id = ?`, [id]);
    if (!s) return null;
    let participantes = []; try { participantes = JSON.parse(s.participantes || '[]'); } catch (e) { participantes = []; }
    let pilares = []; try { pilares = JSON.parse(s.pilares || '[]'); } catch (e) { pilares = []; }
    return { ...s, participantes, pilares, pilaresLabel: pilares.map(p => (DPO_AMBEV_DATA[p] || {}).label || p) };
}

// Convidados = consultor + e-mails escolhidos (usuários da empresa e extras), sem repetir.
function convidadosSessaoDpo(s) {
    const lista = [];
    const add = (email, nome, papel) => { const e = String(email || '').trim().toLowerCase(); if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e) && !lista.some(x => x.email === e)) lista.push({ email: e, nome: nome || '', papel }); };
    if (s.consultorEmail) add(s.consultorEmail, s.consultorNome, 'consultor');
    (s.participantes || []).forEach(p => add(p.email, p.nome, p.papel || 'empresa'));
    return lista;
}

function descricaoSessaoDpo(s, linkTeams) {
    return [
        `Consultoria DPO Ambev — ${s.companyName || ''}`,
        s.pilaresLabel && s.pilaresLabel.length ? `Pilar(es): ${s.pilaresLabel.join(', ')}` : '',
        s.consultorNome ? `Consultor: ${s.consultorNome}${s.consultorTelefone ? ' · ' + s.consultorTelefone : ''}` : '',
        `Data: ${s.data.split('-').reverse().join('/')} · ${s.hora_inicio} às ${s.hora_fim} (horário de Brasília)`,
        linkTeams ? `Entrar na reunião do Teams: ${linkTeams}` : (s.formato === 'presencial' && s.local ? `Local: ${s.local}` : ''),
        s.observacao ? `\n${s.observacao}` : '',
        '\nAgendado pela plataforma Impulsionar V4.'
    ].filter(Boolean).join('\n');
}

function gerarIcsSessaoDpo(s, metodo) {
    const ini = dataHoraUtcDpo(s.data, s.hora_inicio), fim = dataHoraUtcDpo(s.data, s.hora_fim);
    const org = separarRemetenteEmail(EMAIL_API.remetente || process.env.SMTP_FROM || process.env.SMTP_USER || '').email || 'agenda@impulsionarv4.com.br';
    const local = s.teams_link ? 'Microsoft Teams' : (s.formato === 'presencial' ? (s.local || 'Presencial') : (s.local || 'Online'));
    const linhas = [
        'BEGIN:VCALENDAR', 'PRODID:-//Impulsionar V4//Agenda DPO//PT-BR', 'VERSION:2.0', 'CALSCALE:GREGORIAN', 'METHOD:' + metodo,
        'BEGIN:VEVENT',
        'UID:' + s.uid,
        'SEQUENCE:' + (s.sequencia || 0),
        'DTSTAMP:' + icsData(new Date()),
        'DTSTART:' + icsData(ini),
        'DTEND:' + icsData(fim),
        'SUMMARY:' + icsTexto(s.titulo),
        'DESCRIPTION:' + icsTexto(descricaoSessaoDpo(s, s.teams_link)),
        'LOCATION:' + icsTexto(local),
        s.teams_link ? 'URL:' + s.teams_link : '',
        'ORGANIZER;CN=Impulsionar V4:mailto:' + org,
        ...convidadosSessaoDpo(s).map(c => `ATTENDEE;CN=${icsTexto(c.nome || c.email)};ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE:mailto:${c.email}`),
        'STATUS:' + (metodo === 'CANCEL' ? 'CANCELLED' : 'CONFIRMED'),
        'BEGIN:VALARM', 'TRIGGER:-PT30M', 'ACTION:DISPLAY', 'DESCRIPTION:Consultoria DPO em 30 minutos', 'END:VALARM',
        'END:VEVENT', 'END:VCALENDAR'
    ].filter(Boolean);
    return linhas.map(icsDobrar).join('\r\n') + '\r\n';
}

function htmlConviteSessaoDpo(s, tipo) {
    const cab = tipo === 'cancel' ? '❌ Consultoria cancelada' : tipo === 'update' ? '🔄 Consultoria reagendada / atualizada' : '📅 Nova consultoria agendada';
    const linha = (r, v) => v ? `<tr><td style="padding:4px 12px 4px 0;color:#64748b;">${r}</td><td style="padding:4px 0;font-weight:600;">${v}</td></tr>` : '';
    const esc = t => String(t || '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
    return `<div style="font-family:Arial,sans-serif;max-width:560px;margin:auto;border:1px solid #e2e8f0;border-radius:12px;overflow:hidden;">
      <div style="background:#0b1324;color:#fff;padding:16px 20px;font-size:18px;font-weight:700;">${cab}</div>
      <div style="padding:18px 20px;color:#1e293b;">
        <div style="font-size:16px;font-weight:700;margin-bottom:10px;">${esc(s.titulo)}</div>
        <table style="font-size:14px;border-collapse:collapse;">
          ${linha('Empresa', esc(s.companyName))}
          ${linha('Data', esc(s.data.split('-').reverse().join('/')))}
          ${linha('Horário', esc(`${s.hora_inicio} às ${s.hora_fim} (Brasília)`))}
          ${linha('Consultor', esc(s.consultorNome))}
          ${linha('Pilar(es)', esc((s.pilaresLabel || []).join(', ')))}
          ${s.formato === 'presencial' ? linha('Local', esc(s.local)) : ''}
        </table>
        ${s.observacao ? `<p style="font-size:14px;white-space:pre-line;">${esc(s.observacao)}</p>` : ''}
        ${tipo !== 'cancel' && s.teams_link ? `<p style="margin:18px 0;"><a href="${esc(s.teams_link)}" style="background:#5b5fc7;color:#fff;padding:11px 18px;border-radius:8px;text-decoration:none;font-weight:700;">🎥 Entrar na reunião do Teams</a></p>` : ''}
        ${tipo !== 'cancel' ? '<p style="font-size:12.5px;color:#64748b;">O convite vai anexo (convite.ics). No Outlook/Teams clique em <strong>Aceitar</strong> para colocar na sua agenda.</p>' : '<p style="font-size:12.5px;color:#64748b;">Abra o anexo para remover o compromisso da sua agenda.</p>'}
      </div></div>`;
}

// Envia/atualiza/cancela o convite. tipo: 'novo' | 'update' | 'cancel'
async function sincronizarSessaoDpo(id, tipo) {
    const s = await montarSessaoDpo(id);
    if (!s) return { ok: false, erro: 'Agendamento não encontrado.' };
    const convidados = convidadosSessaoDpo(s);
    const cfg = await lerConfigMs365();
    let via = 'email', aviso = '';
    if (ms365Ativo(cfg)) {
        try {
            const base = `/users/${encodeURIComponent(cfg.organizador)}/events`;
            if (tipo === 'cancel') {
                if (s.graph_event_id) await graphMs365(cfg, 'POST', `${base}/${encodeURIComponent(s.graph_event_id)}/cancel`, { comment: 'Consultoria cancelada pela Impulsionar.' });
            } else {
                const corpo = {
                    subject: s.titulo,
                    body: { contentType: 'HTML', content: descricaoSessaoDpo(s, null).replace(/\n/g, '<br>') },
                    start: { dateTime: `${s.data}T${s.hora_inicio}:00`, timeZone: 'America/Sao_Paulo' },
                    end: { dateTime: `${s.data}T${s.hora_fim}:00`, timeZone: 'America/Sao_Paulo' },
                    location: { displayName: s.formato === 'presencial' ? (s.local || 'Presencial') : 'Microsoft Teams' },
                    attendees: convidados.map(c => ({ emailAddress: { address: c.email, name: c.nome || c.email }, type: 'required' })),
                    reminderMinutesBeforeStart: 30, isReminderOn: true
                };
                if (s.formato !== 'presencial') { corpo.isOnlineMeeting = true; corpo.onlineMeetingProvider = 'teamsForBusiness'; }
                let ev;
                if (s.graph_event_id) ev = await graphMs365(cfg, 'PATCH', `${base}/${encodeURIComponent(s.graph_event_id)}`, corpo);
                else ev = await graphMs365(cfg, 'POST', base, corpo);
                const link = (ev.onlineMeeting && ev.onlineMeeting.joinUrl) || s.teams_link || null;
                await new Promise(r => db.run(`UPDATE dpo_sessoes SET graph_event_id = COALESCE(?, graph_event_id), teams_link = ? WHERE id = ?`, [ev.id || null, link, s.id], () => r()));
            }
            via = 'microsoft365';
        } catch (e) {
            aviso = e.message + ' — enviei o convite por e-mail (.ics) no lugar.';
        }
    }
    if (via === 'email') {
        const s2 = await montarSessaoDpo(id);
        if (!convidados.length) return { ok: false, via, erro: 'Nenhum e-mail para convidar (cadastre o e-mail do consultor ou escolha os participantes).' };
        const metodo = tipo === 'cancel' ? 'CANCEL' : 'REQUEST';
        const assunto = (tipo === 'cancel' ? 'Cancelada: ' : tipo === 'update' ? 'Atualizada: ' : 'Convite: ') + s2.titulo + ` — ${s2.data.split('-').reverse().join('/')} ${s2.hora_inicio}`;
        try {
            await transporter.sendMail({
                from: process.env.SMTP_FROM || `"Impulsionar V4" <${process.env.SMTP_USER || EMAIL_API.remetente}>`,
                to: convidados.map(c => c.email).join(', '),
                subject: assunto,
                html: htmlConviteSessaoDpo(s2, tipo === 'novo' ? 'novo' : tipo),
                icalEvent: { filename: 'convite.ics', method: metodo, content: gerarIcsSessaoDpo(s2, metodo) }
            });
        } catch (e) { return { ok: false, via, erro: 'Agendamento salvo, mas o e-mail não saiu: ' + e.message }; }
    }
    await new Promise(r => db.run(`UPDATE dpo_sessoes SET ultimo_envio = ? WHERE id = ?`, [`${via}|${new Date().toISOString()}|${convidados.length}`, s.id], () => r()));
    // Aviso dentro da plataforma para os gestores da empresa.
    const titulo = tipo === 'cancel' ? '❌ Consultoria cancelada' : tipo === 'update' ? '🔄 Consultoria atualizada' : '📅 Consultoria agendada';
    notificarPorCompanyAdmins(s.company_id, titulo, `${s.titulo} — ${s.data.split('-').reverse().join('/')} das ${s.hora_inicio} às ${s.hora_fim}${s.consultorNome ? ' com ' + s.consultorNome : ''}.`, 'dpoHome');
    return { ok: true, via, convidados: convidados.length, aviso };
}

function validarSessaoDpo(b) {
    const erros = [];
    if (!b.company_id) erros.push('Escolha a empresa.');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(b.data || ''))) erros.push('Escolha a data.');
    if (!HORA_RE.test(String(b.hora_inicio || ''))) erros.push('Informe a hora de início.');
    if (!HORA_RE.test(String(b.hora_fim || ''))) erros.push('Informe a hora de término.');
    if (!erros.length && b.hora_fim <= b.hora_inicio) erros.push('A hora de término precisa ser depois do início.');
    if (b.teams_link && !/^https?:\/\/\S+$/i.test(String(b.teams_link).trim())) erros.push('Link do Teams inválido.');
    return erros;
}
function participantesLimposDpo(lista) {
    return (Array.isArray(lista) ? lista : []).map(p => ({ email: String(p.email || '').trim().toLowerCase().slice(0, 200), nome: String(p.nome || '').trim().slice(0, 120), papel: p.papel === 'extra' ? 'extra' : 'empresa' }))
        .filter(p => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(p.email)).slice(0, 40);
}

app.get('/api/admin/dpo/ms365', requireRole('admin'), async (req, res) => {
    const c = await lerConfigMs365();
    res.json({ ativo: ms365Ativo(c), tenant: c.tenant, clientId: c.clientId, organizador: c.organizador, secretPreview: c.secret ? '••••••' + c.secret.slice(-4) : null });
});
app.put('/api/admin/dpo/ms365', requireRole('admin'), async (req, res) => {
    try {
        const atual = await lerConfigMs365();
        const novo = {
            ms365_tenant: String(req.body.tenant || '').trim(),
            ms365_client_id: String(req.body.clientId || '').trim(),
            ms365_client_secret: String(req.body.secret || '').trim() || (req.body.limpar ? '' : atual.secret),
            ms365_organizador: String(req.body.organizador || '').trim()
        };
        if (req.body.limpar) Object.keys(novo).forEach(k => novo[k] = '');
        for (const [k, v] of Object.entries(novo)) await new Promise(r => db.run(`INSERT INTO integration_settings (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value`, [k, v], () => r()));
        CACHE_TOKEN_MS365 = { token: '', expira: 0, chave: '' };
        const c = await lerConfigMs365();
        if (ms365Ativo(c) && req.body.testar !== false) {
            try { await graphMs365(c, 'GET', `/users/${encodeURIComponent(c.organizador)}/calendar`); }
            catch (e) { return res.status(400).json({ error: 'Salvo, mas o teste falhou: ' + e.message }); }
            return res.json({ message: 'Microsoft 365 conectado! Os agendamentos vão direto para o Outlook com link do Teams.' });
        }
        res.json({ message: req.body.limpar ? 'Integração removida — convites seguem por e-mail (.ics).' : 'Configuração salva.' });
    } catch (e) { res.status(400).json({ error: 'Erro ao salvar a integração.' }); }
});

// E-mails da empresa para convidar (gestores e acessos da empresa).
app.get('/api/admin/dpo/sessoes/participantes', requireRole('admin'), async (req, res) => {
    try {
        const users = await dbAll(`SELECT name, email, role FROM users WHERE company_id = ? AND email IS NOT NULL AND email <> '' AND role IN ('client_admin', 'autonomous', 'employee') ORDER BY CASE role WHEN 'client_admin' THEN 0 ELSE 1 END, name`, [req.query.company_id]);
        res.json(users.map(u => ({ nome: u.name, email: u.email, papel: u.role === 'client_admin' ? 'gestor' : 'acesso' })));
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar os participantes.' }); }
});

app.get('/api/dpo/sessoes', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const params = []; let filtro = '';
        if (req.user.role === 'client_admin') { filtro = 'WHERE s.company_id = ?'; params.push(req.user.companyId); }
        else if (req.query.company_id) { filtro = 'WHERE s.company_id = ?'; params.push(req.query.company_id); }
        const ids = await dbAll(`SELECT s.id FROM dpo_sessoes s ${filtro} ORDER BY s.data DESC, s.hora_inicio DESC LIMIT 300`, params);
        const lista = [];
        for (const r of ids) {
            const s = await montarSessaoDpo(r.id);
            if (req.user.role !== 'admin') { delete s.graph_event_id; delete s.ultimo_envio; }
            lista.push(s);
        }
        res.json(lista);
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar a agenda.' }); }
});

app.post('/api/admin/dpo/sessoes', requireRole('admin'), async (req, res) => {
    const b = req.body || {};
    const erros = validarSessaoDpo(b);
    if (erros.length) return res.status(400).json({ error: erros[0] });
    try {
        const emp = await dbGet(`SELECT id, name FROM companies WHERE id = ?`, [b.company_id]);
        if (!emp) return res.status(400).json({ error: 'Empresa inválida.' });
        const pilares = (Array.isArray(b.pilares) ? b.pilares : []).filter(p => DPO_PILARES_ORDEM.includes(p));
        const titulo = String(b.titulo || '').trim().slice(0, 200) || `Consultoria DPO — ${emp.name}`;
        const uid = `dpo-${Date.now()}-${crypto.randomBytes(5).toString('hex')}@impulsionarv4`;
        const id = await new Promise((ok, ko) => db.run(
            `INSERT INTO dpo_sessoes (company_id, consultant_id, titulo, pilares, data, hora_inicio, hora_fim, formato, local, teams_link, participantes, observacao, uid, created_by) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
            [emp.id, b.consultant_id || null, titulo, JSON.stringify(pilares), b.data, b.hora_inicio, b.hora_fim, b.formato === 'presencial' ? 'presencial' : 'teams', String(b.local || '').trim().slice(0, 300) || null, String(b.teams_link || '').trim() || null, JSON.stringify(participantesLimposDpo(b.participantes)), String(b.observacao || '').trim().slice(0, 2000) || null, uid, req.user.userId],
            function (err) { err ? ko(err) : ok(this.lastID); }));
        if (b.solicitacao_id) {
            const sol = await dbGet(`SELECT * FROM dpo_sessao_solicitacoes WHERE id = ? AND company_id = ?`, [b.solicitacao_id, emp.id]);
            if (sol) await new Promise(r => db.run(`UPDATE dpo_sessao_solicitacoes SET status = 'agendada', sessao_id = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?`, [id, sol.id], () => r()));
        }
        const r = await sincronizarSessaoDpo(id, 'novo');
        res.json({ id, message: r.ok ? `Consultoria agendada! Convite enviado para ${r.convidados} e-mail(s)${r.via === 'microsoft365' ? ' pelo Outlook/Teams' : ''}.` : r.erro, aviso: r.aviso || '', enviado: r.ok });
    } catch (e) { res.status(400).json({ error: 'Erro ao agendar a consultoria.' }); }
});

app.put('/api/admin/dpo/sessoes/:id', requireRole('admin'), async (req, res) => {
    const b = req.body || {};
    try {
        const atual = await dbGet(`SELECT * FROM dpo_sessoes WHERE id = ?`, [req.params.id]);
        if (!atual) return res.status(404).json({ error: 'Agendamento não encontrado.' });
        const m = { ...atual, ...b, company_id: atual.company_id };
        const erros = validarSessaoDpo(m);
        if (erros.length) return res.status(400).json({ error: erros[0] });
        const pilares = b.pilares !== undefined ? (Array.isArray(b.pilares) ? b.pilares : []).filter(p => DPO_PILARES_ORDEM.includes(p)) : JSON.parse(atual.pilares || '[]');
        await new Promise((ok, ko) => db.run(
            `UPDATE dpo_sessoes SET consultant_id = ?, titulo = ?, pilares = ?, data = ?, hora_inicio = ?, hora_fim = ?, formato = ?, local = ?, teams_link = ?, participantes = ?, observacao = ?, status = 'agendada', sequencia = sequencia + 1, updated_at = CURRENT_TIMESTAMP WHERE id = ?`,
            [m.consultant_id || null, String(m.titulo || '').trim().slice(0, 200) || atual.titulo, JSON.stringify(pilares), m.data, m.hora_inicio, m.hora_fim, m.formato === 'presencial' ? 'presencial' : 'teams', String(m.local || '').trim().slice(0, 300) || null, String(m.teams_link || '').trim() || null,
             JSON.stringify(b.participantes !== undefined ? participantesLimposDpo(b.participantes) : JSON.parse(atual.participantes || '[]')), String(m.observacao || '').trim().slice(0, 2000) || null, atual.id],
            (err) => err ? ko(err) : ok()));
        const r = b.reenviar === false ? { ok: true, convidados: 0 } : await sincronizarSessaoDpo(atual.id, 'update');
        res.json({ message: r.ok ? `Agendamento atualizado${r.convidados ? ' — convite atualizado para ' + r.convidados + ' e-mail(s)' : ''}.` : r.erro, aviso: r.aviso || '' });
    } catch (e) { res.status(400).json({ error: 'Erro ao atualizar o agendamento.' }); }
});

app.post('/api/admin/dpo/sessoes/:id/reenviar', requireRole('admin'), async (req, res) => {
    try {
        const r = await sincronizarSessaoDpo(req.params.id, 'update');
        if (!r.ok) return res.status(400).json({ error: r.erro });
        res.json({ message: `Convite reenviado para ${r.convidados} e-mail(s).`, aviso: r.aviso || '' });
    } catch (e) { res.status(400).json({ error: 'Erro ao reenviar o convite.' }); }
});

app.post('/api/admin/dpo/sessoes/:id/status', requireRole('admin'), async (req, res) => {
    const st = ['agendada', 'realizada', 'cancelada'].includes(req.body.status) ? req.body.status : null;
    if (!st) return res.status(400).json({ error: 'Status inválido.' });
    try {
        const atual = await dbGet(`SELECT * FROM dpo_sessoes WHERE id = ?`, [req.params.id]);
        if (!atual) return res.status(404).json({ error: 'Agendamento não encontrado.' });
        await new Promise(r => db.run(`UPDATE dpo_sessoes SET status = ?, sequencia = sequencia + 1, updated_at = CURRENT_TIMESTAMP WHERE id = ?`, [st, atual.id], () => r()));
        let msg = 'Status atualizado.';
        if (st === 'agendada' && atual.status === 'cancelada') {
            const r = await sincronizarSessaoDpo(atual.id, 'update');
            msg = r.ok ? 'Consultoria reativada — o convite foi enviado de novo para as agendas.' : 'Reativada, mas ' + r.erro;
        } else if (st === 'agendada') msg = 'Voltou para agendada.';
        if (st === 'cancelada' && atual.status !== 'cancelada') {
            const r = await sincronizarSessaoDpo(atual.id, 'cancel');
            msg = r.ok ? 'Consultoria cancelada — o cancelamento foi enviado para as agendas.' : 'Cancelada, mas ' + r.erro;
        }
        res.json({ message: msg });
    } catch (e) { res.status(400).json({ error: 'Erro ao alterar o status.' }); }
});

app.delete('/api/admin/dpo/sessoes/:id', requireRole('admin'), async (req, res) => {
    try {
        const atual = await dbGet(`SELECT * FROM dpo_sessoes WHERE id = ?`, [req.params.id]);
        if (!atual) return res.status(404).json({ error: 'Agendamento não encontrado.' });
        if (atual.status === 'agendada' && atual.data >= new Date(Date.now() - 3 * 3600 * 1000).toISOString().slice(0, 10)) {
            await new Promise(r => db.run(`UPDATE dpo_sessoes SET sequencia = sequencia + 1 WHERE id = ?`, [atual.id], () => r()));
            await sincronizarSessaoDpo(atual.id, 'cancel').catch(() => {});
        }
        await new Promise(r => db.run(`DELETE FROM dpo_sessoes WHERE id = ?`, [atual.id], () => r()));
        res.json({ message: 'Agendamento excluído.' });
    } catch (e) { res.status(400).json({ error: 'Erro ao excluir.' }); }
});

// ----- Solicitações de consultoria (empresa pede, Master agenda) -----
app.post('/api/dpo/solicitacoes', requireRole('client_admin'), async (req, res) => {
    const b = req.body || {};
    const assunto = String(b.assunto || '').trim().slice(0, 1500);
    if (!assunto) return res.status(400).json({ error: 'Conte o que precisa tratar na consultoria.' });
    if (b.data_sugerida && !/^\d{4}-\d{2}-\d{2}$/.test(b.data_sugerida)) return res.status(400).json({ error: 'Data inválida.' });
    if (b.hora_sugerida && !HORA_RE.test(b.hora_sugerida)) return res.status(400).json({ error: 'Hora inválida.' });
    if (b.hora_fim_sugerida && !HORA_RE.test(b.hora_fim_sugerida)) return res.status(400).json({ error: 'Hora final inválida.' });
    if (b.hora_sugerida && b.hora_fim_sugerida && b.hora_fim_sugerida <= b.hora_sugerida) return res.status(400).json({ error: 'A hora final precisa ser depois do início.' });
    try {
        const pend = await dbGet(`SELECT COUNT(*) n FROM dpo_sessao_solicitacoes WHERE company_id = ? AND status = 'pendente'`, [req.user.companyId]);
        if (pend && pend.n >= 5) return res.status(400).json({ error: 'Você já tem 5 solicitações aguardando o Master. Aguarde o retorno.' });
        const pilares = (Array.isArray(b.pilares) ? b.pilares : []).filter(p => DPO_PILARES_ORDEM.includes(p));
        const id = await new Promise((ok, ko) => db.run(
            `INSERT INTO dpo_sessao_solicitacoes (company_id, user_id, pilares, data_sugerida, hora_sugerida, hora_fim_sugerida, formato, assunto) VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
            [req.user.companyId, req.user.userId, JSON.stringify(pilares), b.data_sugerida || null, b.hora_sugerida || null, b.hora_fim_sugerida || null, b.formato === 'presencial' ? 'presencial' : 'teams', assunto],
            function (err) { err ? ko(err) : ok(this.lastID); }));
        const emp = await dbGet(`SELECT name FROM companies WHERE id = ?`, [req.user.companyId]);
        const admins = await dbAll(`SELECT id FROM users WHERE role = 'admin'`);
        admins.forEach(a => notificar(a.id, '📅 Solicitação de consultoria DPO', `${emp ? emp.name : 'Empresa'} pediu uma consultoria${b.data_sugerida ? ' para ' + b.data_sugerida.split('-').reverse().join('/') + (b.hora_sugerida ? ' às ' + b.hora_sugerida : '') : ''}.`, 'dpoAgenda'));
        res.json({ id, message: 'Solicitação enviada! A Impulsionar vai agendar e você recebe o convite no seu e-mail/Outlook.' });
    } catch (e) { res.status(400).json({ error: 'Erro ao enviar a solicitação.' }); }
});

app.get('/api/dpo/solicitacoes', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const params = []; let filtro = '';
        if (req.user.role === 'client_admin') { filtro = 'WHERE s.company_id = ?'; params.push(req.user.companyId); }
        const l = await dbAll(`SELECT s.*, c.name as companyName, u.name as solicitanteNome, u.email as solicitanteEmail
            FROM dpo_sessao_solicitacoes s LEFT JOIN companies c ON c.id = s.company_id LEFT JOIN users u ON u.id = s.user_id
            ${filtro} ORDER BY CASE s.status WHEN 'pendente' THEN 0 ELSE 1 END, s.created_at DESC LIMIT 200`, params);
        res.json(l.map(x => { let p = []; try { p = JSON.parse(x.pilares || '[]'); } catch (e) {} return { ...x, pilares: p, pilaresLabel: p.map(k => (DPO_AMBEV_DATA[k] || {}).label || k) }; }));
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar as solicitações.' }); }
});

app.post('/api/admin/dpo/solicitacoes/:id/recusar', requireRole('admin'), async (req, res) => {
    try {
        const s = await dbGet(`SELECT * FROM dpo_sessao_solicitacoes WHERE id = ?`, [req.params.id]);
        if (!s) return res.status(404).json({ error: 'Solicitação não encontrada.' });
        const resposta = String(req.body.resposta || '').trim().slice(0, 1000) || null;
        await new Promise(r => db.run(`UPDATE dpo_sessao_solicitacoes SET status = 'recusada', resposta = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?`, [resposta, s.id], () => r()));
        if (s.user_id) notificar(s.user_id, 'Solicitação de consultoria', resposta ? 'A Impulsionar respondeu: ' + resposta : 'Sua solicitação de consultoria não pôde ser atendida nessa data. Fale com a Impulsionar.', 'dpoHome');
        res.json({ message: 'Solicitação respondida.' });
    } catch (e) { res.status(400).json({ error: 'Erro ao responder.' }); }
});

app.delete('/api/dpo/solicitacoes/:id', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const s = await dbGet(`SELECT * FROM dpo_sessao_solicitacoes WHERE id = ?`, [req.params.id]);
        if (!s) return res.status(404).json({ error: 'Solicitação não encontrada.' });
        if (req.user.role === 'client_admin' && (String(s.company_id) !== String(req.user.companyId) || s.status !== 'pendente')) return res.status(403).json({ error: 'Só dá para cancelar solicitações ainda pendentes.' });
        await new Promise(r => db.run(`DELETE FROM dpo_sessao_solicitacoes WHERE id = ?`, [s.id], () => r()));
        res.json({ message: req.user.role === 'client_admin' ? 'Solicitação cancelada.' : 'Solicitação removida.' });
    } catch (e) { res.status(400).json({ error: 'Erro ao remover.' }); }
});

// Arquivo .ics para "Adicionar à minha agenda" (Outlook, Google, Apple).
app.get('/api/dpo/sessoes/:id/ics', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const s = await montarSessaoDpo(req.params.id);
        if (!s) return res.status(404).json({ error: 'Agendamento não encontrado.' });
        if (req.user.role === 'client_admin' && String(s.company_id) !== String(req.user.companyId)) return res.status(403).json({ error: 'Sem acesso.' });
        res.set('Content-Type', 'text/calendar; charset=utf-8');
        res.set('Content-Disposition', `attachment; filename="consultoria-dpo-${s.data}.ics"`);
        res.send(gerarIcsSessaoDpo(s, s.status === 'cancelada' ? 'CANCEL' : 'PUBLISH'));
    } catch (e) { res.status(500).json({ error: 'Erro ao gerar o convite.' }); }
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

// Interrompe um trial que está rodando ANTES do prazo combinado — só marca a
// expiração para agora (não apaga a linha, fica no histórico de quem já teve
// esse pilar liberado em algum momento).
app.post('/api/admin/dpo/trial/:id/cancelar', requireRole('admin'), async (req, res) => {
    try {
        const trial = await dbGet(`SELECT * FROM dpo_purchases WHERE id = ? AND is_trial = 1`, [req.params.id]);
        if (!trial) return res.status(404).json({ error: 'Período de teste não encontrado.' });
        const empresa = await dbGet(`SELECT name FROM companies WHERE id = ?`, [trial.company_id]);
        await new Promise((resolve, reject) => db.run(
            `UPDATE dpo_purchases SET trial_expires_at = CURRENT_TIMESTAMP WHERE id = ?`,
            [req.params.id],
            (err) => err ? reject(err) : resolve()
        ));
        const rotuloEscopo = trial.scope === 'completo' ? 'a Consultoria Completa' : `o pilar ${DPO_AMBEV_DATA[trial.pillar_key] ? DPO_AMBEV_DATA[trial.pillar_key].label : trial.pillar_key}`;
        notificarGestoresDaEmpresa(trial.company_id, 'DPO Ambev — período de teste interrompido', `O Master encerrou antes do prazo o período de teste de ${rotuloEscopo}.`);
        res.json({ message: `Período de teste de ${empresa ? empresa.name : 'empresa'} interrompido!` });
    } catch (e) { res.status(400).json({ error: 'Erro ao interromper o período de teste.' }); }
});

// Exclui de vez o registro do trial (some do histórico também). Diferente do
// "Interromper", que só marca o prazo como já encerrado, isto apaga a linha.
app.delete('/api/admin/dpo/trial/:id', requireRole('admin'), async (req, res) => {
    try {
        const trial = await dbGet(`SELECT * FROM dpo_purchases WHERE id = ? AND is_trial = 1`, [req.params.id]);
        if (!trial) return res.status(404).json({ error: 'Período de teste não encontrado.' });
        await new Promise((resolve, reject) => db.run(`DELETE FROM dpo_purchases WHERE id = ?`, [req.params.id], (err) => err ? reject(err) : resolve()));
        res.json({ message: 'Período de teste excluído!' });
    } catch (e) { res.status(400).json({ error: 'Erro ao excluir o período de teste.' }); }
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
        const empresas = await dbAll(`SELECT id, name, dpo_auditoria_oficial_data, dpo_auditoria_oficial_nota, dpo_primeira_auditoria FROM companies ORDER BY name ASC`);
        const resultado = [];
        for (const emp of empresas) {
            const ativos = await pilaresAtivosDaEmpresa(emp.id);
            if (ativos.length || req.query.all) {
                const trials = await dbAll(`SELECT id, scope, pillar_key, trial_expires_at FROM dpo_purchases WHERE company_id = ? AND status = 'paid' AND is_trial = 1 AND trial_expires_at > datetime('now') ORDER BY trial_expires_at ASC`, [emp.id]);
                resultado.push({
                    id: emp.id, name: emp.name, pilaresAtivos: ativos,
                    compraCompleta: ativos.length === DPO_PILARES_ORDEM.length,
                    auditoriaOficialData: emp.dpo_auditoria_oficial_data || null,
                    auditoriaOficialNota: emp.dpo_auditoria_oficial_nota || null,
                    trialsAtivos: trials.map(t => ({ id: t.id, scope: t.scope, pillarKey: t.pillar_key, expiraEm: t.trial_expires_at })),
                    pastas: await pastasLiberadasDaEmpresa(emp.id),
                    primeiraAuditoria: !!emp.dpo_primeira_auditoria
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
// Master exclui um ciclo (histórico) com as notas e os planos de ação ligados a ele.
app.delete('/api/admin/dpo/cycles/:id', requireRole('admin'), async (req, res) => {
    try {
        const ciclo = await dbGet(`SELECT id FROM dpo_audit_cycles WHERE id = ?`, [req.params.id]);
        if (!ciclo) return res.status(404).json({ error: 'Ciclo não encontrado.' });
        for (const q of [`DELETE FROM dpo_action_plans WHERE cycle_id = ?`, `DELETE FROM dpo_answers WHERE cycle_id = ?`, `DELETE FROM dpo_audit_cycles WHERE id = ?`]) {
            await new Promise((resolve, reject) => db.run(q, [ciclo.id], (err) => err ? reject(err) : resolve()));
        }
        res.json({ message: 'Ciclo excluído do histórico.' });
    } catch (e) {
        console.error('Erro ao excluir ciclo DPO:', e.message);
        res.status(400).json({ error: 'Erro ao excluir o ciclo.' });
    }
});

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
            let pergunta = null;
            if (pilarInfo) {
                for (const g of pilarInfo.grupos) {
                    const achou = g.perguntas.find(q => q.numero === numeroPergunta);
                    if (achou) { pergunta = achou; break; }
                }
            }
            if (!porPilar[pilarKey]) porPilar[pilarKey] = { key: pilarKey, label: pilarInfo ? pilarInfo.label : pilarKey, planos: [] };
            porPilar[pilarKey].planos.push({
                ...p,
                perguntaNumero: numeroPergunta,
                perguntaTexto: pergunta ? pergunta.questao : '',
                verificacaoTexto: pergunta ? pergunta.verificacao : '',
                howToCheck: pergunta ? pergunta.how_to_check : '',
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

// Resumo compacto pro "Painel da Empresa" (dashboard) — quantos planos de
// ação estão em cada situação de prazo e a média geral de pontuação do
// checklist, sem precisar a tela do dashboard buscar cada ciclo um por um.
app.get('/api/dpo/resumo', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const companyId = req.user.role === 'client_admin' ? req.user.companyId : (req.query.company_id || null);
        if (!companyId) return res.status(400).json({ error: 'Informe a empresa (company_id).' });
        const planos = await dbAll(`
            SELECT ap.status, ap.data_prevista
            FROM dpo_action_plans ap
            JOIN dpo_audit_cycles dac ON dac.id = ap.cycle_id
            WHERE dac.company_id = ?
        `, [companyId]);
        const hojeISO = new Date().toISOString().slice(0, 10);
        const acoes = { noPrazo: 0, vencida: 0, emAndamento: 0, concluida: 0 };
        planos.forEach(p => {
            if (p.status === 'concluida') { acoes.concluida++; return; }
            if (p.data_prevista && p.data_prevista.slice(0, 10) < hojeISO) { acoes.vencida++; return; }
            if (p.status === 'em_andamento') { acoes.emAndamento++; return; }
            acoes.noPrazo++;
        });
        // Pontuação = última autoavaliação mensal (não mais o checklist dos ciclos).
        const niveis = await niveisAtuaisDaEmpresaDpo(companyId);
        const pilaresAtivos = await pilaresAtivosDaEmpresa(companyId);
        const geral = niveis.geral;
        res.json({
            totalPlanos: planos.length,
            acoes,
            mediaGeral: geral && geral.categorias.todos !== null ? String(geral.categorias.todos).replace('.', ',') + '%' : null,
            nivelGeral: geral ? geral.nivelGeral : null,
            nivelGeralLabel: geral && geral.nivelGeral ? REGUA_SELOS_DPO.find(x => x.key === geral.nivelGeral).label : null,
            referenciaLabel: geral ? geral.referenciaLabel : null,
            totalPerguntasRespondidas: 0,
            pilaresAtivos: pilaresAtivos.length
        });
    } catch (e) {
        console.error('Erro ao carregar resumo DPO:', e.message);
        res.status(500).json({ error: 'Erro ao carregar o resumo do DPO.' });
    }
});

// Baixa em Excel (.xlsx) todos os planos de ação da empresa, consolidados por
// pilar — mesma base de dados de "Meus Planos DPO", só que num arquivo pra
// levar pra reunião/enviar por e-mail em vez de ficar só na tela.
app.get('/api/dpo/export/planos', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const companyId = req.user.role === 'client_admin' ? req.user.companyId : (req.query.company_id || null);
        if (!companyId) return res.status(400).json({ error: 'Informe a empresa (company_id).' });
        const empresa = await dbGet(`SELECT name FROM companies WHERE id = ?`, [companyId]);
        const planos = await dbAll(`
            SELECT ap.*, dac.referencia as cicloReferencia
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

        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet('Planos de Ação');
        sheet.columns = [
            { header: 'Pilar', key: 'pilar', width: 22 },
            { header: 'Pergunta', key: 'pergunta', width: 50 },
            { header: 'Verificação', key: 'verificacao', width: 14 },
            { header: 'Plano de Ação', key: 'plano', width: 55 },
            { header: 'Dono', key: 'dono', width: 18 },
            { header: 'Status', key: 'status', width: 16 },
            { header: 'Prazo', key: 'prazo', width: 14 },
            { header: 'Ciclo', key: 'ciclo', width: 16 },
            { header: 'Follow-ups', key: 'follows', width: 70 }
        ];
        sheet.getRow(1).font = { bold: true };
        planos.forEach(p => {
            const [pilarKey, numeroPergunta] = String(p.question_key).split(':');
            const pilarInfo = DPO_AMBEV_DATA[pilarKey];
            const follows = followsPorPlano[p.id] || [];
            sheet.addRow({
                pilar: pilarInfo ? pilarInfo.label : pilarKey,
                pergunta: `${numeroPergunta}. ${textoDaPerguntaDpo(pilarKey, numeroPergunta)}`,
                verificacao: p.verificacao_numero || '',
                plano: p.texto || '',
                dono: p.owner || '',
                status: ROTULOS_STATUS_PLANO_DPO_SERVIDOR[p.status] || p.status || '',
                prazo: p.data_prevista ? new Date(p.data_prevista + 'T00:00:00').toLocaleDateString('pt-BR') : '',
                ciclo: p.cicloReferencia || '',
                follows: follows.map(f => `Follow ${f.numero}: ${f.texto}${f.data_prevista ? ' (' + new Date(f.data_prevista).toLocaleDateString('pt-BR') + ')' : ''} [${f.status}]`).join(' | ')
            });
        });
        const buffer = await workbook.xlsx.writeBuffer();
        const nomeArquivo = `planos-dpo-${(empresa ? empresa.name : 'empresa').replace(/[^a-z0-9]+/gi, '-')}.xlsx`;
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${nomeArquivo}"`);
        res.send(Buffer.from(buffer));
    } catch (e) {
        console.error('Erro ao exportar planos DPO em Excel:', e.message);
        res.status(500).json({ error: 'Erro ao exportar os planos em Excel.' });
    }
});

// Baixa em Excel (.xlsx) o checklist completo (todas as perguntas de todos os
// pilares já liberados) com a pontuação dada em cada ciclo já respondido.
// "Excel do Checklist": agora exporta a autoavaliação mensal mais recente
// (resumo por categoria/pilar/bloco + todas as notas 3 / 1 / 0 / N/A).
app.get('/api/dpo/export/checklist', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const companyId = req.user.role === 'client_admin' ? req.user.companyId : (req.query.company_id || null);
        if (!companyId) return res.status(400).json({ error: 'Informe a empresa (company_id).' });
        const av = await dbGet(`SELECT * FROM dpo_self_assessments WHERE company_id = ? ORDER BY referencia DESC LIMIT 1`, [companyId]);
        if (!av) return res.status(404).json({ error: 'Nenhuma autoavaliação mensal feita ainda.' });
        const { buffer, nome } = await gerarExcelAutoavaliacaoDpo(av);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${nome}"`);
        res.send(Buffer.from(buffer));
    } catch (e) {
        console.error('Erro ao exportar checklist DPO em Excel:', e.message);
        res.status(500).json({ error: 'Erro ao exportar em Excel.' });
    }
});

// ---------- DPO Ambev — pasta "Perguntas Bate-Papo" (por pilar) ----------
// Posição de cada pergunta oficial dentro do pilar (ordem do checklist), usada
// para ordenar as perguntas do bate-papo "por pergunta do pilar".
function ordemDasPerguntasDoPilarDpo(pilarKey) {
    const ordem = {};
    const pilarInfo = DPO_AMBEV_DATA[pilarKey];
    if (!pilarInfo) return ordem;
    let i = 0;
    pilarInfo.grupos.forEach(g => g.perguntas.forEach(q => { ordem[q.numero] = i++; }));
    return ordem;
}

// ---------- Permissão por PASTA do DPO (Master libera por empresa) ----------
const PASTAS_DPO = ['checklist', 'batepapo', 'material', 'ferramentas', 'autoavaliacao', 'gop'];
const PASTAS_DPO_PADRAO = { checklist: true, batepapo: false, material: false, ferramentas: false, autoavaliacao: true, gop: true };
const ROTULOS_PASTAS_DPO = { checklist: 'Gestão (Checklist)', batepapo: 'Conhecimento do Time', material: 'Material do Pilar', ferramentas: 'Ferramentas Impulsionar', autoavaliacao: 'Autoavaliação Mensal', gop: 'Gerenciador GOP' };

async function pastasLiberadasDaEmpresa(companyId) {
    const resultado = { ...PASTAS_DPO_PADRAO };
    const linhas = await dbAll(`SELECT folder_key, enabled FROM dpo_company_folders WHERE company_id = ?`, [companyId]);
    linhas.forEach(l => { if (PASTAS_DPO.includes(l.folder_key)) resultado[l.folder_key] = !!l.enabled; });
    return resultado;
}

// Resolve de qual empresa é a pasta: client_admin sempre a própria; Master
// informa ?company_id= (ou company_id no corpo). Para a empresa, também confere
// que o pilar está liberado e que o Master liberou ESTA pasta para ela.
// Retorna null (e já responde o erro) se não pode.
async function resolverEmpresaPastaDpo(req, res, pilarKey, companyIdInformado, pasta) {
    if (!DPO_PILARES_ORDEM.includes(pilarKey)) { res.status(400).json({ error: 'Pilar inválido.' }); return null; }
    const companyId = req.user.role === 'client_admin' ? req.user.companyId : companyIdInformado;
    if (!companyId) { res.status(400).json({ error: 'Informe a empresa (company_id).' }); return null; }
    if (req.user.role === 'client_admin') {
        const ativos = await pilaresAtivosDaEmpresa(companyId);
        if (!ativos.includes(pilarKey)) { res.status(403).json({ error: 'Sua empresa ainda não tem este pilar liberado.' }); return null; }
        if (pasta) {
            const pastas = await pastasLiberadasDaEmpresa(companyId);
            if (!pastas[pasta]) { res.status(403).json({ error: `A pasta "${ROTULOS_PASTAS_DPO[pasta]}" ainda não foi liberada para sua empresa. Fale com o Master.` }); return null; }
        }
    }
    return companyId;
}

async function resolverEmpresaBatePapoDpo(req, res, pilarKey, companyIdInformado) {
    return resolverEmpresaPastaDpo(req, res, pilarKey, companyIdInformado, 'batepapo');
}

// Garante que a empresa (client_admin) pode mexer na pasta de um registro já
// existente (edição/exclusão). Master sempre pode.
async function empresaTemPastaDpo(req, res, companyId, pasta) {
    if (req.user.role !== 'client_admin') return true;
    const pastas = await pastasLiberadasDaEmpresa(companyId);
    if (!pastas[pasta]) { res.status(403).json({ error: `A pasta "${ROTULOS_PASTAS_DPO[pasta]}" não está liberada para sua empresa.` }); return false; }
    return true;
}

app.get('/api/dpo/pastas', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        if (req.user.role === 'admin') {
            // Master enxerga todas as pastas (quando abre o ciclo de uma empresa),
            // mas também recebe o que está liberado para ela, se informar company_id.
            const liberadas = req.query.company_id ? await pastasLiberadasDaEmpresa(req.query.company_id) : null;
            return res.json({ checklist: true, batepapo: true, material: true, ferramentas: true, autoavaliacao: true, gop: true, liberadasParaEmpresa: liberadas });
        }
        res.json(await pastasLiberadasDaEmpresa(req.user.companyId));
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar as pastas liberadas.' }); }
});

app.put('/api/admin/dpo/pastas/:companyId', requireRole('admin'), async (req, res) => {
    try {
        const empresa = await dbGet(`SELECT id, name FROM companies WHERE id = ?`, [req.params.companyId]);
        if (!empresa) return res.status(404).json({ error: 'Empresa não encontrada.' });
        const antes = await pastasLiberadasDaEmpresa(empresa.id);
        for (const pasta of PASTAS_DPO) {
            if (req.body[pasta] === undefined) continue;
            await new Promise((resolve, reject) => db.run(
                `INSERT INTO dpo_company_folders (company_id, folder_key, enabled, updated_at) VALUES (?, ?, ?, CURRENT_TIMESTAMP)
                 ON CONFLICT(company_id, folder_key) DO UPDATE SET enabled = excluded.enabled, updated_at = CURRENT_TIMESTAMP`,
                [empresa.id, pasta, req.body[pasta] ? 1 : 0], (err) => err ? reject(err) : resolve()
            ));
        }
        if (req.body.primeiraAuditoria !== undefined) {
            await new Promise((resolve, reject) => db.run(`UPDATE companies SET dpo_primeira_auditoria = ? WHERE id = ?`, [req.body.primeiraAuditoria ? 1 : 0, empresa.id], (err) => err ? reject(err) : resolve()));
        }
        const depois = await pastasLiberadasDaEmpresa(empresa.id);
        const novas = PASTAS_DPO.filter(p => depois[p] && !antes[p]).map(p => ROTULOS_PASTAS_DPO[p]);
        if (novas.length) notificarGestoresDaEmpresa(empresa.id, 'DPO Ambev — nova pasta liberada', `O Master liberou para sua empresa: ${novas.join(', ')}.`);
        res.json({ message: 'Pastas atualizadas!', pastas: depois });
    } catch (e) { res.status(400).json({ error: 'Erro ao atualizar as pastas da empresa.' }); }
});

async function listarBatePapoDpo(companyId, pilarKey) {
    const linhas = await dbAll(`
        SELECT cq.*, u.name as autorNome
        FROM dpo_chat_questions cq
        LEFT JOIN users u ON u.id = cq.created_by
        WHERE cq.company_id = ? AND cq.pillar_key = ?
    `, [companyId, pilarKey]);
    const ordem = ordemDasPerguntasDoPilarDpo(pilarKey);
    linhas.sort((a, b) => {
        const oa = ordem[a.question_numero] !== undefined ? ordem[a.question_numero] : 99999;
        const ob = ordem[b.question_numero] !== undefined ? ordem[b.question_numero] : 99999;
        if (oa !== ob) return oa - ob;
        return String(a.created_at).localeCompare(String(b.created_at)) || a.id - b.id;
    });
    return linhas.map(l => ({ ...l, perguntaPilarTexto: textoDaPerguntaDpo(pilarKey, l.question_numero) }));
}

app.get('/api/dpo/bate-papo/:pillarKey', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const companyId = await resolverEmpresaBatePapoDpo(req, res, req.params.pillarKey, req.query.company_id);
        if (!companyId) return;
        res.json(await listarBatePapoDpo(companyId, req.params.pillarKey));
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar o conhecimento do time.' }); }
});

// Aceita uma OU várias perguntas de uma vez para a mesma pergunta do pilar:
// { pillarKey, questionNumero, itens: [{ pergunta, resposta }, ...] }
// (também aceita { pergunta, resposta } soltos, para uma só).
app.post('/api/dpo/bate-papo', requireRole('admin', 'client_admin'), async (req, res) => {
    const { pillarKey, questionNumero, company_id } = req.body;
    const itens = Array.isArray(req.body.itens) ? req.body.itens : [{ pergunta: req.body.pergunta, resposta: req.body.resposta }];
    const validos = itens
        .map(i => ({ pergunta: String((i && i.pergunta) || '').trim(), resposta: String((i && i.resposta) || '').trim() }))
        .filter(i => i.pergunta);
    if (!questionNumero) return res.status(400).json({ error: 'Escolha a pergunta do pilar.' });
    if (!validos.length) return res.status(400).json({ error: 'Descreva pelo menos uma pergunta.' });
    try {
        const companyId = await resolverEmpresaBatePapoDpo(req, res, pillarKey, company_id);
        if (!companyId) return;
        if (!textoDaPerguntaDpo(pillarKey, questionNumero)) return res.status(400).json({ error: 'Pergunta do pilar inválida.' });
        for (const item of validos) {
            await new Promise((resolve, reject) => db.run(
                `INSERT INTO dpo_chat_questions (company_id, pillar_key, question_numero, pergunta, resposta, created_by) VALUES (?, ?, ?, ?, ?, ?)`,
                [companyId, pillarKey, questionNumero, item.pergunta, item.resposta || null, req.user.userId],
                (err) => err ? reject(err) : resolve()
            ));
        }
        res.json({ message: validos.length === 1 ? 'Pergunta adicionada!' : `${validos.length} perguntas adicionadas!` });
    } catch (e) { res.status(400).json({ error: 'Erro ao salvar o registro de conhecimento do time.' }); }
});

async function obterBatePapoComAcesso(req, res, id) {
    const item = await dbGet(`SELECT * FROM dpo_chat_questions WHERE id = ?`, [id]);
    if (!item) { res.status(404).json({ error: 'Pergunta não encontrada.' }); return null; }
    if (req.user.role === 'client_admin' && String(item.company_id) !== String(req.user.companyId)) {
        res.status(403).json({ error: 'Esta pergunta não pertence à sua empresa.' });
        return null;
    }
    if (!(await empresaTemPastaDpo(req, res, item.company_id, 'batepapo'))) return null;
    return item;
}

app.put('/api/dpo/bate-papo/:id', requireRole('admin', 'client_admin'), async (req, res) => {
    const { questionNumero, pergunta, resposta } = req.body;
    try {
        const item = await obterBatePapoComAcesso(req, res, req.params.id);
        if (!item) return;
        const novoNumero = questionNumero || item.question_numero;
        if (!textoDaPerguntaDpo(item.pillar_key, novoNumero)) return res.status(400).json({ error: 'Pergunta do pilar inválida.' });
        const novaPergunta = pergunta !== undefined ? String(pergunta).trim() : item.pergunta;
        if (!novaPergunta) return res.status(400).json({ error: 'Descreva a pergunta.' });
        await new Promise((resolve, reject) => db.run(
            `UPDATE dpo_chat_questions SET question_numero = ?, pergunta = ?, resposta = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?`,
            [novoNumero, novaPergunta, resposta !== undefined ? (String(resposta).trim() || null) : item.resposta, item.id],
            (err) => err ? reject(err) : resolve()
        ));
        res.json({ message: 'Pergunta atualizada!' });
    } catch (e) { res.status(400).json({ error: 'Erro ao atualizar a pergunta.' }); }
});

app.delete('/api/dpo/bate-papo/:id', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const item = await obterBatePapoComAcesso(req, res, req.params.id);
        if (!item) return;
        await new Promise((resolve, reject) => db.run(`DELETE FROM dpo_chat_questions WHERE id = ?`, [item.id], (err) => err ? reject(err) : resolve()));
        res.json({ message: 'Pergunta excluída!' });
    } catch (e) { res.status(400).json({ error: 'Erro ao excluir a pergunta.' }); }
});

// ----- Master sobe planilha (anexo) que vira perguntas do "Conhecimento do Time" -----
const uploadPlanilhaMemDpo = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } });
function normCab(t) { return String(t || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' ').trim(); }
function textoCelula(v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'object') {
        if (v.richText) return v.richText.map(r => r.text).join('');
        if (v.text !== undefined) return String(v.text);
        if (v.result !== undefined) return String(v.result);
        if (v instanceof Date) return v.toISOString().slice(0, 10);
    }
    return String(v);
}
function pilarPorRotuloDpo(t) {
    const n = normCab(t); if (!n) return null;
    for (const k of DPO_PILARES_ORDEM) {
        const lb = normCab((DPO_AMBEV_DATA[k] || {}).label);
        if (n === k || n === lb || lb.startsWith(n) || n.includes(lb) || lb.split(' ')[0] === n.split(' ')[0]) return k;
    }
    return null;
}
async function lerPlanilhaConhecimentoDpo(buffer, nomeArquivo, pilarPadrao, numeroPadrao) {
    const wb = new ExcelJS.Workbook();
    if (/\.csv$/i.test(nomeArquivo || '')) {
        const { Readable } = require('stream');
        await wb.csv.read(Readable.from([buffer.toString('utf8').replace(/^﻿/, '')]), { parserOptions: { delimiter: /;/.test(buffer.toString('utf8').split(/\r?\n/)[0]) ? ';' : ',' } });
    } else await wb.xlsx.load(buffer);
    const linhas = [];
    for (const ws of wb.worksheets) {
        // Acha a linha de cabeçalho (nas 10 primeiras) com "pergunta".
        let cab = null, iCab = 0;
        for (let r = 1; r <= Math.min(10, ws.rowCount); r++) {
            const vals = (ws.getRow(r).values || []).map(textoCelula).map(normCab);
            if (vals.some(v => /pergunta/.test(v))) { cab = vals; iCab = r; break; }
        }
        if (!cab) continue;
        const col = (re, excl) => cab.findIndex(v => v && re.test(v) && !(excl && excl.test(v)));
        const cNum = col(/^(n[ºo°.]|no |num\b|numero|item do pilar)/);
        const cPilar = col(/^pilar$/);
        const cPerg = (() => { let c = col(/pergunta ao time|pergunta do time|pergunta feita|^pergunta$|^perguntas$/); if (c < 0) c = col(/pergunta/, /pilar|^n/); return c; })();
        const cResp = col(/resposta/);
        const cPergPilar = col(/pergunta do pilar/, /^n/);
        if (cPerg < 0) continue;
        for (let r = iCab + 1; r <= ws.rowCount; r++) {
            const row = ws.getRow(r);
            const v = c => c > 0 ? textoCelula(row.getCell(c).value).trim() : '';
            const pergunta = v(cPerg), resposta = v(cResp);
            if (!pergunta) continue;
            let pk = (cPilar > 0 && pilarPorRotuloDpo(v(cPilar))) || pilarPadrao;
            let num = v(cNum).replace(',', '.').replace(/[^\d.]/g, ' ').trim().split(' ')[0] || '';
            if (!num && cPergPilar > 0) { const m = v(cPergPilar).match(/^\s*(\d+(?:\.\d+)+)/); if (m) num = m[1]; }
            if (!num) num = numeroPadrao || '';
            const ok = pk && num && textoDaPerguntaDpo(pk, num);
            linhas.push({ linha: r, aba: ws.name, pillar_key: pk || null, question_numero: num || null, perguntaPilarTexto: ok ? textoDaPerguntaDpo(pk, num) : '', pergunta: pergunta.slice(0, 2000), resposta: resposta.slice(0, 4000),
                erro: !pk ? 'Pilar não identificado' : !num ? 'Sem nº da pergunta do pilar' : !ok ? `Pergunta ${num} não existe em ${(DPO_AMBEV_DATA[pk] || {}).label || pk}` : '' });
            if (linhas.length >= 1000) break;
        }
    }
    return linhas;
}

app.get('/api/admin/dpo/bate-papo/modelo/:pillarKey', requireRole('admin'), async (req, res) => {
    try {
        const pk = req.params.pillarKey;
        if (!DPO_AMBEV_DATA[pk]) return res.status(400).json({ error: 'Pilar inválido.' });
        const wb = new ExcelJS.Workbook(); wb.creator = 'Impulsionar V4';
        const ws = wb.addWorksheet('Conhecimento do Time', { views: [{ state: 'frozen', ySplit: 1 }] });
        ws.columns = [
            { header: 'Pilar', key: 'pilar', width: 22 },
            { header: 'Nº Pergunta do Pilar', key: 'numero', width: 12 },
            { header: 'Pergunta do Pilar (referência)', key: 'pp', width: 50 },
            { header: 'Pergunta ao time', key: 'pergunta', width: 55 },
            { header: 'Resposta', key: 'resposta', width: 65 }
        ];
        const cab = ws.getRow(1);
        cab.font = { bold: true, color: { argb: 'FFFFFFFF' } }; cab.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDC4C4C' } };
        cab.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }; cab.height = 30;
        (DPO_AMBEV_DATA[pk].grupos || []).forEach(g => g.perguntas.forEach(q => {
            const row = ws.addRow({ pilar: DPO_AMBEV_DATA[pk].label, numero: q.numero, pp: q.questao, pergunta: '', resposta: '' });
            row.alignment = { vertical: 'top', wrapText: true };
            row.getCell('pp').font = { color: { argb: 'FF64748B' } };
        }));
        const buf = await wb.xlsx.writeBuffer();
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="modelo-conhecimento-do-time-${pk}.xlsx"`);
        res.send(Buffer.from(buf));
    } catch (e) { res.status(500).json({ error: 'Erro ao gerar o modelo.' }); }
});

// multipart: file, pillarKey, questionNumero (opcional), companyIds (JSON) , previa=1 só lê.
app.post('/api/admin/dpo/bate-papo/importar', requireRole('admin'), (req, res) => {
    uploadPlanilhaMemDpo.single('file')(req, res, async (err) => {
        if (err) return res.status(400).json({ error: err.code === 'LIMIT_FILE_SIZE' ? 'Arquivo muito grande (máximo 10MB).' : err.message });
        if (!req.file) return res.status(400).json({ error: 'Envie a planilha (Excel .xlsx ou .csv).' });
        if (!/\.(xlsx|xlsm|csv)$/i.test(req.file.originalname || '')) return res.status(400).json({ error: 'Envie a planilha em Excel (.xlsx) ou .csv — use o modelo.' });
        const pilarPadrao = DPO_PILARES_ORDEM.includes(req.body.pillarKey) ? req.body.pillarKey : null;
        try {
            const linhas = await lerPlanilhaConhecimentoDpo(req.file.buffer, req.file.originalname, pilarPadrao, String(req.body.questionNumero || '').trim());
            const validas = linhas.filter(l => !l.erro);
            if (req.body.previa === '1' || req.body.previa === 'true') return res.json({ linhas, validas: validas.length });
            let ids = []; try { ids = JSON.parse(req.body.companyIds || '[]').map(Number).filter(n => n > 0); } catch (e) { ids = []; }
            if (!ids.length) return res.status(400).json({ error: 'Escolha ao menos uma empresa.' });
            if (!validas.length) return res.status(400).json({ error: 'Nenhuma linha válida na planilha.' });
            let inseridas = 0, puladas = 0;
            for (const cid of ids) {
                for (const l of validas) {
                    const ja = await dbGet(`SELECT id FROM dpo_chat_questions WHERE company_id = ? AND pillar_key = ? AND question_numero = ? AND pergunta = ?`, [cid, l.pillar_key, l.question_numero, l.pergunta]);
                    if (ja) {
                        if (l.resposta) await new Promise(r => db.run(`UPDATE dpo_chat_questions SET resposta = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?`, [l.resposta, ja.id], () => r()));
                        puladas++; continue;
                    }
                    await new Promise((ok, ko) => db.run(`INSERT INTO dpo_chat_questions (company_id, pillar_key, question_numero, pergunta, resposta, created_by, origem) VALUES (?, ?, ?, ?, ?, ?, 'impulsionar')`,
                        [cid, l.pillar_key, l.question_numero, l.pergunta, l.resposta || null, req.user.userId], (e) => e ? ko(e) : ok()));
                    inseridas++;
                }
                const pilares = [...new Set(validas.map(l => (DPO_AMBEV_DATA[l.pillar_key] || {}).label))].join(', ');
                notificarPorCompanyAdmins(cid, '📁 Conhecimento do Time atualizado', `A Impulsionar adicionou ${validas.length} pergunta(s) em ${pilares}.`, 'dpoHome');
            }
            res.json({ message: `${inseridas} pergunta(s) adicionada(s) em ${ids.length} empresa(s)${puladas ? ` · ${puladas} já existia(m) (resposta atualizada)` : ''}.`, inseridas, puladas, ignoradas: linhas.length - validas.length });
        } catch (e) {
            console.error('Importar conhecimento do time:', e.message);
            res.status(400).json({ error: 'Não consegui ler a planilha. Use o modelo (Baixar modelo) e salve em .xlsx.' });
        }
    });
});

// Excel padrão da pasta "Perguntas Bate-Papo" de um pilar, já ordenado pela
// pergunta do pilar.
app.get('/api/dpo/bate-papo/:pillarKey/export', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const pilarKey = req.params.pillarKey;
        const companyId = await resolverEmpresaBatePapoDpo(req, res, pilarKey, req.query.company_id);
        if (!companyId) return;
        const empresa = await dbGet(`SELECT name FROM companies WHERE id = ?`, [companyId]);
        const linhas = await listarBatePapoDpo(companyId, pilarKey);
        const pilarLabel = DPO_AMBEV_DATA[pilarKey].label;

        const workbook = new ExcelJS.Workbook();
        workbook.creator = 'Impulsionar V4';
        workbook.created = new Date();
        const sheet = workbook.addWorksheet('Conhecimento do Time', { views: [{ state: 'frozen', ySplit: 1 }] });
        sheet.columns = [
            { header: 'Empresa', key: 'empresa', width: 24 },
            { header: 'Pilar', key: 'pilar', width: 24 },
            { header: 'Nº Pergunta do Pilar', key: 'numero', width: 12 },
            { header: 'Pergunta do Pilar', key: 'perguntaPilar', width: 40 },
            { header: 'Item', key: 'item', width: 7 },
            { header: 'Pergunta ao time', key: 'pergunta', width: 55 },
            { header: 'Resposta', key: 'resposta', width: 65 },
            { header: 'Registrado por', key: 'autor', width: 22 },
            { header: 'Data', key: 'data', width: 13 }
        ];
        const cabecalho = sheet.getRow(1);
        cabecalho.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cabecalho.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDC4C4C' } };
        cabecalho.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        cabecalho.height = 30;

        const contadorPorPergunta = {};
        linhas.forEach(l => {
            contadorPorPergunta[l.question_numero] = (contadorPorPergunta[l.question_numero] || 0) + 1;
            const row = sheet.addRow({
                empresa: empresa ? empresa.name : '',
                pilar: pilarLabel,
                numero: l.question_numero,
                perguntaPilar: l.perguntaPilarTexto || '',
                item: contadorPorPergunta[l.question_numero],
                pergunta: l.pergunta || '',
                resposta: l.resposta || '',
                autor: l.autorNome || '',
                data: l.created_at ? new Date(String(l.created_at).replace(' ', 'T') + 'Z').toLocaleDateString('pt-BR') : ''
            });
            row.alignment = { vertical: 'top', wrapText: true };
            ['numero', 'item', 'data'].forEach(k => { row.getCell(k).alignment = { vertical: 'top', horizontal: 'center' }; });
        });
        sheet.eachRow(row => row.eachCell(cell => {
            cell.border = { top: { style: 'thin', color: { argb: 'FFE2E8F0' } }, left: { style: 'thin', color: { argb: 'FFE2E8F0' } }, bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } }, right: { style: 'thin', color: { argb: 'FFE2E8F0' } } };
        }));
        sheet.autoFilter = { from: 'A1', to: 'I1' };

        const buffer = await workbook.xlsx.writeBuffer();
        const nomeArquivo = `conhecimento-do-time-${pilarKey}-${(empresa ? empresa.name : 'empresa').replace(/[^a-z0-9]+/gi, '-')}.xlsx`;
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${nomeArquivo}"`);
        res.send(Buffer.from(buffer));
    } catch (e) {
        console.error('Erro ao exportar bate-papo DPO em Excel:', e.message);
        res.status(500).json({ error: 'Erro ao exportar as perguntas em Excel.' });
    }
});

// ---------- DPO Ambev — pasta "Material do Pilar" (evidências p/ auditoria) ----------
// Aberta por PERGUNTA do pilar e, dentro dela, por ITEM da verificação (os
// "1.", "2."... do texto de verificação). Em cada item a empresa sobe o que o
// auditor vai pedir: padrão, ata de treinamento + check de retenção, outras
// evidências e links para sistemas externos.
const STATUS_ITEM_MATERIAL_DPO = ['pendente', 'em_andamento', 'pronto'];
const ROTULOS_STATUS_ITEM_MATERIAL_DPO = { pendente: 'Pendente', em_andamento: 'Em andamento', pronto: 'Pronto p/ auditoria' };
const TIPOS_EVIDENCIA_DPO = ['padrao', 'ata', 'evidencia', 'link'];
const ROTULOS_TIPO_EVIDENCIA_DPO = { padrao: 'Padrão', ata: 'Ata de treinamento', evidencia: 'Outra evidência', link: 'Link sistema externo' };

// Quebra o texto de verificação nos itens numerados ("1. ...", "2. ...").
function itensDaVerificacaoDpo(texto) {
    const t = String(texto || '').replace(/\r/g, '');
    const marcas = [];
    const re = /(^|\n)\s*(\d{1,2})\s*[.)\-]\s+/g;
    let m;
    while ((m = re.exec(t))) marcas.push({ numero: m[2], inicio: m.index + m[1].length, corpo: m.index + m[0].length });
    if (!marcas.length) return t.trim() ? [{ numero: '1', texto: t.trim() }] : [];
    return marcas.map((mk, i) => ({ numero: mk.numero, texto: t.slice(mk.corpo, i + 1 < marcas.length ? marcas[i + 1].inicio : t.length).trim() }));
}

// Sugere o que o item pede, pelo texto: padrão e/ou treinamento.
function sugestoesDoItemDpo(texto) {
    return {
        pedePadrao: /padr[ãa]o|padr[õo]es|padroniz|procedimento|\bPOP\b|\bSOP\b|\bTOR\b/i.test(texto),
        pedeTreinamento: /treinad|treinament|capacita|reciclage|qualifica[çc]/i.test(texto)
    };
}

function perguntaDoPilarDpo(pilarKey, numero) {
    const pilarInfo = DPO_AMBEV_DATA[pilarKey];
    if (!pilarInfo) return null;
    for (const g of pilarInfo.grupos) {
        const q = g.perguntas.find(x => x.numero === numero);
        if (q) return { grupo: g, pergunta: q };
    }
    return null;
}

function itemValidoDpo(pilarKey, questionNumero, itemNumero) {
    const achou = perguntaDoPilarDpo(pilarKey, questionNumero);
    if (!achou) return false;
    return itensDaVerificacaoDpo(achou.pergunta.verificacao).some(i => i.numero === String(itemNumero));
}

// Endereço público do sistema: a URL configurada pelo Master (ou APP_BASE_URL),
// senão o próprio endereço da requisição (respeitando o https do Railway).
function baseUrlPublicaDpo(req) {
    if (urlPublicaValida(appBaseUrlAtiva)) return appBaseUrlAtiva.replace(/\/$/, '');
    const proto = String(req.get('x-forwarded-proto') || req.protocol || 'https').split(',')[0].trim();
    const b = `${proto}://${req.get('host')}`;
    if (typeof ULTIMA_BASE_DPO !== 'undefined' && !/localhost|127\.0\.0\.1/.test(b)) ULTIMA_BASE_DPO = b;
    return b;
}

function urlPublicaDoCheckDpo(req, token) {
    return `${baseUrlPublicaDpo(req)}/retencao.html?t=${token}`;
}

async function montarMaterialDoPilarDpo(req, companyId, pilarKey) {
    const pilarInfo = DPO_AMBEV_DATA[pilarKey];
    const status = await dbAll(`SELECT * FROM dpo_material_status WHERE company_id = ? AND pillar_key = ?`, [companyId, pilarKey]);
    const evidencias = await dbAll(`
        SELECT e.*, u.name as autorNome FROM dpo_material_evidencias e
        LEFT JOIN users u ON u.id = e.created_by
        WHERE e.company_id = ? AND e.pillar_key = ? ORDER BY e.created_at ASC, e.id ASC`, [companyId, pilarKey]);
    const checks = await dbAll(`
        SELECT c.*, (SELECT COUNT(*) FROM dpo_retention_responses r WHERE r.check_id = c.id) as totalRespostas,
               (SELECT SUM(acertos) FROM dpo_retention_responses r WHERE r.check_id = c.id) as somaAcertos,
               (SELECT SUM(total_objetivas) FROM dpo_retention_responses r WHERE r.check_id = c.id) as somaObjetivas
        FROM dpo_retention_checks c WHERE c.company_id = ? AND c.pillar_key = ? ORDER BY c.created_at ASC`, [companyId, pilarKey]);
    const chave = (q, i) => `${q}|${i}`;
    const mapaStatus = {}; status.forEach(s => { mapaStatus[chave(s.question_numero, s.item_numero)] = s; });
    const mapaEvid = {}; evidencias.forEach(e => { (mapaEvid[chave(e.question_numero, e.item_numero)] = mapaEvid[chave(e.question_numero, e.item_numero)] || []).push(e); });
    const mapaChecks = {}; checks.forEach(c => {
        const perguntas = JSON.parse(c.perguntas || '[]');
        (mapaChecks[chave(c.question_numero, c.item_numero)] = mapaChecks[chave(c.question_numero, c.item_numero)] || []).push({
            id: c.id, titulo: c.titulo, ativo: !!c.ativo, created_at: c.created_at, perguntas,
            link: urlPublicaDoCheckDpo(req, c.token),
            totalRespostas: c.totalRespostas || 0,
            mediaAcerto: c.somaObjetivas ? Math.round((c.somaAcertos || 0) * 100 / c.somaObjetivas) : null
        });
    });
    const resumo = { totalItens: 0, pendente: 0, em_andamento: 0, pronto: 0, totalEvidencias: evidencias.length, totalChecks: checks.length };
    const grupos = pilarInfo.grupos.map(g => ({
        numero: g.numero, titulo: g.titulo,
        perguntas: g.perguntas.map(q => {
            const itens = itensDaVerificacaoDpo(q.verificacao).map(it => {
                const st = mapaStatus[chave(q.numero, it.numero)];
                const situacao = st ? st.status : 'pendente';
                resumo.totalItens++; resumo[situacao] = (resumo[situacao] || 0) + 1;
                return {
                    numero: it.numero, texto: it.texto, ...sugestoesDoItemDpo(it.texto),
                    status: situacao, observacao: st ? st.observacao : null,
                    evidencias: mapaEvid[chave(q.numero, it.numero)] || [],
                    checks: mapaChecks[chave(q.numero, it.numero)] || []
                };
            });
            return {
                numero: q.numero, questao: q.questao, mandatoria: !!q.mandatoria, how_to_check: q.how_to_check || '',
                itens, prontos: itens.filter(i => i.status === 'pronto').length
            };
        })
    }));
    return { key: pilarKey, label: pilarInfo.label, resumo, grupos };
}

app.get('/api/dpo/material/:pillarKey', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const companyId = await resolverEmpresaPastaDpo(req, res, req.params.pillarKey, req.query.company_id, 'material');
        if (!companyId) return;
        res.json(await montarMaterialDoPilarDpo(req, companyId, req.params.pillarKey));
    } catch (e) {
        console.error('Erro ao carregar Material do Pilar:', e.message);
        res.status(500).json({ error: 'Erro ao carregar o material do pilar.' });
    }
});

app.put('/api/dpo/material/status', requireRole('admin', 'client_admin'), async (req, res) => {
    const { pillarKey, questionNumero, itemNumero, status, observacao, company_id } = req.body;
    try {
        const companyId = await resolverEmpresaPastaDpo(req, res, pillarKey, company_id, 'material');
        if (!companyId) return;
        if (!itemValidoDpo(pillarKey, questionNumero, itemNumero)) return res.status(400).json({ error: 'Item inválido.' });
        const atual = await dbGet(`SELECT * FROM dpo_material_status WHERE company_id = ? AND pillar_key = ? AND question_numero = ? AND item_numero = ?`, [companyId, pillarKey, questionNumero, String(itemNumero)]);
        const novoStatus = STATUS_ITEM_MATERIAL_DPO.includes(status) ? status : (atual ? atual.status : 'pendente');
        const novaObs = observacao !== undefined ? (String(observacao).trim() || null) : (atual ? atual.observacao : null);
        await new Promise((resolve, reject) => db.run(
            `INSERT INTO dpo_material_status (company_id, pillar_key, question_numero, item_numero, status, observacao, updated_by, updated_at)
             VALUES (?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
             ON CONFLICT(company_id, pillar_key, question_numero, item_numero) DO UPDATE SET status = excluded.status, observacao = excluded.observacao, updated_by = excluded.updated_by, updated_at = CURRENT_TIMESTAMP`,
            [companyId, pillarKey, questionNumero, String(itemNumero), novoStatus, novaObs, req.user.userId], (err) => err ? reject(err) : resolve()
        ));
        res.json({ message: 'Item atualizado!' });
    } catch (e) { res.status(400).json({ error: 'Erro ao atualizar o item.' }); }
});

// Upload próprio do Material do Pilar — aceita também Excel e PowerPoint
// (padrões e atas costumam vir nesses formatos), além de PDF, Word, imagem e vídeo.
const uploadMaterialDpo = multer({
    storage: armazenamentoUpload,
    limits: { fileSize: 100 * 1024 * 1024 },
    fileFilter: (req, file, cb) => {
        const permitidos = /video\/|image\/|application\/pdf|application\/msword|application\/vnd\.openxmlformats-officedocument\.(wordprocessingml|spreadsheetml|presentationml)|application\/vnd\.ms-excel|application\/vnd\.ms-powerpoint|text\/csv|text\/plain/;
        if (permitidos.test(file.mimetype)) return cb(null, true);
        // Planilhas com macro (.xlsm) e afins às vezes chegam como octet-stream.
        if (/\.(xlsm|xlsx|xls|xlsb|csv|pdf|docx?|pptx?)$/i.test(file.originalname || '')) return cb(null, true);
        cb(new Error('Tipo de arquivo não permitido. Envie PDF, Word, Excel, PowerPoint, imagem ou vídeo.'));
    }
});

app.post('/api/dpo/material/upload', requireRole('admin', 'client_admin'), (req, res) => {
    uploadMaterialDpo.single('file')(req, res, (err) => {
        if (err) return res.status(400).json({ error: err.code === 'LIMIT_FILE_SIZE' ? 'Arquivo muito grande (máximo 100MB).' : err.message });
        if (!req.file) return res.status(400).json({ error: 'Nenhum arquivo recebido.' });
        res.json({ url: '/uploads/' + req.file.filename, originalName: req.file.originalname });
    });
});

app.post('/api/dpo/material/evidencias', requireRole('admin', 'client_admin'), async (req, res) => {
    const { pillarKey, questionNumero, itemNumero, tipo, titulo, url, originalName, company_id } = req.body;
    if (!TIPOS_EVIDENCIA_DPO.includes(tipo)) return res.status(400).json({ error: 'Tipo de evidência inválido.' });
    const urlLimpa = String(url || '').trim();
    if (tipo === 'link') {
        if (!/^https?:\/\/\S+$/i.test(urlLimpa)) return res.status(400).json({ error: 'Informe um link válido (começando com http:// ou https://).' });
    } else if (!/^\/uploads\/[\w.\-]+$/.test(urlLimpa)) {
        return res.status(400).json({ error: 'Envie o arquivo antes de salvar.' });
    }
    try {
        const companyId = await resolverEmpresaPastaDpo(req, res, pillarKey, company_id, 'material');
        if (!companyId) return;
        if (!itemValidoDpo(pillarKey, questionNumero, itemNumero)) return res.status(400).json({ error: 'Item inválido.' });
        await new Promise((resolve, reject) => db.run(
            `INSERT INTO dpo_material_evidencias (company_id, pillar_key, question_numero, item_numero, tipo, titulo, url, original_name, created_by) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`,
            [companyId, pillarKey, questionNumero, String(itemNumero), tipo, String(titulo || '').trim().slice(0, 200) || null, urlLimpa, originalName ? String(originalName).slice(0, 200) : null, req.user.userId],
            (err) => err ? reject(err) : resolve()
        ));
        // Primeira evidência num item pendente já o coloca "em andamento".
        db.run(`INSERT INTO dpo_material_status (company_id, pillar_key, question_numero, item_numero, status, updated_by) VALUES (?, ?, ?, ?, 'em_andamento', ?)
                ON CONFLICT(company_id, pillar_key, question_numero, item_numero) DO UPDATE SET status = CASE WHEN status = 'pendente' THEN 'em_andamento' ELSE status END`,
            [companyId, pillarKey, questionNumero, String(itemNumero), req.user.userId], () => {});
        res.json({ message: `${ROTULOS_TIPO_EVIDENCIA_DPO[tipo]} adicionado(a)!` });
    } catch (e) { res.status(400).json({ error: 'Erro ao salvar a evidência.' }); }
});

app.delete('/api/dpo/material/evidencias/:id', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const ev = await dbGet(`SELECT * FROM dpo_material_evidencias WHERE id = ?`, [req.params.id]);
        if (!ev) return res.status(404).json({ error: 'Evidência não encontrada.' });
        if (req.user.role === 'client_admin' && String(ev.company_id) !== String(req.user.companyId)) return res.status(403).json({ error: 'Esta evidência não pertence à sua empresa.' });
        if (!(await empresaTemPastaDpo(req, res, ev.company_id, 'material'))) return;
        await new Promise((resolve, reject) => db.run(`DELETE FROM dpo_material_evidencias WHERE id = ?`, [ev.id], (err) => err ? reject(err) : resolve()));
        res.json({ message: 'Removido!' });
    } catch (e) { res.status(400).json({ error: 'Erro ao remover a evidência.' }); }
});

// ----- Check de retenção (perguntas + link público) -----
// Cada pergunta: { texto, tipo: 'aberta' | 'multipla', opcoes: [..], correta: índice | null }
function normalizarPerguntasCheckDpo(lista) {
    if (!Array.isArray(lista)) return [];
    return lista.slice(0, 50).map(p => {
        const texto = String((p && p.texto) || '').trim().slice(0, 500);
        const tipo = p && p.tipo === 'multipla' ? 'multipla' : 'aberta';
        let opcoes = tipo === 'multipla' && Array.isArray(p.opcoes) ? p.opcoes.map(o => String(o || '').trim().slice(0, 300)).filter(Boolean).slice(0, 8) : [];
        let correta = tipo === 'multipla' && p.correta !== null && p.correta !== undefined && p.correta !== '' ? Number(p.correta) : null;
        if (correta !== null && !(correta >= 0 && correta < opcoes.length)) correta = null;
        return { texto, tipo: opcoes.length >= 2 ? tipo : 'aberta', opcoes: opcoes.length >= 2 ? opcoes : [], correta: opcoes.length >= 2 ? correta : null };
    }).filter(p => p.texto);
}

async function obterCheckDpoComAcesso(req, res, id) {
    const check = await dbGet(`SELECT * FROM dpo_retention_checks WHERE id = ?`, [id]);
    if (!check) { res.status(404).json({ error: 'Check de retenção não encontrado.' }); return null; }
    if (req.user.role === 'client_admin' && String(check.company_id) !== String(req.user.companyId)) { res.status(403).json({ error: 'Este check não pertence à sua empresa.' }); return null; }
    if (!(await empresaTemPastaDpo(req, res, check.company_id, 'material'))) return null;
    return check;
}

app.post('/api/dpo/material/checks', requireRole('admin', 'client_admin'), async (req, res) => {
    const { pillarKey, questionNumero, itemNumero, titulo, company_id } = req.body;
    const perguntas = normalizarPerguntasCheckDpo(req.body.perguntas);
    if (!String(titulo || '').trim()) return res.status(400).json({ error: 'Dê um título ao check (ex.: nome do treinamento).' });
    if (!perguntas.length) return res.status(400).json({ error: 'Cadastre pelo menos uma pergunta.' });
    try {
        const companyId = await resolverEmpresaPastaDpo(req, res, pillarKey, company_id, 'material');
        if (!companyId) return;
        if (!itemValidoDpo(pillarKey, questionNumero, itemNumero)) return res.status(400).json({ error: 'Item inválido.' });
        const token = crypto.randomBytes(16).toString('hex');
        await new Promise((resolve, reject) => db.run(
            `INSERT INTO dpo_retention_checks (company_id, pillar_key, question_numero, item_numero, titulo, token, perguntas, created_by) VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
            [companyId, pillarKey, questionNumero, String(itemNumero), String(titulo).trim().slice(0, 200), token, JSON.stringify(perguntas), req.user.userId],
            (err) => err ? reject(err) : resolve()
        ));
        res.json({ message: 'Check de retenção criado — link gerado!', link: urlPublicaDoCheckDpo(req, token) });
    } catch (e) { res.status(400).json({ error: 'Erro ao criar o check de retenção.' }); }
});

app.put('/api/dpo/material/checks/:id', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const check = await obterCheckDpoComAcesso(req, res, req.params.id);
        if (!check) return;
        const titulo = req.body.titulo !== undefined ? String(req.body.titulo).trim().slice(0, 200) : check.titulo;
        if (!titulo) return res.status(400).json({ error: 'Dê um título ao check.' });
        let perguntas = check.perguntas;
        if (req.body.perguntas !== undefined) {
            const novas = normalizarPerguntasCheckDpo(req.body.perguntas);
            if (!novas.length) return res.status(400).json({ error: 'Cadastre pelo menos uma pergunta.' });
            perguntas = JSON.stringify(novas);
        }
        const ativo = req.body.ativo !== undefined ? (req.body.ativo ? 1 : 0) : check.ativo;
        await new Promise((resolve, reject) => db.run(`UPDATE dpo_retention_checks SET titulo = ?, perguntas = ?, ativo = ? WHERE id = ?`, [titulo, perguntas, ativo, check.id], (err) => err ? reject(err) : resolve()));
        res.json({ message: 'Check atualizado!' });
    } catch (e) { res.status(400).json({ error: 'Erro ao atualizar o check.' }); }
});

app.delete('/api/dpo/material/checks/:id', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const check = await obterCheckDpoComAcesso(req, res, req.params.id);
        if (!check) return;
        await new Promise((resolve, reject) => db.run(`DELETE FROM dpo_retention_responses WHERE check_id = ?`, [check.id], (err) => err ? reject(err) : resolve()));
        await new Promise((resolve, reject) => db.run(`DELETE FROM dpo_retention_checks WHERE id = ?`, [check.id], (err) => err ? reject(err) : resolve()));
        res.json({ message: 'Check excluído!' });
    } catch (e) { res.status(400).json({ error: 'Erro ao excluir o check.' }); }
});

app.get('/api/dpo/material/checks/:id/respostas', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const check = await obterCheckDpoComAcesso(req, res, req.params.id);
        if (!check) return;
        const respostas = await dbAll(`SELECT * FROM dpo_retention_responses WHERE check_id = ? ORDER BY created_at DESC`, [check.id]);
        res.json({
            id: check.id, titulo: check.titulo, perguntas: JSON.parse(check.perguntas || '[]'), link: urlPublicaDoCheckDpo(req, check.token),
            respostas: respostas.map(r => ({ ...r, respostas: JSON.parse(r.respostas || '[]') }))
        });
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar as respostas.' }); }
});

function estilizarCabecalhoExcelDpo(sheet, ultimaColuna) {
    const cabecalho = sheet.getRow(1);
    cabecalho.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cabecalho.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDC4C4C' } };
    cabecalho.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cabecalho.height = 30;
    sheet.autoFilter = { from: 'A1', to: `${ultimaColuna}1` };
}

function dataBrDpo(valor) {
    if (!valor) return '';
    const d = new Date(String(valor).replace(' ', 'T') + (String(valor).includes('Z') ? '' : 'Z'));
    return isNaN(d) ? '' : d.toLocaleDateString('pt-BR');
}

app.get('/api/dpo/material/checks/:id/export', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const check = await obterCheckDpoComAcesso(req, res, req.params.id);
        if (!check) return;
        const perguntas = JSON.parse(check.perguntas || '[]');
        const respostas = await dbAll(`SELECT * FROM dpo_retention_responses WHERE check_id = ? ORDER BY created_at ASC`, [check.id]);
        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet('Check de Retenção', { views: [{ state: 'frozen', ySplit: 1 }] });
        sheet.columns = [
            { header: 'Nome', key: 'nome', width: 28 },
            { header: 'Matrícula', key: 'matricula', width: 14 },
            { header: 'Data', key: 'data', width: 13 },
            { header: 'Acertos', key: 'acertos', width: 12 },
            ...perguntas.map((p, i) => ({ header: `${i + 1}. ${p.texto}`, key: 'p' + i, width: 40 }))
        ];
        respostas.forEach(r => {
            const lista = JSON.parse(r.respostas || '[]');
            const linha = { nome: r.nome, matricula: r.matricula || '', data: dataBrDpo(r.created_at), acertos: r.total_objetivas ? `${r.acertos}/${r.total_objetivas}` : '-' };
            perguntas.forEach((p, i) => {
                const v = lista[i];
                linha['p' + i] = p.tipo === 'multipla' ? (v !== null && v !== undefined && p.opcoes[v] !== undefined ? p.opcoes[v] : '') : (v || '');
            });
            sheet.addRow(linha).alignment = { vertical: 'top', wrapText: true };
        });
        estilizarCabecalhoExcelDpo(sheet, sheet.getColumn(sheet.columns.length).letter);
        const buffer = await workbook.xlsx.writeBuffer();
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="check-retencao-${check.id}.xlsx"`);
        res.send(Buffer.from(buffer));
    } catch (e) {
        console.error('Erro ao exportar check de retenção:', e.message);
        res.status(500).json({ error: 'Erro ao exportar as respostas.' });
    }
});

// Dossiê do pilar para a auditoria: cada pergunta/item com situação e todas as
// evidências (arquivos com link, links externos, checks de retenção).
app.get('/api/dpo/material/:pillarKey/export', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const pilarKey = req.params.pillarKey;
        const companyId = await resolverEmpresaPastaDpo(req, res, pilarKey, req.query.company_id, 'material');
        if (!companyId) return;
        const empresa = await dbGet(`SELECT name FROM companies WHERE id = ?`, [companyId]);
        const material = await montarMaterialDoPilarDpo(req, companyId, pilarKey);
        const base = baseUrlPublicaDpo(req);
        const workbook = new ExcelJS.Workbook();
        workbook.creator = 'Impulsionar V4';
        const sheet = workbook.addWorksheet('Material do Pilar', { views: [{ state: 'frozen', ySplit: 1 }] });
        sheet.columns = [
            { header: 'Pilar', key: 'pilar', width: 20 },
            { header: 'Grupo', key: 'grupo', width: 28 },
            { header: 'Nº Pergunta', key: 'numero', width: 11 },
            { header: 'Pergunta do Pilar', key: 'pergunta', width: 30 },
            { header: 'Item', key: 'item', width: 7 },
            { header: 'O que o item pede', key: 'texto', width: 60 },
            { header: 'Situação', key: 'status', width: 18 },
            { header: 'Observação', key: 'obs', width: 30 },
            { header: 'Tipo de evidência', key: 'tipo', width: 20 },
            { header: 'Evidência', key: 'evidencia', width: 40 },
            { header: 'Link', key: 'link', width: 50 }
        ];
        material.grupos.forEach(g => g.perguntas.forEach(q => q.itens.forEach(it => {
            const base_ = { pilar: material.label, grupo: `${g.numero} ${g.titulo}`, numero: q.numero, pergunta: q.questao, item: it.numero, texto: it.texto, status: ROTULOS_STATUS_ITEM_MATERIAL_DPO[it.status] || it.status, obs: it.observacao || '' };
            const linhas = [
                ...it.evidencias.map(e => ({ tipo: ROTULOS_TIPO_EVIDENCIA_DPO[e.tipo] || e.tipo, evidencia: e.titulo || e.original_name || '', link: e.tipo === 'link' ? e.url : base + e.url })),
                ...it.checks.map(c => ({ tipo: 'Check de retenção', evidencia: `${c.titulo} — ${c.totalRespostas} resposta(s)${c.mediaAcerto !== null ? ` · ${c.mediaAcerto}% de acerto` : ''}`, link: c.link }))
            ];
            if (!linhas.length) linhas.push({ tipo: '', evidencia: 'Sem evidência cadastrada', link: '' });
            linhas.forEach(l => {
                const row = sheet.addRow({ ...base_, ...l });
                row.alignment = { vertical: 'top', wrapText: true };
                if (l.link) row.getCell('link').value = { text: l.link, hyperlink: l.link };
                const cor = it.status === 'pronto' ? 'FFDCFCE7' : it.status === 'em_andamento' ? 'FFFEF9C3' : 'FFFEE2E2';
                row.getCell('status').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: cor } };
            });
        })));
        estilizarCabecalhoExcelDpo(sheet, 'K');
        const buffer = await workbook.xlsx.writeBuffer();
        const nomeArquivo = `material-pilar-${pilarKey}-${(empresa ? empresa.name : 'empresa').replace(/[^a-z0-9]+/gi, '-')}.xlsx`;
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${nomeArquivo}"`);
        res.send(Buffer.from(buffer));
    } catch (e) {
        console.error('Erro ao exportar Material do Pilar:', e.message);
        res.status(500).json({ error: 'Erro ao exportar o material do pilar.' });
    }
});

// ======================================================================
// MINISTRAR TREINAMENTO (Master) — módulos, apresentação, rascunho e
// check de retenção gerado por IA a partir do conteúdo do treinamento.
// ======================================================================
const PUBLICOS_MT = ['administrativo', 'lideranca', 'operacional', 'todos'];
const ROTULOS_PUBLICO_MT = { administrativo: 'Administrativo', lideranca: 'Liderança', operacional: 'Operacional', todos: 'Todos' };
const STATUS_MT = ['rascunho', 'pronto', 'realizado'];

// Leitor de ZIP mínimo (PPTX/DOCX são ZIP) — sem dependência extra.
function lerZipMt(buf) {
    const zlib = require('zlib');
    const arquivos = {};
    let eocd = -1;
    for (let i = buf.length - 22; i >= Math.max(0, buf.length - 70000); i--) { if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; } }
    if (eocd < 0) return arquivos;
    const total = buf.readUInt16LE(eocd + 10);
    let p = buf.readUInt32LE(eocd + 16);
    for (let n = 0; n < total && p + 46 <= buf.length; n++) {
        if (buf.readUInt32LE(p) !== 0x02014b50) break;
        const metodo = buf.readUInt16LE(p + 10), tamComp = buf.readUInt32LE(p + 20);
        const lenNome = buf.readUInt16LE(p + 28), lenExtra = buf.readUInt16LE(p + 30), lenCom = buf.readUInt16LE(p + 32);
        const offLocal = buf.readUInt32LE(p + 42);
        const nome = buf.slice(p + 46, p + 46 + lenNome).toString('utf8');
        p += 46 + lenNome + lenExtra + lenCom;
        if (!/\.xml$/i.test(nome)) continue;
        try {
            const ini = offLocal + 30 + buf.readUInt16LE(offLocal + 26) + buf.readUInt16LE(offLocal + 28);
            const dado = buf.slice(ini, ini + tamComp);
            arquivos[nome] = metodo === 8 ? zlib.inflateRawSync(dado) : dado;
        } catch (e) { /* ignora entrada corrompida */ }
    }
    return arquivos;
}
function xmlTextoMt(t) { return String(t || '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (m, n) => String.fromCharCode(+n)).replace(/&amp;/g, '&'); }
function paragrafosMt(xml, tagP, tagT) {
    return String(xml).split(new RegExp(`</${tagP}>`)).map(par => {
        const ts = par.match(new RegExp(`<${tagT}(?: [^>]*)?>([\\s\\S]*?)</${tagT}>`, 'g')) || [];
        return xmlTextoMt(ts.map(x => x.replace(/<[^>]+>/g, '')).join('')).trim();
    }).filter(Boolean);
}
// PPTX → [{titulo, texto}] por slide; DOCX → texto.
function extrairOfficeMt(buf, nome) {
    const z = lerZipMt(buf);
    if (/\.pptx$/i.test(nome)) {
        const slides = Object.keys(z).filter(k => /^ppt\/slides\/slide\d+\.xml$/.test(k)).sort((a, b) => +a.match(/(\d+)\.xml/)[1] - +b.match(/(\d+)\.xml/)[1]);
        return { slides: slides.map(k => {
            const xml = z[k].toString('utf8');
            const formas = xml.split('</p:sp>');
            let titulo = '', corpo = [];
            formas.forEach(f => {
                const pars = paragrafosMt(f, 'a:p', 'a:t');
                if (!pars.length) return;
                if (!titulo && /type="(title|ctrTitle)"/.test(f)) titulo = pars.join(' ');
                else corpo.push(...pars);
            });
            if (!titulo && corpo.length) titulo = corpo.shift();
            return { titulo: titulo.slice(0, 200), texto: corpo.map(l => '• ' + l).join('\n').slice(0, 3000) };
        }) };
    }
    if (/\.docx$/i.test(nome) && z['word/document.xml']) return { texto: paragrafosMt(z['word/document.xml'].toString('utf8'), 'w:p', 'w:t').join('\n') };
    return {};
}
// Nome de arquivo com acento chega do navegador em latin1 ("GestÃ£o") — conserta.
function consertarAcentoMt(t) {
    const s0 = String(t || '');
    if (!/[ÃÂ][\u0080-\u00ff]/.test(s0)) return s0;
    try { const d = Buffer.from(s0, 'latin1').toString('utf8'); return d.includes('\uFFFD') ? s0 : d; } catch (e) { return s0; }
}
function caminhoUploadMt(url) { return path.join(PASTA_UPLOADS, path.basename(String(url || ''))); }
async function lerArquivoUploadMt(url) {
    const c = caminhoUploadMt(url);
    if (fs.existsSync(c)) return fs.readFileSync(c);
    const r = await restaurarUploadDoBanco(path.basename(String(url || '')));
    return r ? r.dados : null;
}
// Tenta converter Office → PDF (LibreOffice). Se o servidor não tiver, segue sem.
function converterParaPdfMt(caminho) {
    return new Promise(resolve => {
        try {
            const { execFile } = require('child_process');
            const saida = path.join(require('os').tmpdir(), 'mt-' + Date.now());
            fs.mkdirSync(saida, { recursive: true });
            execFile('soffice', ['--headless', '--convert-to', 'pdf', '--outdir', saida, caminho], { timeout: 90000 }, err => {
                if (err) return resolve(null);
                const pdf = fs.readdirSync(saida).find(n => /\.pdf$/i.test(n));
                if (!pdf) return resolve(null);
                const nome = 'mt-' + Date.now() + '-' + Math.round(Math.random() * 1e6) + '.pdf';
                const destino = path.join(PASTA_UPLOADS, nome);
                fs.copyFileSync(path.join(saida, pdf), destino);
                copiarUploadParaBanco({ filename: nome, path: destino, size: fs.statSync(destino).size, mimetype: 'application/pdf' });
                resolve('/uploads/' + nome);
            });
        } catch (e) { resolve(null); }
    });
}

function normalizarModulosMt(lista) {
    return (Array.isArray(lista) ? lista : []).slice(0, 60).map((m, i) => {
        const tipo = m && m.tipo === 'arquivo' ? 'arquivo' : 'slides';
        const limpo = { id: String((m && m.id) || ('m' + Date.now().toString(36) + i)).slice(0, 40), titulo: consertarAcentoMt(String((m && m.titulo) || `Módulo ${i + 1}`).trim().slice(0, 200)), tipo, oculto: !!(m && m.oculto) };
        if (tipo === 'slides') {
            limpo.slides = (Array.isArray(m.slides) ? m.slides : []).slice(0, 200).map(sl => ({
                titulo: String((sl && sl.titulo) || '').slice(0, 300), texto: String((sl && sl.texto) || '').slice(0, 5000),
                imagem: /^\/uploads\/[\w.\-]+$|^https?:\/\/\S+$/.test(String((sl && sl.imagem) || '')) ? sl.imagem : '',
                notas: String((sl && sl.notas) || '').slice(0, 3000), layout: ['padrao', 'destaque', 'imagem'].includes(sl && sl.layout) ? sl.layout : 'padrao', oculto: !!(sl && sl.oculto)
            }));
        } else {
            const a = m.arquivo || {};
            limpo.arquivo = { url: /^\/uploads\/[\w.\-]+$|^https?:\/\/\S+$/.test(String(a.url || '')) ? a.url : '', nome: consertarAcentoMt(String(a.nome || '').slice(0, 200)), mime: String(a.mime || '').slice(0, 120),
                pdfUrl: /^\/uploads\/[\w.\-]+$/.test(String(a.pdfUrl || '')) ? a.pdfUrl : '', slidesExtraidos: Array.isArray(a.slidesExtraidos) ? a.slidesExtraidos.slice(0, 200).map(x => ({ titulo: String(x.titulo || '').slice(0, 300), texto: String(x.texto || '').slice(0, 3000), oculto: !!x.oculto })) : [] };
            limpo.texto = String(m.texto || '').slice(0, 60000);
        }
        return limpo;
    });
}
function camposTreinamentoMt(b) {
    const d = {};
    if (b.titulo !== undefined) d.titulo = consertarAcentoMt(String(b.titulo || '').trim().slice(0, 200));
    if (b.objetivo !== undefined) d.objetivo = String(b.objetivo || '').trim().slice(0, 3000) || null;
    if (b.publico !== undefined) d.publico = PUBLICOS_MT.includes(b.publico) ? b.publico : 'todos';
    if (b.data_treinamento !== undefined) d.data_treinamento = /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2})?$/.test(String(b.data_treinamento || '')) ? b.data_treinamento : null;
    if (b.duracao_min !== undefined) d.duracao_min = Number(b.duracao_min) > 0 ? Math.min(1440, Math.round(Number(b.duracao_min))) : null;
    if (b.local !== undefined) d.local = String(b.local || '').trim().slice(0, 200) || null;
    if (b.instrutor !== undefined) d.instrutor = String(b.instrutor || '').trim().slice(0, 120) || null;
    if (b.company_id !== undefined) d.company_id = Number(b.company_id) > 0 ? Number(b.company_id) : null;
    if (b.participantes !== undefined) d.participantes = Number(b.participantes) >= 0 && b.participantes !== '' && b.participantes !== null ? Math.round(Number(b.participantes)) : null;
    if (b.status !== undefined) d.status = STATUS_MT.includes(b.status) ? b.status : 'rascunho';
    if (b.modulos !== undefined) d.modulos = JSON.stringify(normalizarModulosMt(b.modulos));
    return d;
}
async function resumoChecksMt(ids) {
    if (!ids.length) return {};
    const rows = await dbAll(`SELECT c.id, c.treinamento_id, c.ativo, c.token, c.titulo, c.created_at,
        (SELECT COUNT(*) FROM treinamentos_mt_respostas r WHERE r.check_id = c.id) as respostas,
        (SELECT SUM(acertos) FROM treinamentos_mt_respostas r WHERE r.check_id = c.id) as acertos,
        (SELECT SUM(total_objetivas) FROM treinamentos_mt_respostas r WHERE r.check_id = c.id) as objetivas
        FROM treinamentos_mt_checks c WHERE c.treinamento_id IN (${ids.map(() => '?').join(',')})`, ids);
    const m = {}; rows.forEach(r => { m[r.treinamento_id] = { id: r.id, ativo: !!r.ativo, token: r.token, titulo: r.titulo, respostas: r.respostas || 0, media: r.objetivas ? Math.round((r.acertos || 0) * 100 / r.objetivas) : null }; });
    return m;
}

app.get('/api/admin/treinamentos-mt', requireRole('admin'), async (req, res) => {
    try {
        const l = await dbAll(`SELECT t.*, c.name as companyName FROM treinamentos_mt t LEFT JOIN companies c ON c.id = t.company_id ORDER BY COALESCE(t.updated_at, t.created_at) DESC`);
        const checks = await resumoChecksMt(l.map(t => t.id));
        const assin = {}; (await dbAll(`SELECT treinamento_id, COUNT(*) n FROM treinamentos_mt_presencas GROUP BY treinamento_id`)).forEach(r => { assin[r.treinamento_id] = r.n; });
        res.json(l.map(t => {
            let mods = []; try { mods = JSON.parse(t.modulos || '[]'); } catch (e) {}
            const nSlides = mods.reduce((s, m) => s + (m.tipo === 'slides' ? (m.slides || []).length : 1), 0);
            const { modulos, ...resto } = t;
            return { ...resto, titulo: consertarAcentoMt(resto.titulo), qtdModulos: mods.length, qtdSlides: nSlides, modulosResumo: mods.map(m => ({ titulo: m.titulo, tipo: m.tipo })), check: checks[t.id] || null, assinaturas: assin[t.id] || 0 };
        }));
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar os treinamentos.' }); }
});

app.get('/api/admin/treinamentos-mt/:id', requireRole('admin'), async (req, res) => {
    try {
        const t = await dbGet(`SELECT * FROM treinamentos_mt WHERE id = ?`, [req.params.id]);
        if (!t) return res.status(404).json({ error: 'Treinamento não encontrado.' });
        let modulos = []; try { modulos = JSON.parse(t.modulos || '[]'); } catch (e) {}
        modulos.forEach(m => { m.titulo = consertarAcentoMt(m.titulo); if (m.arquivo) m.arquivo.nome = consertarAcentoMt(m.arquivo.nome); });
        t.titulo = consertarAcentoMt(t.titulo);
        const ch = (await resumoChecksMt([t.id]))[t.id] || null;
        let check = null;
        if (ch) {
            const c = await dbGet(`SELECT * FROM treinamentos_mt_checks WHERE id = ?`, [ch.id]);
            check = { ...ch, perguntas: JSON.parse(c.perguntas || '[]'), gerado_por_ia: !!c.gerado_por_ia, link: `${baseUrlPublicaDpo(req)}/retencao.html?t=${c.token}` };
        }
        const cfg = await lerConfigMt();
        const tk = await garantirAtaTokenMt(t);
        const nAss = await dbGet(`SELECT COUNT(*) n FROM treinamentos_mt_presencas WHERE treinamento_id = ?`, [t.id]);
        res.json({ ...t, modulos, check, ata: { link: `${baseUrlPublicaDpo(req)}/assinatura.html?t=${tk}`, ativa: t.ata_ativa !== 0, assinaturas: nAss ? nAss.n : 0 },
            divulgacao: divulgacaoMt(cfg) });
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar o treinamento.' }); }
});

app.post('/api/admin/treinamentos-mt', requireRole('admin'), async (req, res) => {
    const d = camposTreinamentoMt({ status: 'rascunho', modulos: [], ...req.body });
    if (!d.titulo) return res.status(400).json({ error: 'Dê um nome ao treinamento.' });
    try {
        const cols = Object.keys(d);
        const id = await new Promise((ok, ko) => db.run(`INSERT INTO treinamentos_mt (${cols.join(', ')}, created_by, updated_at) VALUES (${cols.map(() => '?').join(', ')}, ?, CURRENT_TIMESTAMP)`,
            [...cols.map(c => d[c]), req.user.userId], function (err) { err ? ko(err) : ok(this.lastID); }));
        res.json({ id, message: 'Treinamento criado (rascunho).' });
    } catch (e) { res.status(400).json({ error: 'Erro ao criar o treinamento.' }); }
});

app.put('/api/admin/treinamentos-mt/:id', requireRole('admin'), async (req, res) => {
    try {
        const t = await dbGet(`SELECT id FROM treinamentos_mt WHERE id = ?`, [req.params.id]);
        if (!t) return res.status(404).json({ error: 'Treinamento não encontrado.' });
        const d = camposTreinamentoMt(req.body || {});
        if (d.titulo !== undefined && !d.titulo) return res.status(400).json({ error: 'Dê um nome ao treinamento.' });
        if (req.body.marcarApresentado) d.apresentado_em = new Date().toISOString();
        const cols = Object.keys(d);
        if (!cols.length) return res.json({ message: 'Nada para salvar.' });
        await new Promise((ok, ko) => db.run(`UPDATE treinamentos_mt SET ${cols.map(c => c + ' = ?').join(', ')}, updated_at = CURRENT_TIMESTAMP WHERE id = ?`, [...cols.map(c => d[c]), t.id], e => e ? ko(e) : ok()));
        res.json({ message: 'Salvo!', salvoEm: new Date().toISOString() });
    } catch (e) { res.status(400).json({ error: 'Erro ao salvar o treinamento.' }); }
});

app.post('/api/admin/treinamentos-mt/:id/duplicar', requireRole('admin'), async (req, res) => {
    try {
        const t = await dbGet(`SELECT * FROM treinamentos_mt WHERE id = ?`, [req.params.id]);
        if (!t) return res.status(404).json({ error: 'Treinamento não encontrado.' });
        const id = await new Promise((ok, ko) => db.run(`INSERT INTO treinamentos_mt (titulo, objetivo, publico, duracao_min, local, instrutor, company_id, status, modulos, created_by, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, 'rascunho', ?, ?, CURRENT_TIMESTAMP)`,
            [t.titulo + ' (cópia)', t.objetivo, t.publico, t.duracao_min, t.local, t.instrutor, t.company_id, t.modulos, req.user.userId], function (err) { err ? ko(err) : ok(this.lastID); }));
        res.json({ id, message: 'Treinamento duplicado como rascunho.' });
    } catch (e) { res.status(400).json({ error: 'Erro ao duplicar.' }); }
});

app.delete('/api/admin/treinamentos-mt/:id', requireRole('admin'), async (req, res) => {
    try {
        const checks = await dbAll(`SELECT id FROM treinamentos_mt_checks WHERE treinamento_id = ?`, [req.params.id]);
        for (const c of checks) await new Promise(r => db.run(`DELETE FROM treinamentos_mt_respostas WHERE check_id = ?`, [c.id], () => r()));
        await new Promise(r => db.run(`DELETE FROM treinamentos_mt_checks WHERE treinamento_id = ?`, [req.params.id], () => r()));
        await new Promise(r => db.run(`DELETE FROM treinamentos_mt_presencas WHERE treinamento_id = ?`, [req.params.id], () => r()));
        await new Promise(r => db.run(`DELETE FROM treinamentos_mt WHERE id = ?`, [req.params.id], () => r()));
        res.json({ message: 'Treinamento excluído.' });
    } catch (e) { res.status(400).json({ error: 'Erro ao excluir.' }); }
});

// Upload de arquivo do módulo: guarda, extrai o texto (PPTX/DOCX) e tenta gerar PDF para apresentar.
app.post('/api/admin/treinamentos-mt/arquivo', requireRole('admin'), (req, res) => {
    uploadMaterialDpo.single('file')(req, res, async (err) => {
        if (err) return res.status(400).json({ error: err.code === 'LIMIT_FILE_SIZE' ? 'Arquivo muito grande (máximo 100MB).' : err.message });
        if (!req.file) return res.status(400).json({ error: 'Nenhum arquivo recebido.' });
        const nome = consertarAcentoMt(req.file.originalname || req.file.filename);
        const r = { url: '/uploads/' + req.file.filename, nome, mime: req.file.mimetype || '', pdfUrl: '', texto: '', slidesExtraidos: [] };
        try {
            if (/\.(pptx|docx)$/i.test(nome)) {
                const ex = extrairOfficeMt(fs.readFileSync(req.file.path), nome);
                if (ex.slides) { r.slidesExtraidos = ex.slides; r.texto = ex.slides.map((s, i) => `Slide ${i + 1}: ${s.titulo}\n${s.texto}`).join('\n\n'); }
                if (ex.texto) r.texto = ex.texto;
            } else if (/\.(txt|csv)$/i.test(nome)) r.texto = fs.readFileSync(req.file.path, 'utf8').slice(0, 60000);
            if (/\.(pptx?|docx?|odp|odt)$/i.test(nome)) r.pdfUrl = (await converterParaPdfMt(req.file.path)) || '';
            if (/\.pdf$/i.test(nome) || /pdf/.test(r.mime)) r.pdfUrl = r.url;
        } catch (e) { console.error('Treinamento — leitura do arquivo:', e.message); }
        r.texto = String(r.texto || '').slice(0, 60000);
        res.json(r);
    });
});

// Check automático SEM IA: usa os tópicos dos slides. Pergunta: "Sobre <título>, qual
// afirmação faz parte do treinamento?" — certa = um tópico do slide; erradas = tópicos
// de outros slides, transformados em afirmações que NÃO são daquele assunto.
function gerarCheckLocalMt(modulos, qtd, abertas) {
    const slides = [];
    modulos.filter(m => !m.oculto).forEach(m => {
        const lista = m.tipo === 'slides' ? (m.slides || []) : ((m.arquivo && m.arquivo.slidesExtraidos) || []);
        lista.filter(s => !s.oculto).forEach(s => {
            const topicos = String(s.texto || '').split('\n').map(l => l.replace(/^\s*[•\-\*–]\s*/, '').trim()).filter(l => l.length >= 12 && l.length <= 220);
            if (s.titulo && topicos.length) slides.push({ titulo: String(s.titulo).trim(), topicos });
        });
    });
    if (!slides.length) return [];
    const embaralhar = a => { const b = a.slice(); for (let i = b.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [b[i], b[j]] = [b[j], b[i]]; } return b; };
    const todos = slides.flatMap((s, i) => s.topicos.map(t => ({ t, i })));
    const perguntas = [];
    const ordem = embaralhar(slides.map((s, i) => i));
    const nObj = Math.max(0, qtd - abertas);
    for (const i of ordem) {
        if (perguntas.length >= nObj) break;
        const s = slides[i];
        const certa = s.topicos[Math.floor(Math.random() * s.topicos.length)];
        const erradas = embaralhar(todos.filter(x => x.i !== i && x.t !== certa)).slice(0, 3).map(x => x.t);
        if (erradas.length < 2) continue;
        const opcoes = embaralhar([certa, ...erradas]);
        perguntas.push({ texto: `Sobre "${s.titulo}", qual destas afirmações foi apresentada no treinamento?`, tipo: 'multipla', opcoes, correta: opcoes.indexOf(certa) });
    }
    embaralhar(slides).slice(0, Math.max(abertas, perguntas.length ? 0 : Math.min(qtd, 5))).forEach(s => perguntas.push({ texto: `Com suas palavras: o que você vai aplicar no dia a dia sobre "${s.titulo}"?`, tipo: 'aberta', opcoes: [], correta: null }));
    return perguntas.slice(0, qtd);
}

// Configurações do Ministrar Treinamento: redes da Impulsionar (aparecem no fim) e chave da IA.
const CHAVES_CFG_MT = ['impulsionar_instagram', 'impulsionar_whatsapp', 'impulsionar_site', 'impulsionar_instagram_link', 'impulsionar_whatsapp_link', 'anthropic_api_key', 'anthropic_model'];
function divulgacaoMt(cfg) {
    const insta = cfg.impulsionar_instagram || '', zap = cfg.impulsionar_whatsapp || '';
    const dig = zap.replace(/\D/g, '');
    return { instagram: insta, whatsapp: zap, site: cfg.impulsionar_site || '',
        instagramLink: cfg.impulsionar_instagram_link || (insta ? 'https://instagram.com/' + insta.replace(/^@/, '') : ''),
        whatsappLink: cfg.impulsionar_whatsapp_link || (dig ? 'https://wa.me/' + (dig.length <= 11 ? '55' + dig : dig) : '') };
}
async function lerConfigMt() {
    const rows = await dbAll(`SELECT key, value FROM integration_settings WHERE key IN (${CHAVES_CFG_MT.map(() => '?').join(',')})`, CHAVES_CFG_MT);
    return Object.fromEntries(rows.map(r => [r.key, r.value || '']));
}
app.get('/api/admin/treinamentos-mt-config', requireRole('admin'), async (req, res) => {
    try {
        const c = await lerConfigMt();
        res.json({ instagram: c.impulsionar_instagram || '', whatsapp: c.impulsionar_whatsapp || '', site: c.impulsionar_site || '', instagramLink: c.impulsionar_instagram_link || '', whatsappLink: c.impulsionar_whatsapp_link || '',
            iaAtiva: !!ANTHROPIC_API_KEY, iaOrigem: c.anthropic_api_key ? 'sistema' : (process.env.ANTHROPIC_API_KEY ? 'servidor' : ''), iaPreview: ANTHROPIC_API_KEY ? '••••' + ANTHROPIC_API_KEY.slice(-4) : '', modelo: ANTHROPIC_MODEL });
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar as configurações.' }); }
});
app.put('/api/admin/treinamentos-mt-config', requireRole('admin'), async (req, res) => {
    try {
        const salvar = (k, v) => new Promise(ok => db.run(`INSERT INTO integration_settings (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value`, [k, v], () => ok()));
        const insta = String(req.body.instagram || '').trim().replace(/^https?:\/\/(www\.)?instagram\.com\//i, '').replace(/\/$/, '').replace(/^@?/, '').slice(0, 60);
        await salvar('impulsionar_instagram', insta ? '@' + insta : '');
        await salvar('impulsionar_whatsapp', String(req.body.whatsapp || '').trim().slice(0, 30));
        await salvar('impulsionar_site', String(req.body.site || '').trim().slice(0, 120));
        const linkOk = v => { const t = String(v || '').trim(); return /^https?:\/\/\S+$/i.test(t) ? t.slice(0, 300) : ''; };
        await salvar('impulsionar_instagram_link', linkOk(req.body.instagramLink));
        await salvar('impulsionar_whatsapp_link', linkOk(req.body.whatsappLink));
        let aviso = '';
        if (req.body.chaveIa !== undefined && String(req.body.chaveIa).trim()) {
            const chave = String(req.body.chaveIa).trim();
            // Testa a chave antes de guardar.
            const r = await fetch('https://api.anthropic.com/v1/messages', { method: 'POST', headers: { 'x-api-key': chave, 'anthropic-version': '2023-06-01', 'Content-Type': 'application/json' },
                body: JSON.stringify({ model: ANTHROPIC_MODEL, max_tokens: 5, messages: [{ role: 'user', content: 'ok' }] }) }).catch(e => ({ ok: false, json: async () => ({ error: { message: e.message } }) }));
            if (!r.ok) { const j = await r.json().catch(() => ({})); return res.status(400).json({ error: 'A chave da IA não funcionou: ' + ((j.error && j.error.message) || 'erro') + '. As redes foram salvas.' }); }
            await salvar('anthropic_api_key', chave); ANTHROPIC_API_KEY = chave; aviso = ' IA conectada! ✅';
        }
        if (req.body.removerChaveIa) { await salvar('anthropic_api_key', ''); ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY || ''; }
        res.json({ message: 'Configurações salvas!' + aviso });
    } catch (e) { res.status(400).json({ error: 'Erro ao salvar as configurações.' }); }
});

// ----- Controle da apresentação pelo celular (link + Socket.IO) -----
const CONTROLES_MT = new Map(); // codigo -> { treinamentoId, titulo, criado }
app.post('/api/admin/treinamentos-mt/:id/controle', requireRole('admin'), async (req, res) => {
    try {
        const t = await dbGet(`SELECT id, titulo FROM treinamentos_mt WHERE id = ?`, [req.params.id]);
        if (!t) return res.status(404).json({ error: 'Treinamento não encontrado.' });
        const agora = Date.now();
        for (const [k, v] of CONTROLES_MT) if (agora - v.criado > 12 * 3600 * 1000) CONTROLES_MT.delete(k);
        let codigo = [...CONTROLES_MT.entries()].find(([k, v]) => v.treinamentoId === t.id && v.userId === req.user.userId)?.[0];
        if (!codigo) { codigo = crypto.randomBytes(6).toString('hex'); CONTROLES_MT.set(codigo, { treinamentoId: t.id, titulo: consertarAcentoMt(t.titulo), criado: agora, userId: req.user.userId }); }
        res.json({ codigo, link: `${baseUrlPublicaDpo(req)}/controle.html?c=${codigo}` });
    } catch (e) { res.status(500).json({ error: 'Erro ao gerar o controle.' }); }
});

// ----- Ata de presença com assinatura (link público) -----
async function garantirAtaTokenMt(t) {
    if (t.ata_token) return t.ata_token;
    const tk = 'at' + crypto.randomBytes(12).toString('hex');
    await new Promise(ok => db.run(`UPDATE treinamentos_mt SET ata_token = ? WHERE id = ?`, [tk, t.id], () => ok()));
    return tk;
}
app.get('/api/admin/treinamentos-mt/:id/ata', requireRole('admin'), async (req, res) => {
    try {
        const t = await dbGet(`SELECT * FROM treinamentos_mt WHERE id = ?`, [req.params.id]);
        if (!t) return res.status(404).json({ error: 'Treinamento não encontrado.' });
        const tk = await garantirAtaTokenMt(t);
        const lista = await dbAll(`SELECT id, nome, matricula, cargo, empresa, assinatura, created_at FROM treinamentos_mt_presencas WHERE treinamento_id = ? ORDER BY created_at ASC, id ASC`, [t.id]);
        res.json({ link: `${baseUrlPublicaDpo(req)}/assinatura.html?t=${tk}`, ativa: t.ata_ativa !== 0, presencas: lista });
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar a ata.' }); }
});
app.put('/api/admin/treinamentos-mt/:id/ata', requireRole('admin'), async (req, res) => {
    await new Promise(ok => db.run(`UPDATE treinamentos_mt SET ata_ativa = ? WHERE id = ?`, [req.body.ativa ? 1 : 0, req.params.id], () => ok()));
    res.json({ message: req.body.ativa ? 'Ata aberta para assinaturas.' : 'Ata encerrada.' });
});
app.delete('/api/admin/treinamentos-mt/presencas/:pid', requireRole('admin'), async (req, res) => {
    await new Promise(ok => db.run(`DELETE FROM treinamentos_mt_presencas WHERE id = ?`, [req.params.pid], () => ok()));
    res.json({ message: 'Assinatura removida.' });
});
app.get('/api/public/ata/:token', async (req, res) => {
    try {
        const t = await dbGet(`SELECT t.*, c.name as companyName FROM treinamentos_mt t LEFT JOIN companies c ON c.id = t.company_id WHERE t.ata_token = ?`, [String(req.params.token || '')]);
        if (!t) return res.status(404).json({ error: 'Ata não encontrada.' });
        if (t.ata_ativa === 0) return res.status(410).json({ error: 'Esta ata já foi encerrada.' });
        const ck = await dbGet(`SELECT token, ativo FROM treinamentos_mt_checks WHERE treinamento_id = ?`, [t.id]);
        const cfg = await lerConfigMt();
        res.json({ titulo: consertarAcentoMt(t.titulo), publico: ROTULOS_PUBLICO_MT[t.publico] || 'Todos', data: t.data_treinamento, instrutor: t.instrutor, local: t.local, empresa: t.companyName || '',
            checkLink: ck && ck.ativo ? `/retencao.html?t=${ck.token}` : null, instagram: cfg.impulsionar_instagram || '', whatsapp: cfg.impulsionar_whatsapp || '' });
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar a ata.' }); }
});
app.post('/api/public/ata/:token', async (req, res) => {
    try {
        const t = await dbGet(`SELECT * FROM treinamentos_mt WHERE ata_token = ?`, [String(req.params.token || '')]);
        if (!t) return res.status(404).json({ error: 'Ata não encontrada.' });
        if (t.ata_ativa === 0) return res.status(410).json({ error: 'Esta ata já foi encerrada.' });
        const nome = String(req.body.nome || '').trim().slice(0, 120);
        const assinatura = String(req.body.assinatura || '');
        if (nome.length < 3) return res.status(400).json({ error: 'Informe seu nome completo.' });
        if (!/^data:image\/png;base64,[A-Za-z0-9+/=]+$/.test(assinatura) || assinatura.length < 800) return res.status(400).json({ error: 'Faça sua assinatura no quadro.' });
        if (assinatura.length > 400000) return res.status(400).json({ error: 'Assinatura muito grande, limpe e assine de novo.' });
        const ja = await dbGet(`SELECT id FROM treinamentos_mt_presencas WHERE treinamento_id = ? AND lower(nome) = lower(?)`, [t.id, nome]);
        if (ja) return res.status(400).json({ error: 'Esse nome já assinou a ata deste treinamento.' });
        await new Promise((ok, ko) => db.run(`INSERT INTO treinamentos_mt_presencas (treinamento_id, nome, matricula, cargo, empresa, assinatura, ip) VALUES (?, ?, ?, ?, ?, ?, ?)`,
            [t.id, nome, String(req.body.matricula || '').trim().slice(0, 40) || null, String(req.body.cargo || '').trim().slice(0, 80) || null, String(req.body.empresa || '').trim().slice(0, 120) || null, assinatura, String(((req.headers || {})['x-forwarded-for']) || req.ip || '').split(',')[0].slice(0, 60)], e => e ? ko(e) : ok()));
        res.json({ message: 'Presença assinada! Obrigado.' });
    } catch (e) { res.status(400).json({ error: 'Erro ao registrar a assinatura.' }); }
});

// Monta o conteúdo do treinamento para a IA (texto + PDFs como documento).
async function conteudoParaIaMt(t, modulos) {
    const blocos = [];
    let texto = `TREINAMENTO: ${t.titulo}\nPúblico: ${ROTULOS_PUBLICO_MT[t.publico] || 'Todos'}\n${t.objetivo ? 'Objetivo: ' + t.objetivo + '\n' : ''}`;
    let pdfs = 0, bytesPdf = 0;
    for (const [i, m] of modulos.entries()) {
        if (m.oculto) continue;
        texto += `\n\n=== MÓDULO ${i + 1}: ${m.titulo} ===\n`;
        if (m.tipo === 'arquivo' && m.arquivo && (m.arquivo.slidesExtraidos || []).length) { texto += m.arquivo.slidesExtraidos.filter(x => !x.oculto).map((x, j) => `Slide ${j + 1}: ${x.titulo}\n${x.texto}`).join('\n\n'); continue; }
        if (m.tipo === 'slides') texto += (m.slides || []).filter(x => !x.oculto).map((s, j) => `Slide ${j + 1}: ${s.titulo}\n${s.texto}${s.notas ? '\n(Notas do instrutor: ' + s.notas + ')' : ''}`).join('\n\n');
        else if (m.texto) texto += m.texto;
        else if (m.arquivo && /\.pdf$/i.test(m.arquivo.pdfUrl || m.arquivo.url || '') && pdfs < 3) {
            const buf = await lerArquivoUploadMt(m.arquivo.pdfUrl || m.arquivo.url);
            if (buf && buf.length < 20 * 1024 * 1024 && bytesPdf + buf.length < 30 * 1024 * 1024) {
                blocos.push({ type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: buf.toString('base64') } });
                pdfs++; bytesPdf += buf.length; texto += `(conteúdo no PDF anexo "${m.arquivo.nome}")`;
            }
        } else if (m.arquivo && /^image\//.test(m.arquivo.mime || '')) {
            const buf = await lerArquivoUploadMt(m.arquivo.url);
            if (buf && buf.length < 4 * 1024 * 1024) { blocos.push({ type: 'image', source: { type: 'base64', media_type: m.arquivo.mime, data: buf.toString('base64') } }); texto += `(imagem anexa "${m.arquivo.nome}")`; }
        } else texto += `(arquivo "${m.arquivo ? m.arquivo.nome : ''}" sem texto legível)`;
    }
    return { blocos, texto: texto.slice(0, 120000) };
}

app.post('/api/admin/treinamentos-mt/:id/check/gerar', requireRole('admin'), async (req, res) => {
    try {
        const t = await dbGet(`SELECT * FROM treinamentos_mt WHERE id = ?`, [req.params.id]);
        if (!t) return res.status(404).json({ error: 'Treinamento não encontrado.' });
        const modulos = JSON.parse(t.modulos || '[]');
        if (!modulos.length) return res.status(400).json({ error: 'Adicione os módulos do treinamento antes de gerar o check.' });
        const qtd = Math.max(3, Math.min(20, Number(req.body.qtd) || 8));
        const abertas = Math.max(0, Math.min(5, Number(req.body.abertas) || 0));
        const { blocos, texto } = await conteudoParaIaMt(t, modulos);
        const local = () => {
            const perguntas = gerarCheckLocalMt(modulos, qtd, abertas);
            return perguntas.length ? { titulo: `Check de retenção — ${t.titulo}`, perguntas, gerado_por_ia: false } : null;
        };
        if (!ANTHROPIC_API_KEY) {
            const l = local();
            if (!l) return res.status(400).json({ error: 'Não achei texto suficiente nos módulos para montar o check. Adicione texto nos slides.' });
            return res.json({ ...l, aviso: 'Check montado automaticamente pelo conteúdo dos slides (sem IA). Para a IA criar perguntas mais elaboradas, cadastre a chave em ⚙️ Configurações.' });
        }
        const sistema = 'Você é especialista em treinamento corporativo e cria checks de retenção (avaliação de aprendizagem) em português do Brasil, com linguagem simples adequada ao público. Baseie-se SOMENTE no conteúdo fornecido. Responda APENAS com JSON válido, sem texto fora do JSON.';
        const pedido = `${texto}\n\n---\nCrie um check de retenção com ${qtd} perguntas sobre os pontos mais importantes do treinamento acima${abertas ? `, sendo ${abertas} abertas (dissertativas curtas) e o restante de múltipla escolha` : ', todas de múltipla escolha'}.\nMúltipla escolha: 4 alternativas, só 1 correta, alternativas plausíveis e de tamanho parecido, sem "todas as anteriores".\nCubra todos os módulos. Formato exato:\n{"titulo":"Check de retenção — ...","perguntas":[{"texto":"...","tipo":"multipla","opcoes":["...","...","...","..."],"correta":0},{"texto":"...","tipo":"aberta"}]}`;
        const conteudo = [...blocos, { type: 'text', text: pedido }];
        let resposta;
        try {
            const r = await fetch('https://api.anthropic.com/v1/messages', { method: 'POST', headers: { 'x-api-key': ANTHROPIC_API_KEY, 'anthropic-version': '2023-06-01', 'Content-Type': 'application/json' },
                body: JSON.stringify({ model: ANTHROPIC_MODEL, max_tokens: 4000, system: sistema, messages: [{ role: 'user', content: conteudo }] }) });
            const j = await r.json();
            if (!r.ok) {
                if (blocos.length) { // modelo sem suporte a PDF/imagem → tenta só com texto
                    resposta = await perguntarIA(sistema, pedido, 4000);
                } else throw new Error((j.error && j.error.message) || 'Erro na IA.');
            } else resposta = (j.content || []).map(b => b.text || '').join('\n');
        } catch (e) {
            const l = local();
            if (l) return res.json({ ...l, aviso: `A IA não respondeu (${e.message}). Montei o check automaticamente pelo conteúdo dos slides — revise e salve. Confira a chave da IA em ⚙️ Configurações.` });
            return res.status(400).json({ error: 'A IA não conseguiu gerar o check: ' + e.message });
        }
        const ini = resposta.indexOf('{'), fim = resposta.lastIndexOf('}');
        let obj; try { obj = JSON.parse(resposta.slice(ini, fim + 1)); } catch (e) { const l = local(); if (l) return res.json({ ...l, aviso: 'A IA respondeu num formato inesperado; montei o check pelo conteúdo dos slides.' }); return res.status(400).json({ error: 'A IA respondeu num formato inesperado. Tente gerar de novo.' }); }
        const perguntas = normalizarPerguntasCheckDpo(obj.perguntas || []);
        if (!perguntas.length) return res.status(400).json({ error: 'A IA não conseguiu montar perguntas com esse conteúdo. Adicione mais texto nos slides.' });
        res.json({ titulo: String(obj.titulo || `Check de retenção — ${t.titulo}`).slice(0, 200), perguntas, gerado_por_ia: true });
    } catch (e) { res.status(500).json({ error: 'Erro ao gerar o check de retenção.' }); }
});

app.put('/api/admin/treinamentos-mt/:id/check', requireRole('admin'), async (req, res) => {
    try {
        const t = await dbGet(`SELECT * FROM treinamentos_mt WHERE id = ?`, [req.params.id]);
        if (!t) return res.status(404).json({ error: 'Treinamento não encontrado.' });
        const perguntas = normalizarPerguntasCheckDpo(req.body.perguntas || []);
        if (!perguntas.length) return res.status(400).json({ error: 'O check precisa de pelo menos uma pergunta.' });
        const titulo = String(req.body.titulo || '').trim().slice(0, 200) || `Check de retenção — ${t.titulo}`;
        const atual = await dbGet(`SELECT * FROM treinamentos_mt_checks WHERE treinamento_id = ?`, [t.id]);
        if (atual) await new Promise(r => db.run(`UPDATE treinamentos_mt_checks SET titulo = ?, perguntas = ?, ativo = ?, gerado_por_ia = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?`, [titulo, JSON.stringify(perguntas), req.body.ativo === false ? 0 : 1, req.body.gerado_por_ia ? 1 : atual.gerado_por_ia, atual.id], () => r()));
        else await new Promise(r => db.run(`INSERT INTO treinamentos_mt_checks (treinamento_id, titulo, token, perguntas, gerado_por_ia) VALUES (?, ?, ?, ?, ?)`, [t.id, titulo, 'tr' + crypto.randomBytes(12).toString('hex'), JSON.stringify(perguntas), req.body.gerado_por_ia ? 1 : 0], () => r()));
        const c = await dbGet(`SELECT * FROM treinamentos_mt_checks WHERE treinamento_id = ?`, [t.id]);
        res.json({ message: 'Check de retenção salvo!', link: `${baseUrlPublicaDpo(req)}/retencao.html?t=${c.token}`, ativo: !!c.ativo });
    } catch (e) { res.status(400).json({ error: 'Erro ao salvar o check.' }); }
});

app.get('/api/admin/treinamentos-mt/:id/check/respostas', requireRole('admin'), async (req, res) => {
    try {
        const c = await dbGet(`SELECT * FROM treinamentos_mt_checks WHERE treinamento_id = ?`, [req.params.id]);
        if (!c) return res.json({ perguntas: [], respostas: [] });
        const r = await dbAll(`SELECT * FROM treinamentos_mt_respostas WHERE check_id = ? ORDER BY created_at DESC`, [c.id]);
        res.json({ perguntas: JSON.parse(c.perguntas || '[]'), respostas: r.map(x => ({ ...x, respostas: JSON.parse(x.respostas || '[]') })) });
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar as respostas.' }); }
});

// ======================================================================
// DPO — Cadastro de responsáveis (dono da ação) + usuários da empresa
// ======================================================================
function empresaDosResponsaveisDpo(req) {
    if (req.user.role === 'client_admin') return req.user.companyId;
    return Number(req.query.company_id || (req.body && req.body.company_id)) || null;
}
app.get('/api/dpo/responsaveis', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const cid = empresaDosResponsaveisDpo(req);
        if (!cid) return res.json({ cadastrados: [], usuarios: [], funcionarios: [] });
        const cadastrados = await dbAll(`SELECT * FROM dpo_responsaveis WHERE company_id = ? ORDER BY ativo DESC, nome COLLATE NOCASE`, [cid]);
        const usuarios = await dbAll(`SELECT id, name as nome, email, role FROM users WHERE company_id = ? AND role IN ('client_admin', 'autonomous', 'employee') AND name IS NOT NULL ORDER BY name COLLATE NOCASE`, [cid]);
        const funcionarios = await dbAll(`SELECT id, name as nome, role as cargo, email FROM employees WHERE company_id = ? AND name IS NOT NULL ORDER BY name COLLATE NOCASE`, [cid]).catch(() => []);
        // Contagem de ações por dono (para o cadastro mostrar quanto cada um tem).
        const contagem = await dbAll(`SELECT ap.owner, COUNT(*) n, SUM(CASE WHEN ap.status = 'concluida' THEN 0 ELSE 1 END) abertas
            FROM dpo_action_plans ap JOIN dpo_audit_cycles c ON c.id = ap.cycle_id WHERE c.company_id = ? AND ap.owner IS NOT NULL AND ap.owner <> '' GROUP BY ap.owner`, [cid]).catch(() => []);
        const mapa = {}; contagem.forEach(x => { mapa[String(x.owner).trim().toLowerCase()] = { total: x.n, abertas: x.abertas }; });
        const comAcoes = l => l.map(p => ({ ...p, acoes: mapa[String(p.nome || '').trim().toLowerCase()] || { total: 0, abertas: 0 } }));
        res.json({ cadastrados: comAcoes(cadastrados), usuarios: comAcoes(usuarios), funcionarios: comAcoes(funcionarios) });
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar os responsáveis.' }); }
});
function camposResponsavelDpo(b) {
    return { nome: String(b.nome || '').trim().slice(0, 120), cargo: String(b.cargo || '').trim().slice(0, 120) || null, area: String(b.area || '').trim().slice(0, 120) || null,
        email: String(b.email || '').trim().toLowerCase().slice(0, 160) || null, telefone: String(b.telefone || '').trim().slice(0, 40) || null };
}
app.post('/api/dpo/responsaveis', requireRole('admin', 'client_admin'), async (req, res) => {
    const cid = empresaDosResponsaveisDpo(req);
    if (!cid) return res.status(400).json({ error: 'Empresa não informada.' });
    const d = camposResponsavelDpo(req.body || {});
    if (!d.nome) return res.status(400).json({ error: 'Informe o nome do responsável.' });
    if (d.email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(d.email)) return res.status(400).json({ error: 'E-mail inválido.' });
    try {
        const ja = await dbGet(`SELECT id FROM dpo_responsaveis WHERE company_id = ? AND lower(nome) = lower(?)`, [cid, d.nome]);
        if (ja) return res.status(400).json({ error: 'Já existe um responsável com esse nome.' });
        const id = await new Promise((ok, ko) => db.run(`INSERT INTO dpo_responsaveis (company_id, nome, cargo, area, email, telefone, created_by) VALUES (?, ?, ?, ?, ?, ?, ?)`,
            [cid, d.nome, d.cargo, d.area, d.email, d.telefone, req.user.userId], function (err) { err ? ko(err) : ok(this.lastID); }));
        res.json({ id, nome: d.nome, message: 'Responsável cadastrado!' });
    } catch (e) { res.status(400).json({ error: 'Erro ao cadastrar o responsável.' }); }
});
async function responsavelComAcessoDpo(req, res) {
    const r = await dbGet(`SELECT * FROM dpo_responsaveis WHERE id = ?`, [req.params.id]);
    if (!r) { res.status(404).json({ error: 'Responsável não encontrado.' }); return null; }
    if (req.user.role === 'client_admin' && String(r.company_id) !== String(req.user.companyId)) { res.status(403).json({ error: 'Sem acesso.' }); return null; }
    return r;
}
app.put('/api/dpo/responsaveis/:id', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const r = await responsavelComAcessoDpo(req, res); if (!r) return;
        const d = camposResponsavelDpo({ ...r, ...req.body });
        if (!d.nome) return res.status(400).json({ error: 'Informe o nome do responsável.' });
        if (d.email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(d.email)) return res.status(400).json({ error: 'E-mail inválido.' });
        const ativo = req.body.ativo === undefined ? r.ativo : (req.body.ativo ? 1 : 0);
        await new Promise(ok => db.run(`UPDATE dpo_responsaveis SET nome = ?, cargo = ?, area = ?, email = ?, telefone = ?, ativo = ? WHERE id = ?`, [d.nome, d.cargo, d.area, d.email, d.telefone, ativo, r.id], () => ok()));
        // Renomeou: atualiza as ações que estavam com o nome antigo.
        let renomeadas = 0;
        if (d.nome !== r.nome && req.body.atualizarAcoes !== false) {
            renomeadas = await new Promise(ok => db.run(`UPDATE dpo_action_plans SET owner = ? WHERE owner = ? AND cycle_id IN (SELECT id FROM dpo_audit_cycles WHERE company_id = ?)`, [d.nome, r.nome, r.company_id], function () { ok(this.changes || 0); }));
        }
        res.json({ message: 'Responsável atualizado!' + (renomeadas ? ` (${renomeadas} ação(ões) atualizada(s) com o novo nome)` : '') });
    } catch (e) { res.status(400).json({ error: 'Erro ao atualizar o responsável.' }); }
});
app.delete('/api/dpo/responsaveis/:id', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const r = await responsavelComAcessoDpo(req, res); if (!r) return;
        await new Promise(ok => db.run(`DELETE FROM dpo_responsaveis WHERE id = ?`, [r.id], () => ok()));
        res.json({ message: 'Responsável removido (as ações dele continuam com o nome).' });
    } catch (e) { res.status(400).json({ error: 'Erro ao remover.' }); }
});

// ----- Página pública do check de retenção (sem login) -----
app.get('/api/public/retencao/:token', async (req, res) => {
    try {
        const check = await dbGet(`SELECT * FROM dpo_retention_checks WHERE token = ?`, [String(req.params.token || '')]);
        if (!check) {
            const tc = await dbGet(`SELECT c.*, t.titulo as treinamento, t.publico FROM treinamentos_mt_checks c JOIN treinamentos_mt t ON t.id = c.treinamento_id WHERE c.token = ?`, [String(req.params.token || '')]);
            if (!tc) return res.status(404).json({ error: 'Check de retenção não encontrado.' });
            if (!tc.ativo) return res.status(410).json({ error: 'Este check de retenção foi encerrado.' });
            return res.json({ titulo: tc.titulo, empresa: 'Treinamento Impulsionar · ' + (ROTULOS_PUBLICO_MT[tc.publico] || 'Todos'), logo: null, pilar: '', perguntaPilar: tc.treinamento,
                perguntas: JSON.parse(tc.perguntas || '[]').map(p => ({ texto: p.texto, tipo: p.tipo, opcoes: p.opcoes })) });
        }
        if (!check.ativo) return res.status(410).json({ error: 'Este check de retenção foi encerrado.' });
        const empresa = await dbGet(`SELECT name, logo_url FROM companies WHERE id = ?`, [check.company_id]);
        const achou = perguntaDoPilarDpo(check.pillar_key, check.question_numero);
        res.json({
            titulo: check.titulo,
            empresa: empresa ? empresa.name : '',
            logo: empresa ? empresa.logo_url : null,
            pilar: DPO_AMBEV_DATA[check.pillar_key] ? DPO_AMBEV_DATA[check.pillar_key].label : '',
            perguntaPilar: achou ? `${check.question_numero} ${achou.pergunta.questao}` : check.question_numero,
            // Sem o gabarito: quem responde não vê qual é a alternativa correta.
            perguntas: JSON.parse(check.perguntas || '[]').map(p => ({ texto: p.texto, tipo: p.tipo, opcoes: p.opcoes }))
        });
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar o check.' }); }
});

app.post('/api/public/retencao/:token', async (req, res) => {
    try {
        let check = await dbGet(`SELECT * FROM dpo_retention_checks WHERE token = ?`, [String(req.params.token || '')]);
        let tabelaResp = 'dpo_retention_responses';
        if (!check) { check = await dbGet(`SELECT * FROM treinamentos_mt_checks WHERE token = ?`, [String(req.params.token || '')]); tabelaResp = 'treinamentos_mt_respostas'; }
        if (!check) return res.status(404).json({ error: 'Check de retenção não encontrado.' });
        if (!check.ativo) return res.status(410).json({ error: 'Este check de retenção foi encerrado.' });
        const nome = String(req.body.nome || '').trim().slice(0, 120);
        const matricula = String(req.body.matricula || '').trim().slice(0, 40);
        if (!nome) return res.status(400).json({ error: 'Informe seu nome.' });
        const perguntas = JSON.parse(check.perguntas || '[]');
        const enviadas = Array.isArray(req.body.respostas) ? req.body.respostas : [];
        let acertos = 0, totalObjetivas = 0;
        const respostas = perguntas.map((p, i) => {
            const v = enviadas[i];
            if (p.tipo === 'multipla') {
                const idx = v === null || v === undefined || v === '' ? null : Number(v);
                const valido = idx !== null && idx >= 0 && idx < p.opcoes.length ? idx : null;
                if (p.correta !== null && p.correta !== undefined) { totalObjetivas++; if (valido === p.correta) acertos++; }
                return valido;
            }
            return String(v || '').trim().slice(0, 2000);
        });
        const faltando = perguntas.findIndex((p, i) => p.tipo === 'multipla' ? respostas[i] === null : !respostas[i]);
        if (faltando >= 0) return res.status(400).json({ error: `Responda a pergunta ${faltando + 1}.` });
        await new Promise((resolve, reject) => db.run(
            `INSERT INTO ${tabelaResp} (check_id, nome, matricula, respostas, acertos, total_objetivas) VALUES (?, ?, ?, ?, ?, ?)`,
            [check.id, nome, matricula || null, JSON.stringify(respostas), totalObjetivas ? acertos : null, totalObjetivas || null],
            (err) => err ? reject(err) : resolve()
        ));
        res.json({ message: 'Respostas enviadas! Obrigado.', acertos: totalObjetivas ? acertos : null, total: totalObjetivas || null });
    } catch (e) { res.status(400).json({ error: 'Erro ao enviar as respostas.' }); }
});

// ---------- DPO Ambev — AUTOAVALIAÇÃO MENSAL + régua de Selos DPO 2026 ----------
// Cada pergunta recebe 3, 1, 0 ou N/A. % = Σ(nota × peso) / Σ(3 × peso), só com as
// perguntas respondidas e que não são N/A.
const VALORES_AUTOAVALIACAO_DPO = ['3', '1', '0', 'na'];
const NOTA_MAXIMA_DPO = 3;
const MESES_DPO = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];

// Categorias da régua: Segurança; Gente e Gestão (juntos); pilares técnicos.
const CATEGORIA_PILAR_DPO = { seguranca: 'seg', gente: 'gg', gestao: 'gg', planejamento: 'tec', armazem: 'tec', frota: 'tec', entrega: 'tec' };

// Régua de Selos DPO 2026 (do mais alto para o mais baixo). "todos" = todos os
// pilares juntos (usado só no Route Basic, que vale apenas para revendas que
// nunca foram auditadas — 1ª auditoria em 2026).
const REGUA_SELOS_DPO = [
    { key: 'sustainable', label: 'Sustainable', seg: 85, gg: 73, tec: 73 },
    { key: 'certified', label: 'Certified', seg: 80, gg: 68, tec: 68 },
    { key: 'qualified', label: 'Qualified', seg: 73, gg: 57, tec: 64 },
    { key: 'route_basic', label: 'Route Basic', seg: 64, gg: 40, todos: 40, somentePrimeiraAuditoria: true },
    { key: 'not_qualified', label: 'Not Qualified' }
];

function referenciaAtualDpo() {
    const partes = new Intl.DateTimeFormat('en-CA', { timeZone: 'America/Sao_Paulo', year: 'numeric', month: '2-digit' }).formatToParts(new Date());
    return `${partes.find(p => p.type === 'year').value}-${partes.find(p => p.type === 'month').value}`;
}
function rotuloReferenciaDpo(ref) {
    const m = /^(\d{4})-(\d{2})$/.exec(String(ref || ''));
    return m ? `${MESES_DPO[Number(m[2]) - 1]}/${m[1]}` : String(ref || '');
}
const pctDpo = (pontos, max) => max ? Math.round(pontos * 1000 / max) / 10 : null;

// Nível de UM pilar: usa o limite da categoria dele (Segurança / Gente e Gestão /
// técnicos). Para o Route Basic, pilar técnico usa o limite de "todos pilares".
function nivelDoPilarDpo(pilarKey, pct, primeiraAuditoria) {
    if (pct === null || pct === undefined) return null;
    const cat = CATEGORIA_PILAR_DPO[pilarKey];
    for (const selo of REGUA_SELOS_DPO) {
        if (selo.key === 'not_qualified') return selo.key;
        if (selo.somentePrimeiraAuditoria && !primeiraAuditoria) continue;
        const limite = selo[cat] !== undefined ? selo[cat] : selo.todos;
        if (pct >= limite) return selo.key;
    }
    return 'not_qualified';
}

// Selo da OPERAÇÃO: o maior nível em que TODAS as categorias avaliadas batem o limite.
function nivelGeralDpo(cats, primeiraAuditoria) {
    const avaliadas = ['seg', 'gg', 'tec'].filter(c => cats[c] !== null && cats[c] !== undefined);
    if (!avaliadas.length) return null;
    for (const selo of REGUA_SELOS_DPO) {
        if (selo.key === 'not_qualified') return selo.key;
        if (selo.somentePrimeiraAuditoria && !primeiraAuditoria) continue;
        const criterios = selo.todos !== undefined ? ['seg', 'gg'].filter(c => avaliadas.includes(c)).map(c => cats[c] >= selo[c]).concat([cats.todos >= selo.todos])
                                                  : avaliadas.map(c => cats[c] >= selo[c]);
        if (criterios.every(Boolean)) return selo.key;
    }
    return 'not_qualified';
}

// Monta o resumo da operação (por pilar, por bloco e por categoria).
// respostas: { 'pilar:numero': '1' | '3' | 'na' }
function calcularResumoAutoavaliacaoDpo(respostas, pilares, primeiraAuditoria) {
    const somaCat = { seg: [0, 0], gg: [0, 0], tec: [0, 0], todos: [0, 0] };
    let totalPerguntas = 0, totalRespondidas = 0;
    const mandatoriasEm1 = [];
    const porPilar = pilares.filter(k => DPO_AMBEV_DATA[k]).map(pilarKey => {
        const info = DPO_AMBEV_DATA[pilarKey];
        let pontosP = 0, maxP = 0, respP = 0, naP = 0, totalP = 0;
        const grupos = info.grupos.map(g => {
            let pontos = 0, max = 0, resp = 0, na = 0;
            g.perguntas.forEach(q => {
                totalP++;
                const v = respostas[`${pilarKey}:${q.numero}`];
                if (v === undefined || v === null) return;
                resp++;
                if (v === 'na') { na++; return; }
                const peso = Number(q.peso) || 1;
                pontos += Number(v) * peso; max += NOTA_MAXIMA_DPO * peso;
                if (q.mandatoria && (v === '1' || v === '0')) mandatoriasEm1.push({ pilarKey, pilar: info.label, numero: q.numero, questao: q.questao, nota: v });
            });
            pontosP += pontos; maxP += max; respP += resp; naP += na;
            const pct = pctDpo(pontos, max);
            return { numero: g.numero, titulo: g.titulo, pct, nivel: nivelDoPilarDpo(pilarKey, pct, primeiraAuditoria), respondidas: resp, na, total: g.perguntas.length };
        });
        const cat = CATEGORIA_PILAR_DPO[pilarKey];
        somaCat[cat][0] += pontosP; somaCat[cat][1] += maxP;
        somaCat.todos[0] += pontosP; somaCat.todos[1] += maxP;
        totalPerguntas += totalP; totalRespondidas += respP;
        const pct = pctDpo(pontosP, maxP);
        return { key: pilarKey, numero: DPO_PILARES_ORDEM.indexOf(pilarKey) + 1, label: info.label, categoria: cat, pct, nivel: nivelDoPilarDpo(pilarKey, pct, primeiraAuditoria), respondidas: respP, na: naP, total: totalP, grupos };
    });
    const categorias = { seg: pctDpo(...somaCat.seg), gg: pctDpo(...somaCat.gg), tec: pctDpo(...somaCat.tec), todos: pctDpo(...somaCat.todos) };
    // Selo da OPERAÇÃO = sempre o MENOR nível entre os pilares já avaliados.
    const ordemNivel = k => REGUA_SELOS_DPO.findIndex(x => x.key === k); // 0 = mais alto
    const avaliados = porPilar.filter(p => p.nivel);
    const limitantes = avaliados.length ? avaliados.filter(p => ordemNivel(p.nivel) === Math.max(...avaliados.map(x => ordemNivel(x.nivel)))) : [];
    const nivelGeral = limitantes.length ? limitantes[0].nivel : null;
    return {
        pilares: porPilar, categorias, nivelGeral, primeiraAuditoria: !!primeiraAuditoria, totalPerguntas, totalRespondidas, mandatoriasEm1,
        pilaresLimitantes: limitantes.map(p => ({ key: p.key, label: p.label, pct: p.pct })),
        pilaresAvaliados: avaliados.length, totalPilares: porPilar.length
    };
}

async function respostasDaAutoavaliacaoDpo(assessmentId) {
    const linhas = await dbAll(`SELECT question_key, valor FROM dpo_self_answers WHERE assessment_id = ?`, [assessmentId]);
    const mapa = {}; linhas.forEach(l => { mapa[l.question_key] = l.valor; });
    return mapa;
}

async function primeiraAuditoriaDaEmpresaDpo(companyId) {
    const emp = await dbGet(`SELECT dpo_primeira_auditoria FROM companies WHERE id = ?`, [companyId]);
    return !!(emp && emp.dpo_primeira_auditoria);
}

// Resolve a empresa para as rotas da autoavaliação (sem pilar específico).
async function resolverEmpresaAutoavaliacaoDpo(req, res, companyIdInformado) {
    const companyId = req.user.role === 'client_admin' ? req.user.companyId : companyIdInformado;
    if (!companyId) { res.status(400).json({ error: 'Informe a empresa (company_id).' }); return null; }
    if (!(await empresaTemPastaDpo(req, res, companyId, 'autoavaliacao'))) return null;
    return companyId;
}

async function obterAutoavaliacaoComAcesso(req, res, id) {
    const av = await dbGet(`SELECT * FROM dpo_self_assessments WHERE id = ?`, [id]);
    if (!av) { res.status(404).json({ error: 'Autoavaliação não encontrada.' }); return null; }
    if (req.user.role === 'client_admin' && String(av.company_id) !== String(req.user.companyId)) { res.status(403).json({ error: 'Esta autoavaliação não pertence à sua empresa.' }); return null; }
    if (!(await empresaTemPastaDpo(req, res, av.company_id, 'autoavaliacao'))) return null;
    return av;
}

// Nível de cada pilar pela ÚLTIMA autoavaliação em que ele foi avaliado (usado
// nas cores dos botões dos pilares e na pasta de Gestão), + selo geral da última.
async function niveisAtuaisDaEmpresaDpo(companyId) {
    const ativos = await pilaresAtivosDaEmpresa(companyId);
    const primeira = await primeiraAuditoriaDaEmpresaDpo(companyId);
    const avaliacoes = await dbAll(`SELECT * FROM dpo_self_assessments WHERE company_id = ? ORDER BY referencia DESC`, [companyId]);
    const pilares = {};
    let geral = null;
    for (const av of avaliacoes) {
        const respostas = await respostasDaAutoavaliacaoDpo(av.id);
        if (!Object.keys(respostas).length) continue;
        const resumo = calcularResumoAutoavaliacaoDpo(respostas, ativos, primeira);
        if (!geral) geral = { assessmentId: av.id, referencia: av.referencia, referenciaLabel: rotuloReferenciaDpo(av.referencia), status: av.status, nivelGeral: resumo.nivelGeral, categorias: resumo.categorias, pilaresLimitantes: resumo.pilaresLimitantes, pilaresAvaliados: resumo.pilaresAvaliados, totalPilares: resumo.totalPilares };
        resumo.pilares.forEach(p => {
            if (pilares[p.key] || p.respondidas === 0) return;
            const notas = {};
            Object.keys(respostas).filter(k => k.startsWith(p.key + ':')).forEach(k => { notas[k] = respostas[k]; });
            pilares[p.key] = { pct: p.pct, nivel: p.nivel, respondidas: p.respondidas, total: p.total, referencia: av.referencia, referenciaLabel: rotuloReferenciaDpo(av.referencia), notas };
        });
        if (ativos.every(k => pilares[k])) break;
    }
    return { geral, pilares, primeiraAuditoria: primeira };
}

app.get('/api/dpo/niveis', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const companyId = req.user.role === 'client_admin' ? req.user.companyId : req.query.company_id;
        if (!companyId) return res.status(400).json({ error: 'Informe a empresa (company_id).' });
        res.json({ regua: REGUA_SELOS_DPO, ...(await niveisAtuaisDaEmpresaDpo(companyId)) });
    } catch (e) {
        console.error('Erro ao calcular níveis DPO:', e.message);
        res.status(500).json({ error: 'Erro ao calcular os níveis dos pilares.' });
    }
});

app.get('/api/dpo/autoavaliacoes', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const companyId = await resolverEmpresaAutoavaliacaoDpo(req, res, req.query.company_id);
        if (!companyId) return;
        const empresa = await dbGet(`SELECT name, dpo_primeira_auditoria FROM companies WHERE id = ?`, [companyId]);
        const ativos = await pilaresAtivosDaEmpresa(companyId);
        const primeira = !!(empresa && empresa.dpo_primeira_auditoria);
        const avaliacoes = await dbAll(`SELECT * FROM dpo_self_assessments WHERE company_id = ? ORDER BY referencia DESC`, [companyId]);
        const lista = [];
        for (const av of avaliacoes) {
            const r = calcularResumoAutoavaliacaoDpo(await respostasDaAutoavaliacaoDpo(av.id), ativos, primeira);
            lista.push({
                id: av.id, referencia: av.referencia, referenciaLabel: rotuloReferenciaDpo(av.referencia), status: av.status, closed_at: av.closed_at,
                nivelGeral: r.nivelGeral, categorias: r.categorias, totalPerguntas: r.totalPerguntas, totalRespondidas: r.totalRespondidas,
                pilares: r.pilares.map(p => ({ key: p.key, pct: p.pct, nivel: p.nivel }))
            });
        }
        const refAtual = referenciaAtualDpo();
        res.json({
            companyId: Number(companyId), empresa: empresa ? empresa.name : '', primeiraAuditoria: primeira, pilaresAtivos: ativos, regua: REGUA_SELOS_DPO,
            referenciaAtual: refAtual, referenciaAtualLabel: rotuloReferenciaDpo(refAtual), existeMesAtual: avaliacoes.some(a => a.referencia === refAtual), lista
        });
    } catch (e) {
        console.error('Erro ao listar autoavaliações DPO:', e.message);
        res.status(500).json({ error: 'Erro ao carregar as autoavaliações.' });
    }
});

// Master: todas as autoavaliações feitas, de todas as revendas com pacote contratado.
app.get('/api/admin/dpo/autoavaliacoes', requireRole('admin'), async (req, res) => {
    try {
        const empresas = await dbAll(`SELECT id, name, dpo_primeira_auditoria FROM companies ORDER BY name ASC`);
        const resultado = [];
        for (const emp of empresas) {
            const ativos = await pilaresAtivosDaEmpresa(emp.id);
            if (!ativos.length) continue;
            const avaliacoes = await dbAll(`SELECT a.*, u.name as aprovadoPor FROM dpo_self_assessments a LEFT JOIN users u ON u.id = a.approved_by WHERE a.company_id = ? ORDER BY a.referencia DESC`, [emp.id]);
            const lista = [];
            for (const av of avaliacoes) {
                const r = calcularResumoAutoavaliacaoDpo(await respostasDaAutoavaliacaoDpo(av.id), ativos, !!emp.dpo_primeira_auditoria);
                lista.push({ id: av.id, referencia: av.referencia, referenciaLabel: rotuloReferenciaDpo(av.referencia), status: av.status, approved_at: av.approved_at, aprovadoPor: av.aprovadoPor,
                    nivelGeral: r.nivelGeral, pct: r.categorias.todos, totalPerguntas: r.totalPerguntas, totalRespondidas: r.totalRespondidas });
            }
            resultado.push({ companyId: emp.id, empresa: emp.name, pilaresAtivos: ativos, avaliacoes: lista });
        }
        res.json(resultado);
    } catch (e) {
        console.error('Erro ao listar autoavaliações (Master):', e.message);
        res.status(500).json({ error: 'Erro ao carregar as autoavaliações.' });
    }
});

// Abre a autoavaliação do mês atual (ou devolve a que já existe).
app.post('/api/dpo/autoavaliacoes', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const companyId = await resolverEmpresaAutoavaliacaoDpo(req, res, req.body.company_id);
        if (!companyId) return;
        const ativos = await pilaresAtivosDaEmpresa(companyId);
        if (!ativos.length) return res.status(400).json({ error: 'A empresa ainda não tem nenhum pilar do DPO liberado.' });
        const referencia = referenciaAtualDpo();
        const existente = await dbGet(`SELECT id FROM dpo_self_assessments WHERE company_id = ? AND referencia = ?`, [companyId, referencia]);
        if (existente) return res.json({ id: existente.id, message: 'A autoavaliação deste mês já existe — abrindo.' });
        const id = await new Promise((resolve, reject) => db.run(
            `INSERT INTO dpo_self_assessments (company_id, referencia, created_by) VALUES (?, ?, ?)`,
            [companyId, referencia, req.user.userId], function (err) { err ? reject(err) : resolve(this.lastID); }
        ));
        res.json({ id, message: `Autoavaliação de ${rotuloReferenciaDpo(referencia)} aberta!` });
    } catch (e) { res.status(400).json({ error: 'Erro ao abrir a autoavaliação.' }); }
});

function textosNotasDpo(explicacao) {
    const textos = {};
    String(explicacao || '').replace(/\r/g, '').split(/\n\s*\n/).forEach(bloco => {
        const m = bloco.match(/^\s*(\d+)\s*[-.:)]\s*([\s\S]*)$/);
        if (m) textos[m[1]] = m[2].replace(/\n/g, ' ').trim();
    });
    return textos;
}

app.get('/api/dpo/autoavaliacoes/:id', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const av = await obterAutoavaliacaoComAcesso(req, res, req.params.id);
        if (!av) return;
        const ativos = await pilaresAtivosDaEmpresa(av.company_id);
        const primeira = await primeiraAuditoriaDaEmpresaDpo(av.company_id);
        const respostas = await respostasDaAutoavaliacaoDpo(av.id);
        const resumo = calcularResumoAutoavaliacaoDpo(respostas, ativos, primeira);
        const anterior = await dbGet(`SELECT id, referencia FROM dpo_self_assessments WHERE company_id = ? AND referencia < ? ORDER BY referencia DESC LIMIT 1`, [av.company_id, av.referencia]);
        const resumoAnterior = anterior ? calcularResumoAutoavaliacaoDpo(await respostasDaAutoavaliacaoDpo(anterior.id), ativos, primeira) : null;
        const empresa = await dbGet(`SELECT name FROM companies WHERE id = ?`, [av.company_id]);
        const aprovador = av.approved_by ? await dbGet(`SELECT name FROM users WHERE id = ?`, [av.approved_by]) : null;
        const eventos = await dbAll(`SELECT e.acao, e.detalhe, e.created_at, u.name as autorNome, u.role as autorPapel FROM dpo_self_assessment_events e LEFT JOIN users u ON u.id = e.user_id WHERE e.assessment_id = ? ORDER BY e.created_at DESC, e.id DESC LIMIT 60`, [av.id]);
        const chamadoAberto = await dbGet(`SELECT id, status FROM chamados WHERE ref_tipo = 'autoavaliacao' AND ref_id = ? AND status NOT IN ('resolvido', 'fechado') ORDER BY id DESC LIMIT 1`, [av.id]);
        const pilares = ativos.filter(k => DPO_AMBEV_DATA[k]).sort((a, b) => DPO_PILARES_ORDEM.indexOf(a) - DPO_PILARES_ORDEM.indexOf(b)).map(pilarKey => ({
            key: pilarKey, numero: DPO_PILARES_ORDEM.indexOf(pilarKey) + 1, label: DPO_AMBEV_DATA[pilarKey].label,
            grupos: DPO_AMBEV_DATA[pilarKey].grupos.map(g => ({
                numero: g.numero, titulo: g.titulo,
                perguntas: g.perguntas.map(q => {
                    const t = textosNotasDpo(q.explicacao_pontos);
                    return {
                        questionKey: `${pilarKey}:${q.numero}`, numero: q.numero, questao: q.questao, mandatoria: !!q.mandatoria, peso: q.peso,
                        verificacao: q.verificacao || '', how_to_check: q.how_to_check || '', texto0: t['0'] || '', texto1: t['1'] || '', texto3: t['3'] || '',
                        valor: respostas[`${pilarKey}:${q.numero}`] || null
                    };
                })
            }))
        }));
        res.json({
            id: av.id, companyId: av.company_id, empresa: empresa ? empresa.name : '', referencia: av.referencia, referenciaLabel: rotuloReferenciaDpo(av.referencia),
            status: av.status, closed_at: av.closed_at, approved_at: av.approved_at, aprovadoPor: aprovador ? aprovador.name : null,
            eventos, chamadoAberto: chamadoAberto || null, regua: REGUA_SELOS_DPO, resumo, pilares,
            anterior: anterior ? { id: anterior.id, referencia: anterior.referencia, referenciaLabel: rotuloReferenciaDpo(anterior.referencia), resumo: resumoAnterior } : null
        });
    } catch (e) {
        console.error('Erro ao carregar autoavaliação DPO:', e.message);
        res.status(500).json({ error: 'Erro ao carregar a autoavaliação.' });
    }
});

app.put('/api/dpo/autoavaliacoes/:id/respostas', requireRole('admin', 'client_admin'), async (req, res) => {
    const { questionKey } = req.body;
    const valor = req.body.valor === null || req.body.valor === '' || req.body.valor === undefined ? null : String(req.body.valor).toLowerCase();
    if (valor !== null && !VALORES_AUTOAVALIACAO_DPO.includes(valor)) return res.status(400).json({ error: 'Nota inválida. Use 3, 1, 0 ou N/A.' });
    try {
        const av = await obterAutoavaliacaoComAcesso(req, res, req.params.id);
        if (!av) return;
        if (req.user.role !== 'admin' && av.status !== 'em_andamento') {
            return res.status(400).json({ error: av.status === 'aprovada'
                ? 'Esta autoavaliação já foi aprovada — só o Master pode alterar. Abra um chamado pedindo ajuste.'
                : 'A autoavaliação está salva. Clique em "Editar" para alterar as notas.' });
        }
        const [pilarKey, numero] = String(questionKey || '').split(':');
        const ativos = await pilaresAtivosDaEmpresa(av.company_id);
        if (!ativos.includes(pilarKey) || !perguntaDoPilarDpo(pilarKey, numero)) return res.status(400).json({ error: 'Pergunta inválida.' });
        if (valor === null) {
            await new Promise((resolve, reject) => db.run(`DELETE FROM dpo_self_answers WHERE assessment_id = ? AND question_key = ?`, [av.id, questionKey], (err) => err ? reject(err) : resolve()));
        } else {
            await new Promise((resolve, reject) => db.run(
                `INSERT INTO dpo_self_answers (assessment_id, question_key, valor, updated_by, updated_at) VALUES (?, ?, ?, ?, CURRENT_TIMESTAMP)
                 ON CONFLICT(assessment_id, question_key) DO UPDATE SET valor = excluded.valor, updated_by = excluded.updated_by, updated_at = CURRENT_TIMESTAMP`,
                [av.id, questionKey, valor, req.user.userId], (err) => err ? reject(err) : resolve()
            ));
        }
        if (req.user.role === 'admin' && av.status !== 'em_andamento') registrarEventoAutoavaliacaoDpo(av.id, 'ajuste_master', `${questionKey} → ${valor === null ? 'sem nota' : valor === 'na' ? 'N/A' : valor}`, req.user.userId);
        const resumo = calcularResumoAutoavaliacaoDpo(await respostasDaAutoavaliacaoDpo(av.id), ativos, await primeiraAuditoriaDaEmpresaDpo(av.company_id));
        res.json({ message: 'Nota salva!', resumo });
    } catch (e) { res.status(400).json({ error: 'Erro ao salvar a nota.' }); }
});

function registrarEventoAutoavaliacaoDpo(assessmentId, acao, detalhe, userId) {
    db.run(`INSERT INTO dpo_self_assessment_events (assessment_id, acao, detalhe, user_id) VALUES (?, ?, ?, ?)`, [assessmentId, acao, detalhe || null, userId || null], () => {});
}

function notificarMasters(title, message, link) {
    db.all(`SELECT id FROM users WHERE role = 'admin'`, [], (err, admins) => {
        if (!err && admins) admins.forEach(a => notificar(a.id, title, message, link));
    });
}

// Salvar (trava para conferência) / Editar (volta a editar) / Aprovar (final,
// só o Master altera depois) / Reabrir (só Master, devolve para a empresa editar).
app.post('/api/dpo/autoavaliacoes/:id/acao', requireRole('admin', 'client_admin'), async (req, res) => {
    const { acao } = req.body;
    try {
        const av = await obterAutoavaliacaoComAcesso(req, res, req.params.id);
        if (!av) return;
        const ehMaster = req.user.role === 'admin';
        const ref = rotuloReferenciaDpo(av.referencia);
        let novo, msg;
        if (acao === 'salvar') {
            if (av.status !== 'em_andamento') return res.status(400).json({ error: 'Só dá para salvar uma autoavaliação em edição.' });
            novo = 'salva'; msg = 'Autoavaliação salva!';
        } else if (acao === 'editar') {
            if (av.status === 'aprovada') return res.status(400).json({ error: ehMaster ? 'Use "Reabrir para a empresa".' : 'Autoavaliação aprovada — só o Master pode alterar. Abra um chamado pedindo ajuste.' });
            if (av.status !== 'salva') return res.status(400).json({ error: 'A autoavaliação já está em edição.' });
            novo = 'em_andamento'; msg = 'Autoavaliação liberada para edição.';
        } else if (acao === 'aprovar') {
            if (av.status === 'aprovada') return res.status(400).json({ error: 'Esta autoavaliação já está aprovada.' });
            novo = 'aprovada'; msg = 'Autoavaliação aprovada!';
        } else if (acao === 'reabrir') {
            if (!ehMaster) return res.status(403).json({ error: 'Só o Master pode reabrir uma autoavaliação aprovada. Abra um chamado pedindo ajuste.' });
            if (av.status !== 'aprovada') return res.status(400).json({ error: 'A autoavaliação não está aprovada.' });
            novo = 'em_andamento'; msg = 'Autoavaliação reaberta para a empresa ajustar.';
        } else return res.status(400).json({ error: 'Ação inválida.' });

        await new Promise((resolve, reject) => db.run(
            `UPDATE dpo_self_assessments SET status = ?,
                approved_by = CASE WHEN ? = 'aprovada' THEN ? ELSE (CASE WHEN ? = 'em_andamento' THEN NULL ELSE approved_by END) END,
                approved_at = CASE WHEN ? = 'aprovada' THEN CURRENT_TIMESTAMP ELSE (CASE WHEN ? = 'em_andamento' THEN NULL ELSE approved_at END) END,
                closed_at = CASE WHEN ? = 'aprovada' THEN CURRENT_TIMESTAMP ELSE closed_at END
             WHERE id = ?`,
            [novo, novo, req.user.userId, novo, novo, novo, novo, av.id], (err) => err ? reject(err) : resolve()
        ));
        registrarEventoAutoavaliacaoDpo(av.id, acao, req.body.motivo || null, req.user.userId);

        if (acao === 'aprovar' && !ehMaster) {
            notificarMasters('DPO — autoavaliação aprovada', `A empresa aprovou a autoavaliação de ${ref}.`, 'dpoAgenda');
        }
        if (acao === 'reabrir') {
            notificarPorCompanyAdmins(av.company_id, 'DPO — autoavaliação reaberta', `O Master reabriu a autoavaliação de ${ref} para ajuste.`, 'dpoHome');
        }
        // Chamados de ajuste ligados a esta autoavaliação acompanham o fluxo.
        const chamadosAjuste = await dbAll(`SELECT id FROM chamados WHERE ref_tipo = 'autoavaliacao' AND ref_id = ? AND status NOT IN ('resolvido', 'fechado')`, [av.id]);
        for (const c of chamadosAjuste) {
            if (acao === 'reabrir') {
                await registrarMensagemSistemaChamado(c.id, `🔓 O Master liberou a autoavaliação de ${ref} para ajuste. Faça as alterações e aprove novamente.`, 'aguardando_empresa');
            } else if (acao === 'aprovar') {
                await registrarMensagemSistemaChamado(c.id, `✅ A autoavaliação de ${ref} foi aprovada novamente após o ajuste.`, 'resolvido');
            }
        }
        res.json({ message: msg, status: novo });
    } catch (e) {
        console.error('Erro na ação da autoavaliação DPO:', e.message);
        res.status(400).json({ error: 'Erro ao atualizar a autoavaliação.' });
    }
});

app.delete('/api/dpo/autoavaliacoes/:id', requireRole('admin'), async (req, res) => {
    try {
        await new Promise((resolve, reject) => db.run(`DELETE FROM dpo_self_answers WHERE assessment_id = ?`, [req.params.id], (err) => err ? reject(err) : resolve()));
        await new Promise((resolve, reject) => db.run(`DELETE FROM dpo_self_assessments WHERE id = ?`, [req.params.id], (err) => err ? reject(err) : resolve()));
        res.json({ message: 'Autoavaliação excluída!' });
    } catch (e) { res.status(400).json({ error: 'Erro ao excluir a autoavaliação.' }); }
});

// Excel de uma autoavaliação: resumo por categoria/pilar/bloco + todas as notas.
const ROTULO_NIVEL_DPO = Object.fromEntries(REGUA_SELOS_DPO.map(s => [s.key, s.label]));
const ROTULO_CATEGORIA_DPO = { seg: 'Segurança', gg: 'Gente e Gestão', tec: 'Pilares técnicos', todos: 'Todos os pilares' };
const fmtPctDpo = v => v === null || v === undefined ? '—' : `${String(v).replace('.', ',')}%`;

async function gerarExcelAutoavaliacaoDpo(av) {
    const ativos = await pilaresAtivosDaEmpresa(av.company_id);
    const respostas = await respostasDaAutoavaliacaoDpo(av.id);
    const r = calcularResumoAutoavaliacaoDpo(respostas, ativos, await primeiraAuditoriaDaEmpresaDpo(av.company_id));
    const empresa = await dbGet(`SELECT name FROM companies WHERE id = ?`, [av.company_id]);
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Impulsionar V4';
    const resumo = workbook.addWorksheet('Resumo da Operação');
    resumo.columns = [{ header: 'Nível', key: 'a', width: 34 }, { header: 'Item', key: 'b', width: 38 }, { header: '%', key: 'c', width: 12 }, { header: 'Selo / Nível', key: 'd', width: 18 }, { header: 'Respondidas', key: 'e', width: 13 }, { header: 'N/A', key: 'f', width: 8 }];
    estilizarCabecalhoExcelDpo(resumo, 'F');
    resumo.addRow({ a: 'Operação', b: `${empresa ? empresa.name : ''} — ${rotuloReferenciaDpo(av.referencia)}`, c: fmtPctDpo(r.categorias.todos), d: r.nivelGeral ? ROTULO_NIVEL_DPO[r.nivelGeral] : '—', e: `${r.totalRespondidas}/${r.totalPerguntas}` }).font = { bold: true };
    ['seg', 'gg', 'tec'].forEach(c => resumo.addRow({ a: 'Categoria', b: ROTULO_CATEGORIA_DPO[c], c: fmtPctDpo(r.categorias[c]) }));
    r.pilares.forEach(p => {
        resumo.addRow({ a: 'Pilar', b: `${p.numero}. ${p.label}`, c: fmtPctDpo(p.pct), d: p.nivel ? ROTULO_NIVEL_DPO[p.nivel] : '—', e: `${p.respondidas}/${p.total}`, f: p.na }).font = { bold: true };
        p.grupos.forEach(g => resumo.addRow({ a: `   Bloco — ${p.label}`, b: `${g.numero} ${g.titulo}`, c: fmtPctDpo(g.pct), d: g.nivel ? ROTULO_NIVEL_DPO[g.nivel] : '—', e: `${g.respondidas}/${g.total}`, f: g.na }));
    });
    const notas = workbook.addWorksheet('Notas', { views: [{ state: 'frozen', ySplit: 1 }] });
    notas.columns = [{ header: 'Pilar', key: 'pilar', width: 22 }, { header: 'Bloco', key: 'bloco', width: 34 }, { header: 'Nº', key: 'numero', width: 8 }, { header: 'Pergunta', key: 'pergunta', width: 45 }, { header: 'Mandatória', key: 'mand', width: 12 }, { header: 'Peso', key: 'peso', width: 8 }, { header: 'Nota', key: 'nota', width: 10 }];
    estilizarCabecalhoExcelDpo(notas, 'G');
    ativos.filter(k => DPO_AMBEV_DATA[k]).forEach(k => DPO_AMBEV_DATA[k].grupos.forEach(g => g.perguntas.forEach(q => {
        const v = respostas[`${k}:${q.numero}`];
        notas.addRow({ pilar: DPO_AMBEV_DATA[k].label, bloco: `${g.numero} ${g.titulo}`, numero: q.numero, pergunta: q.questao, mand: q.mandatoria ? 'Sim' : 'Não', peso: q.peso, nota: v === 'na' ? 'N/A' : (v || 'Não avaliada') });
    })));
    return { buffer: await workbook.xlsx.writeBuffer(), nome: `autoavaliacao-dpo-${av.referencia}-${(empresa ? empresa.name : 'empresa').replace(/[^a-z0-9]+/gi, '-')}.xlsx` };
}

app.get('/api/dpo/autoavaliacoes/:id/export', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const av = await obterAutoavaliacaoComAcesso(req, res, req.params.id);
        if (!av) return;
        const { buffer, nome } = await gerarExcelAutoavaliacaoDpo(av);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${nome}"`);
        res.send(Buffer.from(buffer));
    } catch (e) {
        console.error('Erro ao exportar autoavaliação DPO:', e.message);
        res.status(500).json({ error: 'Erro ao exportar a autoavaliação.' });
    }
});

// ---------- CHAMADOS (suporte da empresa para o Master) ----------
const CATEGORIAS_CHAMADO = {
    ajuste_autoavaliacao: 'Ajuste de autoavaliação DPO',
    duvida_dpo: 'Dúvida sobre o DPO',
    suporte_tecnico: 'Problema no sistema',
    financeiro: 'Financeiro / pagamento',
    sugestao: 'Sugestão de melhoria',
    liberar_ferramenta: 'Liberação de ferramenta Impulsionar',
    outro: 'Outro assunto'
};
const PRIORIDADES_CHAMADO = ['baixa', 'media', 'alta', 'urgente'];
const STATUS_CHAMADO = { aberto: 'Aberto', em_atendimento: 'Em atendimento', aguardando_empresa: 'Aguardando empresa', resolvido: 'Resolvido', fechado: 'Fechado' };
const numeroChamado = id => '#' + String(id).padStart(5, '0');

async function registrarMensagemSistemaChamado(chamadoId, texto, novoStatus) {
    await new Promise((resolve) => db.run(`INSERT INTO chamado_mensagens (chamado_id, autor_papel, texto) VALUES (?, 'sistema', ?)`, [chamadoId, texto], () => resolve()));
    await new Promise((resolve) => db.run(
        `UPDATE chamados SET updated_at = CURRENT_TIMESTAMP, nao_lido_empresa = 1, nao_lido_master = 1${novoStatus ? `, status = ?, closed_at = CASE WHEN ? IN ('resolvido', 'fechado') THEN CURRENT_TIMESTAMP ELSE NULL END` : ''} WHERE id = ?`,
        novoStatus ? [novoStatus, novoStatus, chamadoId] : [chamadoId], () => resolve()));
}

async function obterChamadoComAcesso(req, res, id) {
    const c = await dbGet(`SELECT * FROM chamados WHERE id = ?`, [id]);
    if (!c) { res.status(404).json({ error: 'Chamado não encontrado.' }); return null; }
    if (req.user.role === 'client_admin' && String(c.company_id) !== String(req.user.companyId)) { res.status(403).json({ error: 'Este chamado não pertence à sua empresa.' }); return null; }
    return c;
}

async function descreverReferenciaChamado(c) {
    if (c.ref_tipo === 'autoavaliacao' && c.ref_id) {
        const av = await dbGet(`SELECT id, referencia, status FROM dpo_self_assessments WHERE id = ?`, [c.ref_id]);
        if (av) return { tipo: 'autoavaliacao', id: av.id, rotulo: `Autoavaliação DPO — ${rotuloReferenciaDpo(av.referencia)}`, status: av.status };
    }
    return null;
}

app.get('/api/chamados', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const filtros = [], params = [];
        if (req.user.role === 'client_admin') { filtros.push('c.company_id = ?'); params.push(req.user.companyId); }
        else if (req.query.company_id) { filtros.push('c.company_id = ?'); params.push(req.query.company_id); }
        if (req.query.categoria && CATEGORIAS_CHAMADO[req.query.categoria]) { filtros.push('c.categoria = ?'); params.push(req.query.categoria); }
        const lista = await dbAll(`
            SELECT c.*, e.name as empresaNome, u.name as autorNome,
                   (SELECT COUNT(*) FROM chamado_mensagens m WHERE m.chamado_id = c.id AND m.autor_papel != 'sistema') as totalMensagens
            FROM chamados c
            LEFT JOIN companies e ON e.id = c.company_id
            LEFT JOIN users u ON u.id = c.created_by
            ${filtros.length ? 'WHERE ' + filtros.join(' AND ') : ''}
            ORDER BY CASE c.status WHEN 'aberto' THEN 0 WHEN 'em_atendimento' THEN 1 WHEN 'aguardando_empresa' THEN 2 WHEN 'resolvido' THEN 3 ELSE 4 END,
                     CASE c.prioridade WHEN 'urgente' THEN 0 WHEN 'alta' THEN 1 WHEN 'media' THEN 2 ELSE 3 END,
                     c.updated_at DESC`, params);
        const contagem = { aberto: 0, em_atendimento: 0, aguardando_empresa: 0, resolvido: 0, fechado: 0 };
        lista.forEach(c => { contagem[c.status] = (contagem[c.status] || 0) + 1; });
        const resolvidos = lista.filter(c => c.closed_at);
        const horas = (a, b) => (new Date(String(b).replace(' ', 'T') + 'Z') - new Date(String(a).replace(' ', 'T') + 'Z')) / 36e5;
        const comResposta = lista.filter(c => c.first_response_at);
        const avaliados = lista.filter(c => c.avaliacao);
        res.json({
            categorias: CATEGORIAS_CHAMADO, statusRotulos: STATUS_CHAMADO,
            indicadores: {
                ...contagem,
                naoLidos: lista.filter(c => req.user.role === 'admin' ? c.nao_lido_master : c.nao_lido_empresa).length,
                horasPrimeiraResposta: comResposta.length ? Math.round(comResposta.reduce((s, c) => s + horas(c.created_at, c.first_response_at), 0) / comResposta.length * 10) / 10 : null,
                horasResolucao: resolvidos.length ? Math.round(resolvidos.reduce((s, c) => s + horas(c.created_at, c.closed_at), 0) / resolvidos.length * 10) / 10 : null,
                satisfacao: avaliados.length ? Math.round(avaliados.reduce((s, c) => s + c.avaliacao, 0) / avaliados.length * 10) / 10 : null
            },
            lista: lista.map(c => ({ ...c, numero: numeroChamado(c.id), categoriaRotulo: CATEGORIAS_CHAMADO[c.categoria] || c.categoria, naoLido: !!(req.user.role === 'admin' ? c.nao_lido_master : c.nao_lido_empresa) }))
        });
    } catch (e) {
        console.error('Erro ao listar chamados:', e.message);
        res.status(500).json({ error: 'Erro ao carregar os chamados.' });
    }
});

app.post('/api/chamados/upload', requireRole('admin', 'client_admin'), (req, res) => {
    uploadMaterialDpo.single('file')(req, res, (err) => {
        if (err) return res.status(400).json({ error: err.code === 'LIMIT_FILE_SIZE' ? 'Arquivo muito grande (máximo 100MB).' : err.message });
        if (!req.file) return res.status(400).json({ error: 'Nenhum arquivo recebido.' });
        res.json({ url: '/uploads/' + req.file.filename, originalName: req.file.originalname });
    });
});

const anexoValidoChamado = url => !url || /^\/uploads\/[\w.\-]+$/.test(String(url));

app.post('/api/chamados', requireRole('admin', 'client_admin'), async (req, res) => {
    const { categoria, prioridade, ref_tipo, ref_id, anexo_url, anexo_nome } = req.body;
    const assunto = String(req.body.assunto || '').trim().slice(0, 200);
    const descricao = String(req.body.descricao || '').trim().slice(0, 5000);
    const companyId = req.user.role === 'client_admin' ? req.user.companyId : req.body.company_id;
    if (!companyId) return res.status(400).json({ error: 'Informe a empresa.' });
    if (!CATEGORIAS_CHAMADO[categoria]) return res.status(400).json({ error: 'Escolha a categoria do chamado.' });
    if (!assunto) return res.status(400).json({ error: 'Informe o assunto.' });
    if (!descricao) return res.status(400).json({ error: 'Descreva o que você precisa.' });
    if (!anexoValidoChamado(anexo_url)) return res.status(400).json({ error: 'Anexo inválido.' });
    try {
        let refTipo = null, refId = null;
        if (categoria === 'liberar_ferramenta') {
            const jaAberto = await dbGet(`SELECT id FROM chamados WHERE company_id = ? AND categoria = 'liberar_ferramenta' AND assunto = ? AND status NOT IN ('resolvido', 'fechado')`, [companyId, assunto]);
            if (jaAberto) return res.status(400).json({ error: `Você já pediu esta liberação — chamado ${numeroChamado(jaAberto.id)} em andamento. O Master vai avisar quando liberar.` });
        }
        if (categoria === 'ajuste_autoavaliacao') {
            const av = ref_id ? await dbGet(`SELECT * FROM dpo_self_assessments WHERE id = ? AND company_id = ?`, [ref_id, companyId]) : null;
            if (!av) return res.status(400).json({ error: 'Escolha qual autoavaliação precisa de ajuste.' });
            const jaAberto = await dbGet(`SELECT id FROM chamados WHERE ref_tipo = 'autoavaliacao' AND ref_id = ? AND status NOT IN ('resolvido', 'fechado')`, [av.id]);
            if (jaAberto) return res.status(400).json({ error: `Já existe o chamado ${numeroChamado(jaAberto.id)} aberto para esta autoavaliação.` });
            refTipo = 'autoavaliacao'; refId = av.id;
        }
        const id = await new Promise((resolve, reject) => db.run(
            `INSERT INTO chamados (company_id, created_by, categoria, assunto, prioridade, ref_tipo, ref_id) VALUES (?, ?, ?, ?, ?, ?, ?)`,
            [companyId, req.user.userId, categoria, assunto, PRIORIDADES_CHAMADO.includes(prioridade) ? prioridade : 'media', refTipo, refId],
            function (err) { err ? reject(err) : resolve(this.lastID); }
        ));
        await new Promise((resolve, reject) => db.run(
            `INSERT INTO chamado_mensagens (chamado_id, user_id, autor_papel, texto, anexo_url, anexo_nome) VALUES (?, ?, ?, ?, ?, ?)`,
            [id, req.user.userId, req.user.role === 'admin' ? 'master' : 'empresa', descricao, anexo_url || null, anexo_nome ? String(anexo_nome).slice(0, 200) : null],
            (err) => err ? reject(err) : resolve()
        ));
        const empresa = await dbGet(`SELECT name FROM companies WHERE id = ?`, [companyId]);
        if (req.user.role === 'client_admin') notificarMasters(`Novo chamado ${numeroChamado(id)}`, `${empresa ? empresa.name : 'Empresa'}: ${assunto}`, 'chamados');
        else notificarPorCompanyAdmins(companyId, `Novo chamado ${numeroChamado(id)}`, assunto, 'chamados');
        res.json({ message: `Chamado ${numeroChamado(id)} aberto!`, id });
    } catch (e) {
        console.error('Erro ao abrir chamado:', e.message);
        res.status(400).json({ error: 'Erro ao abrir o chamado.' });
    }
});

app.get('/api/chamados/:id', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const c = await obterChamadoComAcesso(req, res, req.params.id);
        if (!c) return;
        const mensagens = await dbAll(`SELECT m.*, u.name as autorNome FROM chamado_mensagens m LEFT JOIN users u ON u.id = m.user_id WHERE m.chamado_id = ? ORDER BY m.created_at ASC, m.id ASC`, [c.id]);
        const empresa = await dbGet(`SELECT name FROM companies WHERE id = ?`, [c.company_id]);
        const autor = c.created_by ? await dbGet(`SELECT name FROM users WHERE id = ?`, [c.created_by]) : null;
        db.run(`UPDATE chamados SET ${req.user.role === 'admin' ? 'nao_lido_master' : 'nao_lido_empresa'} = 0 WHERE id = ?`, [c.id], () => {});
        res.json({
            ...c, numero: numeroChamado(c.id), categoriaRotulo: CATEGORIAS_CHAMADO[c.categoria] || c.categoria, statusRotulos: STATUS_CHAMADO,
            empresaNome: empresa ? empresa.name : '', autorNome: autor ? autor.name : '', referencia: await descreverReferenciaChamado(c), mensagens
        });
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar o chamado.' }); }
});

app.post('/api/chamados/:id/mensagens', requireRole('admin', 'client_admin'), async (req, res) => {
    const texto = String(req.body.texto || '').trim().slice(0, 5000);
    const { anexo_url, anexo_nome } = req.body;
    if (!texto && !anexo_url) return res.status(400).json({ error: 'Escreva uma mensagem ou anexe um arquivo.' });
    if (!anexoValidoChamado(anexo_url)) return res.status(400).json({ error: 'Anexo inválido.' });
    try {
        const c = await obterChamadoComAcesso(req, res, req.params.id);
        if (!c) return;
        if (c.status === 'fechado' && req.user.role !== 'admin') return res.status(400).json({ error: 'Este chamado está fechado. Reabra para responder.' });
        const ehMaster = req.user.role === 'admin';
        await new Promise((resolve, reject) => db.run(
            `INSERT INTO chamado_mensagens (chamado_id, user_id, autor_papel, texto, anexo_url, anexo_nome) VALUES (?, ?, ?, ?, ?, ?)`,
            [c.id, req.user.userId, ehMaster ? 'master' : 'empresa', texto || null, anexo_url || null, anexo_nome ? String(anexo_nome).slice(0, 200) : null],
            (err) => err ? reject(err) : resolve()
        ));
        // Master respondendo um chamado novo -> "em atendimento"; empresa
        // respondendo um que aguardava ela (ou já resolvido) -> volta pro Master.
        let novoStatus = c.status;
        if (ehMaster && c.status === 'aberto') novoStatus = 'em_atendimento';
        if (!ehMaster && ['aguardando_empresa', 'resolvido'].includes(c.status)) novoStatus = 'em_atendimento';
        if (ehMaster && req.body.aguardarEmpresa) novoStatus = 'aguardando_empresa';
        await new Promise((resolve, reject) => db.run(
            `UPDATE chamados SET status = ?, updated_at = CURRENT_TIMESTAMP, ${ehMaster ? 'nao_lido_empresa = 1' : 'nao_lido_master = 1'},
                first_response_at = CASE WHEN ? = 1 AND first_response_at IS NULL THEN CURRENT_TIMESTAMP ELSE first_response_at END,
                closed_at = CASE WHEN ? IN ('resolvido', 'fechado') THEN closed_at ELSE NULL END
             WHERE id = ?`,
            [novoStatus, ehMaster ? 1 : 0, novoStatus, c.id], (err) => err ? reject(err) : resolve()
        ));
        if (ehMaster) notificarPorCompanyAdmins(c.company_id, `Resposta no chamado ${numeroChamado(c.id)}`, c.assunto, 'chamados');
        else notificarMasters(`Nova mensagem no chamado ${numeroChamado(c.id)}`, c.assunto, 'chamados');
        res.json({ message: 'Mensagem enviada!', status: novoStatus });
    } catch (e) { res.status(400).json({ error: 'Erro ao enviar a mensagem.' }); }
});

// Master: muda status/prioridade. Empresa: fecha (com avaliação) ou reabre.
app.put('/api/chamados/:id', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const c = await obterChamadoComAcesso(req, res, req.params.id);
        if (!c) return;
        const ehMaster = req.user.role === 'admin';
        const { status, prioridade } = req.body;
        const mudancas = [];
        if (prioridade !== undefined) {
            if (!ehMaster) return res.status(403).json({ error: 'Só o Master altera a prioridade.' });
            if (!PRIORIDADES_CHAMADO.includes(prioridade)) return res.status(400).json({ error: 'Prioridade inválida.' });
            if (prioridade !== c.prioridade) {
                await new Promise((resolve) => db.run(`UPDATE chamados SET prioridade = ? WHERE id = ?`, [prioridade, c.id], () => resolve()));
                mudancas.push(`Prioridade alterada para ${prioridade}.`);
            }
        }
        if (status !== undefined && status !== c.status) {
            if (!STATUS_CHAMADO[status]) return res.status(400).json({ error: 'Status inválido.' });
            if (!ehMaster && !['fechado', 'aberto'].includes(status)) return res.status(403).json({ error: 'A empresa só pode fechar ou reabrir o chamado.' });
            if (!ehMaster && status === 'aberto' && !['resolvido', 'fechado'].includes(c.status)) return res.status(400).json({ error: 'O chamado já está aberto.' });
            const nota = req.body.avaliacao ? Math.max(1, Math.min(5, Number(req.body.avaliacao))) : null;
            if (!ehMaster && status === 'fechado' && nota) {
                await new Promise((resolve) => db.run(`UPDATE chamados SET avaliacao = ?, avaliacao_comentario = ? WHERE id = ?`, [nota, String(req.body.comentario || '').slice(0, 1000) || null, c.id], () => resolve()));
            }
            const quem = ehMaster ? 'Master' : 'Empresa';
            await registrarMensagemSistemaChamado(c.id, `${quem} alterou o status para "${STATUS_CHAMADO[status]}".${nota ? ` Avaliação do atendimento: ${'★'.repeat(nota)}${'☆'.repeat(5 - nota)}` : ''}`, status);
            if (ehMaster) notificarPorCompanyAdmins(c.company_id, `Chamado ${numeroChamado(c.id)}: ${STATUS_CHAMADO[status]}`, c.assunto, 'chamados');
            else notificarMasters(`Chamado ${numeroChamado(c.id)}: ${STATUS_CHAMADO[status]}`, c.assunto, 'chamados');
        }
        if (mudancas.length) await registrarMensagemSistemaChamado(c.id, mudancas.join(' '));
        res.json({ message: 'Chamado atualizado!' });
    } catch (e) { res.status(400).json({ error: 'Erro ao atualizar o chamado.' }); }
});

// Master libera o ajuste pedido: reabre a autoavaliação ligada ao chamado.
app.post('/api/chamados/:id/liberar-ajuste', requireRole('admin'), async (req, res) => {
    try {
        const c = await obterChamadoComAcesso(req, res, req.params.id);
        if (!c) return;
        if (c.ref_tipo !== 'autoavaliacao' || !c.ref_id) return res.status(400).json({ error: 'Este chamado não está ligado a uma autoavaliação.' });
        const av = await dbGet(`SELECT * FROM dpo_self_assessments WHERE id = ?`, [c.ref_id]);
        if (!av) return res.status(404).json({ error: 'Autoavaliação não encontrada.' });
        if (av.status !== 'aprovada') return res.status(400).json({ error: 'A autoavaliação já está liberada para edição.' });
        await new Promise((resolve, reject) => db.run(`UPDATE dpo_self_assessments SET status = 'em_andamento', approved_by = NULL, approved_at = NULL WHERE id = ?`, [av.id], (err) => err ? reject(err) : resolve()));
        registrarEventoAutoavaliacaoDpo(av.id, 'reabrir', `Chamado ${numeroChamado(c.id)}`, req.user.userId);
        await registrarMensagemSistemaChamado(c.id, `🔓 O Master liberou a autoavaliação de ${rotuloReferenciaDpo(av.referencia)} para ajuste. Faça as alterações e aprove novamente.`, 'aguardando_empresa');
        db.run(`UPDATE chamados SET first_response_at = COALESCE(first_response_at, CURRENT_TIMESTAMP) WHERE id = ?`, [c.id], () => {});
        notificarPorCompanyAdmins(c.company_id, `Ajuste liberado — chamado ${numeroChamado(c.id)}`, `A autoavaliação de ${rotuloReferenciaDpo(av.referencia)} foi reaberta para ajuste.`, 'chamados');
        res.json({ message: 'Autoavaliação liberada para a empresa ajustar!' });
    } catch (e) { res.status(400).json({ error: 'Erro ao liberar o ajuste.' }); }
});

app.get('/api/chamados-contagem', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const r = req.user.role === 'admin'
            ? await dbGet(`SELECT COUNT(*) as naoLidos FROM chamados WHERE nao_lido_master = 1 AND status != 'fechado'`)
            : await dbGet(`SELECT COUNT(*) as naoLidos FROM chamados WHERE company_id = ? AND nao_lido_empresa = 1`, [req.user.companyId]);
        res.json(r);
    } catch (e) { res.json({ naoLidos: 0 }); }
});

// ---------- DPO — "Ferramentas Impulsionar" (biblioteca do Master por pilar) ----------
const TIPOS_FERRAMENTA_DPO = ['planilha', 'modelo_padrao', 'treinamento', 'checklist', 'apresentacao', 'link', 'outro'];

function validarFerramentaDpo(body, parcial) {
    const erros = [];
    const dados = {};
    if (!parcial || body.pillarKey !== undefined) {
        if (!DPO_PILARES_ORDEM.includes(body.pillarKey)) erros.push('Pilar inválido.');
        dados.pillar_key = body.pillarKey;
    }
    if (!parcial || body.questionNumero !== undefined) {
        const q = body.questionNumero ? String(body.questionNumero) : null;
        if (q && !perguntaDoPilarDpo(body.pillarKey || dados.pillar_key, q)) erros.push('Pergunta do pilar inválida.');
        dados.question_numero = q;
    }
    if (!parcial || body.tipo !== undefined) {
        if (!TIPOS_FERRAMENTA_DPO.includes(body.tipo)) erros.push('Tipo inválido.');
        dados.tipo = body.tipo;
    }
    if (!parcial || body.titulo !== undefined) {
        const t = String(body.titulo || '').trim().slice(0, 200);
        if (!t) erros.push('Informe o título.');
        dados.titulo = t;
    }
    if (body.descricao !== undefined) dados.descricao = String(body.descricao || '').trim().slice(0, 2000) || null;
    if (!parcial || body.url !== undefined) {
        const u = String(body.url || '').trim();
        if (!/^https?:\/\/\S+$/i.test(u) && !/^\/uploads\/[\w.\-]+$/.test(u)) erros.push('Envie um arquivo ou informe um link válido (http/https).');
        dados.url = u;
        dados.original_name = body.originalName ? String(body.originalName).slice(0, 200) : null;
    }
    if (body.ativo !== undefined) dados.ativo = body.ativo ? 1 : 0;
    return { erros, dados };
}

app.get('/api/dpo/ferramentas/:pillarKey', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const pilarKey = req.params.pillarKey;
        if (req.user.role === 'client_admin') {
            const ok = await resolverEmpresaPastaDpo(req, res, pilarKey, null, 'ferramentas');
            if (!ok) return;
        } else if (!DPO_PILARES_ORDEM.includes(pilarKey)) return res.status(400).json({ error: 'Pilar inválido.' });
        const lista = await dbAll(`SELECT * FROM dpo_ferramentas WHERE pillar_key = ? ${req.user.role === 'admin' ? '' : 'AND ativo = 1'} ORDER BY created_at DESC`, [pilarKey]);
        const ordem = ordemDasPerguntasDoPilarDpo(pilarKey);
        lista.sort((a, b) => {
            const oa = a.question_numero ? (ordem[a.question_numero] ?? 9999) : -1;
            const ob = b.question_numero ? (ordem[b.question_numero] ?? 9999) : -1;
            return oa - ob || String(b.created_at).localeCompare(String(a.created_at));
        });
        res.json(lista.map(f => ({ ...f, perguntaTexto: f.question_numero ? textoDaPerguntaDpo(pilarKey, f.question_numero) : null })));
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar as ferramentas.' }); }
});

app.post('/api/admin/dpo/ferramentas', requireRole('admin'), async (req, res) => {
    const { erros, dados } = validarFerramentaDpo(req.body, false);
    if (erros.length) return res.status(400).json({ error: erros[0] });
    try {
        await new Promise((resolve, reject) => db.run(
            `INSERT INTO dpo_ferramentas (pillar_key, question_numero, tipo, titulo, descricao, url, original_name, created_by) VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
            [dados.pillar_key, dados.question_numero, dados.tipo, dados.titulo, dados.descricao || null, dados.url, dados.original_name, req.user.userId],
            (err) => err ? reject(err) : resolve()
        ));
        res.json({ message: 'Ferramenta publicada!' });
    } catch (e) { res.status(400).json({ error: 'Erro ao publicar a ferramenta.' }); }
});

app.put('/api/admin/dpo/ferramentas/:id', requireRole('admin'), async (req, res) => {
    try {
        const atual = await dbGet(`SELECT * FROM dpo_ferramentas WHERE id = ?`, [req.params.id]);
        if (!atual) return res.status(404).json({ error: 'Ferramenta não encontrada.' });
        const { erros, dados } = validarFerramentaDpo({ pillarKey: atual.pillar_key, ...req.body }, true);
        if (erros.length) return res.status(400).json({ error: erros[0] });
        delete dados.pillar_key;
        const campos = Object.keys(dados);
        if (!campos.length) return res.json({ message: 'Nada para alterar.' });
        await new Promise((resolve, reject) => db.run(
            `UPDATE dpo_ferramentas SET ${campos.map(c => c + ' = ?').join(', ')}, updated_at = CURRENT_TIMESTAMP WHERE id = ?`,
            [...campos.map(c => dados[c]), atual.id], (err) => err ? reject(err) : resolve()
        ));
        res.json({ message: 'Ferramenta atualizada!' });
    } catch (e) { res.status(400).json({ error: 'Erro ao atualizar a ferramenta.' }); }
});

app.delete('/api/admin/dpo/ferramentas/:id', requireRole('admin'), async (req, res) => {
    try {
        await new Promise((resolve, reject) => db.run(`DELETE FROM dpo_ferramentas WHERE id = ?`, [req.params.id], (err) => err ? reject(err) : resolve()));
        res.json({ message: 'Ferramenta removida!' });
    } catch (e) { res.status(400).json({ error: 'Erro ao remover a ferramenta.' }); }
});

// Conta quantas vezes cada ferramenta foi aberta/baixada (para o Master ver o uso).
app.post('/api/dpo/ferramentas/:id/acesso', requireRole('admin', 'client_admin'), (req, res) => {
    if (req.user.role === 'admin') return res.json({ ok: true });
    db.run(`UPDATE dpo_ferramentas SET acessos = acessos + 1 WHERE id = ?`, [req.params.id], () => res.json({ ok: true }));
});

// ======================================================================
// DPO — MATERIAL DPO (biblioteca do Master + compartilhamento por empresa)
// ======================================================================
function validarMaterialImpulsionarDpo(body, parcial) {
    const erros = []; const dados = {};
    if (!parcial || body.pillarKey !== undefined) {
        if (!DPO_PILARES_ORDEM.includes(body.pillarKey)) erros.push('Escolha o pilar.');
        else dados.pillar_key = body.pillarKey;
    }
    if (!parcial || body.questionNumero !== undefined) {
        const pk = body.pillarKey;
        const q = String(body.questionNumero || '').trim();
        if (!q || (pk && !textoDaPerguntaDpo(pk, q))) erros.push('Escolha a pergunta do pilar.');
        else dados.question_numero = q;
    }
    if (!parcial || body.titulo !== undefined) {
        const t = String(body.titulo || '').trim().slice(0, 200);
        if (!t) erros.push('Informe o título do material.'); else dados.titulo = t;
    }
    if (body.descricao !== undefined) dados.descricao = String(body.descricao || '').trim().slice(0, 2000) || null;
    if (!parcial || body.url !== undefined) {
        const u = String(body.url || '').trim();
        if (!/^https?:\/\/\S+$/i.test(u) && !/^\/uploads\/[\w.\-]+$/.test(u)) erros.push('Envie o arquivo ou informe um link válido.');
        else { dados.url = u; dados.original_name = body.originalName ? String(body.originalName).slice(0, 200) : null; }
    }
    return { erros, dados };
}

function estruturaPilaresDpo() {
    return DPO_PILARES_ORDEM.filter(k => DPO_AMBEV_DATA[k]).map(k => ({
        key: k, label: DPO_AMBEV_DATA[k].label,
        grupos: (DPO_AMBEV_DATA[k].grupos || []).map(g => ({ numero: g.numero, titulo: g.titulo, perguntas: g.perguntas.map(q => ({ numero: q.numero, questao: q.questao })) }))
    }));
}

app.get('/api/admin/material-dpo', requireRole('admin'), async (req, res) => {
    try {
        const materiais = await dbAll(`SELECT * FROM dpo_material_impulsionar ORDER BY created_at DESC, id DESC`);
        const shares = await dbAll(`SELECT s.*, c.name as companyName FROM dpo_material_impulsionar_share s LEFT JOIN companies c ON c.id = s.company_id`);
        const porMat = {}; shares.forEach(s => (porMat[s.material_id] = porMat[s.material_id] || []).push({ companyId: s.company_id, companyName: s.companyName, sharedAt: s.shared_at, acessos: s.acessos || 0 }));
        const empresas = await dbAll(`SELECT id, name, enabled_modules FROM companies ORDER BY name COLLATE NOCASE`);
        res.json({
            materiais: materiais.map(m => ({ ...m, perguntaTexto: textoDaPerguntaDpo(m.pillar_key, m.question_numero), compartilhado: porMat[m.id] || [] })),
            empresas: empresas.map(e => {
                let dpo = true;
                try { const mods = e.enabled_modules ? JSON.parse(e.enabled_modules) : null; if (Array.isArray(mods)) dpo = mods.includes('dpoAmbev'); } catch (x) { /* sem restrição */ }
                return { id: e.id, name: e.name, dpo };
            }),
            pilares: estruturaPilaresDpo()
        });
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar o Material DPO.' }); }
});

app.post('/api/admin/material-dpo', requireRole('admin'), async (req, res) => {
    const { erros, dados } = validarMaterialImpulsionarDpo(req.body, false);
    if (erros.length) return res.status(400).json({ error: erros[0] });
    try {
        const id = await new Promise((resolve, reject) => db.run(
            `INSERT INTO dpo_material_impulsionar (pillar_key, question_numero, titulo, descricao, url, original_name, created_by) VALUES (?, ?, ?, ?, ?, ?, ?)`,
            [dados.pillar_key, dados.question_numero, dados.titulo, dados.descricao || null, dados.url, dados.original_name, req.user.userId],
            function (err) { err ? reject(err) : resolve(this.lastID); }));
        res.json({ message: 'Material salvo!', id });
    } catch (e) { res.status(400).json({ error: 'Erro ao salvar o material.' }); }
});

app.put('/api/admin/material-dpo/:id', requireRole('admin'), async (req, res) => {
    try {
        const atual = await dbGet(`SELECT * FROM dpo_material_impulsionar WHERE id = ?`, [req.params.id]);
        if (!atual) return res.status(404).json({ error: 'Material não encontrado.' });
        const corpo = { ...req.body };
        if (corpo.questionNumero !== undefined && corpo.pillarKey === undefined) corpo.pillarKey = atual.pillar_key;
        const { erros, dados } = validarMaterialImpulsionarDpo(corpo, true);
        if (erros.length) return res.status(400).json({ error: erros[0] });
        const campos = Object.keys(dados);
        if (!campos.length) return res.json({ message: 'Nada para alterar.' });
        await new Promise((resolve, reject) => db.run(
            `UPDATE dpo_material_impulsionar SET ${campos.map(c => c + ' = ?').join(', ')}, updated_at = CURRENT_TIMESTAMP WHERE id = ?`,
            [...campos.map(c => dados[c]), atual.id], (err) => err ? reject(err) : resolve()));
        res.json({ message: 'Material atualizado!' });
    } catch (e) { res.status(400).json({ error: 'Erro ao atualizar o material.' }); }
});

app.delete('/api/admin/material-dpo/:id', requireRole('admin'), async (req, res) => {
    try {
        await new Promise((resolve, reject) => db.run(`DELETE FROM dpo_material_impulsionar_share WHERE material_id = ?`, [req.params.id], (err) => err ? reject(err) : resolve()));
        await new Promise((resolve, reject) => db.run(`DELETE FROM dpo_material_impulsionar WHERE id = ?`, [req.params.id], (err) => err ? reject(err) : resolve()));
        res.json({ message: 'Material removido!' });
    } catch (e) { res.status(400).json({ error: 'Erro ao remover o material.' }); }
});

// Disponibiliza / retira o material para uma ou mais empresas.
// body: { companyIds: [..], compartilhar: true|false }  (ou { todas: true, compartilhar })
app.put('/api/admin/material-dpo/:id/compartilhar', requireRole('admin'), async (req, res) => {
    try {
        const mat = await dbGet(`SELECT * FROM dpo_material_impulsionar WHERE id = ?`, [req.params.id]);
        if (!mat) return res.status(404).json({ error: 'Material não encontrado.' });
        const compartilhar = req.body.compartilhar !== false;
        let ids = Array.isArray(req.body.companyIds) ? req.body.companyIds.map(Number).filter(n => n > 0) : [];
        if (req.body.todas) ids = (await dbAll(`SELECT id FROM companies`)).map(c => c.id);
        if (!ids.length) return res.status(400).json({ error: 'Escolha ao menos uma empresa.' });
        let novos = 0;
        for (const cid of ids) {
            if (compartilhar) {
                const r = await new Promise((resolve, reject) => db.run(
                    `INSERT OR IGNORE INTO dpo_material_impulsionar_share (material_id, company_id, shared_by) VALUES (?, ?, ?)`,
                    [mat.id, cid, req.user.userId], function (err) { err ? reject(err) : resolve(this.changes); }));
                if (r) {
                    novos++;
                    const pilar = DPO_AMBEV_DATA[mat.pillar_key];
                    notificarPorCompanyAdmins(cid, '📂 Novo Material Impulsionar',
                        `A Impulsionar disponibilizou "${mat.titulo}" na pergunta ${mat.question_numero} de ${pilar ? pilar.label : mat.pillar_key}.`, 'dpoHome');
                }
            } else {
                await new Promise((resolve, reject) => db.run(`DELETE FROM dpo_material_impulsionar_share WHERE material_id = ? AND company_id = ?`, [mat.id, cid], (err) => err ? reject(err) : resolve()));
            }
        }
        res.json({ message: compartilhar ? (novos ? `Disponibilizado para ${novos} empresa(s)!` : 'Já estava disponível.') : 'Compartilhamento removido.' });
    } catch (e) { res.status(400).json({ error: 'Erro ao alterar o compartilhamento.' }); }
});

// Materiais disponibilizados para a empresa (só os compartilhados com ela).
app.get('/api/dpo/material-impulsionar', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const companyId = req.user.role === 'admin' ? Number(req.query.company_id) : req.user.companyId;
        if (!companyId) return res.json([]);
        const params = [companyId];
        let filtro = '';
        if (req.query.pillar) { filtro = ' AND m.pillar_key = ?'; params.push(String(req.query.pillar)); }
        const lista = await dbAll(`SELECT m.id, m.pillar_key, m.question_numero, m.titulo, m.descricao, m.url, m.original_name, s.shared_at
            FROM dpo_material_impulsionar m JOIN dpo_material_impulsionar_share s ON s.material_id = m.id
            WHERE s.company_id = ?${filtro} ORDER BY s.shared_at DESC, m.id DESC`, params);
        res.json(lista.map(m => ({ ...m, perguntaTexto: textoDaPerguntaDpo(m.pillar_key, m.question_numero) })));
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar o Material Impulsionar.' }); }
});

app.post('/api/dpo/material-impulsionar/:id/acesso', requireRole('admin', 'client_admin'), (req, res) => {
    if (req.user.role === 'admin') return res.json({ ok: true });
    db.run(`UPDATE dpo_material_impulsionar_share SET acessos = acessos + 1 WHERE material_id = ? AND company_id = ?`, [req.params.id, req.user.companyId], () => res.json({ ok: true }));
});

// ======================================================================
// DPO — FERRAMENTAS DIGITAIS por pergunta do checklist
//   swot         -> Gestão 1.3 (SWOT por área: Armazém, Distribuição, Frota, Gente)
//   Simulador de Dimensionamento -> Planejamento 1.1, em pastas:
//     sim_armazem, sim_entrega, sim_puxada (simuladores) + orcamento (RACI, KPIs, PDCA)
// Cada pasta é salva por ANO (um registro JSON por empresa/pasta/ano).
// ======================================================================
const FERRAMENTAS_DIGITAIS_DPO = {
    swot: { pilar: 'gestao', pergunta: '1.3', titulo: 'Análise SWOT digital' },
    sim_armazem: { pilar: 'planejamento', pergunta: '1.1', titulo: 'Simulador de Dimensionamento — Armazém' },
    sim_entrega: { pilar: 'planejamento', pergunta: '1.1', titulo: 'Simulador de Dimensionamento — Entrega' },
    sim_puxada: { pilar: 'planejamento', pergunta: '1.1', titulo: 'Simulador de Dimensionamento — Puxada' },
    orcamento: { pilar: 'planejamento', pergunta: '1.1', titulo: 'Simulador de Dimensionamento — Orçamento, RACI e KPIs' }
};
const CHAVES_DIMENSIONAMENTO_DPO = ['sim_armazem', 'sim_entrega', 'sim_puxada', 'orcamento'];
const AREAS_SWOT_PADRAO_DPO = ['Armazém', 'Distribuição', 'Frota', 'Gente'];

// Escalas da planilha "Modelo SWOT 2026": pontuação = critério1 × critério2 × critério3 (1 a 125).
const ESCALAS_SWOT_DPO = {
    importancia: { 'Sem importância': 1, 'Pouco importante': 2, 'Importante': 3, 'Muito importante': 4, 'Totalmente importante': 5 },
    intensidadePos: { 'Muito fraca': 1, 'Fraca': 2, 'Média': 3, 'Forte': 4, 'Muito forte': 5 },
    intensidadeNeg: { 'Muito fraca': 5, 'Fraca': 4, 'Média': 3, 'Forte': 2, 'Muito forte': 1 },
    urgencia: { 'Nada urgente': 1, 'Pouco urgente': 2, 'Urgente': 3, 'Muito urgente': 4, 'Pra ontem': 5 },
    tendenciaPos: { 'Piora muito': 1, 'Piora': 2, 'Mantém': 3, 'Melhora': 4, 'Melhora muito': 5 },
    tendenciaNeg: { 'Piora muito': 5, 'Piora': 4, 'Mantém': 3, 'Melhora': 2, 'Melhora muito': 1 }
};
const QUADRANTES_SWOT_DPO = {
    forcas: { rotulo: 'Força', c2: 'intensidadePos', c3: 'tendenciaPos' },
    fraquezas: { rotulo: 'Fraqueza', c2: 'intensidadeNeg', c3: 'tendenciaNeg' },
    oportunidades: { rotulo: 'Oportunidade', c2: 'urgencia', c3: 'tendenciaPos' },
    ameacas: { rotulo: 'Ameaça', c2: 'urgencia', c3: 'tendenciaNeg' }
};
function pontuacaoItemSwotDpo(quadrante, item) {
    const q = QUADRANTES_SWOT_DPO[quadrante];
    const a = ESCALAS_SWOT_DPO.importancia[item.c1], b = ESCALAS_SWOT_DPO[q.c2][item.c2], c = ESCALAS_SWOT_DPO[q.c3][item.c3];
    return a && b && c ? a * b * c : null;
}
const numDpo = v => (v === null || v === undefined || v === '' || isNaN(Number(v))) ? null : Number(v);
const temTexto = v => !!String(v || '').trim();

function areasSwotDpo(d) { return (Array.isArray(d.areas) && d.areas.filter(temTexto).length) ? d.areas.filter(temTexto) : AREAS_SWOT_PADRAO_DPO.slice(); }
// Itens da SWOT por área: { area: { forcas: [...], ... } } (versões antigas guardavam em d.itens com campo "area").
function itensSwotPorAreaDpo(d) {
    const porArea = {};
    const add = (area, q, item) => { porArea[area] = porArea[area] || {}; (porArea[area][q] = porArea[area][q] || []).push({ ...item, area }); };
    Object.entries(d.itensPorArea || {}).forEach(([area, qs]) => Object.keys(QUADRANTES_SWOT_DPO).forEach(q => (qs && qs[q] || []).forEach(i => add(area, q, i))));
    Object.keys(QUADRANTES_SWOT_DPO).forEach(q => (d.itens && d.itens[q] || []).forEach(i => add(i.area || 'Sem área', q, i)));
    return porArea;
}

function validarSwotDpo(d, ano) {
    d = d || {};
    const areas = areasSwotDpo(d);
    const porArea = itensSwotPorAreaDpo(d);
    const pontuados = {};
    Object.keys(QUADRANTES_SWOT_DPO).forEach(q => { pontuados[q] = Object.values(porArea).flatMap(a => (a[q] || [])).filter(i => temTexto(i.texto) && pontuacaoItemSwotDpo(q, i)); });
    const quadrantesOk = Object.keys(QUADRANTES_SWOT_DPO).filter(k => pontuados[k].length >= 1);
    const objetivos = (d.objetivos || []).filter(o => temTexto(o.texto));
    const planos = (d.planos || []).filter(p => temTexto(p.oque));
    const desdobrados = [...objetivos, ...planos].filter(x => temTexto(x.desdobramento));
    const prazo = `${ano}-03-31`;

    const v1faltas = [];
    if (quadrantesOk.length < 4) v1faltas.push(`Pontue pelo menos 1 item em cada quadrante (faltam: ${Object.keys(QUADRANTES_SWOT_DPO).filter(k => !quadrantesOk.includes(k)).map(k => QUADRANTES_SWOT_DPO[k].rotulo).join(', ')}).`);
    if (!objetivos.length) v1faltas.push('Cadastre os objetivos estratégicos definidos a partir da SWOT.');
    if (objetivos.length && objetivos.some(o => !numDpo(o.prioridade))) v1faltas.push('Defina a prioridade de todos os objetivos estratégicos.');
    const avisos1 = [];
    if (!d.dataPriorizacao) avisos1.push('Informe a data da priorização (ciclo encerra até o fim de março).');
    else if (d.dataPriorizacao > prazo) avisos1.push(`Priorização em ${d.dataPriorizacao.split('-').reverse().join('/')} — depois do prazo do ciclo (31/03/${ano}).`);

    const v2faltas = [];
    if (!temTexto(d.sonho)) v2faltas.push('Escreva o Sonho da unidade.');
    if (!objetivos.length) v2faltas.push('Sem objetivos estratégicos cadastrados.');
    if (objetivos.some(o => !temTexto(o.obstaculo))) v2faltas.push('Em cada objetivo, informe qual obstáculo ao Sonho ele ataca.');
    if (objetivos.some(o => !temTexto(o.itemRelacionado))) v2faltas.push('Relacione cada objetivo a um item da SWOT.');

    const v3faltas = [], avisos3 = [];
    areas.forEach(area => {
        const faltando = Object.keys(QUADRANTES_SWOT_DPO).filter(q => !((porArea[area] || {})[q] || []).some(i => temTexto(i.texto) && pontuacaoItemSwotDpo(q, i)));
        if (faltando.length === 4) v3faltas.push(`${area}: SWOT da área ainda não preenchida.`);
        else if (faltando.length) v3faltas.push(`${area}: falta pontuar ${faltando.map(q => QUADRANTES_SWOT_DPO[q].rotulo.toLowerCase()).join(', ')}.`);
        else if (Object.keys(QUADRANTES_SWOT_DPO).some(q => ((porArea[area] || {})[q] || []).length < 3)) avisos3.push(`${area}: poucos itens em algum quadrante — o ideal é chegar no Top 5.`);
    });
    if (!d.vinculoDNMP || !temTexto(d.vinculoDNMPTexto)) v3faltas.push('Marque e descreva o vínculo da SWOT com a Descrição de Negócio e o Mapeamento de Processo.');

    const itensV = [
        { numero: 'V.1', texto: 'Definição e priorização dos objetivos estratégicos usando a análise SWOT.', ok: !v1faltas.length, faltas: v1faltas, avisos: avisos1 },
        { numero: 'V.2', texto: 'Objetivos estratégicos relacionados e abordando os principais obstáculos para alcançar o Sonho.', ok: !v2faltas.length, faltas: v2faltas, avisos: [] },
        { numero: 'V.3', texto: `SWOT executada por áreas (${areas.join(', ')}), com link claro com a Descrição de Negócio e o Mapeamento de Processo.`, ok: !v3faltas.length, faltas: v3faltas, avisos: avisos3 }
    ];
    const complemento = { texto: 'How to check 3: os resultados da SWOT definem CAPEX, cascateamento de metas e/ou PDCA.', ok: desdobrados.length > 0, faltas: desdobrados.length ? [] : ['Indique o desdobramento (PDCA, Projeto, CAPEX ou Cascateamento de metas) nos objetivos ou planos de ação.'] };
    const notaSugerida = !itensV[0].ok ? '0' : (itensV[1].ok && itensV[2].ok ? '3' : '1');
    return { itens: itensV, complementos: [complemento], notaSugerida, regra: '3 = todas atendidas · 1 = V.1 atendida mas V.2 ou V.3 não · 0 = V.1 não atendida' };
}

const MESES_CURTOS_DPO = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
function mesesPreenchidosDpo(lista, campos) {
    return (lista || []).filter(m => m && campos.every(c => numDpo(m[c]) !== null && numDpo(m[c]) > 0)).length;
}

// Validação da pergunta Planejamento 1.1 lendo as 4 pastas do mesmo ano.
function validarDimensionamentoDpo(t, arquivosOrc) {
    const ent = t.sim_entrega || {}, arm = t.sim_armazem || {}, pux = t.sim_puxada || {}, orc = t.orcamento || {};
    const raci = (orc.raci || []).filter(r => temTexto(r.pacote));
    const kpis = (orc.kpis || []).filter(k => temTexto(k.nome));
    const pdca = (orc.pdca || []).filter(p => temTexto(p.oque));

    const v1 = [];
    if (raci.length < 3) v1.push('Cadastre os pacotes orçamentários na matriz RACI (pelo menos 3).');
    if (raci.some(r => !temTexto(r.r) || !temTexto(r.a))) v1.push('Todo pacote precisa de um Responsável (R) e um Aprovador (A).');
    const responsaveis = new Set(raci.map(r => String(r.r || '').trim().toLowerCase()).filter(Boolean));
    if (raci.length && responsaveis.size < 2) v1.push('As responsabilidades estão concentradas em uma pessoa — distribua os pacotes (não só a gerência).');
    if (raci.length && !raci.some(r => temTexto(r.kpi))) v1.push('Conecte os pacotes ao DPO: informe o KPI/resultado de cada pacote.');

    const v2 = [];
    if (!temTexto(orc.processo) && !(arquivosOrc || []).some(a => a.tipo === 'processo')) v2.push('Descreva (ou anexe) o processo orçamentário formalizado.');
    const mktp = [ent, arm, pux].reduce((s, b) => s + (b.plan || []).reduce((x, m) => x + (numDpo(m && m.volume_mktp_hl) || 0), 0), 0);
    if (!mktp) v2.push('Inclua o volume de Marketplace no orçamento (em algum dos simuladores).');
    if (!(ent.plan || []).some(m => m && numDpo(m.ff_ativa))) v2.push('Planeje a frota da entrega (headcount de motoristas e ajudantes sai dela).');
    if (!(arm.plan || []).some(m => m && numDpo(m.volume_hl))) v2.push('Planeje o volume do armazém (QLP e empilhadeiras saem dele).');

    const v3 = [];
    const mEnt = mesesPreenchidosDpo(ent.plan, ['volume_hl', 'ff_ativa']), mArm = mesesPreenchidosDpo(arm.plan, ['volume_hl', 'dias_trab']), mPux = mesesPreenchidosDpo(pux.plan, ['volume_hl', 'carretas_ativas']);
    if (mEnt < 12) v3.push(`Simulador da Entrega: volume e frota no orçamento em ${mEnt}/12 meses.`);
    if (mArm < 12) v3.push(`Simulador do Armazém: volume e dias no orçamento em ${mArm}/12 meses.`);
    if (mPux < 12) v3.push(`Simulador da Puxada: volume e carretas no orçamento em ${mPux}/12 meses.`);

    const v4 = [];
    if (!orc.dataNegociacao) v4.push('Informe a data da negociação final do orçamento.');
    if (!orc.versaoFinal) v4.push('Confirme que o orçamento é a versão final, feita após a negociação.');
    if (!kpis.length || kpis.some(k => !temTexto(k.meta))) v4.push('Cadastre os KPIs de sustentabilidade com meta.');
    if (kpis.length && !kpis.some(k => temTexto(k.acao)) && !pdca.length) v4.push('Mostre a conexão orçamento × KPI × ações (ação ligada ao KPI ou PDCA).');

    const v5 = [];
    const reais = { Entrega: mesesPreenchidosDpo(ent.real, ['volume_hl']), Armazém: mesesPreenchidosDpo(arm.real, ['volume_hl']), Puxada: mesesPreenchidosDpo(pux.real, ['volume_hl']) };
    const semReal = Object.entries(reais).filter(([, n]) => !n).map(([k]) => k);
    if (semReal.length) v5.push(`Preencha o Realizado mensal (rotina de custos) em: ${semReal.join(', ')}.`);
    const areas = new Set(raci.map(r => String(r.area || '').trim().toLowerCase()).filter(Boolean));
    if (areas.size < 2) v5.push('A RACI precisa envolver pelo menos 2 áreas.');
    const donosPdca = new Set(pdca.map(p => String(p.responsavel || '').trim().toLowerCase()).filter(Boolean));
    if (!pdca.length) v5.push('Registre o PDCA de custos com as ações das áreas.');
    else if (donosPdca.size < 2) v5.push('As ações do PDCA estão com um único responsável — envolva as áreas.');

    const itensV = [
        { numero: 'V.1', texto: 'Responsáveis pelos pacotes orçamentários definidos em matriz RACI, não concentrados na gerência e conectados ao DPO.', ok: !v1.length, faltas: v1, avisos: [] },
        { numero: 'V.2', texto: 'Processo orçamentário estruturado e formalizado, com foco no Marketplace, headcount e equipamentos.', ok: !v2.length, faltas: v2, avisos: [] },
        { numero: 'V.3', texto: 'Simulador com projeção de volume, QLP e frota para armazém, puxada e entrega.', ok: !v3.length, faltas: v3, avisos: [] },
        { numero: 'V.4', texto: 'Orçamento feito após a negociação final e alinhado aos KPIs de sustentabilidade, com ações.', ok: !v4.length, faltas: v4, avisos: [] },
        { numero: 'V.5', texto: 'As áreas participam da construção e manutenção das rotinas de gestão de custos.', ok: !v5.length, faltas: v5, avisos: [] }
    ];
    const notaSugerida = (!itensV[0].ok || !itensV[1].ok || !itensV[2].ok) ? '0' : (itensV[3].ok && itensV[4].ok ? '3' : '1');
    return { itens: itensV, complementos: [], notaSugerida, regra: '3 = todas atendidas · 1 = V.1, V.2 e V.3 atendidas mas V.4 ou V.5 não · 0 = V.1, V.2 ou V.3 não atendidas' };
}

async function resolverFerramentaDigitalDpo(req, res, chave, companyIdInformado) {
    if (String(chave).startsWith('acomp:')) {
        const [, pilar, numero] = String(chave).split(':');
        const t = montarTemplateAcompDpo(pilar, numero);
        if (!t) { res.status(404).json({ error: 'Pergunta não encontrada.' }); return null; }
        const companyId = await resolverEmpresaPastaDpo(req, res, pilar, companyIdInformado, 'checklist');
        if (!companyId) return null;
        if (req.user.role === 'client_admin' && !(await acompLiberadoDpo(companyId, `${pilar}:${numero}`))) {
            res.status(403).json({ error: 'Este Acompanhamento Impulsionar ainda não foi liberado para sua empresa. Fale com o Master.' });
            return null;
        }
        return { pilar, pergunta: numero, titulo: `Acompanhamento Impulsionar — ${t.pilarLabel} ${numero} ${t.questao}`, companyId, template: t };
    }
    if (chave === 'cinco_s') {
        const companyId = await resolverEmpresaPastaDpo(req, res, 'gestao', companyIdInformado, null);
        if (!companyId) return null;
        if (req.user.role === 'client_admin' && !(await acompLiberadoDpo(companyId, 'gestao:3.1'))) { res.status(403).json({ error: 'Esta ferramenta Impulsionar ainda não foi liberada para sua empresa. Abra um chamado para o Master autorizar.' }); return null; }
        return { titulo: 'Gerenciador 5S', pilarLabel: 'Gestão Revenda', pergunta: '3.1', perguntaTexto: '5S', companyId, modelo5s: CINCO_S_MODELO_DPO };
    }
    if (chave === 'gop') {
        const companyId = req.user.role === 'client_admin' ? req.user.companyId : companyIdInformado;
        if (!companyId) { res.status(400).json({ error: 'Informe a empresa (company_id).' }); return null; }
        if (req.user.role === 'client_admin' && !(await pilaresAtivosDaEmpresa(companyId)).length) { res.status(403).json({ error: 'Sua empresa ainda não tem o DPO contratado.' }); return null; }
        if (req.user.role === 'client_admin' && !(await acompLiberadoDpo(companyId, 'gestao:4.6'))) { res.status(403).json({ error: 'Esta ferramenta Impulsionar ainda não foi liberada para sua empresa. Abra um chamado para o Master autorizar.' }); return null; }
        return { titulo: 'Gerenciador GOP — Revendas', pilarLabel: 'GOP', pergunta: '', perguntaTexto: '', companyId, modeloGop: GOP_MODELO_DPO };
    }
    const f = FERRAMENTAS_DIGITAIS_DPO[chave];
    if (!f) { res.status(404).json({ error: 'Ferramenta não encontrada.' }); return null; }
    const companyId = await resolverEmpresaPastaDpo(req, res, f.pilar, companyIdInformado, TOOLS_EXCLUSIVAS_DPO[chave] ? null : 'checklist');
    if (!companyId) return null;
    if (TOOLS_EXCLUSIVAS_DPO[chave] && req.user.role === 'client_admin' && !(await acompLiberadoDpo(companyId, TOOLS_EXCLUSIVAS_DPO[chave]))) {
        res.status(403).json({ error: 'Esta ferramenta Impulsionar ainda não foi liberada para sua empresa. Abra um chamado para o Master autorizar.' });
        return null;
    }
    if (!TOOLS_EXCLUSIVAS_DPO[chave] && req.user.role === 'client_admin' && !(await acompLiberadoDpo(companyId, `${f.pilar}:${f.pergunta}`))) {
        res.status(403).json({ error: 'Esta ferramenta Impulsionar ainda não foi liberada para sua empresa. Abra um chamado para o Master autorizar.' });
        return null;
    }
    return { ...f, companyId };
}

const anoValidoDpo = v => { const n = Number(v); return n >= 2020 && n <= 2100 ? n : new Date().getFullYear(); };

async function carregarFerramentaDigitalDpo(companyId, chave, ano) {
    const reg = await dbGet(`SELECT fd.*, u.name as autorNome FROM dpo_ferramentas_digitais fd LEFT JOIN users u ON u.id = fd.updated_by WHERE fd.company_id = ? AND fd.chave = ? AND fd.ano = ?`, [companyId, chave, ano]);
    const arquivos = await dbAll(`SELECT a.*, u.name as autorNome FROM dpo_ferramentas_digitais_arquivos a LEFT JOIN users u ON u.id = a.created_by WHERE a.company_id = ? AND a.chave = ? AND a.ano = ? ORDER BY a.created_at DESC, a.id DESC`, [companyId, chave, ano]);
    return { dados: reg ? JSON.parse(reg.dados || '{}') : {}, arquivos, updated_at: reg ? reg.updated_at : null, autorNome: reg ? reg.autorNome : null, existe: !!reg };
}

async function validacaoDaChaveDpo(companyId, chave, ano, dadosAtual) {
    if (String(chave).startsWith('acomp:')) return null; // calculada na tela (blocos do acompanhamento)
    if (chave === 'gop' || chave === 'cinco_s' || TOOLS_EXCLUSIVAS_DPO[chave]) return null; // calculada na tela
    if (chave === 'swot') return validarSwotDpo(dadosAtual, ano);
    const todos = {};
    for (const c of CHAVES_DIMENSIONAMENTO_DPO) todos[c] = c === chave && dadosAtual ? dadosAtual : (await carregarFerramentaDigitalDpo(companyId, c, ano)).dados;
    const arqs = await dbAll(`SELECT tipo FROM dpo_ferramentas_digitais_arquivos WHERE company_id = ? AND chave = 'orcamento' AND ano = ?`, [companyId, ano]);
    return validarDimensionamentoDpo(todos, arqs);
}

async function anosDaChaveDpo(companyId, chave) {
    const r = await dbAll(`SELECT ano FROM dpo_ferramentas_digitais WHERE company_id = ? AND chave = ? UNION SELECT ano FROM dpo_ferramentas_digitais_arquivos WHERE company_id = ? AND chave = ? ORDER BY ano DESC`, [companyId, chave, companyId, chave]);
    return r.map(x => x.ano);
}

// Migração: versões anteriores guardavam tudo num registro "ppr" — separa nas pastas novas.
async function migrarPprAntigoDpo(companyId) {
    const antigos = await dbAll(`SELECT * FROM dpo_ferramentas_digitais WHERE company_id = ? AND chave = 'ppr'`, [companyId]);
    for (const reg of antigos) {
        const d = JSON.parse(reg.dados || '{}');
        const partes = { sim_entrega: d.entrega, sim_armazem: d.armazem, orcamento: d.orcamento };
        for (const [chave, dados] of Object.entries(partes)) {
            if (!dados) continue;
            await new Promise((resolve) => db.run(`INSERT OR IGNORE INTO dpo_ferramentas_digitais (company_id, chave, ano, dados, updated_by, updated_at) VALUES (?, ?, ?, ?, ?, ?)`,
                [companyId, chave, reg.ano, JSON.stringify(dados), reg.updated_by, reg.updated_at], () => resolve()));
        }
        await new Promise((resolve) => db.run(`UPDATE dpo_ferramentas_digitais SET chave = 'ppr_legado' WHERE company_id = ? AND chave = 'ppr' AND ano = ?`, [companyId, reg.ano], () => resolve()));
    }
    await new Promise((resolve) => db.run(`UPDATE dpo_ferramentas_digitais_arquivos SET chave = CASE WHEN tipo = 'ppr_armazem' THEN 'sim_armazem' WHEN tipo = 'ppr_entrega' THEN 'sim_entrega' ELSE 'orcamento' END WHERE company_id = ? AND chave = 'ppr'`, [companyId], () => resolve()));
}

async function respostaFerramentaDigitalDpo(f, chave, ano, ctxReq) {
    if (CHAVES_DIMENSIONAMENTO_DPO.includes(chave)) await migrarPprAntigoDpo(f.companyId);
    const r = await carregarFerramentaDigitalDpo(f.companyId, chave, ano);
    if (chave === 'gop') r.dados = garantirPlanosGopDpo(r.dados);
    if (chave === 'cinco_s') r.dados = garantirPlanos5sDpo(r.dados);
    const pergunta = f.pilar ? perguntaDoPilarDpo(f.pilar, f.pergunta) : null;
    let ctx;
    if (ctxReq && (chave === 'capex' || chave === 'p3a')) {
        const principal = await ehGestorPrincipalDpo(f.companyId, ctxReq.user.userId);
        const usuarios = chave === 'capex' ? await dbAll(`SELECT id, name FROM users WHERE company_id = ? AND role = 'client_admin' ORDER BY name`, [f.companyId]) : undefined;
        const emp = await dbGet(`SELECT gestor_principal_id FROM companies WHERE id = ?`, [f.companyId]);
        ctx = { ehMaster: ctxReq.user.role === 'admin', souGestorPrincipal: principal, meuUserId: ctxReq.user.userId, usuarios, temGestorPrincipal: !!(emp && emp.gestor_principal_id) };
    }
    return {
        ctx, chave, titulo: f.titulo, pilar: f.pilar || null, pilarLabel: f.pilar ? DPO_AMBEV_DATA[f.pilar].label : (f.pilarLabel || ''), pergunta: f.pergunta, modeloGop: f.modeloGop || undefined, modelo5s: f.modelo5s || undefined, modeloManutencao: chave === 'manutencao' ? MODELO_MANUTENCAO_DPO : undefined,
        perguntaTexto: pergunta ? pergunta.pergunta.questao : '', verificacao: pergunta ? pergunta.pergunta.verificacao : '',
        companyId: Number(f.companyId), ano, anos: await anosDaChaveDpo(f.companyId, chave), ...r, template: f.template || null,
        validacao: await validacaoDaChaveDpo(f.companyId, chave, ano, r.dados)
    };
}

app.get('/api/dpo/ferramentas-digitais/:chave', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const f = await resolverFerramentaDigitalDpo(req, res, req.params.chave, req.query.company_id);
        if (!f) return;
        res.json(await respostaFerramentaDigitalDpo(f, req.params.chave, anoValidoDpo(req.query.ano), req));
    } catch (e) {
        console.error('Erro ao carregar ferramenta digital DPO:', e.message);
        res.status(500).json({ error: 'Erro ao carregar a ferramenta.' });
    }
});

// Visão das pastas do Simulador de Dimensionamento num ano (os 4 registros + anos de cada pasta).
app.get('/api/dpo/dimensionamento', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const f = await resolverFerramentaDigitalDpo(req, res, 'orcamento', req.query.company_id);
        if (!f) return;
        await migrarPprAntigoDpo(f.companyId);
        const ano = anoValidoDpo(req.query.ano);
        const pastas = {};
        for (const c of CHAVES_DIMENSIONAMENTO_DPO) {
            const r = await carregarFerramentaDigitalDpo(f.companyId, c, ano);
            pastas[c] = { ...r, anos: await anosDaChaveDpo(f.companyId, c), titulo: FERRAMENTAS_DIGITAIS_DPO[c].titulo };
        }
        const pergunta = perguntaDoPilarDpo('planejamento', '1.1');
        res.json({
            ano, companyId: Number(f.companyId), pastas, pilarLabel: DPO_AMBEV_DATA.planejamento.label, pergunta: '1.1',
            perguntaTexto: pergunta ? pergunta.pergunta.questao : '', verificacao: pergunta ? pergunta.pergunta.verificacao : '',
            validacao: validarDimensionamentoDpo(Object.fromEntries(Object.entries(pastas).map(([k, v]) => [k, v.dados])), pastas.orcamento.arquivos)
        });
    } catch (e) {
        console.error('Erro ao carregar o Simulador de Dimensionamento:', e.message);
        res.status(500).json({ error: 'Erro ao carregar o Simulador de Dimensionamento.' });
    }
});

app.put('/api/dpo/ferramentas-digitais/:chave', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const f = await resolverFerramentaDigitalDpo(req, res, req.params.chave, req.body.company_id);
        if (!f) return;
        const ano = anoValidoDpo(req.body.ano);
        let dados = req.body.dados && typeof req.body.dados === 'object' ? req.body.dados : {};
        if (req.params.chave === 'gop') dados = garantirPlanosGopDpo(dados);
        if (req.params.chave === 'cinco_s') dados = garantirPlanos5sDpo(dados);
        if ((req.params.chave === 'p3a' || req.params.chave === 'capex') && req.user.role !== 'admin') dados = await protegerCamposDpo(req.params.chave, f.companyId, ano, dados, req.user);
        const json = JSON.stringify(dados);
        if (json.length > 1500000) return res.status(400).json({ error: 'Dados grandes demais para salvar.' });
        await new Promise((resolve, reject) => db.run(
            `INSERT INTO dpo_ferramentas_digitais (company_id, chave, ano, dados, updated_by, updated_at) VALUES (?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
             ON CONFLICT(company_id, chave, ano) DO UPDATE SET dados = excluded.dados, updated_by = excluded.updated_by, updated_at = CURRENT_TIMESTAMP`,
            [f.companyId, req.params.chave, ano, json, req.user.userId], (err) => err ? reject(err) : resolve()
        ));
        res.json({ message: 'Salvo!', validacao: await validacaoDaChaveDpo(f.companyId, req.params.chave, ano, dados), acoes: req.params.chave === 'gop' || req.params.chave === 'cinco_s' ? dados.acoes : undefined });
    } catch (e) {
        console.error('Erro ao salvar ferramenta digital DPO:', e.message);
        res.status(400).json({ error: 'Erro ao salvar.' });
    }
});

// Campos que só o Master (parâmetros do P1A/P3A) ou o Gestor principal (aprovadores de CAPEX) podem mudar;
// aprovação de CAPEX só vale quando feita pelo próprio aprovador cadastrado.
async function protegerCamposDpo(chave, companyId, ano, dados, user) {
    const atual = (await carregarFerramentaDigitalDpo(companyId, chave, ano)).dados || {};
    if (chave === 'p3a') { dados.paramMaster = atual.paramMaster; return dados; }
    const principal = await ehGestorPrincipalDpo(companyId, user.userId);
    if (!principal) dados.aprovadores = atual.aprovadores;
    const aprovadores = Array.isArray(dados.aprovadores) ? dados.aprovadores : [];
    const antes = Object.fromEntries((atual.itens || []).map(i => [i.id, i]));
    (dados.itens || []).forEach(i => {
        const a = antes[i.id] || {}, novo = JSON.stringify(i.aprovacao || null), velho = JSON.stringify(a.aprovacao || null);
        if (novo === velho) return;
        const ok = i.aprovacao && Number(i.aprovacao.userId) === Number(user.userId) && aprovadores.some(x => Number(x.userId) === Number(user.userId));
        if (!ok && i.aprovacao) i.aprovacao = a.aprovacao; // aprovação forjada: mantém a anterior
    });
    return dados;
}

// Abre a pasta de um ano novo. Copia do ano de origem só o que é "estrutura":
// parâmetros dos simuladores; RACI e KPIs do orçamento; áreas, Sonho e responsáveis da SWOT.
app.post('/api/dpo/ferramentas-digitais/:chave/novo-ano', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const chaves = req.params.chave === 'dimensionamento' ? CHAVES_DIMENSIONAMENTO_DPO : [req.params.chave];
        const f = await resolverFerramentaDigitalDpo(req, res, chaves[0], req.body.company_id);
        if (!f) return;
        const ano = anoValidoDpo(req.body.ano);
        const criados = [];
        for (const chave of chaves) {
            const existe = await dbGet(`SELECT 1 FROM dpo_ferramentas_digitais WHERE company_id = ? AND chave = ? AND ano = ?`, [f.companyId, chave, ano]);
            if (existe) continue;
            let base = {};
            if (req.body.copiarDe) {
                const origem = (await carregarFerramentaDigitalDpo(f.companyId, chave, anoValidoDpo(req.body.copiarDe))).dados;
                if (String(chave).startsWith('acomp:')) base = {
                    reunioes: origem.reunioes, padroes: origem.padroes, riscos: origem.riscos, processos: origem.processos, __seeded: true,
                    indicadores: (origem.indicadores || []).map(i => ({ id: i.id, nome: i.nome, unidade: i.unidade, meta: i.meta, sentido: i.sentido }))
                };
                else if (chave === 'swot') base = { areas: origem.areas, sonho: origem.sonho, responsaveis: origem.responsaveis, vinculoDNMP: origem.vinculoDNMP, vinculoDNMPTexto: origem.vinculoDNMPTexto };
                else if (chave === 'orcamento') base = { raci: origem.raci, kpis: origem.kpis, processo: origem.processo };
                else if (chave === 'sonho') base = { frase: origem.frase, conexao: origem.conexao, exposicao: origem.exposicao, kpis: (origem.kpis || []).map(k => ({ id: k.id, nome: k.nome, pilar: k.pilar, sentido: k.sentido, meta: k.meta, unidade: k.unidade, acumula: k.acumula, nivel: k.nivel, pai: k.pai, cor: k.cor, rotuloMeta: k.rotuloMeta, slot: k.slot, real: [] })) };
                else if (chave === 'ans') base = { acordo: origem.acordo, volPadrao: origem.volPadrao, tolerancia: origem.tolerancia };
                else if (chave === 'visibilidade') base = { colaboradores: origem.colaboradores, indicadores: origem.indicadores, incentivo: origem.incentivo };
                else if (chave === 'riscos') base = { riscos: origem.riscos, respostas: origem.respostas, retomada: origem.retomada, retomadaRevisao: origem.retomadaRevisao, retomadaLocal: origem.retomadaLocal, revisoes: origem.revisoes };
                else if (chave === 'capex') base = { aprovadores: origem.aprovadores, cadastros: origem.cadastros, premissas: origem.premissas, emergencial: origem.emergencial, itens: (origem.itens || []).filter(i => i.etapa !== 'Concluído' && i.etapa !== 'Cancelado') };
                else if (chave === 'p3a') base = { ...origem, __resumo: undefined };
                else if (chave === 'manutencao') base = { modelo: origem.modelo, fornecedores: origem.fornecedores, raci: origem.raci, slaChamados: origem.slaChamados, chamados: (origem.chamados || []).filter(c => c.status !== 'Concluído' && c.status !== 'Cancelado'), acoes: (origem.acoes || []).filter(a => a.status !== 'Concluída') };
                else if (chave === 'cinco_s') { const o = prepararDados5sDpo(origem); base = { modelo: o.modelo, areas: o.areas, motoristas: o.motoristas, departamentos: o.departamentos, auditorias: [], acoes: o.acoes.filter(a => a.status !== 'Concluída' && !a.auto).map(a => ({ ...a, id: idGop() })) }; }
                else if (chave === 'gop') { const o = prepararDadosGopDpo(origem); base = { gops: Object.fromEntries(Object.entries(o.gops).map(([k, g]) => [k, { titulo: g.titulo, area: g.area, meta: g.meta, itens: g.itens, resp: {} }])), acoes: o.acoes.filter(a => a.status !== 'Concluída' && !a.auto).map(a => ({ ...a, id: idGop() })) }; }
                else base = { params: origem.params };
            }
            await new Promise((resolve, reject) => db.run(`INSERT INTO dpo_ferramentas_digitais (company_id, chave, ano, dados, updated_by) VALUES (?, ?, ?, ?, ?)`,
                [f.companyId, chave, ano, JSON.stringify(base), req.user.userId], (err) => err ? reject(err) : resolve()));
            criados.push(chave);
        }
        res.json({ message: criados.length ? `Pasta ${ano} criada${req.body.copiarDe ? ` (parâmetros copiados de ${req.body.copiarDe})` : ''}!` : `A pasta ${ano} já existe.`, criados });
    } catch (e) { res.status(400).json({ error: 'Erro ao criar a pasta do ano.' }); }
});

app.post('/api/dpo/ferramentas-digitais/:chave/arquivos', requireRole('admin', 'client_admin'), async (req, res) => {
    const { url, originalName, comentario, tipo } = req.body;
    if (!/^\/uploads\/[\w.\-]+$/.test(String(url || ''))) return res.status(400).json({ error: 'Envie o arquivo antes de salvar.' });
    try {
        const f = await resolverFerramentaDigitalDpo(req, res, req.params.chave, req.body.company_id);
        if (!f) return;
        const ano = anoValidoDpo(req.body.ano);
        await new Promise((resolve, reject) => db.run(
            `INSERT INTO dpo_ferramentas_digitais_arquivos (company_id, chave, ano, tipo, url, original_name, comentario, created_by) VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
            [f.companyId, req.params.chave, ano, String(tipo || 'outro').slice(0, 40), url, originalName ? String(originalName).slice(0, 200) : null, String(comentario || '').trim().slice(0, 500) || null, req.user.userId],
            (err) => err ? reject(err) : resolve()
        ));
        res.json({ message: 'Arquivo de atualização enviado!' });
    } catch (e) { res.status(400).json({ error: 'Erro ao registrar o arquivo.' }); }
});

app.delete('/api/dpo/ferramentas-digitais/arquivos/:id', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const a = await dbGet(`SELECT * FROM dpo_ferramentas_digitais_arquivos WHERE id = ?`, [req.params.id]);
        if (!a) return res.status(404).json({ error: 'Arquivo não encontrado.' });
        if (req.user.role === 'client_admin' && String(a.company_id) !== String(req.user.companyId)) return res.status(403).json({ error: 'Este arquivo não pertence à sua empresa.' });
        await new Promise((resolve, reject) => db.run(`DELETE FROM dpo_ferramentas_digitais_arquivos WHERE id = ?`, [a.id], (err) => err ? reject(err) : resolve()));
        res.json({ message: 'Arquivo removido!' });
    } catch (e) { res.status(400).json({ error: 'Erro ao remover o arquivo.' }); }
});

// ======================================================================
// DPO — GERENCIADOR GOP (planilha "Gerenciador GOP Revendas")
// Cada aba da planilha é uma GOP (Inventário, Obsolescence, TQI Armazém,
// Rating, TQI Distribuição, OTIF): perguntas com peso e SIM/NÃO por mês.
// % do mês = soma dos pesos SIM ÷ soma dos pesos (N/A fora); meta 80%.
// Todo NÃO de um mês abre automaticamente um plano de ação daquele mês
// (100% dos meses NOK com plano). Salvo por ano em dpo_ferramentas_digitais.
// ======================================================================
// Modelo embutido (as 6 abas da planilha) — não depende de arquivo externo.
const GOP_MODELO_DPO = [{"chave": "inventario", "titulo": "Inventário", "aba": "INVENTÁRIO", "area": "Armazém", "meta": 80.0, "itens": [{"id": "inventario-1", "num": 1, "texto": "A contagem de estoque é realizada em todos os produtos ( Matéria prima, embalagens, retornáveis e produto acabado)", "peso": 100}, {"id": "inventario-2", "num": 2, "texto": "A contagem de estoque inclui produtos fracionados (Lastros e caixas)", "peso": 100}, {"id": "inventario-3", "num": 3, "texto": "A contagem de estoque abrange produtos em processo de reembalagem (Repack)", "peso": 100}, {"id": "inventario-4", "num": 4, "texto": "A contagem de estoque cobre todos os armazéns pelos quais a unidade é responsável - incluindo armazéns de terceiros, armazéns de exportação e importação, caminhões carregados no local e em trânsito", "peso": 100}, {"id": "inventario-5", "num": 5, "texto": "O Gerente de Logística e Gerente de Financeiro estão assinando todas as contagens de inventário.", "peso": 100}, {"id": "inventario-6", "num": 6, "texto": "Segregam adequadamente as funções existentes entre o pessoal responsável pela contagem e recontagem de estoque.", "peso": 100}, {"id": "inventario-7", "num": 7, "texto": "Existe um Padrão bem documentado para descrever adequadamente o processo de entrada de materiais (tanto da área de recebimento quanto da cervejaria)", "peso": 100}, {"id": "inventario-8", "num": 8, "texto": "Há um Padrão bem documentado para descrever adequadamente o processo de captura de todas as perdas e quebras de mercadorias", "peso": 100}, {"id": "inventario-9", "num": 9, "texto": "Há um Padrão bem documentado para descrever adequadamente o processo de inventário (cobrindo pré-reunião, processo real de tomada de inventário, pós-reunião) e processo de aprovação mensal.", "peso": 100}, {"id": "inventario-10", "num": 10, "texto": "Existe um Padrão bem documentado para descrever adequadamente o processo de inventário enquanto as operações não são (totalmente) interrompidas.", "peso": 100}, {"id": "inventario-11", "num": 11, "texto": "Há um Padrão bem documentado para descrever adequadamente o procedimento de recontagem", "peso": 100}, {"id": "inventario-12", "num": 12, "texto": "Há um Padrão bem documentado para descrever adequadamente o sistema de gerenciamento de peças de reposição", "peso": 100}, {"id": "inventario-13", "num": 13, "texto": "Calendário DTO esta em vigor e seguido para todos os Padrões mencionados acima.", "peso": 100}, {"id": "inventario-14", "num": 14, "texto": "Existe um sistema integrado de contagem digital (WMS, tablets, handhelds) que está sendo utilizado para o controle de estoque e ciclos de contagem.", "peso": 100}, {"id": "inventario-15", "num": 15, "texto": "Caso o WMS e o sistema contábil (Promax) não tenham interface automática, é realizada uma reconciliação diária entre os sistemas e os ajustes são registrados em tempo hábil para garantir o alinhamento total entre os dois sistemas.", "peso": 100}, {"id": "inventario-16", "num": 16, "texto": "Não há estoque negativo registrado no sistema (Promax)", "peso": 100}, {"id": "inventario-17", "num": 17, "texto": "Existe uma solução de backup - (Plano de Contingência) para os processos de gerenciamento de estoque em caso de falha do sistema PROMAX /WMS.", "peso": 100}, {"id": "inventario-18", "num": 18, "texto": "Todos os registros de ajustes de estoque são mantidos - contendo data, usuários e perdas detalhadas por tipo de produto.", "peso": 100}, {"id": "inventario-19", "num": 19, "texto": "Todas as perdas relacionadas são registradas corretamente de acordo com o plano de contas do SCL / Pacote de prejuízo (Supply Chain Loss)", "peso": 100}, {"id": "inventario-20", "num": 20, "texto": "Todos os materiais são revisados para confirmar que ainda estão alinhados com a identidade da marca atual e, portanto, ainda podem ser usados. Caso contrário, eles são cancelados.", "peso": 100}, {"id": "inventario-21", "num": 21, "texto": "Todas as peças sobressalentes obsoletas identificadas são baixadas de acordo com as regras definidas de qualidade", "peso": 100}, {"id": "inventario-22", "num": 22, "texto": "O KPI de precisão do inventário (resultados financeiros e operacionais) está sendo rastreado no sistema de inventário oficial", "peso": 100}, {"id": "inventario-23", "num": 23, "texto": "Pelo menos 4 vezes por ano, a contagem do inventário está sendo feita junto com a equipe do financeiro com uma contagem independente.", "peso": 100}, {"id": "inventario-24", "num": 24, "texto": "O processo FEFO é medido e os resultados estão melhor que a meta", "peso": 100}, {"id": "inventario-25", "num": 25, "texto": "A unidade possui o time necessário para garantir o controle de estoque", "peso": 100}, {"id": "inventario-26", "num": 26, "texto": "Existem medidas extras para proteger o estoque contra roubo em momentos excepcionais.", "peso": 100}, {"id": "inventario-27", "num": 27, "texto": "Treinamentos foram dados aos funcionários que irão participar do inventários e equipes de contagem.", "peso": 100}, {"id": "inventario-28", "num": 28, "texto": "Os resultados do inventário acima do limite definido são escalados (Logística, Suprimento e Finanças)", "peso": 100}, {"id": "inventario-29", "num": 29, "texto": "Para qualquer resultado de inventário acima do limite, relatórios estão sendo feitos para auxiliar na investigação.", "peso": 100}]}, {"chave": "obsolescence", "titulo": "Obsolescence", "aba": "OBSOLESCENCE", "area": "Armazém", "meta": 80.0, "itens": [{"id": "obsolescence-1", "num": 1, "texto": "A unidade compartilha os objetivos de obsolescência com as áreas de vendas, marketing, logística e outras áreas relevantes para a unidade ?", "peso": 100}, {"id": "obsolescence-2", "num": 2, "texto": "A unidade tem um orçamento de obsolescência alocado para cada inovação antes do lançamento e, em seguida, os valores atuais são monitorados em relação a eles?", "peso": 100}, {"id": "obsolescence-3", "num": 3, "texto": "A equipe de vendas tem visibilidade para ver os estoques do armazém em tempo real?", "peso": 100}, {"id": "obsolescence-4", "num": 4, "texto": "A equipe de vendas tem visibilidade para ver as datas previstas de chegada para SKUs fora de estoque?", "peso": 100}, {"id": "obsolescence-5", "num": 5, "texto": "A equipe de vendas tem visibilidade semanal (no mínimo) para analisar as SKUs fora do intervalo acima em risco de obsolescência, com dias de vendas, contagem total de produtos e valor monetário em moeda local?", "peso": 100}, {"id": "obsolescence-6", "num": 6, "texto": "A unidade tem uma rotina entre liderança de Vendas e Logística para caminhar fisicamente pelo armazém e observar SKUs novos e fora de alcance?", "peso": 100}, {"id": "obsolescence-7", "num": 7, "texto": "Existem tarefas de check diário de idade dos produtos pela equipe do armazém?", "peso": 100}, {"id": "obsolescence-8", "num": 8, "texto": "Existe um processo para prever ou encomendar novos SKUs que não tenham tendências históricas de vendas?", "peso": 100}]}, {"chave": "tqi_arm", "titulo": "TQI - ARMAZÉM", "aba": "TQI - ARM", "area": "Armazém", "meta": 80.0, "itens": [{"id": "tqi_arm-1", "num": 1, "texto": "A unidade treinou os funcionários do armazém para completar várias habilidades/tarefas que poderiam ocorrer durante suas horas de trabalho para melhorar a flexibilidade do trabalho?", "peso": 100}, {"id": "tqi_arm-2", "num": 2, "texto": "A unidade concluiu uma avaliação do local com o fabricante de equipamento para determinar os tipos de equipamentos mais produtivos antes da compra de novos MHE (empilhadeira, paleteira...) ? Todas as novas compras da MHE ( (empilhadeira, paleteira...) estão de acordo com as Especificações Técnicas Globais?", "peso": 100}, {"id": "tqi_arm-3", "num": 3, "texto": "A unidade utilizou alguma Ferramenta de Otimização de espaço para analisar o layout do armazém em busca de oportunidades de melhoria de produtividade e testar diferentes cenários?", "peso": 100}, {"id": "tqi_arm-4", "num": 4, "texto": "A unidade fornece EPIs durante o processo de saída do caminhão ou de forma self service (ou seja, armários com código do funcionário) no caso de EPIs serem necessários durante o turno?", "peso": 100}, {"id": "tqi_arm-5", "num": 5, "texto": "Os funcionários do armazém podem visualizar seus IVs de produtividade em tempo real por meio de estações de trabalho do operador ou salas de equipe?", "peso": 100}, {"id": "tqi_arm-6", "num": 6, "texto": "A unidade utiliza intercalação de tarefas dentro de seus turnos?", "peso": 100}, {"id": "tqi_arm-7", "num": 7, "texto": "A unidade analisa a performance mais baixa para cada atividade/tarefa do armazém e conduz a análise de causa raiz para melhorar os resultados?", "peso": 100}, {"id": "tqi_arm-8", "num": 8, "texto": "A unidade rastreia os KPIs de complexidade da carga de saída (palete cheio, escolha manual, camada, tamanho da queda, caixas/palete, etc.) e revisa com as principais partes interessadas? As ações estão definidas para diminuir a complexidade da carga de saída e melhorar a produtividade?", "peso": 100}, {"id": "tqi_arm-9", "num": 9, "texto": "A unidade rastreia erros de carregamento que causam retrabalho e conduz a análise de causa raiz nos maiores contribuidores?", "peso": 100}, {"id": "tqi_arm-10", "num": 10, "texto": "A unidade utiliza o Sistema de Gerenciamento de Empilhadeira (FMS)/ Telemetria para monitorar os contribuintes de produtividade, como tempo ocioso e utilização? As horas de empilhadeira estão relacionadas com as horas de trabalho para avaliar oportunidades?", "peso": 100}]}, {"chave": "rating", "titulo": "RATING", "aba": "RATING", "area": "Entrega", "meta": 80.0, "itens": [{"id": "rating-1", "num": 1, "texto": "A unidade tem um processo/sigla para ajudar a equipe de entrega a entender facilmente os passos para fornecer um atendimento de qualidade ao cliente da ABI? (ou seja, SERVIR)", "peso": 100}, {"id": "rating-2", "num": 2, "texto": "A unidade utiliza feedback Rate my Delivery (RMD) para treinar equipes profissionais de entrega?", "peso": 100}, {"id": "rating-3", "num": 3, "texto": "A unidade tem um treinamento para educar os clientes sobre RMD? (ou seja, finalidade para RMD, como completar, quando, etc.)", "peso": 100}, {"id": "rating-4", "num": 4, "texto": "A unidade tem um treinamento para orientar os clientes sobre como verificar o status de seus pedidos? (ou seja, acessar o aplicativo/página da Web, comunicações esperadas vindas da unidade, como se inscrever para comunicação)", "peso": 100}, {"id": "rating-5", "num": 5, "texto": "A unidade tem uma rotina para validar se os dados dos PDVs para entrega estão corretos? (ou seja, detalhes de contato, método de entrega, local de entrega)", "peso": 100}, {"id": "rating-6", "num": 6, "texto": "A unidade tem uma maneira de comunicar a pesquisa RMD que reduz o viés dos clientes? (ou seja, o profissional de entrega não está na frente do entrevistado)", "peso": 100}, {"id": "rating-7", "num": 7, "texto": "A equipe de gestão de entregas tem uma avaliação para orientar as observações sobre como os profissionais de entrega atendem os clientes?", "peso": 100}, {"id": "rating-8", "num": 8, "texto": "A unidade analisa a taxa de resposta e tem um plano para aumentar as respostas? (ou seja, visitas aos PDVs, equipe de vendas, plano de comunicação, dia(s) de foco da linha de frente)", "peso": 100}, {"id": "rating-9", "num": 9, "texto": "A unidade tem um incentivo ou reconhecimento mensal do profissional de entrega com a maior nota RMD?", "peso": 100}]}, {"chave": "tqi_dis", "titulo": "TQI - Distribuição", "aba": "TQI - DIS", "area": "Entrega", "meta": 80.0, "itens": [{"id": "tqi_dis-1", "num": 1, "texto": "A unidade possui um mapeamento da quantidade de equipes de entrega (motoristas e ajudantes) necessários para executar as entregas em cada cliente (fator FTE médio)? A unidade tem um processo para negociar com a equipe de manutenção (ou seja, empilhadeira, paleteira) para aumentar a produtividade e reduzir os problemas de segurança?", "peso": 100}, {"id": "tqi_dis-2", "num": 2, "texto": "Os profissionais de entrega são designados para a mesma área/zona/setor mais de 75% do tempo, para que possam garantir o atendimento ao cliente e a produtividade das rotas?", "peso": 100}, {"id": "tqi_dis-3", "num": 3, "texto": "A equipe de entrega forneceu à equipe de armazém um cronograma de carregamento de caminhões de como priorizar o picking e o carregamento de caminhões? (ou seja, horários de início da rota, distância) Se houver problemas com caminhões que não estão prontos para uma saída de entrega no prazo, a unidade usa uma ANS para identificar a lacuna e trabalhar em conjunto para fechar?", "peso": 100}, {"id": "tqi_dis-4", "num": 4, "texto": "A unidade analisa as restrições realistas da unidade e cria horários de início escalonados/diferentes para a equipe de entrega? (ou seja, acesso ao estacionamento, tráfego para as instalações, horários de início do cliente, saída do portão do caminhão, validação de inventário).", "peso": 100}, {"id": "tqi_dis-5", "num": 5, "texto": "A unidade tem os equipamentos necessários para execução da rotina de entrega das equipes liberados para uso, carregados, fidelizados por motorista e disponíveis antes do início da rota? A unidade possui backups desses equipamentos em caso de falha?", "peso": 100}, {"id": "tqi_dis-6", "num": 6, "texto": "A unidade tem um layout de estacionamento que é respeitado assim que os caminhões são devolvidos e depois que o armazém os carrega, para que os profissionais de entrega possam encontrar rapidamente seu caminhão antes da partida?", "peso": 100}, {"id": "tqi_dis-7", "num": 7, "texto": "A unidade fornece EPIs durante o processo de saída do caminhão ou de forma self service (ou seja, armários com código do funcionário) no caso de EPIs serem necessários antes da partida?", "peso": 100}, {"id": "tqi_dis-8", "num": 8, "texto": "O time da liderança das equipes da unidade organiza algum mapa de rota ou documentação necessária antes do team room para garantir uma saída rápida do team room de entrega?", "peso": 100}]}, {"chave": "otif", "titulo": "OTIF", "aba": "OTIF", "area": "Entrega", "meta": 80.0, "itens": [{"id": "otif-1", "num": 1, "texto": "A unidade só tem SKUs em estoque visíveis para os clientes dentro da plataforma de pedidos? Ao dar visibilidade de estoque aos clientes, a unidade está utilizando dados de precisão de estoque para fornecer a disponibilidade de nível de SKU mais precisa aos clientes?", "peso": 100}, {"id": "otif-2", "num": 2, "texto": "A unidade permite que os motoristas alterem os tipos de pagamento com base em tipos aprovados, dividam os tipos de pagamento, apliquem créditos e/ou apliquem recompensas durante a entrega?", "peso": 100}, {"id": "otif-3", "num": 3, "texto": "A equipe de operações fornece à equipe comercial visibilidade/notificação ao vivo de uma possível recusa, cronômetro como tempo para resolver a possível recusa e as etapas seguidas no processo para salvar uma possível recusa?", "peso": 100}, {"id": "otif-4", "num": 4, "texto": "Existe um IP para monitorar o tempo desde a abertura até o fechamento de uma possível recusa? A equipe propõe melhores processos/rotinas de comunicação com diferentes equipes que possam acelerar esse processo de abertura para fechamento? (ou seja, WhatsApp com logística/comercial)", "peso": 100}, {"id": "otif-5", "num": 5, "texto": "O profissional de entrega tem impacto salarial quando ocorrem recusas? (ou seja, não é pago por certos tipos de recusa relacionados à logística - janela de tempo perdida, quebra profissional de entrega, etc.)", "peso": 100}, {"id": "otif-6", "num": 6, "texto": "A unidade tem uma maneira de informar o cliente de uma tentativa de entrega que não foi concluída, um motivo e carimbo de data/hora da tentativa? Se a unidade não tem essa tecnologia, eles deixam uma nota de estávamos aqui com as informações acima?", "peso": 100}, {"id": "otif-7", "num": 7, "texto": "A unidade tem uma rotina semanal de se reunir com os 10% melhores profissionais de entrega da unidade com as principais recusas logísticas? (quebra, perda de janela de tempo, etc.)", "peso": 100}, {"id": "otif-8", "num": 8, "texto": "A unidade tem uma rotina semanal com o time de vendas para tratar as devoluções que não são motivo logística e encontrarem juntos a causa raiz e criarem um plano de ação para tratarem?", "peso": 100}, {"id": "otif-9", "num": 9, "texto": "A unidade tem um incentivo/reconhecimento mensal para as equipes e time de vendas com menor percentual de devolução?", "peso": 100}, {"id": "otif-10", "num": 10, "texto": "A unidade tem um plano para visitar clientes críticos que possuem baixo resultado de OTIF e/ou devoluções reincidentes? (Quantidade de PDVs deve ser definida de acordo com o tamanho da unidade e análise do pareto com a criticidade)", "peso": 100}]}];
const STATUS_ACAO_GOP_DPO = ['Não iniciado', 'Em andamento', 'Concluída', 'Atrasado'];
const semAcentoGop = s => String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').trim().toUpperCase();
function normalizarRespostaGop(v) {
    const t = semAcentoGop(v);
    if (t === 'SIM' || t === 'S' || t === 'OK') return 'SIM';
    if (t === 'NAO' || t === 'N' || t === 'NOK') return 'NÃO';
    if (t === 'NA' || t === 'N/A' || t === 'N.A.') return 'N/A';
    return '';
}
function idGop() { return crypto.randomBytes(5).toString('hex'); }
function chaveGopPorNome(nome) {
    const n = semAcentoGop(nome).replace(/[^A-Z0-9]/g, '');
    const m = GOP_MODELO_DPO.find(g => semAcentoGop(g.aba).replace(/[^A-Z0-9]/g, '') === n || semAcentoGop(g.titulo).replace(/[^A-Z0-9]/g, '') === n);
    return m ? m.chave : null;
}
// Garante a estrutura das 6 GOPs (a partir do modelo) sem apagar o que a revenda já tem.
function prepararDadosGopDpo(dados) {
    const d = dados && typeof dados === 'object' ? dados : {};
    d.gops = d.gops && typeof d.gops === 'object' ? d.gops : {};
    GOP_MODELO_DPO.forEach(m => {
        if (!d.gops[m.chave]) d.gops[m.chave] = { titulo: m.titulo, area: m.area, meta: m.meta, itens: m.itens.map(i => ({ ...i })), resp: {} };
        const g = d.gops[m.chave];
        g.itens = Array.isArray(g.itens) ? g.itens : [];
        g.resp = g.resp && typeof g.resp === 'object' ? g.resp : {};
        if (g.meta === undefined || g.meta === null || g.meta === '') g.meta = m.meta;
    });
    d.acoes = Array.isArray(d.acoes) ? d.acoes.filter(a => a && typeof a === 'object') : [];
    return d;
}
// Todo NÃO (item × mês) precisa de plano: cria o que falta; remove planos automáticos
// ainda vazios cujo item voltou a SIM/N/A.
function garantirPlanosGopDpo(dados) {
    const d = prepararDadosGopDpo(dados);
    const chaveAcao = a => `${a.gop}|${a.itemId}|${a.mes}`;
    const existentes = new Set(d.acoes.filter(a => a.itemId).map(chaveAcao));
    const nok = new Set();
    Object.entries(d.gops).forEach(([gk, g]) => g.itens.forEach(it => (g.resp[it.id] || []).forEach((v, mes) => {
        if (normalizarRespostaGop(v) !== 'NÃO') return;
        const k = `${gk}|${it.id}|${mes}`;
        nok.add(k);
        if (!existentes.has(k)) { d.acoes.push({ id: idGop(), gop: gk, itemId: it.id, mes, acao: '', causa: '', responsavel: '', prevista: '', status: 'Não iniciado', realizada: '', obs: '', auto: true }); existentes.add(k); }
    })));
    d.acoes = d.acoes.filter(a => !(a.auto && a.itemId && !nok.has(chaveAcao(a)) && !String(a.acao || '').trim() && !String(a.responsavel || '').trim()));
    return d;
}
function pctMesGopDpo(g, mes) {
    let soma = 0, total = 0, respondidas = 0;
    g.itens.forEach(it => {
        const v = normalizarRespostaGop((g.resp[it.id] || [])[mes]);
        const peso = Number(it.peso) || 100;
        if (v === 'N/A') return;
        if (v) respondidas++;
        total += peso; if (v === 'SIM') soma += peso;
    });
    return respondidas && total ? soma / total : null;
}
function valorCelulaGop(v) {
    if (v && typeof v === 'object' && !(v instanceof Date)) {
        if (Array.isArray(v.richText)) return v.richText.map(t => t.text).join('');
        if ('result' in v) return v.result;
        if (v.text !== undefined) return v.text;
        if (v.error) return null;
    }
    return v;
}
function dataCelulaGop(v) {
    v = valorCelulaGop(v);
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    if (typeof v === 'number' && v > 20000 && v < 80000) return new Date(Math.round((v - 25569) * 86400000)).toISOString().slice(0, 10);
    const t = String(v || '').trim();
    const br = t.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
    return br ? `${br[3]}-${br[2]}-${br[1]}` : (/^\d{4}-\d{2}-\d{2}/.test(t) ? t.slice(0, 10) : '');
}
// Lê a planilha (xlsx/xlsm): uma aba por GOP + aba AÇÕES.
function lerWorkbookGopDpo(wb) {
    const gops = {}, acoes = [], avisos = [];
    const MESES = ['JAN', 'FEV', 'MAR', 'ABR', 'MAI', 'JUN', 'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ'];
    wb.worksheets.forEach(ws => {
        const cel = (r, c) => valorCelulaGop(ws.getRow(r).getCell(c).value);
        const nomeAba = semAcentoGop(ws.name);
        if (nomeAba === 'ACOES' || nomeAba === 'ACAO') {
            let hdr = null;
            for (let r = 1; r <= Math.min(ws.rowCount, 40) && !hdr; r++) for (let c = 1; c <= 12; c++) if (semAcentoGop(cel(r, c)) === 'ACAO') { hdr = { r, c }; break; }
            if (!hdr) return;
            for (let r = hdr.r + 1; r <= ws.rowCount; r++) {
                const gopNome = cel(r, hdr.c - 1), acao = cel(r, hdr.c);
                if (!acao || !String(acao).trim()) continue;
                const gk = chaveGopPorNome(gopNome) || (gopNome ? String(gopNome).trim() : '');
                const st = STATUS_ACAO_GOP_DPO.find(s => semAcentoGop(s) === semAcentoGop(cel(r, hdr.c + 3))) || 'Não iniciado';
                acoes.push({ id: idGop(), gop: gk, itemId: null, mes: null, acao: String(acao).trim(), causa: '', responsavel: String(cel(r, hdr.c + 1) || '').trim(), prevista: dataCelulaGop(ws.getRow(r).getCell(hdr.c + 2).value), status: st, realizada: dataCelulaGop(ws.getRow(r).getCell(hdr.c + 4).value), obs: String(cel(r, hdr.c + 5) || '').trim(), origem: 'planilha' });
            }
            return;
        }
        let hdr = null;
        for (let r = 1; r <= Math.min(ws.rowCount, 80) && !hdr; r++) for (let c = 1; c <= 8; c++) if (semAcentoGop(cel(r, c)) === 'PERGUNTAS') { hdr = { r, c }; break; }
        if (!hdr) return;
        let jan = null;
        for (let c = hdr.c; c <= hdr.c + 8; c++) if (semAcentoGop(cel(hdr.r, c)) === 'JAN') { jan = c; break; }
        if (!jan) { avisos.push(`Aba "${ws.name}": não achei a coluna JAN.`); return; }
        const chave = chaveGopPorNome(ws.name) || ('gop_' + nomeAba.toLowerCase().replace(/[^a-z0-9]+/g, '_'));
        const modelo = GOP_MODELO_DPO.find(m => m.chave === chave);
        const titulo = String(cel(hdr.r - 2, hdr.c) || (modelo ? modelo.titulo : ws.name)).trim();
        const itens = [], resp = {};
        let meta = modelo ? modelo.meta : 80, n = 0;
        for (let r = hdr.r + 1; r <= ws.rowCount; r++) {
            const num = cel(r, hdr.c), txt = cel(r, hdr.c + 1);
            if (semAcentoGop(txt) === 'META') { const mt = Number(cel(r, jan)); if (mt) meta = mt <= 1 ? Math.round(mt * 1000) / 10 : mt; break; }
            if (num === null || num === undefined || num === '' || !txt || !String(txt).trim()) continue;
            n++;
            const id = modelo && modelo.itens[n - 1] ? modelo.itens[n - 1].id : `${chave}-${n}`;
            itens.push({ id, num: n, texto: String(txt).replace(/\s+/g, ' ').trim(), peso: Number(cel(r, jan - 1)) || 100 });
            resp[id] = MESES.map((_, i) => normalizarRespostaGop(cel(r, jan + i)));
        }
        if (!itens.length) { avisos.push(`Aba "${ws.name}": nenhuma pergunta encontrada.`); return; }
        gops[chave] = { titulo, area: modelo ? modelo.area : '', meta, itens, resp, aba: ws.name };
    });
    return { gops, acoes, avisos };
}

app.post('/api/dpo/gop/importar', requireRole('admin', 'client_admin'), async (req, res) => {
    const { url, originalName } = req.body;
    if (!/^\/uploads\/[\w.\-]+$/.test(String(url || ''))) return res.status(400).json({ error: 'Envie a planilha antes de importar.' });
    try {
        const f = await resolverFerramentaDigitalDpo(req, res, 'gop', req.body.company_id);
        if (!f) return;
        const ano = anoValidoDpo(req.body.ano);
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.readFile(path.join(PASTA_UPLOADS, path.basename(url)));
        const lido = lerWorkbookGopDpo(wb);
        if (!Object.keys(lido.gops).length) return res.status(400).json({ error: 'Não encontrei nenhuma aba de GOP (cabeçalho "Perguntas" + meses JAN a DEZ).', avisos: lido.avisos });
        const atual = prepararDadosGopDpo((await carregarFerramentaDigitalDpo(f.companyId, 'gop', ano)).dados);
        Object.entries(lido.gops).forEach(([k, g]) => { atual.gops[k] = g; });
        const jaTem = new Set(atual.acoes.map(a => `${a.gop}|${semAcentoGop(a.acao)}`));
        lido.acoes.forEach(a => { if (!jaTem.has(`${a.gop}|${semAcentoGop(a.acao)}`)) atual.acoes.push(a); });
        const dados = garantirPlanosGopDpo(atual);
        await new Promise((resolve, reject) => db.run(
            `INSERT INTO dpo_ferramentas_digitais (company_id, chave, ano, dados, updated_by, updated_at) VALUES (?, 'gop', ?, ?, ?, CURRENT_TIMESTAMP)
             ON CONFLICT(company_id, chave, ano) DO UPDATE SET dados = excluded.dados, updated_by = excluded.updated_by, updated_at = CURRENT_TIMESTAMP`,
            [f.companyId, ano, JSON.stringify(dados), req.user.userId], (err) => err ? reject(err) : resolve()));
        db.run(`INSERT INTO dpo_ferramentas_digitais_arquivos (company_id, chave, ano, tipo, url, original_name, comentario, created_by) VALUES (?, 'gop', ?, 'planilha_gop', ?, ?, ?, ?)`,
            [f.companyId, ano, url, originalName || 'Gerenciador GOP.xlsm', 'Importada para o sistema', req.user.userId], () => {});
        const planos = dados.acoes.filter(a => a.auto).length;
        res.json({ message: `${Object.keys(lido.gops).length} GOP(s) importada(s) e ${lido.acoes.length} ação(ões) da aba AÇÕES. ${planos} plano(s) por mês abertos para os itens NÃO.`, dados, avisos: lido.avisos });
    } catch (e) {
        console.error('Erro ao importar GOP:', e.message);
        res.status(400).json({ error: 'Não foi possível ler a planilha. Envie o Gerenciador GOP em .xlsx ou .xlsm.' });
    }
});

function exportarGopDpo(add, dados) {
    const d = prepararDadosGopDpo(dados);
    const MESES = MESES_CURTOS_DPO;
    Object.entries(d.gops).forEach(([k, g]) => {
        const linhas = g.itens.map(it => ({ n: it.num, t: it.texto, p: it.peso, ...Object.fromEntries(MESES.map((m, i) => ['m' + i, (g.resp[it.id] || [])[i] || ''])) }));
        linhas.push({ t: `% ${g.titulo}`, ...Object.fromEntries(MESES.map((m, i) => { const v = pctMesGopDpo(g, i); return ['m' + i, v === null ? '' : `${String(Math.round(v * 1000) / 10).replace('.', ',')}%`]; })) });
        linhas.push({ t: 'Meta', ...Object.fromEntries(MESES.map((m, i) => ['m' + i, `${g.meta}%`])) });
        add(String(g.titulo || k).slice(0, 28).replace(/[\\/*?:[\]]/g, '-'), [['Nº', 'n', 6], ['Pergunta', 't', 70], ['Peso', 'p', 8], ...MESES.map((m, i) => [m, 'm' + i, 8])], linhas);
    });
    const textoItem = a => { const g = d.gops[a.gop]; const it = g && g.itens.find(i => i.id === a.itemId); return it ? `${it.num}. ${it.texto}` : ''; };
    add('Ações', [['GOP', 'gop', 18], ['Mês', 'mes', 8], ['Item NOK', 'item', 50], ['Causa', 'causa', 30], ['Ação', 'acao', 45], ['Responsável', 'responsavel', 18], ['Data prevista', 'prevista', 13], ['Status', 'status', 14], ['Data realizada', 'realizada', 13], ['Observação', 'obs', 40]],
        d.acoes.map(a => ({ ...a, gop: d.gops[a.gop] ? d.gops[a.gop].titulo : a.gop, mes: a.mes === null || a.mes === undefined ? '' : MESES[a.mes], item: textoItem(a) })));
}

// ======================================================================
// DPO — GERENCIADOR 5S (planilha "5S Armazém e ADM") — Gestão 3.1
// Áreas com dono e auditor; auditoria mensal por área com as perguntas dos
// 5 sensos (S / N / NA). % do senso = S ÷ (S + N); 5S da área = média dos
// sensos avaliados; revenda = média das áreas. Meta 85%. Todo N abre na hora
// um plano de ação (dono da área como responsável).
// ======================================================================
const CINCO_S_MODELO_DPO = {"meta": 85, "sensos": [{"chave": "selecao", "titulo": "Seleção", "numero": 1, "perguntas": [{"id": "selecao-1", "num": "1.1", "texto": "A área está livre de equipamentos e/ou objetos (ex. máquinas, cadeiras, mesas, trava paletes, cones de sinalização, quadros de gestão à vista) quebrados e/ou sem utilização na área? Todos os equipamentos e/ou objetos são necessários?"}, {"id": "selecao-2", "num": "1.2", "texto": "A área está livre de Cópias desnecessárias (Padrões vencidos, Books sem utilização) de materiais de consulta?"}, {"id": "selecao-3", "num": "1.3", "texto": "A área está livre de objetos desnecessários nos armários, gavetas equipamentos?"}, {"id": "selecao-4", "num": "1.4", "texto": "Os objetos pessoais estão nos lugares corretos? (Não deve ter objetos pessoais nos postos de trabalho)"}]}, {"chave": "organizacao", "titulo": "Organização", "numero": 2, "perguntas": [{"id": "organizacao-1", "num": "2.1", "texto": "Existe identificação de materiais (mesas, salas, cadeiras, armarios)"}, {"id": "organizacao-2", "num": "2.2", "texto": "Os telefones estão identificados com o número do ramal?"}, {"id": "organizacao-3", "num": "2.3", "texto": "Os arquivos da Rede (Pastas de trabalho) estão organizados e de fácil acesso. Mostrando uma organização lógica, com nomes, para que todos consigam acessar (respeitando os limites de acesso)?"}, {"id": "organizacao-4", "num": "2.4", "texto": "O desktop do funcionário está devidamente organizado?"}, {"id": "organizacao-5", "num": "2.5", "texto": "Os padrões da área se encontram em local de fácil acesso, conhecido por todos? Os padrões estão organizados, quaisquer padrões podem ser encontrados facilmente?"}, {"id": "organizacao-6", "num": "2.6", "texto": "Os itens da area estão nos seus locais destinados? Existem placas/identificações para todos os itens?"}]}, {"chave": "limpeza", "titulo": "Limpeza", "numero": 3, "perguntas": [{"id": "limpeza-1", "num": "3.1", "texto": "Existe cronograma de limpeza na área? Está sendo cumprido?"}, {"id": "limpeza-2", "num": "3.2", "texto": "O lixo é recolhido com frequência?"}, {"id": "limpeza-3", "num": "3.3", "texto": "As mesas e o piso estão limpos?"}, {"id": "limpeza-4", "num": "3.4", "texto": "De modo geral a área passa a impressão de ser um ambiente limpo?"}, {"id": "limpeza-5", "num": "3.5", "texto": "A área está livre de alimentos ou restos de alimentos?"}]}, {"chave": "conservacao", "titulo": "Conservação", "numero": 4, "perguntas": [{"id": "conservacao-1", "num": "4.1", "texto": "Os equipamentos, utensílios, ferramentas e materiais estão em bom estado de conservação?"}, {"id": "conservacao-2", "num": "4.2", "texto": "As luminárias estão funcionando e estão em bom estado de conservação?"}, {"id": "conservacao-3", "num": "4.3", "texto": "Existem cabos de energia ou outros tipo de cabos soltos pela area?"}, {"id": "conservacao-4", "num": "4.4", "texto": "O piso da área está em bom estado? (sem buracos, cerâmica faltando e/ou quebradas, etc). As paredes da área estão em bom estado (pintura não deve estar descascando, não deve ter azulejos faltando ou quebrados, não deve ter manchas)? O telhado da área está em bom estado? As tubulações e escadas estão em bom estado?"}, {"id": "conservacao-5", "num": "4.5", "texto": "As tomadas e interruptores estão em bom estado e funcionando?"}]}, {"chave": "autodisciplina", "titulo": "Auto-Disciplina", "numero": 5, "perguntas": [{"id": "autodisciplina-1", "num": "5.1", "texto": "A operação conhece a sua responsabilidade na área? Sabe explicar o quadro de 5S? Qual a área sob sua responsabilidade e quais as atividades de 5S que precisa executar?"}, {"id": "autodisciplina-2", "num": "5.2", "texto": "Existe um quadro de gestão à vista com o resultado da ultima auditoria de 5s e ele esta atualizado?"}, {"id": "autodisciplina-3", "num": "5.3", "texto": "Todos os quadros de gestão à vista estão preenchidos e atualizados?"}, {"id": "autodisciplina-4", "num": "5.4", "texto": "As não conformidades levantadas nas auditorias passadas foram tratadas? (Só pontuar se houve não conformidade na auditoria anterior)"}, {"id": "autodisciplina-5", "num": "5.5", "texto": "As ações da ultima auditoria de 5s estão escritas no quadro de 5s da area e estão atualizadas?"}]}]};
const MESES_LONGOS_5S_DPO = ['JANEIRO', 'FEVEREIRO', 'MARCO', 'ABRIL', 'MAIO', 'JUNHO', 'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO'];
function resp5sDpo(v) { const t = semAcentoGop(v); return t === 'S' || t === 'SIM' ? 'S' : t === 'N' || t === 'NAO' ? 'N' : t === 'NA' || t === 'N/A' ? 'NA' : ''; }
function prepararDados5sDpo(dados) {
    const d = dados && typeof dados === 'object' ? dados : {};
    if (!d.modelo || !Array.isArray(d.modelo.sensos)) d.modelo = JSON.parse(JSON.stringify(CINCO_S_MODELO_DPO));
    d.areas = Array.isArray(d.areas) ? d.areas.filter(a => a && a.id) : [];
    d.auditorias = Array.isArray(d.auditorias) ? d.auditorias.filter(a => a && a.id) : [];
    d.acoes = Array.isArray(d.acoes) ? d.acoes.filter(a => a && a.id) : [];
    return d;
}
function garantirPlanos5sDpo(dados) {
    const d = prepararDados5sDpo(dados);
    const chave = a => `${a.auditoriaId}|${a.qid}`;
    const existentes = new Set(d.acoes.filter(a => a.auditoriaId).map(chave)), nok = new Set();
    const sensoDe = qid => { const s = d.modelo.sensos.find(x => x.perguntas.some(p => p.id === qid)); return s ? s.chave : ''; };
    // Auditorias antigas importadas da planilha: só a última de cada área abre plano automático.
    const ultimoMes = {};
    d.auditorias.forEach(au => { if (ultimoMes[au.areaId] === undefined || au.mes > ultimoMes[au.areaId]) ultimoMes[au.areaId] = au.mes; });
    d.auditorias.forEach(au => Object.entries(au.resp || {}).forEach(([qid, v]) => {
        if (resp5sDpo(v) !== 'N') return;
        if (au.origem === 'planilha' && au.mes !== ultimoMes[au.areaId]) return;
        const k = `${au.id}|${qid}`; nok.add(k);
        if (existentes.has(k)) return;
        const area = d.areas.find(a => a.id === au.areaId);
        d.acoes.push({ id: idGop(), auditoriaId: au.id, areaId: au.areaId, mes: au.mes, senso: sensoDe(qid), qid, acao: '', dono: area ? area.dono || '' : '', prevista: '', tratativa: '', status: 'Não iniciado', auto: true });
        existentes.add(k);
    }));
    d.acoes = d.acoes.filter(a => !(a.auto && a.auditoriaId && !nok.has(chave(a)) && !String(a.acao || '').trim()));
    return d;
}
// Lê a planilha 5S (xlsx/xlsm): "Donos de área", "Resultado Geral" (departamento), "Base" (auditorias) e "Gerenciador de Ações".
function lerWorkbook5sDpo(wb, ano) {
    const d = prepararDados5sDpo({}), avisos = [];
    const aba = nome => wb.worksheets.find(w => semAcentoGop(w.name).replace(/\s+/g, '') === semAcentoGop(nome).replace(/\s+/g, ''));
    const val = (ws, r, c) => valorCelulaGop(ws.getRow(r).getCell(c).value);
    const acharCab = (ws, rotulo, maxL) => { for (let r = 1; r <= Math.min(ws.rowCount, maxL || 30); r++) for (let c = 1; c <= 30; c++) if (semAcentoGop(val(ws, r, c)) === semAcentoGop(rotulo)) return { r, c }; return null; };
    const areaPorNome = {};
    const addArea = (nome, extra) => { const n = String(nome || '').trim(); if (!n || /^TOTAL$|^REVENDA$/i.test(n)) return null; const k = semAcentoGop(n); if (!areaPorNome[k]) { areaPorNome[k] = { id: idGop(), nome: n, depto: '', dono: '', auditor: '' }; d.areas.push(areaPorNome[k]); } Object.entries(extra || {}).forEach(([c, v]) => { if (v && String(v).trim()) areaPorNome[k][c] = String(v).trim(); }); return areaPorNome[k]; };
    const donos = aba('Donos de área');
    if (donos) { const h = acharCab(donos, 'Area'); if (h) for (let r = h.r + 1; r <= donos.rowCount; r++) addArea(val(donos, r, h.c), { dono: val(donos, r, h.c + 1), auditor: val(donos, r, h.c + 3) }); }
    else avisos.push('Aba "Donos de área" não encontrada.');
    const rg = aba('Resultado Geral');
    if (rg) { const h = acharCab(rg, 'Departamento'); if (h) for (let r = h.r + 1; r <= rg.rowCount; r++) { const a = areaPorNome[semAcentoGop(val(rg, r, h.c))]; if (a && val(rg, r, h.c + 1)) a.depto = String(val(rg, r, h.c + 1)).trim(); } }
    const base = aba('Base');
    const qids = {};
    d.modelo.sensos.forEach(s => s.perguntas.forEach((p, i) => { qids[semAcentoGop(`${i + 1} - ${s.titulo}`).replace(/[^A-Z0-9]/g, '')] = p.id; }));
    if (base) {
        const h = acharCab(base, 'Auditor');
        if (h) {
            const cols = {};
            for (let c = h.c; c <= h.c + 200; c++) { const t = semAcentoGop(val(base, h.r, c)).replace(/[^A-Z0-9]/g, ''); if (qids[t] && cols[qids[t]] === undefined) cols[qids[t]] = c; }
            for (let r = h.r + 1; r <= base.rowCount; r++) {
                const nomeArea = val(base, r, h.c + 1), mesTxt = semAcentoGop(val(base, r, h.c + 2));
                const mes = MESES_LONGOS_5S_DPO.indexOf(mesTxt);
                if (!nomeArea || mes < 0) continue;
                const area = addArea(nomeArea);
                const resp = {};
                Object.entries(cols).forEach(([qid, c]) => { const v = resp5sDpo(val(base, r, c)); if (v) resp[qid] = v; });
                if (!Object.keys(resp).length) continue;
                const ja = d.auditorias.find(a => a.areaId === area.id && a.mes === mes);
                const au = { id: ja ? ja.id : idGop(), areaId: area.id, mes, auditor: String(val(base, r, h.c) || '').trim(), resp, origem: 'planilha' };
                if (ja) Object.assign(ja, au); else d.auditorias.push(au);
            }
        }
    } else avisos.push('Aba "Base" não encontrada — auditorias não importadas.');
    // Resultado Geral: nota do mês informada na planilha (vale como resultado oficial da área naquele mês).
    if (rg) {
        const h = acharCab(rg, 'Departamento');
        if (h) {
            const colMes = {};
            for (let c = h.c; c <= h.c + 40; c++) { const m = MESES_LONGOS_5S_DPO.indexOf(semAcentoGop(val(rg, h.r, c))); if (m >= 0 && colMes[m] === undefined) colMes[m] = c; }
            for (let r = h.r + 1; r <= rg.rowCount; r++) {
                const area = areaPorNome[semAcentoGop(val(rg, r, h.c))];
                if (!area) continue;
                Object.entries(colMes).forEach(([m, c]) => {
                    const v = Number(val(rg, r, c));
                    if (!isFinite(v) || v <= 0 || v > 1.0001 || val(rg, r, c) === null || val(rg, r, c) === '') return;
                    let au = d.auditorias.find(a => a.areaId === area.id && a.mes === Number(m));
                    if (!au) { au = { id: idGop(), areaId: area.id, mes: Number(m), auditor: area.auditor || '', resp: {}, origem: 'planilha' }; d.auditorias.push(au); }
                    au.notaInformada = Math.round(v * 10000) / 100;
                });
            }
        }
    }
    const ga = aba('Gerenciador de Ações');
    if (ga) {
        const h = acharCab(ga, 'Dono');
        const sensoPorTexto = t => { const n = semAcentoGop(t).replace(/[^A-Z]/g, ''); const s = d.modelo.sensos.find(x => n.includes(semAcentoGop(x.titulo).replace(/[^A-Z]/g, ''))); return s ? s.chave : ''; };
        if (h) for (let r = h.r + 1; r <= ga.rowCount; r++) {
            const acao = val(ga, r, h.c + 2);
            if (!acao || !String(acao).trim()) continue;
            const dt = dataCelulaGop(ga.getRow(r).getCell(h.c - 1).value);
            const st = semAcentoGop(val(ga, r, h.c + 4));
            d.acoes.push({ id: idGop(), auditoriaId: null, areaId: null, mes: dt && Number(dt.slice(5, 7)) ? Number(dt.slice(5, 7)) - 1 : null, senso: sensoPorTexto(val(ga, r, h.c + 1)), qid: null, acao: String(acao).trim(), dono: String(val(ga, r, h.c) || '').trim(), prevista: dt || '', tratativa: String(val(ga, r, h.c + 3) || '').trim(), status: /CONCL|CONLC/.test(st) ? 'Concluída' : /ANDAM/.test(st) ? 'Em andamento' : 'Não iniciado', origem: 'planilha' });
        }
    }
    return { dados: garantirPlanos5sDpo(d), avisos };
}
async function lerArquivoPlanilhaDpo(caminho) {
    let arquivo = caminho;
    if (/\.xlsb$/i.test(caminho)) {
        // .xlsb (binário) — converte com o LibreOffice do servidor, se existir.
        const { execFile } = require('child_process');
        const saida = path.join(require('os').tmpdir(), 'conv-' + Date.now());
        fs.mkdirSync(saida, { recursive: true });
        await new Promise((resolve, reject) => execFile('soffice', ['--headless', '--convert-to', 'xlsx', '--outdir', saida, caminho], { timeout: 120000 }, err => err ? reject(new Error('xlsb')) : resolve()));
        arquivo = path.join(saida, path.basename(caminho).replace(/\.xlsb$/i, '.xlsx'));
    }
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(arquivo);
    return wb;
}
app.post('/api/dpo/cinco-s/importar', requireRole('admin', 'client_admin'), async (req, res) => {
    const { url, originalName } = req.body;
    if (!/^\/uploads\/[\w.\-]+$/.test(String(url || ''))) return res.status(400).json({ error: 'Envie a planilha antes de importar.' });
    try {
        const f = await resolverFerramentaDigitalDpo(req, res, 'cinco_s', req.body.company_id);
        if (!f) return;
        const ano = anoValidoDpo(req.body.ano);
        let wb;
        try { wb = await lerArquivoPlanilhaDpo(path.join(PASTA_UPLOADS, path.basename(url))); }
        catch (e) { return res.status(400).json({ error: /xlsb/.test(e.message) || /\.xlsb$/i.test(url) ? 'Não consegui ler o formato .xlsb aqui. Abra a planilha no Excel, use "Salvar como" → Pasta de Trabalho do Excel (.xlsx) e importe de novo.' : 'Não foi possível ler a planilha.' }); }
        const lido = lerWorkbook5sDpo(wb, ano);
        if (!lido.dados.areas.length) return res.status(400).json({ error: 'Não encontrei as áreas (aba "Donos de área").', avisos: lido.avisos });
        const atual = prepararDados5sDpo((await carregarFerramentaDigitalDpo(f.companyId, 'cinco_s', ano)).dados);
        // mescla: áreas por nome, auditorias por área+mês, ações novas
        const mapa = {};
        lido.dados.areas.forEach(a => { const ex = atual.areas.find(x => semAcentoGop(x.nome) === semAcentoGop(a.nome)); if (ex) { mapa[a.id] = ex.id; ['depto', 'dono', 'auditor'].forEach(c => { if (a[c]) ex[c] = a[c]; }); } else { atual.areas.push(a); mapa[a.id] = a.id; } });
        lido.dados.auditorias.forEach(au => { au.areaId = mapa[au.areaId] || au.areaId; const ex = atual.auditorias.find(x => x.areaId === au.areaId && x.mes === au.mes); if (ex) { ex.resp = au.resp; ex.auditor = au.auditor || ex.auditor; } else atual.auditorias.push(au); });
        const jaTem = new Set(atual.acoes.map(a => semAcentoGop(a.acao)));
        lido.dados.acoes.filter(a => !a.auditoriaId && !jaTem.has(semAcentoGop(a.acao))).forEach(a => atual.acoes.push(a));
        const dados = garantirPlanos5sDpo(atual);
        await new Promise((resolve, reject) => db.run(
            `INSERT INTO dpo_ferramentas_digitais (company_id, chave, ano, dados, updated_by, updated_at) VALUES (?, 'cinco_s', ?, ?, ?, CURRENT_TIMESTAMP)
             ON CONFLICT(company_id, chave, ano) DO UPDATE SET dados = excluded.dados, updated_by = excluded.updated_by, updated_at = CURRENT_TIMESTAMP`,
            [f.companyId, ano, JSON.stringify(dados), req.user.userId], (err) => err ? reject(err) : resolve()));
        db.run(`INSERT INTO dpo_ferramentas_digitais_arquivos (company_id, chave, ano, tipo, url, original_name, comentario, created_by) VALUES (?, 'cinco_s', ?, 'planilha_5s', ?, ?, ?, ?)`,
            [f.companyId, ano, url, originalName || 'Planilha 5S', 'Importada para o sistema', req.user.userId], () => {});
        res.json({ message: `${lido.dados.areas.length} área(s), ${lido.dados.auditorias.length} auditoria(s) e ${lido.dados.acoes.filter(a => !a.auditoriaId).length} ação(ões) importadas. ${dados.acoes.filter(a => a.auto && !String(a.acao || '').trim()).length} plano(s) aberto(s) para os itens N.`, dados, avisos: lido.avisos });
    } catch (e) {
        console.error('Erro ao importar 5S:', e.message);
        res.status(400).json({ error: 'Não foi possível importar a planilha 5S.' });
    }
});
// Auditoria 5S pelo celular: link público por empresa/ano; respostas e fotos entram na ferramenta ao abrir.
async function link5sDpo(token) {
    const link = await dbGet(`SELECT * FROM dpo_5s_links WHERE token = ?`, [String(token || '')]);
    if (!link) return null;
    const empresa = await dbGet(`SELECT name FROM companies WHERE id = ?`, [link.company_id]);
    const dados = prepararDados5sDpo((await carregarFerramentaDigitalDpo(link.company_id, 'cinco_s', link.ano)).dados);
    return { link, empresa, dados };
}
app.post('/api/dpo/cinco-s/link', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const f = await resolverFerramentaDigitalDpo(req, res, 'cinco_s', req.body.company_id);
        if (!f) return;
        const ano = anoValidoDpo(req.body.ano);
        let reg = await dbGet(`SELECT token FROM dpo_5s_links WHERE company_id = ? AND ano = ?`, [f.companyId, ano]);
        if (!reg) { reg = { token: crypto.randomBytes(16).toString('hex') }; await new Promise((resolve, reject) => db.run(`INSERT INTO dpo_5s_links (token, company_id, ano, criado_por) VALUES (?, ?, ?, ?)`, [reg.token, f.companyId, ano, req.user.userId], e => e ? reject(e) : resolve())); }
        res.json({ url: `${baseUrlPublicaDpo(req)}/auditoria-5s.html?t=${reg.token}`, token: reg.token });
    } catch (e) { res.status(400).json({ error: 'Erro ao gerar o link da auditoria.' }); }
});
app.get('/api/dpo/cinco-s/celular', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const f = await resolverFerramentaDigitalDpo(req, res, 'cinco_s', req.query.company_id);
        if (!f) return;
        res.json(await dbAll(`SELECT area_id, mes, qid, resp, foto, autor, updated_at FROM dpo_5s_resp WHERE company_id = ? AND ano = ?`, [f.companyId, anoValidoDpo(req.query.ano)]));
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar a auditoria do celular.' }); }
});
app.get('/api/public/5s/:token', async (req, res) => {
    try {
        const r = await link5sDpo(req.params.token);
        if (!r) return res.status(404).json({ error: 'Link da auditoria inválido.' });
        const mes = Math.max(0, Math.min(11, Number(req.query.mes ?? (r.link.ano === new Date().getFullYear() ? new Date().getMonth() : 11)) || 0));
        const resp = {};
        r.dados.auditorias.filter(a => a.mes === mes).forEach(a => { resp[a.areaId] = {}; Object.entries(a.resp || {}).forEach(([q, v]) => { resp[a.areaId][q] = { resp: v, foto: (a.fotos || {})[q] || null, em: (a.respEm || {})[q] || '' }; }); });
        const cel = await dbAll(`SELECT area_id, qid, resp, foto, updated_at FROM dpo_5s_resp WHERE company_id = ? AND ano = ? AND mes = ?`, [r.link.company_id, r.link.ano, mes]);
        cel.forEach(c => { const a = resp[c.area_id] = resp[c.area_id] || {}, x = a[c.qid] = a[c.qid] || { resp: '', foto: null, em: '' }; if (c.resp !== null && String(c.updated_at) > String(x.em)) { x.resp = c.resp; x.em = c.updated_at; } if (c.foto) x.foto = c.foto; });
        res.json({ empresa: r.empresa ? r.empresa.name : '', ano: r.link.ano, mes, meta: r.dados.modelo.meta || 85,
            sensos: r.dados.modelo.sensos.map(s => ({ chave: s.chave, numero: s.numero, titulo: s.titulo, perguntas: s.perguntas.map(p => ({ id: p.id, num: p.num, texto: p.texto })) })),
            areas: r.dados.areas.map(a => ({ id: a.id, nome: a.nome, depto: a.depto || '', dono: a.dono || '', auditor: a.auditor || '', placa: a.placa || '', supervisor: a.supervisor || '' })), resp });
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar a auditoria.' }); }
});
function validar5sPublicoDpo(r, body) {
    const areaId = String(body.areaId || ''), qid = String(body.qid || ''), mes = Number(body.mes);
    if (!r.dados.areas.some(a => a.id === areaId)) return 'Área inválida.';
    if (!r.dados.modelo.sensos.some(s => s.perguntas.some(p => p.id === qid))) return 'Pergunta inválida.';
    if (!(mes >= 0 && mes <= 11)) return 'Mês inválido.';
    return null;
}
app.post('/api/public/5s/:token/resp', async (req, res) => {
    try {
        const r = await link5sDpo(req.params.token);
        if (!r) return res.status(404).json({ error: 'Link da auditoria inválido.' });
        const erro = validar5sPublicoDpo(r, req.body); if (erro) return res.status(400).json({ error: erro });
        const resp = String(req.body.resp ?? ''); if (!['S', 'N', 'NA', ''].includes(resp)) return res.status(400).json({ error: 'Resposta inválida.' });
        const quando = new Date().toISOString();
        await new Promise((resolve, reject) => db.run(`INSERT INTO dpo_5s_resp (company_id, ano, area_id, mes, qid, resp, autor, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(company_id, ano, area_id, mes, qid) DO UPDATE SET resp = excluded.resp, autor = excluded.autor, updated_at = excluded.updated_at`,
            [r.link.company_id, r.link.ano, req.body.areaId, Number(req.body.mes), req.body.qid, resp, String(req.body.autor || '').slice(0, 80), quando], e => e ? reject(e) : resolve()));
        res.json({ message: 'Resposta salva!', updated_at: quando });
    } catch (e) { res.status(400).json({ error: 'Erro ao salvar a resposta.' }); }
});
app.post('/api/public/5s/:token/foto', (req, res) => {
    uploadMaterialDpo.single('file')(req, res, async (err) => {
        if (err) return res.status(400).json({ error: err.code === 'LIMIT_FILE_SIZE' ? 'Foto muito grande.' : err.message });
        if (!req.file) return res.status(400).json({ error: 'Nenhuma foto recebida.' });
        if (!/^image\//.test(req.file.mimetype || '') && !/\.(jpe?g|png|heic|webp)$/i.test(req.file.originalname || '')) return res.status(400).json({ error: 'Envie uma imagem.' });
        try {
            const r = await link5sDpo(req.params.token);
            if (!r) return res.status(404).json({ error: 'Link da auditoria inválido.' });
            const erro = validar5sPublicoDpo(r, req.body); if (erro) return res.status(400).json({ error: erro });
            const url = '/uploads/' + req.file.filename;
            await new Promise((resolve, reject) => db.run(`INSERT INTO dpo_5s_resp (company_id, ano, area_id, mes, qid, resp, foto, autor, updated_at) VALUES (?, ?, ?, ?, ?, NULL, ?, ?, ?)
                ON CONFLICT(company_id, ano, area_id, mes, qid) DO UPDATE SET foto = excluded.foto`,
                [r.link.company_id, r.link.ano, req.body.areaId, Number(req.body.mes), req.body.qid, url, String(req.body.autor || '').slice(0, 80), new Date().toISOString()], e => e ? reject(e) : resolve()));
            res.json({ message: 'Foto enviada!', url });
        } catch (e) { res.status(400).json({ error: 'Erro ao salvar a foto.' }); }
    });
});
function exportar5sDpo(add, dados) {
    const d = prepararDados5sDpo(dados), M = MESES_CURTOS_DPO;
    const pctSenso = (au, s) => { let S = 0, N = 0; s.perguntas.forEach(p => { const v = resp5sDpo((au.resp || {})[p.id]); if (v === 'S') S++; if (v === 'N') N++; }); return S + N ? S / (S + N) : null; };
    const pctAud = au => { const v = d.modelo.sensos.map(s => pctSenso(au, s)).filter(x => x !== null); return v.length ? v.reduce((a, b) => a + b, 0) / v.length : null; };
    const fmt = v => v === null || v === undefined ? '' : `${String(Math.round(v * 1000) / 10).replace('.', ',')}%`;
    add('Donos de área', [['Área', 'nome', 26], ['Departamento', 'depto', 18], ['Placa', 'placa', 12], ['Dono / motorista', 'dono', 22], ['Supervisor', 'supervisor', 20], ['Auditor', 'auditor', 20]], d.areas);
    if (Array.isArray(d.motoristas) && d.motoristas.length) add('Motoristas', [['Motorista', 'nome', 26], ['Supervisor', 'supervisor', 22], ['Placa', 'placa', 12], ['Telefone', 'telefone', 16]], d.motoristas);
    add('Resultado Geral', [['Área', 'a', 26], ['Departamento', 'dep', 16], ...M.map((m, i) => [m, 'm' + i, 9])], d.areas.map(a => ({ a: a.nome, dep: a.depto, ...Object.fromEntries(M.map((m, i) => { const au = d.auditorias.find(x => x.areaId === a.id && x.mes === i); return ['m' + i, au ? fmt(pctAud(au)) : '']; })) })));
    const linhas = [];
    d.auditorias.forEach(au => { const a = d.areas.find(x => x.id === au.areaId); d.modelo.sensos.forEach(s => s.perguntas.forEach(p => linhas.push({ area: a ? a.nome : '', mes: M[au.mes], auditor: au.auditor, senso: s.titulo, q: `${p.num} ${p.texto}`, r: (au.resp || {})[p.id] || '' }))); });
    add('Base', [['Área', 'area', 22], ['Mês', 'mes', 8], ['Auditor', 'auditor', 18], ['Senso', 'senso', 14], ['Pergunta', 'q', 60], ['Resposta', 'r', 10]], linhas);
    add('Gerenciador de Ações', [['Mês', 'mes', 8], ['Área', 'area', 22], ['Senso', 'senso', 16], ['Item', 'item', 50], ['Ação', 'acao', 45], ['Dono', 'dono', 18], ['Prevista', 'prevista', 12], ['Tratativa', 'tratativa', 40], ['Status', 'status', 14]],
        d.acoes.map(a => { const ar = d.areas.find(x => x.id === a.areaId), s = d.modelo.sensos.find(x => x.chave === a.senso), p = s && a.qid ? s.perguntas.find(x => x.id === a.qid) : null; return { ...a, mes: a.mes === null || a.mes === undefined ? '' : M[a.mes], area: ar ? ar.nome : '', senso: s ? s.titulo : a.senso, item: p ? `${p.num} ${p.texto}` : '' }; }));
}

// ======================================================================
// DPO — FERRAMENTAS EXCLUSIVAS POR PERGUNTA (substituem o acompanhamento genérico)
//   sonho      -> Gestão 1.2        (DPO Day — Sonho: frase, KPIs meta x real)
//   manutencao -> Planejamento 2.2  (Check Global de Manutenção + ronda com fotos)
//   p3a        -> Planejamento 2.3  (P1A / P3A: pessoas, estrutura e frota)
//   capex      -> Planejamento 2.4  (Gestão de CAPEX com NF na finalização)
//   ans        -> Planejamento 3.2  (ANS de Volume com vendas, preenchimento diário)
// A nota sugerida é calculada na tela e gravada em dados.__resumo.
// ======================================================================
const MODELO_MANUTENCAO_DPO = [{"titulo": "Fundamentos", "grupos": [{"numero": "1", "titulo": "Gestão De Áreas E Equipamentos Críticos", "itens": [{"id": "m1_1", "num": "1.1", "texto": "A Estrutura das Coberturas existentes apresentam bom estado de conservação?", "verificacao": "Estrutura com ausência de anomalias (ferrugem, colisão, danos em geral) que coloquem em risco a segurança e operacionalidade do Armazém.\n\nLaudo Técnico Estrutural/Mapeamento e Tratativa das Anomalias.\n\nCheck do cronograma de manutenção de 6 em 6 meses.", "pontos": "3 - Estruturas sem risco de queda e cronograma de manutenção em dia.\n\n1- Algumas anomalias encontradas, porém todas mapeadas com plano de ação gerando visibilidade com follow.\n\n0 - Plano de ação inconsistente e/ou anomalias não mapeadas.", "peso": 4, "foto": true, "critico": true}, {"id": "m1_2", "num": "1.2", "texto": "As telhas e calhas das coberturas existentes estão em bom estado de conservação?", "verificacao": "Checar se as telhas e calhas estão livres de danos, vazamento e goteiras.\n\nVerificar se existe vazão apropriada para as águas oriundas do telhado.\n\nCheck do cronograma de manutenção de 6 em 6 meses.", "pontos": "3 - As telhas e calhas das coberturas existentes estão em bom estado de conservação.\n\n1- Algumas anomalias encontradas, porém todas mapeadas com plano de ação gerando visibilidade com follow.\n\n0 - Plano de ação inconsistente e/ou anomalias não mapeadas.", "peso": 1, "foto": true, "critico": true}, {"id": "m1_3", "num": "1.3", "texto": "Itens críticos de segurança patrimonial como: portão de acesso, porta do banco, torniquete e CFTV estão em perfeita funcionalidade?", "verificacao": "Checar (in loco) a funcionalidade dos itens e entrevistar usuários (controle, portaria e financeiro).\n\nCFTV com 60 dias de imagens e contrato de manutenção de segurança (controle de acesso e CFTV).\n\nPara todos os itens funcionarem perfeitamente é necessário um plano de manutenção.", "pontos": "3 - Todos os itens críticos de segurança patrimonial estão funcionando perfeitamente.\n\n1 - Pelo menos 2 itens estão funcionando perfeitamente e os outros estão mapeados para solução. \n\n0 – Menos de 2 itens funcionando e/ou problemas não mapeados.", "peso": 1, "foto": true, "critico": true}, {"id": "m1_4", "num": "1.4", "texto": "Áreas e equipamentos críticos relacionados à qualidade (câmara fria, gerador, flowracks, racks, equipamentos de limpeza, áreas e equipamentos de controle de PNC) são inspecionados regularmente e estão em perfeita funcionalidade?", "verificacao": "Checar (in loco) a funcionalidade e conservação dos itens e áreas e entrevistar usuários.\n\nChecar se todos os itens estão inventariados em perfeita condição e se tem plano de correção para os que não estejam.\n\nVerificar a existência de laudo que ateste a segurança dos racks e flowracks.\n\nVerificar o cronograma de manutenção preventiva da câmara fria e gerador.", "pontos": "3 - Todos os itens críticos de qualidade estão funcionando perfeitamente.\n\n1 - Pelo menos 3 itens estão funcionando perfeitamente e os outros estão mapeados para solução.\n\n0 - Dois itens apenas funcionando e/ou problemas não mapeados.", "peso": 1, "foto": true, "critico": true}, {"id": "m1_5", "num": "1.5", "texto": "Itens e equipamentos críticos de segurança (hidrantes e extintores, trava rodas, linha de vida, paleteira manual e carrinhos) estão em perfeita funcionalidade?", "verificacao": "Checar (in loco) a funcionalidade e conservação dos itens e entrevistar usuários.\n\nChecar se todos os itens estão inventariados em perfeitas condições e se tem plano de correção para os que não estejam.\n\nVerificar a existência de laudo que ateste a segurança da linha de vida.\n\nVerificar cronograma/plano de inspeção dos hidrantes, extintores, paleteiras e carrinhos.", "pontos": "3 - Todos os itens críticos de segurança estão funcionando perfeitamente.\n\n1 - Pelo menos 3 itens estão funcionando perfeitamente e os outros estão mapeados para solução.\n\n0 - Dois itens apenas funcionando e/ou problemas não mapeados.", "peso": 1, "foto": true, "critico": true}, {"id": "m1_6", "num": "1.6", "texto": "A unidade possui um plano de contingência único, integrado e atualizado para casos críticos?", "verificacao": "Apresentar contratos/tratativa para abastecimento de água, abastecimento dos geradores, falta de energia e intertravamento das portas (caixa).\n\nVerificar se o plano está atualizado com o contato dos responsáveis atuais.\n\nVerificar plano para compra, transporte e abastecimento de diesel do gerador.", "pontos": "3 - Possui plano de contingência único, integrado e atualizado de 100% dos itens críticos.\n\n1 - Há um plano, mas este não está completo ou não está atualizado.\n\n0 - Não há um plano.", "peso": 3, "foto": false, "critico": true}, {"id": "m1_7", "num": "1.7", "texto": "A unidade executa a rotina estabelecida (rondas e reuniões), criando plano de ação correto e consistente?", "verificacao": "Verificar se o Quadro de Gestão à Vista da Matinal do GOD está atualizado com indicador da área.\n\nVerificar se a Reunião de Estrutura acontece na frequência correta e conforme TOR.\n\nVerificar se a Ronda de Blindagem é feita na periodicidade e qualidade correta.\n\nVerificar se a Reunião do Pilar acontece conforme TOR da Reunião DPO.\n\nVerificar se a Super Matinal/Vespertina/Noturna aborda o assunto conforme TOR.\n\nVerificar se o assunto é tratado em MPR do GOD e da GEO.\n\nAnalisar os planos de ação para as anomalias e serviços já realizados e a serem executados e/ou planejados.", "pontos": "3 - Reuniões e rondas são realizadas regularmente conforme padrão e há plano de ação para a evolução da estrutura. \n\n1 - Reuniões e rondas são realizadas parcialmente (pelo menos 75%, com a frequência especificada) –  analisar os últimos 3 meses.\n\n0 - Reuniões e rondas não são realizadas regularmente, de acordo com as orientações, ou estão com frequência abaixo de 75%  –  analisar os últimos 3 meses.\n\n\nNOTA: Se a Reunião de Estrutura não estiver sendo realizada na frequência correta, a questão 1.7 deverá ser Zero.", "peso": 3, "foto": false, "critico": true}]}, {"numero": "2", "titulo": "Gestão Do Plano De Tráfego", "itens": [{"id": "m2_1", "num": "2.1", "texto": "As cancelas e guarda-corpos de proteção estão em bom estado de conservação e de acordo com a especificação padrão?", "verificacao": "Verificar se as segregações, guarda-corpos e cancelas estão conforme padrão e sem apresentar anomalias.\n\nVerificar se existe registro de quando aconteceu à anomalia e qual foi à tratativa e/ou fluxo de cobrança da avaria.\n\nVerificar se os guarda-corpos e proteções de pilares estão devidamente fixados ao piso com todos os parafusos bem apertados.", "pontos": "3 - Segregações, guarda-corpos e cancelas em boas condições de utilização.\n\n1 - Conservação com algumas falhas na pintura, pequenas manutenções ou mal afixados no piso, mas com plano de ação para adequação.\n\n0 - Conservação em estado ruim sem plano de ação.", "peso": 3, "foto": true, "critico": false}, {"id": "m2_2", "num": "2.2", "texto": "As áreas de segregação do picking, espera dos motoristas, refugo, sala dos conferentes, retorno de rota, pit stop e armazenamento de gás (P20) estão em bom estado de conservação?", "verificacao": "Verificar em ronda as condições das áreas que devem estar em perfeita condições de uso. Caso haja anomalia deve haver registro das ocorrências com plano de ação para tratamento.\n\nO prazo entre registro de anomalia e tratamento deve respeitar padrão.", "pontos": "3 - As áreas estão em bom estado de conservação. Problemas sendo tratado em plano de ação.\n\n1 - Conservação com algumas falhas na pintura ou pequenas manutenções, mas sendo tratado via plano de ação.\n\n0 - Conservação em estado ruim.", "peso": 3, "foto": true, "critico": false}, {"id": "m2_3", "num": "2.3", "texto": "Pinturas de faixa de pedestre, fluxo de circulação, separação de lotes, vagas de veículos (leves e pesados), guia de conferente, redzone e sinalizações dos equipamentos de combate a incêndio estão em bom estado de conservação?", "verificacao": "Unidade deve ter cronograma para pinturas novas e manutenção das antigas. Controle eletrônico com plano de ação para tratamento dos problemas quando necessário.\n\nEm ronda na unidade verifique o estado de conservação das pinturas. As mesmas devem seguir padrão.\n\nPlacas de sinalização em bom estado de conservação.", "pontos": "3 - Pinturas e sinalizações presentes em todos os locais definidos e estão em bom estado de conservação.\n\n1 - Pinturas e sinalizações presentes em todos os locais definidos, contudo existem falhas na conservação sendo tratados com plano de ação. \n\n0 - Falta pintura e/ou sinalizações em locais definidos e/ou má conservação nas existentes.", "peso": 1, "foto": true, "critico": false}]}, {"numero": "3", "titulo": "Conservação Civil", "itens": [{"id": "m3_1", "num": "3.1", "texto": "As áreas de estocagem, circulação de veículos, Pit Stop, pedestres e áreas ADM estão livres de buracos ou outras interferências?", "verificacao": "Verificar (in loco) pisos, rampas e áreas de circulação da unidade, analisar VBZ de quebra e checar chamados com motivos buracos. \n\nCheck de apontamento de condições inseguras no Credit sem tratamento.", "pontos": "3 - Pisos em bom estado de conservação, não apresentando risco de quedas ou tropeços.\n\n1 - Algumas anomalias encontradas, porém todas mapeadas com plano de ação gerando visibilidade com follow.\n\n0 - Plano de ação inconsistente e/ou anomalias não mapeadas.", "peso": 1, "foto": true, "critico": false}, {"id": "m3_2", "num": "3.2", "texto": "Tetos e forros estão livres de rachaduras, buracos, deslocamentos (PVC ou Placas), umidade, manchas e goteiras?", "verificacao": "Verificar in loco.\n\nCheck de chamados abertos com plano de ação para solucionar.\n\nCheck de GSAs com plano de ação para solucionar.", "pontos": "3 - Tetos e forros em bom estado de conservação, não apresentando risco de queda.\n\n1 - Alguma anomalia encontrada, porém todas mapeadas com plano de ação gerando visibilidade com follow.\n\n0 – Anomalias identificadas em ronda e não mapeadas e/ou plano de ação inconsistente.", "peso": 1, "foto": true, "critico": false}, {"id": "m3_3", "num": "3.3", "texto": "O Muro de Fechamento do perímetro está livre de buracos, rachaduras e com pintura em bom estado (quando houver pintura)?", "verificacao": "Verificar se existem anomalias (buracos, concertinas amassadas, respeitar especificações do check list patrimonial dos muros). \n\nCheck de chamados abertos com plano de ação para solucionar.", "pontos": "3 - Muro em bom estado de conservação, não apresentando trincas ou risco de queda.\n\n1 - Alguma anomalia encontrada, porém todas mapeadas com plano de ação gerando visibilidade com follow.\n\n0 - Anomalias identificadas em ronda e não mapeadas e/ou plano de ação inconsistente.", "peso": 3, "foto": true, "critico": false}, {"id": "m3_4", "num": "3.4", "texto": "O revestimento das paredes das salas e acessos estão livres de buracos e rachaduras?", "verificacao": "Verificar se existem anomalias (pintura envelhecida ou danificada, mofos, buracos, trincas, rachaduras) nas paredes.\n\nCheck de chamados abertos com plano de ação para solucionar.", "pontos": "3 - Paredes em bom estado de conservação, não apresentando trincas ou risco de quedas. \n\n1 - Alguma anomalia encontrada, porém todas mapeadas com plano de ação gerando visibilidade com follow.\n\n0 - Anomalias identificadas em ronda e não mapeadas e/ou plano de ação inconsistente.", "peso": 3, "foto": true, "critico": false}, {"id": "m3_5", "num": "3.5", "texto": "As portas e janelas das salas estão em perfeitas condições de uso?", "verificacao": "Verificar se existem anomalias (maçaneta quebrada, trinco quebrado, sem porta, vidro quebrado das janelas, persiana rasgada entrando sol, insulfim rasgado, porta emperrando ou fazendo barulho ao movimentar). \n\nCheck de chamados abertos com plano de ação para solucionar.", "pontos": "3 - Portas e janelas em bom estado de conservação.\n\n1 - Alguma anomalia encontrada, porém todas mapeadas com plano de ação gerando visibilidade com follow.\n\n0 - Anomalias identificadas em ronda e não mapeadas e/ou plano de ação inconsistente.", "peso": 3, "foto": false, "critico": false}]}, {"numero": "4", "titulo": "Elétrica", "itens": [{"id": "m4_1", "num": "4.1", "texto": "A Cabine Primária, Quadros de Energia e Infraestrutura Elétrica para distribuição de Força da Unidade estão em bom estado de conservação?", "verificacao": "Unidade possui laudo técnico válido com as condições infraestruturais elétrica da unidade? Incluem verificação das condições das cabines e quadros? Verificar documentação.\n\nA infraestrutura elétrica esta em boas condições de uso? No caso de anomalia existe plano de ação para tratamento com prazos coerentes?\n\nA infraestrutura elétrica tem revisão periódica conforme orientação de engenheiro eletricista?\n\nOs locais onde estão instalados as cabines e quadros são adequados (a frente de qualquer quadro / painel elétrico deve ter 1m² de acesso livre de interferências para a execução da manutenção do mesmo)? Existe risco de batida por máquinas e veículos? Verifique as condições físicas em ronda.\n\nToda a rede elétrica está protegida por eletrocalha, eletrodutos, canaletas, conduítes e suportes.", "pontos": "3 - Unidade com toda documentação válida. Instalações em perfeitas condições de uso e revisões periódicas acontecendo conforme orientação técnica garantindo a segurança da unidade. \n\n1 - Unidade com toda documentação válida. Instalação com algumas anomalias sendo tratadas via plano de ação e revisões acontecendo conforme orientação técnica.\n\n0 - Unidade sem documentação válida. Instalações com problemas ou sem revisões.", "peso": 1, "foto": true, "critico": true}, {"id": "m4_2", "num": "4.2", "texto": "Iluminação interna e externa está em perfeitas condições de uso? Existe cronograma de verificação dos mesmos?", "verificacao": "Realize ronda pelas áreas e verifique a existência de lâmpadas, refletores, etc. queimadas e/ou danificados.\n\nA unidade possui agenda definida para verificação das condições da iluminação? O mesmo atende a necessidade da unidade?\n\nVerifique RACIs e ANSs definidas para realização de rondas e tratamento de problemas. A unidade deve ter análise e plano de ação para tratamento de problemas.", "pontos": "3 - Poucas lâmpadas queimadas ou danificadas, unidade com agenda definida para verificação e ANSs e RACIs definidas.\n\n1 – Algumas lâmpadas queimadas, falha no cronograma de verificação e sem plano de ação consistente. \n\n0 - Lâmpadas queimadas, falha no cronograma verificação e/ou plano de ação inconsistente.", "peso": 1, "foto": true, "critico": false}, {"id": "m4_3", "num": "4.3", "texto": "As instalações elétricas como: ar condicionado, ventiladores, chuveiros elétricos, tomadas e VDs estão em bom estado de conservação? Existe cronograma de verificação?", "verificacao": "Realize ronda pelas áreas e verifique a existência de equipamentos queimados e/ou danificados.\n\nVerifique o cronograma de revisão das instalações elétricas. A unidade deve ter controle formal das revisões.\n\nVerificar se o ar condicionado está inventariado, possui plano de manutenção preventiva e se as anomalias estão mapeadas.", "pontos": "3 – Pelo menos 3 dos itens de instalações elétricas estão em boas condições de uso e cronograma de revisões ocorrendo e sendo registradas.\n\n1 – Menos de 3 itens de instalações elétricas estão em boas condições, contudo não existem falhas no cronograma de revisões.\n\n0 - Instalações elétricas com problemas e/ou falha no cronograma de revisões.", "peso": 3, "foto": true, "critico": false}]}, {"numero": "5", "titulo": "Hidráulica, Áreas Molhadas E Molháveis", "itens": [{"id": "m5_1", "num": "5.1", "texto": "Os Banheiros, Vestiários e Refeitórios estão garantindo as condições básicas de funcionamento e utilização?", "verificacao": "Banheiros, vestiários e refeitórios livres de odores.\n\nVasos, Pias e Torneiras em funcionamento e livre de vazamentos.\n\nChuveiros em funcionamento.\n\nArmários em bom estado de conservação.\n\nSuportes para sabão, papéis e espelhos bem fixados e sem anomalias.", "pontos": "3 - Atendimento de 5 itens de verificação.\n\n1- Atendimento de 3 a 4 itens de verificação.\n\n0 - Atendimento menor que 3 itens de verificação.", "peso": 1, "foto": true, "critico": false}, {"id": "m5_2", "num": "5.2", "texto": "A unidade está realizando a limpeza e possui registro de manutenção preventiva nos ralos, sifões, grelhas e galerias (condutor de água pluvial)?", "verificacao": "Verificar se a unidade executa um cronograma padrão de limpeza  e manutenção e se existem ações para tratamento de anomalias. \n\nEntreviste algumas pessoas e verifique se existe registro de anomalias relacionadas a entupimentos, inundações e vazamentos.", "pontos": "3 - Cronograma de limpeza e manutenção acontecendo conforme padrão. O mesmo é gerenciado via plano de ação. Não existem sinas de entupimentos, inundações e vazamentos.\n\n1 - Não existem sinais de entupimentos ou vazamentos, porém cronograma de limpeza com falhas e sem plano de ação para tratamento.\n\n0 - Existência de sinais de vazamentos e entupimentos e sem cronograma de limpeza e manutenção.", "peso": 3, "foto": false, "critico": false}, {"id": "m5_3", "num": "5.3", "texto": "Reservatório de água potável da unidade está em perfeita condição de uso e possui registro de manutenção e limpezas periódicas? (Se aplicável)", "verificacao": "Existe plano para abastecimento de água no caso de falta? Existe laudo para controle de PH?\n\nExiste gestão das manutenções dos reservatórios? A mesma está em boas condições físicas?\n\nUnidade realiza manutenção e limpeza conforme padrões de conservação sanitários? Existe controle eletrônico (planilha)? Atenção para unidades com poço artesiano. \n\nSmall OP:  Regional  ajudar na elaboração / execução do plano.", "pontos": "3 - Plano de abastecimento consistente, laudo emitido e controlado, livre de vazamentos e limpeza ocorrendo conforme cronograma e padrões sanitários.\n1 – Laudo emitido, livre de vazamentos, mas com falhas no cronograma de limpeza e manutenção.  \n0 – Não atende os requisitos acima.", "peso": 3, "foto": false, "critico": true}]}]}, {"titulo": "Gerenciar para Manter", "grupos": [{"numero": "6", "titulo": "Manutenção Preventiva", "itens": [{"id": "m6_1", "num": "6.1", "texto": "A unidade possui um Plano de manutenção preventiva (com periodicidade e atividades) para cada tipo de equipamentos e áreas criticas?", "verificacao": "Plano de manutenção com periodicidade das seguintes atividades preventivas:\n\nEquipamentos críticos: geradores, portões de acesso, ar condicionado, quadros elétricos, flowracks, porta pallet, linha de vida, câmara fria, SPDA, bomba d'água e recalque (alimentação dos hidrantes).\n\nÁreas criticas: Caixa Financeiro, Pit Stop,  tanque de abastecimento, oficinas, telhados.\n\nVerificar o controle de produtividade do técnico, anomalias, controle de fotos e OS parada por falta de material. \n\nServiços a serem realizados em cada tipo de manutenção (com pelo menos as recomendações contidas no manual do equipamento - quando aplicável).", "pontos": "3 - O plano de manutenção é seguido e engloba todas as áreas e equipamentos críticos, com gestão da produtividade dos técnicos, anomalias e materiais. \n\n1 - O plano de manutenção existe, mas não é detalhado ou está faltando componentes-chaves, gestão insuficiente da produtividade, anomalias e materiais. \n\n0 - Não existem evidências de plano de manutenção.", "peso": 3, "foto": false, "critico": false}, {"id": "m6_2", "num": "6.2", "texto": "A unidade utiliza os aprendizados das manutenções corretivas para atualizar os planos de manutenção preventiva?", "verificacao": "Verificar a Lista de Manutenção Corretivas com as principais ocorrências.\n\nVerificar a realização de Relatos de Anomalia para itens críticos que passaram por manutenção corretiva com uma preventiva feita. \n\nVerificar se a unidade controla os indicadores MTBF e MTTR e se existem ações para melhorar a performance.\n\nVerificar evolução dos indicadores atrelados a atualização dos planos.\n\nSmall OP:  Regional  apurar o KPI  de forma centralizada.", "pontos": "3 - Os planos de Manuteção Preventiva são atualizados com os aprendizados das manutenção corretivas, relatos de anomalia são gerados quando necessário e a unidade controla o MTBF e MTTR.\n\n1 - Os planos de Manuteção Preventiva são atualizados com os aprendizados da Manutenção corretiva, mas não apresentam evolução, existem falhas na geração dos relatos e/ou no acompanhamento de MTBF e MTTR.\n\n0 - Os planos não foram atualizados, não são feitos relatos de anomalia e/ou o MTBF e MTTR não é controlado.", "peso": 3, "foto": false, "critico": false}]}, {"numero": "7", "titulo": "Gestão Dos Custos De Manutenção", "itens": [{"id": "m7_1", "num": "7.1", "texto": "As manutenções recorrentes executadas na unidade possuem contrato de prestação de serviço validado via CSU?", "verificacao": "Verificar os contratos (Ex.: Manserv, Talentos, etc.).\n\nVerificar exceções aprovados pelo corporativo AC.\n\nVerificar Matriz de Serviços disponibilizada pelo Coorporativo.\n\nVerificar execução de manutenção sem pedido ou contrato.\n\nVerificar existência de regularização de notas fiscais.", "pontos": "3 - Possui todos os contratos de serviços recorrentes devidamente aprovados.\n\n1 - Não possui todos os contratos, no entanto estão em processo de aprovação.\n\n0 - Não possui contrato, não estão em fluxo de aprovação e/ou possui notas fiscais não regularizadas.", "peso": 3, "foto": false, "critico": false}, {"id": "m7_2", "num": "7.2", "texto": "Existe um controle e estratificação dos maiores gastos por área/equipamento/serviço?", "verificacao": "Apresentar controle com histórico mínimo de 6 meses.\n\nVerificar a realização da Reunião de OBZ com análises e estratificações. \n\nPlano de Ação com follow dos maiores gastos.", "pontos": "3 - Possui estratificação aberto por área, Reunião de OBZ acontece com Plano de Ação e follow para os itens de maior impacto.\n\n1 - Possui estratificação, mas com plano de ação e follow inconsistentes.\n\n0 - Não possui estratificação.", "peso": 3, "foto": false, "critico": false}, {"id": "m7_3", "num": "7.3", "texto": "A unidade possui uma gestão do pacote de manutenção?", "verificacao": "Verificar se o dono tem o acompanhamento do resultado (PLAN x TEND x REAL).\n\nVerificar se a unidade tem estouros no Pacote Manutenção.\n\nVerificar se o dono pode explicar os principais impactos.\n\nVerificar se a unidade tem um controle da tendência do LE.\n\nVerificar se existem alocações indevidas no pacote sem tratativa.", "pontos": "3 – A unidade possui gestão e acompanhamento do pacote, garante a correta alocação das despesas, sem estouro, possui evidências e estratificações com Plano de Ação e follow.\n\n1 – A unidade possui gestão e acompanhamento do pacote, no entanto, existem falhas de alocação das despesas e nas tratativas. \n\n0 – A unidade não possui gestão do pacote e/ou possui lançamentos indevidos.", "peso": 3, "foto": false, "critico": false}, {"id": "m7_4", "num": "7.4", "texto": "A unidade possui áreas internas comodatadas para parceiros? A mesma possui evidências de cobranças?", "verificacao": "Verificar existência do contrato das áreas comodatadas, cobrar existência física do comodato assinado e reconhecimento de firma. \n\nVerificar se as áreas estão em bom estado de conservação e/ou foram realizados os reparos necessários.\n\nVerificar se os reparos de responsabilidade do comodatado foram devidamente cobrados.", "pontos": "3 - Possui comodato e as obras com responsabilidade do parceiro é devidamente cobrado.\n\n1 - Possui comodato, mas não é cobrado.\n\n0 - Não possui comodato.", "peso": 3, "foto": false, "critico": false}, {"id": "m7_5", "num": "7.5", "texto": "A unidade possui um processo de aquisição de equipamentos e peças para a execução das atividades de manutenção?", "verificacao": "Verificar se os funcionários conhecem e utilizam o Portal do Fornecedor Local.\n\nApresentar controle que evidencie a contratação de serviços e/ou aquisição de peças x lista de ordens de serviço. \n\nVerificar se existem OS abertas por falta de material.\n\nAvaliar o prazo de cumprimento das ordens de serviço x disponibilidade dos materiais.\n\nVerificar se a unidade faz gestão de estoque de peças e materiais com inventários regulares.", "pontos": "3 - 90% das aquisições concretizadas <= de 30 dias da abertura da ordem de serviço.\n\n1 - 90% das aquisições concretizadas <=60 dias da abertura da ordem de serviço.\n\n0 - Não atende os requisitos.", "peso": 3, "foto": false, "critico": false}, {"id": "m7_6", "num": "7.6", "texto": "A unidade possui um processo definido para o planejamento orçamentário de obras, serviços e aquisição de peças?", "verificacao": "Verificar a existência de orçamentos padronizados que contemplem todos os itens.\n\nEntrevistar se o dono entende os benefícios de se realizar um orçamento padronizado.\n\nValidar se o serviço descriminado corresponde ao orçado.\n\nVerificar se o Aceite Final da Obra reflete o orçamento aprovado. \n\nVerificar se foi prospectado mais de um fornecedor.", "pontos": "3 – Obras e serviços realizados atenderam os requisitos de cotação e prospecção orçamentária.\n\n1 – A unidade possui processo definido, no entanto existem falhas e oportunidades. \n\n0 – Não atende os requisitos.", "peso": 3, "foto": false, "critico": false}]}, {"numero": "8", "titulo": "Gestão De Ordens De Serviços", "itens": [{"id": "m8_1", "num": "8.1", "texto": "Existe um fluxo definido e amplamente divulgado da ferramenta de abertura de chamados?", "verificacao": "Verificar se existe o fluxo de abertura, disponibilidade da ferramenta e plano de comunicação (visão 6 meses).\n\nEntrevistar in loco 3 usuários para checar o conhecimento da ferramenta.\n\nVerificar a gestão e conservação dos QRCodes.\n\nLUP com passo a passo para abertura de chamado.\n\nNota: Small OP poderá realizar o processo através do Clic.", "pontos": "3 - Evidências da disponibilidade, comunicação e utilização da ferramenta.\n\n1 - A ferramenta existe, mas não é utilizada ou bem comunicada.\n\n0 - Sem evidências da utilização e comunicação.", "peso": 3, "foto": false, "critico": false}, {"id": "m8_2", "num": "8.2", "texto": "A unidade garante gestão das ordens de serviços (corretivas e preventivas) com prazo de execução, priorização das demandas e follow nas reuniões de rotina?", "verificacao": "Verificar se existe algum sistema de gestão implatando (Exppe - Optimus) e este é utilizado frequentemente. \n\nVerificar follow na reunião de estrutura semanal.\n\nVerificar se existem chamados fechados indevidamente, sem a solução definitiva do problema. \n\nVerificar Plano de Ação para as anomalias e serviços não atendidos.", "pontos": "3 - Há sistema de gestão e acompanhamento, sem chamados fechados indevidamente, anomalias são tratadas em reunião com Plano de Ação e follow.\n\n1 – Há sistema de gestão e acompanhamento, com chamados fechados indevidamente, falhas na tratativa das anomalias.\n\n0 - Não possui sistema de gestão e anomalias não são tratadas.", "peso": 3, "foto": false, "critico": false}, {"id": "m8_3", "num": "8.3", "texto": "A unidade realiza check com os prestadores de serviço com foco em planejamento, execução e nível de serviço?", "verificacao": "Ata de reunião com o desdobramento de atividades, acompanhamento da execução, qualidade dos serviços. \n\nVerificar se a unidade exige e faz a gestão da garantia de peças, materiais e serviços.", "pontos": "3 – Tem gestão da garantia de peças e serviços e evidências de reunião com Plano de Ação e follow. \n\n1 – Algumas peças e serviços possuem garantias, mas existem falhas no Plano de Ação e reuniões.\n\n0 - Não possui evidências de gestão de garantia e não ocorre reunião.", "peso": 1, "foto": false, "critico": false}]}]}, {"titulo": "Gerenciar para Melhorar", "grupos": [{"numero": "9", "titulo": "Nível De Serviço", "itens": [{"id": "m9_1", "num": "9.1", "texto": "A unidade aplica e tem acompanhamento da Pesquisa de Nível de Serviço de manutenção e Serviços Gerais?", "verificacao": "Verificar se a unidade aplica periodicamente uma Pesquisa de Nível de Serviço.\n\nVerificar se existe evolução entre uma pesquisa e outra.\n\nChecar a existência de plano de ação e evidências que respaldem e enderecem os itens da pesquisa com follow mensal.", "pontos": "3 – A Pesquisa é aplicada, existe plano consistente com follow mensal e apresenta evolução no resultado.\n\n1 – A Pesquisa é aplicada, existe um plano sem follow e sem evolução no resultado.\n\n0 – A pesquisa não é aplicada e/ou não existe acompanhamento.", "peso": 1, "foto": false, "critico": false}, {"id": "m9_2", "num": "9.2", "texto": "A unidade garante os chamados de manutenção predial fechados no prazo?", "verificacao": "Apresentar gestão e estratificação dos chamados.\n\nChecar se o resultado da unidade é maior ou igual à meta desdobrada. \n\nVerificar % de chamados reabertos. \n\nEntrevista com no mínimo 3 usuários e verificação in loco.", "pontos": "3 – A unidade atinge a meta de chamados, possui acompanhamentos gerenciais e controla o % de chamados reabertos. \n\n1 – A unidade atinge a meta de chamados, mas os controles não são eficientes e não há informações sobre os chamados reabertos. \n\n0 - Não atende os requisitos.", "peso": 3, "foto": false, "critico": false}, {"id": "m9_3", "num": "9.3", "texto": "A unidade possui um plano efetivo, amplo e frequente de comunicação dos processos de manutenção, obras e serviços concluídos e feedback de chamados?", "verificacao": "Verificar se a unidade possui uma rotina de comunicar as frentes da área. \n\nChecar se existe plano de comunicação para iniciar atividades de manutenção e obras  constando: macro atividades, cronograma, plano de tráfego e áreas a serem isoladas, sendo obrigatória a presença dos responsáveis abaixo:\nAmbev - Gerente da área, Prefeito  5S, TST e Tec. de Manutenção;\nTerceiros - Responsável pela área e TST;\nConstrutora - Encarregado e TST.\n\nConsultar materiais de RCOG, R. Estrutura, Super Matinal. \n\nChecar a ativação e comunicação via WorkPlace e Comunicação Interna. \n\nEntrevista in loco.", "pontos": "3 – Comunicação eficiente, frequente e ampla, utilizando as reuniões corretas e promovendo mudança na percepção da área. \n\n1 - Comunicação inconsistente e Plano de Ação com falhas. \n\n0 - Não existe comunicação.", "peso": 3, "foto": false, "critico": false}, {"id": "m9_4", "num": "9.4", "texto": "Foi feito algum benckmark de processo, indicadores, melhores práticas ou iniciativas com outras operações?", "verificacao": "Verificar o processo de busca e compartilhamento de Melhores Práticas de Manutenção entre as unidades.  \n\nVerificar evolução direta ou indireta no processo que foi aplicado a Melhor Prática.\n\nVerificar planos de ação.", "pontos": "3 – A unidade adotou/compartilhou alguma melhor prática e consegue evidenciar melhoria nos processos. \n\n1 – A unidade adotou/compartilhou alguma melhor prática, mas ainda não houve melhoria nos processos.  \n\n0 - Não existem evidências.", "peso": 1, "foto": false, "critico": false}]}]}];
const TOOLS_EXCLUSIVAS_DPO = { visibilidade: 'entrega:3.1', riscos: 'planejamento:2.1', sonho: 'gestao:1.2', manutencao: 'planejamento:2.2', p3a: 'planejamento:2.3', capex: 'planejamento:2.4', ans: 'planejamento:3.2' };
Object.assign(FERRAMENTAS_DIGITAIS_DPO, {
    visibilidade: { pilar: 'entrega', pergunta: '3.1', titulo: 'Visibilidade de Resultados da Distribuição' },
    riscos: { pilar: 'planejamento', pergunta: '2.1', titulo: 'Riscos, Resposta e Retomada de Negócios' },
    sonho: { pilar: 'gestao', pergunta: '1.2', titulo: 'DPO Sonho' },
    manutencao: { pilar: 'planejamento', pergunta: '2.2', titulo: 'Check Global de Manutenção' },
    p3a: { pilar: 'planejamento', pergunta: '2.3', titulo: 'P1A / P3A — pessoas, estrutura e frota' },
    capex: { pilar: 'planejamento', pergunta: '2.4', titulo: 'Gestão de CAPEX' },
    ans: { pilar: 'planejamento', pergunta: '3.2', titulo: 'ANS de Volume com Vendas' }
});

function exportarExclusivaDpo(chave, add, dados) {
    const M = MESES_CURTOS_DPO, d = dados || {};
    const tabela = (nome, lista, cols) => { if (Array.isArray(lista) && lista.length) add(nome, cols.map(([h, k, w]) => [h, k, w || 18]), lista); };
    if (chave === 'visibilidade') {
        const inds = (d.indicadores || []).filter(i => i.nome);
        tabela('Motoristas e ajudantes', d.colaboradores, [['Matrícula', 'matricula', 12], ['Nome', 'nome', 30], ['Função', 'funcao', 14], ['Supervisor', 'supervisor', 22], ['Placa', 'placa', 12], ['Telefone', 'telefone', 16], ['Ativo', 'ativo', 8]]);
        add('Metas', [['Indicador', 'nome', 26], ['Unidade', 'unidade', 10], ['Sentido', 'sentido', 14], ['Meta', 'meta', 10], ['Prêmio por dia na meta (R$)', 'premio', 16]], inds.map(i => ({ ...i, sentido: i.sentido === 'menor' ? 'menor melhor' : 'maior melhor' })));
        const lin = [], nomes = Object.fromEntries((d.colaboradores || []).map(c => [c.matricula, c.nome]));
        Object.entries(d.resultados || {}).sort().forEach(([data, porMat]) => Object.entries(porMat || {}).forEach(([mat, v]) => lin.push({ data, mat, nome: nomes[mat] || '', ...Object.fromEntries(inds.map(i => [i.id, numDpo((v || {})[i.id])])) })));
        add('Resultados diários', [['Data', 'data', 12], ['Matrícula', 'mat', 12], ['Nome', 'nome', 28], ...inds.map(i => [i.nome, i.id, 14])], lin);
    } else if (chave === 'riscos') {
        const riscos = (d.riscos || []).filter(r => r.risco);
        add('Matriz de Riscos', [['Nº', 'n', 6], ['Risco', 'risco', 34], ['Descrição', 'descricao', 50], ['Plano de ação', 'plano', 50], ['Impacto/Dano', 'impacto', 14], ['Probabilidade', 'prob', 16], ['Tipo de risco', 'classe', 12], ['Tipo', 'tipo', 14], ['Perigo', 'perigo', 26], ['Mecanismo', 'mecanismo', 22], ['Frequência', 'freq', 12], ...M.map((m, i) => [m, 'm' + i, 6])],
            riscos.map(r => ({ ...r, classe: classeRiscoDpo(r), freq: freqRiscoDpo(classeRiscoDpo(r)), ...Object.fromEntries(M.map((m, i) => ['m' + i, ({ P: 'Plan', R: 'Real', X: 'Real (extra)' })[(r.verif || {})[i]] || ''])) })));
        add('Plano de resposta', [['Risco', 'risco', 30], ['Tipo', 'classe', 10], ['Proprietários', 'prop', 26], ['Procedimento quando o risco surgir', 'proc', 50], ['Ações no final do episódio', 'pos', 40], ['Nível de serviço / mão de obra', 'ns', 36], ['Responsável', 'resp', 24], ['Contato', 'cont', 24]],
            riscos.map(r => { const p = (d.respostas || {})[r.id] || {}; return { risco: r.risco, classe: classeRiscoDpo(r), prop: p.proprietarios, proc: p.procedimento, pos: p.posEpisodio, ns: p.nivelServico, resp: p.responsavel, cont: p.contato }; }));
        tabela('Plano de Retomada', d.retomada, [['Parada', 'parada', 30], ['Categoria', 'categoria', 20], ['Fornecedor', 'fornecedor', 24], ['Atividades', 'atividades', 26], ['Contato', 'contato', 26], ['Ações', 'acoes', 60]]);
        tabela('Histórico de Ocorrências', d.ocorrencias, [['Data', 'data', 12], ['Risco', 'risco', 28], ['Descrição', 'descricao', 50], ['Modificar procedimento?', 'mudar', 12], ['Ações corretivas', 'acoes', 40], ['Impacto R$', 'impacto', 12]]);
        tabela('Revisões da matriz', d.revisoes, [['Data', 'data', 12], ['Responsável', 'responsavel', 22], ['Time de segurança participou', 'seguranca', 12], ['Alterações', 'alteracoes', 70]]);
        tabela('Conversa com o time', d.entrevistas, [['Data', 'data', 12], ['Nome', 'nome', 22], ['Cargo', 'cargo', 20], ['Conhece os 3 riscos', 'top3', 12], ['Conhece o plano de resposta', 'plano', 12], ['Sabe onde está a retomada', 'local', 12]]);
    } else if (chave === 'sonho') {
        add('Sonho', [['Campo', 'c', 30], ['Conteúdo', 'v', 90]], [
            { c: 'Frase do Sonho', v: d.frase }, { c: 'Data da revisão', v: d.dataRevisao }, { c: 'Conexão com a estratégia ABI', v: d.conexao },
            { c: 'Onde está exposto', v: d.exposicao }, { c: 'Como foi construído (envolvimento)', v: d.construcao }]);
        const ytdSonho = k => { const v = (k.real || []).map(numDpo).filter(x => x !== null && x !== ''); if (!v.length) return ''; const r = k.acumula === 'soma' ? v.reduce((a, b) => a + b, 0) : k.acumula === 'ultimo' ? v[v.length - 1] : v.reduce((a, b) => a + b, 0) / v.length; return Math.round(r * 100) / 100; };
        const linhasSonho = [];
        (d.kpis || []).filter(k => k.nome).forEach(k => {
            const meta = numDpo(k.meta);
            const pai = (d.kpis || []).find(x => x.id === k.pai);
            linhasSonho.push({ nivel: k.nivel === 'estrategia' ? 'Estratégia' : 'Sonho', ligado: pai ? pai.nome : '', nome: k.nome, pilar: k.pilar || '', sentido: k.sentido === 'menor' ? '↓ menor melhor' : '↑ maior melhor', unidade: k.unidade || '', tipo: 'Meta', ...Object.fromEntries(M.map((m, i) => { const v = numDpo((k.metaMes || [])[i]); return ['m' + i, v === null || v === '' ? meta : v]; })), ytd: meta });
            linhasSonho.push({ nivel: '', ligado: '', nome: '', pilar: '', sentido: '', unidade: '', tipo: 'Real', ...Object.fromEntries(M.map((m, i) => ['m' + i, numDpo((k.real || [])[i])])), ytd: ytdSonho(k) });
        });
        add('Meta e Real', [['Nível', 'nivel', 11], ['Ligado a', 'ligado', 22], ['KPI', 'nome', 30], ['Pilar', 'pilar', 20], ['Sentido', 'sentido', 16], ['Unid.', 'unidade', 7], ['', 'tipo', 7], ...M.map((m, i) => [m, 'm' + i, 9]), ['Ano / YTD', 'ytd', 11]], linhasSonho);
        tabela('Propostas dos grupos', d.propostas, [['Grupo', 'grupo', 14], ['Proposta de Sonho', 'texto', 90], ['Escolhida?', 'escolhida', 10]]);
        tabela('Comunicação', d.comunicacoes, [['Data', 'data', 12], ['Canal', 'canal', 20], ['Público', 'publico', 25], ['Descrição', 'descricao', 60]]);
    } else if (chave === 'ans') {
        const linhas = [];
        Object.entries(d.meses || {}).forEach(([m, mes]) => Object.entries((mes || {}).dias || {}).forEach(([dia, v]) => {
            const neg = numDpo(v.negAuto) ?? numDpo(v.neg), real = numDpo(v.real);
            linhas.push({ mes: M[m], dia: Number(dia), neg, real, fora: numDpo(v.fora), buffer: numDpo(v.buffer), limiteLink: v.limiteLink || '', disp: neg && real !== null ? `${String(Math.round((real / neg - 1) * 1000) / 10).replace('.', ',')}%` : '', just: v.just || '' });
        }));
        linhas.sort((a, b) => M.indexOf(a.mes) - M.indexOf(b.mes) || a.dia - b.dia);
        add('Acompanhamento diário', [['Mês', 'mes', 8], ['Dia', 'dia', 6], ['Volume negociado', 'neg', 14], ['Volume realizado', 'real', 14], ['Fora de rota', 'fora', 12], ['Buffer %', 'buffer', 10], ['Limite fechamento link (hs)', 'limiteLink', 14], ['Dispersão', 'disp', 10], ['Justificativa', 'just', 40]], linhas);
        tabela('Ações', d.acoes, [['Mês', 'mesNome', 8], ['Dia', 'dia', 6], ['Dispersão', 'dispTxt', 10], ['Causa', 'causa', 35], ['Ação', 'acao', 45], ['Responsável', 'responsavel', 18], ['Prazo', 'prazo', 12], ['Status', 'status', 14]]);
        add('Meta mensal', [['Mês', 'mes', 10], ['Volume meta (hl)', 'meta', 16]], M.map((m, i) => ({ mes: m, meta: numDpo((d.metaMes || {})[i]) })));
        tabela('Reuniões de check', d.reunioes, [['Data', 'data', 12], ['Tipo', 'tipo', 12], ['Participantes', 'participantes', 40], ['Pontos / decisões', 'decisoes', 60]]);
    } else if (chave === 'capex') {
        tabela('Registro de solicitações', d.itens, [['Nº', 'n', 6], ['Data', 'data', 12], ['Descrição', 'descricao', 40], ['Área', 'area', 14], ['Categoria', 'categoria', 16], ['Tipo', 'tipo', 14], ['Origem', 'origem', 18], ['Justificativa', 'justificativa', 50], ['Qtde', 'qtd', 8], ['Custo unitário', 'unitario', 14], ['Custo total', 'total', 14], ['G', 'g', 5], ['U', 'u', 5], ['T', 't', 5], ['Etapa', 'etapa', 14], ['Orçado', 'orcado', 14], ['Aprovado', 'aprovado', 14], ['Realizado', 'realizado', 14], ['Início', 'inicio', 12], ['Fim previsto', 'fimPrevisto', 12], ['Fim real', 'fimReal', 12], ['NF', 'nfNome', 30]]);
        tabela('Checks trimestrais', d.checks, [['Trimestre', 'tri', 10], ['Data', 'data', 12], ['Participantes', 'participantes', 40], ['Decisões', 'decisoes', 60]]);
        tabela('Plano de ação', d.plano, [['Data', 'data', 12], ['Área', 'area', 14], ['Assunto', 'assunto', 20], ['Ação', 'acao', 45], ['Responsável', 'responsavel', 18], ['Prazo', 'prazo', 12], ['Status', 'status', 14]]);
    } else if (chave === 'p3a') {
        tabela('Projetos P1A-P3A', d.projetos, [['Horizonte', 'horizonte', 10], ['Descrição', 'descricao', 45], ['Recurso', 'recurso', 14], ['Tipo', 'tipo', 22], ['Qtde', 'qtd', 8], ['Valor unit.', 'unitario', 14], ['Previsto', 'previsto', 14], ['Realizado', 'realizado', 14], ['Ano', 'ano', 8], ['Data execução', 'execucao', 12], ['Prioridade', 'prioridade', 10], ['Status', 'status', 18], ['Registro', 'registro', 18], ['Impacto no negócio', 'impacto', 45]]);
        const volLin = [];
        Object.entries(d.volReal || {}).forEach(([ano, v]) => volLin.push({ ano: Number(ano), tipo: 'Realizado', volume: numDpo(v.volume), mktp: numDpo(v.mktp) }));
        Object.entries(d.volAdic || {}).forEach(([ano, v]) => volLin.push({ ano: Number(ano), tipo: 'Adicional planejado', volume: numDpo(v.vol), mktp: numDpo(v.mktp), motivo: v.motivo || '' }));
        volLin.sort((a, b) => a.ano - b.ano);
        tabela('Volume', volLin.length ? volLin : d.volumes, [['Ano', 'ano', 8], ['Tipo', 'tipo', 20], ['Volume (hl)', 'volume', 14], ['Marketplace (hl)', 'mktp', 14], ['Motivo do adicional', 'motivo', 50]]);
        tabela('Parâmetros da frota', d.paramFrota, [['Tipo', 'tipo', 14], ['hl por caminhão/dia', 'hlDia', 14], ['Entregas por caminhão/dia', 'entregasDia', 14], ['% do volume', 'part', 10], ['Qtde atual', 'atual', 10], ['Pessoas por caminhão', 'tripulacao', 12], ['Valor de compra', 'valor', 14]]);
        tabela('QLP', d.qlp, [['Função', 'funcao', 22], ['Atual', 'atual', 10], ['Ano 1', 'a1', 10], ['Ano 2', 'a2', 10], ['Ano 3', 'a3', 10], ['Justificativa', 'just', 45]]);
        tabela('Frota', d.frota, [['Placa', 'placa', 12], ['Tipo', 'tipo', 12], ['Ano fabricação', 'ano', 10], ['Utilização %', 'util', 12], ['Decisão', 'decisao', 22], ['Observação', 'obs', 40]]);
        tabela('Impactos financeiros', d.impactos, [['Investimento', 'descricao', 40], ['Valor', 'valor', 14], ['Parcelas', 'parcelas', 10], ['Juros % a.m.', 'juros', 10], ['Economia anual', 'economia', 14], ['Receita/ganho anual', 'ganho', 14]]);
    } else if (chave === 'manutencao') {
        const modelo = d.modelo || MODELO_MANUTENCAO_DPO;
        const linhas = [];
        modelo.forEach(s => s.grupos.forEach(g => g.itens.forEach(it => linhas.push({ secao: s.titulo, grupo: `${g.numero} ${g.titulo}`, n: it.num, q: it.texto, peso: it.peso, critico: it.critico ? 'Sim' : '', ...Object.fromEntries([0, 1, 2, 3].map(t => ['t' + t, ((d.notas || {})[t] || {})[it.id] ?? ''])) }))));
        add('Checklist', [['Seção', 'secao', 18], ['Grupo', 'grupo', 30], ['Nº', 'n', 6], ['Questão', 'q', 70], ['Peso', 'peso', 6], ['Crítico', 'critico', 8], ['T1', 't0', 6], ['T2', 't1', 6], ['T3', 't2', 6], ['T4', 't3', 6]], linhas);
        tabela('Plano de ação', d.acoes, [['Trimestre', 'triNome', 10], ['Item', 'itemTxt', 45], ['Crítico', 'criticoTxt', 8], ['Tratativa', 'destino', 22], ['Ação', 'acao', 45], ['Responsável', 'responsavel', 18], ['Prazo', 'prazo', 12], ['Status', 'status', 14]]);
        tabela('Base de fornecedores', d.fornecedores, [['Fornecedor', 'nome', 26], ['Contato', 'contato', 18], ['Tipo de serviço', 'servico', 24], ['Frequência', 'frequencia', 12], ['Cidade', 'cidade', 16], ['ANS / prazo de atendimento', 'ans', 24], ['Custo / contrato', 'custo', 18]]);
        if (Array.isArray(d.chamados) && d.chamados.length) {
            const hoje = new Date().toISOString().slice(0, 10), dias = c => c.abertura ? Math.max(0, Math.floor((new Date((['Concluído', 'Cancelado'].includes(c.status) && c.fechamento ? c.fechamento : hoje) + 'T12:00:00') - new Date(c.abertura + 'T12:00:00')) / 86400000)) : '';
            add('Chamados de manutenção', [['Nº', 'n', 6], ['Abertura', 'abertura', 12], ['Chamado', 'titulo', 40], ['Local', 'local', 18], ['Categoria', 'categoria', 20], ['Prioridade', 'prioridade', 12], ['Responsável', 'responsavel', 20], ['Prazo', 'prazo', 12], ['Status', 'status', 18], ['Fechamento', 'fechamento', 12], ['Dias abertos', 'dias', 10], ['Solução', 'solucao', 40], ['Custo', 'custo', 12]],
                d.chamados.map(c => ({ ...c, prazo: c.prazoManual || c.prazoCalc || '', dias: dias(c) })));
            const acs = []; d.chamados.forEach(c => (c.acoes || []).forEach(a => acs.push({ n: c.n, chamado: c.titulo, ...a })));
            tabela('Ações dos chamados', acs, [['Chamado nº', 'n', 8], ['Chamado', 'chamado', 34], ['Ação', 'acao', 40], ['Responsável', 'responsavel', 20], ['Prazo', 'prazo', 12], ['Status', 'status', 14]]);
            tabela('Prazos de fechamento', d.slaChamados, [['Categoria', 'categoria', 26], ['Prioridade', 'prioridade', 14], ['Dias', 'dias', 8]]);
        }
        tabela('RACI', d.raci, [['Atividade / item', 'atividade', 34], ['R', 'r', 18], ['A', 'a', 18], ['C', 'c', 18], ['I', 'i', 18], ['Fornecedor', 'fornecedor', 20]]);
    }
}

const IMPACTOS_RISCO_DPO = ['Insignificante', 'Menor', 'Moderado', 'Maior', 'Extremo'];
const PROBS_RISCO_DPO = ['Raro', 'Improvável', 'Possível', 'Provável', 'Quase certamente'];
const MATRIZ_RISCO_DPO = { Extremo: ['Alto', 'Crítico', 'Crítico', 'Crítico', 'Crítico'], Maior: ['Alto', 'Alto', 'Crítico', 'Crítico', 'Crítico'], Moderado: ['Moderado', 'Moderado', 'Alto', 'Alto', 'Crítico'], Menor: ['Baixo', 'Baixo', 'Moderado', 'Alto', 'Alto'], Insignificante: ['Baixo', 'Baixo', 'Baixo', 'Moderado', 'Alto'] };
function normRiscoDpo(v, lista) { const t = semAcentoGop(v); if (!t) return ''; if (t === 'ALTO') return 'Provável'; if (t === 'ESTRANHO') return 'Raro'; return lista.find(x => semAcentoGop(x) === t) || lista.find(x => t.startsWith(semAcentoGop(x).slice(0, 4))) || ''; }
function classeRiscoDpo(r) { if (r.classeManual) return r.classeManual; const l = MATRIZ_RISCO_DPO[r.impacto], i = PROBS_RISCO_DPO.indexOf(r.prob); return l && i >= 0 ? l[i] : ''; }
function freqRiscoDpo(c) { return c === 'Crítico' || c === 'Alto' ? 'Trimestral' : c === 'Moderado' ? 'Semestral' : c === 'Baixo' ? 'Anual' : ''; }
// Lê a Matriz de Riscos Externos (abas Matriz de Riscos, Histórico, Plano de Retomada, Matriz de Contatos e uma aba por plano de ação).
function lerWorkbookRiscosDpo(wb) {
    const txt = v => { const x = valorCelulaGop(v); return x === null || x === undefined ? '' : String(x instanceof Date ? x.toISOString().slice(0, 10) : x).trim(); };
    const out = { riscos: [], ocorrencias: [], retomada: [], contatos: {}, planos: [] }, avisos = [];
    const acharCab = (ws, testes) => { for (let r = 1; r <= Math.min(ws.rowCount, 15); r++) { const row = ws.getRow(r), cols = {}; row.eachCell((c, n) => { const t = semAcentoGop(txt(c.value)); Object.entries(testes).forEach(([k, re]) => { if (cols[k] === undefined && re.test(t)) cols[k] = n; }); }); if (Object.keys(cols).length >= Math.min(3, Object.keys(testes).length)) return { linha: r, cols }; } return null; };
    wb.worksheets.forEach(ws => {
        const nome = semAcentoGop(ws.name);
        if (/^MATRIZ DE RISCOS/.test(nome)) {
            const h = acharCab(ws, { n: /^N[°º]?$/, risco: /^RISCO$/, desc: /^DESCRICAO/, plano: /^PLANO DE ACAO/, imp: /^IMPACTO/, prob: /^PROBABILIDADE/, tipo: /^TIPO$/, perigo: /^PERIGO/, mec: /^MECANISMO/, pr: /^PLAN\s*\//, jan: /^JAN/ });
            if (!h) { avisos.push('Aba Matriz de Riscos sem cabeçalho reconhecido.'); return; }
            let atual = null;
            for (let r = h.linha + 1; r <= ws.rowCount; r++) {
                const g = k => h.cols[k] ? txt(ws.getRow(r).getCell(h.cols[k]).value) : '';
                const risco = g('risco'), pr = semAcentoGop(g('pr'));
                if (risco) {
                    atual = { id: idGop(), n: Number(g('n')) || out.riscos.length + 1, risco, descricao: g('desc'), plano: g('plano'), impacto: normRiscoDpo(g('imp'), IMPACTOS_RISCO_DPO), prob: normRiscoDpo(g('prob'), PROBS_RISCO_DPO), tipo: g('tipo'), perigo: g('perigo'), mecanismo: g('mec'), verif: {} };
                    out.riscos.push(atual);
                }
                if (atual && h.cols.jan && (pr === 'PLAN' || pr === 'REAL')) {
                    for (let m = 0; m < 12; m++) {
                        const v = Number(txt(ws.getRow(r).getCell(h.cols.jan + m).value).replace(',', '.'));
                        if (!(v > 0)) continue;
                        const ant = atual.verif[m];
                        atual.verif[m] = pr === 'PLAN' ? (ant === 'R' || ant === 'X' ? 'R' : 'P') : (ant === 'P' || ant === 'R' ? 'R' : 'X');
                    }
                }
            }
        } else if (/^HISTORICO/.test(nome)) {
            const h = acharCab(ws, { data: /^DATA$/, risco: /^RISCO$/, desc: /^DESCRICAO/, mudar: /MODIFICAR/, acoes: /^ACOES/, imp: /^IMPACTO/ });
            if (!h) return;
            for (let r = h.linha + 1; r <= ws.rowCount; r++) {
                const g = k => h.cols[k] ? ws.getRow(r).getCell(h.cols[k]).value : '';
                const risco = txt(g('risco')); if (!risco) continue;
                out.ocorrencias.push({ id: idGop(), data: dataCelulaGop(g('data')), risco, descricao: txt(g('desc')), mudar: /^S/i.test(txt(g('mudar'))) ? 'Sim' : 'Não', acoes: txt(g('acoes')), impacto: Number(txt(g('imp')).replace(',', '.')) || '' });
            }
        } else if (/^PLANO DE RETOMADA/.test(nome)) {
            const h = acharCab(ws, { parada: /^PARADA/, forn: /^FORNECEDOR/, ativ: /^ATIVIDADE/, cont: /^CONTATO/, acoes: /^ACOES/ });
            if (!h) return;
            for (let r = h.linha + 1; r <= ws.rowCount; r++) {
                const g = k => h.cols[k] ? txt(ws.getRow(r).getCell(h.cols[k]).value) : '';
                if (!g('parada')) continue;
                out.retomada.push({ id: idGop(), parada: g('parada'), categoria: categoriaRetomadaDpo(g('parada') + ' ' + g('ativ')), fornecedor: g('forn'), atividades: g('ativ'), contato: g('cont'), acoes: g('acoes') });
            }
        } else if (/^MATRIZ DE CONTATOS/.test(nome)) {
            const h = acharCab(ws, { risco: /^MATRIZ DE CONTATOS/, acao: /^ACAO IMEDIATA/, resp: /^RESPONSAVEL/, cont: /^CONTATO/ });
            if (!h) return;
            for (let r = h.linha + 1; r <= ws.rowCount; r++) {
                const g = k => h.cols[k] ? txt(ws.getRow(r).getCell(h.cols[k]).value) : '';
                if (g('risco')) out.contatos[semAcentoGop(g('risco'))] = { acao: g('acao'), resp: g('resp'), cont: g('cont') };
            }
        } else {
            // aba de plano de ação por risco ("Plano de Ação contra Riscos Externos")
            let ehPlano = false; const campos = {};
            for (let r = 1; r <= Math.min(ws.rowCount, 40); r++) ws.getRow(r).eachCell(c => { const t = txt(c.value); if (/Plano de A[cç][aã]o contra Riscos/i.test(t)) ehPlano = true; });
            if (!ehPlano) return;
            const rotulos = { prop: /^PROPRIETARIO/, desc: /^DESCRICAO/, proc: /^PROCEDIMENTO/, pos: /^ACOES A SEREM/, imp: /^IMPACTO NA OPERACAO/, freq: /^FREQUENCIA/ };
            let risco = '', chaveAtual = null;
            for (let r = 1; r <= ws.rowCount; r++) {
                ws.getRow(r).eachCell(c => {
                    const t = txt(c.value); if (!t) return;
                    const m = t.match(/^Risco\s*:\s*(.+)$/i); if (m) { risco = m[1].trim(); return; }
                    const k = Object.keys(rotulos).find(k => rotulos[k].test(semAcentoGop(t)));
                    if (k) { chaveAtual = k; return; }
                    if (chaveAtual && !/^N[°º]\s*RISCO/i.test(t) && !/^\d+$/.test(t)) campos[chaveAtual] = (campos[chaveAtual] ? campos[chaveAtual] + '\n' : '') + t;
                });
            }
            if (risco) out.planos.push({ risco, ...campos });
        }
    });
    return { ...out, avisos };
}
function categoriaRetomadaDpo(t) {
    const x = semAcentoGop(t);
    if (/ARMAZENAMENTO|ARMAZEM EXTERNO/.test(x)) return 'Armazenamento externo';
    if (/ALUGUE|LOCADORA|FROTA EXTRA/.test(x)) return 'Aluguel de veículos';
    if (/SISTEMA|INTERNET|SERVIDOR|WMS|PROMAX|BEES/.test(x)) return 'Reparo de sistemas';
    if (/CFTV|AR-CONDICIONADO|AR CONDICIONADO|EQUIPAMENTO|ENERGIA|ELETRIC|GERADOR|PORTAO|EMPILHADEIRA/.test(x)) return 'Equipamentos';
    if (/BOMBEIRO|DEFESA CIVIL|ACIDENTE|ALAGAMENTO|POLICIA/.test(x)) return 'Emergência / órgãos públicos';
    return 'Fornecedores';
}
app.post('/api/dpo/riscos/importar', requireRole('admin', 'client_admin'), async (req, res) => {
    const { url, originalName } = req.body;
    if (!/^\/uploads\/[\w.\-]+$/.test(String(url || ''))) return res.status(400).json({ error: 'Envie a planilha antes de importar.' });
    try {
        const f = await resolverFerramentaDigitalDpo(req, res, 'riscos', req.body.company_id);
        if (!f) return;
        const ano = anoValidoDpo(req.body.ano);
        let wb;
        try { wb = await lerArquivoPlanilhaDpo(path.join(PASTA_UPLOADS, path.basename(url))); }
        catch (e) { return res.status(400).json({ error: /\.xlsb$/i.test(url) ? 'Não consegui ler o formato .xlsb. Salve como .xlsx e importe de novo.' : 'Não foi possível ler a planilha.' }); }
        const lido = lerWorkbookRiscosDpo(wb);
        if (!lido.riscos.length) return res.status(400).json({ error: 'Não encontrei a aba "Matriz de Riscos".', avisos: lido.avisos });
        const d = (await carregarFerramentaDigitalDpo(f.companyId, 'riscos', ano)).dados || {};
        d.riscos = Array.isArray(d.riscos) ? d.riscos : []; d.respostas = d.respostas || {}; d.retomada = Array.isArray(d.retomada) ? d.retomada : []; d.ocorrencias = Array.isArray(d.ocorrencias) ? d.ocorrencias : [];
        const porNome = Object.fromEntries(d.riscos.map(r => [semAcentoGop(r.risco), r]));
        lido.riscos.forEach(r => { const ex = porNome[semAcentoGop(r.risco)]; if (ex) { Object.assign(ex, { ...r, id: ex.id }); } else { d.riscos.push(r); porNome[semAcentoGop(r.risco)] = r; } });
        const acharRisco = nome => { const t = semAcentoGop(nome); return porNome[t] || d.riscos.find(r => { const a = semAcentoGop(r.risco); return a.includes(t) || t.includes(a) || a.split(' ')[0] === t.split(' ')[0] && a.split(' ')[0].length > 5; }); };
        Object.entries(lido.contatos).forEach(([k, c]) => { const r = porNome[k] || acharRisco(k); if (!r) return; const p = d.respostas[r.id] = d.respostas[r.id] || {}; if (c.resp) p.responsavel = c.resp; if (c.cont) p.contato = c.cont; if (c.acao && !p.procedimento) p.procedimento = c.acao; });
        lido.planos.forEach(pl => { const r = acharRisco(pl.risco); if (!r) return; const p = d.respostas[r.id] = d.respostas[r.id] || {}; if (pl.prop) p.proprietarios = pl.prop; if (pl.proc) p.procedimento = pl.proc; if (pl.pos) p.posEpisodio = pl.pos; if (pl.imp) p.impactoOperacao = pl.imp; });
        const jaRet = new Set(d.retomada.map(x => semAcentoGop(x.parada)));
        lido.retomada.forEach(x => { const ex = d.retomada.find(y => semAcentoGop(y.parada) === semAcentoGop(x.parada)); if (ex) Object.assign(ex, { ...x, id: ex.id }); else if (!jaRet.has(semAcentoGop(x.parada))) d.retomada.push(x); });
        const jaOc = new Set(d.ocorrencias.map(o => o.data + '|' + semAcentoGop(o.risco) + '|' + semAcentoGop(o.descricao)));
        lido.ocorrencias.forEach(o => { if (!jaOc.has(o.data + '|' + semAcentoGop(o.risco) + '|' + semAcentoGop(o.descricao))) d.ocorrencias.push(o); });
        await new Promise((resolve, reject) => db.run(
            `INSERT INTO dpo_ferramentas_digitais (company_id, chave, ano, dados, updated_by, updated_at) VALUES (?, 'riscos', ?, ?, ?, CURRENT_TIMESTAMP)
             ON CONFLICT(company_id, chave, ano) DO UPDATE SET dados = excluded.dados, updated_by = excluded.updated_by, updated_at = CURRENT_TIMESTAMP`,
            [f.companyId, ano, JSON.stringify(d), req.user.userId], (err) => err ? reject(err) : resolve()));
        db.run(`INSERT INTO dpo_ferramentas_digitais_arquivos (company_id, chave, ano, tipo, url, original_name, comentario, created_by) VALUES (?, 'riscos', ?, 'matriz', ?, ?, ?, ?)`,
            [f.companyId, ano, url, originalName || 'Matriz de Riscos Externos', 'Importada para o sistema', req.user.userId], () => {});
        res.json({ message: `${lido.riscos.length} risco(s), ${lido.ocorrencias.length} ocorrência(s), ${lido.retomada.length} item(ns) do plano de retomada e ${lido.planos.length} plano(s) de resposta importados.`, dados: d, avisos: lido.avisos });
    } catch (e) {
        console.error('Erro ao importar matriz de riscos:', e.message);
        res.status(400).json({ error: 'Não foi possível importar a matriz de riscos.' });
    }
});

// ======================= CHAMADO DE MANUTENÇÃO PELO LINK (sem login) =======================
app.post('/api/dpo/manutencao/chamado-link', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const f = await resolverFerramentaDigitalDpo(req, res, 'manutencao', req.body.company_id);
        if (!f) return;
        let reg = await dbGet(`SELECT token FROM dpo_chamado_links WHERE company_id = ?`, [f.companyId]);
        if (!reg) { reg = { token: crypto.randomBytes(16).toString('hex') }; await new Promise((resolve, reject) => db.run(`INSERT INTO dpo_chamado_links (token, company_id, criado_por) VALUES (?, ?, ?)`, [reg.token, f.companyId, req.user.userId], e => e ? reject(e) : resolve())); }
        res.json({ url: `${baseUrlPublicaDpo(req)}/abrir-chamado.html?t=${reg.token}`, token: reg.token });
    } catch (e) { res.status(400).json({ error: 'Erro ao gerar o link de chamados.' }); }
});
app.get('/api/dpo/manutencao/chamados-celular', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const f = await resolverFerramentaDigitalDpo(req, res, 'manutencao', req.query.company_id);
        if (!f) return;
        const l = await dbAll(`SELECT * FROM dpo_chamados_pub WHERE company_id = ? ORDER BY id`, [f.companyId]);
        res.json(l.map(r => ({ id: r.id, foto: r.foto, created_at: r.created_at, ...JSON.parse(r.dados || '{}') })));
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar os chamados do link.' }); }
});
async function linkChamadoDpo(token) { const l = await dbGet(`SELECT * FROM dpo_chamado_links WHERE token = ?`, [String(token || '')]); if (!l) return null; const emp = await dbGet(`SELECT name FROM companies WHERE id = ?`, [l.company_id]); return { link: l, empresa: emp ? emp.name : '' }; }
app.get('/api/public/chamado/:token', async (req, res) => {
    const r = await linkChamadoDpo(req.params.token).catch(() => null);
    if (!r) return res.status(404).json({ error: 'Link de chamados inválido.' });
    const meus = req.query.tel ? await dbAll(`SELECT id, dados, created_at FROM dpo_chamados_pub WHERE company_id = ? ORDER BY id DESC LIMIT 200`, [r.link.company_id]) : [];
    const tel = String(req.query.tel || '').replace(/\D/g, '');
    // situação atual dos chamados deste telefone (lida da ferramenta)
    let status = {};
    if (tel) { const ano = new Date().getFullYear(); for (const a of [ano, ano - 1]) { const d = (await carregarFerramentaDigitalDpo(r.link.company_id, 'manutencao', a)).dados || {}; (d.chamados || []).forEach(c => { if (c.pubId) status[c.pubId] = { n: c.n, status: c.status || 'Aberto', prazo: c.prazoManual || c.prazoCalc || '', responsavel: c.responsavel || '' }; }); } }
    res.json({ empresa: r.empresa, meus: meus.map(m => ({ id: m.id, created_at: m.created_at, ...JSON.parse(m.dados || '{}') })).filter(m => tel && String(m.telefone || '').replace(/\D/g, '') === tel).slice(0, 20).map(m => ({ protocolo: 'C-' + m.id, titulo: m.titulo, created_at: m.created_at, ...(status[m.id] || { status: 'Recebido' }) })) });
});
app.post('/api/public/chamado/:token', (req, res) => {
    uploadMaterialDpo.single('file')(req, res, async (err) => {
        if (err) return res.status(400).json({ error: err.code === 'LIMIT_FILE_SIZE' ? 'Foto muito grande.' : err.message });
        try {
            const r = await linkChamadoDpo(req.params.token);
            if (!r) return res.status(404).json({ error: 'Link de chamados inválido.' });
            const b = req.body || {}, t = (k, n) => String(b[k] || '').trim().slice(0, n);
            if (!t('titulo', 200) || !t('solicitante', 80)) return res.status(400).json({ error: 'Informe seu nome e o que precisa ser feito.' });
            const dados = { solicitante: t('solicitante', 80), telefone: t('telefone', 30), local: t('local', 120), categoria: t('categoria', 60), prioridade: ['Emergencial', 'Alta', 'Média', 'Baixa'].includes(b.prioridade) ? b.prioridade : 'Média', titulo: t('titulo', 200), descricao: t('descricao', 2000) };
            const foto = req.file ? '/uploads/' + req.file.filename : null;
            const id = await new Promise((resolve, reject) => db.run(`INSERT INTO dpo_chamados_pub (company_id, dados, foto) VALUES (?, ?, ?)`, [r.link.company_id, JSON.stringify(dados), foto], function (e) { e ? reject(e) : resolve(this.lastID); }));
            notificarPorCompanyAdmins(r.link.company_id, 'Novo chamado de manutenção', `C-${id}: ${dados.titulo} (${dados.local || 'sem local'}) — aberto por ${dados.solicitante}.`);
            res.json({ message: 'Chamado aberto!', protocolo: 'C-' + id });
        } catch (e) { res.status(400).json({ error: 'Erro ao abrir o chamado.' }); }
    });
});

// ======================= 5S: CALENDÁRIO E AVISOS AUTOMÁTICOS AOS AUDITORES =======================
let ULTIMA_BASE_DPO = '';
function baseAvisosDpo() { return urlPublicaValida(appBaseUrlAtiva) ? appBaseUrlAtiva.replace(/\/$/, '') : ULTIMA_BASE_DPO; }
function agoraLocalDpo() { const s = new Date().toLocaleString('en-US', { timeZone: 'America/Sao_Paulo' }); return new Date(s); }
function dataLocalDpo(d) { return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`; }
async function token5sDpo(companyId, ano) {
    let reg = await dbGet(`SELECT token FROM dpo_5s_links WHERE company_id = ? AND ano = ?`, [companyId, ano]);
    if (!reg) { reg = { token: crypto.randomBytes(16).toString('hex') }; await new Promise((resolve, reject) => db.run(`INSERT INTO dpo_5s_links (token, company_id, ano) VALUES (?, ?, ?)`, [reg.token, companyId, ano], e => e ? reject(e) : resolve())); }
    return reg.token;
}
async function auditoria5sFeitaDpo(companyId, ano, dados, areaId, mes) {
    const perguntas = dados.modelo.sensos.flatMap(s => s.perguntas.map(p => p.id));
    const au = (dados.auditorias || []).find(a => a.areaId === areaId && a.mes === mes);
    const resp = new Set(Object.entries((au && au.resp) || {}).filter(([, v]) => v).map(([q]) => q));
    if (au && au.origem === 'planilha' && au.notaInformada !== undefined) return true;
    (await dbAll(`SELECT qid FROM dpo_5s_resp WHERE company_id = ? AND ano = ? AND area_id = ? AND mes = ? AND resp IS NOT NULL AND resp != ''`, [companyId, ano, areaId, mes])).forEach(r => resp.add(r.qid));
    return resp.size >= perguntas.length;
}
function auditor5sDpo(dados, area) {
    const nome = String(area.auditor || '').trim().toLowerCase();
    return (dados.auditores || []).find(a => String(a.nome || '').trim().toLowerCase() === nome) || null;
}
async function enviarAviso5sDpo(companyId, ano, dados, area, mes, data, base) {
    const aud = auditor5sDpo(dados, area);
    const token = await token5sDpo(companyId, ano);
    const url = `${base || ''}/auditoria-5s.html?t=${token}&area=${encodeURIComponent(area.id)}&mes=${mes}`;
    const texto = `🧹 *Auditoria 5S pendente*\nOlá${aud && aud.nome ? ' ' + aud.nome.split(' ')[0] : ''}! A auditoria 5S da área *${area.nome}*${area.placa ? ` (${area.placa})` : ''} está agendada para ${data.split('-').reverse().join('/')}.\nFaça pelo celular: ${url}\nEste aviso se repete até a auditoria ser concluída.`;
    const out = { whats: false, email: false, url, texto, auditor: aud ? aud.nome : area.auditor || '', telefone: aud ? aud.whatsapp || '' : '', emailDest: aud ? aud.email || '' : '' };
    if (aud && aud.whatsapp && (dados.avisos || {}).whats !== false) { const n = String(aud.whatsapp).replace(/\D/g, ''); out.whats = await enviarWhatsApp('+' + (n.length <= 11 ? '55' + n : n), texto); }
    if (aud && aud.email && (dados.avisos || {}).email !== false) {
        try { await transporter.sendMail({ from: process.env.SMTP_FROM || '"Impulsionar V4" <no-reply@impulsionar.com>', to: aud.email, subject: `Auditoria 5S pendente — ${area.nome}`, text: texto.replace(/\*/g, ''), html: `<p>Olá${aud.nome ? ' ' + escapeHtmlEmailDpo(aud.nome.split(' ')[0]) : ''}!</p><p>A auditoria 5S da área <b>${escapeHtmlEmailDpo(area.nome)}</b> está agendada para <b>${data.split('-').reverse().join('/')}</b>.</p><p><a href="${url}" style="background:#6d28d9;color:#fff;padding:10px 16px;border-radius:8px;text-decoration:none;font-weight:bold;">Fazer a auditoria pelo celular</a></p><p style="color:#64748b;font-size:12px;">Este aviso se repete até a auditoria ser concluída.</p>` }); out.email = true; }
        catch (e) { console.warn('⚠️  Aviso 5S por e-mail falhou:', e.message); }
    }
    await new Promise(resolve => db.run(`INSERT INTO dpo_5s_avisos (company_id, ano, mes, area_id, data, envios, ultimo, canais) VALUES (?, ?, ?, ?, ?, 1, ?, ?)
        ON CONFLICT(company_id, ano, mes, area_id) DO UPDATE SET envios = envios + 1, ultimo = excluded.ultimo, data = excluded.data, canais = excluded.canais`,
        [companyId, ano, mes, area.id, data, new Date().toISOString(), [out.whats ? 'WhatsApp' : '', out.email ? 'E-mail' : ''].filter(Boolean).join(' + ') || 'sem canal'], () => resolve()));
    return out;
}
function escapeHtmlEmailDpo(s) { return String(s || '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
let AVISOS_5S_RODANDO = false;
async function rodarAvisos5sDpo() {
    if (AVISOS_5S_RODANDO) return; AVISOS_5S_RODANDO = true;
    try {
        const agora = agoraLocalDpo(), hoje = dataLocalDpo(agora), ano = agora.getFullYear(), hora = agora.getHours() + agora.getMinutes() / 60;
        const base = baseAvisosDpo();
        const regs = await dbAll(`SELECT company_id, dados FROM dpo_ferramentas_digitais WHERE chave = 'cinco_s' AND ano = ?`, [ano]);
        for (const reg of regs) {
            const dados = prepararDados5sDpo(JSON.parse(reg.dados || '{}')), cfg = dados.avisos || {};
            if (cfg.ativo === false) continue;
            const ini = Number(cfg.horaInicio ?? 8), fim = Number(cfg.horaFim ?? 18), intervalo = Math.max(1, Number(cfg.intervalo ?? 2));
            if (hora < ini || hora >= fim) continue;
            for (const [mesTxt, cal] of Object.entries(dados.calendario || {})) {
                const mes = Number(mesTxt);
                for (const [areaId, data] of Object.entries(cal || {})) {
                    if (!data || data > hoje || data.slice(0, 4) !== String(ano)) continue;
                    if (data < hoje && cfg.atrasadas === false) continue;
                    const area = dados.areas.find(a => a.id === areaId); if (!area) continue;
                    if (await auditoria5sFeitaDpo(reg.company_id, ano, dados, areaId, mes)) continue;
                    const av = await dbGet(`SELECT ultimo FROM dpo_5s_avisos WHERE company_id = ? AND ano = ? AND mes = ? AND area_id = ?`, [reg.company_id, ano, mes, areaId]);
                    if (av && av.ultimo && Date.now() - new Date(av.ultimo).getTime() < intervalo * 3600000 - 60000) continue;
                    if (!auditor5sDpo(dados, area)) continue; // sem contato cadastrado
                    await enviarAviso5sDpo(reg.company_id, ano, dados, area, mes, data, base);
                }
            }
        }
    } catch (e) { console.warn('⚠️  Rotina de avisos 5S:', e.message); }
    finally { AVISOS_5S_RODANDO = false; }
}
setInterval(rodarAvisos5sDpo, 10 * 60 * 1000);
setTimeout(rodarAvisos5sDpo, 60 * 1000);
app.get('/api/dpo/cinco-s/avisos', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const f = await resolverFerramentaDigitalDpo(req, res, 'cinco_s', req.query.company_id);
        if (!f) return;
        res.json({ avisos: await dbAll(`SELECT mes, area_id, data, envios, ultimo, canais FROM dpo_5s_avisos WHERE company_id = ? AND ano = ?`, [f.companyId, anoValidoDpo(req.query.ano)]), whatsConfigurado: !!(TWILIO_ACCOUNT_SID && TWILIO_AUTH_TOKEN && TWILIO_WHATSAPP_FROM) });
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar os avisos.' }); }
});
app.post('/api/dpo/cinco-s/avisar', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const f = await resolverFerramentaDigitalDpo(req, res, 'cinco_s', req.body.company_id);
        if (!f) return;
        const ano = anoValidoDpo(req.body.ano), mes = Number(req.body.mes);
        const dados = prepararDados5sDpo((await carregarFerramentaDigitalDpo(f.companyId, 'cinco_s', ano)).dados);
        const area = dados.areas.find(a => a.id === req.body.areaId); if (!area) return res.status(400).json({ error: 'Área não encontrada.' });
        const data = ((dados.calendario || {})[mes] || {})[area.id] || dataLocalDpo(agoraLocalDpo());
        const out = await enviarAviso5sDpo(f.companyId, ano, dados, area, mes, data, baseUrlPublicaDpo(req));
        res.json({ ...out, message: out.whats || out.email ? `Aviso enviado por ${[out.whats ? 'WhatsApp' : '', out.email ? 'e-mail' : ''].filter(Boolean).join(' e ')}.` : 'Não foi possível enviar automaticamente (contato do auditor ou WhatsApp/e-mail do servidor não configurados) — use o botão do WhatsApp.' });
    } catch (e) { res.status(400).json({ error: 'Erro ao enviar o aviso.' }); }
});

// ======================= VISIBILIDADE DE RESULTADOS DA DISTRIBUIÇÃO (Entrega 3.1) =======================
async function linkVisDpo(token) { const l = await dbGet(`SELECT * FROM dpo_vis_links WHERE token = ?`, [String(token || '')]); if (!l) return null; const emp = await dbGet(`SELECT name FROM companies WHERE id = ?`, [l.company_id]); return { link: l, empresa: emp ? emp.name : '' }; }
async function dadosVisDpo(companyId, ano) {
    let d = (await carregarFerramentaDigitalDpo(companyId, 'visibilidade', ano)).dados || {};
    if (!d.colaboradores) { const ult = await dbGet(`SELECT ano FROM dpo_ferramentas_digitais WHERE company_id = ? AND chave = 'visibilidade' ORDER BY ano DESC LIMIT 1`, [companyId]); if (ult && ult.ano !== ano) d = (await carregarFerramentaDigitalDpo(companyId, 'visibilidade', ult.ano)).dados || {}; }
    return d;
}
app.post('/api/dpo/visibilidade/link', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const f = await resolverFerramentaDigitalDpo(req, res, 'visibilidade', req.body.company_id);
        if (!f) return;
        let reg = await dbGet(`SELECT token FROM dpo_vis_links WHERE company_id = ?`, [f.companyId]);
        if (!reg) { reg = { token: crypto.randomBytes(16).toString('hex') }; await new Promise((resolve, reject) => db.run(`INSERT INTO dpo_vis_links (token, company_id, criado_por) VALUES (?, ?, ?)`, [reg.token, f.companyId, req.user.userId], e => e ? reject(e) : resolve())); }
        res.json({ url: `${baseUrlPublicaDpo(req)}/resultados.html?t=${reg.token}`, token: reg.token });
    } catch (e) { res.status(400).json({ error: 'Erro ao gerar o link.' }); }
});
app.get('/api/dpo/visibilidade/registros', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const f = await resolverFerramentaDigitalDpo(req, res, 'visibilidade', req.query.company_id);
        if (!f) return;
        const ano = anoValidoDpo(req.query.ano);
        const just = await dbAll(`SELECT * FROM dpo_vis_just WHERE company_id = ? AND substr(data, 1, 4) = ? ORDER BY data DESC, id DESC`, [f.companyId, String(ano)]);
        const acessos = await dbAll(`SELECT matricula, data, qtd FROM dpo_vis_acessos WHERE company_id = ? AND substr(data, 1, 4) = ?`, [f.companyId, String(ano)]);
        res.json({ justificativas: just.map(j => ({ ...j, dados: JSON.parse(j.dados || '{}') })), acessos });
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar as justificativas.' }); }
});
app.put('/api/dpo/visibilidade/justificativas/:id', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const j = await dbGet(`SELECT * FROM dpo_vis_just WHERE id = ?`, [req.params.id]); if (!j) return res.status(404).json({ error: 'Justificativa não encontrada.' });
        const f = await resolverFerramentaDigitalDpo(req, res, 'visibilidade', j.company_id); if (!f) return;
        if (String(f.companyId) !== String(j.company_id)) return res.status(403).json({ error: 'Sem acesso.' });
        await new Promise((resolve, reject) => db.run(`UPDATE dpo_vis_just SET status = ?, retorno = ? WHERE id = ?`, [String(req.body.status || 'Validada').slice(0, 30), String(req.body.retorno || '').slice(0, 1000), j.id], e => e ? reject(e) : resolve()));
        res.json({ message: 'Justificativa atualizada.' });
    } catch (e) { res.status(400).json({ error: 'Erro ao atualizar.' }); }
});
function colaboradorVisDpo(d, matricula) { const m = String(matricula || '').trim().replace(/^0+/, '').toUpperCase(); return (d.colaboradores || []).find(c => String(c.matricula || '').trim().replace(/^0+/, '').toUpperCase() === m && c.ativo !== 'Não'); }
app.get('/api/public/vis/:token', async (req, res) => {
    const r = await linkVisDpo(req.params.token).catch(() => null);
    if (!r) return res.status(404).json({ error: 'Link inválido.' });
    res.json({ empresa: r.empresa });
});
app.post('/api/public/vis/:token/consulta', async (req, res) => {
    try {
        const r = await linkVisDpo(req.params.token);
        if (!r) return res.status(404).json({ error: 'Link inválido.' });
        const hoje = dataLocalDpo(agoraLocalDpo()), mesRef = /^\d{4}-\d{2}$/.test(String(req.body.mes || '')) ? req.body.mes : hoje.slice(0, 7), ano = Number(mesRef.slice(0, 4));
        const d = await dadosVisDpo(r.link.company_id, ano), c = colaboradorVisDpo(d, req.body.matricula);
        if (!c) return res.status(404).json({ error: 'Matrícula não encontrada. Confira o número ou fale com seu supervisor.' });
        const dias = Object.entries(d.resultados || {}).filter(([data]) => data.startsWith(mesRef)).map(([data, porMat]) => ({ data, valores: (porMat || {})[c.matricula] || null })).filter(x => x.valores).sort((a, b) => a.data.localeCompare(b.data));
        const just = await dbAll(`SELECT data, ind_id, dados, status, retorno, created_at FROM dpo_vis_just WHERE company_id = ? AND matricula = ? AND data LIKE ?`, [r.link.company_id, c.matricula, mesRef + '%']);
        const anexos = (d.anexosInd || []).filter(a => String(a.data || '').startsWith(mesRef)).map(a => ({ indId: a.indId, data: a.data, url: a.url, nome: a.nome }));
        await new Promise(resolve => db.run(`INSERT INTO dpo_vis_acessos (company_id, matricula, data, qtd) VALUES (?, ?, ?, 1) ON CONFLICT(company_id, matricula, data) DO UPDATE SET qtd = qtd + 1`, [r.link.company_id, c.matricula, hoje], () => resolve()));
        res.json({ empresa: r.empresa, mes: mesRef, colaborador: { matricula: c.matricula, nome: c.nome, funcao: c.funcao, supervisor: c.supervisor, placa: c.placa },
            indicadores: (d.indicadores || []).filter(i => i.nome), incentivo: d.incentivo || {}, dias, justificativas: just.map(j => ({ ...j, dados: JSON.parse(j.dados || '{}') })), anexos });
    } catch (e) { res.status(500).json({ error: 'Erro ao consultar os resultados.' }); }
});
app.post('/api/public/vis/:token/justificar', async (req, res) => {
    try {
        const r = await linkVisDpo(req.params.token);
        if (!r) return res.status(404).json({ error: 'Link inválido.' });
        const data = String(req.body.data || ''); if (!/^\d{4}-\d{2}-\d{2}$/.test(data)) return res.status(400).json({ error: 'Data inválida.' });
        const d = await dadosVisDpo(r.link.company_id, Number(data.slice(0, 4))), c = colaboradorVisDpo(d, req.body.matricula);
        if (!c) return res.status(404).json({ error: 'Matrícula não encontrada.' });
        const ind = (d.indicadores || []).find(i => i.id === req.body.indId); if (!ind) return res.status(400).json({ error: 'Indicador inválido.' });
        const p = Array.isArray(req.body.porques) ? req.body.porques.map(x => String(x || '').trim().slice(0, 500)) : [];
        if (p.filter(Boolean).length < 3 || !String(req.body.causa || '').trim() || !String(req.body.acao || '').trim()) return res.status(400).json({ error: 'Preencha pelo menos 3 porquês, a causa raiz e a ação.' });
        const dados = { problema: String(req.body.problema || '').slice(0, 500), porques: p.slice(0, 5), causa: String(req.body.causa).trim().slice(0, 500), acao: String(req.body.acao).trim().slice(0, 500), prazo: String(req.body.prazo || '').slice(0, 10), valor: req.body.valor, meta: req.body.meta, nome: c.nome };
        await new Promise((resolve, reject) => db.run(`INSERT INTO dpo_vis_just (company_id, data, matricula, ind_id, dados, status) VALUES (?, ?, ?, ?, ?, 'Enviada')
            ON CONFLICT(company_id, data, matricula, ind_id) DO UPDATE SET dados = excluded.dados, status = 'Enviada', created_at = CURRENT_TIMESTAMP`, [r.link.company_id, data, c.matricula, ind.id, JSON.stringify(dados)], e => e ? reject(e) : resolve()));
        res.json({ message: 'Relato de anomalia enviado!' });
    } catch (e) { res.status(400).json({ error: 'Erro ao enviar o relato.' }); }
});

// Ronda de manutenção: link público para tirar as fotos no celular e subir direto no checklist.
app.post('/api/dpo/manutencao/ronda-link', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const f = await resolverFerramentaDigitalDpo(req, res, 'manutencao', req.body.company_id);
        if (!f) return;
        const ano = anoValidoDpo(req.body.ano), tri = Math.max(0, Math.min(3, Number(req.body.trimestre) || 0));
        let reg = await dbGet(`SELECT token FROM dpo_ronda_links WHERE company_id = ? AND ano = ? AND trimestre = ?`, [f.companyId, ano, tri]);
        if (!reg) {
            const token = crypto.randomBytes(16).toString('hex');
            await new Promise((resolve, reject) => db.run(`INSERT INTO dpo_ronda_links (token, company_id, ano, trimestre, criado_por) VALUES (?, ?, ?, ?, ?)`, [token, f.companyId, ano, tri, req.user.userId], (err) => err ? reject(err) : resolve()));
            reg = { token };
        }
        res.json({ url: `${baseUrlPublicaDpo(req)}/checklist-global.html?t=${reg.token}`, token: reg.token });
    } catch (e) { res.status(400).json({ error: 'Erro ao gerar o link da ronda.' }); }
});
app.get('/api/dpo/manutencao/fotos', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const f = await resolverFerramentaDigitalDpo(req, res, 'manutencao', req.query.company_id);
        if (!f) return;
        res.json(await dbAll(`SELECT * FROM dpo_ronda_fotos WHERE company_id = ? AND ano = ? ORDER BY created_at DESC`, [f.companyId, anoValidoDpo(req.query.ano)]));
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar as fotos.' }); }
});
app.delete('/api/dpo/manutencao/fotos/:id', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const foto = await dbGet(`SELECT * FROM dpo_ronda_fotos WHERE id = ?`, [req.params.id]);
        if (!foto) return res.status(404).json({ error: 'Foto não encontrada.' });
        const f = await resolverFerramentaDigitalDpo(req, res, 'manutencao', foto.company_id);
        if (!f || String(f.companyId) !== String(foto.company_id)) { if (!res.headersSent) res.status(403).json({ error: 'Sem acesso.' }); return; }
        await new Promise((resolve) => db.run(`DELETE FROM dpo_ronda_fotos WHERE id = ?`, [foto.id], () => resolve()));
        res.json({ message: 'Foto removida.' });
    } catch (e) { res.status(400).json({ error: 'Erro ao remover a foto.' }); }
});
async function rondaPorTokenDpo(token) {
    const link = await dbGet(`SELECT * FROM dpo_ronda_links WHERE token = ?`, [String(token || '')]);
    if (!link) return null;
    const empresa = await dbGet(`SELECT name, logo_url FROM companies WHERE id = ?`, [link.company_id]);
    const reg = await dbGet(`SELECT dados FROM dpo_ferramentas_digitais WHERE company_id = ? AND chave = 'manutencao' AND ano = ?`, [link.company_id, link.ano]);
    const dados = reg ? JSON.parse(reg.dados || '{}') : {};
    // nota mais recente entre a tela (dados.notasEm) e o celular (dpo_ronda_notas)
    const notas = { ...(((dados.notas || {})[link.trimestre]) || {}) }, em = ((dados.notasEm || {})[link.trimestre]) || {};
    const cel = await dbAll(`SELECT item, nota, updated_at FROM dpo_ronda_notas WHERE company_id = ? AND ano = ? AND trimestre = ?`, [link.company_id, link.ano, link.trimestre]);
    cel.forEach(c => { if (!em[c.item] || String(c.updated_at) > String(em[c.item])) notas[c.item] = c.nota; });
    return { link, empresa, modelo: dados.modelo || MODELO_MANUTENCAO_DPO, notas };
}
app.get('/api/public/ronda/:token', async (req, res) => {
    try {
        const r = await rondaPorTokenDpo(req.params.token);
        if (!r) return res.status(404).json({ error: 'Link de ronda inválido.' });
        const fotos = await dbAll(`SELECT item, url, obs, autor, created_at FROM dpo_ronda_fotos WHERE company_id = ? AND ano = ? AND trimestre = ? ORDER BY created_at DESC`, [r.link.company_id, r.link.ano, r.link.trimestre]);
        res.json({ empresa: r.empresa ? r.empresa.name : '', logo: r.empresa ? r.empresa.logo_url : null, ano: r.link.ano, trimestre: r.link.trimestre + 1,
            secoes: r.modelo.map(s => ({ titulo: s.titulo, grupos: s.grupos.map(g => ({ numero: g.numero, titulo: g.titulo, itens: g.itens.map(i => ({ id: i.id, num: i.num, texto: i.texto, verificacao: i.verificacao, pontos: i.pontos, peso: i.peso, foto: !!i.foto, critico: !!i.critico, fotoIdeal: i.fotoIdeal || null, nota: r.notas[i.id] === undefined ? '' : String(r.notas[i.id]) })) })) })),
            fotos });
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar a ronda.' }); }
});
app.post('/api/public/ronda/:token/nota', async (req, res) => {
    try {
        const r = await rondaPorTokenDpo(req.params.token);
        if (!r) return res.status(404).json({ error: 'Link do checklist inválido.' });
        const item = String(req.body.item || '').slice(0, 40), nota = String(req.body.nota ?? '');
        if (!['3', '1', '0', 'na', ''].includes(nota)) return res.status(400).json({ error: 'Nota inválida.' });
        if (!r.modelo.some(s => s.grupos.some(g => g.itens.some(i => i.id === item)))) return res.status(400).json({ error: 'Item inválido.' });
        const quando = new Date().toISOString();
        await new Promise((resolve, reject) => db.run(`INSERT INTO dpo_ronda_notas (company_id, ano, trimestre, item, nota, autor, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(company_id, ano, trimestre, item) DO UPDATE SET nota = excluded.nota, autor = excluded.autor, updated_at = excluded.updated_at`,
            [r.link.company_id, r.link.ano, r.link.trimestre, item, nota, String(req.body.autor || '').slice(0, 80), quando], e => e ? reject(e) : resolve()));
        res.json({ message: 'Nota registrada!', nota, updated_at: quando });
    } catch (e) { res.status(400).json({ error: 'Erro ao salvar a nota.' }); }
});
app.get('/api/dpo/manutencao/ronda-notas', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const f = await resolverFerramentaDigitalDpo(req, res, 'manutencao', req.query.company_id);
        if (!f) return;
        res.json(await dbAll(`SELECT trimestre, item, nota, autor, updated_at FROM dpo_ronda_notas WHERE company_id = ? AND ano = ?`, [f.companyId, anoValidoDpo(req.query.ano)]));
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar as notas do celular.' }); }
});
app.post('/api/public/ronda/:token/foto', (req, res) => {
    uploadMaterialDpo.single('file')(req, res, async (err) => {
        if (err) return res.status(400).json({ error: err.code === 'LIMIT_FILE_SIZE' ? 'Foto muito grande.' : err.message });
        if (!req.file) return res.status(400).json({ error: 'Nenhuma foto recebida.' });
        if (!/^image\//.test(req.file.mimetype || '') && !/\.(jpe?g|png|heic|webp)$/i.test(req.file.originalname || '')) return res.status(400).json({ error: 'Envie uma imagem.' });
        try {
            const r = await rondaPorTokenDpo(req.params.token);
            if (!r) return res.status(404).json({ error: 'Link de ronda inválido.' });
            const item = String(req.body.item || '').slice(0, 40);
            const existe = r.modelo.some(s => s.grupos.some(g => g.itens.some(i => i.id === item)));
            if (!existe) return res.status(400).json({ error: 'Item inválido.' });
            await new Promise((resolve, reject) => db.run(`INSERT INTO dpo_ronda_fotos (company_id, ano, trimestre, item, url, obs, autor) VALUES (?, ?, ?, ?, ?, ?, ?)`,
                [r.link.company_id, r.link.ano, r.link.trimestre, item, '/uploads/' + req.file.filename, String(req.body.obs || '').slice(0, 500), String(req.body.autor || '').slice(0, 80)], (e) => e ? reject(e) : resolve()));
            res.json({ message: 'Foto enviada!', url: '/uploads/' + req.file.filename });
        } catch (e) { res.status(400).json({ error: 'Erro ao salvar a foto.' }); }
    });
});

// Rótulos dos campos digitados em cada simulador (para o Excel).
const CAMPOS_SIM_DPO = {
    sim_entrega: {
        plan: [['volume_hl', 'Volume (hl)'], ['volume_mktp_hl', 'Volume Marketplace (hl)'], ['dias_uteis', 'Dias úteis'], ['dias_extra', 'Feriados/domingos trabalhados'], ['ff_ativa', 'Frota fixa ativa'], ['ff_parada', 'Frota parada'], ['cap_media_cx', 'Capacidade média (cx)']],
        real: [['volume_hl', 'Volume (hl)'], ['volume_mktp_hl', 'Volume Marketplace (hl)'], ['dias_uteis', 'Dias trabalhados'], ['ff_ativa', 'Frota fixa ativa'], ['viagens_ff', 'Viagens frota fixa'], ['viagens_spot', 'Viagens spot'], ['km_rodado', 'Km rodado'], ['motoristas', 'Motoristas'], ['ajudantes', 'Ajudantes'], ['custo_fixo', 'Custo fixo frota (R$)'], ['custo_pessoal', 'Custo pessoal (R$)'], ['custo_variavel', 'Custo variável (R$)'], ['custo_spot', 'Custo spot (R$)'], ['custo_outros', 'Outros custos (R$)']]
    },
    sim_armazem: {
        plan: [['volume_hl', 'Volume (hl)'], ['volume_mktp_hl', 'Volume Marketplace (hl)'], ['dias_trab', 'Dias trabalhados'], ['viagens_rota', 'Viagens de rota (carregamento)'], ['viagens_puxada', 'Viagens de puxada (descarga)'], ['qlp_extra', 'QLP extra (limpeza, picking, repack)'], ['emp_alugadas', 'Empilhadeiras alugadas'], ['custo_prejuizos', 'Prejuízos (R$)'], ['custo_outros', 'Outros custos (R$)']],
        real: [['volume_hl', 'Volume (hl)'], ['volume_mktp_hl', 'Volume Marketplace (hl)'], ['dias_trab', 'Dias trabalhados'], ['viagens_rota', 'Viagens de rota'], ['viagens_puxada', 'Viagens de puxada'], ['empilhadeiras', 'Empilhadeiras'], ['horimetro', 'Horímetro (h)'], ['operadores', 'Operadores'], ['conferentes', 'Conferentes'], ['ajudantes', 'Ajudantes'], ['qlp_extra', 'QLP extra'], ['qlp_adm', 'QLP administrativo'], ['custo_empilhadeiras', 'Custo empilhadeiras (R$)'], ['custo_pessoal', 'Custo pessoal (R$)'], ['custo_prejuizos', 'Prejuízos (R$)'], ['custo_outros', 'Outros custos (R$)']]
    },
    sim_puxada: {
        plan: [['volume_hl', 'Volume puxado (hl)'], ['volume_mktp_hl', 'Volume Marketplace (hl)'], ['dias', 'Dias de operação'], ['carretas_ativas', 'Carretas próprias ativas']],
        real: [['volume_hl', 'Volume puxado (hl)'], ['volume_mktp_hl', 'Volume Marketplace (hl)'], ['dias', 'Dias de operação'], ['carretas_ativas', 'Carretas ativas'], ['viagens_programadas', 'Viagens programadas'], ['viagens_ff', 'Viagens frota própria'], ['viagens_spot', 'Viagens spot'], ['viagens_furadas', 'Viagens furadas'], ['motoristas', 'Motoristas'], ['km_rodado', 'Km rodado'], ['custo_fixo', 'Custo fixo carretas (R$)'], ['custo_pessoal', 'Custo pessoal (R$)'], ['custo_variavel', 'Custo variável (R$)'], ['custo_spot', 'Custo spot (R$)']]
    }
};

app.get('/api/dpo/ferramentas-digitais/:chave/export', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const chave = req.params.chave;
        const f = await resolverFerramentaDigitalDpo(req, res, chave, req.query.company_id);
        if (!f) return;
        const ano = anoValidoDpo(req.query.ano);
        const { dados } = await carregarFerramentaDigitalDpo(f.companyId, chave, ano);
        const empresa = await dbGet(`SELECT name FROM companies WHERE id = ?`, [f.companyId]);
        const wb = new ExcelJS.Workbook();
        wb.creator = 'Impulsionar V4';
        const add = (nome, colunas, linhas) => {
            const sh = wb.addWorksheet(nome, { views: [{ state: 'frozen', ySplit: 1 }] });
            sh.columns = colunas.map(([header, key, width]) => ({ header, key, width }));
            estilizarCabecalhoExcelDpo(sh, sh.getColumn(colunas.length).letter);
            (linhas || []).forEach(l => { sh.addRow(l).alignment = { vertical: 'top', wrapText: true }; });
            return sh;
        };
        const val = (await validacaoDaChaveDpo(f.companyId, chave, ano, dados)) || (dados.__resumo ? { itens: [], notaSugerida: dados.__resumo.nota, regra: 'Nota sugerida calculada pelos blocos do acompanhamento' } : { itens: [], notaSugerida: '—', regra: '' });
        if (String(chave).startsWith('acomp:')) {
            const t = f.template;
            add('Itens da verificação', [['Item', 'n', 8], ['Verificação', 't', 70], ['Blocos', 'b', 30], ['Autoavaliação', 's', 16], ['Evidência', 'e', 50]],
                t.itens.map(i => ({ n: i.numero, t: i.texto, b: i.blocos.join(', '), s: ((dados.itens || {})[i.numero] || {}).status || '', e: ((dados.itens || {})[i.numero] || {}).evidencia || '' })));
            ['reunioes', 'realizacoes', 'processos', 'padroes', 'treinamentos', 'ocorrencias', 'riscos', 'inspecoes', 'acoes'].forEach(b => {
                const l = (dados[b] || []).filter(x => x && Object.keys(x).length > 1);
                if (!l.length) return;
                const cols = [...new Set(l.flatMap(x => Object.keys(x)))].filter(k => k !== 'id');
                add(b.charAt(0).toUpperCase() + b.slice(1), cols.map(c => [c, c, 22]), l);
            });
            if ((dados.indicadores || []).length) add('Indicadores', [['Indicador', 'nome', 30], ['Unidade', 'unidade', 10], ['Meta', 'meta', 10], ['Sentido', 'sentido', 12], ...MESES_CURTOS_DPO.map((m, i) => [m, 'm' + i, 10])],
                dados.indicadores.map(i => ({ ...i, ...Object.fromEntries(MESES_CURTOS_DPO.map((m, k) => ['m' + k, numDpo((i.valores || [])[k])])) })));
        } else if (chave === 'swot') {
            const porArea = itensSwotPorAreaDpo(dados);
            const linhas = [];
            Object.entries(porArea).forEach(([area, qs]) => Object.keys(QUADRANTES_SWOT_DPO).forEach(q => (qs[q] || []).forEach(i => linhas.push({ area, q: QUADRANTES_SWOT_DPO[q].rotulo, item: i.texto, c1: i.c1, c2: i.c2, c3: i.c3, pts: pontuacaoItemSwotDpo(q, i) }))));
            linhas.sort((a, b) => a.area.localeCompare(b.area) || a.q.localeCompare(b.q) || (b.pts || 0) - (a.pts || 0));
            add('Itens SWOT por área', [['Área', 'area', 16], ['Fator', 'q', 14], ['Item', 'item', 50], ['Importância', 'c1', 20], ['Intensidade/Urgência', 'c2', 20], ['Tendência', 'c3', 16], ['Pontuação', 'pts', 11]], linhas);
            add('Objetivos', [['Prioridade', 'prioridade', 10], ['Objetivo estratégico', 'texto', 50], ['Obstáculo ao Sonho', 'obstaculo', 40], ['Item SWOT relacionado', 'itemRelacionado', 40], ['Desdobramento', 'desdobramento', 18]], dados.objetivos);
            add('Cruzamentos', [['Força/Fraqueza', 'interno', 45], ['Oportunidade/Ameaça', 'externo', 45], ['Estratégia', 'estrategia', 22]], dados.cruzamentos);
            add('Planos de ação', [['O quê', 'oque', 45], ['Fator', 'fator', 14], ['Item', 'item', 35], ['Responsável', 'responsavel', 22], ['Área', 'area', 16], ['Início', 'inicio', 12], ['Fim', 'fim', 12], ['Andamento', 'andamento', 14], ['Desdobramento', 'desdobramento', 18]], dados.planos);
        } else if (chave === 'gop') {
            exportarGopDpo(add, dados);
        } else if (chave === 'cinco_s') {
            exportar5sDpo(add, dados);
        } else if (TOOLS_EXCLUSIVAS_DPO[chave]) {
            exportarExclusivaDpo(chave, add, dados);
        } else if (chave === 'orcamento') {
            add('RACI', [['Pacote orçamentário', 'pacote', 30], ['Área', 'area', 16], ['R - Responsável', 'r', 22], ['A - Aprovador', 'a', 22], ['C - Consultado', 'c', 22], ['I - Informado', 'i', 22], ['KPI / resultado', 'kpi', 28]], dados.raci);
            add('KPIs sustentabilidade', [['KPI', 'nome', 30], ['Unidade', 'unidade', 12], ['Meta', 'meta', 14], ['Ação / pacote ligado', 'acao', 40]], dados.kpis);
            add('PDCA custos', [['O quê', 'oque', 40], ['Como', 'como', 36], ['Resultado esperado', 'resultado', 30], ['Responsável', 'responsavel', 20], ['Área', 'area', 14], ['Início', 'inicio', 12], ['Fim', 'fim', 12], ['Status', 'status', 18]], dados.pdca);
        } else {
            const cfg = CAMPOS_SIM_DPO[chave];
            add('Parâmetros', [['Parâmetro', 'k', 40], ['Valor', 'v', 16]], Object.entries(dados.params || {}).map(([k, v]) => ({ k, v })));
            ['plan', 'real'].forEach(cen => {
                const extras = Object.fromEntries((dados['custos_extra_' + cen] || []).filter(x => x && x.id).map(x => ['x_' + x.id, 'Custo criado: ' + x.nome]));
                const conhecidos = new Set(cfg[cen].map(c => c[0]));
                const outros = [...new Set((dados[cen] || []).flatMap(m => Object.keys(m || {})))].filter(k => !conhecidos.has(k));
                const rotular = k => extras[k] || ('OBZ — ' + k.replace(/^(o_|custo_)/, '').replace(/_/g, ' ').replace(/^./, c => c.toUpperCase()) + ' (R$)');
                const campos = [...cfg[cen], ...outros.map(k => [k, rotular(k)])];
                const linhas = campos.map(([campo, rot]) => {
                    const l = { ind: rot };
                    MESES_CURTOS_DPO.forEach((m, i) => { l['m' + i] = numDpo(((dados[cen] || [])[i] || {})[campo]); });
                    return l;
                });
                add(cen === 'plan' ? 'Orçamento (Plan)' : 'Realizado', [['Indicador', 'ind', 36], ...MESES_CURTOS_DPO.map((m, i) => [m, 'm' + i, 11])], linhas);
            });
        }
        if (TOOLS_EXCLUSIVAS_DPO[chave] && dados.__validacao) add('Validação checklist', [['Item', 'n', 8], ['Status', 's', 12], ['Como a ferramenta avaliou', 't', 90]], dados.__validacao.map(v => ({ n: v.v, s: v.st, t: v.txt })));
        if (chave !== 'gop' && chave !== 'cinco_s' && !TOOLS_EXCLUSIVAS_DPO[chave]) add('Validação checklist', [['Item', 'n', 8], ['Verificação', 't', 60], ['Atendido', 'ok', 10], ['O que falta', 'f', 70]],
            [...val.itens.map(i => ({ n: i.numero, t: i.texto, ok: i.ok ? 'Sim' : 'Não', f: i.faltas.join(' | ') })), { n: 'Nota', t: 'Nota sugerida pela ferramenta', ok: val.notaSugerida, f: val.regra }]);
        const buffer = await wb.xlsx.writeBuffer();
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${chave}-${ano}-${(empresa ? empresa.name : 'empresa').replace(/[^a-z0-9]+/gi, '-')}.xlsx"`);
        res.send(Buffer.from(buffer));
    } catch (e) {
        console.error('Erro ao exportar ferramenta digital DPO:', e.message);
        res.status(500).json({ error: 'Erro ao exportar.' });
    }
});

// ======================================================================
// DPO — ACOMPANHAMENTOS IMPULSIONAR (um por pergunta do checklist)
// Montados automaticamente a partir dos itens da verificação (V.1, V.2...):
// cada item vira blocos de acompanhamento (reuniões, indicadores, padrões,
// treinamentos, processos, ocorrências, riscos, inspeções, plano de ação).
// Liberados pelo Master por REVENDA e por PERGUNTA. Salvos por ano.
// ======================================================================
const BLOCOS_ACOMP_DPO = {
    reunioes: /reuni|comit[êe]|\brps\b|matinal|kick ?off|\bdds\b|di[áa]logo di[áa]rio|f[óo]rum|rotina de reuni/i,
    indicadores: /indicador|\bkpi|\bmeta\b|\bmetas\b|%|[íi]ndice|\btaxa|tempo m[ée]dio|\bytd\b|[úu]ltimos \d|tend[êe]ncia|dispers|\bnps\b|turnover|rotatividade|absente[íi]smo|produtividade|r\$\/hl|acuracidade|n[íi]vel de servi/i,
    padroes: /padr[ãa]o|padr[õo]es|procedimento|\bpop\b|\bsop\b|instru[çc][ãa]o de trabalho|pol[íi]tica/i,
    treinamentos: /treina|capacita|reciclag|qualifica[çc]/i,
    processos: /existe (um )?processo|processo (em vigor|formal|estruturado|definido)|sistem[áa]tica|cronograma|fluxo |respons[áa]veis (claros|definidos)|crit[ée]rios definidos|rotina (de|estruturada|definida)/i,
    acoes: /plano de a[çc]|planos de a[çc]|\bpdca\b|contramedida|tratativa|a[çc][õo]es corretivas|a[çc][õo]es preventivas/i,
    ocorrencias: /acidente|incidente|ocorr[êe]ncia|investiga|quase.acidente|\bdesvios?\b|reclama[çc]|avaria|sinistro|\blti\b|\bsif\b/i,
    riscos: /\briscos?\b|perigo|\bapr\b|mapa de risco/i,
    inspecoes: /inspe[çc]|auditoria|checklist|check-list|ronda|gemba|blitz|observa[çc][ãa]o comportamental|sinaliza[çc]/i
};
const ORDEM_BLOCOS_ACOMP_DPO = Object.keys(BLOCOS_ACOMP_DPO);
const ACOMP_ESPECIAIS_DPO = { 'entrega:3.1': 'visibilidade', 'planejamento:2.1': 'riscos', 'gestao:1.3': 'swot', 'planejamento:1.1': 'dimensionamento', 'gestao:4.6': 'gop', 'gestao:3.1': 'cinco_s', 'gestao:1.2': 'sonho', 'planejamento:2.2': 'manutencao', 'planejamento:2.3': 'p3a', 'planejamento:2.4': 'capex', 'planejamento:3.2': 'ans' };
const KPIS_CONHECIDOS_DPO = ['TML', 'TMA', 'NPS', 'eNPS', 'OTIF', 'LTI', 'MDI', 'MTI', 'SIF', 'TRI', 'TRIFR', 'PNP', 'FNP', 'VMI', 'EFC', 'OEE', 'DPMO', 'IRL', 'GPS'];

function sugestoesAcompDpo(texto) {
    const t = String(texto || '');
    const conectores = /\s+(de|do|da|dos|das|e|a|o|ao|aos|no|na|com|para|que|é|são|está|estão|foi|deve|inclui|rastreada|revisad[ao]|executad[ao]|\(|-)$/i;
    const limpar = s => { let x = s.replace(/\s+/g, ' ').trim().split(/\s+(?:é|são|está|estão|foi|deve|inclui|rastread[ao]|revisad[ao]|executad[ao]|com|para|que|onde)\s/i)[0]; x = x.split(' ').slice(0, 7).join(' '); while (conectores.test(x)) x = x.replace(conectores, ''); return x; };
    const uniq = arr => [...new Set(arr.map(limpar).filter(s => s.length > 2))];
    const nome = '[A-ZÀ-Ú][\\wÀ-ú]*(?:\\s+(?:de|do|da|dos|das|e|[A-ZÀ-Ú][\\wÀ-ú]*)){0,5}';
    const reunioes = uniq([
        ...(t.match(/Comit[êe] de [A-ZÀ-Úa-zà-ú]+(?: [A-ZÀ-Ú][a-zà-ú]+)?/g) || []),
        ...(t.match(/\b(RPS|Supermatinal|Kick ?off|Kickoff|DDS|Matinal|Team Room|Semana d[ae] [A-ZÀ-Ú][a-zà-ú]+)\b/g) || []),
        ...(t.match(/Reuni[ãa]o (?:de|do|da|mensal|semanal|di[áa]ria) [a-zà-ú ]{3,28}/gi) || [])
    ]).slice(0, 6);
    const kpis = uniq([
        ...KPIS_CONHECIDOS_DPO.filter(k => new RegExp('\\b' + k + '\\b').test(t)),
        ...(t.match(new RegExp('Tempo M[ée]dio de ' + nome, 'g')) || []),
        ...(t.match(/(?:[ÍI]ndice|Taxa|% de) [a-zà-ú][a-zà-ú ]{3,45}/gi) || []),
        ...(/absente[íi]smo/i.test(t) ? ['Absenteísmo'] : []), ...(/turnover|rotatividade/i.test(t) ? ['Turnover'] : []),
        ...(/acuracidade/i.test(t) ? ['Acuracidade de estoque'] : []), ...(/n[íi]vel de servi/i.test(t) ? ['Nível de serviço'] : []),
        ...(/devolu/i.test(t) ? ['% Devolução'] : []), ...(/produtividade/i.test(t) ? ['Produtividade'] : [])
    ]).slice(0, 6);
    const padroes = uniq(t.match(new RegExp('Padr[ãa]o (?:de |do |da )?' + nome, 'g')) || []).slice(0, 6);
    return { reunioes, indicadores: kpis, padroes };
}

function montarTemplateAcompDpo(pilarKey, numero) {
    const achou = perguntaDoPilarDpo(pilarKey, numero);
    if (!achou) return null;
    const q = achou.pergunta;
    const itens = itensDaVerificacaoDpo(q.verificacao).map(it => ({
        numero: 'V.' + it.numero, texto: it.texto,
        blocos: ORDEM_BLOCOS_ACOMP_DPO.filter(b => BLOCOS_ACOMP_DPO[b].test(it.texto))
    }));
    const blocos = ORDEM_BLOCOS_ACOMP_DPO.filter(b => itens.some(i => i.blocos.includes(b)));
    if (!blocos.includes('acoes')) blocos.push('acoes'); // todo acompanhamento tem plano de ação
    // Regra da nota 1: itens citados antes do "MAS" na explicação da pontuação 1.
    const linha1 = String(q.explicacao_pontos || '').split(/\n\s*\n/).find(l => /^\s*1\s*[-.:)]/.test(l)) || '';
    const req1 = [...new Set((linha1.split(/\bMAS\b/i)[0].match(/V\.?\s*\d+/gi) || []).map(v => 'V.' + v.replace(/\D/g, '')))];
    return {
        chave: `acomp:${pilarKey}:${numero}`, pilar: pilarKey, pilarLabel: DPO_AMBEV_DATA[pilarKey].label, grupo: `${achou.grupo.numero} ${achou.grupo.titulo}`,
        numero, questao: q.questao, mandatoria: !!q.mandatoria, verificacao: q.verificacao || '', how_to_check: q.how_to_check || '', explicacao: q.explicacao_pontos || '',
        itens, blocos, sugestoes: sugestoesAcompDpo(q.verificacao + '\n' + (q.how_to_check || '')), regra: { req1 }, especial: ACOMP_ESPECIAIS_DPO[`${pilarKey}:${numero}`] || null
    };
}

async function acompLiberadoDpo(companyId, questionKey) {
    return !!(await dbGet(`SELECT 1 FROM dpo_acomp_liberacoes WHERE company_id = ? AND question_key = ?`, [companyId, questionKey]));
}

// Catálogo (Master): todas as perguntas com os blocos que o acompanhamento terá.
app.get('/api/admin/dpo/acompanhamentos/catalogo', requireRole('admin'), (req, res) => {
    res.json(DPO_PILARES_ORDEM.map(k => ({
        key: k, label: DPO_AMBEV_DATA[k].label, numero: DPO_PILARES_ORDEM.indexOf(k) + 1,
        perguntas: DPO_AMBEV_DATA[k].grupos.flatMap(g => g.perguntas.map(q => {
            const t = montarTemplateAcompDpo(k, q.numero);
            return { questionKey: `${k}:${q.numero}`, numero: q.numero, questao: q.questao, grupo: `${g.numero} ${g.titulo}`, mandatoria: !!q.mandatoria, blocos: t.blocos, itens: t.itens.length, especial: t.especial };
        }))
    })));
});

async function resumosAcompDpo(companyId) {
    const regs = await dbAll(`SELECT chave, ano, dados, updated_at FROM dpo_ferramentas_digitais WHERE company_id = ? AND chave LIKE 'acomp:%' ORDER BY ano DESC`, [companyId]);
    const porChave = {};
    regs.forEach(r => { if (porChave[r.chave]) return; const d = JSON.parse(r.dados || '{}'); porChave[r.chave] = { ano: r.ano, updated_at: r.updated_at, resumo: d.__resumo || null }; });
    // ferramentas exclusivas: o resumo da pergunta vem da própria ferramenta
    const ex = await dbAll(`SELECT chave, ano, dados, updated_at FROM dpo_ferramentas_digitais WHERE company_id = ? AND chave IN (${Object.keys(TOOLS_EXCLUSIVAS_DPO).map(() => '?').join(',')}) ORDER BY ano DESC`, [companyId, ...Object.keys(TOOLS_EXCLUSIVAS_DPO)]);
    ex.forEach(r => { const k = 'acomp:' + TOOLS_EXCLUSIVAS_DPO[r.chave]; if (porChave[k] && porChave[k].exclusiva) return; const d = JSON.parse(r.dados || '{}'); porChave[k] = { ano: r.ano, updated_at: r.updated_at, resumo: d.__resumo || null, exclusiva: r.chave }; });
    return porChave;
}

app.get('/api/admin/dpo/acompanhamentos/liberacoes', requireRole('admin'), async (req, res) => {
    try {
        if (!req.query.company_id) return res.status(400).json({ error: 'Informe a revenda.' });
        const libs = await dbAll(`SELECT l.question_key, l.liberado_em, u.name as liberadoPor FROM dpo_acomp_liberacoes l LEFT JOIN users u ON u.id = l.liberado_por WHERE l.company_id = ?`, [req.query.company_id]);
        res.json({ liberacoes: libs, resumos: await resumosAcompDpo(req.query.company_id), pilaresAtivos: await pilaresAtivosDaEmpresa(req.query.company_id) });
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar as liberações.' }); }
});

app.put('/api/admin/dpo/acompanhamentos/liberacoes', requireRole('admin'), async (req, res) => {
    const { company_id, keys, liberado } = req.body;
    if (!company_id || !Array.isArray(keys) || !keys.length) return res.status(400).json({ error: 'Informe a revenda e as perguntas.' });
    try {
        const empresa = await dbGet(`SELECT id FROM companies WHERE id = ?`, [company_id]);
        if (!empresa) return res.status(404).json({ error: 'Revenda não encontrada.' });
        const validas = keys.filter(k => { const [p, n] = String(k).split(':'); return !!perguntaDoPilarDpo(p, n); });
        for (const k of validas) {
            if (liberado) await new Promise((resolve) => db.run(`INSERT OR IGNORE INTO dpo_acomp_liberacoes (company_id, question_key, liberado_por) VALUES (?, ?, ?)`, [company_id, k, req.user.userId], () => resolve()));
            else await new Promise((resolve) => db.run(`DELETE FROM dpo_acomp_liberacoes WHERE company_id = ? AND question_key = ?`, [company_id, k], () => resolve()));
        }
        if (liberado && validas.length) {
            const nomes = validas.slice(0, 3).map(k => { const [p, n] = k.split(':'); return `${DPO_AMBEV_DATA[p].label} ${n}`; }).join(', ');
            notificarPorCompanyAdmins(company_id, 'DPO — novo Acompanhamento Impulsionar liberado', `${validas.length} acompanhamento(s) liberado(s): ${nomes}${validas.length > 3 ? '...' : ''}.`, 'dpoHome');
        }
        res.json({ message: liberado ? `${validas.length} acompanhamento(s) liberado(s)!` : `${validas.length} acompanhamento(s) bloqueado(s).` });
    } catch (e) { res.status(400).json({ error: 'Erro ao atualizar as liberações.' }); }
});

// Empresa: acompanhamentos liberados para ela (Master pode consultar com company_id).
app.get('/api/dpo/acompanhamentos', requireRole('admin', 'client_admin'), async (req, res) => {
    try {
        const companyId = req.user.role === 'client_admin' ? req.user.companyId : req.query.company_id;
        if (!companyId) return res.status(400).json({ error: 'Informe a empresa.' });
        const ativos = await pilaresAtivosDaEmpresa(companyId);
        const libs = await dbAll(`SELECT question_key, liberado_em FROM dpo_acomp_liberacoes WHERE company_id = ?`, [companyId]);
        const resumos = await resumosAcompDpo(companyId);
        const lista = libs.map(l => {
            const [p, n] = l.question_key.split(':');
            if (req.user.role === 'client_admin' && !ativos.includes(p)) return null;
            const t = montarTemplateAcompDpo(p, n);
            if (!t) return null;
            const r = resumos[`acomp:${l.question_key}`];
            return { questionKey: l.question_key, pilar: p, pilarLabel: t.pilarLabel, numero: n, questao: t.questao, mandatoria: t.mandatoria, blocos: t.blocos, especial: t.especial, liberado_em: l.liberado_em, resumo: r ? r.resumo : null, ano: r ? r.ano : null, updated_at: r ? r.updated_at : null };
        }).filter(Boolean).sort((a, b) => DPO_PILARES_ORDEM.indexOf(a.pilar) - DPO_PILARES_ORDEM.indexOf(b.pilar) || String(a.numero).localeCompare(String(b.numero), undefined, { numeric: true }));
        res.json(lista);
    } catch (e) { res.status(500).json({ error: 'Erro ao carregar os acompanhamentos.' }); }
});

app.put('/api/dpo/cycles/:id/answers', requireRole('admin', 'client_admin'), async (req, res) => {
    const { questionKey, score } = req.body;
    if (!questionKey) return res.status(400).json({ error: 'Informe a pergunta.' });
    try {
        const ciclo = await obterCicloComAcesso(req, res, req.params.id);
        if (!ciclo) return;
        if (ciclo.status === 'concluido') return res.status(400).json({ error: 'Este ciclo já foi encerrado e não pode mais ser editado.' });
        if (!(await empresaTemPastaDpo(req, res, ciclo.company_id, 'checklist'))) return;
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
const ROTULOS_STATUS_PLANO_DPO_SERVIDOR = { nao_iniciada: 'Não iniciada', em_andamento: 'Em andamento', concluida: 'Concluída' };

// Acha o texto da pergunta (dado estático) a partir da question_key salva no
// plano/resposta — usado nas exportações em Excel, que precisam do texto por
// fora do JSON já montado pra tela.
function textoDaPerguntaDpo(pilarKey, numeroPergunta) {
    const pilarInfo = DPO_AMBEV_DATA[pilarKey];
    if (!pilarInfo) return '';
    for (const g of pilarInfo.grupos) {
        const achou = g.perguntas.find(q => q.numero === numeroPergunta);
        if (achou) return achou.questao;
    }
    return '';
}

app.post('/api/dpo/action-plans', requireRole('admin', 'client_admin'), async (req, res) => {
    const { cycle_id, questionKey, texto, verificacao_numero, owner, status, data_prevista } = req.body;
    if (!cycle_id || !questionKey || !texto) return res.status(400).json({ error: 'Preencha o plano de ação.' });
    const statusFinal = STATUS_PLANO_ACAO_DPO.includes(status) ? status : 'nao_iniciada';
    try {
        const ciclo = await obterCicloComAcesso(req, res, cycle_id);
        if (!ciclo) return;
        if (ciclo.status === 'concluido') return res.status(400).json({ error: 'Este ciclo já foi encerrado.' });
        if (!(await empresaTemPastaDpo(req, res, ciclo.company_id, 'checklist'))) return;
        const resultado = await new Promise((resolve, reject) => db.run(
            `INSERT INTO dpo_action_plans (cycle_id, question_key, texto, created_by, verificacao_numero, owner, status, data_prevista) VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
            [cycle_id, questionKey, texto, req.user.userId, verificacao_numero || null, owner || null, statusFinal, data_prevista || null], function (err) { err ? reject(err) : resolve(this.lastID); }
        ));
        res.json({ message: 'Plano de ação criado!', id: resultado });
    } catch (e) { res.status(400).json({ error: 'Erro ao criar o plano de ação.' }); }
});

app.put('/api/dpo/action-plans/:id', requireRole('admin', 'client_admin'), async (req, res) => {
    const { texto, verificacao_numero, owner, status, data_prevista } = req.body;
    try {
        const plano = await dbGet(`SELECT * FROM dpo_action_plans WHERE id = ?`, [req.params.id]);
        if (!plano) return res.status(404).json({ error: 'Plano de ação não encontrado.' });
        const ciclo = await obterCicloComAcesso(req, res, plano.cycle_id);
        if (!ciclo) return;
        const statusFinal = status !== undefined ? (STATUS_PLANO_ACAO_DPO.includes(status) ? status : plano.status) : plano.status;
        db.run(`UPDATE dpo_action_plans SET texto = ?, verificacao_numero = ?, owner = ?, status = ?, data_prevista = ? WHERE id = ?`,
            [texto !== undefined ? texto : plano.texto, verificacao_numero !== undefined ? verificacao_numero : plano.verificacao_numero, owner !== undefined ? owner : plano.owner, statusFinal, data_prevista !== undefined ? data_prevista : plano.data_prevista, req.params.id],
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

// Correção de texto (ortografia/acentuação/pontuação em pt-BR). Usa a IA quando
// a chave estiver configurada; sem ela, aplica um corretor local de regras.
const CORRECOES_PTBR = { concluida: 'concluída', concluido: 'concluído', concluidas: 'concluídas', concluidos: 'concluídos', saida: 'saída', saidas: 'saídas', ferias: 'férias', obrigatorio: 'obrigatório', obrigatoria: 'obrigatória', necessario: 'necessário', necessaria: 'necessária', inventario: 'inventário', usuario: 'usuário', usuarios: 'usuários', horario: 'horário', horarios: 'horários', semanal: 'semanal', mensal: 'mensal', experiencia: 'experiência', frequencia: 'frequência', sequencia: 'sequência', ocorrencia: 'ocorrência', ocorrencias: 'ocorrências', auditoria: 'auditoria', ciclo: 'ciclo', nao: 'não', voce: 'você', voces: 'vocês', tambem: 'também', ate: 'até', entao: 'então', ja: 'já', sera: 'será', acao: 'ação', acoes: 'ações', avaliacao: 'avaliação', autoavaliacao: 'autoavaliação', validacao: 'validação', correcao: 'correção', reuniao: 'reunião', reunioes: 'reuniões', gestao: 'gestão', producao: 'produção', manutencao: 'manutenção', distribuicao: 'distribuição', informacao: 'informação', informacoes: 'informações', comunicacao: 'comunicação', operacao: 'operação', operacoes: 'operações', area: 'área', areas: 'áreas', analise: 'análise', numero: 'número', proximo: 'próximo', proxima: 'próxima', periodo: 'período', inicio: 'início', conclusao: 'conclusão', armazem: 'armazém', veiculo: 'veículo', veiculos: 'veículos', caminhao: 'caminhão', caminhoes: 'caminhões', padrao: 'padrão', padroes: 'padrões', revisao: 'revisão', implantacao: 'implantação', execucao: 'execução', confeccao: 'confecção', evidencia: 'evidência', evidencias: 'evidências', responsavel: 'responsável', responsaveis: 'responsáveis', funcionario: 'funcionário', funcionarios: 'funcionários', indicador: 'indicador', calendario: 'calendário', relatorio: 'relatório', relatorios: 'relatórios', reuniao: 'reunião', diario: 'diário', diaria: 'diária', mes: 'mês', tres: 'três', sao: 'são', apos: 'após', tambem: 'também', etica: 'ética', anticorrupcao: 'anticorrupção', corrupcao: 'corrupção', seguranca: 'segurança', lideranca: 'liderança', organizacao: 'organização', manutençao: 'manutenção', atraves: 'através', alem: 'além', especifico: 'específico', minimo: 'mínimo', maximo: 'máximo', unico: 'único', otimo: 'ótimo', pratica: 'prática', praticas: 'práticas', tecnico: 'técnico', tecnica: 'técnica', logistica: 'logística', critico: 'crítico', criticos: 'críticos', historico: 'histórico', basico: 'básico', publico: 'público', saude: 'saúde', possivel: 'possível', disponivel: 'disponível', nivel: 'nível', niveis: 'níveis', util: 'útil', facil: 'fácil', dificil: 'difícil', agua: 'água', voce: 'você', porem: 'porém', alguem: 'alguém', ninguem: 'ninguém', tem: 'tem', ultimo: 'último', ultima: 'última', pagina: 'página', estrategia: 'estratégia', cascateamento: 'cascateamento', kpis: 'KPIs', kpi: 'KPI', swot: 'SWOT', dpo: 'DPO', vpo: 'VPO', rh: 'RH', ti: 'TI' };
function corrigirTextoLocal(t) {
    let x = String(t || '').replace(/\r/g, '');
    x = x.replace(/[ \t]+/g, ' ').replace(/ *\n */g, '\n').trim();
    x = x.replace(/\b(anti|pré|pós|pró|ex|vice|auto|semi)-\s+(\p{L})/giu, '$1-$2').replace(/\banti-\s*corrup/giu, 'anticorrup');
    x = x.replace(/\s+([,.;:!?%)])/g, '$1').replace(/([,;:!?])(?=[^\s\d\n)])/g, '$1 ').replace(/\.(?=[A-Za-zÀ-ú]{2})/g, '. ').replace(/\(\s+/g, '(');
    x = x.replace(/\s*,\s*/g, ', ').replace(/\s+\n/g, '\n');
    x = x.replace(/\p{L}+/gu, w => { const k = w.toLowerCase(); const c = CORRECOES_PTBR[k]; if (!c) return w; if (c === c.toUpperCase()) return c; return w[0] === w[0].toUpperCase() ? c[0].toUpperCase() + c.slice(1) : c; });
    x = x.replace(/(^|[.!?]\s+|\n)(\p{Ll})/gu, (m, a, b) => a + b.toUpperCase());
    return x;
}
app.post('/api/corrigir-texto', async (req, res) => {
    const texto = String((req.body || {}).texto || '').slice(0, 4000);
    if (!texto.trim()) return res.status(400).json({ error: 'Escreva o texto primeiro.' });
    if (ANTHROPIC_API_KEY) {
        try {
            const r = await perguntarIA('Você é um revisor de textos em português do Brasil. Corrija ortografia, acentuação, concordância, pontuação e espaçamento, mantendo o sentido, o tom e os termos técnicos/siglas (DPO, KPI, SWOT, 5S, PDV etc.). Não acrescente informação nem explicações. Responda SOMENTE com o texto corrigido.', texto, 900);
            if (r) return res.json({ texto: r.replace(/^["“]|["”]$/g, '').trim(), viaIA: true });
        } catch (e) { /* cai no corretor local */ }
    }
    res.json({ texto: corrigirTextoLocal(texto), viaIA: false });
});
app.post('/api/dpo/action-plans/:id/follow-ups', requireRole('admin', 'client_admin'), async (req, res) => {
    const { texto, data_prevista, foto_url } = req.body;
    if (!texto) return res.status(400).json({ error: 'Descreva o follow-up.' });
    try {
        const plano = await dbGet(`SELECT * FROM dpo_action_plans WHERE id = ?`, [req.params.id]);
        if (!plano) return res.status(404).json({ error: 'Plano de ação não encontrado.' });
        const ciclo = await obterCicloComAcesso(req, res, plano.cycle_id);
        if (!ciclo) return;
        const ultimo = await dbGet(`SELECT MAX(numero) as maximo FROM dpo_follow_ups WHERE action_plan_id = ?`, [req.params.id]);
        const numero = (ultimo && ultimo.maximo) ? ultimo.maximo + 1 : 1;
        const resultado = await new Promise((resolve, reject) => db.run(
            `INSERT INTO dpo_follow_ups (action_plan_id, numero, texto, data_prevista, foto_url, autor) VALUES (?, ?, ?, ?, ?, (SELECT name FROM users WHERE id = ?))`,
            [req.params.id, numero, texto, data_prevista || null, foto_url || null, req.user.userId], function (err) { err ? reject(err) : resolve(this.lastID); }
        ));
        res.json({ message: `Follow ${numero} adicionado!`, id: resultado, numero });
    } catch (e) { res.status(400).json({ error: 'Erro ao adicionar o follow-up.' }); }
});

app.put('/api/dpo/follow-ups/:id', requireRole('admin', 'client_admin'), async (req, res) => {
    const { texto, data_prevista, status, foto_url } = req.body;
    try {
        const follow = await dbGet(`SELECT * FROM dpo_follow_ups WHERE id = ?`, [req.params.id]);
        if (!follow) return res.status(404).json({ error: 'Follow-up não encontrado.' });
        const plano = await dbGet(`SELECT * FROM dpo_action_plans WHERE id = ?`, [follow.action_plan_id]);
        const ciclo = await obterCicloComAcesso(req, res, plano.cycle_id);
        if (!ciclo) return;
        db.run(`UPDATE dpo_follow_ups SET texto = COALESCE(?, texto), data_prevista = COALESCE(?, data_prevista), status = COALESCE(?, status), foto_url = COALESCE(?, foto_url) WHERE id = ?`,
            [texto || null, data_prevista || null, status || null, foto_url || null, req.params.id], () => res.json({ message: 'Follow-up atualizado!' }));
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
    // Controle remoto da apresentação (Ministrar Treinamento).
    socket.on('mt-apresentador', ({ codigo, token } = {}) => {
        try {
            const payload = jwt.verify(String(token || ''), JWT_SECRET);
            if (payload.role !== 'admin' || !CONTROLES_MT.has(String(codigo))) return;
            socket.join('mtc:' + codigo); socket.data.mtc = { codigo: String(codigo), papel: 'apresentador' };
            const n = (io.sockets.adapter.rooms.get('mtc:' + codigo) || new Set()).size - 1;
            socket.emit('mt-controles', { conectados: Math.max(0, n) });
        } catch (e) { /* token inválido */ }
    });
    socket.on('mt-controle-entrar', ({ codigo } = {}) => {
        const c = CONTROLES_MT.get(String(codigo || ''));
        if (!c) return socket.emit('mt-controle-erro', { error: 'Link de controle inválido ou expirado. Gere um novo na apresentação.' });
        socket.join('mtc:' + codigo); socket.data.mtc = { codigo: String(codigo), papel: 'controle' };
        socket.emit('mt-controle-ok', { titulo: c.titulo });
        socket.to('mtc:' + codigo).emit('mt-controle-conectou', {});
    });
    socket.on('mt-comando', ({ cmd, i } = {}) => {
        const m = socket.data.mtc; if (!m || m.papel !== 'controle') return;
        if (!['proximo', 'anterior', 'ir', 'pedir-estado'].includes(cmd)) return;
        socket.to('mtc:' + m.codigo).emit('mt-comando', { cmd, i: Number(i) || 0 });
    });
    socket.on('mt-estado', (estado = {}) => {
        const m = socket.data.mtc; if (!m || m.papel !== 'apresentador') return;
        socket.to('mtc:' + m.codigo).emit('mt-estado', estado);
    });
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