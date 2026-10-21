// Service worker do app Impulsionar V4: deixa o sistema instalável no celular e mostra
// uma tela amigável quando o aparelho está sem internet. Não guarda dados do sistema
// (tudo continua vindo sempre atualizado do servidor).
const VERSAO = 'impulsionar-v1';
self.addEventListener('install', e => { e.waitUntil(caches.open(VERSAO).then(c => c.addAll(['/assets/app-icone-192.png']).catch(() => {}))); self.skipWaiting(); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(l => Promise.all(l.filter(k => k !== VERSAO).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  const r = e.request;
  if (r.method !== 'GET' || r.mode !== 'navigate') return; // API, vídeos e arquivos: direto do servidor
  e.respondWith(fetch(r).catch(() => new Response(`<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Sem internet</title></head>
<body style="margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center;background:#0b1329;color:#fff;font-family:system-ui,Arial,sans-serif;text-align:center;padding:24px;">
<div><img src="/assets/app-icone-192.png" width="96" height="96" style="border-radius:22px" alt=""><h2>Sem conexão com a internet</h2><p style="color:#cbd5e1">Assim que a internet voltar, o Impulsionar abre normalmente.</p>
<button onclick="location.reload()" style="border:none;border-radius:12px;padding:12px 22px;font-weight:800;background:#22c55e;color:#0b1329;font-size:15px">Tentar de novo</button></div></body></html>`, { headers: { 'Content-Type': 'text/html; charset=utf-8' } })));
});
