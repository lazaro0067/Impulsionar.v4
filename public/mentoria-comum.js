/* Comum às páginas públicas da mentoria "Do Chão à Gestão" */
const MT = { dados: null };
const mtEsc = s => String(s === undefined || s === null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const mtMoeda = v => 'R$ ' + Number(v || 0).toLocaleString('pt-BR', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
const mtData = s => { if (!s) return ''; const d = new Date(String(s).length <= 10 ? s + 'T12:00:00' : s); return isNaN(d) ? s : d.toLocaleDateString('pt-BR', { day: '2-digit', month: 'long' }); };
function mtWhats(texto) { const n = String((MT.dados && MT.dados.marca.whatsapp) || '').replace(/\D/g, ''); if (!n) return ''; return 'https://wa.me/' + (n.length <= 11 ? '55' + n : n) + (texto ? '?text=' + encodeURIComponent(texto) : ''); }
function mtIniciais(nome) { return String(nome || 'LS').split(/\s+/).filter(Boolean).slice(0, 2).map(p => p[0]).join('').toUpperCase(); }
async function mtCarregar() {
  const r = await fetch('/api/public/mentoria', { cache: 'no-store' }); if (!r.ok) throw new Error('Não consegui carregar.');
  MT.dados = await r.json(); return MT.dados;
}
function mtTopo(ativo) {
  const m = MT.dados.marca;
  return `<header class="topo"><div class="cx"><a class="marca" href="/bio"><i>${mtEsc(mtIniciais(m.mentor))}</i>${mtEsc(m.mentor)}</a>
    <nav><a href="/mentoria" ${ativo === 'mentoria' ? 'style="color:var(--ambar2)"' : ''}>Mentoria</a><a href="/palestra" ${ativo === 'palestra' ? 'style="color:var(--ambar2)"' : ''}>Palestras</a><a href="/aula" ${ativo === 'aula' ? 'style="color:var(--ambar2)"' : ''}>Aula gratuita</a><a href="/raio-x" ${ativo === 'raio' ? 'style="color:var(--ambar2)"' : ''}>Raio-X do Líder</a></nav>
    ${ativo === 'palestra' ? `<a class="bt ambar peq" href="#contratar">Contratar palestra</a>` : `<a class="bt ambar peq" href="/mentoria#inscricao">Quero participar</a>`}</div></header>`;
}
function mtRodape() {
  const m = MT.dados.marca;
  return `<footer class="rodape"><div class="cx"><div><b style="color:#fff">${mtEsc(m.mentor)}</b> · Consultor ${mtEsc(m.empresa || 'Impulsionar V4')}<br>${mtEsc(m.titulo)} · Mentoria ${mtEsc(m.nome)}<br><a href="https://instagram.com/${mtEsc(String(m.instagramEmpresa || 'impulsionarv4').replace(/^@/, ''))}" target="_blank" rel="noopener">@${mtEsc(String(m.instagramEmpresa || 'impulsionarv4').replace(/^@/, ''))}</a> · <a href="https://instagram.com/${mtEsc(String(m.instagram || 'lazaroneto.oficial').replace(/^@/, ''))}" target="_blank" rel="noopener">@${mtEsc(String(m.instagram || 'lazaroneto.oficial').replace(/^@/, ''))}</a></div>
    <div>${m.instagram ? `<a href="https://instagram.com/${mtEsc(String(m.instagram).replace(/^@/, ''))}" target="_blank" rel="noopener">Instagram</a>` : ''}${m.linkedin ? `<a href="${mtEsc(m.linkedin)}" target="_blank" rel="noopener">LinkedIn</a>` : ''}${mtWhats() ? `<a href="${mtWhats('Olá, Lázaro! Vim pelo site.')}" target="_blank" rel="noopener">WhatsApp</a>` : ''}${m.email ? `<a href="mailto:${mtEsc(m.email)}">${mtEsc(m.email)}</a>` : ''}</div></div></footer>
    ${mtWhats() ? `<a class="flutuante" href="${mtWhats('Olá, Lázaro! Vim pelo site e quero saber mais.')}" target="_blank" rel="noopener" title="Falar no WhatsApp">💬</a>` : ''}`;
}
// envia um formulário (lead, inscrição) e mostra a mensagem
async function mtEnviar(form, url, extra) {
  const bt = form.querySelector('button[type=submit]'), msg = form.querySelector('.msg-area');
  const dados = Object.assign(Object.fromEntries(new FormData(form).entries()), { visitante: mtVisitante() }, extra || {});
  if (bt) { bt.disabled = true; bt.dataset.t = bt.textContent; bt.textContent = 'Enviando…'; }
  if (msg) msg.innerHTML = '';
  try {
    const r = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(dados) });
    const j = await r.json().catch(() => ({}));
    if (!r.ok) throw new Error(j.error || 'Não consegui enviar. Tente de novo.');
    return j;
  } catch (e) { if (msg) msg.innerHTML = `<div class="msg erro">${mtEsc(e.message)}</div>`; throw e; }
  finally { if (bt) { bt.disabled = false; bt.textContent = bt.dataset.t; } }
}
function mtModal(id, on) { const m = document.getElementById(id); if (m) m.classList.toggle('on', on !== false); }
function mtErroPagina(e) { document.body.innerHTML = `<div class="cx" style="padding:80px 20px;text-align:center;"><h2>Página indisponível no momento</h2><p>${mtEsc(e.message)}</p></div>`; }

/* ---------- Rastreio de visitas e cliques (anônimo até a pessoa se cadastrar) ---------- */
function mtVisitante() {
  try { let v = localStorage.getItem('mt_v'); if (!v || !/^[a-z0-9]{8,40}$/i.test(v)) { v = (Date.now().toString(36) + Math.random().toString(36).slice(2, 12)).slice(0, 24); localStorage.setItem('mt_v', v); } return v; } catch (e) { return 'anon' + Math.random().toString(36).slice(2, 12); }
}
const MT_RS = (() => {
  const q = new URLSearchParams(location.search); let origem = q.get('utm_source') || q.get('o') || '', l = q.get('l') || '';
  try { if (origem) sessionStorage.setItem('mt_o', origem); else origem = sessionStorage.getItem('mt_o') || ''; if (l) localStorage.setItem('mt_l', l); else l = localStorage.getItem('mt_l') || ''; } catch (e) {}
  return { origem, l, ref: document.referrer && !document.referrer.startsWith(location.origin) ? document.referrer : '' };
})();
function mtEvento(tipo, alvo, dados) {
  try {
    const corpo = JSON.stringify({ v: mtVisitante(), tipo, pagina: location.pathname, alvo: alvo || '', origem: MT_RS.origem, ref: MT_RS.ref, l: MT_RS.l, dados: dados || undefined });
    if (navigator.sendBeacon && navigator.sendBeacon('/api/public/ev', new Blob([corpo], { type: 'application/json' }))) return;
    fetch('/api/public/ev', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: corpo, keepalive: true }).catch(() => {});
  } catch (e) {}
}
mtEvento('visita');
document.addEventListener('click', ev => {
  const a = ev.target.closest('a, button'); if (!a || a.type === 'submit' || a.closest('.escala')) return;
  const href = a.getAttribute('href') || '', txt = (a.innerText || a.title || '').replace(/\s+/g, ' ').trim().slice(0, 60);
  const alvo = /wa\.me|whatsapp/.test(href) ? 'WhatsApp: ' + txt : /instagram\.com/.test(href) ? 'Instagram: ' + txt : href && !href.startsWith('#') ? txt + ' → ' + href.replace(location.origin, '').slice(0, 80) : txt;
  if (alvo) mtEvento('clique', alvo);
}, true);
