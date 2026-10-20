/* Câmera "igual à do celular" — usada pelo editor do Estúdio e pela gravação com teleprompter.
   - Abre a câmera na orientação em que o aparelho está (em pé = vídeo em pé), testando as combinações
     até o aparelho entregar a imagem certa (alguns Android insistem em mandar deitado).
   - Desenha no formato do projeto com zoom: "Tudo" (imagem inteira + fundo desfocado), 1x, 2x… e pinça.
   - Fundo desfocado leve que funciona em qualquer navegador (inclusive iPhone, que não tem ctx.filter). */
(function () {
  const CC = {};
  const parar = s => { try { s && s.getTracks().forEach(t => t.stop()); } catch (e) {} };

  // tamanho real dos quadros que a câmera está mandando (o getSettings mente em alguns aparelhos)
  CC.dims = s => new Promise(ok => {
    const v = document.createElement('video'); v.muted = true; v.playsInline = true; v.setAttribute('playsinline', ''); v.srcObject = s;
    let feito = false;
    const fim = () => {
      if (feito) return; feito = true;
      let w = v.videoWidth, h = v.videoHeight;
      if (!w) { const st = (s.getVideoTracks()[0] || { getSettings: () => ({}) }).getSettings(); w = st.width || 0; h = st.height || 0; }
      v.srcObject = null; ok({ w, h });
    };
    v.onloadedmetadata = () => setTimeout(fim, 60); v.play().catch(() => {}); setTimeout(fim, 2000);
  });

  // retrato: true = quer quadros em pé · alt: 1080 (Full HD) ou 720/2160 · base: { facingMode } ou { deviceId }
  CC.abrir = async function ({ base, audio, retrato, alt }) {
    alt = alt || 1080; const lg = Math.round(alt * 16 / 9);
    const tentativas = retrato
      ? [{ width: { ideal: alt }, height: { ideal: lg } }, { width: { ideal: lg }, height: { ideal: alt } }, { aspectRatio: { ideal: 9 / 16 } }, {}]
      : [{ width: { ideal: lg }, height: { ideal: alt } }, {}];
    let reserva = null, erro = null;
    for (const t of tentativas) {
      let s;
      try { s = await navigator.mediaDevices.getUserMedia({ video: Object.assign({ frameRate: { ideal: 30 } }, base || {}, t), audio: audio || false }); }
      catch (e) { erro = e; continue; }
      const d = await CC.dims(s);
      const certo = d.w > 0 && (retrato ? d.h >= d.w : d.w >= d.h);
      if (certo) return { stream: s, w: d.w, h: d.h, certo: true };
      if (!reserva) reserva = t;
      parar(s); // fecha antes de tentar de novo (o Android não abre a mesma câmera duas vezes)
    }
    if (reserva) {
      const s = await navigator.mediaDevices.getUserMedia({ video: Object.assign({ frameRate: { ideal: 30 } }, base || {}, reserva), audio: audio || false });
      const d = await CC.dims(s); return { stream: s, w: d.w, h: d.h, certo: false };
    }
    throw erro || new Error('Câmera indisponível');
  };

  // fundo desfocado: reduz a imagem para poucos pixels e amplia de novo (desfoque que funciona em todo navegador)
  CC.fundo = function (g, el, x, y, W, H) {
    const mw = el.videoWidth || el.naturalWidth || el.width, mh = el.videoHeight || el.naturalHeight || el.height; if (!mw || !mh) return;
    const c = CC._mini || (CC._mini = document.createElement('canvas'));
    const cw = 28, ch = Math.max(10, Math.round(cw * H / W)); if (c.width !== cw || c.height !== ch) { c.width = cw; c.height = ch; }
    const m = c.getContext('2d'), k = Math.max(cw / mw, ch / mh);
    m.drawImage(el, (cw - mw * k) / 2, (ch - mh * k) / 2, mw * k, mh * k);
    g.save(); g.imageSmoothingEnabled = true; try { g.imageSmoothingQuality = 'high'; } catch (e) {}
    g.drawImage(c, x, y, W, H); g.fillStyle = 'rgba(0,0,0,.32)'; g.fillRect(x, y, W, H); g.restore();
  };

  // desenha a câmera ocupando W×H: z = 1 preenche a tela (como a câmera do celular); z < 1 mostra mais (com fundo desfocado)
  CC.desenhar = function (g, el, W, H, z) {
    const mw = el.videoWidth, mh = el.videoHeight; if (!mw) return;
    const s = Math.max(W / mw, H / mh) * (z || 1), w = mw * s, h = mh * s;
    if (w < W - 1 || h < H - 1) CC.fundo(g, el, 0, 0, W, H);
    g.drawImage(el, (W - w) / 2, (H - h) / 2, w, h);
  };

  // zoom: usa o zoom da própria câmera quando o aparelho oferece (mais nítido) e completa com zoom digital
  CC.criarZoom = function (aoMudar) {
    const Z = { z: 1, min: 1, max: 4, base: 1, hw: null, dig: 1, track: null, hwAtual: null, fila: null, aplicando: false };
    Z.config = function (track, vw, vh, W, H, mostrarTudo) {
      Z.track = track; const cap = track && track.getCapabilities ? track.getCapabilities() : {};
      Z.hw = cap.zoom && cap.zoom.max > cap.zoom.min ? cap.zoom : null; Z.hwAtual = null;
      Z.base = Z.hw ? Math.max(Z.hw.min, Math.min(1, Z.hw.max)) : 1;
      const cob = Math.max(W / vw, H / vh), con = Math.min(W / vw, H / vh), fit = vw && vh ? con / cob : 1;
      Z.minHw = Z.hw ? Z.hw.min : 1;
      Z.min = Math.min(Z.base, Z.minHw * fit); Z.max = Z.hw ? Math.max(Z.hw.max, Z.base * 2) : 5;
      Z.setar(mostrarTudo ? Z.min : Z.base);
    };
    const aplicarHw = v => {
      if (!Z.track || !Z.hw) return;
      if (Z.aplicando) { Z.fila = v; return; }
      Z.aplicando = true;
      Z.track.applyConstraints({ advanced: [{ zoom: v }] }).catch(() => {}).then(() => { Z.aplicando = false; if (Z.fila !== null) { const f = Z.fila; Z.fila = null; aplicarHw(f); } });
    };
    Z.setar = function (z) {
      z = Math.max(Z.min, Math.min(Z.max, Number(z) || 1)); Z.z = z;
      if (Z.hw) { const hv = Math.max(Z.hw.min, Math.min(Z.hw.max, z)); if (Math.abs(hv - (Z.hwAtual || -1)) > 0.009) { Z.hwAtual = hv; aplicarHw(hv); } Z.dig = z / hv; }
      else Z.dig = z;
      aoMudar && aoMudar(Z);
    };
    Z.rotulo = function () { return Z.z <= Z.min + 0.005 && Z.min < Z.base - 0.03 ? 'Tudo' : (Math.round(Z.z * 10) / 10).toString().replace('.', ',') + 'x'; };
    Z.chips = function () {
      const l = []; if (Z.min < Z.base - 0.03) l.push([Z.min, '⤢ Tudo']);
      l.push([Z.base, '1x']); if (Z.base * 2 <= Z.max + 0.01) l.push([Z.base * 2, '2x']); if (Z.base * 4 <= Z.max + 0.01) l.push([Z.base * 4, '4x']);
      return l;
    };
    // pinça com dois dedos (e Ctrl + rodinha no computador); dois toques alternam entre "Tudo" e 1x
    Z.ligarPinca = function (el) {
      if (!el || el.__pinca) return; el.__pinca = true; el.style.touchAction = 'none';
      const pts = new Map(); let d0 = 0, z0 = 1, ultToque = 0;
      const dist = () => { const [a, b] = [...pts.values()]; return Math.hypot(a.x - b.x, a.y - b.y); };
      el.addEventListener('pointerdown', e => {
        pts.set(e.pointerId, { x: e.clientX, y: e.clientY });
        if (pts.size === 2) { d0 = dist(); z0 = Z.z; }
        if (pts.size === 1) { const ag = Date.now(); if (ag - ultToque < 300) Z.setar(Z.z > Z.base + 0.01 || Z.z < Z.base - 0.01 ? Z.base : Z.min); ultToque = ag; }
      });
      el.addEventListener('pointermove', e => { if (!pts.has(e.pointerId)) return; pts.set(e.pointerId, { x: e.clientX, y: e.clientY }); if (pts.size === 2 && d0 > 10) Z.setar(z0 * dist() / d0); });
      ['pointerup', 'pointercancel', 'pointerleave'].forEach(t => el.addEventListener(t, e => { pts.delete(e.pointerId); if (pts.size < 2) d0 = 0; }));
      el.addEventListener('wheel', e => { if (!e.ctrlKey) return; e.preventDefault(); Z.setar(Z.z * (e.deltaY < 0 ? 1.08 : 0.92)); }, { passive: false });
    };
    return Z;
  };
  window.CameraCerta = CC;
})();
