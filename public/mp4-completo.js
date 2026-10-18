// Transforma um MP4 "em pedaços" (fragmentado — é assim que o Chrome/Safari gravam) num MP4 normal e completo.
// Sem isso, a galeria do celular, o Instagram e alguns players leem só o primeiro pedaço (os "3 segundos").
// Não reencoda nada: só reorganiza o índice (rápido e sem perder qualidade). Funciona no navegador e no servidor.
(function (raiz) {
  'use strict';
  const u32 = (b, o) => ((b[o] << 24) >>> 0) + (b[o + 1] << 16) + (b[o + 2] << 8) + b[o + 3];
  const u64 = (b, o) => u32(b, o) * 4294967296 + u32(b, o + 4);
  const i32 = (b, o) => (b[o] << 24) | (b[o + 1] << 16) | (b[o + 2] << 8) | b[o + 3];
  const tipo = (b, o) => String.fromCharCode(b[o], b[o + 1], b[o + 2], b[o + 3]);

  // filhos de uma caixa (dentro de um buffer)
  function filhos(b, ini, fim) {
    const l = []; let p = ini;
    while (p + 8 <= fim) {
      let tam = u32(b, p), cab = 8; const t = tipo(b, p + 4);
      if (tam === 1) { tam = u64(b, p + 8); cab = 16; } else if (tam === 0) tam = fim - p;
      if (tam < cab || p + tam > fim) break;
      l.push({ t, ini: p, cab, fim: p + tam, dados: p + cab });
      p += tam;
    }
    return l;
  }
  const achar = (b, cx, t) => filhos(b, cx.dados, cx.fim).find(c => c.t === t);

  // escrita de caixas
  function caixa(t, ...partes) {
    const corpo = partes.reduce((s, x) => s + x.length, 0), out = new Uint8Array(8 + corpo);
    const dv = new DataView(out.buffer); dv.setUint32(0, 8 + corpo); for (let i = 0; i < 4; i++) out[4 + i] = t.charCodeAt(i);
    let p = 8; partes.forEach(x => { out.set(x, p); p += x.length; }); return out;
  }
  function inteiros(lista, bits) { // lista de números (32 ou 64 bits)
    const out = new Uint8Array(lista.length * (bits / 8)), dv = new DataView(out.buffer);
    lista.forEach((v, i) => { if (bits === 64) { dv.setUint32(i * 8, Math.floor(v / 4294967296)); dv.setUint32(i * 8 + 4, v >>> 0); } else dv.setUint32(i * 4, v >>> 0); });
    return out;
  }
  const cabFull = (versao, flags) => new Uint8Array([versao, (flags >> 16) & 255, (flags >> 8) & 255, flags & 255]);

  // Lê o arquivo por "janelas" (no navegador: Blob; no servidor: arquivo aberto).
  async function analisar(ler, tamanho, aoProgresso) {
    let p = 0, ftyp = null, moov = null; const moofs = [];
    while (p + 8 <= tamanho) {
      let h = await ler(p, Math.min(65536, tamanho - p));
      let tam = u32(h, 0), cab = 8; const t = tipo(h, 4);
      if (tam === 1) { tam = u64(h, 8); cab = 16; } else if (tam === 0) tam = tamanho - p;
      if (tam < cab) break;
      if (t === 'ftyp' || t === 'moov' || t === 'moof') {
        if (h.length < tam) h = await ler(p, tam);
        const b = h.slice(0, tam);
        if (t === 'ftyp') ftyp = b; else if (t === 'moov') moov = b; else moofs.push({ pos: p, b });
      }
      p += tam;
      if (aoProgresso && moofs.length % 200 === 0) aoProgresso(p / tamanho);
    }
    return { ftyp, moov, moofs };
  }

  function montarPlano(ftyp, moov, moofs) {
    const mv = { t: 'moov', ini: 0, cab: 8, fim: moov.length, dados: 8 };
    if (!filhos(moov, 8, moov.length).some(c => c.t === 'mvex') || !moofs.length) return null; // já é um MP4 normal
    const trex = {};
    const mvex = achar(moov, mv, 'mvex');
    filhos(moov, mvex.dados, mvex.fim).filter(c => c.t === 'trex').forEach(c => { const d = c.dados + 4; trex[u32(moov, d)] = { dur: u32(moov, d + 8), tam: u32(moov, d + 12), flags: u32(moov, d + 16) }; });
    const faixas = {};
    filhos(moov, 8, moov.length).filter(c => c.t === 'trak').forEach(tk => {
      const tkhd = achar(moov, tk, 'tkhd'), v = moov[tkhd.dados], id = u32(moov, tkhd.dados + (v === 1 ? 20 : 12));
      const mdia = achar(moov, tk, 'mdia'), mdhd = achar(moov, mdia, 'mdhd'), vm = moov[mdhd.dados];
      faixas[id] = { id, tk, mdia, escala: u32(moov, mdhd.dados + (vm === 1 ? 20 : 12)), amostras: [], tempo: 0, video: !!achar(moov, achar(moov, mdia, 'minf'), 'vmhd') };
    });
    const pedacos = []; // trechos contínuos de dados, na ordem do arquivo
    for (const mf of moofs) {
      const b = mf.b, raizMf = { dados: 8, fim: b.length };
      for (const tf of filhos(b, raizMf.dados, raizMf.fim).filter(c => c.t === 'traf')) {
        const tfhd = achar(b, tf, 'tfhd'); if (!tfhd) continue;
        const fl = u32(b, tfhd.dados) & 0xffffff; let q = tfhd.dados + 4; const id = u32(b, q); q += 4;
        const F = faixas[id]; if (!F) continue; const T = trex[id] || { dur: 0, tam: 0, flags: 0 };
        let base = mf.pos; if (fl & 1) { base = u64(b, q); q += 8; } if (fl & 2) q += 4;
        let dDur = T.dur, dTam = T.tam, dFl = T.flags;
        if (fl & 8) { dDur = u32(b, q); q += 4; } if (fl & 0x10) { dTam = u32(b, q); q += 4; } if (fl & 0x20) { dFl = u32(b, q); q += 4; }
        const tfdt = achar(b, tf, 'tfdt');
        if (tfdt) { const vt = b[tfdt.dados], bt = vt === 1 ? u64(b, tfdt.dados + 4) : u32(b, tfdt.dados + 4); if (bt > F.tempo && F.amostras.length) F.amostras[F.amostras.length - 1].dur += bt - F.tempo; F.tempo = Math.max(F.tempo, bt); }
        for (const tr of filhos(b, tf.dados, tf.fim).filter(c => c.t === 'trun')) {
          const vr = b[tr.dados], tfl = u32(b, tr.dados) & 0xffffff; let r = tr.dados + 4; const n = u32(b, r); r += 4;
          let off = base; if (tfl & 1) { off = base + i32(b, r); r += 4; }
          let primFl = null; if (tfl & 4) { primFl = u32(b, r); r += 4; }
          const ini = off;
          for (let i = 0; i < n; i++) {
            let dur = dDur, tam = dTam, sf = i === 0 && primFl !== null ? primFl : dFl, cto = 0;
            if (tfl & 0x100) { dur = u32(b, r); r += 4; } if (tfl & 0x200) { tam = u32(b, r); r += 4; }
            if (tfl & 0x400) { sf = u32(b, r); r += 4; } if (tfl & 0x800) { cto = vr === 1 ? i32(b, r) : u32(b, r); r += 4; }
            F.amostras.push({ dur, tam, sync: !(sf & 0x10000), cto, pedaco: pedacos.length });
            off += tam; F.tempo += dur;
          }
          if (n) pedacos.push({ ini, fim: off, id });
        }
      }
    }
    return { faixas, pedacos };
  }

  function montarMoov(ftyp, moov, plano, inicioMdat) {
    const mv = { dados: 8, fim: moov.length };
    // posição de cada trecho no arquivo novo
    let p = inicioMdat; plano.pedacos.forEach(pc => { pc.novo = p; p += pc.fim - pc.ini; });
    const grande = p > 4294967000;
    let durFilme = 0; const traks = [];
    Object.values(plano.faixas).forEach(F => {
      if (!F.amostras.length) return;
      const tot = F.amostras.reduce((s, a) => s + a.dur, 0), durMs = Math.round(tot * 1000 / F.escala); durFilme = Math.max(durFilme, durMs);
      // stts
      const stts = []; F.amostras.forEach(a => { const u = stts[stts.length - 1]; if (u && u[1] === a.dur) u[0]++; else stts.push([1, a.dur]); });
      // ctts
      const temCto = F.amostras.some(a => a.cto), ctts = []; if (temCto) F.amostras.forEach(a => { const u = ctts[ctts.length - 1]; if (u && u[1] === a.cto) u[0]++; else ctts.push([1, a.cto]); });
      // chunks = trechos desta faixa
      const chunks = []; F.amostras.forEach(a => { const u = chunks[chunks.length - 1]; if (u && u.pedaco === a.pedaco) u.n++; else chunks.push({ pedaco: a.pedaco, n: 1 }); });
      const stsc = []; chunks.forEach((c, i) => { const u = stsc[stsc.length - 1]; if (!u || u[1] !== c.n) stsc.push([i + 1, c.n, 1]); });
      const sync = F.amostras.every(a => a.sync) ? null : F.amostras.map((a, i) => a.sync ? i + 1 : 0).filter(Boolean);
      const mdia = F.mdia, minf = achar(moov, mdia, 'minf'), stbl = achar(moov, minf, 'stbl'), stsd = achar(moov, stbl, 'stsd');
      const caixasStbl = [moov.slice(stsd.ini, stsd.fim),
        caixa('stts', cabFull(0, 0), inteiros([stts.length], 32), inteiros(stts.flat(), 32))];
      if (temCto) caixasStbl.push(caixa('ctts', cabFull(1, 0), inteiros([ctts.length], 32), inteiros(ctts.flat().map(v => v < 0 ? v + 4294967296 : v), 32)));
      if (sync) caixasStbl.push(caixa('stss', cabFull(0, 0), inteiros([sync.length], 32), inteiros(sync, 32)));
      caixasStbl.push(caixa('stsc', cabFull(0, 0), inteiros([stsc.length], 32), inteiros(stsc.flat(), 32)));
      caixasStbl.push(caixa('stsz', cabFull(0, 0), inteiros([0, F.amostras.length], 32), inteiros(F.amostras.map(a => a.tam), 32)));
      const offs = chunks.map(c => plano.pedacos[c.pedaco].novo);
      caixasStbl.push(grande ? caixa('co64', cabFull(0, 0), inteiros([offs.length], 32), inteiros(offs, 64)) : caixa('stco', cabFull(0, 0), inteiros([offs.length], 32), inteiros(offs, 32)));
      const novoStbl = caixa('stbl', ...caixasStbl);
      const novoMinf = caixa('minf', ...filhos(moov, minf.dados, minf.fim).map(c => c.t === 'stbl' ? novoStbl : moov.slice(c.ini, c.fim)));
      const novoMdia = caixa('mdia', ...filhos(moov, mdia.dados, mdia.fim).map(c => {
        if (c.t === 'minf') return novoMinf;
        const x = moov.slice(c.ini, c.fim);
        if (c.t === 'mdhd') { const dv = new DataView(x.buffer), v = x[8]; if (v === 1) { dv.setUint32(32, Math.floor(tot / 4294967296)); dv.setUint32(36, tot >>> 0); } else dv.setUint32(24, Math.min(tot, 4294967295)); }
        return x;
      }));
      const partesTrak = filhos(moov, F.tk.dados, F.tk.fim).filter(c => c.t !== 'edts').map(c => {
        if (c.t === 'mdia') return novoMdia;
        const x = moov.slice(c.ini, c.fim);
        if (c.t === 'tkhd') { const dv = new DataView(x.buffer), v = x[8]; if (v === 1) { dv.setUint32(36, 0); dv.setUint32(40, durMs); } else dv.setUint32(28, durMs); }
        return x;
      });
      traks.push(caixa('trak', ...partesTrak));
    });
    const outros = filhos(moov, mv.dados, mv.fim).filter(c => !['trak', 'mvex', 'mvhd'].includes(c.t)).map(c => moov.slice(c.ini, c.fim));
    const mvhdOrig = achar(moov, mv, 'mvhd'), mvhd = moov.slice(mvhdOrig.ini, mvhdOrig.fim), dv = new DataView(mvhd.buffer);
    if (mvhd[8] === 1) { dv.setUint32(28, 1000); dv.setUint32(32, 0); dv.setUint32(36, durFilme); } else { dv.setUint32(20, 1000); dv.setUint32(24, durFilme); }
    return { moov: caixa('moov', mvhd, ...traks, ...outros), grande, duracao: durFilme / 1000 };
  }

  // Resultado: { cabecalho: Uint8Array (ftyp + moov + início do mdat), pedacos: [[ini, fim], ...] do arquivo original, duracao }
  // ou null quando o arquivo já é um MP4 normal.
  async function planejarMp4Completo(ler, tamanho, aoProgresso) {
    const { ftyp, moov, moofs } = await analisar(ler, tamanho, aoProgresso);
    if (!ftyp || !moov) return null;
    const plano = montarPlano(ftyp, moov, moofs); if (!plano || !plano.pedacos.length) return null;
    const totalDados = plano.pedacos.reduce((s, p) => s + p.fim - p.ini, 0);
    // o tamanho do moov depende de stco x co64; calcula duas vezes para acertar os endereços
    let r = montarMoov(ftyp, moov, plano, 0);
    const cabMdat = (r.grande || totalDados + 16 > 4294967000) ? 16 : 8;
    r = montarMoov(ftyp, moov, plano, ftyp.length + r.moov.length + cabMdat);
    const r2 = montarMoov(ftyp, moov, plano, ftyp.length + r.moov.length + cabMdat); // estável (stco/co64 já definido)
    const mdatCab = new Uint8Array(cabMdat), dv = new DataView(mdatCab.buffer);
    if (cabMdat === 16) { dv.setUint32(0, 1); mdatCab.set([109, 100, 97, 116], 4); dv.setUint32(8, Math.floor((totalDados + 16) / 4294967296)); dv.setUint32(12, (totalDados + 16) >>> 0); }
    else { dv.setUint32(0, totalDados + 8); mdatCab.set([109, 100, 97, 116], 4); }
    const cab = new Uint8Array(ftyp.length + r2.moov.length + cabMdat); cab.set(ftyp, 0); cab.set(r2.moov, ftyp.length); cab.set(mdatCab, ftyp.length + r2.moov.length);
    return { cabecalho: cab, pedacos: plano.pedacos.map(p => [p.ini, p.fim]), duracao: r2.duracao };
  }

  // Navegador: recebe um Blob/File e devolve um Blob MP4 completo (ou o mesmo, se já estiver certo).
  async function mp4CompletoBlob(blob, aoProgresso) {
    const ler = async (pos, n) => new Uint8Array(await blob.slice(pos, pos + n).arrayBuffer());
    const pl = await planejarMp4Completo(ler, blob.size, aoProgresso);
    if (!pl) return blob;
    return new Blob([pl.cabecalho, ...pl.pedacos.map(([a, b]) => blob.slice(a, b))], { type: 'video/mp4' });
  }

  // Rápido: só olha o começo do arquivo para saber se ele está "em pedaços".
  async function ehFragmentado(ler, tamanho) {
    let p = 0;
    for (let n = 0; n < 50 && p + 8 <= tamanho; n++) {
      let h = await ler(p, Math.min(65536, tamanho - p)); let tam = u32(h, 0); const t = tipo(h, 4);
      if (tam === 1) tam = u64(h, 8); else if (tam === 0) tam = tamanho - p;
      if (tam < 8) return false;
      if (t === 'moof') return true;
      if (t === 'moov') { if (h.length < tam) h = await ler(p, tam); return filhos(h, 8, tam).some(c => c.t === 'mvex'); }
      p += tam;
    }
    return false;
  }
  const api = { planejarMp4Completo, mp4CompletoBlob, ehFragmentado };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else raiz.Mp4Completo = api;
})(typeof window !== 'undefined' ? window : globalThis);
