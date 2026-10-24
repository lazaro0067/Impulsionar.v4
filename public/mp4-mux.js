/* Muxer MP4 simples (não fragmentado, "faststart": o índice vai antes dos dados).
   Recebe os pedaços já codificados pelo navegador (WebCodecs) e monta um MP4 completo,
   do jeito que o Instagram, a galeria e o WhatsApp esperam.
   Vídeo: H.264 (avc1) — e VP9 (vp09) só para testes. Áudio: AAC (mp4a) — e Opus só para testes. */
(function () {
  const TXT = new TextEncoder();
  const u8 = n => [n & 255];
  const u16 = n => [(n >>> 8) & 255, n & 255];
  const u24 = n => [(n >>> 16) & 255, (n >>> 8) & 255, n & 255];
  const u32 = n => [(n >>> 24) & 255, (n >>> 16) & 255, (n >>> 8) & 255, n & 255];
  const juntar = partes => {
    let tam = 0; const ps = partes.map(p => (p instanceof Uint8Array ? p : Array.isArray(p) ? Uint8Array.from(p) : typeof p === 'string' ? TXT.encode(p) : new Uint8Array(p)));
    ps.forEach(p => { tam += p.length; }); const out = new Uint8Array(tam); let o = 0; ps.forEach(p => { out.set(p, o); o += p.length; }); return out;
  };
  const caixa = (tipo, ...partes) => { const corpo = juntar(partes); return juntar([u32(corpo.length + 8), tipo, corpo]); };
  const cheia = (tipo, versao, flags, ...partes) => caixa(tipo, [versao], u24(flags), ...partes);
  const MATRIZ = [...u32(0x00010000), ...u32(0), ...u32(0), ...u32(0), ...u32(0x00010000), ...u32(0), ...u32(0), ...u32(0), ...u32(0x40000000)];
  const desc = (tag, corpo) => juntar([[tag], tamDesc(corpo.length), corpo]);
  const tamDesc = n => { const b = [n & 127]; n >>>= 7; while (n) { b.unshift((n & 127) | 128); n >>>= 7; } return b; };

  class Mp4Mux {
    constructor(op) {
      this.v = op.video ? Object.assign({ amostras: [], dados: [] }, op.video) : null;
      this.a = op.audio ? Object.assign({ amostras: [], dados: [] }, op.audio) : null;
      if (this.v) this.v.escala = 90000; if (this.a) this.a.escala = this.a.sampleRate;
    }
    addVideoChunk(chunk, meta) {
      const v = this.v; if (meta && meta.decoderConfig && meta.decoderConfig.description && !v.descricao) v.descricao = new Uint8Array(meta.decoderConfig.description.buffer ? meta.decoderConfig.description.buffer.slice(meta.decoderConfig.description.byteOffset || 0, (meta.decoderConfig.description.byteOffset || 0) + meta.decoderConfig.description.byteLength) : meta.decoderConfig.description);
      const b = new Uint8Array(chunk.byteLength); chunk.copyTo(b);
      v.dados.push(b); v.amostras.push({ tam: b.length, pts: chunk.timestamp, dur: chunk.duration || 0, chave: chunk.type === 'key' });
    }
    addAudioChunk(chunk, meta) {
      const a = this.a; if (meta && meta.decoderConfig && meta.decoderConfig.description && !a.descricao) { const d = meta.decoderConfig.description; a.descricao = new Uint8Array(d.buffer ? d.buffer.slice(d.byteOffset || 0, (d.byteOffset || 0) + d.byteLength) : d); }
      const b = new Uint8Array(chunk.byteLength); chunk.copyTo(b);
      a.dados.push(b); a.amostras.push({ tam: b.length, pts: chunk.timestamp, dur: chunk.duration || 0 });
    }
    // durações de cada amostra na escala da faixa (pela diferença entre os tempos; a última usa a própria duração)
    duracoes(tr) {
      const am = tr.amostras, esc = tr.escala, out = [];
      for (let i = 0; i < am.length; i++) {
        const prox = am[i + 1]; let d = prox ? (prox.pts - am[i].pts) : am[i].dur;
        if (!(d > 0)) d = am[i].dur || (tr === this.v ? 1e6 / 30 : 1024 / tr.sampleRate * 1e6);
        out.push(Math.max(1, Math.round(d * esc / 1e6)));
      }
      return out;
    }
    stts(durs) { const runs = []; durs.forEach(d => { const u = runs[runs.length - 1]; if (u && u[1] === d) u[0]++; else runs.push([1, d]); }); return cheia('stts', 0, 0, u32(runs.length), ...runs.map(([n, d]) => [...u32(n), ...u32(d)])); }
    entradaVideo() {
      const v = this.v, base = [...new Array(6).fill(0), ...u16(1), ...u16(0), ...u16(0), ...u32(0), ...u32(0), ...u32(0), ...u16(v.width), ...u16(v.height), ...u32(0x00480000), ...u32(0x00480000), ...u32(0), ...u16(1), ...new Array(32).fill(0), ...u16(0x18), ...u16(0xffff)];
      if (v.codec === 'vp9') return caixa('vp09', base, cheia('vpcC', 1, 0, [0, 10, (8 << 4) | (1 << 1) | 0, 1, 1, 1], u16(0)));
      if (!v.descricao) throw new Error('o codificador não entregou a configuração do H.264');
      return caixa('avc1', base, caixa('avcC', v.descricao));
    }
    entradaAudio() {
      const a = this.a, base = [...new Array(6).fill(0), ...u16(1), ...u32(0), ...u32(0), ...u16(a.channels), ...u16(16), ...u16(0), ...u16(0), ...u32((a.codec === 'opus' ? 48000 : a.sampleRate) * 65536 >>> 0)];
      if (a.codec === 'opus') {
        let preSkip = 312, entrada = a.sampleRate; const d = a.descricao;
        if (d && d.length >= 19 && String.fromCharCode(...d.slice(0, 8)) === 'OpusHead') { preSkip = d[10] | (d[11] << 8); entrada = d[12] | (d[13] << 8) | (d[14] << 16) | (d[15] << 24); }
        return caixa('Opus', base, caixa('dOps', [0, a.channels], u16(preSkip), u32(entrada >>> 0), u16(0), [0]));
      }
      // AudioSpecificConfig: o que o codificador mandou ou AAC-LC montado na mão
      let asc = a.descricao; if (!asc || !asc.length) { const idx = [96000, 88200, 64000, 48000, 44100, 32000, 24000, 22050, 16000, 12000, 11025, 8000].indexOf(a.sampleRate); asc = Uint8Array.from([(2 << 3) | (idx >> 1), ((idx & 1) << 7) | (a.channels << 3)]); }
      const dcd = desc(0x04, juntar([[0x40, 0x15], u24(0), u32(a.bitrate || 192000), u32(a.bitrate || 192000), desc(0x05, asc)]));
      const es = desc(0x03, juntar([u16(0), [0], dcd, desc(0x06, Uint8Array.from([2]))]));
      return caixa('mp4a', base, cheia('esds', 0, 0, es));
    }
    trak(tr, id, video, offsets, durMovie) {
      const durs = this.duracoes(tr), durMidia = durs.reduce((s, d) => s + d, 0);
      const tkhd = cheia('tkhd', 0, 3, u32(0), u32(0), u32(id), u32(0), u32(durMovie), u32(0), u32(0), u16(0), u16(video ? 0 : 1), u16(video ? 0 : 0x0100), u16(0), MATRIZ, u32((video ? tr.width : 0) * 65536 >>> 0), u32((video ? tr.height : 0) * 65536 >>> 0));
      const mdhd = cheia('mdhd', 0, 0, u32(0), u32(0), u32(tr.escala), u32(durMidia), u16(0x55c4), u16(0));
      const hdlr = cheia('hdlr', 0, 0, u32(0), video ? 'vide' : 'soun', u32(0), u32(0), u32(0), video ? 'VideoHandler\0' : 'SoundHandler\0');
      const xmhd = video ? cheia('vmhd', 0, 1, u16(0), u16(0), u16(0), u16(0)) : cheia('smhd', 0, 0, u16(0), u16(0));
      const dinf = caixa('dinf', cheia('dref', 0, 0, u32(1), cheia('url ', 0, 1)));
      const stsd = cheia('stsd', 0, 0, u32(1), video ? this.entradaVideo() : this.entradaAudio());
      const chaves = video ? tr.amostras.map((s, i) => s.chave ? i + 1 : 0).filter(Boolean) : null;
      const stbl = caixa('stbl', stsd, this.stts(durs), video && chaves.length && chaves.length < tr.amostras.length ? cheia('stss', 0, 0, u32(chaves.length), ...chaves.map(u32)) : [],
        cheia('stsc', 0, 0, u32(1), u32(1), u32(1), u32(1)),
        cheia('stsz', 0, 0, u32(0), u32(tr.amostras.length), ...tr.amostras.map(s => u32(s.tam))),
        cheia('stco', 0, 0, u32(offsets.length), ...offsets.map(o => u32(o))));
      return caixa('trak', tkhd, caixa('mdia', mdhd, hdlr, caixa('minf', xmhd, dinf, stbl)));
    }
    moov(base) {
      const faixas = [this.v, this.a].filter(t => t && t.amostras.length); let pos = base; const offs = [];
      faixas.forEach(t => { const o = []; t.amostras.forEach(s => { o.push(pos); pos += s.tam; }); offs.push(o); });
      const durMovie = Math.max(...faixas.map(t => Math.round(this.duracoes(t).reduce((s, d) => s + d, 0) * 1000 / t.escala)));
      const mvhd = cheia('mvhd', 0, 0, u32(0), u32(0), u32(1000), u32(durMovie), u32(0x00010000), u16(0x0100), u16(0), u32(0), u32(0), MATRIZ, ...new Array(6).fill(u32(0)), u32(faixas.length + 1));
      return caixa('moov', mvhd, ...faixas.map((t, i) => this.trak(t, i + 1, t === this.v, offs[i], durMovie)));
    }
    finish() {
      const ftyp = caixa('ftyp', 'isom', u32(512), 'isom', 'iso2', this.v && this.v.codec === 'vp9' ? 'vp09' : 'avc1', 'mp41');
      const dados = [this.v, this.a].filter(t => t && t.amostras.length).flatMap(t => t.dados), tamDados = dados.reduce((s, d) => s + d.length, 0);
      const provisorio = this.moov(0), base = ftyp.length + provisorio.length + 8;
      const moov = this.moov(base);
      return new Blob([ftyp, moov, Uint8Array.from(u32(tamDados + 8)), TXT.encode('mdat'), ...dados], { type: 'video/mp4' });
    }
  }
  window.Mp4Mux = Mp4Mux;
})();
