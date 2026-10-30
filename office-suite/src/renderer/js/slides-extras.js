'use strict';
// Isadek Slides — Morph, Designer (design ideas), SmartArt, handouts and narration audio.
//   Morph: objects that appear on two slides in a row glide from one to the other. Objects
//          match when they share a morph key (el.mk: a duplicated slide's copies share it) or
//          have the same text or picture.
//   Designer: other arrangements of the slide's title, text and pictures.
//   SmartArt: a list, process, cycle, hierarchy or pyramid built from an outline, as a group
//          of ordinary shapes that remembers its outline (el.smart) so it can be rebuilt.
(function () {
  const W = 960, H = 540;
  const plain = html => { const d = document.createElement('div'); d.innerHTML = html || ''; return d.textContent.replace(/\s+/g, ' ').trim(); };
  const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  // =====================================================================================
  // Morph
  // =====================================================================================
  const signature = el => {
    if (el.type === 'image' || el.type === 'media') return el.type + ':' + (el.src || '').slice(-200) + ':' + (el.src || '').length;
    if (el.type === 'text' || el.type === 'shape' || el.type === 'table') { const t = plain(el.html); return t ? el.type + ':' + (el.shape || '') + ':' + t : null; }
    return null;
  };
  /** Pairs of objects on two slides: Map(id on `to` → element on `from`). */
  function pairs(from, to) {
    const out = new Map(), used = new Set();
    const cands = from.elements.filter(e => e.type !== 'connector');
    for (const b of to.elements) {
      if (b.type === 'connector' || !b.mk) continue;
      const a = cands.find(x => !used.has(x.id) && (x.mk === b.mk || x.id === b.mk || b.id === x.mk));
      if (a) { out.set(b.id, a); used.add(a.id); }
    }
    for (const b of to.elements) {
      if (b.type === 'connector' || out.has(b.id)) continue;
      const sig = signature(b);
      if (!sig) continue;
      const a = cands.find(x => !used.has(x.id) && signature(x) === sig);
      if (a) { out.set(b.id, a); used.add(a.id); }
    }
    return out;
  }
  /**
   * Plays Morph in `container` (which shows `from`): draws `to` with render(slide, container), then
   * moves matched objects from where they were, fades the rest in or out. Returns the animations.
   */
  function morph(container, from, to, render, ms) {
    const old = new Map([...container.querySelectorAll(':scope > [data-id]')].map(n => [n.dataset.id, n.cloneNode(true)]));
    const oldBg = container.style.backgroundColor, oldImg = container.style.backgroundImage;
    render(to, container);
    const match = pairs(from, to);
    const anims = [];
    const opts = { duration: ms, easing: 'cubic-bezier(.4,0,.2,1)' };
    const ghosts = document.createElement('div');
    ghosts.className = 'morph-ghosts';
    container.appendChild(ghosts);
    const matched = new Set([...match.values()].map(a => a.id));
    for (const a of from.elements) {
      if (matched.has(a.id)) continue;
      const n = old.get(a.id);
      if (!n) continue;
      ghosts.appendChild(n);
      anims.push(n.animate([{ opacity: 1 }, { opacity: 0 }], { ...opts, duration: ms * 0.6, fill: 'forwards' }));
    }
    for (const b of to.elements) {
      const node = container.querySelector(`:scope > [data-id="${b.id}"]`);
      if (!node) continue;
      const a = match.get(b.id);
      if (!a) { anims.push(node.animate([{ opacity: 0 }, { opacity: 0, offset: 0.35 }, { opacity: 1 }], opts)); continue; }
      const box = e => ({ left: e.x + 'px', top: e.y + 'px', width: e.w + 'px', height: e.h + 'px', transform: e.rot ? `rotate(${e.rot}deg)` : 'rotate(0deg)' });
      anims.push(node.animate([box(a), box(b)], opts));
      const c = node.querySelector(':scope > .content');
      if (c && (a.size !== b.size || a.color !== b.color)) anims.push(c.animate([{ fontSize: (a.size || b.size) + 'pt', ...(a.color && b.color ? { color: a.color } : {}) }, { fontSize: b.size + 'pt', ...(a.color && b.color ? { color: b.color } : {}) }], opts));
      const g = node.querySelector(':scope > svg g');
      if (g && a.fill && b.fill && a.fill !== b.fill) anims.push(g.animate([{ fill: a.fill }, { fill: b.fill }], opts));
    }
    if (oldBg !== container.style.backgroundColor || oldImg !== container.style.backgroundImage) {
      const bg = document.createElement('div');
      bg.className = 'morph-bg';
      Object.assign(bg.style, { backgroundColor: oldBg, backgroundImage: oldImg });
      container.insertBefore(bg, container.firstChild);
      anims.push(bg.animate([{ opacity: 1 }, { opacity: 0 }], { ...opts, fill: 'forwards' }));
    }
    const done = () => { ghosts.remove(); container.querySelectorAll(':scope > .morph-bg').forEach(n => n.remove()); };
    Promise.all(anims.map(x => x.finished.catch(() => {}))).then(done);
    return { anims, done, pairs: match };
  }

  // =====================================================================================
  // Designer
  // =====================================================================================
  const id = () => 'e' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
  const fit = (img, box) => {
    const k = Math.min(box.w / img.w, box.h / img.h);
    const w = Math.round(img.w * k), h = Math.round(img.h * k);
    return { x: Math.round(box.x + (box.w - w) / 2), y: Math.round(box.y + (box.h - h) / 2), w, h };
  };
  const cover = (img, box) => {
    const k = Math.max(box.w / img.w, box.h / img.h);
    const w = Math.round(img.w * k), h = Math.round(img.h * k);
    return { x: Math.round(box.x + (box.w - w) / 2), y: Math.round(box.y + (box.h - h) / 2), w, h };
  };
  /** Arrangements for a slide: [{ name, elements, bg }]. theme: { accent, text, title, bg }. */
  function ideas(slide, theme) {
    const els = slide.elements;
    const title = els.find(e => e.ph === 'title') || els.filter(e => e.type === 'text' && plain(e.html)).sort((a, b) => (b.size || 0) - (a.size || 0))[0];
    const texts = els.filter(e => e !== title && e.type === 'text' && plain(e.html));
    const imgs = els.filter(e => e.type === 'image');
    const others = els.filter(e => e !== title && !texts.includes(e) && !imgs.includes(e) && !(e.type === 'text' && !plain(e.html)));
    const T = (o, extra) => (o ? { ...JSON.parse(JSON.stringify(o)), ...extra } : null);
    const band = (x, y, w, h, fill) => ({ id: id(), type: 'shape', shape: 'rect', x, y, w, h, fill, line: null, html: '', size: 18, align: 'center', valign: 'middle', decorative: true });
    const out = [];
    const keep = list => list.filter(Boolean);
    const bodyText = texts[0];
    if (imgs.length === 1) {
      out.push({ name: 'Picture right', elements: keep([T(title, { x: 50, y: 40, w: 440, h: 110, align: 'left', valign: 'bottom' }), ...texts.map((t, i) => T(t, { x: 50, y: 165 + i * 10, w: 440, h: 330 - i * 10 })), T(imgs[0], fit(imgs[0], { x: 520, y: 40, w: 400, h: 460 })), ...others]) });
      out.push({ name: 'Picture left', elements: keep([T(imgs[0], fit(imgs[0], { x: 40, y: 40, w: 400, h: 460 })), T(title, { x: 470, y: 40, w: 450, h: 110, align: 'left', valign: 'bottom' }), ...texts.map((t, i) => T(t, { x: 470, y: 165 + i * 10, w: 450, h: 330 - i * 10 })), ...others]) });
      const cv = cover(imgs[0], { x: 0, y: 0, w: W, h: H });
      out.push({ name: 'Full-slide picture', elements: keep([T(imgs[0], { ...cv, decorative: imgs[0].decorative }), band(0, 380, W, 160, theme.accent), T(title, { x: 50, y: 392, w: 860, h: 80, color: '#ffffff', align: 'left', valign: 'middle' }), bodyText && T(bodyText, { x: 50, y: 470, w: 860, h: 56, color: '#ffffff', size: Math.min(bodyText.size || 20, 20) }), ...others]) });
    }
    if (imgs.length >= 2) {
      const n = Math.min(imgs.length, 6), cols = n <= 3 ? n : Math.ceil(n / 2), rows = Math.ceil(n / cols);
      const gw = (880 - (cols - 1) * 16) / cols, gh = (380 - (rows - 1) * 16) / rows;
      out.push({ name: 'Picture grid', elements: keep([T(title, { x: 40, y: 20, w: 880, h: 90, align: 'left' }), ...imgs.slice(0, n).map((im, i) => T(im, cover(im, { x: 40 + (i % cols) * (gw + 16), y: 130 + Math.floor(i / cols) * (gh + 16), w: gw, h: gh }))), ...imgs.slice(n), ...texts, ...others]) });
    }
    if (title) {
      out.push({ name: 'Title band', elements: keep([band(0, 0, W, 130, theme.accent), T(title, { x: 50, y: 20, w: 860, h: 95, color: '#ffffff', align: 'left', valign: 'middle' }), ...texts.map((t, i) => T(t, { x: 60, y: 160 + i * 10, w: 840, h: 340 - i * 10 })), ...imgs, ...others]) });
      out.push({ name: 'Side bar', elements: keep([band(0, 0, 300, H, theme.accent), T(title, { x: 30, y: 60, w: 250, h: 420, color: '#ffffff', align: 'left', valign: 'middle', size: Math.min(title.size || 40, 36) }), ...texts.map((t, i) => T(t, { x: 340, y: 60 + i * 10, w: 580, h: 420 - i * 10, valign: 'middle' })), ...imgs.map(im => T(im, fit(im, { x: 340, y: 60, w: 580, h: 420 }))), ...others]) });
      if (!imgs.length && (!bodyText || plain(bodyText.html).length < 160)) {
        out.push({ name: 'Statement', elements: keep([T(title, { x: 80, y: 150, w: 800, h: 150, align: 'center', valign: 'bottom', size: Math.max(title.size || 40, 48) }), band(400, 318, 160, 6, theme.accent), bodyText && T(bodyText, { x: 120, y: 340, w: 720, h: 120, align: 'center' }), ...texts.slice(1), ...others]) });
      }
      // Long bullet lists read better in two columns.
      if (bodyText) {
        const d = document.createElement('div');
        d.innerHTML = bodyText.html || '';
        const items = [...d.querySelectorAll('li')];
        if (items.length >= 6) {
          const half = Math.ceil(items.length / 2);
          const list = part => { const tag = (items[0].parentElement && items[0].parentElement.tagName.toLowerCase()) || 'ul'; return `<${tag}>${part.map(li => li.outerHTML).join('')}</${tag}>`; };
          out.push({ name: 'Two columns', elements: keep([T(title, { x: 60, y: 28, w: 840, h: 90 }), T(bodyText, { x: 60, y: 135, w: 405, h: 370, html: list(items.slice(0, half)) }), T(bodyText, { id: id(), mk: null, ph: null, x: 495, y: 135, w: 405, h: 370, html: list(items.slice(half)) }), ...texts.slice(1), ...imgs, ...others]) });
        }
      }
    }
    return out.map(o => ({ ...o, elements: o.elements.map(e => (e.id ? e : { ...e, id: id() })) }));
  }

  // =====================================================================================
  // SmartArt
  // =====================================================================================
  const SMART = { list: 'Vertical List', process: 'Process', cycle: 'Cycle', hierarchy: 'Hierarchy', pyramid: 'Pyramid' };
  /** Outline text → [{ text, level, children }] (two spaces or a tab per level). */
  function outline(text) {
    const root = { children: [] }, stack = [{ level: -1, node: root }];
    for (const raw of String(text || '').split(/\r?\n/)) {
      if (!raw.trim()) continue;
      const level = Math.min(5, Math.floor(raw.match(/^[\t ]*/)[0].replace(/\t/g, '  ').length / 2));
      const node = { text: raw.trim().replace(/^[-*•]\s*/, '').slice(0, 300), level, children: [] };
      while (stack.length > 1 && stack[stack.length - 1].level >= level) stack.pop();
      stack[stack.length - 1].node.children.push(node);
      stack.push({ level, node });
    }
    return root.children;
  }
  /** A SmartArt group in `box` ({x,y,w,h}); colours from the theme. */
  function smartArt(type, text, theme, box = { x: 80, y: 130, w: 800, h: 370 }) {
    const items = outline(text).slice(0, 12);
    if (!items.length) return null;
    const kids = [];
    const accent = theme.accent, ink = '#ffffff';
    const shape = (shapeName, x, y, w, h, label, o = {}) => kids.push({ id: id(), type: 'shape', shape: shapeName, x: Math.round(x), y: Math.round(y), w: Math.max(4, Math.round(w)), h: Math.max(4, Math.round(h)), fill: o.fill || accent, line: o.line || null, html: label ? `<p>${esc(label)}</p>${(o.sub || []).length ? `<ul>${o.sub.map(s => `<li>${esc(s)}</li>`).join('')}</ul>` : ''}` : '', size: o.size || 18, align: 'center', valign: 'middle', color: o.color || ink, ...(o.rot ? { rot: o.rot } : {}) });
    const n = items.length;
    const fontFor = (w, h, words) => Math.max(10, Math.min(28, Math.round(Math.min(h / 2.6, w / Math.max(4, Math.sqrt(words) * 5)))));
    const words = s => s.split(/\s+/).length;
    if (type === 'list') {
      const gap = 12, h = (box.h - gap * (n - 1)) / n;
      items.forEach((it, i) => {
        shape('roundRect', 0, i * (h + gap), box.w * 0.32, h, it.text, { size: fontFor(box.w * 0.32, h, words(it.text)) });
        if (it.children.length) shape('rect', box.w * 0.34, i * (h + gap), box.w * 0.66, h, it.children.map(c => c.text).join(' · '), { fill: '#f2f2f2', color: theme.text || '#222222', size: Math.min(18, fontFor(box.w * 0.66, h, 8)) });
      });
    } else if (type === 'process') {
      const arrowW = Math.min(50, box.w * 0.06), w = (box.w - arrowW * 1.6 * (n - 1)) / n, h = Math.min(box.h, w * 0.75);
      const y = (box.h - h) / 2;
      items.forEach((it, i) => {
        const x = i * (w + arrowW * 1.6);
        shape('roundRect', x, y, w, h, it.text, { sub: it.children.map(c => c.text), size: fontFor(w, h, words(it.text) + it.children.length * 3) });
        if (i < n - 1) shape('rightArrow', x + w + arrowW * 0.3, box.h / 2 - arrowW * 0.6, arrowW, arrowW * 1.2, '', { fill: '#a5a5a5' });
      });
    } else if (type === 'cycle') {
      const R = Math.min(box.w, box.h) / 2 - 10, cx = box.w / 2, cy = box.h / 2;
      const node = Math.min(150, R * 1.15 * Math.sin(Math.PI / Math.max(n, 3)) * 1.4);
      items.forEach((it, i) => {
        const t = -Math.PI / 2 + (i * 2 * Math.PI) / n;
        const x = cx + (R - node / 2) * Math.cos(t) - node / 2, y = cy + (R - node / 2) * Math.sin(t) - node * 0.35;
        shape('ellipse', x, y, node, node * 0.7, it.text, { size: fontFor(node, node * 0.7, words(it.text)) });
        // An arrow half-way to the next one, along the circle.
        const tm = t + Math.PI / n, ar = R - node / 2, aw = Math.max(18, node * 0.28);
        let deg = Math.round(((tm * 180) / Math.PI + 90) % 360);
        if (deg > 180) deg -= 360;
        shape('rightArrow', cx + ar * Math.cos(tm) - aw / 2, cy + ar * Math.sin(tm) - aw * 0.35, aw, aw * 0.7, '', { fill: '#a5a5a5', rot: deg });
      });
    } else if (type === 'hierarchy') {
      // Top level items side by side as roots; their children below each.
      const levels = [];
      const walk = (list, d) => list.forEach(it => { (levels[d] = levels[d] || []).push(it); walk(it.children, d + 1); });
      walk(items, 0);
      const depth = Math.min(levels.length, 4), lh = box.h / depth, bh = Math.min(70, lh * 0.6);
      const pos = new Map();
      // Leaves get equal widths; parents are centred over their children.
      const leaves = [];
      const collect = (it, d) => { if (!it.children.length || d === depth - 1) leaves.push(it); else it.children.forEach(c => collect(c, d + 1)); };
      items.forEach(it => collect(it, 0));
      const lw = box.w / leaves.length;
      leaves.forEach((it, i) => pos.set(it, i * lw + lw / 2));
      const place = (it, d) => { if (pos.has(it)) return pos.get(it); const xs = it.children.map(c => place(c, d + 1)); const x = xs.reduce((a, b) => a + b, 0) / xs.length; pos.set(it, x); return x; };
      items.forEach(it => place(it, 0));
      const bw = Math.min(180, lw * 0.85);
      const line = (x, y, w, h) => kids.push({ id: id(), type: 'shape', shape: 'rect', x: Math.round(x), y: Math.round(y), w: Math.max(2, Math.round(w)), h: Math.max(2, Math.round(h)), fill: '#a5a5a5', line: null, html: '', size: 10, align: 'center', valign: 'middle', decorative: true });
      const draw = (it, d) => {
        if (d >= depth) return;
        const x = pos.get(it), y = d * lh + (lh - bh) / 2;
        if (it.children.length && d + 1 < depth) {
          const ys = y + bh, yc = (d + 1) * lh + (lh - bh) / 2, mid = (ys + yc) / 2;
          line(x - 1, ys, 2, mid - ys);
          const cx = it.children.map(c => pos.get(c));
          line(Math.min(...cx), mid - 1, Math.max(...cx) - Math.min(...cx) + 2, 2);
          cx.forEach(px => line(px - 1, mid, 2, yc - mid));
        }
        shape('roundRect', x - bw / 2, y, bw, bh, it.text, { fill: d === 0 ? accent : shadeOf(accent, d), size: fontFor(bw, bh, words(it.text)) });
        it.children.forEach(c => draw(c, d + 1));
      };
      items.forEach(it => draw(it, 0));
    } else if (type === 'pyramid') {
      const h = box.h / n;
      items.forEach((it, i) => {
        const wTop = (box.w * 0.7 * i) / n, wBot = (box.w * 0.7 * (i + 1)) / n;
        const w = (wTop + wBot) / 2 + box.w * 0.05;
        shape(i === 0 ? 'triangle' : 'rect', (box.w - (i === 0 ? wBot : w)) / 2, i * h + 2, i === 0 ? wBot : w, h - 4, it.text, { fill: shadeOf(accent, i * 0.6), size: fontFor(w, h, words(it.text)) });
      });
    }
    if (!kids.length) return null;
    return { id: id(), type: 'group', x: box.x, y: box.y, w: box.w, h: box.h, smart: { type, text: String(text).slice(0, 5000) }, children: kids };
  }
  function shadeOf(hex, k) {
    const m = /^#?([0-9a-f]{6})$/i.exec(hex || '');
    if (!m) return hex;
    const n = parseInt(m[1], 16), f = Math.min(0.6, 0.18 * k);
    const ch = s => Math.round(((n >> s) & 255) + (255 - ((n >> s) & 255)) * f);
    return '#' + [16, 8, 0].map(s => ch(s).toString(16).padStart(2, '0')).join('');
  }

  // =====================================================================================
  // Handouts
  // =====================================================================================
  /**
   * Printable handout pages into `box`: perPage 1, 2, 3, 4, 6 or 9 slides, or 'notes' (a slide and
   * its notes per page). render(slide, el) draws a slide into an element 960×540.
   */
  function handouts(box, slides, perPage, render, { title = '', date = new Date().toLocaleDateString() } = {}) {
    box.innerHTML = '';
    const page = () => { const p = document.createElement('div'); p.className = 'handout-page'; box.appendChild(p); return p; };
    const frame = (p, x, y, w) => {
      const f = document.createElement('div');
      f.className = 'handout-slide';
      Object.assign(f.style, { left: x + 'in', top: y + 'in', width: w + 'in', height: (w * 9 / 16) + 'in' });
      const s = document.createElement('div');
      s.className = 'slide';
      s.style.transform = `scale(${(w * 96) / W})`;
      f.appendChild(s);
      p.appendChild(f);
      return s;
    };
    const head = (p, n, total) => {
      const h = document.createElement('div');
      h.className = 'handout-head';
      h.innerHTML = `<span>${esc(title)}</span><span>${esc(date)}</span>`;
      p.appendChild(h);
      const f = document.createElement('div');
      f.className = 'handout-foot';
      f.textContent = `${n} / ${total}`;
      p.appendChild(f);
    };
    const per = perPage === 'notes' ? 1 : perPage;
    const total = Math.ceil(slides.length / per);
    for (let i = 0, pn = 1; i < slides.length; i += per, pn++) {
      const p = page();
      head(p, pn, total);
      const chunk = slides.slice(i, i + per);
      if (perPage === 'notes') {
        render(chunk[0], frame(p, 0.75, 0.9, 7));
        const nt = document.createElement('div');
        nt.className = 'handout-notes';
        nt.textContent = chunk[0].notes || '';
        p.appendChild(nt);
      } else if (perPage === 3) {
        chunk.forEach((s, j) => {
          render(s, frame(p, 0.6, 0.9 + j * 3.2, 3.6));
          const lines = document.createElement('div');
          lines.className = 'handout-lines';
          Object.assign(lines.style, { left: '4.5in', top: (0.95 + j * 3.2) + 'in', width: '3.4in', height: '2in' });
          p.appendChild(lines);
        });
      } else {
        const cols = { 1: 1, 2: 1, 4: 2, 6: 2, 9: 3 }[perPage] || 2, rows = Math.ceil(perPage / cols);
        const areaW = 7.3, areaH = 9.2, gapX = 0.3, gapY = 0.35;
        let w = (areaW - (cols - 1) * gapX) / cols;
        w = Math.min(w, ((areaH - (rows - 1) * gapY) / rows) * 16 / 9);
        const h = w * 9 / 16;
        const x0 = 0.6 + (areaW - (cols * w + (cols - 1) * gapX)) / 2, y0 = 0.9 + (areaH - (rows * h + (rows - 1) * gapY)) / 2;
        chunk.forEach((s, j) => render(s, frame(p, x0 + (j % cols) * (w + gapX), y0 + Math.floor(j / cols) * (h + gapY), w)));
      }
    }
  }

  // =====================================================================================
  // Narration: recorded audio as WAV (which PowerPoint plays), mono at 22 kHz
  // =====================================================================================
  async function toWav(blob, rate = 22050) {
    const buf = await blob.arrayBuffer();
    const ctx = new (window.OfflineAudioContext || window.webkitOfflineAudioContext)(1, 1, rate);
    const audio = await ctx.decodeAudioData(buf.slice(0));
    const len = Math.max(1, Math.ceil(audio.duration * rate));
    const off = new OfflineAudioContext(1, len, rate);
    const src = off.createBufferSource();
    src.buffer = audio;
    src.connect(off.destination);
    src.start();
    const out = await off.startRendering();
    const pcm = out.getChannelData(0);
    const bytes = new Uint8Array(44 + pcm.length * 2);
    const v = new DataView(bytes.buffer);
    const str = (o, s) => [...s].forEach((ch, i) => v.setUint8(o + i, ch.charCodeAt(0)));
    str(0, 'RIFF'); v.setUint32(4, 36 + pcm.length * 2, true); str(8, 'WAVE'); str(12, 'fmt ');
    v.setUint32(16, 16, true); v.setUint16(20, 1, true); v.setUint16(22, 1, true); v.setUint32(24, rate, true); v.setUint32(28, rate * 2, true); v.setUint16(32, 2, true); v.setUint16(34, 16, true);
    str(36, 'data'); v.setUint32(40, pcm.length * 2, true);
    for (let i = 0; i < pcm.length; i++) v.setInt16(44 + i * 2, Math.max(-1, Math.min(1, pcm[i])) * 0x7fff, true);
    return bytes;
  }

  window.SlidesExtras = { pairs, morph, ideas, SMART, outline, smartArt, handouts, toWav };
})();
