'use strict';
// Atlas Slides — presentation editor and slide show.
(function () {
  const W = 960, H = 540; // logical slide size (10in × 5.625in at 96 px/in)
  const stage = document.getElementById('stage');
  const stageWrap = document.getElementById('stageWrap');
  const thumbs = document.getElementById('thumbs');
  const notes = document.getElementById('notes');
  const state = UI.docState('Atlas Slides', 'Presentation1');
  const SAVE_FILTERS = [{ name: 'PowerPoint Presentation', extensions: ['pptx'] }, { name: 'Atlas Presentation', extensions: ['aslides'] }];
  const OPEN_FILTERS = [{ name: 'Presentations', extensions: ['pptx', 'aslides'] }];

  const THEMES = {
    office: { name: 'Office', bg: '#ffffff', title: '#262626', text: '#404040', accent: '#4472c4', font: 'Calibri', titleFont: 'Calibri Light' },
    dark: { name: 'Midnight', bg: '#1f2430', title: '#ffffff', text: '#d0d4dc', accent: '#5fb3f9', font: 'Segoe UI', titleFont: 'Segoe UI' },
    ocean: { name: 'Ocean', bg: '#0f4c81', title: '#ffffff', text: '#dbe9f6', accent: '#ffd166', font: 'Calibri', titleFont: 'Georgia' },
    ion: { name: 'Ion', bg: '#13665c', title: '#ffffff', text: '#e3f4f1', accent: '#f4b942', font: 'Trebuchet MS', titleFont: 'Trebuchet MS' },
    sunset: { name: 'Sunset', bg: '#fff4ec', title: '#b7472a', text: '#5b3a29', accent: '#e07a5f', font: 'Georgia', titleFont: 'Georgia' },
    slate: { name: 'Slate', bg: '#f2f2f2', title: '#2f3e46', text: '#354f52', accent: '#52796f', font: 'Verdana', titleFont: 'Verdana' }
  };

  UI.initRibbon();
  document.execCommand('styleWithCSS', false, true);

  // ---------- Model ----------
  let uid = 1;
  const newId = () => 'e' + Date.now().toString(36) + (uid++);
  let pres = { theme: 'office', slides: [], master: null, layouts: {} };
  let current = 0;
  let selectedId = null;
  const extraSel = new Set(); // more objects selected with Shift/Ctrl+click
  let editingId = null;
  const theme = () => (pres.theme === 'custom' && pres.customTheme) || THEMES[pres.theme] || THEMES.office;
  // The audience window of Presenter View runs this page with ?audience=1 and only shows slides.
  const AUDIENCE = new URLSearchParams(location.search).get('audience') === '1';
  // In Slide Master view the "slide" being edited is the master or one of the layouts.
  let masterTarget = null;
  const slide = () => (masterTarget ? masterObj(masterTarget) : pres.slides[current]);
  const elById = id => slide().elements.find(e => e.id === id);

  const BUILTIN_LAYOUTS = { title: 'Title Slide', content: 'Title and Content', section: 'Section Header', two: 'Two Content', titleonly: 'Title Only', blank: 'Blank' };
  const layoutKeys = () => [...Object.keys(BUILTIN_LAYOUTS), ...Object.keys(pres.layouts || {}).filter(k => !BUILTIN_LAYOUTS[k])];
  const layoutName = k => (pres.layouts && pres.layouts[k] && pres.layouts[k].name) || BUILTIN_LAYOUTS[k] || 'Layout';
  /** A layout's placeholders for a new slide: the edited layout if there is one, else the built-in. */
  function layoutFor(kind) {
    const L = pres && pres.layouts && pres.layouts[kind];
    if (!L) return makeLayout(kind);
    return L.elements.map(el => ({ ...JSON.parse(JSON.stringify(el)), id: newId(), ...(el.ph ? { html: '' } : {}) }));
  }
  function defaultMaster() {
    const f = (ph, x, align) => ({ id: newId(), type: 'text', ph, html: '', x, y: 498, w: ph === 'footer' ? 360 : 240, h: 30, size: 12, align, valign: 'middle' });
    return { id: 'master', elements: [f('dt', 40, 'left'), f('footer', 300, 'center'), f('sldNum', 680, 'right')], bg: null, bgImage: null, footer: { text: '', number: false, date: false, skipTitle: true } };
  }
  function masterObj(t) {
    if (t === 'master') return pres.master;
    if (!pres.layouts[t]) pres.layouts[t] = { name: layoutName(t), elements: makeLayout(t) };
    const L = pres.layouts[t];
    return Object.assign(L, { id: 'layout-' + t, isLayout: true, bg: L.bg || null, bgImage: L.bgImage || null, notes: '' });
  }
  function makeLayout(kind) {
    const T = (o) => ({ id: newId(), type: 'text', html: '', align: 'left', valign: 'top', ...o });
    const title = o => T({ ph: 'title', size: 40, valign: 'middle', ...o });
    const body = o => T({ ph: 'body', size: 24, ...o });
    switch (kind) {
      case 'title': return [title({ x: 80, y: 140, w: 800, h: 150, size: 54, align: 'center', valign: 'bottom' }), T({ ph: 'subtitle', x: 120, y: 305, w: 720, h: 80, size: 24, align: 'center' })];
      case 'section': return [title({ x: 80, y: 190, w: 800, h: 110, size: 48, valign: 'bottom' }), T({ ph: 'subtitle', x: 80, y: 305, w: 800, h: 60, size: 20 })];
      case 'two': return [title({ x: 60, y: 28, w: 840, h: 90 }), body({ x: 60, y: 130, w: 410, h: 370 }), body({ x: 490, y: 130, w: 410, h: 370 })];
      case 'titleonly': return [title({ x: 60, y: 28, w: 840, h: 90 })];
      case 'blank': return [];
      case 'content': return [title({ x: 60, y: 28, w: 840, h: 90 }), body({ x: 60, y: 130, w: 840, h: 370 })];
      default:
        if (pres && pres.layouts && pres.layouts[kind]) return layoutFor(kind); return [title({ x: 60, y: 28, w: 840, h: 90 }), body({ x: 60, y: 130, w: 840, h: 370 })];
    }
  }
  /**
   * Presentations from files are untrusted: rebuild them from known fields with checked
   * types, fresh ids and sanitized text, so nothing can inject markup into the editor.
   */
  const SHAPE_NAMES = ['rect', 'roundRect', 'ellipse', 'triangle', 'rightArrow', 'line', 'star5', 'diamond'];
  const num = (v, d, lo = -5000, hi = 5000) => (typeof v === 'number' && isFinite(v) ? Math.max(lo, Math.min(hi, v)) : d);
  const color = v => (Platform.isColor(v) ? v : null);
  const imageUrl = v => (typeof v === 'string' && /^data:image\/(png|jpe?g|gif|webp|bmp|svg\+xml)[;,]/i.test(v) ? v : null);
  const MEDIA_TYPES = { mp4: 'video/mp4', m4v: 'video/mp4', webm: 'video/webm', ogv: 'video/ogg', mov: 'video/quicktime', mp3: 'audio/mpeg', m4a: 'audio/mp4', wav: 'audio/wav', ogg: 'audio/ogg', oga: 'audio/ogg', aac: 'audio/aac', flac: 'audio/flac' };
  const mediaUrl = v => (typeof v === 'string' && /^data:(video\/(mp4|webm|ogg|quicktime)|audio\/(mpeg|mp3|mp4|x-m4a|wav|x-wav|ogg|aac|flac));base64,/i.test(v) ? v : null);
  const oneOf = (v, list, d) => (list.includes(v) ? v : d);
  function cleanPres(p) {
    if (!p || !Array.isArray(p.slides)) throw new Error('Not a presentation');
    const slides = p.slides.slice(0, 2000).map(sl => {
      const s = newSlide('blank');
      if (!sl || typeof sl !== 'object') return s;
      s.layout = typeof sl.layout === 'string' && (BUILTIN_LAYOUTS[sl.layout] || (p.layouts && Object.hasOwn(p.layouts, sl.layout))) ? sl.layout : 'blank';
      s.bg = color(sl.bg);
      s.bgImage = imageUrl(sl.bgImage);
      s.transition = oneOf(sl.transition, SlidesAnim.TRANSITIONS, 'none');
      s.transDur = num(sl.transDur, 1, 0.1, 30);
      s.advance = typeof sl.advance === 'number' && isFinite(sl.advance) ? Math.max(0, Math.min(3600, sl.advance)) : null;
      s.notes = typeof sl.notes === 'string' ? sl.notes.slice(0, 100000) : '';
      s.elements = cleanElements(sl.elements);
      if (sl.hideMaster === true) s.hideMaster = true;
      return s;
    });
    if (!slides.length) slides.push(newSlide('title'));
    const out = { theme: Object.prototype.hasOwnProperty.call(THEMES, p.theme) || p.theme === 'custom' ? p.theme : 'office', slides };
    const ct = p.customTheme;
    if (out.theme === 'custom') {
      const fontOk = f => (typeof f === 'string' && /^[\w .-]{1,64}$/.test(f) ? f : null);
      if (ct && typeof ct === 'object' && color(ct.bg) && color(ct.title) && color(ct.text) && color(ct.accent)) out.customTheme = { name: 'Custom', bg: ct.bg, title: ct.title, text: ct.text, accent: ct.accent, font: fontOk(ct.font) || 'Calibri', titleFont: fontOk(ct.titleFont) || 'Calibri Light' };
      else out.theme = 'office';
    }
    const m = p.master && typeof p.master === 'object' ? p.master : null;
    out.master = defaultMaster();
    if (m) {
      out.master.elements = cleanElements(m.elements, ['footer', 'sldNum', 'dt']);
      out.master.bg = color(m.bg); out.master.bgImage = imageUrl(m.bgImage);
      const f = m.footer || {};
      out.master.footer = { text: typeof f.text === 'string' ? f.text.slice(0, 500) : '', number: f.number === true, date: f.date === true, skipTitle: f.skipTitle !== false };
    }
    out.layouts = {};
    if (p.layouts && typeof p.layouts === 'object') {
      Object.keys(p.layouts).slice(0, 100).forEach(k => {
        if (!/^[\w-]{1,40}$/.test(k)) return;
        const L = p.layouts[k];
        if (!L || typeof L !== 'object') return;
        out.layouts[k] = { name: typeof L.name === 'string' ? L.name.slice(0, 100) : layoutName(k), elements: cleanElements(L.elements), bg: color(L.bg), bgImage: imageUrl(L.bgImage) };
        if (L.custom === true || !BUILTIN_LAYOUTS[k]) out.layouts[k].custom = true;
      });
    }
    return out;
  }
  /** Elements from a file, rebuilt from known fields (phs: extra placeholder kinds allowed, e.g. footers on the master). */
  function cleanElements(list, phs = [], depth = 0) {
    const ids = new Map(); // old id -> new id, so connectors keep pointing at their shapes
    const out = (Array.isArray(list) ? list : []).slice(0, 500).map(el => cleanElement(el, phs, depth, ids)).filter(Boolean);
    out.forEach(c => {
      if (c.type !== 'connector') return;
      ['from', 'to'].forEach(k => { if (c[k]) { const id = ids.get(c[k].id); if (id) c[k].id = id; else delete c[k]; } });
    });
    return out;
  }
  function cleanElement(el, phs, depth, ids) {
    {
      if (!el || typeof el !== 'object') return null;
      const type = oneOf(el.type, ['text', 'shape', 'image', 'table', 'media', 'group', 'connector'], null);
      if (!type) return null;
      const out = { id: newId(), type, x: num(el.x, 100), y: num(el.y, 100), w: num(el.w, 200, 1), h: num(el.h, 100, 1) };
      if (typeof el.id === 'string') ids.set(el.id, out.id);
      if (el.flipH === true) out.flipH = true;
      if (el.flipV === true) out.flipV = true;
      if (type === 'group') {
        if (depth >= 5) return null;
        out.children = cleanElements(el.children, phs, depth + 1).filter(c => c.type !== 'connector');
        if (!out.children.length) return null;
        if (el.rot) out.rot = num(el.rot, 0, -360, 360);
        const anim = SlidesAnim.clean(el.anim);
        if (anim) out.anim = anim;
        return out;
      }
      if (type === 'connector') {
        out.kind = oneOf(el.kind, ['straight', 'elbow'], 'straight');
        out.arrow = oneOf(el.arrow, ['none', 'end', 'both'], 'end');
        if (color(el.color)) out.color = el.color;
        out.width = num(el.width, 2, 0.5, 20);
        ['x1', 'y1', 'x2', 'y2'].forEach(k => { out[k] = num(el[k], 0); });
        ['from', 'to'].forEach(k => { const e = el[k]; if (e && typeof e === 'object' && typeof e.id === 'string') out[k] = { id: e.id, site: oneOf(e.site, ['n', 's', 'e', 'w', 'auto'], 'auto') }; });
        return out;
      }
      if (type === 'media') {
        out.src = mediaUrl(el.src);
        if (!out.src) return null;
        out.kind = /^data:audio/i.test(out.src) ? 'audio' : 'video';
        if (imageUrl(el.poster)) out.poster = el.poster;
        ['autoplay', 'loop', 'hideIcon'].forEach(k => { if (el[k] === true) out[k] = true; });
        if (el.rot) out.rot = num(el.rot, 0, -360, 360);
        const anim = SlidesAnim.clean(el.anim);
        if (anim) out.anim = anim;
        return out;
      }
      if (el.rot) out.rot = num(el.rot, 0, -360, 360);
      const anim = SlidesAnim.clean(el.anim);
      if (anim) out.anim = anim;
      if (type === 'image') { out.src = imageUrl(el.src); return out.src ? out : null; }
      out.html = Platform.sanitizeHtml(el.html);
      out.size = num(el.size, 18, 4, 400);
      out.align = oneOf(el.align, ['left', 'center', 'right'], 'left');
      out.valign = oneOf(el.valign, ['top', 'middle', 'bottom'], 'top');
      if (typeof el.font === 'string' && /^[\w .-]{1,64}$/.test(el.font)) out.font = el.font;
      if (color(el.color)) out.color = el.color;
      if (['title', 'subtitle', 'body', ...phs].includes(el.ph)) out.ph = el.ph;
      if (el.bold === false) out.bold = false;
      if (type === 'shape') { out.shape = oneOf(el.shape, SHAPE_NAMES, 'rect'); out.fill = color(el.fill); out.line = color(el.line); }
      return out;
    }
  }
  const newSlide = kind => ({ id: newId(), layout: kind, bg: null, bgImage: null, transition: 'none', transDur: 1, advance: null, notes: '', elements: layoutFor(kind) });
  const PLACEHOLDER = { title: 'Click to add title', subtitle: 'Click to add subtitle', body: 'Click to add text', footer: 'Footer', sldNum: '‹#›', dt: 'Date' };

  // ---------- Undo ----------
  const undoStack = [], redoStack = [];
  // Live editing: remote changes that arrived after a step are put back after undoing it.
  const undoMarks = [], redoMarks = [];
  let live = null, livePeers = [];
  const liveMark = () => (live ? live.mark() : 0);
  function snapshot() {
    undoStack.push(JSON.stringify({ pres, current }));
    undoMarks.push(liveMark());
    if (undoStack.length > 100) { undoStack.shift(); undoMarks.shift(); }
    redoStack.length = redoMarks.length = 0;
  }
  function restore(json) {
    const s = JSON.parse(json);
    pres = s.pres; current = Math.min(s.current, pres.slides.length - 1);
    selectedId = editingId = null;
    renderAll();
    state.markDirty();
  }
  function undo() {
    if (!undoStack.length) return;
    redoStack.push(JSON.stringify({ pres, current })); redoMarks.push(liveMark());
    const m = undoMarks.pop();
    restore(undoStack.pop());
    if (live) live.replaySince(m);
  }
  function redo() {
    if (!redoStack.length) return;
    undoStack.push(JSON.stringify({ pres, current })); undoMarks.push(liveMark());
    const m = redoMarks.pop();
    restore(redoStack.pop());
    if (live) live.replaySince(m);
  }

  // ---------- Rendering ----------
  const SHAPES = {
    rect: () => '<rect x="0" y="0" width="100" height="100"/>',
    roundRect: () => '<rect x="0" y="0" width="100" height="100" rx="12" ry="12"/>',
    ellipse: () => '<ellipse cx="50" cy="50" rx="50" ry="50"/>',
    triangle: () => '<polygon points="50,0 100,100 0,100"/>',
    rightArrow: () => '<polygon points="0,25 60,25 60,0 100,50 60,100 60,75 0,75"/>',
    line: () => '<line x1="0" y1="100" x2="100" y2="0"/>',
    star5: () => '<polygon points="50,0 61,35 98,35 68,57 79,91 50,70 21,91 32,57 2,35 39,35"/>',
    diamond: () => '<polygon points="50,0 100,50 50,100 0,50"/>'
  };
  function textColor(el) { return el.color || (el.ph === 'title' ? theme().title : theme().text); }
  function fontOf(el) { return el.font || (el.ph === 'title' ? theme().titleFont : theme().font); }

  /** The master's shapes and visible footers as they appear on slide s (none on the master itself). */
  function masterItems(s) {
    const M = pres.master, out = [];
    if (!M || s === M || s.hideMaster) return out;
    const n = pres.slides.indexOf(s) + 1, f = M.footer;
    const showFooters = !s.isLayout && !(f.skipTitle && s.layout === 'title');
    M.elements.forEach(el => {
      if (el.ph === 'footer' || el.ph === 'sldNum' || el.ph === 'dt') {
        if (!showFooters) return;
        const text = el.ph === 'footer' ? (f.text && escapeHtml(f.text)) : el.ph === 'sldNum' ? (f.number && String(n)) : (f.date && escapeHtml(new Date().toLocaleDateString()));
        if (text) out.push({ ...el, html: `<div>${text}</div>`, master: true });
        return;
      }
      out.push({ ...el, master: true });
    });
    return out;
  }
  /** Builds the DOM for one slide. interactive=false for thumbnails, print and the show. */
  function renderSlide(s, container, interactive) {
    container.innerHTML = '';
    const M = pres.master;
    const isMaster = s === M;
    // Backgrounds come from the slide, else its layout, else the master, else the theme.
    const L = !isMaster && !s.isLayout && pres.layouts && pres.layouts[s.layout];
    const bgFrom = [s, L, !isMaster && M].find(x => x && (x.bg || x.bgImage)) || {};
    container.style.backgroundColor = bgFrom.bg || theme().bg;
    container.style.backgroundImage = imageUrl(bgFrom.bgImage) ? `url("${bgFrom.bgImage}")` : '';
    const list = [...masterItems(s), ...s.elements];
    list.forEach(el => {
      const div = buildEl(el, s, interactive);
      if (!el.master) div.dataset.id = el.id;
      if (interactive && !el.master && el.anim && document.body.classList.contains('show-anims')) {
        const b = document.createElement('span');
        b.className = 'anim-badge';
        b.textContent = animNumber(s, el);
        div.appendChild(b);
      }
      if (interactive && (el.id === selectedId || extraSel.has(el.id))) {
        div.classList.add('selected');
        if (el.id === editingId) div.classList.add('editing');
        if (el.id === selectedId && !extraSel.size && el.type !== 'connector') {
          ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w', 'rot'].forEach(h => { const d = document.createElement('div'); d.className = 'handle ' + h; d.dataset.h = h; if (h === 'rot') d.title = 'Rotate (Shift: 15° steps)'; div.appendChild(d); });
        }
      }
      container.appendChild(div);
    });
  }
  /** Where a connector's ends are: on the facing sides of the shapes it joins, else its own points. */
  function connectorPoints(c, s) {
    const box = id => s.elements.find(e => e.id === id && e.type !== 'connector');
    const site = (e, side) => ({ n: [e.x + e.w / 2, e.y], s: [e.x + e.w / 2, e.y + e.h], w: [e.x, e.y + e.h / 2], e: [e.x + e.w, e.y + e.h / 2] }[side] || [e.x + e.w / 2, e.y + e.h / 2]);
    const a = c.from && box(c.from.id), b = c.to && box(c.to.id);
    let p1 = a ? site(a, c.from.site) : [c.x1, c.y1], p2 = b ? site(b, c.to.site) : [c.x2, c.y2];
    if (a && b && (c.from.site === 'auto' || c.to.site === 'auto' || !c.from.site)) {
      const sides = facingSides(a, b);
      p1 = site(a, sides[0]); p2 = site(b, sides[1]);
    }
    return [p1, p2];
  }
  function facingSides(a, b) {
    const dx = (b.x + b.w / 2) - (a.x + a.w / 2), dy = (b.y + b.h / 2) - (a.y + a.h / 2);
    return Math.abs(dx) >= Math.abs(dy) ? (dx >= 0 ? ['e', 'w'] : ['w', 'e']) : (dy >= 0 ? ['s', 'n'] : ['n', 's']);
  }
  /** The DOM for one element (and, for groups, their children). */
  function buildEl(el, s, interactive) {
    const div = document.createElement('div');
    div.className = 'el ' + el.type + (el.master ? ' master-el' : '');
    if (el.master) div.dataset.masterId = el.id;
    if (el.type === 'connector') {
      const [[x1, y1], [x2, y2]] = connectorPoints(el, s);
      const pad = 12, minX = Math.min(x1, x2) - pad, minY = Math.min(y1, y2) - pad;
      Object.assign(el, { x: Math.round(Math.min(x1, x2)), y: Math.round(Math.min(y1, y2)), w: Math.max(1, Math.round(Math.abs(x2 - x1))), h: Math.max(1, Math.round(Math.abs(y2 - y1))) });
      Object.assign(div.style, { left: minX + 'px', top: minY + 'px', width: Math.abs(x2 - x1) + 2 * pad + 'px', height: Math.abs(y2 - y1) + 2 * pad + 'px' });
      const lx1 = x1 - minX, ly1 = y1 - minY, lx2 = x2 - minX, ly2 = y2 - minY;
      const horizontalFirst = Math.abs(x2 - x1) >= Math.abs(y2 - y1);
      const d = el.kind === 'elbow'
        ? (horizontalFirst ? `M${lx1},${ly1} H${(lx1 + lx2) / 2} V${ly2} H${lx2}` : `M${lx1},${ly1} V${(ly1 + ly2) / 2} H${lx2} V${ly2}`)
        : `M${lx1},${ly1} L${lx2},${ly2}`;
      const col = color(el.color) || theme().text, wd = num(el.width, 2, 0.5, 20);
      const mk = `m${el.id}`;
      div.innerHTML = `<svg width="100%" height="100%" overflow="visible"><defs><marker id="${mk}e" markerWidth="10" markerHeight="10" refX="9" refY="5" orient="auto" markerUnits="userSpaceOnUse"><path d="M0,0 L10,5 L0,10 z" fill="${col}"/></marker><marker id="${mk}s" markerWidth="10" markerHeight="10" refX="1" refY="5" orient="auto-start-reverse" markerUnits="userSpaceOnUse"><path d="M0,0 L10,5 L0,10 z" fill="${col}"/></marker></defs>`
        + `<path class="conn-hit" d="${d}" fill="none" stroke="transparent" stroke-width="12"/><path d="${d}" fill="none" stroke="${col}" stroke-width="${wd}"${el.arrow === 'end' || el.arrow === 'both' ? ` marker-end="url(#${mk}e)"` : ''}${el.arrow === 'both' ? ` marker-start="url(#${mk}s)"` : ''}/></svg>`;
      return div;
    }
    const flip = `${el.flipH ? ' scaleX(-1)' : ''}${el.flipV ? ' scaleY(-1)' : ''}`;
    Object.assign(div.style, { left: el.x + 'px', top: el.y + 'px', width: el.w + 'px', height: el.h + 'px', transform: el.rot ? `rotate(${el.rot}deg)` : '' });
    if (el.type === 'group') {
      (el.children || []).forEach(ch => div.appendChild(buildEl(ch, s, false)));
      return div;
    }
    if (el.type === 'media') {
      if (el.kind === 'video') {
        const v = document.createElement('video');
        v.src = el.src; v.preload = 'metadata'; v.playsInline = true;
        if (el.poster) v.poster = el.poster;
        v.loop = !!el.loop;
        v.className = 'media-video';
        div.appendChild(v);
      } else {
        const a = document.createElement('audio');
        a.src = el.src; a.preload = 'metadata'; a.loop = !!el.loop;
        div.appendChild(a);
        const icon = document.createElement('div');
        icon.className = 'media-audio-icon';
        icon.textContent = '🔊';
        div.appendChild(icon);
      }
      if (interactive) { const tag = document.createElement('div'); tag.className = 'media-badge'; tag.textContent = (el.kind === 'video' ? '▶ Video' : '♪ Audio') + (el.autoplay ? ' · auto' : ''); div.appendChild(tag); }
    } else if (el.type === 'image') {
      const img = document.createElement('img');
      img.src = el.src;
      img.alt = '';
      if (flip) img.style.transform = flip.trim();
      div.appendChild(img);
    } else {
      if (el.type === 'shape') {
        // Colours are checked here too, since they go into SVG markup.
        const stroke = color(el.line) || 'none';
        const fill = el.shape === 'line' ? 'none' : (color(el.fill) || 'none');
        div.insertAdjacentHTML('afterbegin', `<svg viewBox="0 0 100 100" preserveAspectRatio="none"${flip ? ` style="transform:${flip.trim()}"` : ''}><g fill="${fill}" stroke="${el.shape === 'line' ? (color(el.line) || theme().accent) : stroke}" stroke-width="${el.shape === 'line' ? 3 : 1.5}" vector-effect="non-scaling-stroke">${SHAPES[el.shape] ? SHAPES[el.shape]() : SHAPES.rect()}</g></svg>`);
        div.querySelectorAll('svg *').forEach(n => n.setAttribute('vector-effect', 'non-scaling-stroke'));
      }
      const c = document.createElement('div');
      c.className = 'content v-' + (el.valign || 'top');
      c.innerHTML = el.html || '';
      Object.assign(c.style, { fontSize: el.size + 'pt', color: textColor(el), textAlign: el.align || 'left', fontFamily: fontOf(el) });
      if (el.ph === 'title') c.style.fontWeight = el.bold === false ? '400' : '';
      if (el.type === 'table') div.style.setProperty('--tbl-accent', theme().accent);
      if ((interactive || s.isLayout || s === pres.master) && el.type === 'text' && el.ph && !el.master) c.dataset.placeholder = PLACEHOLDER[el.ph];
      else if (interactive && el.type === 'text') c.dataset.placeholder = 'Type here';
      div.appendChild(c);
    }
    return div;
  }

  function renderStage() {
    renderSlide(slide(), stage, true);
    if (editingId) {
      const c = stage.querySelector(`[data-id="${editingId}"] .content`);
      if (c) c.contentEditable = 'true';
    }
    notes.value = slide().notes || '';
    document.getElementById('stSlide').textContent = masterTarget ? (masterTarget === 'master' ? 'Slide Master' : 'Layout: ' + layoutName(masterTarget)) : `Slide ${current + 1} of ${pres.slides.length}`;
    document.getElementById('hideMaster').classList.toggle('on', !masterTarget && !!slide().hideMaster);
    syncToolbar();
    drawLivePeers();
    if (live) live.moved();
  }
  function renderThumbs() {
    thumbs.innerHTML = '';
    if (masterTarget) return renderMasterThumbs();
    pres.slides.forEach((s, i) => {
      const t = document.createElement('div');
      t.className = 'thumb' + (i === current ? ' active' : '');
      t.draggable = true;
      t.innerHTML = `<div class="num">${i + 1}</div><div class="frame"><div class="slide"></div></div>`;
      renderSlide(s, t.querySelector('.slide'), false);
      t.addEventListener('click', () => goTo(i));
      t.addEventListener('dragstart', e => { e.dataTransfer.setData('text/plain', String(i)); });
      t.addEventListener('dragover', e => { e.preventDefault(); t.classList.add('drag-over'); });
      t.addEventListener('dragleave', () => t.classList.remove('drag-over'));
      t.addEventListener('drop', e => {
        e.preventDefault();
        const from = +e.dataTransfer.getData('text/plain');
        if (from === i) return;
        snapshot();
        const [m] = pres.slides.splice(from, 1);
        pres.slides.splice(i, 0, m);
        current = i;
        renderAll();
        state.markDirty();
      });
      thumbs.appendChild(t);
    });
    thumbs.querySelector('.thumb.active')?.scrollIntoView({ block: 'nearest' });
  }
  function renderMasterThumbs() {
    [['master', 'Slide Master'], ...layoutKeys().map(k => [k, layoutName(k)])].forEach(([key, name]) => {
      const t = document.createElement('div');
      t.className = 'thumb' + (key === 'master' ? ' master-thumb' : ' layout-thumb') + (key === masterTarget ? ' active' : '');
      t.title = name;
      t.innerHTML = '<div class="num"></div><div class="frame"><div class="slide"></div></div><div class="thumb-name"></div>';
      t.querySelector('.thumb-name').textContent = name;
      renderSlide(masterObj(key), t.querySelector('.slide'), false);
      t.addEventListener('click', () => { commitEditing(); masterTarget = key; selectedId = null; renderAll(); });
      thumbs.appendChild(t);
    });
    thumbs.querySelector('.thumb.active')?.scrollIntoView({ block: 'nearest' });
  }
  function setMasterView(on) {
    commitEditing();
    masterTarget = on ? 'master' : null;
    selectedId = editingId = null;
    document.body.classList.toggle('master-view', on);
    const tab = document.querySelector('.ribbon-tabs [data-tab="master"]');
    tab.hidden = !on;
    (on ? tab : document.querySelector('.ribbon-tabs [data-tab="home"]')).click();
    renderAll();
  }
  document.getElementById('masterView').addEventListener('click', () => setMasterView(true));
  document.getElementById('masterView2').addEventListener('click', () => setMasterView(true));
  document.getElementById('closeMaster').addEventListener('click', () => setMasterView(false));
  document.getElementById('insLayout').addEventListener('click', async () => {
    const r = await UI.form('Insert Layout', [{ name: 'name', label: 'Layout name', value: 'Custom Layout' }, { name: 'from', label: 'Start from', type: 'select', options: layoutKeys().map(k => ({ value: k, label: layoutName(k) })) }], 'Insert');
    if (!r) return;
    snapshot();
    const key = 'L' + Date.now().toString(36);
    pres.layouts[key] = { name: r.name.trim().slice(0, 100) || 'Custom Layout', custom: true, elements: JSON.parse(JSON.stringify(masterObj(r.from).elements)).map(el => ({ ...el, id: newId() })) };
    masterTarget = key;
    renderAll(); state.markDirty();
  });
  document.getElementById('renameLayout').addEventListener('click', async () => {
    if (!masterTarget || masterTarget === 'master') return UI.toast('Select a layout first');
    const name = await UI.prompt('Rename Layout', 'Layout name', layoutName(masterTarget));
    if (!name) return;
    snapshot(); masterObj(masterTarget).name = name.slice(0, 100); renderAll(); state.markDirty();
  });
  document.getElementById('delLayout').addEventListener('click', async () => {
    const k = masterTarget;
    if (!k || k === 'master' || BUILTIN_LAYOUTS[k]) return UI.toast('Only custom layouts can be deleted (built-in ones can be edited)');
    if (pres.slides.some(sl => sl.layout === k) && !(await UI.confirm('Delete Layout', 'Slides use this layout. They will keep their content and become Blank-layout slides.'))) return;
    snapshot();
    pres.slides.forEach(sl => { if (sl.layout === k) sl.layout = 'blank'; });
    delete pres.layouts[k];
    masterTarget = 'master';
    renderAll(); state.markDirty();
  });
  document.getElementById('insPlaceholder').addEventListener('change', e => {
    const kind = e.target.value; e.target.value = '';
    if (!kind || !masterTarget || masterTarget === 'master') { if (kind) UI.toast('Placeholders go on layouts: select a layout first'); return; }
    snapshot();
    const base = { id: newId(), type: 'text', html: '', align: 'left', valign: 'top', ph: kind };
    slide().elements.push(kind === 'title' ? { ...base, x: 60, y: 28, w: 840, h: 90, size: 40, valign: 'middle' } : kind === 'subtitle' ? { ...base, x: 120, y: 300, w: 720, h: 80, size: 24, align: 'center' } : { ...base, x: 100, y: 150, w: 400, h: 250, size: 24 });
    selectedId = slide().elements[slide().elements.length - 1].id;
    renderAll(); state.markDirty();
  });
  document.getElementById('resetSlide').addEventListener('click', () => {
    // Put the slide's placeholders back where its layout has them, as PowerPoint's Reset does.
    snapshot();
    const s = slide(), fresh = layoutFor(s.layout);
    s.elements.filter(el => el.ph).forEach(el => {
      const f = fresh.find(x => x.ph === el.ph && !x._used);
      if (f) { f._used = true; Object.assign(el, { x: f.x, y: f.y, w: f.w, h: f.h, size: f.size, align: f.align, valign: f.valign }); }
    });
    renderAll(); state.markDirty();
  });
  document.getElementById('headerFooter').addEventListener('click', async () => {
    const f = pres.master.footer, yn = [{ value: '1', label: 'Yes' }, { value: '', label: 'No' }], ny = [{ value: '', label: 'No' }, { value: '1', label: 'Yes' }];
    const r = await UI.form('Header and Footer', [
      { name: 'date', label: 'Date', type: 'select', options: f.date ? yn : ny },
      { name: 'number', label: 'Slide number', type: 'select', options: f.number ? yn : ny },
      { name: 'text', label: 'Footer text', value: f.text },
      { name: 'skip', label: 'Don\'t show on title slide', type: 'select', options: f.skipTitle ? yn : ny }
    ], 'Apply to All');
    if (!r) return;
    snapshot();
    pres.master.footer = { text: r.text.slice(0, 500), number: !!r.number, date: !!r.date, skipTitle: !!r.skip };
    if (r.text && !pres.master.elements.some(el => el.ph === 'footer')) pres.master.elements.push(...defaultMaster().elements.filter(el => el.ph === 'footer'));
    if (r.number && !pres.master.elements.some(el => el.ph === 'sldNum')) pres.master.elements.push(...defaultMaster().elements.filter(el => el.ph === 'sldNum'));
    if (r.date && !pres.master.elements.some(el => el.ph === 'dt')) pres.master.elements.push(...defaultMaster().elements.filter(el => el.ph === 'dt'));
    renderAll(); state.markDirty();
  });
  document.getElementById('headerFooter2').addEventListener('click', () => document.getElementById('headerFooter').click());
  document.getElementById('hideMaster').addEventListener('click', e => {
    if (masterTarget) return;
    snapshot(); slide().hideMaster = !slide().hideMaster; e.currentTarget.classList.toggle('on', !!slide().hideMaster); renderAll(); state.markDirty();
  });
  function fillLayoutSelect() {
    const sel = document.getElementById('layout');
    const keys = layoutKeys();
    if (sel.options.length === keys.length && keys.every((k, i) => sel.options[i].value === k && sel.options[i].textContent === layoutName(k))) return;
    sel.textContent = '';
    keys.forEach(k => { const o = document.createElement('option'); o.value = k; o.textContent = layoutName(k); sel.appendChild(o); });
  }
  let thumbTimer;
  function refreshThumbSoon() {
    clearTimeout(thumbTimer);
    thumbTimer = setTimeout(() => {
      const frame = (masterTarget ? thumbs.querySelector('.thumb.active') : thumbs.children[current])?.querySelector('.slide');
      if (frame) renderSlide(slide(), frame, false);
    }, 250);
  }
  function renderAll() { renderStage(); renderThumbs(); fitStage(); drawLivePeers(); }
  function fitStage() {
    const r = stageWrap.getBoundingClientRect();
    const s = Math.max(0.1, Math.min((r.width - 40) / W, (r.height - 40) / H));
    stage.style.transform = `scale(${s})`;
    stage.dataset.scale = s;
  }
  window.addEventListener('resize', fitStage);
  new ResizeObserver(fitStage).observe(stageWrap);

  function goTo(i) {
    commitEditing();
    if (masterTarget) setMasterView(false);
    current = Math.max(0, Math.min(pres.slides.length - 1, i));
    selectedId = null;
    renderAll();
  }

  // ---------- Slide operations ----------
  function addSlide(kind) {
    commitEditing();
    if (masterTarget) return UI.toast('Close Master View to add slides');
    snapshot();
    // Like PowerPoint: after a title slide comes "Title and Content", otherwise repeat the current layout.
    const cur = slide() && slide().layout;
    pres.slides.splice(current + 1, 0, newSlide(kind || (!cur || cur === 'title' ? 'content' : cur)));
    current++;
    selectedId = null;
    renderAll();
    state.markDirty();
  }
  document.getElementById('newSlide').addEventListener('click', () => addSlide());
  document.getElementById('newSlide2').addEventListener('click', () => addSlide());
  document.getElementById('layout').addEventListener('change', e => {
    // Changing layout on the current slide: re-flow placeholders, keep other content.
    snapshot();
    const s = slide();
    const old = s.elements.filter(el => el.ph);
    const fresh = layoutFor(e.target.value);
    fresh.forEach(f => { const match = old.find(o => o.ph === f.ph && !o._used); if (match) { f.html = match.html; match._used = true; } });
    s.elements = [...fresh, ...s.elements.filter(el => !el.ph)];
    s.layout = e.target.value;
    renderAll();
    state.markDirty();
  });
  document.getElementById('dupSlide').addEventListener('click', () => {
    snapshot();
    const copy = JSON.parse(JSON.stringify(slide()));
    copy.id = newId();
    copy.elements.forEach(el => { el.id = newId(); });
    pres.slides.splice(current + 1, 0, copy);
    current++;
    renderAll();
    state.markDirty();
  });
  function deleteSlide() {
    if (masterTarget) return;
    if (pres.slides.length === 1) return UI.toast('A presentation needs at least one slide');
    snapshot();
    pres.slides.splice(current, 1);
    current = Math.min(current, pres.slides.length - 1);
    selectedId = null;
    renderAll();
    state.markDirty();
  }
  document.getElementById('delSlide').addEventListener('click', deleteSlide);
  thumbs.tabIndex = 0;
  thumbs.addEventListener('keydown', e => {
    if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); deleteSlide(); }
    if (e.key === 'ArrowDown') { e.preventDefault(); goTo(current + 1); }
    if (e.key === 'ArrowUp') { e.preventDefault(); goTo(current - 1); }
    if (e.key === 'Enter') { e.preventDefault(); addSlide(); }
  });

  // ---------- Element interaction ----------
  const scale = () => +stage.dataset.scale || 1;
  function select(id) {
    if (editingId && editingId !== id) commitEditing();
    selectedId = id;
    renderStage();
  }
  function startEditing(id, caretAtEnd = true) {
    const el = elById(id);
    if (!el || el.type === 'image') return;
    snapshot();
    selectedId = editingId = id;
    renderStage();
    const c = stage.querySelector(`[data-id="${id}"] .content`);
    c.contentEditable = 'true';
    // Body placeholders start as a bullet list, like PowerPoint.
    if (!c.innerHTML && el.ph === 'body') c.innerHTML = '<ul><li><br></li></ul>';
    c.focus();
    const r = document.createRange();
    r.selectNodeContents(c.querySelector('li') || c);
    r.collapse(!caretAtEnd);
    getSelection().removeAllRanges();
    getSelection().addRange(r);
  }
  function commitEditing() {
    if (!editingId) return;
    const el = elById(editingId);
    const c = stage.querySelector(`[data-id="${editingId}"] .content`);
    if (el && c) {
      let html = c.innerHTML;
      if (!c.textContent.trim() && !c.querySelector('img')) html = '';
      el.html = html;
    }
    editingId = null;
    renderStage();
    refreshThumbSoon();
  }
  stage.addEventListener('input', e => {
    const elDiv = e.target.closest('.el');
    if (!elDiv) return;
    const el = elById(elDiv.dataset.id);
    if (el) { el.html = e.target.closest('.content').innerHTML; state.markDirty(); refreshThumbSoon(); }
  });

  stage.addEventListener('mousedown', e => {
    const elDiv = e.target.closest('.el');
    if (connectMode) { e.preventDefault(); connectClick(elDiv && elDiv.dataset.id); return; }
    if (!elDiv) { commitEditing(); extraSel.clear(); select(null); return; }
    const id = elDiv.dataset.id;
    if (!id) return;
    // Shift/Ctrl+click adds or removes an object from the selection.
    if ((e.shiftKey || e.ctrlKey || e.metaKey) && selectedId && id !== selectedId && !editingId) {
      e.preventDefault();
      if (extraSel.has(id)) extraSel.delete(id); else extraSel.add(id);
      renderStage();
      return;
    }
    if (extraSel.size && (extraSel.has(id) || id === selectedId) && !e.target.classList.contains('handle')) { e.preventDefault(); startMove(e, elById(id)); return; }
    extraSel.clear();
    if (id === editingId && !e.target.classList.contains('handle')) return; // let the caret move
    const el = elById(id);
    if (e.target.classList.contains('handle')) return e.target.dataset.h === 'rot' ? startRotate(e, el) : startResize(e, el, e.target.dataset.h);
    if (el.type === 'text' && !el.html && id !== editingId) { e.preventDefault(); startEditing(id); return; }
    if (el.type === 'group' || el.type === 'connector') { e.preventDefault(); if (selectedId !== id) select(id); if (el.type === 'group' || !el.from) startMove(e, el); return; }
    e.preventDefault();
    if (selectedId !== id) select(id);
    startMove(e, el);
  });
  stage.addEventListener('dblclick', e => {
    const elDiv = e.target.closest('.el');
    const el = elDiv && elById(elDiv.dataset.id);
    if (!el || ['group', 'connector', 'image', 'media'].includes(el.type)) return;
    if (elDiv.dataset.id !== editingId) startEditing(elDiv.dataset.id);
  });

  function clearGuides() { stage.querySelectorAll('.guide').forEach(g => g.remove()); }
  function guide(vertical, pos) {
    const g = document.createElement('div');
    g.className = 'guide';
    Object.assign(g.style, vertical ? { left: pos + 'px', top: 0, width: '1px', height: H + 'px' } : { top: pos + 'px', left: 0, height: '1px', width: W + 'px' });
    stage.appendChild(g);
  }
  function startMove(e, el) {
    const sx = e.clientX, sy = e.clientY, ox = el.x, oy = el.y;
    let moved = false;
    const div = stage.querySelector(`[data-id="${el.id}"]`);
    // Everything selected moves together; connectors follow their shapes.
    const group = [el, ...[...extraSel].map(elById).filter(x => x && x !== el)];
    const start = group.map(x => ({ x, ox: x.x, oy: x.y, pts: x.type === 'connector' ? [x.x1, x.y1, x.x2, x.y2] : null, div: stage.querySelector(`[data-id="${x.id}"]`) }));
    const others = slide().elements.filter(o => !group.includes(o) && o.type !== 'connector');
    const mv = ev => {
      let nx = ox + (ev.clientX - sx) / scale();
      let ny = oy + (ev.clientY - sy) / scale();
      if (!moved && Math.abs(ev.clientX - sx) + Math.abs(ev.clientY - sy) < 3) return;
      if (!moved) { snapshot(); moved = true; }
      clearGuides();
      if (!ev.altKey) {
        // Snap centres and edges to the slide and to other objects.
        const xs = [[W / 2, 'c'], [0, 's'], [W, 'e'], ...others.flatMap(o => [[o.x, 's'], [o.x + o.w, 'e'], [o.x + o.w / 2, 'c']])];
        const ys = [[H / 2, 'c'], [0, 's'], [H, 'e'], ...others.flatMap(o => [[o.y, 's'], [o.y + o.h, 'e'], [o.y + o.h / 2, 'c']])];
        for (const [t] of xs) {
          for (const [edge, off] of [[nx, 0], [nx + el.w / 2, el.w / 2], [nx + el.w, el.w]]) if (Math.abs(edge - t) < 6) { nx = t - off; guide(true, t); break; }
        }
        for (const [t] of ys) {
          for (const [edge, off] of [[ny, 0], [ny + el.h / 2, el.h / 2], [ny + el.h, el.h]]) if (Math.abs(edge - t) < 6) { ny = t - off; guide(false, t); break; }
        }
      }
      const dx = Math.round(nx) - ox, dy = Math.round(ny) - oy;
      start.forEach(o => {
        // An unattached connector end moves with the drag; an attached end stays on its shape.
        if (o.pts) { if (!o.x.from) { o.x.x1 = o.pts[0] + dx; o.x.y1 = o.pts[1] + dy; } if (!o.x.to) { o.x.x2 = o.pts[2] + dx; o.x.y2 = o.pts[3] + dy; } return; }
        o.x.x = o.ox + dx; o.x.y = o.oy + dy;
        if (o.div) { o.div.style.left = o.x.x + 'px'; o.div.style.top = o.x.y + 'px'; }
      });
      if (div && el.type !== 'connector') { div.style.left = el.x + 'px'; div.style.top = el.y + 'px'; }
      if (hasConnectors() || el.type === 'connector') renderStage();
    };
    const up = () => {
      document.removeEventListener('mousemove', mv); document.removeEventListener('mouseup', up);
      clearGuides();
      if (moved) { state.markDirty(); refreshThumbSoon(); }
    };
    document.addEventListener('mousemove', mv); document.addEventListener('mouseup', up);
  }
  const hasConnectors = () => slide().elements.some(x => x.type === 'connector');
  /** Drag the round handle above an object to turn it; Shift snaps to 15°, and it settles on right angles. */
  function startRotate(e, el) {
    e.preventDefault(); e.stopPropagation();
    snapshot();
    const r = stage.getBoundingClientRect(), k = scale();
    const cx = r.left + (el.x + el.w / 2) * k, cy = r.top + (el.y + el.h / 2) * k;
    const div = stage.querySelector(`[data-id="${el.id}"]`);
    const mv = ev => {
      let a = Math.atan2(ev.clientY - cy, ev.clientX - cx) * 180 / Math.PI + 90;
      if (ev.shiftKey) a = Math.round(a / 15) * 15;
      else { const near = Math.round(a / 90) * 90; if (Math.abs(a - near) < 3) a = near; }
      a = ((Math.round(a) % 360) + 540) % 360 - 180;
      el.rot = a || 0;
      if (!el.rot) delete el.rot;
      div.style.transform = el.rot ? `rotate(${el.rot}deg)` : '';
    };
    const up = () => { document.removeEventListener('mousemove', mv); document.removeEventListener('mouseup', up); renderStage(); state.markDirty(); refreshThumbSoon(); };
    document.addEventListener('mousemove', mv); document.addEventListener('mouseup', up);
  }
  function startResize(e, el, h) {
    e.preventDefault(); e.stopPropagation();
    snapshot();
    const sx = e.clientX, sy = e.clientY, o = { x: el.x, y: el.y, w: el.w, h: el.h };
    const kids = el.type === 'group' ? JSON.parse(JSON.stringify(el.children)) : null;
    const ratio = o.w / o.h;
    const div = stage.querySelector(`[data-id="${el.id}"]`);
    const mv = ev => {
      const dx = (ev.clientX - sx) / scale(), dy = (ev.clientY - sy) / scale();
      let { x, y, w, h: hh } = o;
      if (h.includes('e')) w = o.w + dx;
      if (h.includes('s')) hh = o.h + dy;
      if (h.includes('w')) { w = o.w - dx; x = o.x + dx; }
      if (h.includes('n')) { hh = o.h - dy; y = o.y + dy; }
      if ((ev.shiftKey || el.type === 'image') && h.length === 2) { // keep aspect ratio on corners
        if (w / hh > ratio) w = hh * ratio; else hh = w / ratio;
        if (h.includes('w')) x = o.x + o.w - w;
        if (h.includes('n')) y = o.y + o.h - hh;
      }
      el.w = Math.max(10, Math.round(w)); el.h = Math.max(10, Math.round(hh));
      el.x = Math.round(x); el.y = Math.round(y);
      // A group scales what is inside it.
      if (kids) { const fx = el.w / o.w, fy = el.h / o.h; const scaleKids = (list, src) => list.forEach((c, i) => { const s0 = src[i]; c.x = Math.round(s0.x * fx); c.y = Math.round(s0.y * fy); c.w = Math.max(1, Math.round(s0.w * fx)); c.h = Math.max(1, Math.round(s0.h * fy)); }); scaleKids(el.children, kids); renderStage(); return; }
      Object.assign(div.style, { left: el.x + 'px', top: el.y + 'px', width: el.w + 'px', height: el.h + 'px' });
      if (hasConnectors()) renderStage();
    };
    const up = () => { document.removeEventListener('mousemove', mv); document.removeEventListener('mouseup', up); state.markDirty(); refreshThumbSoon(); };
    document.addEventListener('mousemove', mv); document.addEventListener('mouseup', up);
  }

  let elClipboard = null;
  document.addEventListener('keydown', e => {
    if (!document.getElementById('presenter').hidden) return;
    if (document.querySelector('.modal-backdrop') || document.querySelector('.backstage.open')) return;
    const mod = e.ctrlKey || e.metaKey;
    if (e.key === 'F5') { e.preventDefault(); present(e.shiftKey ? current : 0); return; }
    if (editingId) { if (e.key === 'Escape') { e.preventDefault(); commitEditing(); } return; }
    const tag = document.activeElement.tagName;
    if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT') return;
    if (mod && e.key.toLowerCase() === 'm') { e.preventDefault(); addSlide(); return; }
    if (!Platform.desktop && mod && e.key.toLowerCase() === 'z') { e.preventDefault(); e.shiftKey ? redo() : undo(); return; }
    if (!Platform.desktop && mod && e.key.toLowerCase() === 'y') { e.preventDefault(); redo(); return; }
    if (document.activeElement === thumbs) return;
    const el = selectedId && elById(selectedId);
    if (!el) {
      if (e.key === 'PageDown') goTo(current + 1);
      if (e.key === 'PageUp') goTo(current - 1);
      return;
    }
    if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); deleteElement(); }
    else if (e.key === 'Escape') { extraSel.clear(); select(null); }
    else if (mod && e.key.toLowerCase() === 'g') { e.preventDefault(); if (e.shiftKey) ungroup(); else groupSelected(); }
    else if (e.key === 'Enter' || e.key === 'F2') { e.preventDefault(); startEditing(el.id); }
    else if (/^Arrow/.test(e.key)) {
      e.preventDefault();
      snapshot();
      const d = e.shiftKey ? 10 : 1;
      [el, ...[...extraSel].map(elById).filter(Boolean)].forEach(x => {
        if (x.type === 'connector') return;
        if (e.key === 'ArrowLeft') x.x -= d; if (e.key === 'ArrowRight') x.x += d;
        if (e.key === 'ArrowUp') x.y -= d; if (e.key === 'ArrowDown') x.y += d;
      });
      renderStage(); refreshThumbSoon(); state.markDirty();
    } else if (mod && e.key.toLowerCase() === 'd') { e.preventDefault(); duplicateElement(); }
    else if (mod && e.key.toLowerCase() === 'c') { elClipboard = JSON.parse(JSON.stringify(el)); }
    else if (mod && e.key.toLowerCase() === 'x') { elClipboard = JSON.parse(JSON.stringify(el)); deleteElement(); }
    else if (mod && e.key.toLowerCase() === 'b') { e.preventDefault(); wholeBox('bold'); }
    else if (mod && e.key.toLowerCase() === 'i') { e.preventDefault(); wholeBox('italic'); }
    else if (mod && e.key.toLowerCase() === 'u') { e.preventDefault(); wholeBox('underline'); }
    else if (e.key.length === 1 && !mod && !e.altKey && !['image', 'media', 'group', 'connector'].includes(el.type)) {
      // Typing on a selected box replaces its text, like PowerPoint.
      e.preventDefault();
      snapshot();
      el.html = '';
      startEditing(el.id);
      document.execCommand('insertText', false, e.key);
    }
  });
  document.addEventListener('paste', async e => {
    if (editingId || !document.getElementById('presenter').hidden) return;
    const tag = document.activeElement.tagName;
    if (tag === 'INPUT' || tag === 'TEXTAREA') return;
    const file = [...e.clipboardData.files].find(f => f.type.startsWith('image/'));
    if (file) { e.preventDefault(); addImage(await Platform.bytesToDataUrl(new Uint8Array(await file.arrayBuffer()), file.type)); return; }
    if (elClipboard) {
      e.preventDefault();
      snapshot();
      const c = JSON.parse(JSON.stringify(elClipboard));
      c.id = newId(); c.x += 20; c.y += 20; delete c.ph;
      elClipboard = JSON.parse(JSON.stringify(c));
      slide().elements.push(c);
      selectedId = c.id;
      renderStage(); refreshThumbSoon(); state.markDirty();
      return;
    }
    const text = e.clipboardData.getData('text/plain');
    if (text) { e.preventDefault(); addText(text.split('\n').map(escapeHtml).join('<br>')); }
  });
  function deleteElement() {
    if (!selectedId) return;
    snapshot();
    const gone = new Set([selectedId, ...extraSel]);
    // Connectors to deleted shapes stay where they were, unattached.
    slide().elements.forEach(c => {
      if (c.type !== 'connector' || gone.has(c.id)) return;
      const [[x1, y1], [x2, y2]] = connectorPoints(c, slide());
      if (c.from && gone.has(c.from.id)) { delete c.from; c.x1 = x1; c.y1 = y1; }
      if (c.to && gone.has(c.to.id)) { delete c.to; c.x2 = x2; c.y2 = y2; }
    });
    slide().elements = slide().elements.filter(e => !gone.has(e.id));
    extraSel.clear();
    selectedId = editingId = null;
    renderStage(); refreshThumbSoon(); state.markDirty();
  }
  function duplicateElement() {
    const el = selectedId && elById(selectedId);
    if (!el) return;
    snapshot();
    const c = JSON.parse(JSON.stringify(el));
    c.id = newId(); c.x += 20; c.y += 20; delete c.ph;
    slide().elements.push(c);
    selectedId = c.id;
    renderStage(); refreshThumbSoon(); state.markDirty();
  }
  // ---------- Group, rotate, flip ----------
  function groupSelected() {
    const items = [selectedId, ...extraSel].map(elById).filter(x => x && x.type !== 'connector');
    if (items.length < 2) return UI.toast('Select two or more objects (Shift+click) to group them');
    snapshot();
    const x = Math.min(...items.map(i => i.x)), y = Math.min(...items.map(i => i.y));
    const g = { id: newId(), type: 'group', x, y, w: Math.max(...items.map(i => i.x + i.w)) - x, h: Math.max(...items.map(i => i.y + i.h)) - y,
      children: items.map(i => ({ ...i, x: i.x - x, y: i.y - y })) };
    const list = slide().elements;
    const at = Math.max(...items.map(i => list.indexOf(i)));
    list.splice(at + 1, 0, g);
    // Connectors pointing at grouped shapes now point at the group.
    const ids = new Set(items.map(i => i.id));
    list.forEach(c => { if (c.type === 'connector') ['from', 'to'].forEach(k => { if (c[k] && ids.has(c[k].id)) c[k].id = g.id; }); });
    slide().elements = list.filter(e => !items.includes(e));
    extraSel.clear(); selectedId = g.id;
    renderStage(); refreshThumbSoon(); state.markDirty();
  }
  function ungroup() {
    const g = selectedId && elById(selectedId);
    if (!g || g.type !== 'group') return UI.toast('Select a group to ungroup');
    snapshot();
    const list = slide().elements, at = list.indexOf(g);
    const cx = g.x + g.w / 2, cy = g.y + g.h / 2, rad = (g.rot || 0) * Math.PI / 180;
    const kids = g.children.map(c => {
      // Keep each child where it appears when the group is turned.
      let nx = g.x + c.x, ny = g.y + c.y;
      if (g.rot) { const ox = nx + c.w / 2 - cx, oy = ny + c.h / 2 - cy; nx = cx + ox * Math.cos(rad) - oy * Math.sin(rad) - c.w / 2; ny = cy + ox * Math.sin(rad) + oy * Math.cos(rad) - c.h / 2; }
      const out = { ...c, id: newId(), x: Math.round(nx), y: Math.round(ny) };
      if (g.rot) out.rot = (((c.rot || 0) + g.rot + 540) % 360) - 180 || undefined;
      return out;
    });
    list.splice(at, 1, ...kids);
    selectedId = kids[0].id; extraSel.clear(); kids.slice(1).forEach(k => extraSel.add(k.id));
    renderStage(); refreshThumbSoon(); state.markDirty();
  }
  document.getElementById('groupEls').addEventListener('click', groupSelected);
  document.getElementById('ungroupEls').addEventListener('click', ungroup);
  document.getElementById('rotRight').addEventListener('click', () => withSelected(el => { el.rot = ((((el.rot || 0) + 90) + 540) % 360) - 180; if (!el.rot) delete el.rot; }));
  document.getElementById('rotLeft').addEventListener('click', () => withSelected(el => { el.rot = ((((el.rot || 0) - 90) + 540) % 360) - 180; if (!el.rot) delete el.rot; }));
  document.getElementById('flipH').addEventListener('click', () => withSelected(el => { el.flipH = !el.flipH; if (!el.flipH) delete el.flipH; }));
  document.getElementById('flipV').addEventListener('click', () => withSelected(el => { el.flipV = !el.flipV; if (!el.flipV) delete el.flipV; }));

  // ---------- Connectors ----------
  let connectMode = null; // { kind, from }
  document.querySelectorAll('[data-connector]').forEach(b => b.addEventListener('click', () => {
    commitEditing();
    connectMode = { kind: b.dataset.connector, from: null };
    document.body.classList.add('connecting');
    UI.toast('Click the shape to connect from (Esc to cancel)', 3000);
  }));
  function connectClick(id) {
    const el = id && elById(id);
    if (!el || el.type === 'connector') { UI.toast('Click a shape'); return; }
    if (!connectMode.from) { connectMode.from = id; UI.toast('Now click the shape to connect to', 3000); return; }
    if (id === connectMode.from) return;
    snapshot();
    const c = { id: newId(), type: 'connector', kind: connectMode.kind, arrow: 'end', from: { id: connectMode.from, site: 'auto' }, to: { id, site: 'auto' }, width: 2, x: 0, y: 0, w: 1, h: 1 };
    slide().elements.push(c);
    endConnect();
    selectedId = c.id;
    renderStage(); refreshThumbSoon(); state.markDirty();
  }
  function endConnect() { connectMode = null; document.body.classList.remove('connecting'); }
  document.addEventListener('keydown', e => { if (connectMode && e.key === 'Escape') { endConnect(); UI.toast('Connector cancelled'); } });

  document.getElementById('delEl').addEventListener('click', deleteElement);
  document.getElementById('dupEl').addEventListener('click', duplicateElement);
  document.getElementById('toFront').addEventListener('click', () => reorder(1));
  document.getElementById('toBack').addEventListener('click', () => reorder(-1));
  function reorder(dir) {
    const el = selectedId && elById(selectedId);
    if (!el) return UI.toast('Select an object first');
    snapshot();
    const list = slide().elements.filter(e => e !== el);
    slide().elements = dir > 0 ? [...list, el] : [el, ...list];
    renderStage(); refreshThumbSoon(); state.markDirty();
  }
  document.getElementById('alignCenterH').addEventListener('click', () => withSelected(el => { el.x = Math.round((W - el.w) / 2); }));
  document.getElementById('alignCenterV').addEventListener('click', () => withSelected(el => { el.y = Math.round((H - el.h) / 2); }));
  function withSelected(fn) {
    const el = selectedId && elById(selectedId);
    if (!el) return UI.toast('Select an object first');
    snapshot();
    fn(el);
    renderStage(); refreshThumbSoon(); state.markDirty();
  }

  // ---------- Insert ----------
  function escapeHtml(s) { return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
  function addText(html = '') {
    commitEditing();
    snapshot();
    const el = { id: newId(), type: 'text', x: 280, y: 220, w: 400, h: 80, html, size: 20, align: 'left', valign: 'top' };
    slide().elements.push(el);
    if (html) { selectedId = el.id; renderStage(); } else startEditing(el.id);
    refreshThumbSoon(); state.markDirty();
  }
  document.getElementById('insText').addEventListener('click', () => addText());
  function addImage(src) {
    const img = new Image();
    img.onload = () => {
      snapshot();
      let w = img.naturalWidth, h = img.naturalHeight;
      const k = Math.min(1, 600 / w, 400 / h);
      w = Math.round(w * k); h = Math.round(h * k);
      const el = { id: newId(), type: 'image', src, x: Math.round((W - w) / 2), y: Math.round((H - h) / 2), w, h };
      slide().elements.push(el);
      selectedId = el.id;
      renderStage(); refreshThumbSoon(); state.markDirty();
    };
    img.src = src;
  }
  async function pickImage() {
    const f = await Platform.open([{ name: 'Images', extensions: ['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'svg'] }]);
    if (!f) return null;
    const ext = Platform.extOf(f.name);
    return Platform.bytesToDataUrl(f.data, ext === 'svg' ? 'image/svg+xml' : ext === 'jpg' ? 'image/jpeg' : 'image/' + ext);
  }
  document.getElementById('insImage').addEventListener('click', async () => { const src = await pickImage(); if (src) addImage(src); });
  // ---------- Video and audio ----------
  /** The first frame of a video (at 0.5 s), as a JPEG data URL for its poster. */
  function videoPoster(url) {
    return new Promise(resolve => {
      const v = document.createElement('video');
      v.muted = true; v.preload = 'auto'; v.src = url;
      const done = x => { v.removeAttribute('src'); v.load(); resolve(x); };
      v.addEventListener('loadeddata', () => { v.currentTime = Math.min(0.5, (v.duration || 1) / 2); }, { once: true });
      v.addEventListener('seeked', () => {
        try {
          const c = document.createElement('canvas');
          c.width = Math.min(1280, v.videoWidth || 640); c.height = Math.round(c.width * (v.videoHeight || 360) / (v.videoWidth || 640));
          c.getContext('2d').drawImage(v, 0, 0, c.width, c.height);
          done({ poster: c.toDataURL('image/jpeg', 0.8), w: v.videoWidth, h: v.videoHeight });
        } catch { done({}); }
      }, { once: true });
      v.addEventListener('error', () => done({}), { once: true });
      setTimeout(() => done({}), 8000);
    });
  }
  async function insertMedia(kind) {
    commitEditing();
    const exts = Object.keys(MEDIA_TYPES).filter(e => MEDIA_TYPES[e].startsWith(kind));
    const f = await Platform.open([{ name: kind === 'video' ? 'Videos' : 'Audio', extensions: exts }]);
    if (!f) return;
    const mime = MEDIA_TYPES[Platform.extOf(f.name)];
    if (!mime || !mime.startsWith(kind)) return UI.toast('That file type isn\'t supported');
    if (f.data.length > 200e6) return UI.toast('Media files up to 200 MB can be inserted', 4000);
    UI.toast('Inserting ' + kind + '…');
    const src = await Platform.bytesToDataUrl(f.data, mime);
    let el;
    if (kind === 'video') {
      const meta = await videoPoster(URL.createObjectURL(new Blob([f.data], { type: mime })));
      let w = meta.w || 640, h = meta.h || 360;
      const k = Math.min(640 / w, 400 / h); w = Math.round(w * k); h = Math.round(h * k);
      el = { id: newId(), type: 'media', kind, src, ...(meta.poster ? { poster: meta.poster } : {}), x: Math.round((W - w) / 2), y: Math.round((H - h) / 2), w, h };
    } else el = { id: newId(), type: 'media', kind, src, x: W - 120, y: H - 120, w: 64, h: 64 };
    snapshot();
    slide().elements.push(el);
    selectedId = el.id;
    renderStage(); refreshThumbSoon(); state.markDirty();
  }
  document.getElementById('insVideo').addEventListener('click', () => insertMedia('video'));
  document.getElementById('insAudio').addEventListener('click', () => insertMedia('audio'));
  ['mediaAuto', 'mediaLoop', 'mediaHide'].forEach(id => document.getElementById(id).addEventListener('change', e => {
    const el = selectedId && elById(selectedId);
    if (!el || el.type !== 'media') return UI.toast('Select a video or audio clip first');
    snapshot();
    el[{ mediaAuto: 'autoplay', mediaLoop: 'loop', mediaHide: 'hideIcon' }[id]] = e.target.checked;
    renderStage(); state.markDirty();
  }));
  document.getElementById('mediaPlay').addEventListener('click', () => {
    const node = selectedId && stage.querySelector(`[data-id="${selectedId}"] video, [data-id="${selectedId}"] audio`);
    if (!node) return UI.toast('Select a video or audio clip first');
    if (node.paused) node.play().catch(() => {}); else node.pause();
  });
  document.getElementById('insTable').addEventListener('click', async () => {
    const r = await UI.form('Insert Table', [
      { name: 'rows', label: 'Rows', type: 'number', value: 3, min: 1, max: 15 },
      { name: 'cols', label: 'Columns', type: 'number', value: 3, min: 1, max: 10 }
    ], 'Insert');
    if (!r) return;
    const rows = Math.max(1, Math.min(15, +r.rows)), cols = Math.max(1, Math.min(10, +r.cols));
    snapshot();
    const html = '<table>' + Array.from({ length: rows }, (_, i) => '<tr>' + Array.from({ length: cols }, (_, j) => `<td>${i === 0 ? 'Header ' + (j + 1) : ''}</td>`).join('') + '</tr>').join('') + '</table>';
    const el = { id: newId(), type: 'table', x: 80, y: 140, w: 800, h: Math.min(380, rows * 44), html, size: 16, align: 'left', valign: 'top' };
    slide().elements.push(el);
    selectedId = el.id;
    renderStage(); refreshThumbSoon(); state.markDirty();
  });
  document.querySelectorAll('[data-shape]').forEach(b => b.addEventListener('click', () => {
    commitEditing();
    snapshot();
    const shape = b.dataset.shape;
    const isLine = shape === 'line';
    const el = { id: newId(), type: 'shape', shape, x: 380, y: 190, w: isLine ? 200 : 200, h: isLine ? 4 : 160, html: '', size: 18, align: 'center', valign: 'middle', color: '#ffffff',
      fill: isLine ? null : theme().accent, line: isLine ? theme().accent : shade(theme().accent, -0.25) };
    if (isLine) { el.h = 120; }
    slide().elements.push(el);
    selectedId = el.id;
    renderStage(); refreshThumbSoon(); state.markDirty();
  }));
  function shade(hex, amt) {
    const n = parseInt(hex.slice(1), 16);
    const f = c => Math.max(0, Math.min(255, Math.round(c + (amt < 0 ? c * amt : (255 - c) * amt))));
    return '#' + [n >> 16, (n >> 8) & 255, n & 255].map(f).map(v => v.toString(16).padStart(2, '0')).join('');
  }

  // ---------- Text formatting ----------
  function wholeBox(cmd, val) {
    // Without an active caret, apply formatting to the whole selected box.
    const el = selectedId && elById(selectedId);
    if (!el || el.type === 'image') return UI.toast('Select a text box first');
    snapshot();
    const c = stage.querySelector(`[data-id="${el.id}"] .content`);
    c.contentEditable = 'true';
    c.focus();
    const r = document.createRange();
    r.selectNodeContents(c);
    getSelection().removeAllRanges();
    getSelection().addRange(r);
    document.execCommand(cmd, false, val);
    el.html = c.innerHTML;
    c.contentEditable = 'false';
    getSelection().removeAllRanges();
    renderStage(); refreshThumbSoon(); state.markDirty();
  }
  function textCmd(cmd, val) {
    if (editingId) { document.execCommand(cmd, false, val); const el = elById(editingId); el.html = stage.querySelector(`[data-id="${editingId}"] .content`).innerHTML; state.markDirty(); refreshThumbSoon(); }
    else wholeBox(cmd, val);
  }
  document.querySelectorAll('[data-exec]').forEach(b => {
    b.addEventListener('mousedown', e => e.preventDefault());
    b.addEventListener('click', () => textCmd(b.dataset.exec));
  });
  document.getElementById('textColor').addEventListener('change', e => {
    const el = selectedId && elById(selectedId);
    if (!el) return;
    if (editingId && !getSelection().isCollapsed) textCmd('foreColor', e.target.value);
    else { snapshot(); el.color = e.target.value; renderStage(); refreshThumbSoon(); state.markDirty(); }
  });
  document.getElementById('textColor').addEventListener('input', e => { e.target.parentElement.querySelector('.bar').style.background = e.target.value; });
  document.getElementById('fontName').addEventListener('change', e => {
    const el = selectedId && elById(selectedId);
    if (!el) return;
    if (editingId && !getSelection().isCollapsed) textCmd('fontName', e.target.value);
    else { snapshot(); el.font = e.target.value; renderStage(); refreshThumbSoon(); state.markDirty(); }
  });
  document.getElementById('fontSize').addEventListener('change', e => {
    const el = selectedId && elById(selectedId);
    if (!el) return;
    const pt = +e.target.value;
    if (editingId && !getSelection().isCollapsed) {
      document.execCommand('styleWithCSS', false, false);
      document.execCommand('fontSize', false, '7');
      document.execCommand('styleWithCSS', false, true);
      const c = stage.querySelector(`[data-id="${editingId}"] .content`);
      c.querySelectorAll('font[size="7"]').forEach(f => { const s = document.createElement('span'); s.style.fontSize = pt + 'pt'; s.innerHTML = f.innerHTML; f.replaceWith(s); });
      el.html = c.innerHTML;
      state.markDirty();
    } else {
      snapshot();
      el.size = pt;
      // Clear per-run sizes so the box size wins.
      const tmp = document.createElement('div'); tmp.innerHTML = el.html || '';
      tmp.querySelectorAll('[style]').forEach(n => n.style.removeProperty('font-size'));
      el.html = tmp.innerHTML;
      renderStage(); refreshThumbSoon(); state.markDirty();
    }
  });
  document.querySelectorAll('[data-align]').forEach(b => b.addEventListener('click', () => withSelected(el => { el.align = b.dataset.align; })));
  document.querySelectorAll('[data-valign]').forEach(b => b.addEventListener('click', () => withSelected(el => { el.valign = b.dataset.valign; })));
  document.getElementById('shapeFill').addEventListener('change', e => withSelected(el => { if (el.type === 'image') return; el.fill = e.target.value; if (el.type === 'text') { el.type = 'shape'; el.shape = 'rect'; } }));
  document.getElementById('shapeLine').addEventListener('change', e => withSelected(el => { if (el.type === 'image') return; el.line = e.target.value; if (el.type === 'text') { el.type = 'shape'; el.shape = 'rect'; } }));
  document.getElementById('noFill').addEventListener('click', e => { e.preventDefault(); withSelected(el => { el.fill = null; }); });
  document.getElementById('noLine').addEventListener('click', e => { e.preventDefault(); withSelected(el => { el.line = null; }); });
  function syncToolbar() {
    const el = selectedId && elById(selectedId);
    document.querySelectorAll('[data-align]').forEach(b => b.classList.toggle('on', !!el && (el.align || 'left') === b.dataset.align));
    document.querySelectorAll('[data-valign]').forEach(b => b.classList.toggle('on', !!el && (el.valign || 'top') === b.dataset.valign));
    if (el && el.size) document.getElementById('fontSize').value = String(el.size);
    if (el && el.type !== 'image') document.getElementById('fontName').value = fontOf(el);
    document.getElementById('slideBg').value = slide().bg || theme().bg;
    document.querySelectorAll('[data-trans]').forEach(b => b.classList.toggle('on', (slide().transition || 'none') === b.dataset.trans));
    document.getElementById('transDur').value = slide().transDur || 1;
    document.getElementById('advanceOn').checked = typeof slide().advance === 'number';
    document.getElementById('advanceAfter').value = typeof slide().advance === 'number' ? slide().advance : 0;
    const media = el && el.type === 'media';
    document.getElementById('mediaAuto').checked = !!(media && el.autoplay);
    document.getElementById('mediaLoop').checked = !!(media && el.loop);
    document.getElementById('mediaHide').checked = !!(media && el.hideIcon);
    ['mediaAuto', 'mediaLoop', 'mediaPlay'].forEach(id => { document.getElementById(id).disabled = !media; });
    document.getElementById('mediaHide').disabled = !(media && el.kind === 'audio');
    const a = el && el.anim;
    document.getElementById('animEffect').value = a ? a.effect : '';
    ['animDir', 'animStart', 'animDur', 'animDelay', 'animEarlier', 'animLater'].forEach(id => { document.getElementById(id).disabled = !a; });
    if (a) { document.getElementById('animDir').value = a.dir; document.getElementById('animStart').value = a.start; document.getElementById('animDur').value = a.dur; document.getElementById('animDelay').value = a.delay; document.getElementById('animDir').disabled = !SlidesAnim.EFFECTS[a.effect].dirs; }
    fillLayoutSelect();
    document.getElementById('layout').value = layoutKeys().includes(slide().layout) ? slide().layout : 'content';
  }

  // ---------- Design ----------
  const themesBox = document.getElementById('themes');
  Object.entries(THEMES).forEach(([key, t]) => {
    const b = document.createElement('button');
    b.className = 'theme-swatch';
    b.style.background = t.bg;
    b.style.color = t.text;
    b.style.borderBottom = `4px solid ${t.accent}`;
    b.innerHTML = `<span class="t" style="color:${t.title};font-family:${t.titleFont}">Aa</span><span>${t.name}</span>`;
    b.addEventListener('click', () => {
      snapshot();
      pres.theme = key;
      pres.slides.forEach(s => { s.bg = null; });
      renderAll();
      state.markDirty();
    });
    themesBox.appendChild(b);
  });
  document.getElementById('slideBg').addEventListener('change', e => { snapshot(); slide().bg = e.target.value; slide().bgImage = null; renderAll(); state.markDirty(); });
  document.getElementById('bgAll').addEventListener('click', () => { snapshot(); const { bg, bgImage } = slide(); pres.slides.forEach(s => { s.bg = bg; s.bgImage = bgImage; }); renderAll(); state.markDirty(); });
  document.getElementById('bgImage').addEventListener('click', async () => { const src = await pickImage(); if (!src) return; snapshot(); slide().bgImage = src; renderAll(); state.markDirty(); });
  document.querySelectorAll('[data-trans]').forEach(b => b.addEventListener('click', () => {
    snapshot();
    slide().transition = b.dataset.trans;
    syncToolbar(); state.markDirty();
    previewTransition();
  }));
  document.getElementById('transAll').addEventListener('click', () => { snapshot(); const { transition, transDur, advance } = slide(); pres.slides.forEach(s => Object.assign(s, { transition, transDur, advance })); state.markDirty(); UI.toast('Applied to all slides'); });
  document.getElementById('transDur').addEventListener('change', e => { snapshot(); slide().transDur = Math.max(0.1, Math.min(30, +e.target.value || 1)); state.markDirty(); });
  document.getElementById('advanceOn').addEventListener('change', e => { snapshot(); slide().advance = e.target.checked ? Math.max(0, +document.getElementById('advanceAfter').value || 0) : null; state.markDirty(); });
  document.getElementById('advanceAfter').addEventListener('change', e => { snapshot(); if (document.getElementById('advanceOn').checked) slide().advance = Math.max(0, +e.target.value || 0); state.markDirty(); });
  function previewTransition() {
    const frames = TRANSITION_FRAMES[slide().transition];
    if (frames) stage.animate(frames, { duration: (slide().transDur || 1) * 1000, easing: 'ease-out' });
  }

  // ---------- Animations ----------
  /** The click an element's animation plays on (0 = as the slide appears), as PowerPoint numbers them. */
  function animNumber(s, el) {
    let n = 0;
    for (const x of s.elements.filter(e => e.anim).sort((a, b) => a.anim.order - b.anim.order)) {
      if (x.anim.start === 'click') n++;
      if (x === el) return n;
    }
    return n;
  }
  function setAnim(fn) {
    const el = selectedId && elById(selectedId);
    if (!el) { UI.toast('Select an object to animate'); return; }
    snapshot();
    fn(el);
    if (el.anim) el.anim = SlidesAnim.clean(el.anim);
    renderAll(); state.markDirty();
  }
  document.getElementById('animEffect').addEventListener('change', e => {
    const v = e.target.value;
    setAnim(el => {
      if (!v) { delete el.anim; return; }
      const max = Math.max(-1, ...slide().elements.filter(x => x.anim && x !== el).map(x => x.anim.order));
      el.anim = { ...(el.anim || { start: 'click', dur: v === 'appear' ? 0 : 0.5, delay: 0, dir: 'bottom', order: max + 1 }), effect: v };
      if (v === 'appear') el.anim.dur = 0; else if (!el.anim.dur) el.anim.dur = 0.5;
    });
    previewAnim();
  });
  document.getElementById('animDir').addEventListener('change', e => setAnim(el => { if (el.anim) el.anim.dir = e.target.value; }));
  document.getElementById('animStart').addEventListener('change', e => setAnim(el => { if (el.anim) el.anim.start = e.target.value; }));
  document.getElementById('animDur').addEventListener('change', e => setAnim(el => { if (el.anim) el.anim.dur = +e.target.value || 0; }));
  document.getElementById('animDelay').addEventListener('change', e => setAnim(el => { if (el.anim) el.anim.delay = +e.target.value || 0; }));
  function moveAnim(d) {
    setAnim(el => {
      if (!el.anim) return;
      const list = slide().elements.filter(x => x.anim).sort((a, b) => a.anim.order - b.anim.order);
      const i = list.indexOf(el), j = i + d;
      if (j < 0 || j >= list.length) return;
      [list[i], list[j]] = [list[j], list[i]];
      list.forEach((x, k) => { x.anim.order = k; });
    });
  }
  document.getElementById('animEarlier').addEventListener('click', () => moveAnim(-1));
  document.getElementById('animLater').addEventListener('click', () => moveAnim(1));
  /** Plays the slide's builds one after another on the editing stage. */
  function previewAnim() {
    const groups = SlidesAnim.sequence(slide());
    if (!groups.length) return;
    let t = 0;
    const timeline = { items: [] };
    groups.forEach(g => {
      g.items.forEach(x => timeline.items.push({ ...x, at: x.at + t }));
      t += Math.max(...g.items.map(x => x.at + x.anim.dur * 1000)) + 200;
    });
    SlidesAnim.play(timeline, id => stage.querySelector(`[data-id="${id}"]`));
    setTimeout(() => stage.querySelectorAll('.el').forEach(n => { n.style.visibility = ''; }), t + 50);
  }
  document.getElementById('animPreview').addEventListener('click', previewAnim);
  document.querySelector('.ribbon-tabs').addEventListener('click', () => setTimeout(() => {
    document.body.classList.toggle('show-anims', document.querySelector('.ribbon-panel[data-panel="animations"]').classList.contains('active'));
    renderStage();
  }));
  notes.addEventListener('input', () => { slide().notes = notes.value; state.markDirty(); });
  document.getElementById('toggleNotes').addEventListener('click', e => { document.body.classList.toggle('hide-notes'); e.currentTarget.classList.toggle('on', !document.body.classList.contains('hide-notes')); fitStage(); });

  // ---------- Slide show ----------
  const presenter = document.getElementById('presenter');
  const presentSlide = document.getElementById('presentSlide');
  let showIndex = 0, cursorTimer;
  function present(from) {
    commitEditing();
    if (masterTarget) setMasterView(false);
    showIndex = from;
    presenter.hidden = false;
    presenter.classList.remove('black');
    Platform.fullscreen(true);
    showSlide(showIndex, false);
    setTimeout(() => showSlide(showIndex, false), 300); // refit after going full screen
    document.getElementById('presentHint').textContent = '→ / Space: next   ← : previous   B: black   Esc: end';
    setTimeout(() => { document.getElementById('presentHint').textContent = ''; }, 3000);
  }
  let groups = [], built = 0, running = [], advanceTimer = null;
  const showNode = id => presentSlide.querySelector(`[data-id="${id}"]`);
  const TRANSITION_FRAMES = {
    fade: [{ opacity: 0 }, { opacity: 1 }],
    push: [{ translate: '100% 0' }, { translate: '0 0' }],
    cover: [{ translate: '100% 0' }, { translate: '0 0' }],
    uncover: [{ opacity: 0, scale: '1.04' }, { opacity: 1, scale: '1' }],
    wipe: [{ clipPath: 'inset(0 100% 0 0)' }, { clipPath: 'inset(0 0 0 0)' }],
    split: [{ clipPath: 'inset(0 50% 0 50%)' }, { clipPath: 'inset(0 0 0 0)' }],
    zoom: [{ scale: '0.6', opacity: 0 }, { scale: '1', opacity: 1 }]
  };
  /** Shows slide i with its builds done up to `upTo` (0 = none, Infinity = all). */
  function showSlide(i, animate = true, upTo = 0) {
    finishRunning();
    clearTimeout(advanceTimer);
    const s = pres.slides[i];
    renderSlide(s, presentSlide, false);
    const k = Math.min(window.innerWidth / W, window.innerHeight / H);
    presentSlide.style.setProperty('--scale', `scale(${k})`);
    presentSlide.style.transform = `scale(${k})`;
    presentSlide.className = 'slide';
    groups = SlidesAnim.sequence(s);
    built = Math.min(upTo, groups.length);
    applyVisibility(s);
    startMedia(s);
    const frames = TRANSITION_FRAMES[s.transition];
    let wait = 0;
    if (animate && frames) { presentSlide.animate(frames, { duration: (s.transDur || 1) * 1000, easing: 'ease-out' }); wait = (s.transDur || 1) * 1000; }
    if (built === 0 && groups[0] && groups[0].auto) setTimeout(() => { if (built === 0 && pres.slides[showIndex] === s) build(); }, wait);
    else scheduleAdvance();
  }
  /** Media in the show: clips set to start automatically play; a click on a clip plays or pauses it. */
  function startMedia(s) {
    s.elements.filter(e => e.type === 'media').forEach(e => {
      const node = showNode(e.id);
      const m = node && node.querySelector('video, audio');
      if (!m) return;
      if (e.kind === 'audio' && e.hideIcon) node.querySelector('.media-audio-icon').style.visibility = 'hidden';
      node.addEventListener('click', ev => { ev.stopPropagation(); if (m.paused) m.play().catch(() => {}); else m.pause(); });
      if (e.autoplay) m.play().catch(() => {});
    });
  }
  function applyVisibility(s) {
    const vis = SlidesAnim.visibilityAfter(s, groups, built);
    vis.forEach((v, id) => { const n = showNode(id); if (n) n.style.visibility = v ? '' : 'hidden'; });
  }
  function finishRunning() { running.forEach(r => { try { r.a.finish(); } catch { /* gone */ } r.done(); }); running = []; }
  /** Plays the next build; returns false when there is none. */
  function build() {
    if (built >= groups.length) return false;
    finishRunning();
    running = SlidesAnim.play(groups[built], showNode);
    built++;
    // A following group that starts "with/after previous" can't exist (groups start on a click), so wait for a click or the timer.
    scheduleAdvance();
    return true;
  }
  function scheduleAdvance() {
    clearTimeout(advanceTimer);
    if (AUDIENCE) return; // the presenter's window decides when to move on
    const s = pres.slides[showIndex];
    if (typeof s.advance !== 'number' || built < groups.length) return;
    const end = Math.max(0, ...running.map(r => { const t = r.a.effect.getComputedTiming(); return (t.endTime || 0); }));
    advanceTimer = setTimeout(() => { if (!presenter.hidden) step(1); }, end + s.advance * 1000);
  }
  function endShow() {
    finishRunning();
    clearTimeout(advanceTimer);
    presenter.hidden = true;
    Platform.fullscreen(false);
    goTo(showIndex);
  }
  function step(d) {
    if (d > 0 && build()) return;
    if (d < 0 && built > 0) { running.forEach(r => { r.a.dropped = true; r.a.cancel(); }); running = []; built--; applyVisibility(pres.slides[showIndex]); return; }
    const n = showIndex + d;
    if (n >= pres.slides.length) return endShow();
    if (n < 0) return;
    showIndex = n;
    showSlide(n, d > 0, d > 0 ? 0 : Infinity);
  }
  presenter.addEventListener('click', e => { if (e.button !== 0) return; if (AUDIENCE) audience.toPresenter({ cmd: 'key', key: 'ArrowRight' }); else step(1); });
  presenter.addEventListener('contextmenu', e => { e.preventDefault(); if (AUDIENCE) audience.toPresenter({ cmd: 'key', key: 'ArrowLeft' }); else step(-1); });
  presenter.addEventListener('mousemove', () => { presenter.classList.add('show-cursor'); clearTimeout(cursorTimer); cursorTimer = setTimeout(() => presenter.classList.remove('show-cursor'), 1500); });
  document.addEventListener('keydown', e => {
    if (presenter.hidden) return;
    e.preventDefault();
    const k = e.key;
    if (AUDIENCE) { audience.toPresenter({ cmd: 'key', key: k }); return; }
    if (['ArrowRight', 'ArrowDown', ' ', 'PageDown', 'Enter', 'n', 'N'].includes(k)) step(1);
    else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace', 'p', 'P'].includes(k)) step(-1);
    else if (k === 'Home') { showIndex = 0; showSlide(0); }
    else if (k === 'End') { showIndex = pres.slides.length - 1; showSlide(showIndex); }
    else if (k === 'b' || k === 'B' || k === '.') presenter.classList.toggle('black');
    else if (k === 'Escape') endShow();
  }, true);
  window.addEventListener('resize', () => { if (!presenter.hidden) showSlide(showIndex, false); });
  document.addEventListener('fullscreenchange', () => { if (!document.fullscreenElement && !presenter.hidden && !Platform.desktop) endShow(); });

  // ---------- Presenter View ----------
  // The presenter's window shows the current slide, the next one, notes and a timer; the
  // audience window (another screen when there is one) shows the slides and follows along.
  const audience = {
    win: null,
    open() {
      if (Platform.desktop) return window.atlas.showOpen();
      this.win = window.open('slides.html?audience=1', 'atlas-audience', 'popup,width=960,height=580');
      return !!this.win;
    },
    send(msg) {
      if (Platform.desktop) return window.atlas.showSend(msg);
      if (this.win && !this.win.closed) this.win.postMessage(msg, location.origin);
    },
    close() { if (Platform.desktop) window.atlas.showClose(); else if (this.win && !this.win.closed) this.win.close(); this.win = null; },
    listen(cb) {
      if (Platform.desktop) { window.atlas.onShowMessage(cb); return; }
      window.addEventListener('message', e => {
        if (e.origin !== location.origin || !e.data || typeof e.data.cmd !== 'string') return;
        if (AUDIENCE ? e.source !== window.opener : e.source !== this.win) return;
        cb(e.data);
      });
    },
    toPresenter(msg) { if (Platform.desktop) window.atlas.showSend(msg); else if (window.opener) window.opener.postMessage(msg, location.origin); }
  };
  const pv = { el: document.getElementById('presenterView'), index: 0, built: 0, groups: [], started: 0, paused: 0, pausedAt: 0, timer: null, advance: null, noteSize: 18 };
  const pvq = id => document.getElementById(id);
  async function presenterView(from) {
    commitEditing();
    pv.index = from;
    pv.started = Date.now(); pv.paused = 0; pv.pausedAt = 0;
    pv.el.hidden = false;
    document.body.classList.add('presenting');
    pvShow(0);
    clearInterval(pv.timer);
    pv.timer = setInterval(pvClock, 500);
    pvClock();
    try { await audience.open(); } catch { /* the presenter view still works for rehearsal */ }
  }
  /** Shows pv.index with `upTo` builds done (Infinity = all) and tells the audience window. */
  function pvShow(upTo, animate = true) {
    clearTimeout(pv.advance);
    const s = pres.slides[pv.index];
    pv.groups = SlidesAnim.sequence(s);
    // A first build that starts by itself has already played once the slide is up.
    const auto = pv.groups[0] && pv.groups[0].auto ? 1 : 0;
    pv.built = Math.max(auto, Math.min(upTo, pv.groups.length));
    pvRender();
    audience.send({ cmd: 'goto', index: pv.index, built: pv.built === auto ? 0 : pv.built, animate });
    pvScheduleAdvance();
  }
  function pvRender() {
    const s = pres.slides[pv.index];
    const cur = pvq('pvCurrent'), next = pvq('pvNext');
    renderSlide(s, cur, false);
    SlidesAnim.visibilityAfter(s, pv.groups, pv.built).forEach((v, id) => { const n = cur.querySelector(`[data-id="${id}"]`); if (n) n.style.visibility = v ? '' : 'hidden'; });
    // "Next": the next build on this slide, else the next slide.
    if (pv.built < pv.groups.length) {
      renderSlide(s, next, false);
      SlidesAnim.visibilityAfter(s, pv.groups, pv.built + 1).forEach((v, id) => { const n = next.querySelector(`[data-id="${id}"]`); if (n) n.style.visibility = v ? '' : 'hidden'; });
      pvq('pvNextLabel').textContent = 'Next: animation';
    } else if (pres.slides[pv.index + 1]) {
      const ns = pres.slides[pv.index + 1];
      renderSlide(ns, next, false);
      SlidesAnim.visibilityAfter(ns, SlidesAnim.sequence(ns), 0).forEach((v, id) => { const n = next.querySelector(`[data-id="${id}"]`); if (n) n.style.visibility = v ? '' : 'hidden'; });
      pvq('pvNextLabel').textContent = 'Next slide';
    } else { next.innerHTML = ''; next.style.background = '#000'; pvq('pvNextLabel').textContent = 'End of slide show'; }
    pvq('pvNotes').textContent = s.notes || 'No notes for this slide.';
    pvq('pvNotes').classList.toggle('empty', !s.notes);
    pvq('pvCount').textContent = `Slide ${pv.index + 1} of ${pres.slides.length}`;
    pvFit();
  }
  function pvFit() {
    [['pvCurrent', 'pvCurrentBox'], ['pvNext', 'pvNextBox']].forEach(([id, box]) => {
      const b = pvq(box).getBoundingClientRect();
      const k = Math.min(b.width / W, b.height / H) || 0.1;
      Object.assign(pvq(id).style, { transform: `scale(${k})`, left: (b.width - W * k) / 2 + 'px', top: (b.height - H * k) / 2 + 'px' });
    });
  }
  function pvStep(d) {
    if (pv.el.hidden) return;
    if (d > 0) {
      if (pv.built < pv.groups.length) { pv.built++; pvRender(); audience.send({ cmd: 'build' }); pvScheduleAdvance(); return; }
      if (pv.index + 1 >= pres.slides.length) return pvEnd();
      pv.index++; pvShow(0);
    } else {
      const auto = pv.groups[0] && pv.groups[0].auto ? 1 : 0;
      if (pv.built > auto) { pv.built--; pvRender(); audience.send({ cmd: 'goto', index: pv.index, built: pv.built, animate: false }); return; }
      if (pv.index === 0) return;
      pv.index--; pvShow(Infinity, false);
    }
  }
  function pvScheduleAdvance() {
    clearTimeout(pv.advance);
    const s = pres.slides[pv.index];
    if (typeof s.advance !== 'number' || pv.built < pv.groups.length) return;
    const last = pv.groups[pv.built - 1];
    const tail = last ? Math.max(0, ...last.items.map(x => x.at + x.anim.dur * 1000)) : (s.transition && s.transition !== 'none' ? (s.transDur || 1) * 1000 : 0);
    pv.advance = setTimeout(() => pvStep(1), tail + s.advance * 1000);
  }
  function pvClock() {
    const now = Date.now();
    const ms = (pv.pausedAt || now) - pv.started - pv.paused;
    const t = Math.max(0, Math.floor(ms / 1000));
    pvq('pvTimer').textContent = `${String(Math.floor(t / 3600)).padStart(1, '0')}:${String(Math.floor(t / 60) % 60).padStart(2, '0')}:${String(t % 60).padStart(2, '0')}`;
    pvq('pvClock').textContent = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
  }
  function pvEnd() {
    clearInterval(pv.timer); clearTimeout(pv.advance);
    pv.el.hidden = true;
    document.body.classList.remove('presenting');
    audience.send({ cmd: 'end' });
    audience.close();
    goTo(pv.index);
  }
  pvq('pvPrev').addEventListener('click', () => pvStep(-1));
  pvq('pvNextBtn').addEventListener('click', () => pvStep(1));
  pvq('pvBlack').addEventListener('click', e => { const on = e.currentTarget.classList.toggle('on'); pvq('pvCurrentBox').classList.toggle('black', on); audience.send({ cmd: 'black', on }); });
  pvq('pvEnd').addEventListener('click', pvEnd);
  pvq('pvPause').addEventListener('click', e => {
    if (pv.pausedAt) { pv.paused += Date.now() - pv.pausedAt; pv.pausedAt = 0; e.currentTarget.textContent = '⏸'; }
    else { pv.pausedAt = Date.now(); e.currentTarget.textContent = '▶'; }
    pvClock();
  });
  pvq('pvReset').addEventListener('click', () => { pv.started = Date.now(); pv.paused = 0; pv.pausedAt = pv.pausedAt ? Date.now() : 0; pvClock(); });
  pvq('pvZoomIn').addEventListener('click', () => { pv.noteSize = Math.min(48, pv.noteSize + 2); pvq('pvNotes').style.fontSize = pv.noteSize + 'px'; });
  pvq('pvZoomOut').addEventListener('click', () => { pv.noteSize = Math.max(10, pv.noteSize - 2); pvq('pvNotes').style.fontSize = pv.noteSize + 'px'; });
  pvq('pvCurrentBox').addEventListener('click', () => pvStep(1));
  document.addEventListener('keydown', e => {
    if (pv.el.hidden) return;
    e.preventDefault(); e.stopPropagation();
    pvKey(e.key);
  }, true);
  function pvKey(k) {
    if (['ArrowRight', 'ArrowDown', ' ', 'PageDown', 'Enter', 'n', 'N'].includes(k)) pvStep(1);
    else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace', 'p', 'P'].includes(k)) pvStep(-1);
    else if (k === 'Home') { pv.index = 0; pvShow(0); }
    else if (k === 'End') { pv.index = pres.slides.length - 1; pvShow(0); }
    else if (k === 'b' || k === 'B' || k === '.') pvq('pvBlack').click();
    else if (k === 'Escape') pvEnd();
  }
  window.addEventListener('resize', () => { if (!pv.el.hidden) pvFit(); });
  audience.listen(msg => {
    if (AUDIENCE) return audienceMessage(msg);
    if (msg.cmd === 'ready') audience.send({ cmd: 'load', pres: JSON.parse(JSON.stringify(pres)), index: pv.index, built: pv.built });
    else if (msg.cmd === 'key') pvKey(msg.key);
    else if (msg.cmd === 'closed' && !pv.el.hidden) UI.toast('The audience window was closed', 3000);
  });
  /** Audience window: shows what the presenter's window says. */
  function audienceMessage(msg) {
    switch (msg.cmd) {
      case 'load':
        try { pres = cleanPres(msg.pres); } catch { return; }
        showIndex = Math.max(0, Math.min(pres.slides.length - 1, msg.index | 0));
        presenter.hidden = false;
        showSlide(showIndex, false, msg.built === Infinity || msg.built > 1000 ? Infinity : msg.built | 0);
        break;
      case 'goto':
        showIndex = Math.max(0, Math.min(pres.slides.length - 1, msg.index | 0));
        showSlide(showIndex, !!msg.animate, msg.built === Infinity || msg.built > 1000 ? Infinity : msg.built | 0);
        break;
      case 'build': build(); break;
      case 'black': presenter.classList.toggle('black', !!msg.on); break;
      case 'end': window.close(); break;
    }
  }

  // ---------- PPTX export ----------
  const PX_PER_IN = 96;
  const inch = px => +(px / PX_PER_IN).toFixed(3);
  const hex = c => {
    if (!c) return undefined;
    if (c.startsWith('#')) return (c.length === 4 ? c.slice(1).split('').map(x => x + x).join('') : c.slice(1, 7)).toUpperCase();
    const m = c.match(/\d+/g);
    return m ? m.slice(0, 3).map(v => (+v).toString(16).padStart(2, '0')).join('').toUpperCase() : undefined;
  };
  /** Converts a text box's HTML into PptxGenJS text runs, keeping per-run formatting and bullets. */
  function htmlToRuns(html, base) {
    const root = document.createElement('div');
    root.innerHTML = html || '';
    const paras = [];
    let cur = null;
    const startPara = (opts = {}) => { cur = { runs: [], opts }; paras.push(cur); };
    const walk = (node, fmt, paraOpts) => {
      if (node.nodeType === 3) {
        if (!cur) startPara(paraOpts);
        const text = node.nodeValue.replace(/\u00a0/g, ' ');
        if (text) cur.runs.push({ text, fmt: { ...fmt } });
        return;
      }
      if (node.nodeType !== 1) return;
      const tag = node.tagName;
      const f = { ...fmt };
      const st = node.style || {};
      if (tag === 'B' || tag === 'STRONG' || +st.fontWeight >= 600 || st.fontWeight === 'bold') f.bold = true;
      if (tag === 'I' || tag === 'EM' || st.fontStyle === 'italic') f.italic = true;
      if (tag === 'U' || (st.textDecorationLine || st.textDecoration || '').includes('underline')) f.underline = true;
      if (tag === 'S' || tag === 'STRIKE' || (st.textDecorationLine || st.textDecoration || '').includes('line-through')) f.strike = true;
      if (st.color) f.color = hex(st.color);
      if (node.getAttribute && node.getAttribute('color')) f.color = hex(node.getAttribute('color'));
      if (st.fontSize && st.fontSize.endsWith('pt')) f.size = parseFloat(st.fontSize);
      else if (st.fontSize && st.fontSize.endsWith('px')) f.size = parseFloat(st.fontSize) * 0.75;
      if (st.fontFamily) f.font = st.fontFamily.split(',')[0].replace(/["']/g, '');
      if (tag === 'BR') { startPara(paraOpts); return; }
      if (tag === 'UL' || tag === 'OL') {
        [...node.children].forEach(li => { startPara({ bullet: tag === 'OL' ? { type: 'number' } : true }); [...li.childNodes].forEach(n => walk(n, f, cur.opts)); });
        cur = null;
        return;
      }
      if (tag === 'DIV' || tag === 'P' || /^H\d$/.test(tag)) {
        startPara(paraOpts);
        [...node.childNodes].forEach(n => walk(n, f, paraOpts));
        cur = null;
        return;
      }
      [...node.childNodes].forEach(n => walk(n, f, paraOpts));
    };
    [...root.childNodes].forEach(n => walk(n, {}, {}));
    const out = [];
    paras.forEach((p, i) => {
      const last = i === paras.length - 1;
      if (!p.runs.length) p.runs.push({ text: '', fmt: {} });
      p.runs.forEach((r, j) => {
        const options = { ...p.opts };
        if (r.fmt.bold) options.bold = true;
        if (r.fmt.italic) options.italic = true;
        if (r.fmt.underline) options.underline = { style: 'sng' };
        if (r.fmt.strike) options.strike = 'sngStrike';
        if (r.fmt.color) options.color = r.fmt.color;
        if (r.fmt.size) options.fontSize = r.fmt.size;
        if (r.fmt.font) options.fontFace = r.fmt.font;
        if (j === p.runs.length - 1 && !last) options.breakLine = true;
        if (j > 0) delete options.bullet;
        out.push({ text: r.text, options });
      });
    });
    return out.length ? out : [{ text: '', options: {} }];
  }
  async function toPptx() {
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_16x9';
    pptx.title = Platform.baseName(state.name);
    pptx.company = 'Atlas Office';
    for (const s of pres.slides) {
      const ps = pptx.addSlide();
      const L = pres.layouts[s.layout];
      const bgFrom = [s, L, !s.hideMaster && pres.master].find(x => x && (x.bg || x.bgImage)) || {};
      if (bgFrom.bgImage) ps.background = { data: bgFrom.bgImage };
      else ps.background = { color: hex(bgFrom.bg || theme().bg) };
      for (const el of [...masterItems(s), ...s.elements.flatMap(e => flattenGroup(e))]) {
        if (el.type === 'connector') {
          // A line now; addTimingToPptx turns it into a PowerPoint connector joined to its shapes.
          const [[x1, y1], [x2, y2]] = connectorPoints(el, s);
          ps.addShape(pptx.ShapeType.line, { x: inch(Math.min(x1, x2)), y: inch(Math.min(y1, y2)), w: inch(Math.max(1, Math.abs(x2 - x1))), h: inch(Math.max(1, Math.abs(y2 - y1))), flipH: x2 < x1, flipV: y2 < y1, objectName: 'atlas-' + el.id,
            line: { color: hex(el.color || theme().text), width: (el.width || 2) * 0.75, endArrowType: el.arrow !== 'none' ? 'triangle' : undefined, beginArrowType: el.arrow === 'both' ? 'triangle' : undefined } });
          continue;
        }
        const box = { x: inch(el.x), y: inch(el.y), w: inch(el.w), h: inch(el.h), rotate: el.rot || 0, objectName: (el.master ? 'atlas-master-' : 'atlas-') + el.id, ...(el.flipH ? { flipH: true } : {}), ...(el.flipV ? { flipV: true } : {}) };
        if (el.type === 'image') { ps.addImage({ data: el.src, ...box }); continue; }
        if (el.type === 'media') { ps.addMedia({ type: el.kind, data: el.src, cover: el.poster ? await asPng(el.poster) : (el.kind === 'audio' ? AUDIO_ICON() : undefined), ...box }); continue; }
        if (el.type === 'table') {
          const tmp = document.createElement('div'); tmp.innerHTML = el.html;
          const rows = [...tmp.querySelectorAll('tr')].map((tr, ri) => [...tr.children].map(td => ({
            text: td.innerText, options: ri === 0 ? { bold: true, color: 'FFFFFF', fill: { color: hex(theme().accent) } } : { fill: { color: 'DAE3F3' }, color: '262626' }
          })));
          if (rows.length) ps.addTable(rows, { ...box, fontSize: el.size, fontFace: fontOf(el), border: { type: 'solid', color: 'FFFFFF', pt: 1 } });
          continue;
        }
        const textOpts = {
          ...box, fontSize: el.size, fontFace: fontOf(el), color: hex(textColor(el)), align: el.align || 'left',
          valign: el.valign || 'top', margin: [6, 10, 6, 10].map(p => p * 0.75), lineSpacingMultiple: 1.0
        };
        if (el.ph === 'title' && el.bold !== false) { /* titles inherit bold from runs */ }
        if (el.type === 'shape') {
          const type = { roundRect: 'roundRect', rect: 'rect', ellipse: 'ellipse', triangle: 'triangle', rightArrow: 'rightArrow', line: 'line', star5: 'star5', diamond: 'diamond' }[el.shape] || 'rect';
          const shapeOpts = { ...textOpts, shape: pptx.ShapeType[type] || type };
          if (el.fill && el.shape !== 'line') shapeOpts.fill = { color: hex(el.fill) };
          if (el.line || el.shape === 'line') shapeOpts.line = { color: hex(el.line || theme().accent), width: el.shape === 'line' ? 2.25 : 1 };
          if (el.shape === 'line') { shapeOpts.flipV = !el.flipV; }
          ps.addText(el.html ? htmlToRuns(el.html) : '', shapeOpts);
          continue;
        }
        if (!el.html) continue; // empty placeholders aren't exported
        ps.addText(htmlToRuns(el.html), textOpts);
      }
      if (s.notes) ps.addNotes(s.notes);
    }
    const buf = await pptx.write({ outputType: 'arraybuffer', compression: true });
    return addTimingToPptx(new Uint8Array(buf));
  }
  /** PptxGenJS doesn't write transitions or animations: add them to each slide's XML. */
  async function addTimingToPptx(bytes) {
    const zip = await JSZip.loadAsync(bytes);
    // The master, layouts and theme go in a custom document property so Atlas can restore them exactly.
    const meta = JSON.stringify({ theme: pres.theme, customTheme: pres.customTheme || null, master: pres.master, layouts: pres.layouts, slides: pres.slides.map(sl => ({
      layout: sl.layout, hideMaster: !!sl.hideMaster,
      media: Object.fromEntries(sl.elements.flatMap(e => flattenGroup(e)).filter(e => e.type === 'media').map(e => [e.id, { autoplay: !!e.autoplay, loop: !!e.loop, hideIcon: !!e.hideIcon }])),
      groups: sl.elements.filter(e => e.type === 'group').map(groupMeta),
      connectors: Object.fromEntries(sl.elements.filter(e => e.type === 'connector').map(c => [c.id, { kind: c.kind, arrow: c.arrow, color: c.color || null, width: c.width, from: c.from || null, to: c.to || null }]))
    })) });
    const xmlEsc = t => t.replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
    zip.file('docProps/custom.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="AtlasSlidesMeta"><vt:lpwstr>${xmlEsc(meta)}</vt:lpwstr></property></Properties>`);
    const ct = await zip.file('[Content_Types].xml').async('string');
    if (!ct.includes('/docProps/custom.xml')) zip.file('[Content_Types].xml', ct.replace('</Types>', '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/></Types>'));
    const rels = await zip.file('_rels/.rels').async('string');
    if (!rels.includes('docProps/custom.xml')) zip.file('_rels/.rels', rels.replace('</Relationships>', '<Relationship Id="rIdAtlasCustom" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/></Relationships>'));
    for (let i = 0; i < pres.slides.length; i++) {
      const path = `ppt/slides/slide${i + 1}.xml`;
      const f = zip.file(path);
      if (!f) continue;
      let xml = await f.async('string');
      const ids = {};
      for (const m of xml.matchAll(/<p:cNvPr id="(\d+)" name="atlas-([^"]+)"/g)) ids[m[2]] = m[1];
      const s = pres.slides[i];
      // PptxGenJS writes audio as <a:videoFile>; PowerPoint wants <a:audioFile> for its audio relationships.
      if (s.elements.some(e => e.type === 'media' && e.kind === 'audio')) {
        const rels = await zip.file(`ppt/slides/_rels/slide${i + 1}.xml.rels`).async('string');
        const audioIds = new Set([...rels.matchAll(/<Relationship Id="([^"]+)" Type="[^"]*\/relationships\/audio"/g)].map(m => m[1]));
        xml = xml.replace(/<a:videoFile r:link="([^"]+)"\/>/g, (all, id) => (audioIds.has(id) ? `<a:audioFile r:link="${id}"/>` : all));
      }
      // Connectors: PowerPoint connectors joined to the shapes' connection sites.
      for (const c of s.elements.filter(e => e.type === 'connector')) {
        const site = (id, side) => {
          const t = s.elements.flatMap(e => flattenGroup(e)).find(e => e.id === id);
          const k = t && t.type === 'shape' && t.shape === 'ellipse' ? { n: 0, w: 2, s: 4, e: 6 } : { n: 0, w: 1, s: 2, e: 3 };
          return k[side];
        };
        const sides = c.from && c.to ? facingSides(s.elements.find(e => e.id === c.from.id) || {}, s.elements.find(e => e.id === c.to.id) || {}) : [];
        // A connector joined to a group links to the group's child nearest that end.
        const [[px1, py1], [px2, py2]] = connectorPoints(c, s);
        const target = (end, x, y) => {
          const t = end && s.elements.find(e => e.id === end.id);
          if (!t || t.type !== 'group') return end;
          const kids = flattenGroup(t);
          const best = kids.reduce((a, k) => (Math.hypot(k.x + k.w / 2 - x, k.y + k.h / 2 - y) < Math.hypot(a.x + a.w / 2 - x, a.y + a.h / 2 - y) ? k : a), kids[0]);
          return { id: best.id, site: end.site };
        };
        const link = (tag, end0, side, x, y) => { const end = target(end0, x, y); return end && ids[end.id] ? `<a:${tag} id="${ids[end.id]}" idx="${site(end.id, end.site === 'auto' ? side : end.site)}"/>` : ''; };
        xml = xml.replace(new RegExp(`<p:sp>((?:(?!<p:sp>)[\\s\\S])*?<p:cNvPr id="(\\d+)" name="atlas-${c.id}"[\\s\\S]*?)</p:sp>`), (all, body) => {
          const spPr = (body.match(/<p:spPr>[\s\S]*?<\/p:spPr>/) || [''])[0].replace(/prst="line"/, `prst="${c.kind === 'elbow' ? 'bentConnector3' : 'straightConnector1'}"`);
          const nv = (body.match(/<p:cNvPr [^>]*\/>|<p:cNvPr [^>]*>[\s\S]*?<\/p:cNvPr>/) || [''])[0];
          return `<p:cxnSp><p:nvCxnSpPr>${nv}<p:cNvCxnSpPr>${link('stCxn', c.from, sides[0], px1, py1)}${link('endCxn', c.to, sides[1], px2, py2)}</p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr>${spPr}</p:cxnSp>`;
        });
      }
      const extra = SlidesAnim.transitionXml(s) + SlidesAnim.timingXml(s, el => ids[el.id]);
      if (!extra) { zip.file(path, xml); continue; }
      xml = xml.includes('</p:clrMapOvr>') ? xml.replace('</p:clrMapOvr>', '</p:clrMapOvr>' + extra) : xml.replace('</p:cSld>', '</p:cSld>' + extra);
      zip.file(path, xml);
    }
    return new Uint8Array(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
  }

  /** A group's children with slide coordinates (turned with the group), for formats without groups. */
  function flattenGroup(el, parent) {
    if (el.type !== 'group') return [parent ? el : el];
    const cx = el.x + el.w / 2, cy = el.y + el.h / 2, rad = (el.rot || 0) * Math.PI / 180;
    return el.children.flatMap(c => {
      let nx = el.x + c.x, ny = el.y + c.y;
      if (el.rot) { const ox = nx + c.w / 2 - cx, oy = ny + c.h / 2 - cy; nx = cx + ox * Math.cos(rad) - oy * Math.sin(rad) - c.w / 2; ny = cy + ox * Math.sin(rad) + oy * Math.cos(rad) - c.h / 2; }
      const abs = { ...c, x: Math.round(nx), y: Math.round(ny), ...(el.rot ? { rot: ((((c.rot || 0) + el.rot) + 540) % 360) - 180 } : {}), ...(el.anim && !c.anim ? {} : {}) };
      return flattenGroup(abs, el);
    });
  }
  /** Groups as {id, x, y, w, h, rot, kids: [id | group]} for the Atlas metadata in .pptx. */
  const groupMeta = g => ({ id: g.id, x: g.x, y: g.y, w: g.w, h: g.h, rot: g.rot || 0, anim: g.anim || null, kids: g.children.map(c => (c.type === 'group' ? groupMeta(c) : c.id)) });
  /** Rebuilds groups from that metadata out of the flattened elements Atlas wrote (matched by shape name). */
  function regroup(sl, metas) {
    const byName = new Map(sl.elements.filter(e => e.atlasName).map(e => [e.atlasName, e]));
    const build = m => {
      const kids = m.kids.map(k => (typeof k === 'string' ? byName.get(k) : build(k))).filter(Boolean);
      if (!kids.length) return null;
      const g = { id: newId(), type: 'group', x: m.x, y: m.y, w: m.w, h: m.h, atlasName: m.id, ...(m.rot ? { rot: m.rot } : {}), ...(m.anim ? { anim: m.anim } : {}) };
      const cx = g.x + g.w / 2, cy = g.y + g.h / 2, rad = -(g.rot || 0) * Math.PI / 180;
      g.children = kids.map(c => {
        let nx = c.x, ny = c.y;
        if (g.rot) { const ox = nx + c.w / 2 - cx, oy = ny + c.h / 2 - cy; nx = cx + ox * Math.cos(rad) - oy * Math.sin(rad) - c.w / 2; ny = cy + ox * Math.sin(rad) + oy * Math.cos(rad) - c.h / 2; }
        const out = { ...c, x: Math.round(nx - g.x), y: Math.round(ny - g.y) };
        if (g.rot) { out.rot = ((((c.rot || 0) - g.rot) + 540) % 360) - 180; if (!out.rot) delete out.rot; }
        return out;
      });
      return g;
    };
    (Array.isArray(metas) ? metas : []).forEach(m => {
      if (!m || !Array.isArray(m.kids)) return;
      const g = build(m);
      if (!g) return;
      const members = new Set();
      const collect = mm => mm.kids.forEach(k => (typeof k === 'string' ? members.add(k) : collect(k)));
      collect(m);
      const at = sl.elements.findIndex(e => members.has(e.atlasName));
      sl.elements = sl.elements.filter(e => !members.has(e.atlasName));
      sl.elements.splice(Math.max(0, at), 0, g);
    });
  }
  /** PptxGenJS stores a media cover as .png whatever it is, so convert other pictures first. */
  function asPng(url) {
    if (/^data:image\/png/i.test(url)) return Promise.resolve(url);
    return new Promise(resolve => {
      const img = new Image();
      img.onload = () => { const c = document.createElement('canvas'); c.width = img.naturalWidth; c.height = img.naturalHeight; c.getContext('2d').drawImage(img, 0, 0); resolve(c.toDataURL('image/png')); };
      img.onerror = () => resolve(AUDIO_ICON());
      img.src = url;
    });
  }
  /** A small speaker picture for audio clips in PowerPoint. */
  const AUDIO_ICON = () => {
    const c = document.createElement('canvas'); c.width = c.height = 96;
    const g = c.getContext('2d');
    g.fillStyle = '#ffffff'; g.fillRect(0, 0, 96, 96);
    g.fillStyle = '#404040'; g.beginPath(); g.moveTo(16, 36); g.lineTo(34, 36); g.lineTo(56, 16); g.lineTo(56, 80); g.lineTo(34, 60); g.lineTo(16, 60); g.closePath(); g.fill();
    g.strokeStyle = '#404040'; g.lineWidth = 5; g.beginPath(); g.arc(58, 48, 16, -0.9, 0.9); g.stroke(); g.beginPath(); g.arc(58, 48, 28, -0.9, 0.9); g.stroke();
    return c.toDataURL('image/png');
  };
  // ---------- PPTX import ----------
  const EMU = 9525; // EMU per CSS px
  async function fromPptx(bytes) {
    const zip = await JSZip.loadAsync(bytes);
    const read = p => zip.file(p) ? zip.file(p).async('string') : Promise.resolve(null);
    const parse = s => new DOMParser().parseFromString(s, 'application/xml');
    const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
    const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const presXml = parse(await read('ppt/presentation.xml'));
    const presRels = parse(await read('ppt/_rels/presentation.xml.rels'));
    const relMap = rels => Object.fromEntries([...rels.getElementsByTagName('Relationship')].map(r => [r.getAttribute('Id'), r.getAttribute('Target')]));
    const pr = relMap(presRels);
    const sz = presXml.getElementsByTagNameNS(P, 'sldSz')[0];
    const sx = sz ? W / (+sz.getAttribute('cx') / EMU) : 1;
    const sy = sz ? H / (+sz.getAttribute('cy') / EMU) : 1;
    // Theme colours for schemeClr references.
    const themeXml = parse((await read('ppt/theme/theme1.xml')) || '<a/>');
    const scheme = {};
    const cs = themeXml.getElementsByTagNameNS(A, 'clrScheme')[0];
    // System colours (windowText, window) carry their RGB in lastClr.
    if (cs) [...cs.children].forEach(c => { const v = c.firstElementChild; if (v) scheme[c.localName] = v.localName === 'sysClr' ? v.getAttribute('lastClr') : v.getAttribute('val'); });
    scheme.tx1 = scheme.tx1 || scheme.dk1; scheme.bg1 = scheme.bg1 || scheme.lt1; scheme.tx2 = scheme.tx2 || scheme.dk2; scheme.bg2 = scheme.bg2 || scheme.lt2;
    const colorOf = node => {
      if (!node) return null;
      const srgb = node.getElementsByTagNameNS(A, 'srgbClr')[0];
      if (srgb) return color('#' + srgb.getAttribute('val'));
      const sc = node.getElementsByTagNameNS(A, 'schemeClr')[0];
      if (sc && Object.prototype.hasOwnProperty.call(scheme, sc.getAttribute('val'))) return color('#' + scheme[sc.getAttribute('val')]);
      return null;
    };
    const slides = [];
    const imported = {}; // master, layouts and theme read from PowerPoint's own parts
    let spids = new Map(); // PowerPoint shape id -> imported element, for animations
    const docs = new Map();
    const loadPart = async p => { if (!docs.has(p)) docs.set(p, { xml: parse((await read(p)) || '<x/>'), rels: relMap(parse((await read(p.replace(/([^/]+)$/, '_rels/$1.rels'))) || '<Relationships/>')) }); return docs.get(p); };
    const resolve = (from, target) => (target.startsWith('/') ? target.slice(1) : from.replace(/[^/]+$/, '') + target).split('/').reduce((a, x) => { if (x === '..') a.pop(); else if (x !== '.') a.push(x); return a; }, []).join('/');
    const partRelsTyped = async p => {
      const raw = parse((await read(p.replace(/([^/]+)$/, '_rels/$1.rels'))) || '<Relationships/>');
      return [...raw.getElementsByTagName('Relationship')].map(r => ({ id: r.getAttribute('Id'), type: (r.getAttribute('Type') || '').split('/').pop(), target: resolve(p, r.getAttribute('Target') || '') }));
    };
    // Theme fonts and colours become the presentation's theme.
    const fontScheme = themeXml.getElementsByTagNameNS(A, 'fontScheme')[0];
    const typeface = tag => { const f = fontScheme && fontScheme.getElementsByTagNameNS(A, tag)[0]; const l = f && f.getElementsByTagNameNS(A, 'latin')[0]; const t = l && l.getAttribute('typeface'); return t && /^[\w .-]{1,64}$/.test(t) ? t : null; };
    if (scheme.dk1 && scheme.lt1) {
      imported.theme = 'custom';
      imported.customTheme = { bg: '#' + (scheme.bg1 || scheme.lt1), title: '#' + (scheme.tx1 || scheme.dk1), text: '#' + (scheme.tx1 || scheme.dk1), accent: '#' + (scheme.accent1 || '4472C4'), font: typeface('minorFont') || 'Calibri', titleFont: typeface('majorFont') || 'Calibri Light' };
    }
    const phOf = node => { const ph = node.getElementsByTagNameNS(P, 'ph')[0]; return ph ? { type: ph.getAttribute('type') || 'body', idx: ph.getAttribute('idx') } : null; };
    const shapesOf = doc => { const t = doc.getElementsByTagNameNS(P, 'spTree')[0]; return t ? [...t.children] : []; };
    // PowerPoint matches a placeholder to its layout by idx, else by type (title/ctrTitle and body/obj alike).
    const sameKind = (a, b) => a === b || (/title/i.test(a) && /title/i.test(b)) || (['body', 'obj', 'subTitle'].includes(a) && ['body', 'obj'].includes(b));
    function findPh(doc, type, idx) {
      const list = shapesOf(doc).map(n => ({ n, ph: phOf(n) })).filter(x => x.ph);
      return (idx != null && list.find(x => x.ph.idx === idx && sameKind(type, x.ph.type))) || (idx != null && list.find(x => x.ph.idx === idx)) || list.find(x => x.ph.type === type) || list.find(x => sameKind(type, x.ph.type));
    }
    const xfrmBox = node => {
      const x = [...node.getElementsByTagNameNS(A, 'xfrm')].find(e => e.parentNode && /spPr|grpSpPr/.test(e.parentNode.localName)) || node.getElementsByTagNameNS(P, 'xfrm')[0] || node.getElementsByTagNameNS(A, 'xfrm')[0];
      const off = x && x.getElementsByTagNameNS(A, 'off')[0], ext = x && x.getElementsByTagNameNS(A, 'ext')[0];
      return off && ext ? { x: +off.getAttribute('x'), y: +off.getAttribute('y'), w: +ext.getAttribute('cx'), h: +ext.getAttribute('cy'), rot: (+x.getAttribute('rot') || 0) / 60000, flipH: x.getAttribute('flipH') === '1', flipV: x.getAttribute('flipV') === '1' } : null;
    };
    const lvl1Size = el => { const d = el && [...el.getElementsByTagNameNS(A, 'lvl1pPr')].map(l => l.getElementsByTagNameNS(A, 'defRPr')[0]).find(Boolean); return d && d.getAttribute('sz') ? +d.getAttribute('sz') / 100 : null; };
    /** Inheritance context for a slide (or the master/layout itself): where placeholders sit and their text size. */
    function inheritance(layoutDoc, masterDoc) {
      const txStyles = masterDoc && masterDoc.getElementsByTagNameNS(P, 'txStyles')[0];
      const styleSize = phType => {
        if (!txStyles) return null;
        const tag = phType !== 'subTitle' && /title/i.test(phType) ? 'titleStyle' : ['body', 'obj', 'subTitle'].includes(phType) ? 'bodyStyle' : 'otherStyle';
        return lvl1Size(txStyles.getElementsByTagNameNS(P, tag)[0]);
      };
      return {
        phBox(type, idx) {
          for (const d of [layoutDoc, masterDoc]) { if (!d) continue; const f = findPh(d, type, idx); const b = f && xfrmBox(f.n); if (b) return b; }
          return null;
        },
        inheritSize(type, idx) {
          if (!type) return null;
          for (const d of [layoutDoc, masterDoc]) { if (!d) continue; const f = findPh(d, type, idx); const sz = f && lvl1Size(f.n.getElementsByTagNameNS(A, 'lstStyle')[0]); if (sz) return sz; }
          return styleSize(type);
        }
      };
    }
    const bgOf = async (doc, docPath, docRels) => {
      const bg = doc.getElementsByTagNameNS(P, 'bg')[0];
      if (!bg) return null;
      const blip = bg.getElementsByTagNameNS(A, 'blip')[0];
      if (blip) { const src = await mediaUrl(zip, docPath, docRels[blip.getAttributeNS(R, 'embed')]); return src ? { bgImage: src } : null; }
      const c = colorOf(bg);
      return c ? { bg: c } : null;
    };
    const masters = new Map(), layoutsSeen = new Map();
    let masterParsed = null;
    for (const sid of ids_()) {
      const target = pr[sid.getAttributeNS(R, 'id') || sid.getAttribute('r:id')];
      if (!target) continue;
      const path = 'ppt/' + target.replace(/^\/?ppt\//, '').replace(/^\.\.\//, '');
      const xml = parse(await read(path));
      const relsPath = path.replace(/slides\/(slide\d+\.xml)$/, 'slides/_rels/$1.rels');
      const rels = relMap(parse((await read(relsPath)) || '<Relationships/>'));
      // The slide's layout and master.
      const layoutRel = (await partRelsTyped(path)).find(r => r.type === 'slideLayout');
      const layoutPath = layoutRel ? layoutRel.target : null;
      const layout = layoutPath ? await loadPart(layoutPath) : null;
      const masterRel = layoutPath ? (await partRelsTyped(layoutPath)).find(r => r.type === 'slideMaster') : null;
      const masterPath = masterRel ? masterRel.target : null;
      const master = masterPath ? await loadPart(masterPath) : null;
      if (master && !masters.has(masterPath)) masters.set(masterPath, master);
      const s = newSlide('blank');
      s.layout = 'blank';
      const ctx = { ...inheritance(layout && layout.xml, master && master.xml), tf: b => b };
      // Background: the slide's, else the layout's (the master's is the master background).
      Object.assign(s, (await bgOf(xml, path, rels)) || (layout && await bgOf(layout.xml, layoutPath, layout.rels)) || {});
      const hidden = attrIs(xml.documentElement, 'showMasterSp', '0') || (layout && attrIs(layout.xml.documentElement, 'showMasterSp', '0'));
      if (hidden) s.hideMaster = true;
      spids = new Map();
      // Shapes the layout draws (not placeholders) appear behind the slide's own.
      if (layout && !attrIs(xml.documentElement, 'showMasterSp', '0')) {
        for (const node of shapesOf(layout.xml)) { if (phOf(node)) continue; const el = await importShape(node, zip, layoutPath, layout.rels, { ...inheritance(null, master && master.xml), tf: b => b }); if (el) slide_push(s, el); }
      }
      for (const node of shapesOf(xml)) {
        const el = await importShape(node, zip, path, rels, ctx);
        if (el) slide_push(s, el);
      }
      // Connectors joined to shapes on this slide.
      s.elements.forEach(c => {
        if (c.type !== 'connector') return;
        const a = c.stId && spids.get(c.stId), b = c.enId && spids.get(c.enId);
        if (a && a.type !== 'connector') c.from = { id: a.id, site: 'auto' };
        if (b && b.type !== 'connector') c.to = { id: b.id, site: 'auto' };
        delete c.stId; delete c.enId;
      });
      // The layout as an Atlas layout, so new slides can use it.
      if (layoutPath) {
        if (!layoutsSeen.has(layoutPath)) {
          const key = 'L' + (layoutsSeen.size + 1);
          layoutsSeen.set(layoutPath, key);
          const lctx = { ...inheritance(layout.xml, master && master.xml), tf: b => b };
          const els = [];
          for (const node of shapesOf(layout.xml)) {
            const ph = phOf(node);
            if (!ph || !['title', 'ctrTitle', 'body', 'obj', 'subTitle'].includes(ph.type)) continue;
            const el = await importShape(node, zip, layoutPath, layout.rels, lctx, true);
            if (el) { el.html = ''; els.push(el); }
          }
          const cSld = layout.xml.getElementsByTagNameNS(P, 'cSld')[0];
          imported.layouts = imported.layouts || {};
          imported.layouts[key] = { name: (cSld && cSld.getAttribute('name')) || 'Layout ' + (layoutsSeen.size), custom: true, elements: els, ...((await bgOf(layout.xml, layoutPath, layout.rels)) || {}) };
        }
        s.layout = layoutsSeen.get(layoutPath);
      }
      // Transition and animations
      Object.assign(s, SlidesAnim.fromTransition(xml.getElementsByTagNameNS(P, 'transition')[0]) || {});
      SlidesAnim.fromTiming(xml.getElementsByTagNameNS(P, 'timing')[0]).forEach(({ spid, anim }) => { const el = spids.get(spid); if (el && !el.anim) el.anim = anim; });
      // Speaker notes
      const notesTarget = Object.values(rels).find(t => /notesSlide/.test(t));
      if (notesTarget) {
        const nx = parse((await read('ppt/' + notesTarget.replace(/^\.\.\//, ''))) || '<a/>');
        const bodies = [...nx.getElementsByTagNameNS(P, 'sp')].filter(sp => { const ph = sp.getElementsByTagNameNS(P, 'ph')[0]; return ph && ph.getAttribute('type') === 'body'; });
        s.notes = bodies.map(b => [...b.getElementsByTagNameNS(A, 'p')].map(p => [...p.getElementsByTagNameNS(A, 't')].map(t => t.textContent).join('')).join('\n')).join('\n').trim();
      }
      slides.push(s);
      if (!masterParsed && master) masterParsed = { master, masterPath };
    }
    // The (first) master: its background, its shapes and where its footers go.
    if (masterParsed) {
      const { master, masterPath } = masterParsed;
      const mctx = { ...inheritance(null, master.xml), tf: b => b };
      const els = [];
      for (const node of shapesOf(master.xml)) {
        const ph = phOf(node);
        if (ph && !['dt', 'ftr', 'sldNum'].includes(ph.type)) continue;
        const el = await importShape(node, zip, masterPath, master.rels, mctx, true);
        if (!el) continue;
        if (ph) { el.ph = { dt: 'dt', ftr: 'footer', sldNum: 'sldNum' }[ph.type]; el.html = ''; el.type = 'text'; delete el.shape; delete el.fill; delete el.line; }
        els.push(el);
      }
      imported.master = { elements: els, ...((await bgOf(master.xml, masterPath, master.rels)) || {}), footer: { text: '', number: false, date: false, skipTitle: true } };
    }
    function attrIs(el, name, v) { return !!el && el.getAttribute(name) === v; }
    function ids_() { return [...presXml.getElementsByTagNameNS(P, 'sldId')]; }
    function slide_push(s, el) { if (Array.isArray(el)) s.elements.push(...el); else s.elements.push(el); }
    /**
     * One shape as Atlas elements. ctx.tf maps a box from the shape's coordinates (inside groups)
     * to the slide's; ctx gives inherited placeholder positions and sizes. keepEmpty keeps empty
     * placeholders (for layouts and the master).
     */
    async function importShape(node, zipF, path, rels, ctx, keepEmpty) {
      const ln = node.localName;
      if (ln === 'grpSp') {
        // Children use the group's child coordinate space (chOff/chExt), scaled into its box.
        const gx = [...node.children].find(c => c.localName === 'grpSpPr');
        const x = gx && gx.getElementsByTagNameNS(A, 'xfrm')[0];
        const g = (tag, a) => { const e = x && x.getElementsByTagNameNS(A, tag)[0]; return e ? +e.getAttribute(a) : 0; };
        const off = { x: g('off', 'x'), y: g('off', 'y') }, ext = { w: g('ext', 'cx'), h: g('ext', 'cy') }, chOff = { x: g('chOff', 'x'), y: g('chOff', 'y') }, chExt = { w: g('chExt', 'cx') || ext.w, h: g('chExt', 'cy') || ext.h };
        const kx = chExt.w ? ext.w / chExt.w : 1, ky = chExt.h ? ext.h / chExt.h : 1;
        const tf = b => ctx.tf({ x: off.x + (b.x - chOff.x) * kx, y: off.y + (b.y - chOff.y) * ky, w: b.w * kx, h: b.h * ky, rot: b.rot, flipH: b.flipH, flipV: b.flipV });
        const out = [];
        for (const child of [...node.children]) { const e = await importShape(child, zipF, path, rels, { ...ctx, tf }, keepEmpty); if (e) out.push(...[].concat(e)); }
        if (keepEmpty || out.length < 2) return out; // layouts and the master keep their shapes loose
        // A group: its box in slide pixels, children relative to it.
        const gb = ctx.tf({ x: off.x, y: off.y, w: ext.w, h: ext.h });
        const gpx = Math.round(gb.x / EMU * sx), gpy = Math.round(gb.y / EMU * sy);
        const grp = { id: newId(), type: 'group', x: gpx, y: gpy, w: Math.round(gb.w / EMU * sx), h: Math.round(gb.h / EMU * sy), children: out.filter(c => c.type !== 'connector').map(c => ({ ...c, x: c.x - gpx, y: c.y - gpy })) };
        const grot = x ? (+x.getAttribute('rot') || 0) / 60000 : 0;
        if (grot) grp.rot = grot;
        const gnv = node.getElementsByTagNameNS(P, 'cNvPr')[0];
        if (gnv && gnv.getAttribute('id')) spids.set(gnv.getAttribute('id'), grp);
        return [grp, ...out.filter(c => c.type === 'connector')];
      }
      if (!['sp', 'cxnSp', 'pic', 'graphicFrame'].includes(ln)) return null;
      const nv = node.getElementsByTagNameNS(P, 'cNvPr')[0];
      const spid = nv ? nv.getAttribute('id') : null;
      if (nv && /^atlas-master-/.test(nv.getAttribute('name') || '')) return null; // drawn from the master Atlas saved
      const nm = (nv && nv.getAttribute('name')) || '';
      const tag = el => { if (el && spid) spids.set(spid, el); if (el && /^atlas-/.test(nm)) el.atlasName = nm.slice(6); if (el && raw && raw.flipH && el.type !== 'connector') el.flipH = true; if (el && raw && raw.flipV && el.type !== 'connector') el.flipV = true; return el; };
      const phInfo = phOf(node);
      let phType = phInfo ? phInfo.type : null;
      const phIdx = phInfo ? phInfo.idx : null;
      // Footer placeholders on a slide are ordinary text there.
      if (['dt', 'ftr', 'sldNum', 'hdr'].includes(phType) && !keepEmpty) phType = null;
      let raw = xfrmBox(node) || (phInfo && ctx.phBox(phInfo.type, phIdx));
      // Connectors: their ends come from the box and its flips; joined shapes are linked after the slide is read.
      if (ln === 'cxnSp' && raw) {
        const r2 = ctx.tf(raw);
        const px = v => v / EMU;
        const X = r2.x / EMU * sx, Y = r2.y / EMU * sy, Wd = r2.w / EMU * sx, Ht = r2.h / EMU * sy;
        const spPr0 = node.getElementsByTagNameNS(P, 'spPr')[0];
        const prst0 = (spPr0 && spPr0.getElementsByTagNameNS(A, 'prstGeom')[0] || { getAttribute: () => '' }).getAttribute('prst') || '';
        const lnEl = spPr0 && spPr0.getElementsByTagNameNS(A, 'ln')[0];
        const endType = n => { const e = lnEl && lnEl.getElementsByTagNameNS(A, n)[0]; return e && e.getAttribute('type') && e.getAttribute('type') !== 'none'; };
        const c = {
          id: newId(), type: 'connector', kind: /bent|elbow/i.test(prst0) ? 'elbow' : 'straight', arrow: endType('tailEnd') ? (endType('headEnd') ? 'both' : 'end') : (endType('headEnd') ? 'end' : 'none'),
          x1: Math.round(X + (raw.flipH ? Wd : 0)), y1: Math.round(Y + (raw.flipV ? Ht : 0)), x2: Math.round(X + (raw.flipH ? 0 : Wd)), y2: Math.round(Y + (raw.flipV ? 0 : Ht)),
          width: lnEl && lnEl.getAttribute('w') ? Math.max(0.5, Math.round(px(+lnEl.getAttribute('w')) * 10) / 10) : 2, x: 0, y: 0, w: 1, h: 1
        };
        if (endType('headEnd') && !endType('tailEnd')) { [c.x1, c.x2] = [c.x2, c.x1]; [c.y1, c.y2] = [c.y2, c.y1]; }
        const col = lnEl && colorOf(lnEl);
        if (col) c.color = col;
        const st = node.getElementsByTagNameNS(A, 'stCxn')[0], en = node.getElementsByTagNameNS(A, 'endCxn')[0];
        if (st) c.stId = st.getAttribute('id');
        if (en) c.enId = en.getAttribute('id');
        return tag(c);
      }
      let box;
      if (raw) {
        raw = ctx.tf(raw);
        box = { x: Math.round(raw.x / EMU * sx), y: Math.round(raw.y / EMU * sy), w: Math.round(raw.w / EMU * sx), h: Math.round(raw.h / EMU * sy) };
        if (raw.rot) box.rot = raw.rot;
      } else if (phType) {
        box = phType === 'subTitle' ? { x: 120, y: 305, w: 720, h: 80 } : /title/i.test(phType) ? (phType === 'ctrTitle' ? { x: 80, y: 140, w: 800, h: 150 } : { x: 60, y: 28, w: 840, h: 90 }) : { x: 60, y: 130, w: 840, h: 370 };
      } else return null;
      if (ln === 'pic') {
        const blip = node.getElementsByTagNameNS(A, 'blip')[0];
        const src = blip && await mediaUrl(zipF, path, rels[blip.getAttributeNS(R, 'embed')]);
        // Video and audio: a picture (the poster) with a link to the media part.
        const mf = [...node.getElementsByTagName('*')].find(e => ['videoFile', 'audioFile', 'media'].includes(e.localName) && (e.getAttributeNS(R, 'link') || e.getAttributeNS(R, 'embed')));
        if (mf) {
          const target = rels[mf.getAttributeNS(R, 'embed') || mf.getAttributeNS(R, 'link')];
          const clip = target && await mediaFile(zipF, path, target);
          if (clip) {
            const el = tag({ id: newId(), type: 'media', kind: clip.startsWith('data:audio') ? 'audio' : 'video', src: clip, ...(src ? { poster: src } : {}), ...box });
            const nm = (nv && nv.getAttribute('name')) || '';
            if (/^atlas-/.test(nm)) el.atlasName = nm.slice(6);
            return el;
          }
        }
        return src ? tag({ id: newId(), type: 'image', src, ...box }) : null;
      }
      if (ln === 'graphicFrame') {
        const tbl = node.getElementsByTagNameNS(A, 'tbl')[0];
        if (!tbl) return null;
        const html = '<table>' + [...tbl.getElementsByTagNameNS(A, 'tr')].map(tr => '<tr>' + [...tr.getElementsByTagNameNS(A, 'tc')].map(tc => `<td>${escapeHtml([...tc.getElementsByTagNameNS(A, 't')].map(t => t.textContent).join(' '))}</td>`).join('') + '</tr>').join('') + '</table>';
        return tag({ id: newId(), type: 'table', html, size: 14, align: 'left', valign: 'top', ...box });
      }
      const spPr = node.getElementsByTagNameNS(P, 'spPr')[0];
      const geom = spPr && spPr.getElementsByTagNameNS(A, 'prstGeom')[0];
      const prst = geom ? geom.getAttribute('prst') : null;
      const fillNode = spPr && [...spPr.children].find(c => c.localName === 'solidFill');
      const lnNode = spPr && spPr.getElementsByTagNameNS(A, 'ln')[0];
      // Shapes drawn with PowerPoint's default style take their colours from <p:style> theme references.
      const styleEl = [...node.children].find(c => c.localName === 'style');
      const styleRef = name => { const r = styleEl && [...styleEl.children].find(c => c.localName === name); return r && +r.getAttribute('idx') > 0 ? colorOf(r) : null; };
      const noFill = spPr && [...spPr.children].some(c => c.localName === 'noFill');
      const fill = fillNode ? colorOf(fillNode) : noFill ? null : styleRef('fillRef');
      const line = lnNode && lnNode.getElementsByTagNameNS(A, 'noFill')[0] ? null : (lnNode && colorOf(lnNode)) || styleRef('lnRef');
      const txBody = node.getElementsByTagNameNS(P, 'txBody')[0];
      let html = '', size = ctx.inheritSize(phType, phIdx) || (phType && /title/i.test(phType) ? 40 : 18), align = 'left', valign = 'top';
      if (txBody) {
        const bodyPr = txBody.getElementsByTagNameNS(A, 'bodyPr')[0];
        const anchor = bodyPr && bodyPr.getAttribute('anchor');
        valign = anchor === 'ctr' ? 'middle' : anchor === 'b' ? 'bottom' : (phType && /title/i.test(phType) ? 'middle' : 'top');
        const parts = [];
        let listOpen = false;
        const sizes = [];
        for (const p of txBody.getElementsByTagNameNS(A, 'p')) {
          const pPr = p.getElementsByTagNameNS(A, 'pPr')[0];
          const algn = pPr && pPr.getAttribute('algn');
          if (algn) align = { ctr: 'center', r: 'right', just: 'left' }[algn] || 'left';
          const bullet = pPr && (pPr.getElementsByTagNameNS(A, 'buChar')[0] || pPr.getElementsByTagNameNS(A, 'buAutoNum')[0]);
          const isBullet = (bullet || (phType === 'body' && !(pPr && pPr.getElementsByTagNameNS(A, 'buNone')[0])));
          const runs = [...p.children].filter(c => c.localName === 'r' || c.localName === 'br' || c.localName === 'fld').map(r => {
            if (r.localName === 'br') return '<br>';
            const t = r.getElementsByTagNameNS(A, 't')[0];
            let text = escapeHtml(t ? t.textContent : '');
            const rPr = r.getElementsByTagNameNS(A, 'rPr')[0];
            if (rPr) {
              const style = [];
              if (rPr.getAttribute('sz')) { const pt = +rPr.getAttribute('sz') / 100; sizes.push(pt); style.push(`font-size:${pt}pt`); }
              const c = colorOf([...rPr.children].find(ch => ch.localName === 'solidFill'));
              if (c) style.push(`color:${c}`);
              const latin = rPr.getElementsByTagNameNS(A, 'latin')[0];
              const face = latin && latin.getAttribute('typeface');
              if (face && !face.startsWith('+') && /^[\w .-]{1,64}$/.test(face)) style.push(`font-family:'${face}'`);
              if (rPr.getAttribute('b') === '1') text = `<b>${text}</b>`;
              if (rPr.getAttribute('i') === '1') text = `<i>${text}</i>`;
              if (rPr.getAttribute('u') && rPr.getAttribute('u') !== 'none') text = `<u>${text}</u>`;
              if (style.length) text = `<span style="${style.join(';')}">${text}</span>`;
            }
            return text;
          }).join('');
          if (isBullet && runs) { if (!listOpen) { parts.push('<ul>'); listOpen = true; } parts.push(`<li>${runs}</li>`); }
          else { if (listOpen) { parts.push('</ul>'); listOpen = false; } parts.push(`<div>${runs || '<br>'}</div>`); }
        }
        if (listOpen) parts.push('</ul>');
        html = parts.join('');
        if (!txBody.textContent.trim()) html = '';
        if (sizes.length) {
          // Use the most common run size as the box size and drop it from the runs.
          const common = sizes.sort((a, b) => sizes.filter(v => v === a).length - sizes.filter(v => v === b).length).pop();
          size = common;
          html = html.replace(new RegExp(`font-size:${common}pt;?`, 'g'), '').replace(/ style=""/g, '');
        }
      }
      if (!html && !fill && !line && !prst && !(keepEmpty && phType)) return null;
      const isShape = (fill || line || (prst && prst !== 'rect')) && !phType;
      const el = { id: newId(), type: isShape ? 'shape' : 'text', html, size, align, valign, ...box };
      if (phType) el.ph = phType === 'subTitle' ? 'subtitle' : /title/i.test(phType) ? 'title' : 'body';
      tag(el);
      if (isShape) { el.shape = { rect: 'rect', roundRect: 'roundRect', ellipse: 'ellipse', triangle: 'triangle', rightArrow: 'rightArrow', line: 'line', star5: 'star5', diamond: 'diamond', straightConnector1: 'line' }[prst] || 'rect'; el.fill = fill; el.line = line; }
      return el;
    }
    async function mediaFile(zipF, slidePath, target) {
      if (!target || /^https?:/i.test(target)) return null;
      const p = target.startsWith('/') ? target.slice(1) : slidePath.replace(/[^/]+$/, '') + target;
      const norm = p.split('/').reduce((acc, part) => { if (part === '..') acc.pop(); else if (part !== '.') acc.push(part); return acc; }, []).join('/');
      const f = zipF.file(norm);
      const mime = MEDIA_TYPES[norm.split('.').pop().toLowerCase()];
      if (!f || !mime) return null;
      return 'data:' + mime + ';base64,' + await f.async('base64');
    }
    async function mediaUrl(zipF, slidePath, target) {
      if (!target) return null;
      const p = target.startsWith('/') ? target.slice(1) : slidePath.replace(/[^/]+$/, '') + target;
      const norm = p.split('/').reduce((acc, part) => { if (part === '..') acc.pop(); else if (part !== '.') acc.push(part); return acc; }, []).join('/');
      const f = zipF.file(norm);
      if (!f) return null;
      const ext = norm.split('.').pop().toLowerCase();
      const mime = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', svg: 'image/svg+xml', bmp: 'image/bmp', webp: 'image/webp' }[ext];
      if (!mime) return null;
      return 'data:' + mime + ';base64,' + await f.async('base64');
    }
    if (!slides.length) slides.push(newSlide('title'));
    let meta = null;
    const custom = await read('docProps/custom.xml');
    if (custom) {
      const prop = [...parse(custom).getElementsByTagName('*')].find(e => e.localName === 'property' && e.getAttribute('name') === 'AtlasSlidesMeta');
      try { meta = prop ? JSON.parse(prop.textContent) : null; } catch { meta = null; }
    }
    if (meta && typeof meta === 'object') {
      (Array.isArray(meta.slides) ? meta.slides : []).forEach((m, i) => {
        if (!slides[i] || !m) return;
        slides[i].layout = m.layout; slides[i].hideMaster = m.hideMaster === true;
        // Playback options, matched by the shape names Atlas wrote.
        if (m.media && typeof m.media === 'object') slides[i].elements.forEach(e => { const o = e.type === 'media' && e.atlasName && Object.hasOwn(m.media, e.atlasName) ? m.media[e.atlasName] : null; if (o) ['autoplay', 'loop', 'hideIcon'].forEach(k => { e[k] = o[k] === true; }); });
        if (m.connectors && typeof m.connectors === 'object') slides[i].elements.forEach(e => {
          const o = e.type === 'connector' && e.atlasName && Object.hasOwn(m.connectors, e.atlasName) ? m.connectors[e.atlasName] : null;
          if (!o) return;
          Object.assign(e, { kind: o.kind, arrow: o.arrow, width: o.width, ...(o.color ? { color: o.color } : {}) });
          ['from', 'to'].forEach(k => { if (o[k]) e[k] = { id: o[k].id, site: o[k].site }; });
        });
        regroup(slides[i], m.groups);
        // Connector ends named by Atlas ids: point them at the reopened elements.
        const byName = new Map(slides[i].elements.filter(e => e.atlasName).map(e => [e.atlasName, e.id]));
        slides[i].elements.forEach(e => { if (e.type === 'connector') ['from', 'to'].forEach(k => { if (e[k] && byName.has(e[k].id)) e[k] = { ...e[k], id: byName.get(e[k].id) }; else if (e[k] && !slides[i].elements.some(x => x.id === e[k].id)) delete e[k]; }); });
      });
      return cleanPres({ theme: meta.theme, customTheme: meta.customTheme, master: meta.master, layouts: meta.layouts, slides });
    }
    return cleanPres({ theme: 'office', slides, ...imported });
  }

  // ---------- Files ----------
  async function loadFile(f) {
    try {
      const ext = Platform.extOf(f.name);
      if (ext === 'aslides') {
        pres = cleanPres(JSON.parse(Platform.bytesToText(f.data)));
      } else pres = await fromPptx(f.data);
    } catch (err) {
      UI.toast('Could not open file: ' + err.message, 4000);
      return;
    }
    current = 0; selectedId = editingId = null;
    undoStack.length = redoStack.length = undoMarks.length = redoMarks.length = 0;
    state.set({ name: f.name, path: f.path, dirty: false });
    Platform.rememberRecent(f.name, 'slides');
    renderAll();
  }
  async function serialize(ext) {
    if (ext === 'aslides') return JSON.stringify(pres);
    return toPptx();
  }
  async function save(saveAs, quiet) {
    commitEditing();
    let name = state.name;
    if (!/\.(pptx|aslides)$/i.test(name)) name = Platform.baseName(name) + '.pptx';
    const ext = Platform.extOf(name);
    const filters = [...SAVE_FILTERS].sort((a, b) => (b.extensions.includes(ext) ? 1 : 0) - (a.extensions.includes(ext) ? 1 : 0));
    try {
      if (saveAs || !state.path) {
        const target = await Platform.chooseSavePath(name, filters);
        if (!target) return false;
        await Platform.save({ data: await serialize(Platform.extOf(target.name) || 'pptx'), name: target.name, path: target.path });
        state.set({ name: target.name, path: target.path, dirty: false });
      } else {
        await Platform.save({ data: await serialize(ext), name, path: state.path });
        state.set({ dirty: false });
      }
      if (!quiet) UI.toast('Saved');
      return true;
    } catch (err) {
      UI.toast('Save failed: ' + err.message, 4000);
      return false;
    }
  }
  function preparePrint() {
    const box = document.getElementById('printSlides');
    box.innerHTML = '';
    pres.slides.forEach(s => { const d = document.createElement('div'); d.className = 'slide'; renderSlide(s, d, false); box.appendChild(d); });
  }
  window.addEventListener('beforeprint', preparePrint);

  const CLOUD_EXTS = ['pptx', 'aslides'];
  Platform.onMenu(async cmd => {
    switch (cmd) {
      case 'new': Platform.newWindow('slides'); break;
      case 'open': {
        const f = await Platform.open(OPEN_FILTERS);
        if (!f) break;
        const empty = !state.dirty && pres.slides.length === 1 && !pres.slides[0].elements.some(e => e.html || e.type !== 'text');
        if (Platform.shouldOpenInPlace(empty)) {
          if (state.dirty && !(await UI.confirm('Open file', 'Discard unsaved changes in this presentation?'))) break;
          loadFile(f);
        } else Platform.openPath(f.path);
        break;
      }
      case 'save': save(false); break;
      case 'save-as': save(true); break;
      case 'open-cloud': Cloud.openDialog({ app: 'slides', load: loadFile, isEmpty: async () => Platform.shouldOpenInPlace(!state.dirty && pres.slides.length === 1 && !pres.slides[0].elements.some(e => e.html || e.type !== 'text')) && (!state.dirty || UI.confirm('Open file', 'Discard unsaved changes in this presentation?')) }); break;
      case 'save-cloud': Cloud.saveAsDialog({ state, serialize: ext => serialize(ext), exts: CLOUD_EXTS, before: () => commitEditing() }); break;
      case 'share': Cloud.shareDialog({ state, serialize: ext => serialize(ext), exts: CLOUD_EXTS, before: () => commitEditing(), save: () => save(false) }); break;
      case 'save-and-close': if (await save(false)) Platform.closeNow(); break;
      case 'export-pdf': commitEditing(); preparePrint(); Platform.exportPdf(Platform.baseName(state.name) + '.pdf', true); break;
      case 'print': commitEditing(); preparePrint(); Platform.print(); break;
      case 'undo': if (editingId) document.execCommand('undo'); else undo(); break;
      case 'redo': if (editingId) document.execCommand('redo'); else redo(); break;
      case 'present': present(0); break;
      case 'presenter-view': presenterView(current); break;
      case 'present-current': present(current); break;
      case 'find': UI.toast('Find is available in Write and Sheets'); break;
      case 'home': Platform.desktop ? Platform.newWindow('home') : (location.href = 'index.html'); break;
    }
  });

  (async () => {
    if (AUDIENCE) {
      document.body.classList.add('audience');
      presenter.hidden = false;
      audience.toPresenter({ cmd: 'ready' });
      return;
    }
    pres.master = defaultMaster();
    pres.slides = [newSlide('title')];
    renderAll();
    const f = await Platform.pending();
    if (f && f.error) UI.toast('Could not open file: ' + f.error, 4000);
    else if (f && f.recovered) loadRecovered(f.recovered);
    else if (f) await loadFile(f);
  })();

  // ---------- Crash recovery & AutoSave ----------
  const recovery = AUDIENCE ? { adopt() {} } : Platform.enableRecovery({
    app: 'slides', state,
    snapshot: () => { commitEditingQuietly(); return JSON.stringify(pres); },
    autosave: () => save(false, true),
    canAutoSave: () => !Platform.desktop && !Platform.isCloudPath(state.path) ? 'AutoSave needs the desktop app' : !state.path ? 'Save the presentation first to turn on AutoSave'
      : /\.(pptx|aslides)$/i.test(state.path) ? true : 'AutoSave works with .pptx and .aslides files'
  });
  /** Copies the text being typed into the model without leaving edit mode. */
  function commitEditingQuietly() {
    if (!editingId) return;
    const el = elById(editingId);
    const c = stage.querySelector(`[data-id="${editingId}"] .content`);
    if (el && c) el.html = c.innerHTML;
  }
  function loadRecovered(d) {
    try { pres = cleanPres(JSON.parse(d.data)); } catch { return UI.toast('The recovered presentation could not be read', 4000); }
    current = 0; selectedId = editingId = null;
    undoStack.length = redoStack.length = undoMarks.length = redoMarks.length = 0;
    renderAll();
    state.set({ name: d.name, path: d.path, dirty: true });
    recovery.adopt(d.draftId);
    UI.toast('Recovered your unsaved changes', 4000);
  }

  // ---------- Live editing ----------
  // Keys: deck -> theme, master and layouts; s:<id> -> a slide's settings and position;
  // e:<slide id>:<element id> -> a shape, picture, table… with its stacking order.
  let pointerDown = false;
  stage.addEventListener('mousedown', () => { pointerDown = true; }, true);
  document.addEventListener('mouseup', () => setTimeout(() => { pointerDown = false; }), true);
  function drawLivePeers() {
    stage.querySelectorAll('.live-flag').forEach(f => f.remove());
    thumbs.querySelectorAll('.live-dot').forEach(f => f.remove());
    if (!livePeers.length) return;
    const ids = pres.slides.map(s => s.id);
    for (const p of livePeers) {
      const pr = p.presence;
      if (!pr) continue;
      const i = ids.indexOf(pr.slide);
      const t = !masterTarget && i >= 0 && thumbs.children[i];
      if (t) {
        const d = document.createElement('span');
        d.className = 'live-dot';
        d.title = p.name;
        d.style.background = p.color;
        t.querySelector('.num').appendChild(d);
      }
      if (masterTarget || i !== current || !pr.el) continue;
      const el = slide().elements.find(e => e.id === pr.el);
      if (!el) continue;
      const f = document.createElement('div');
      f.className = 'live-flag';
      Object.assign(f.style, { left: (el.x - 3) + 'px', top: (el.y - 3) + 'px', width: (el.w + 2) + 'px', height: (el.h + 2) + 'px', border: `2px solid ${p.color}`, transform: el.rot ? `rotate(${el.rot}deg)` : '' });
      const tag = document.createElement('span');
      tag.className = 'tag';
      tag.style.background = p.color;
      tag.textContent = p.name;
      f.appendChild(tag);
      stage.appendChild(f);
    }
  }
  live = Collab.attach({
    app: 'slides', doc: state,
    dirty: () => state.dirty || pres.slides.length > 1 || pres.slides.some(s => s.elements.some(e => e.html)),
    busy: () => pointerDown,
    state() {
      commitEditingQuietly();
      const out = {};
      const deck = {};
      for (const [k, v] of Object.entries(pres)) if (k !== 'slides') deck[k] = v;
      out.deck = deck;
      pres.slides.forEach((s, i) => {
        const props = { pos: i };
        for (const [k, v] of Object.entries(s)) if (k !== 'elements') props[k] = v;
        out['s:' + s.id] = props;
        s.elements.forEach((el, z) => { out[`e:${s.id}:${el.id}`] = { ...el, z }; });
      });
      return out;
    },
    apply(changes) {
      const curId = pres.slides[current] && pres.slides[current].id;
      // Keep the text being typed and the caret where they are.
      let caret = null;
      if (editingId) {
        commitEditingQuietly();
        const c = stage.querySelector(`[data-id="${editingId}"] .content`);
        const sel = getSelection();
        if (c && sel.rangeCount && c.contains(sel.anchorNode)) { const r = document.createRange(); r.setStart(c, 0); r.setEnd(sel.anchorNode, sel.anchorOffset); caret = r.toString().length; }
      }
      const byId = new Map(pres.slides.map((s, i) => { s.pos = i; s.elements.forEach((e, z) => { e.z = z; }); return [s.id, s]; }));
      const orphans = new Map();
      for (const [k, v] of Object.entries(changes)) {
        if (k === 'deck' && v) {
          for (const key of Object.keys(pres)) if (key !== 'slides') delete pres[key];
          Object.assign(pres, JSON.parse(JSON.stringify(v)));
        } else if (k.startsWith('s:')) {
          const id = k.slice(2);
          if (v === null) { byId.delete(id); continue; }
          let sl = byId.get(id);
          if (!sl) { sl = { id, elements: [] }; byId.set(id, sl); }
          for (const key of Object.keys(sl)) if (key !== 'elements') delete sl[key];
          Object.assign(sl, JSON.parse(JSON.stringify(v)), { id });
        } else if (k.startsWith('e:')) {
          const i = k.indexOf(':', 2);
          const sid = k.slice(2, i), eid = k.slice(i + 1);
          let sl = byId.get(sid);
          if (!sl) { if (v === null) continue; if (!orphans.has(sid)) orphans.set(sid, []); orphans.get(sid).push(v); continue; }
          const at = sl.elements.findIndex(e => e.id === eid);
          if (v === null) { if (at >= 0) sl.elements.splice(at, 1); continue; }
          const el = JSON.parse(JSON.stringify(v));
          if (at >= 0) sl.elements[at] = el; else sl.elements.push(el);
        }
      }
      for (const [sid, els] of orphans) { const sl = byId.get(sid); if (sl) sl.elements.push(...els); }
      const list = [...byId.values()].sort((a, b) => (a.pos - b.pos) || (a.id < b.id ? -1 : 1));
      if (list.length) pres.slides = list;
      pres.slides.forEach(s => { delete s.pos; s.elements.sort((a, b) => (a.z - b.z) || (a.id < b.id ? -1 : 1)); s.elements.forEach(e => { delete e.z; }); });
      const ci = pres.slides.findIndex(s => s.id === curId);
      current = ci >= 0 ? ci : Math.min(current, pres.slides.length - 1);
      if (selectedId && !slide().elements.some(e => e.id === selectedId)) selectedId = null;
      if (editingId && !slide().elements.some(e => e.id === editingId)) editingId = null;
      renderAll();
      if (editingId) {
        const c = stage.querySelector(`[data-id="${editingId}"] .content`);
        if (c) {
          c.focus();
          if (caret !== null) {
            const walk = document.createTreeWalker(c, NodeFilter.SHOW_TEXT);
            let t, n = caret, done = false;
            while ((t = walk.nextNode())) { if (n <= t.length) { getSelection().collapse(t, n); done = true; break; } n -= t.length; }
            if (!done) { const r = document.createRange(); r.selectNodeContents(c); r.collapse(false); getSelection().removeAllRanges(); getSelection().addRange(r); }
          }
        }
      }
      state.markDirty();
    },
    presence: () => ({ slide: pres.slides[current] && pres.slides[current].id, el: selectedId }),
    showPeers(list) { livePeers = list; drawLivePeers(); }
  });

  // ---------- Proofing ----------
  const proofing = Proofing.attach({
    app: 'slides',
    roots: () => {
      commitEditing();
      const out = [];
      pres.slides.forEach((s, si) => {
        s.elements.forEach(e => {
          if (!e.html || !(e.type === 'text' || e.type === 'shape')) return;
          const el = document.createElement('div');
          el.innerHTML = Platform.sanitizeHtml(e.html);
          out.push({
            el,
            commit: () => { snapshot(); e.html = el.innerHTML; renderAll(); state.markDirty(); },
            focus: () => { if (current !== si || masterTarget) goTo(si); select(e.id); }
          });
        });
        if (s.notes && /\p{L}{2}/u.test(s.notes)) {
          const el = document.createElement('div');
          el.textContent = s.notes;
          out.push({ el, commit: () => { snapshot(); s.notes = el.textContent; renderStage(); state.markDirty(); }, focus: () => { if (current !== si) goTo(si); } });
        }
      });
      return out;
    },
    lang: () => pres.lang || '',
    setLang: code => { pres.lang = code || undefined; state.markDirty(); }
  });
  document.getElementById('proofEditor').addEventListener('click', () => proofing.open());
  document.getElementById('proofLang').addEventListener('click', () => proofing.languageDialog());

  // ---------- Assistant ----------
  const plain = html => { const d = document.createElement('div'); d.innerHTML = html || ''; d.querySelectorAll('li, p, div, br').forEach(n => n.before('\n')); return d.textContent.replace(/\n{2,}/g, '\n').trim(); };
  function aiContext() {
    commitEditingQuietly();
    const out = pres.slides.map((s, i) => {
      const texts = s.elements.filter(e => e.type === 'text' || e.type === 'shape').map(e => plain(e.html)).filter(Boolean);
      return `Slide ${i + 1}${i === current ? ' (current)' : ''}:\n${texts.map(t => '  ' + t.replace(/\n/g, '\n  ')).join('\n')}${s.notes ? `\n  Notes: ${s.notes}` : ''}`;
    }).join('\n\n');
    const el = selectedId && elById(selectedId);
    const selection = el && el.html ? plain(el.html) : '';
    return { document: out, selection, hasSelection: !!selection, elId: el ? el.id : null, slideId: slide().id };
  }
  function toListHtml(lines) {
    return '<ul>' + lines.map(l => `<li>${escapeHtml(l.replace(/^\s*(?:[-*•]|\d+[.)])\s+/, '').replace(/\*\*([^*]+)\*\*/g, '$1'))}</li>`).join('') + '</ul>';
  }
  /** Turns an outline (## Title, then bullet lines; "Notes:" lines become speaker notes) into new slides. */
  function aiSlides(text) {
    const sections = [];
    for (const raw of text.split('\n')) {
      const line = raw.trim();
      if (!line || /^---+$/.test(line)) continue;
      const h = /^(?:#{1,3}\s+|\*\*)?(?:slide\s*\d+\s*[:.\-–]\s*)?(.*?)(?:\*\*)?$/i.exec(line);
      if (/^#{1,3}\s/.test(line) || /^(\*\*)?slide\s*\d+/i.test(line)) { sections.push({ title: h[1].replace(/^#+\s*/, ''), body: [], notes: [] }); continue; }
      if (!sections.length) sections.push({ title: line.replace(/^\s*[-*•]\s+/, ''), body: [], notes: [] });
      else if (/^notes?:/i.test(line)) sections[sections.length - 1].notes.push(line.replace(/^notes?:\s*/i, ''));
      else sections[sections.length - 1].body.push(line);
    }
    if (!sections.length) return UI.toast('No slides found in the reply', 3000);
    if (masterTarget) return UI.toast('Close Master View to add slides');
    commitEditing();
    snapshot();
    for (const sec of sections.slice(0, 50)) {
      const sl = newSlide(sec.body.length ? 'content' : 'title');
      const title = sl.elements.find(e => e.ph === 'title');
      const body = sl.elements.find(e => e.ph === 'body' || e.ph === 'subtitle');
      if (title) title.html = escapeHtml(sec.title);
      if (body && sec.body.length) body.html = body.ph === 'body' ? toListHtml(sec.body) : escapeHtml(sec.body.join(' '));
      sl.notes = sec.notes.join('\n');
      pres.slides.splice(++current, 0, sl);
    }
    selectedId = null;
    renderAll();
    state.markDirty();
    UI.toast(`Added ${Math.min(sections.length, 50)} slide${sections.length === 1 ? '' : 's'}`);
  }
  Assistant.attach({
    app: 'slides',
    context: aiContext,
    beforeAsk: () => commitEditingQuietly(),
    format: 'In Atlas Slides, write slides as an outline: a "## " heading line with each slide\'s title, then its points as "- " lines (at most six short points), and optionally a "Notes:" line with speaker notes.',
    actions: [
      { label: 'Draft slides', prompt: 'Draft 5 to 8 slides on the topic I describe next, as an outline. Topic: ', fill: true },
      { label: 'Speaker notes', prompt: 'Write speaker notes for the current slide: what to say, in 3 to 5 natural sentences. Reply with only the notes.' },
      { label: 'Improve this text', prompt: 'Rewrite the selected text box for a slide: shorter, punchier, parallel bullet points. Keep it as bullet lines if it has several points.', selection: true },
      { label: 'Summarize the deck', prompt: 'Summarize this presentation in three sentences, then suggest anything missing.' },
      { label: 'Add a closing slide', prompt: 'Write one closing slide that sums up the presentation, as an outline.' }
    ],
    buttons: [
      { label: 'Add as slides', run: text => aiSlides(text) },
      { label: 'Add to notes', run: text => { snapshot(); const s = slide(); s.notes = (s.notes ? s.notes + '\n' : '') + text.trim(); renderStage(); state.markDirty(); UI.toast('Added to the speaker notes'); } },
      { label: 'Replace text box', selection: true, run: (text, ctx) => {
        const s = pres.slides.find(x => x.id === ctx.slideId) || slide();
        const el = s.elements.find(e => e.id === ctx.elId);
        if (!el) return UI.toast('The text box is gone');
        snapshot();
        const lines = text.trim().split('\n').filter(l => l.trim());
        el.html = lines.length > 1 || /^\s*[-*•]/.test(lines[0] || '') ? toListHtml(lines) : escapeHtml(lines[0] || '');
        renderAll(); state.markDirty(); UI.toast('Replaced');
      } }
    ]
  });

  window.AtlasSlides = { loadFile, serialize, get pres() { return pres; }, present, goTo, addSlide, get live() { return live; } };
})();
