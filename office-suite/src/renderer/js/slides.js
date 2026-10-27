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
  let pres = { theme: 'office', slides: [] };
  let current = 0;
  let selectedId = null;
  let editingId = null;
  const theme = () => THEMES[pres.theme] || THEMES.office;
  const slide = () => pres.slides[current];
  const elById = id => slide().elements.find(e => e.id === id);

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
      default: return [title({ x: 60, y: 28, w: 840, h: 90 }), body({ x: 60, y: 130, w: 840, h: 370 })];
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
  const oneOf = (v, list, d) => (list.includes(v) ? v : d);
  function cleanPres(p) {
    if (!p || !Array.isArray(p.slides)) throw new Error('Not a presentation');
    const slides = p.slides.slice(0, 2000).map(sl => {
      const s = newSlide('blank');
      if (!sl || typeof sl !== 'object') return s;
      s.layout = oneOf(sl.layout, ['title', 'content', 'section', 'two', 'titleonly', 'blank'], 'blank');
      s.bg = color(sl.bg);
      s.bgImage = imageUrl(sl.bgImage);
      s.transition = oneOf(sl.transition, ['none', 'fade', 'push', 'zoom'], 'none');
      s.notes = typeof sl.notes === 'string' ? sl.notes.slice(0, 100000) : '';
      s.elements = (Array.isArray(sl.elements) ? sl.elements : []).slice(0, 500).map(el => {
        if (!el || typeof el !== 'object') return null;
        const type = oneOf(el.type, ['text', 'shape', 'image', 'table'], null);
        if (!type) return null;
        const out = { id: newId(), type, x: num(el.x, 100), y: num(el.y, 100), w: num(el.w, 200, 1), h: num(el.h, 100, 1) };
        if (el.rot) out.rot = num(el.rot, 0, -360, 360);
        if (type === 'image') { out.src = imageUrl(el.src); return out.src ? out : null; }
        out.html = Platform.sanitizeHtml(el.html);
        out.size = num(el.size, 18, 4, 400);
        out.align = oneOf(el.align, ['left', 'center', 'right'], 'left');
        out.valign = oneOf(el.valign, ['top', 'middle', 'bottom'], 'top');
        if (typeof el.font === 'string' && /^[\w .-]{1,64}$/.test(el.font)) out.font = el.font;
        if (color(el.color)) out.color = el.color;
        if (['title', 'subtitle', 'body'].includes(el.ph)) out.ph = el.ph;
        if (el.bold === false) out.bold = false;
        if (type === 'shape') { out.shape = oneOf(el.shape, SHAPE_NAMES, 'rect'); out.fill = color(el.fill); out.line = color(el.line); }
        return out;
      }).filter(Boolean);
      return s;
    });
    if (!slides.length) slides.push(newSlide('title'));
    return { theme: Object.prototype.hasOwnProperty.call(THEMES, p.theme) ? p.theme : 'office', slides };
  }
  const newSlide = kind => ({ id: newId(), layout: kind, bg: null, bgImage: null, transition: 'none', notes: '', elements: makeLayout(kind) });
  const PLACEHOLDER = { title: 'Click to add title', subtitle: 'Click to add subtitle', body: 'Click to add text' };

  // ---------- Undo ----------
  const undoStack = [], redoStack = [];
  function snapshot() {
    undoStack.push(JSON.stringify({ pres, current }));
    if (undoStack.length > 100) undoStack.shift();
    redoStack.length = 0;
  }
  function restore(json) {
    const s = JSON.parse(json);
    pres = s.pres; current = Math.min(s.current, pres.slides.length - 1);
    selectedId = editingId = null;
    renderAll();
    state.markDirty();
  }
  function undo() { if (!undoStack.length) return; redoStack.push(JSON.stringify({ pres, current })); restore(undoStack.pop()); }
  function redo() { if (!redoStack.length) return; undoStack.push(JSON.stringify({ pres, current })); restore(redoStack.pop()); }

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

  /** Builds the DOM for one slide. interactive=false for thumbnails, print and the show. */
  function renderSlide(s, container, interactive) {
    container.innerHTML = '';
    container.style.backgroundColor = s.bg || theme().bg;
    container.style.backgroundImage = imageUrl(s.bgImage) ? `url("${s.bgImage}")` : '';
    s.elements.forEach(el => {
      const div = document.createElement('div');
      div.className = 'el ' + el.type;
      div.dataset.id = el.id;
      Object.assign(div.style, { left: el.x + 'px', top: el.y + 'px', width: el.w + 'px', height: el.h + 'px', transform: el.rot ? `rotate(${el.rot}deg)` : '' });
      if (el.type === 'image') {
        const img = document.createElement('img');
        img.src = el.src;
        img.alt = '';
        div.appendChild(img);
      } else {
        if (el.type === 'shape') {
          // Colours are checked here too, since they go into SVG markup.
          const stroke = color(el.line) || 'none';
          const fill = el.shape === 'line' ? 'none' : (color(el.fill) || 'none');
          div.insertAdjacentHTML('afterbegin', `<svg viewBox="0 0 100 100" preserveAspectRatio="none"><g fill="${fill}" stroke="${el.shape === 'line' ? (color(el.line) || theme().accent) : stroke}" stroke-width="${el.shape === 'line' ? 3 : 1.5}" vector-effect="non-scaling-stroke">${SHAPES[el.shape] ? SHAPES[el.shape]() : SHAPES.rect()}</g></svg>`);
          div.querySelectorAll('svg *').forEach(n => n.setAttribute('vector-effect', 'non-scaling-stroke'));
        }
        const c = document.createElement('div');
        c.className = 'content v-' + (el.valign || 'top');
        c.innerHTML = el.html || '';
        Object.assign(c.style, { fontSize: el.size + 'pt', color: textColor(el), textAlign: el.align || 'left', fontFamily: fontOf(el) });
        if (el.ph === 'title') c.style.fontWeight = el.bold === false ? '400' : '';
        if (el.type === 'table') div.style.setProperty('--tbl-accent', theme().accent);
        if (interactive && el.type === 'text' && el.ph) c.dataset.placeholder = PLACEHOLDER[el.ph];
        else if (interactive && el.type === 'text') c.dataset.placeholder = 'Type here';
        div.appendChild(c);
      }
      if (interactive && el.id === selectedId) {
        div.classList.add('selected');
        if (el.id === editingId) div.classList.add('editing');
        ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'].forEach(h => { const d = document.createElement('div'); d.className = 'handle ' + h; d.dataset.h = h; div.appendChild(d); });
      }
      container.appendChild(div);
    });
  }

  function renderStage() {
    renderSlide(slide(), stage, true);
    if (editingId) {
      const c = stage.querySelector(`[data-id="${editingId}"] .content`);
      if (c) c.contentEditable = 'true';
    }
    notes.value = slide().notes || '';
    document.getElementById('stSlide').textContent = `Slide ${current + 1} of ${pres.slides.length}`;
    syncToolbar();
  }
  function renderThumbs() {
    thumbs.innerHTML = '';
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
  let thumbTimer;
  function refreshThumbSoon() {
    clearTimeout(thumbTimer);
    thumbTimer = setTimeout(() => {
      const frame = thumbs.children[current]?.querySelector('.slide');
      if (frame) renderSlide(slide(), frame, false);
    }, 250);
  }
  function renderAll() { renderStage(); renderThumbs(); fitStage(); }
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
    current = Math.max(0, Math.min(pres.slides.length - 1, i));
    selectedId = null;
    renderAll();
  }

  // ---------- Slide operations ----------
  function addSlide(kind) {
    commitEditing();
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
    const fresh = makeLayout(e.target.value);
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
    if (!elDiv) { commitEditing(); select(null); return; }
    const id = elDiv.dataset.id;
    if (id === editingId && !e.target.classList.contains('handle')) return; // let the caret move
    const el = elById(id);
    if (e.target.classList.contains('handle')) return startResize(e, el, e.target.dataset.h);
    if (el.type === 'text' && !el.html && id !== editingId) { e.preventDefault(); startEditing(id); return; }
    e.preventDefault();
    if (selectedId !== id) select(id);
    startMove(e, el);
  });
  stage.addEventListener('dblclick', e => {
    const elDiv = e.target.closest('.el');
    if (elDiv && elDiv.dataset.id !== editingId) startEditing(elDiv.dataset.id);
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
    const others = slide().elements.filter(o => o.id !== el.id);
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
      el.x = Math.round(nx); el.y = Math.round(ny);
      div.style.left = el.x + 'px'; div.style.top = el.y + 'px';
    };
    const up = () => {
      document.removeEventListener('mousemove', mv); document.removeEventListener('mouseup', up);
      clearGuides();
      if (moved) { state.markDirty(); refreshThumbSoon(); }
    };
    document.addEventListener('mousemove', mv); document.addEventListener('mouseup', up);
  }
  function startResize(e, el, h) {
    e.preventDefault(); e.stopPropagation();
    snapshot();
    const sx = e.clientX, sy = e.clientY, o = { x: el.x, y: el.y, w: el.w, h: el.h };
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
      Object.assign(div.style, { left: el.x + 'px', top: el.y + 'px', width: el.w + 'px', height: el.h + 'px' });
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
    else if (e.key === 'Escape') select(null);
    else if (e.key === 'Enter' || e.key === 'F2') { e.preventDefault(); startEditing(el.id); }
    else if (/^Arrow/.test(e.key)) {
      e.preventDefault();
      snapshot();
      const d = e.shiftKey ? 10 : 1;
      if (e.key === 'ArrowLeft') el.x -= d; if (e.key === 'ArrowRight') el.x += d;
      if (e.key === 'ArrowUp') el.y -= d; if (e.key === 'ArrowDown') el.y += d;
      renderStage(); refreshThumbSoon(); state.markDirty();
    } else if (mod && e.key.toLowerCase() === 'd') { e.preventDefault(); duplicateElement(); }
    else if (mod && e.key.toLowerCase() === 'c') { elClipboard = JSON.parse(JSON.stringify(el)); }
    else if (mod && e.key.toLowerCase() === 'x') { elClipboard = JSON.parse(JSON.stringify(el)); deleteElement(); }
    else if (mod && e.key.toLowerCase() === 'b') { e.preventDefault(); wholeBox('bold'); }
    else if (mod && e.key.toLowerCase() === 'i') { e.preventDefault(); wholeBox('italic'); }
    else if (mod && e.key.toLowerCase() === 'u') { e.preventDefault(); wholeBox('underline'); }
    else if (e.key.length === 1 && !mod && !e.altKey && el.type !== 'image') {
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
    slide().elements = slide().elements.filter(e => e.id !== selectedId);
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
    document.getElementById('layout').value = slide().layout && makeLayout(slide().layout) ? slide().layout : 'content';
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
  document.querySelectorAll('[data-trans]').forEach(b => b.addEventListener('click', () => { slide().transition = b.dataset.trans; syncToolbar(); state.markDirty(); UI.toast('Transition: ' + b.textContent.trim()); }));
  document.getElementById('transAll').addEventListener('click', () => { const t = slide().transition; pres.slides.forEach(s => { s.transition = t; }); state.markDirty(); UI.toast('Applied to all slides'); });
  notes.addEventListener('input', () => { slide().notes = notes.value; state.markDirty(); });
  document.getElementById('toggleNotes').addEventListener('click', e => { document.body.classList.toggle('hide-notes'); e.currentTarget.classList.toggle('on', !document.body.classList.contains('hide-notes')); fitStage(); });

  // ---------- Slide show ----------
  const presenter = document.getElementById('presenter');
  const presentSlide = document.getElementById('presentSlide');
  let showIndex = 0, cursorTimer;
  function present(from) {
    commitEditing();
    showIndex = from;
    presenter.hidden = false;
    presenter.classList.remove('black');
    Platform.fullscreen(true);
    showSlide(showIndex, false);
    setTimeout(() => showSlide(showIndex, false), 300); // refit after going full screen
    document.getElementById('presentHint').textContent = '→ / Space: next   ← : previous   B: black   Esc: end';
    setTimeout(() => { document.getElementById('presentHint').textContent = ''; }, 3000);
  }
  function showSlide(i, animate = true) {
    const s = pres.slides[i];
    renderSlide(s, presentSlide, false);
    const k = Math.min(window.innerWidth / W, window.innerHeight / H);
    presentSlide.style.setProperty('--scale', `scale(${k})`);
    presentSlide.style.transform = `scale(${k})`;
    presentSlide.className = 'slide';
    if (animate && s.transition && s.transition !== 'none') { void presentSlide.offsetWidth; presentSlide.classList.add('t-' + s.transition); }
  }
  function endShow() {
    presenter.hidden = true;
    Platform.fullscreen(false);
    goTo(showIndex);
  }
  function step(d) {
    const n = showIndex + d;
    if (n >= pres.slides.length) return endShow();
    if (n < 0) return;
    showIndex = n;
    showSlide(n, d > 0);
  }
  presenter.addEventListener('click', e => { if (e.button === 0) step(1); });
  presenter.addEventListener('contextmenu', e => { e.preventDefault(); step(-1); });
  presenter.addEventListener('mousemove', () => { presenter.classList.add('show-cursor'); clearTimeout(cursorTimer); cursorTimer = setTimeout(() => presenter.classList.remove('show-cursor'), 1500); });
  document.addEventListener('keydown', e => {
    if (presenter.hidden) return;
    e.preventDefault();
    const k = e.key;
    if (['ArrowRight', 'ArrowDown', ' ', 'PageDown', 'Enter', 'n', 'N'].includes(k)) step(1);
    else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace', 'p', 'P'].includes(k)) step(-1);
    else if (k === 'Home') { showIndex = 0; showSlide(0); }
    else if (k === 'End') { showIndex = pres.slides.length - 1; showSlide(showIndex); }
    else if (k === 'b' || k === 'B' || k === '.') presenter.classList.toggle('black');
    else if (k === 'Escape') endShow();
  }, true);
  window.addEventListener('resize', () => { if (!presenter.hidden) showSlide(showIndex, false); });
  document.addEventListener('fullscreenchange', () => { if (!document.fullscreenElement && !presenter.hidden && !Platform.desktop) endShow(); });

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
      if (s.bgImage) ps.background = { data: s.bgImage };
      else ps.background = { color: hex(s.bg || theme().bg) };
      if (s.transition && s.transition !== 'none') ps.transition = { type: s.transition };
      for (const el of s.elements) {
        const box = { x: inch(el.x), y: inch(el.y), w: inch(el.w), h: inch(el.h), rotate: el.rot || 0 };
        if (el.type === 'image') { ps.addImage({ data: el.src, ...box }); continue; }
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
          if (el.shape === 'line') { shapeOpts.flipV = true; }
          ps.addText(el.html ? htmlToRuns(el.html) : '', shapeOpts);
          continue;
        }
        if (!el.html) continue; // empty placeholders aren't exported
        ps.addText(htmlToRuns(el.html), textOpts);
      }
      if (s.notes) ps.addNotes(s.notes);
    }
    const buf = await pptx.write({ outputType: 'arraybuffer', compression: true });
    return new Uint8Array(buf);
  }

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
    if (cs) [...cs.children].forEach(c => { const v = c.firstElementChild; if (v) scheme[c.localName] = v.getAttribute('val') || v.getAttribute('lastClr'); });
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
    const ids = [...presXml.getElementsByTagNameNS(P, 'sldId')];
    for (const sid of ids) {
      const target = pr[sid.getAttributeNS(R, 'id') || sid.getAttribute('r:id')];
      if (!target) continue;
      const path = 'ppt/' + target.replace(/^\/?ppt\//, '').replace(/^\.\.\//, '');
      const xml = parse(await read(path));
      const relsPath = path.replace(/slides\/(slide\d+\.xml)$/, 'slides/_rels/$1.rels');
      const rels = relMap(parse((await read(relsPath)) || '<Relationships/>'));
      const s = newSlide('blank');
      s.layout = 'blank';
      const bg = xml.getElementsByTagNameNS(P, 'bg')[0];
      if (bg) {
        const blip = bg.getElementsByTagNameNS(A, 'blip')[0];
        if (blip) s.bgImage = await mediaUrl(zip, path, rels[blip.getAttributeNS(R, 'embed')]);
        else s.bg = colorOf(bg);
      }
      const spTree = xml.getElementsByTagNameNS(P, 'spTree')[0];
      if (spTree) {
        for (const node of [...spTree.children]) {
          const el = await importShape(node, zip, path, rels);
          if (el) slide_push(s, el);
        }
      }
      // Speaker notes
      const notesTarget = Object.values(rels).find(t => /notesSlide/.test(t));
      if (notesTarget) {
        const nx = parse((await read('ppt/' + notesTarget.replace(/^\.\.\//, ''))) || '<a/>');
        const bodies = [...nx.getElementsByTagNameNS(P, 'sp')].filter(sp => { const ph = sp.getElementsByTagNameNS(P, 'ph')[0]; return ph && ph.getAttribute('type') === 'body'; });
        s.notes = bodies.map(b => [...b.getElementsByTagNameNS(A, 'p')].map(p => [...p.getElementsByTagNameNS(A, 't')].map(t => t.textContent).join('')).join('\n')).join('\n').trim();
      }
      slides.push(s);
    }
    function slide_push(s, el) { if (Array.isArray(el)) s.elements.push(...el); else s.elements.push(el); }
    async function importShape(node, zipF, path, rels) {
      const ln = node.localName;
      if (ln === 'grpSp') {
        const out = [];
        for (const child of [...node.children]) { const e = await importShape(child, zipF, path, rels); if (e) out.push(...[].concat(e)); }
        return out;
      }
      const xfrm = node.getElementsByTagNameNS(A, 'xfrm')[0];
      const ph = node.getElementsByTagNameNS(P, 'ph')[0];
      const phType = ph ? (ph.getAttribute('type') || 'body') : null;
      let box;
      if (xfrm && xfrm.getElementsByTagNameNS(A, 'off')[0]) {
        const off = xfrm.getElementsByTagNameNS(A, 'off')[0], ext = xfrm.getElementsByTagNameNS(A, 'ext')[0];
        box = { x: Math.round(+off.getAttribute('x') / EMU * sx), y: Math.round(+off.getAttribute('y') / EMU * sy), w: Math.round(+ext.getAttribute('cx') / EMU * sx), h: Math.round(+ext.getAttribute('cy') / EMU * sy) };
        const rot = +xfrm.getAttribute('rot') || 0;
        if (rot) box.rot = rot / 60000;
      } else if (phType) {
        // Placeholder positions come from the layout; use sensible defaults.
        box = /title/i.test(phType) ? (phType === 'ctrTitle' ? { x: 80, y: 140, w: 800, h: 150 } : { x: 60, y: 28, w: 840, h: 90 }) : phType === 'subTitle' ? { x: 120, y: 305, w: 720, h: 80 } : { x: 60, y: 130, w: 840, h: 370 };
      } else return null;
      if (ln === 'pic') {
        const blip = node.getElementsByTagNameNS(A, 'blip')[0];
        const src = blip && await mediaUrl(zipF, path, rels[blip.getAttributeNS(R, 'embed')]);
        return src ? { id: newId(), type: 'image', src, ...box } : null;
      }
      if (ln === 'graphicFrame') {
        const tbl = node.getElementsByTagNameNS(A, 'tbl')[0];
        if (!tbl) return null;
        const html = '<table>' + [...tbl.getElementsByTagNameNS(A, 'tr')].map(tr => '<tr>' + [...tr.getElementsByTagNameNS(A, 'tc')].map(tc => `<td>${escapeHtml([...tc.getElementsByTagNameNS(A, 't')].map(t => t.textContent).join(' '))}</td>`).join('') + '</tr>').join('') + '</table>';
        return { id: newId(), type: 'table', html, size: 14, align: 'left', valign: 'top', ...box };
      }
      if (ln !== 'sp' && ln !== 'cxnSp') return null;
      const spPr = node.getElementsByTagNameNS(P, 'spPr')[0];
      const geom = spPr && spPr.getElementsByTagNameNS(A, 'prstGeom')[0];
      const prst = geom ? geom.getAttribute('prst') : null;
      const fillNode = spPr && [...spPr.children].find(c => c.localName === 'solidFill');
      const lnNode = spPr && spPr.getElementsByTagNameNS(A, 'ln')[0];
      const fill = fillNode ? colorOf(fillNode) : null;
      const line = lnNode && !lnNode.getElementsByTagNameNS(A, 'noFill')[0] ? colorOf(lnNode) : null;
      const txBody = node.getElementsByTagNameNS(P, 'txBody')[0];
      let html = '', size = phType && /title/i.test(phType) ? 40 : 18, align = 'left', valign = 'top';
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
      if (!html && !fill && !line && !prst) return null;
      const isShape = (fill || line || (prst && prst !== 'rect')) && !phType;
      const el = { id: newId(), type: isShape ? 'shape' : 'text', html, size, align, valign, ...box };
      if (phType) el.ph = /title/i.test(phType) ? 'title' : phType === 'subTitle' ? 'subtitle' : 'body';
      if (isShape) { el.shape = { rect: 'rect', roundRect: 'roundRect', ellipse: 'ellipse', triangle: 'triangle', rightArrow: 'rightArrow', line: 'line', star5: 'star5', diamond: 'diamond', straightConnector1: 'line' }[prst] || 'rect'; el.fill = fill; el.line = line; }
      return el;
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
    return cleanPres({ theme: 'office', slides });
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
    undoStack.length = redoStack.length = 0;
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
      case 'save-and-close': if (await save(false)) Platform.closeNow(); break;
      case 'export-pdf': commitEditing(); preparePrint(); Platform.exportPdf(Platform.baseName(state.name) + '.pdf', true); break;
      case 'print': commitEditing(); preparePrint(); Platform.print(); break;
      case 'undo': if (editingId) document.execCommand('undo'); else undo(); break;
      case 'redo': if (editingId) document.execCommand('redo'); else redo(); break;
      case 'present': present(0); break;
      case 'present-current': present(current); break;
      case 'find': UI.toast('Find is available in Write and Sheets'); break;
      case 'home': Platform.desktop ? Platform.newWindow('home') : (location.href = 'index.html'); break;
    }
  });

  (async () => {
    pres.slides = [newSlide('title')];
    renderAll();
    const f = await Platform.pending();
    if (f && f.error) UI.toast('Could not open file: ' + f.error, 4000);
    else if (f && f.recovered) loadRecovered(f.recovered);
    else if (f) await loadFile(f);
  })();

  // ---------- Crash recovery & AutoSave ----------
  const recovery = Platform.enableRecovery({
    app: 'slides', state,
    snapshot: () => { commitEditingQuietly(); return JSON.stringify(pres); },
    autosave: () => save(false, true),
    canAutoSave: () => !Platform.desktop ? 'AutoSave needs the desktop app' : !state.path ? 'Save the presentation first to turn on AutoSave'
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
    undoStack.length = redoStack.length = 0;
    renderAll();
    state.set({ name: d.name, path: d.path, dirty: true });
    recovery.adopt(d.draftId);
    UI.toast('Recovered your unsaved changes', 4000);
  }

  window.AtlasSlides = { loadFile, serialize, get pres() { return pres; }, present, goTo, addSlide };
})();
