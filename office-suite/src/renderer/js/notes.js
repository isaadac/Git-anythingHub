'use strict';
// Atlas Notes — notebooks of sections and pages, like OneNote. Everything saves as you type:
// the desktop app keeps each notebook as a file in its data folder, the web version in the
// browser's storage. Notebooks move between computers as .anote files.
(function () {
  const $ = id => document.getElementById(id);
  const body = $('body'), titleEl = $('pageTitle'), ink = $('ink'), pageEl = $('page');
  const desktop = Platform.desktop;
  const sanitize = html => Platform.sanitizeHtml(html);
  const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const newId = p => p + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
  const COLORS = ['#7719aa', '#2b579a', '#217346', '#b7472a', '#c19c00', '#038387', '#c239b3', '#4a5459'];
  UI.initRibbon();

  // ---------- Storage ----------
  const Store = desktop ? {
    list: () => window.atlas.notesList(),
    read: async id => { const j = await window.atlas.notesRead(id); return j ? JSON.parse(j) : null; },
    write: (id, nb) => window.atlas.notesWrite(id, JSON.stringify(nb)),
    del: id => window.atlas.notesDelete(id)
  } : (() => {
    let dbp = null;
    const db = () => dbp || (dbp = new Promise((res, rej) => {
      const r = indexedDB.open('atlas-notes', 1);
      r.onupgradeneeded = () => r.result.createObjectStore('nb', { keyPath: 'id' });
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
    }));
    const tx = async (mode, fn) => { const d = await db(); return new Promise((res, rej) => { const t = d.transaction('nb', mode); const out = fn(t.objectStore('nb')); t.oncomplete = () => res(out && out.result); t.onerror = () => rej(t.error); }); };
    return {
      list: async () => ((await tx('readonly', s => s.getAll())) || []).map(x => ({ id: x.id, name: x.name, color: x.color, modified: x.modified })).sort((a, b) => b.modified - a.modified),
      read: async id => { const x = await tx('readonly', s => s.get(id)); return x ? JSON.parse(x.json) : null; },
      write: (id, nb) => tx('readwrite', s => s.put({ id, name: nb.name, color: nb.color, modified: Date.now(), json: JSON.stringify(nb) })),
      del: id => tx('readwrite', s => s.delete(id))
    };
  })();

  // ---------- Model ----------
  const newPage = (title = '') => ({ id: newId('p'), title, html: '<p><br></p>', ink: [], level: 0, created: Date.now(), modified: Date.now() });
  const newSection = (name, color) => ({ id: newId('s'), name, color, pages: [newPage()] });
  const newNotebook = name => ({ format: 'atlas-notes', version: 1, id: newId('n'), name, color: COLORS[0], sections: [newSection('Quick Notes', COLORS[1])] });

  /** Checks a notebook from a file: known fields only, page HTML sanitized, sizes limited. */
  function cleanNotebook(x) {
    if (!x || x.format !== 'atlas-notes' || !Array.isArray(x.sections)) throw new Error('This is not an Atlas Notes notebook');
    const str = (v, n) => (typeof v === 'string' ? v.slice(0, n) : '');
    const color = v => (/^#[0-9a-fA-F]{6}$/.test(v) ? v : COLORS[0]);
    const num = v => (Number.isFinite(v) ? v : Date.now());
    const stroke = s => (s && typeof s.d === 'string' && /^[MLQ0-9 .,-]{1,200000}$/.test(s.d) ? { c: color(s.c), w: Math.max(0.5, Math.min(40, +s.w || 2)), o: Math.max(0.1, Math.min(1, +s.o || 1)), d: s.d } : null);
    return {
      format: 'atlas-notes', version: 1,
      id: /^[\w-]{1,40}$/.test(x.id) ? x.id : newId('n'),
      name: str(x.name, 120) || 'Notebook', color: color(x.color),
      sections: x.sections.slice(0, 500).map(s => ({
        id: /^[\w-]{1,40}$/.test(s && s.id) ? s.id : newId('s'), name: str(s && s.name, 120) || 'Section', color: color(s && s.color),
        pages: (Array.isArray(s && s.pages) ? s.pages : []).slice(0, 5000).map(p => ({
          id: /^[\w-]{1,40}$/.test(p && p.id) ? p.id : newId('p'), title: str(p && p.title, 300),
          html: sanitize(str(p && p.html, 50e6)) || '<p><br></p>',
          ink: (Array.isArray(p && p.ink) ? p.ink : []).slice(0, 20000).map(stroke).filter(Boolean),
          level: p && p.level === 1 ? 1 : 0, created: num(p && p.created), modified: num(p && p.modified)
        }))
      })).map(s => (s.pages.length ? s : { ...s, pages: [newPage()] }))
    };
  }

  let nb = null, sec = null, page = null, books = [];

  // ---------- Saving ----------
  let saveTimer = null, saving = Promise.resolve();
  function markChanged() {
    window.__atlasDirty = true;
    $('savedState').textContent = 'Saving…';
    clearTimeout(saveTimer);
    saveTimer = setTimeout(flush, 500);
  }
  async function flush() {
    clearTimeout(saveTimer);
    if (!nb) return;
    commit();
    const snapshot = nb;
    saving = saving.then(() => Store.write(snapshot.id, snapshot)).then(() => {
      if (!saveTimer || snapshot !== nb) { window.__atlasDirty = false; $('savedState').textContent = 'Saved'; }
    }).catch(err => { $('savedState').textContent = 'Not saved'; UI.toast('Could not save the notebook: ' + err.message, 5000); });
    saveTimer = null;
    return saving;
  }
  document.addEventListener('visibilitychange', () => { if (document.hidden) flush(); });
  window.addEventListener('pagehide', () => { flush(); });

  /** Copies what is on screen into the current page. */
  function commit() {
    if (!page) return;
    const html = body.innerHTML, title = titleEl.value.trim();
    if (html !== page.html || title !== page.title) { page.html = html; page.title = title; page.modified = Date.now(); }
  }

  // ---------- Rendering ----------
  const fmtDate = t => new Date(t).toLocaleString(undefined, { dateStyle: 'full', timeStyle: 'short' });
  function renderBooks() {
    const sel = $('nbSelect');
    sel.innerHTML = '';
    const list = books.some(b => b.id === nb.id) ? books : [{ id: nb.id, name: nb.name }, ...books];
    list.forEach(b => { const o = document.createElement('option'); o.value = b.id; o.textContent = b.id === nb.id ? nb.name : b.name; sel.appendChild(o); });
    sel.value = nb.id;
    $('nbName').value = nb.name;
    Platform.setTitle(`${nb.name} — Atlas Notes`);
  }
  function li(text, on, extra) {
    const el = document.createElement('li');
    // A plain list (the ⋯ button can't sit inside an ARIA option); the current item is marked.
    if (on) el.setAttribute('aria-current', 'true');
    el.tabIndex = on ? 0 : -1;
    if (on) el.classList.add('on');
    if (extra) extra(el);
    const t = document.createElement('span');
    t.className = 't';
    t.textContent = text;
    el.appendChild(t);
    const more = document.createElement('button');
    more.className = 'more';
    more.textContent = '⋯';
    more.title = 'More';
    more.setAttribute('aria-label', 'More actions for ' + text);
    el.appendChild(more);
    return el;
  }
  function renderSections() {
    const ul = $('sections');
    ul.innerHTML = '';
    nb.sections.forEach(s => {
      const el = li(s.name, s === sec, e => { const d = document.createElement('span'); d.className = 'dot'; d.style.background = s.color; e.appendChild(d); e.style.borderLeftColor = s === sec ? s.color : 'transparent'; });
      el.addEventListener('click', e => { if (e.target.classList.contains('more')) return sectionMenu(s); openSection(s); });
      el.addEventListener('dblclick', () => renameSection(s));
      ul.appendChild(el);
    });
  }
  function renderPages() {
    const ul = $('pages');
    ul.innerHTML = '';
    sec.pages.forEach(p => {
      const el = li(p.title || 'Untitled page', p === page, e => { if (p.level) e.classList.add('sub'); });
      el.addEventListener('click', e => { if (e.target.classList.contains('more')) return pageMenu(p); openPage(p); });
      ul.appendChild(el);
    });
  }
  function listKeys(ul, items, open) {
    ul.addEventListener('keydown', e => {
      const list = items();
      const i = [...ul.children].indexOf(document.activeElement.closest('li'));
      if (i < 0) return;
      let n = null;
      if (e.key === 'ArrowDown') n = Math.min(list.length - 1, i + 1);
      if (e.key === 'ArrowUp') n = Math.max(0, i - 1);
      if (n !== null) { e.preventDefault(); open(list[n]); ul.children[n] && ul.children[n].focus(); }
      if (e.key === 'Enter') { e.preventDefault(); body.focus(); }
    });
  }
  listKeys($('sections'), () => nb.sections, s => openSection(s));
  listKeys($('pages'), () => sec.pages, p => openPage(p));

  function openSection(s) {
    commit();
    sec = s;
    renderSections();
    openPage(s.pages[0]);
  }
  function openPage(p) {
    commit();
    page = p;
    body.innerHTML = sanitize(p.html) || '<p><br></p>';
    titleEl.value = p.title;
    $('pageDate').textContent = fmtDate(p.created);
    drawInk();
    renderPages();
    updateCount();
    $('canvas').scrollTop = 0;
    try { localStorage.setItem('atlas-notes-last', JSON.stringify({ nb: nb.id, sec: sec.id, page: p.id })); } catch { /* ignore */ }
  }
  async function openNotebook(id) {
    if (nb) await flush();
    let x = null;
    try { x = id && await Store.read(id); } catch { /* unreadable */ }
    nb = x ? cleanNotebook(x) : null;
    if (!nb) { nb = newNotebook('My Notebook'); await Store.write(nb.id, nb); }
    books = await Store.list();
    let last = {};
    try { last = JSON.parse(localStorage.getItem('atlas-notes-last') || '{}'); } catch { /* ignore */ }
    sec = (last.nb === nb.id && nb.sections.find(s => s.id === last.sec)) || nb.sections[0];
    renderBooks();
    renderSections();
    page = null;
    openPage((last.nb === nb.id && sec.pages.find(p => p.id === last.page)) || sec.pages[0]);
  }
  function updateCount() {
    const words = (body.innerText.match(/\S+/g) || []).length;
    const todos = body.querySelectorAll('.todo').length, done = body.querySelectorAll('.todo.done').length;
    $('stCount').textContent = `${words} word${words === 1 ? '' : 's'}${todos ? ` · ${done}/${todos} to-dos done` : ''}`;
  }

  // ---------- Editing ----------
  body.addEventListener('input', () => { markChanged(); updateCount(); });
  titleEl.addEventListener('input', () => {
    markChanged();
    const i = sec.pages.indexOf(page);
    const el = $('pages').children[i];
    if (el) el.querySelector('.t').textContent = titleEl.value.trim() || 'Untitled page';
  });
  titleEl.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); body.focus(); } });
  $('nbName').addEventListener('change', () => {
    nb.name = $('nbName').value.trim().slice(0, 120) || 'Notebook';
    renderBooks();
    markChanged();
  });
  // Pasted content is cleaned; pasted pictures are kept in the page.
  body.addEventListener('paste', async e => {
    const file = [...e.clipboardData.files].find(f => f.type.startsWith('image/'));
    if (file) { e.preventDefault(); insertHtml(`<img src="${await Platform.bytesToDataUrl(new Uint8Array(await file.arrayBuffer()), file.type)}" alt="">`); return; }
    const html = e.clipboardData.getData('text/html');
    if (html) { e.preventDefault(); insertHtml(sanitize(html)); }
  });
  function insertHtml(html) { body.focus(); document.execCommand('insertHTML', false, html); markChanged(); }
  // Ticking a to-do: click its box.
  body.addEventListener('mousedown', e => {
    const t = e.target.closest('.todo');
    if (!t || !body.contains(t)) return;
    const r = t.getBoundingClientRect();
    if (e.clientX - r.left < 24) { e.preventDefault(); t.classList.toggle('done'); markChanged(); updateCount(); }
  });

  // Ribbon
  let savedRange = null;
  document.addEventListener('selectionchange', () => { const s = getSelection(); if (s.rangeCount && body.contains(s.anchorNode)) savedRange = s.getRangeAt(0).cloneRange(); });
  function restoreSel() {
    // Keyboard shortcuts run before selectionchange reports the caret, so only restore when focus left the page.
    const s = getSelection();
    if (s.rangeCount && body.contains(s.anchorNode)) return;
    body.focus();
    if (savedRange) { s.removeAllRanges(); s.addRange(savedRange); }
  }
  document.querySelectorAll('.ribbon [data-exec]').forEach(b => { b.addEventListener('mousedown', e => e.preventDefault()); b.addEventListener('click', () => { restoreSel(); document.execCommand(b.dataset.exec); markChanged(); }); });
  document.querySelectorAll('.ribbon [data-block]').forEach(b => { b.addEventListener('mousedown', e => e.preventDefault()); b.addEventListener('click', () => { restoreSel(); document.execCommand('formatBlock', false, b.dataset.block); markChanged(); }); });
  $('highlight').addEventListener('mousedown', e => e.preventDefault());
  $('highlight').addEventListener('click', () => { restoreSel(); document.execCommand('hiliteColor', false, '#ffff00'); markChanged(); });
  $('fontColor').addEventListener('change', e => { restoreSel(); document.execCommand('foreColor', false, e.target.value); markChanged(); });

  /** The paragraphs (or list items) the selection touches. */
  function selectedBlocks() {
    restoreSel();
    const s = getSelection();
    if (!s.rangeCount) return [];
    const blockOf = n => { let e = n.nodeType === 1 ? n : n.parentElement; while (e && e !== body && !/^(P|LI|H[1-6]|DIV|BLOCKQUOTE)$/.test(e.tagName)) e = e.parentElement; return e === body ? null : e; };
    let a = blockOf(s.anchorNode);
    if (!a) { document.execCommand('formatBlock', false, 'p'); a = blockOf(getSelection().anchorNode); }
    const f = blockOf(s.focusNode) || a;
    if (!a) return [];
    const all = [...body.querySelectorAll('p, li, h1, h2, h3, h4, h5, h6, div, blockquote')].filter(x => !x.querySelector('p, li, div, h1, h2, h3, blockquote'));
    const i = all.indexOf(a), j = all.indexOf(f);
    return i < 0 || j < 0 ? [a] : all.slice(Math.min(i, j), Math.max(i, j) + 1);
  }
  const TAGS = ['todo', 'tag-important', 'tag-question', 'tag-idea'];
  function toggleTag(tag) {
    const blocks = selectedBlocks();
    if (!blocks.length) return;
    const on = !blocks.every(b => b.classList.contains(tag));
    blocks.forEach(b => {
      TAGS.forEach(t => b.classList.remove(t));
      b.classList.remove('done');
      if (on) b.classList.add(tag);
      if (!b.classList.length) b.removeAttribute('class');
    });
    markChanged(); updateCount();
  }
  document.querySelectorAll('.ribbon [data-tag]').forEach(b => { b.addEventListener('mousedown', e => e.preventDefault()); b.addEventListener('click', () => toggleTag(b.dataset.tag)); });

  $('insTable').addEventListener('click', async () => {
    const v = await UI.form('Insert Table', [{ name: 'r', label: 'Rows', type: 'number', value: 3, min: 1, max: 50 }, { name: 'c', label: 'Columns', type: 'number', value: 3, min: 1, max: 20 }], 'Insert');
    if (!v) return;
    const r = Math.max(1, Math.min(50, +v.r || 3)), c = Math.max(1, Math.min(20, +v.c || 3));
    restoreSel();
    insertHtml('<table>' + Array.from({ length: r }, () => '<tr>' + '<td><br></td>'.repeat(c) + '</tr>').join('') + '</table><p><br></p>');
  });
  $('insPicture').addEventListener('click', async () => {
    const f = await Platform.open([{ name: 'Pictures', extensions: ['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'svg'] }]);
    if (!f) return;
    const ext = Platform.extOf(f.name);
    const mime = { jpg: 'image/jpeg', jpeg: 'image/jpeg', svg: 'image/svg+xml' }[ext] || 'image/' + ext;
    restoreSel();
    insertHtml(`<img src="${await Platform.bytesToDataUrl(f.data, mime)}" alt="${esc(Platform.baseName(f.name))}">`);
  });
  $('insLink').addEventListener('click', async () => {
    const sel = savedRange && !savedRange.collapsed ? savedRange.toString() : '';
    const v = await UI.form('Link', [{ name: 'url', label: 'Address', value: 'https://' }, { name: 'text', label: 'Text to show', value: sel }], 'Insert');
    if (!v || !/^(https?:|mailto:)/i.test(v.url.trim())) return;
    restoreSel();
    insertHtml(`<a href="${esc(v.url.trim())}">${esc(v.text.trim() || v.url.trim())}</a>`);
  });
  $('insDate').addEventListener('click', () => { restoreSel(); document.execCommand('insertText', false, new Date().toLocaleString()); markChanged(); });
  $('insRule').addEventListener('click', () => { restoreSel(); document.execCommand('insertHorizontalRule'); markChanged(); });
  $('fullPage').addEventListener('click', () => document.body.classList.toggle('full-page'));
  $('ruled').addEventListener('click', e => { document.body.classList.toggle('ruled'); e.currentTarget.classList.toggle('on'); });

  // ---------- Ink ----------
  let pen = null, erasing = false;
  function drawInk() {
    ink.innerHTML = '';
    page.ink.forEach((s, i) => ink.appendChild(strokeEl(s, i)));
  }
  function strokeEl(s, i) {
    const p = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    p.setAttribute('d', s.d);
    p.setAttribute('stroke', s.c);
    p.setAttribute('stroke-width', s.w);
    if (s.o < 1) p.setAttribute('stroke-opacity', s.o);
    p.dataset.i = i;
    return p;
  }
  function setTool(tool) {
    document.querySelectorAll('.ribbon .pen, #eraser, #drawOff').forEach(b => b.classList.remove('on'));
    pen = null; erasing = false;
    if (tool === 'eraser') { erasing = true; $('eraser').classList.add('on'); }
    else if (tool) { pen = { c: tool.dataset.pen, w: +tool.dataset.w, o: tool.dataset.hl ? 0.45 : 1 }; tool.classList.add('on'); }
    else $('drawOff').classList.add('on');
    document.body.classList.toggle('drawing', !!(pen || erasing));
    document.body.classList.toggle('erasing', erasing);
    $('stMode').textContent = pen ? 'Drawing' : erasing ? 'Erasing' : 'Typing';
  }
  document.querySelectorAll('.ribbon .pen').forEach(b => b.addEventListener('click', () => setTool(b)));
  $('eraser').addEventListener('click', () => setTool('eraser'));
  $('drawOff').addEventListener('click', () => setTool(null));
  $('clearInk').addEventListener('click', async () => {
    if (!page.ink.length || !(await UI.confirm('Clear drawing', 'Remove all drawing from this page?'))) return;
    page.ink = []; drawInk(); markChanged();
  });
  const at = e => { const r = pageEl.getBoundingClientRect(); return [Math.round((e.clientX - r.left) * 10) / 10, Math.round((e.clientY - r.top) * 10) / 10]; };
  ink.addEventListener('pointerdown', e => {
    if (!pen && !erasing) return;
    e.preventDefault();
    ink.setPointerCapture(e.pointerId);
    if (erasing) { erase(e); return; }
    const [x, y] = at(e);
    const s = { ...pen, d: `M${x} ${y}` };
    const el = strokeEl(s, page.ink.length);
    ink.appendChild(el);
    let last = [x, y];
    const move = ev => {
      const [nx, ny] = at(ev);
      if (Math.hypot(nx - last[0], ny - last[1]) < 1.5) return;
      s.d += ` L${nx} ${ny}`;
      last = [nx, ny];
      el.setAttribute('d', s.d);
    };
    const up = () => {
      ink.removeEventListener('pointermove', move);
      ink.removeEventListener('pointerup', up);
      ink.removeEventListener('pointercancel', up);
      if (!/L/.test(s.d)) s.d += ` L${x + 0.1} ${y + 0.1}`; // a dot
      el.setAttribute('d', s.d);
      page.ink.push(s);
      page.modified = Date.now();
      markChanged();
    };
    ink.addEventListener('pointermove', move);
    ink.addEventListener('pointerup', up);
    ink.addEventListener('pointercancel', up);
  });
  function erase(e) {
    const hit = () => {
      const [x, y] = at(e);
      // A stroke is hit when the eraser passes near any of its points.
      const near = page.ink.findIndex(s => {
        const pts = (s.d.match(/-?[\d.]+ -?[\d.]+/g) || []).map(pt => pt.split(' ').map(Number));
        const r = 8 + s.w / 2;
        // Distance to each segment of the stroke (fast strokes have points far apart).
        return pts.some((a, i) => {
          const b = pts[i + 1] || a;
          const dx = b[0] - a[0], dy = b[1] - a[1], len = dx * dx + dy * dy;
          const t = len ? Math.max(0, Math.min(1, ((x - a[0]) * dx + (y - a[1]) * dy) / len)) : 0;
          return Math.hypot(a[0] + t * dx - x, a[1] + t * dy - y) < r;
        });
      });
      if (near >= 0) { page.ink.splice(near, 1); page.modified = Date.now(); drawInk(); markChanged(); }
    };
    hit();
    const move = ev => { e = ev; hit(); };
    const up = () => { ink.removeEventListener('pointermove', move); ink.removeEventListener('pointerup', up); };
    ink.addEventListener('pointermove', move);
    ink.addEventListener('pointerup', up);
  }

  // ---------- Sections and pages ----------
  async function renameSection(s) {
    const n = await UI.prompt('Rename Section', 'Section name', s.name);
    if (!n || !n.trim()) return;
    s.name = n.trim().slice(0, 120); renderSections(); markChanged();
  }
  async function sectionMenu(s) {
    const i = nb.sections.indexOf(s);
    const acts = ['Rename', 'Colour', 'Move up', 'Move down', 'Delete', 'Cancel'];
    const a = acts[await UI.choose(s.name, 'Section', acts)];
    if (a === 'Rename') return renameSection(s);
    if (a === 'Colour') { s.color = COLORS[(COLORS.indexOf(s.color) + 1) % COLORS.length]; }
    if (a === 'Move up' && i > 0) nb.sections.splice(i - 1, 0, nb.sections.splice(i, 1)[0]);
    if (a === 'Move down' && i < nb.sections.length - 1) nb.sections.splice(i + 1, 0, nb.sections.splice(i, 1)[0]);
    if (a === 'Delete') {
      if (nb.sections.length === 1) return UI.toast('A notebook needs at least one section');
      if (!(await UI.confirm('Delete section', `Delete “${s.name}” and its ${s.pages.length} page${s.pages.length === 1 ? '' : 's'}?`))) return;
      nb.sections.splice(i, 1);
      if (sec === s) { page = null; sec = nb.sections[Math.min(i, nb.sections.length - 1)]; openPage(sec.pages[0]); }
    }
    if (!a || a === 'Cancel') return;
    renderSections(); markChanged();
  }
  async function pageMenu(p) {
    const i = sec.pages.indexOf(p);
    const acts = [p.level ? 'Promote' : 'Make subpage', 'Move up', 'Move down', 'Move to section…', 'Delete', 'Cancel'];
    const a = acts[await UI.choose(p.title || 'Untitled page', 'Page', acts)];
    if (a === 'Make subpage' && i > 0) p.level = 1;
    if (a === 'Promote') p.level = 0;
    if (a === 'Move up' && i > 0) sec.pages.splice(i - 1, 0, sec.pages.splice(i, 1)[0]);
    if (a === 'Move down' && i < sec.pages.length - 1) sec.pages.splice(i + 1, 0, sec.pages.splice(i, 1)[0]);
    if (a === 'Move to section…') {
      const others = nb.sections.filter(s => s !== sec);
      if (!others.length) return UI.toast('Add another section first');
      const v = await UI.form('Move page', [{ name: 's', label: 'To section', type: 'select', options: others.map(s => ({ value: s.id, label: s.name })) }], 'Move');
      if (!v) return;
      commit();
      const to = others.find(s => s.id === v.s);
      sec.pages.splice(i, 1);
      p.level = 0;
      to.pages.push(p);
      if (!sec.pages.length) sec.pages.push(newPage());
      if (page === p) { page = null; openPage(sec.pages[Math.min(i, sec.pages.length - 1)]); }
    }
    if (a === 'Delete') {
      if (!(await UI.confirm('Delete page', `Delete “${p.title || 'Untitled page'}”?`))) return;
      sec.pages.splice(i, 1);
      if (!sec.pages.length) sec.pages.push(newPage());
      if (page === p) { page = null; openPage(sec.pages[Math.min(i, sec.pages.length - 1)]); }
    }
    if (!a || a === 'Cancel') return;
    renderPages(); markChanged();
  }
  function addPage() {
    commit();
    const p = newPage();
    const i = sec.pages.indexOf(page);
    sec.pages.splice(i + 1, 0, p);
    openPage(p);
    markChanged();
    titleEl.focus();
  }
  $('pageAdd').addEventListener('click', addPage);
  $('secAdd').addEventListener('click', async () => {
    const n = await UI.prompt('New Section', 'Section name', 'New Section ' + (nb.sections.length + 1));
    if (!n || !n.trim()) return;
    commit();
    const s = newSection(n.trim().slice(0, 120), COLORS[nb.sections.length % COLORS.length]);
    nb.sections.push(s);
    openSection(s);
    markChanged();
  });
  async function addNotebook() {
    const n = await UI.prompt('New Notebook', 'Notebook name', 'Notebook ' + (books.length + 1));
    if (!n || !n.trim()) return;
    await flush();
    const x = newNotebook(n.trim().slice(0, 120));
    await Store.write(x.id, x);
    await openNotebook(x.id);
  }
  $('nbAdd').addEventListener('click', addNotebook);
  $('nbSelect').addEventListener('change', e => openNotebook(e.target.value));

  // ---------- Search ----------
  const searchEl = $('search');
  let searchTimer;
  searchEl.addEventListener('input', () => { clearTimeout(searchTimer); searchTimer = setTimeout(search, 150); });
  searchEl.addEventListener('keydown', e => { if (e.key === 'Escape') { searchEl.value = ''; search(); body.focus(); } });
  function textOf(html) { const d = document.createElement('div'); d.innerHTML = html; d.querySelectorAll('p, li, h1, h2, h3, div, tr, br').forEach(n => n.after(' ')); return d.textContent.replace(/\s+/g, ' ').trim(); }
  function search() {
    commit();
    const q = searchEl.value.trim().toLowerCase();
    const res = $('results');
    $('nav').querySelector('.nb-cols').hidden = !!q;
    res.hidden = !q;
    if (!q) return;
    res.innerHTML = '';
    let n = 0;
    for (const s of nb.sections) for (const p of s.pages) {
      const text = textOf(p.html), t = (p.title || '').toLowerCase();
      const i = text.toLowerCase().indexOf(q);
      if (i < 0 && !t.includes(q)) continue;
      n++;
      const el = document.createElement('li');
      el.tabIndex = 0;
      const h = document.createElement('b'); h.textContent = p.title || 'Untitled page';
      const w = document.createElement('span'); w.className = 'where'; w.textContent = s.name;
      const sn = document.createElement('span'); sn.className = 'snip';
      if (i >= 0) {
        const a = Math.max(0, i - 40);
        sn.append((a ? '…' : '') + text.slice(a, i));
        const m = document.createElement('mark'); m.textContent = text.slice(i, i + q.length);
        sn.append(m, text.slice(i + q.length, i + q.length + 60) + '…');
      }
      el.append(h, w, sn);
      const go = () => { if (sec !== s) { commit(); sec = s; renderSections(); } openPage(p); findInPage(q); };
      el.addEventListener('click', go);
      el.addEventListener('keydown', e => { if (e.key === 'Enter') go(); });
      res.appendChild(el);
    }
    if (!n) { const el = document.createElement('li'); el.className = 'none'; el.textContent = 'No pages match'; res.appendChild(el); }
    A11y.announce(`${n} page${n === 1 ? '' : 's'} found`);
  }
  function findInPage(q) {
    const walk = document.createTreeWalker(body, NodeFilter.SHOW_TEXT);
    let t;
    while ((t = walk.nextNode())) {
      const i = t.data.toLowerCase().indexOf(q);
      if (i >= 0) {
        const r = document.createRange(); r.setStart(t, i); r.setEnd(t, i + q.length);
        body.focus();
        const s = getSelection(); s.removeAllRanges(); s.addRange(r);
        t.parentElement.scrollIntoView({ block: 'center' });
        return;
      }
    }
  }

  // ---------- Export and import ----------
  function pageHtmlDoc(p) {
    const inkSvg = p.ink.length ? `<svg xmlns="http://www.w3.org/2000/svg" style="position:absolute;left:0;top:0;width:100%;height:100%;overflow:visible;pointer-events:none">${p.ink.map(s => `<path d="${s.d}" stroke="${s.c}" stroke-width="${s.w}"${s.o < 1 ? ` stroke-opacity="${s.o}"` : ''} fill="none" stroke-linecap="round" stroke-linejoin="round"/>`).join('')}</svg>` : '';
    return `<!doctype html><html><head><meta charset="utf-8"><title>${esc(p.title || 'Untitled page')}</title>
<style>body{font:14px/1.6 Calibri,Arial,sans-serif;max-width:860px;margin:auto;padding:28px 48px;position:relative}.todo{list-style:none}.todo::before{content:'☐ '}.todo.done::before{content:'☑ '}.todo.done{text-decoration:line-through;color:#777}.tag-important::before{content:'★ ';color:#e3a400}.tag-question::before{content:'? '}.tag-idea::before{content:'💡 '}table{border-collapse:collapse}td,th{border:1px solid #ccc;padding:4px 8px}</style></head>
<body><h1 style="font-weight:300">${esc(p.title || 'Untitled page')}</h1><p style="color:#777;font-size:12px">${esc(fmtDate(p.created))}</p>${p.html}${inkSvg}</body></html>`;
  }
  function toMarkdown(html) {
    const d = document.createElement('div');
    d.innerHTML = html;
    const inline = n => [...n.childNodes].map(c => {
      if (c.nodeType === 3) return c.data.replace(/([*_`[\]])/g, '\\$1');
      if (c.nodeType !== 1) return '';
      const t = inline(c);
      const tag = c.tagName;
      if (tag === 'B' || tag === 'STRONG') return t.trim() ? `**${t}**` : t;
      if (tag === 'I' || tag === 'EM') return t.trim() ? `*${t}*` : t;
      if (tag === 'S' || tag === 'STRIKE' || tag === 'DEL') return `~~${t}~~`;
      if (tag === 'A') return `[${t}](${c.getAttribute('href') || ''})`;
      if (tag === 'IMG') return `![${c.getAttribute('alt') || ''}](${(c.getAttribute('src') || '').startsWith('data:') ? 'image' : c.getAttribute('src')})`;
      if (tag === 'BR') return '  \n';
      return t;
    }).join('');
    const prefix = el => (el.classList.contains('todo') ? (el.classList.contains('done') ? '- [x] ' : '- [ ] ') : el.classList.contains('tag-important') ? '★ ' : el.classList.contains('tag-question') ? '? ' : el.classList.contains('tag-idea') ? '💡 ' : '');
    const out = [];
    const block = (el, depth = 0) => {
      const tag = el.tagName;
      if (/^H[1-6]$/.test(tag)) out.push('#'.repeat(+tag[1] + 1) + ' ' + inline(el).trim(), '');
      else if (tag === 'UL' || tag === 'OL') { [...el.children].forEach((li, i) => { const pre = li.classList.contains('todo') ? prefix(li) : (tag === 'OL' ? `${i + 1}. ` : '- '); out.push('  '.repeat(depth) + pre + inline(li).trim()); li.querySelectorAll(':scope > ul, :scope > ol').forEach(sub => block(sub, depth + 1)); }); out.push(''); }
      else if (tag === 'TABLE') { const rows = [...el.rows].map(r => [...r.cells].map(c => inline(c).trim().replace(/\|/g, '\\|'))); if (rows.length) { out.push('| ' + rows[0].join(' | ') + ' |', '|' + rows[0].map(() => ' --- |').join('')); rows.slice(1).forEach(r => out.push('| ' + r.join(' | ') + ' |')); out.push(''); } }
      else if (tag === 'HR') out.push('---', '');
      else { const t = inline(el).trim(); if (t || prefix(el)) out.push(prefix(el) + t, ''); }
    };
    [...d.children].forEach(el => block(el));
    return out.join('\n').replace(/\n{3,}/g, '\n\n').trim() + '\n';
  }
  const safeName = s => (s || 'Untitled page').replace(/[\\/:*?"<>|]/g, '_').slice(0, 100);

  async function importNotebook(name, data) {
    let x;
    try { x = cleanNotebook(JSON.parse(new TextDecoder().decode(data))); } catch (err) { return UI.toast('Could not open the notebook: ' + err.message, 5000); }
    if (books.some(b => b.id === x.id)) {
      const i = await UI.choose('Open notebook', `“${x.name}” is already in Atlas Notes.`, ['Replace it', 'Keep both', 'Cancel']);
      if (i === 2 || i < 0) return;
      if (i === 1) { x.id = newId('n'); x.name += ' (copy)'; }
    }
    await flush();
    await Store.write(x.id, x);
    await openNotebook(x.id);
    UI.toast(`Opened “${x.name}”`);
  }

  Platform.onMenu(async cmd => {
    switch (cmd) {
      case 'new': addPage(); break;
      case 'new-notebook': addNotebook(); break;
      case 'save': await flush(); UI.toast('Saved'); break;
      case 'save-and-close': await flush(); Platform.closeNow(); break;
      case 'open': {
        const f = await Platform.open([{ name: 'Atlas Notes notebook', extensions: ['anote'] }]);
        if (f) importNotebook(f.name, f.data);
        break;
      }
      case 'save-as': {
        await flush();
        await Platform.save({ data: new TextEncoder().encode(JSON.stringify(nb)), name: safeName(nb.name) + '.anote', filters: [{ name: 'Atlas Notes notebook', extensions: ['anote'] }], saveAs: true });
        break;
      }
      case 'export-md': commit(); await Platform.save({ data: new TextEncoder().encode(`# ${page.title || 'Untitled page'}\n\n${toMarkdown(page.html)}`), name: safeName(page.title) + '.md', filters: [{ name: 'Markdown', extensions: ['md'] }], saveAs: true }); break;
      case 'export-html': commit(); await Platform.save({ data: new TextEncoder().encode(pageHtmlDoc(page)), name: safeName(page.title) + '.html', filters: [{ name: 'Web page', extensions: ['html'] }], saveAs: true }); break;
      case 'open-in-write': {
        commit();
        // Write has no tags: they become text it keeps.
        const d = document.createElement('div');
        d.innerHTML = `<h1>${esc(page.title || 'Untitled page')}</h1>` + page.html;
        d.querySelectorAll('.todo').forEach(t => { t.prepend(t.classList.contains('done') ? '☑ ' : '☐ '); t.removeAttribute('class'); });
        d.querySelectorAll('.tag-important, .tag-question, .tag-idea').forEach(t => { t.prepend(t.classList.contains('tag-important') ? '★ ' : t.classList.contains('tag-question') ? '? ' : '💡 '); t.removeAttribute('class'); });
        Platform.openData('write', safeName(page.title) + '.html', new TextEncoder().encode(`<!doctype html><html><body>${d.innerHTML}</body></html>`), !Platform.desktop);
        break;
      }
      case 'print': await flush(); Platform.print(); break;
      case 'delete-notebook': {
        if (!(await UI.confirm('Delete notebook', `Delete “${nb.name}” with all its sections and pages?${desktop ? ' (It goes to the notebooks trash folder.)' : ''}`))) return;
        clearTimeout(saveTimer); saveTimer = null;
        await Store.del(nb.id);
        books = await Store.list();
        nb = null;
        await openNotebook(books[0] && books[0].id);
        break;
      }
      case 'undo': document.execCommand('undo'); break;
      case 'redo': document.execCommand('redo'); break;
      case 'find': searchEl.focus(); break;
      case 'home': await flush(); Platform.desktop ? Platform.newWindow('home') : (location.href = 'index.html'); break;
    }
  });
  document.addEventListener('keydown', e => {
    const mod = e.ctrlKey || e.metaKey;
    if (!mod || document.querySelector('.modal-backdrop')) return;
    const k = e.key.toLowerCase();
    if (/^[1-4]$/.test(k) && !e.shiftKey && !e.altKey) { e.preventDefault(); toggleTag(TAGS[+k - 1]); }
    else if (k === 'e') { e.preventDefault(); searchEl.focus(); searchEl.select(); }
    else if (k === 'n' && !Platform.desktop) { e.preventDefault(); addPage(); }
  });
  document.addEventListener('keydown', e => { if (e.key === 'F11') { e.preventDefault(); document.body.classList.toggle('full-page'); } });

  // Phones: the notebook list slides over the page.
  const navBtn = document.createElement('button');
  navBtn.className = 'share-btn nav-btn';
  navBtn.textContent = '☰ Pages';
  navBtn.setAttribute('aria-label', 'Show sections and pages');
  document.querySelector('.titlebar').insertBefore(navBtn, document.querySelector('.titlebar .qa'));
  navBtn.addEventListener('click', () => document.body.classList.toggle('nav-open'));
  $('nav').addEventListener('click', e => { if (innerWidth <= 760 && e.target.closest('#pages li, .results li')) document.body.classList.remove('nav-open'); });

  // ---------- Assistant and proofing ----------
  Assistant.attach({
    app: 'notes',
    context: () => { commit(); const r = savedRange && !savedRange.collapsed ? savedRange.toString() : ''; return { document: `# ${page.title || 'Untitled page'}\n\n${toMarkdown(page.html)}`, selection: r, hasSelection: !!r.trim() }; },
    format: 'In Atlas Notes you may use Markdown: # headings, **bold**, - lists, and "- [ ] " for to-do items.',
    actions: [
      { label: 'Summarize page', prompt: 'Summarize this page in a few bullet points.' },
      { label: 'To-do list', prompt: 'Turn the action items on this page into a to-do list ("- [ ] " items). Reply with only the list.' },
      { label: 'Tidy up notes', prompt: 'Rewrite these notes so they are well organized with short headings and bullet points. Keep every fact.' },
      { label: 'Questions to ask', prompt: 'What open questions or gaps are in these notes? List them.' }
    ],
    buttons: [
      { label: 'Insert', run: text => { restoreSel(); insertHtml(mdToHtml(text)); } },
      { label: 'New page', run: text => { addPage(); const first = (/^#\s+(.+)$/m.exec(text) || [])[1]; titleEl.value = first || 'From the Assistant'; body.innerHTML = sanitize(mdToHtml(text.replace(/^#\s+.+\n?/, ''))); markChanged(); renderPages(); } }
    ]
  });
  function mdToHtml(md) {
    const lines = md.replace(/\r/g, '').split('\n');
    let out = '', list = null;
    const inl = s => esc(s).replace(/\*\*([^*]+)\*\*/g, '<b>$1</b>').replace(/\*([^*]+)\*/g, '<i>$1</i>').replace(/\[([^\]]+)\]\((https?:[^)\s]+)\)/g, '<a href="$2">$1</a>');
    const close = () => { if (list) { out += `</${list}>`; list = null; } };
    for (const l of lines) {
      let m;
      if ((m = /^\s*[-*] \[( |x|X)\] (.*)$/.exec(l))) { close(); out += `<p class="todo${m[1].trim() ? ' done' : ''}">${inl(m[2])}</p>`; }
      else if ((m = /^\s*[-*•] (.*)$/.exec(l))) { if (list !== 'ul') { close(); out += '<ul>'; list = 'ul'; } out += `<li>${inl(m[1])}</li>`; }
      else if ((m = /^\s*\d+[.)] (.*)$/.exec(l))) { if (list !== 'ol') { close(); out += '<ol>'; list = 'ol'; } out += `<li>${inl(m[1])}</li>`; }
      else if ((m = /^(#{1,3}) (.*)$/.exec(l))) { close(); out += `<h${Math.min(2, m[1].length)}>${inl(m[2])}</h${Math.min(2, m[1].length)}>`; }
      else if (l.trim()) { close(); out += `<p>${inl(l)}</p>`; }
      else close();
    }
    close();
    return out || '<p><br></p>';
  }
  Proofing.attach({ app: 'notes', roots: () => [{ el: body, live: true }], lang: () => '', setLang: () => {} });

  // ---------- Start ----------
  (async () => {
    setTool(null);
    const f = await Platform.pending();
    books = await Store.list();
    if (f && f.data && /\.anote$/i.test(f.name)) { await openNotebook(books[0] && books[0].id); await importNotebook(f.name, f.data); return; }
    let last = {};
    try { last = JSON.parse(localStorage.getItem('atlas-notes-last') || '{}'); } catch { /* ignore */ }
    await openNotebook((books.find(b => b.id === last.nb) || books[0] || {}).id);
    $('savedState').textContent = 'Saved';
  })();

  window.AtlasNotes = { get notebook() { commit(); return nb; }, flush, toMarkdown, open: openNotebook, get page() { return page; } };
})();
