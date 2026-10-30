// Isadek PDF — read, annotate, fill, sign and organize PDF documents.
//
// Pages are drawn with Mozilla's pdf.js (with selectable text and fillable form fields); saving
// uses pdf-lib: page order, rotation and deletions, form values, and the annotations made here
// (highlights, drawings, text, comments and signatures) written into a standard PDF.
//
// library folders: node_modules/pdfjs-dist/cmaps/ node_modules/pdfjs-dist/standard_fonts/ node_modules/pdfjs-dist/wasm/ node_modules/pdfjs-dist/web/images/
// library files: node_modules/pdfjs-dist/build/pdf.worker.min.mjs node_modules/pdfjs-dist/web/pdf_viewer.mjs
import * as pdfjs from '../../../node_modules/pdfjs-dist/build/pdf.min.mjs';

globalThis.pdfjsLib = pdfjs;
const LIB = new URL('../../../node_modules/pdfjs-dist/', import.meta.url).href;
pdfjs.GlobalWorkerOptions.workerSrc = LIB + 'build/pdf.worker.min.mjs';
const viewer = await import('../../../node_modules/pdfjs-dist/web/pdf_viewer.mjs');
const { PDFDocument, rgb, degrees, StandardFonts, PDFName, PDFString, PDFArray, BlendMode, PDFTextField, PDFCheckBox, PDFRadioGroup, PDFDropdown, PDFOptionList } = window.PDFLib;

const $ = id => document.getElementById(id);
const state = UI.docState('Isadek PDF', 'Document.pdf');
UI.initRibbon();
const eventBus = new viewer.EventBus();
const linkService = new viewer.SimpleLinkService({ eventBus, externalLinkTarget: 2 /* new window */, externalLinkRel: 'noopener noreferrer nofollow' });

// ---------- The document ----------
// sources: the PDFs pages come from ({ bytes, doc }); pages: in order, each with its annotations,
// in the page's own unrotated coordinates (points, from the top left).
let sources = [], pages = [], pageSeq = 0;
let selected = new Set();
let scale = 1, zoomMode = '1', tool = 'select';
const undoStack = [];
const OPTS = { cMapUrl: LIB + 'cmaps/', cMapPacked: true, standardFontDataUrl: LIB + 'standard_fonts/', wasmUrl: LIB + 'wasm/', isEvalSupported: false, enableXfa: false };

async function openSource(bytes, name) {
  const data = new Uint8Array(bytes);
  const task = pdfjs.getDocument({ data: data.slice(), ...OPTS });
  // A password-protected PDF asks for its password.
  task.onPassword = async (update, reason) => {
    const r = await UI.form('Password', [{ name: 'pw', label: `${reason === pdfjs.PasswordResponses.INCORRECT_PASSWORD ? 'That password isn’t right. ' : ''}“${name}” is protected with a password. Enter it to open the document.`, type: 'password' }], 'Open');
    if (!r) { task.destroy(); return; }
    update(r.pw);
  };
  const doc = await task.promise;
  return { bytes: data, doc, encrypted: !!(await doc.getMetadata().catch(() => ({})))?.info?.EncryptFilterName };
}

async function loadFile(f) {
  let src;
  try {
    if (!/^%PDF-/.test(new TextDecoder().decode(new Uint8Array(f.data).subarray(0, 1024)).replace(/^[^%]*/, ''))) throw new Error('This isn’t a PDF file');
    src = await openSource(f.data, f.name);
  } catch (err) {
    if (err && /destroyed|cancel/i.test(err.message)) return;
    UI.toast('Could not open the PDF: ' + (err.message || err), 5000);
    return;
  }
  sources.forEach(s => s.doc.destroy());
  sources = [src];
  linkService.setDocument(src.doc);
  pages = [];
  for (let i = 0; i < src.doc.numPages; i++) pages.push({ id: ++pageSeq, src: 0, index: i, rot: 0, annots: [] });
  selected = new Set([pages[0] && pages[0].id]);
  undoStack.length = 0;
  state.set({ name: f.name, path: f.path || null, dirty: false });
  Platform.rememberRecent(f.name, 'pdf', f.path);
  $('empty').hidden = true;
  await layout();
}

// ---------- Drawing the pages ----------
const view = $('view'), pagesEl = $('pages'), thumbs = $('thumbs');
let observer = null;
const pdfPage = p => sources[p.src].doc.getPage(p.index + 1);
async function baseSize(p) {
  const pg = await pdfPage(p);
  const v = pg.getViewport({ scale: 1, rotation: 0 });
  return { w: v.width, h: v.height, pg };
}
/** The viewport for page p as shown (zoom and rotation, including the page's own). */
const viewportOf = (pg, p) => pg.getViewport({ scale, rotation: (pg.rotate + p.rot) % 360 });

async function layout() {
  if (observer) observer.disconnect();
  pagesEl.innerHTML = '';
  thumbs.innerHTML = '';
  if (!pages.length) { $('empty').hidden = false; updateStatus(); return; }
  // Zoom to fit, from the first page's size.
  if (zoomMode === 'fit' || zoomMode === 'page') {
    const { pg } = await baseSize(pages[0]);
    const v = pg.getViewport({ scale: 1, rotation: (pg.rotate + pages[0].rot) % 360 });
    const avail = view.clientWidth - 48, availH = view.clientHeight - 32;
    scale = Math.max(0.2, Math.min(4, zoomMode === 'fit' ? avail / v.width : Math.min(avail / v.width, availH / v.height)));
  } else scale = +zoomMode;
  observer = new IntersectionObserver(entries => entries.forEach(e => { if (e.isIntersecting) renderPage(e.target); }), { root: view, rootMargin: '600px 0px' });
  for (const p of pages) {
    const pg = await pdfPage(p);
    const v = viewportOf(pg, p);
    const box = document.createElement('div');
    box.className = 'pdf-page';
    box.dataset.id = p.id;
    box.style.width = Math.floor(v.width) + 'px';
    box.style.height = Math.floor(v.height) + 'px';
    box.setAttribute('aria-label', `Page ${pages.indexOf(p) + 1}`);
    pagesEl.appendChild(box);
    observer.observe(box);
    thumbs.appendChild(thumbFor(p));
  }
  updateStatus();
  $('stZoom').textContent = Math.round(scale * 100) + '%';
}

async function renderPage(box) {
  const p = pages.find(x => x.id === +box.dataset.id);
  if (!p || box.dataset.drawn === String(scale) + ':' + p.rot) return;
  box.dataset.drawn = String(scale) + ':' + p.rot;
  const pg = await pdfPage(p);
  const v = viewportOf(pg, p);
  box.innerHTML = '';
  const canvas = document.createElement('canvas');
  const ratio = window.devicePixelRatio || 1;
  canvas.width = Math.floor(v.width * ratio); canvas.height = Math.floor(v.height * ratio);
  canvas.style.width = Math.floor(v.width) + 'px'; canvas.style.height = Math.floor(v.height) + 'px';
  box.appendChild(canvas);
  await pg.render({ canvas, viewport: v, transform: ratio !== 1 ? [ratio, 0, 0, ratio, 0, 0] : null, annotationMode: pdfjs.AnnotationMode.ENABLE_FORMS }).promise.catch(() => {});
  // Selectable text.
  const text = document.createElement('div');
  text.className = 'textLayer';
  box.appendChild(text);
  box.style.setProperty('--scale-factor', scale);
  box.style.setProperty('--total-scale-factor', scale);
  const tl = new pdfjs.TextLayer({ textContentSource: pg.streamTextContent(), container: text, viewport: v });
  await tl.render().catch(() => {});
  // Fillable form fields (and links).
  const ab = new viewer.AnnotationLayerBuilder({ pdfPage: pg, linkService, annotationStorage: sources[p.src].doc.annotationStorage, renderForms: true, onAppend: div => box.appendChild(div) });
  await ab.render({ viewport: v.clone({ dontFlip: true }) }).catch(() => {});
  // Annotations made here, drawn over everything.
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.classList.add('pdf-annots');
  svg.setAttribute('width', v.width); svg.setAttribute('height', v.height);
  box.appendChild(svg);
  drawAnnots(p, box);
  sources[p.src].doc.annotationStorage.onSetModified = () => state.markDirty();
}

// Page coordinates (unrotated, points from the top left) ↔ the screen box.
async function toPage(p, clientX, clientY) {
  const box = boxOf(p), r = box.getBoundingClientRect(), pg = await pdfPage(p), v = viewportOf(pg, p);
  const [x, y] = v.convertToPdfPoint(clientX - r.left, clientY - r.top);
  const h = pg.getViewport({ scale: 1, rotation: 0 }).height, view0 = pg.view;
  return [x - view0[0], h - (y - view0[1])];
}
function toScreenSync(v, pg, x, y) {
  const h = pg.getViewport({ scale: 1, rotation: 0 }).height, view0 = pg.view;
  return v.convertToViewportPoint(x + view0[0], h - y + view0[1]);
}
const boxOf = p => pagesEl.querySelector(`.pdf-page[data-id="${p.id}"]`);

async function drawAnnots(p, box = boxOf(p)) {
  const svg = box && box.querySelector('.pdf-annots');
  if (!svg) return;
  const pg = await pdfPage(p), v = viewportOf(pg, p);
  svg.innerHTML = '';
  box.querySelectorAll('.pdf-note, .pdf-textbox').forEach(n => n.remove());
  const NS = 'http://www.w3.org/2000/svg';
  const pt = (x, y) => toScreenSync(v, pg, x, y);
  p.annots.forEach((a, i) => {
    if (a.type === 'highlight') {
      for (const r of a.rects) {
        const [x1, y1] = pt(r[0], r[1]), [x2, y2] = pt(r[0] + r[2], r[1] + r[3]);
        const e = document.createElementNS(NS, 'rect');
        e.setAttribute('x', Math.min(x1, x2)); e.setAttribute('y', Math.min(y1, y2)); e.setAttribute('width', Math.abs(x2 - x1)); e.setAttribute('height', Math.abs(y2 - y1));
        e.setAttribute('fill', a.color); e.setAttribute('class', 'hl'); e.dataset.i = i;
        svg.appendChild(e);
      }
    } else if (a.type === 'ink') {
      for (const st of a.strokes) {
        const d = st.d.replace(/(-?[\d.]+) (-?[\d.]+)/g, (m, x, y) => pt(+x, +y).map(n => n.toFixed(1)).join(' '));
        const e = Ink.pathEl({ ...st, d, w: st.w * scale });
        e.dataset.i = i;
        svg.appendChild(e);
      }
    } else if (a.type === 'image') {
      const [x1, y1] = pt(a.x, a.y), [x2, y2] = pt(a.x + a.w, a.y + a.h);
      const e = document.createElementNS(NS, 'image');
      e.setAttribute('href', a.src); e.setAttribute('x', Math.min(x1, x2)); e.setAttribute('y', Math.min(y1, y2)); e.setAttribute('width', Math.abs(x2 - x1)); e.setAttribute('height', Math.abs(y2 - y1));
      e.dataset.i = i;
      svg.appendChild(e);
    } else if (a.type === 'text') {
      const [x, y] = pt(a.x, a.y);
      const t = document.createElement('div');
      t.className = 'pdf-textbox';
      t.textContent = a.text;
      Object.assign(t.style, { left: x + 'px', top: y + 'px', fontSize: a.size * scale + 'px', color: a.color });
      t.dataset.i = i;
      box.appendChild(t);
    } else if (a.type === 'note') {
      const [x, y] = pt(a.x, a.y);
      const n = document.createElement('button');
      n.className = 'pdf-note';
      n.title = a.text;
      n.setAttribute('aria-label', 'Comment: ' + a.text);
      n.innerHTML = '<span class="ico" data-icon="message-square"></span>';
      Object.assign(n.style, { left: x + 'px', top: y + 'px' });
      n.dataset.i = i;
      n.addEventListener('click', async e => {
        if (tool === 'erase') return;
        e.stopPropagation();
        const r = await UI.form('Comment', [{ name: 't', label: 'Comment', type: 'textarea', value: a.text }], 'Save');
        if (r) { snapshot(); a.text = r.t.slice(0, 5000); changed(p); }
      });
      box.appendChild(n);
    }
  });
  if (window.Icons) Icons.fill(box);
}

// ---------- Thumbnails and pages ----------
function thumbFor(p) {
  const t = document.createElement('button');
  t.className = 'pdf-thumb' + (selected.has(p.id) ? ' on' : '');
  t.dataset.id = p.id;
  t.draggable = true;
  t.innerHTML = '<canvas></canvas><span class="n"></span>';
  t.querySelector('.n').textContent = pages.indexOf(p) + 1;
  t.setAttribute('aria-label', `Page ${pages.indexOf(p) + 1}`);
  (async () => {
    const pg = await pdfPage(p), v0 = pg.getViewport({ scale: 1, rotation: (pg.rotate + p.rot) % 360 });
    const v = pg.getViewport({ scale: 120 / v0.width, rotation: (pg.rotate + p.rot) % 360 });
    const c = t.querySelector('canvas');
    c.width = Math.floor(v.width); c.height = Math.floor(v.height);
    await pg.render({ canvas: c, viewport: v }).promise.catch(() => {});
  })();
  t.addEventListener('click', e => {
    if (e.shiftKey && selected.size) {
      const ids = pages.map(x => x.id), a = ids.indexOf([...selected].pop()), b = ids.indexOf(p.id);
      for (let i = Math.min(a, b); i <= Math.max(a, b); i++) selected.add(ids[i]);
    } else if (e.ctrlKey || e.metaKey) { if (selected.has(p.id)) selected.delete(p.id); else selected.add(p.id); }
    else selected = new Set([p.id]);
    thumbs.querySelectorAll('.pdf-thumb').forEach(x => x.classList.toggle('on', selected.has(+x.dataset.id)));
    const box = boxOf(p);
    if (box) box.scrollIntoView({ block: 'start' });
  });
  // Drag a thumbnail to put the page somewhere else.
  t.addEventListener('dragstart', e => { e.dataTransfer.setData('text/x-isadek-page', String(p.id)); e.dataTransfer.effectAllowed = 'move'; });
  t.addEventListener('dragover', e => { if (e.dataTransfer.types.includes('text/x-isadek-page')) { e.preventDefault(); t.classList.add('drop'); } });
  t.addEventListener('dragleave', () => t.classList.remove('drop'));
  t.addEventListener('drop', e => {
    e.preventDefault(); t.classList.remove('drop');
    const id = +e.dataTransfer.getData('text/x-isadek-page');
    const from = pages.findIndex(x => x.id === id), to = pages.indexOf(p);
    if (from < 0 || from === to) return;
    snapshot();
    const [moved] = pages.splice(from, 1);
    pages.splice(to, 0, moved);
    pagesChanged();
  });
  return t;
}
const selectedPages = () => pages.filter(p => selected.has(p.id));
function pagesChanged() { state.markDirty(); layout(); }

$('rotLeft').addEventListener('click', () => { const s = selectedPages(); if (!s.length) return; snapshot(); s.forEach(p => { p.rot = (p.rot + 270) % 360; }); pagesChanged(); });
$('rotRight').addEventListener('click', () => { const s = selectedPages(); if (!s.length) return; snapshot(); s.forEach(p => { p.rot = (p.rot + 90) % 360; }); pagesChanged(); });
$('delPages').addEventListener('click', async () => {
  const s = selectedPages();
  if (!s.length) return;
  if (s.length === pages.length) return UI.toast('A PDF needs at least one page');
  if (!(await UI.confirm('Delete pages', `Delete ${s.length === 1 ? 'page ' + (pages.indexOf(s[0]) + 1) : s.length + ' pages'}?`))) return;
  snapshot();
  pages = pages.filter(p => !selected.has(p.id));
  selected = new Set([pages[0].id]);
  pagesChanged();
});
function move(d) {
  const s = selectedPages();
  if (!s.length) return;
  const idx = s.map(p => pages.indexOf(p));
  if ((d < 0 && Math.min(...idx) === 0) || (d > 0 && Math.max(...idx) === pages.length - 1)) return;
  snapshot();
  const order = d < 0 ? idx : idx.slice().reverse();
  for (const i of order) { const [p] = pages.splice(i, 1); pages.splice(i + d, 0, p); }
  pagesChanged();
}
$('upPage').addEventListener('click', () => move(-1));
$('downPage').addEventListener('click', () => move(1));
$('insertPdf').addEventListener('click', async () => {
  if (!pages.length) return Platform.emit('open');
  const f = await Platform.open([{ name: 'PDF', extensions: ['pdf'] }]);
  if (!f) return;
  let src;
  try { src = await openSource(f.data, f.name); } catch (err) { return UI.toast('Could not open that PDF: ' + err.message, 5000); }
  snapshot();
  sources.push(src);
  const after = Math.max(...selectedPages().map(p => pages.indexOf(p)), pages.length - 1);
  const added = Array.from({ length: src.doc.numPages }, (_, i) => ({ id: ++pageSeq, src: sources.length - 1, index: i, rot: 0, annots: [] }));
  pages.splice(after + 1, 0, ...added);
  UI.toast(`${added.length} page${added.length === 1 ? '' : 's'} added from “${f.name}”`);
  pagesChanged();
});
$('extractPages').addEventListener('click', async () => {
  const s = selectedPages();
  if (!s.length) return UI.toast('Select the pages to extract in the list on the left');
  const bytes = await build(s);
  await Platform.save({ data: bytes, name: Platform.baseName(state.name) + ` (pages ${s.map(p => pages.indexOf(p) + 1).join(',')}).pdf`, filters: [{ name: 'PDF', extensions: ['pdf'] }], saveAs: true });
});
$('splitPdf').addEventListener('click', async () => {
  if (!pages.length) return;
  const r = await UI.form('Split PDF', [{ name: 'n', label: 'Pages in each new PDF', type: 'number', value: 1, min: 1, max: pages.length }], 'Split');
  if (!r) return;
  const n = Math.max(1, Math.min(pages.length, Math.floor(+r.n || 1)));
  const parts = [];
  for (let i = 0; i < pages.length; i += n) parts.push(pages.slice(i, i + n));
  for (const [k, part] of parts.entries()) {
    const bytes = await build(part);
    await Platform.save({ data: bytes, name: `${Platform.baseName(state.name)} part ${k + 1}.pdf`, filters: [{ name: 'PDF', extensions: ['pdf'] }], saveAs: Platform.desktop, noVersion: true });
  }
  UI.toast(`Split into ${parts.length} PDF${parts.length === 1 ? '' : 's'}`);
});

// ---------- Tools on the page ----------
let color = $('annColor').value;
$('annColor').addEventListener('input', e => { color = e.target.value; });
function setTool(t) {
  tool = t;
  document.querySelectorAll('.ribbon .tool').forEach(b => b.classList.toggle('on', b.dataset.tool === t));
  document.body.dataset.pdfTool = t;
}
document.querySelectorAll('.ribbon .tool').forEach(b => b.addEventListener('click', () => setTool(b.dataset.tool)));
setTool('select');

const pageAt = target => { const box = target.closest && target.closest('.pdf-page'); return box && pages.find(p => p.id === +box.dataset.id); };
function snapshot() { undoStack.push(JSON.stringify({ pages, sources: sources.length })); if (undoStack.length > 100) undoStack.shift(); }
function undo() {
  const s = undoStack.pop();
  if (!s) return;
  const d = JSON.parse(s);
  pages = d.pages;
  sources.length = Math.min(sources.length, d.sources);
  pagesChanged();
}
function changed(p) { state.markDirty(); drawAnnots(p); }

// Highlight: the selected text, else a dragged area.
document.addEventListener('mouseup', async e => {
  if (tool !== 'highlight') return;
  const sel = getSelection();
  if (!sel.rangeCount || sel.isCollapsed) return;
  const p = pageAt(sel.anchorNode.parentElement || sel.anchorNode);
  if (!p) return;
  const rects = [];
  for (const r of sel.getRangeAt(0).getClientRects()) {
    if (r.width < 1 || r.height < 1) continue;
    const [x1, y1] = await toPage(p, r.left, r.top), [x2, y2] = await toPage(p, r.right, r.bottom);
    rects.push([Math.min(x1, x2), Math.min(y1, y2), Math.abs(x2 - x1), Math.abs(y2 - y1)]);
  }
  sel.removeAllRanges();
  if (!rects.length) return;
  snapshot();
  p.annots.push({ type: 'highlight', color: '#ffe600', rects });
  changed(p);
});
pagesEl.addEventListener('pointerdown', async e => {
  const p = pageAt(e.target);
  if (!p) return;
  if (tool === 'hand') {
    const sx = e.clientX, sy = e.clientY, st = view.scrollTop, sl = view.scrollLeft;
    const mv = ev => { view.scrollTop = st - (ev.clientY - sy); view.scrollLeft = sl - (ev.clientX - sx); };
    const up = () => { removeEventListener('pointermove', mv); removeEventListener('pointerup', up); };
    addEventListener('pointermove', mv); addEventListener('pointerup', up);
    return;
  }
  if (tool === 'erase') {
    const i = e.target.dataset && e.target.dataset.i;
    if (i === undefined) return;
    e.preventDefault();
    snapshot();
    p.annots.splice(+i, 1);
    changed(p);
    return;
  }
  if (tool === 'highlight' && !e.target.closest('.textLayer span')) {
    // An area: drag a box.
    e.preventDefault();
    const [x0, y0] = await toPage(p, e.clientX, e.clientY);
    const a = { type: 'highlight', color: '#ffe600', rects: [[x0, y0, 0, 0]] };
    snapshot();
    p.annots.push(a);
    const mv = async ev => { const [x, y] = await toPage(p, ev.clientX, ev.clientY); a.rects[0] = [Math.min(x0, x), Math.min(y0, y), Math.abs(x - x0), Math.abs(y - y0)]; drawAnnots(p); };
    const up = () => { removeEventListener('pointermove', mv); removeEventListener('pointerup', up); if (a.rects[0][2] < 2 || a.rects[0][3] < 2) { p.annots.splice(p.annots.indexOf(a), 1); undoStack.pop(); } changed(p); };
    addEventListener('pointermove', mv); addEventListener('pointerup', up);
    return;
  }
  if (tool === 'text' || tool === 'note' || tool === 'date' || tool === 'signature') {
    e.preventDefault();
    const [x, y] = await toPage(p, e.clientX, e.clientY);
    if (tool === 'note') {
      const r = await UI.form('Comment', [{ name: 't', label: 'Comment', type: 'textarea' }], 'Add');
      if (!r || !r.t.trim()) return;
      snapshot(); p.annots.push({ type: 'note', x, y, text: r.t.slice(0, 5000) });
    } else if (tool === 'text' || tool === 'date') {
      let text = tool === 'date' ? new Date().toLocaleDateString() : '';
      if (tool === 'text') { const r = await UI.form('Add Text', [{ name: 't', label: 'Text' }, { name: 's', label: 'Size (points)', type: 'number', value: 12, min: 6, max: 72 }], 'Add'); if (!r || !r.t.trim()) return; text = r.t; pending.size = Math.max(6, Math.min(72, +r.s || 12)); }
      snapshot(); p.annots.push({ type: 'text', x, y, text: text.slice(0, 2000), size: pending.size || 12, color: tool === 'date' ? '#1f1f1f' : color });
      if (tool === 'date') setTool('select');
    } else if (tool === 'signature' && pending.signature) {
      const s = pending.signature, w = 160, h = w * s.h / s.w;
      snapshot(); p.annots.push({ type: 'image', x, y: y - h / 2, w, h, src: s.src });
      setTool('select');
    }
    pending.size = 0;
    changed(p);
  }
});
const pending = { size: 0, signature: null };
// Drawing: strokes in page coordinates, with the pen's pressure.
Ink.attach(pagesEl, {
  tool: () => (tool === 'pen' ? { c: color, w: 2, o: 1 } : null),
  // On-screen pixels within the page while drawing; turned into page coordinates when done.
  at: e => { if (e.type === 'pointerdown' || !inkPage) inkPage = pageAt(document.elementFromPoint(e.clientX, e.clientY) || e.target); const p = inkPage; if (!p) return [0, 0]; const r = boxOf(p).getBoundingClientRect(); return [e.clientX - r.left, e.clientY - r.top]; },
  strokes: () => [],
  liveLayer: () => (inkPage && boxOf(inkPage).querySelector('.pdf-annots')) || pagesEl,
  add: async st => {
    const p = inkPage;
    if (!p) return;
    // Screen-scaled stroke → page coordinates.
    const box = boxOf(p), r = box.getBoundingClientRect();
    const pts = [];
    for (const m of st.d.matchAll(/(-?[\d.]+) (-?[\d.]+)/g)) pts.push(await toPage(p, r.left + +m[1], r.top + +m[2]));
    let i = 0;
    const d = st.d.replace(/(-?[\d.]+) (-?[\d.]+)/g, () => { const q = pts[i++]; return q[0].toFixed(1) + ' ' + q[1].toFixed(1); });
    const stroke = { ...st, d, w: st.w / scale };
    snapshot();
    const last = p.annots[p.annots.length - 1];
    if (last && last.type === 'ink' && Date.now() - (last.at || 0) < 3000) { last.strokes.push(stroke); last.at = Date.now(); }
    else p.annots.push({ type: 'ink', strokes: [stroke], at: Date.now() });
    inkPage = null;
    changed(p);
  },
  erase: () => {}
});
let inkPage = null;

// Sign: draw a signature once (kept on this device), then click where it goes.
$('signBtn').addEventListener('click', async () => {
  let sig = null;
  try { sig = JSON.parse(localStorage.getItem('atlas-signature') || 'null'); } catch { /* none */ }
  if (sig) {
    const i = await UI.choose('Sign', 'Use your saved signature, or draw a new one?', ['Use saved signature', 'Draw a new one', 'Cancel'], 'signature');
    if (i === 2 || i < 0) return;
    if (i === 1) sig = null;
  }
  if (!sig) sig = await drawSignature();
  if (!sig) return;
  pending.signature = sig;
  setTool('signature');
  UI.toast('Click where the signature goes');
});
$('dateBtn').addEventListener('click', () => { setTool('date'); UI.toast('Click where the date goes'); });
$('formInfo').addEventListener('click', () => UI.toast(sources.length ? 'Click a form field on the page to fill it in. The values are saved with the PDF.' : 'Open a PDF form first', 5000));
async function drawSignature() {
  const back = document.createElement('div');
  back.className = 'modal-backdrop';
  back.innerHTML = `<div class="modal sig-dialog" role="dialog" aria-label="Signature"><h2>Draw your signature</h2>
    <div class="sig-pad"><svg viewBox="0 0 600 200" width="600" height="200"></svg></div>
    <label class="row"><input type="checkbox" class="keep" checked> Keep it on this device for next time</label>
    <div class="actions"><button class="clear">Clear</button><span class="sp"></span><button class="cancel">Cancel</button><button class="primary ok">Use</button></div></div>`;
  const svg = back.querySelector('svg'), strokes = [];
  Ink.attach(svg, {
    tool: () => ({ c: '#1a237e', w: 3, o: 1 }),
    at: e => { const r = svg.getBoundingClientRect(); return [(e.clientX - r.left) * 600 / r.width, (e.clientY - r.top) * 200 / r.height]; },
    strokes: () => strokes,
    add: st => { strokes.push(st); svg.appendChild(Ink.pathEl(st)); },
    erase: () => {}
  });
  document.body.appendChild(back);
  return new Promise(resolve => {
    const close = v => { back.remove(); resolve(v); };
    back.querySelector('.clear').onclick = () => { strokes.length = 0; svg.innerHTML = ''; };
    back.querySelector('.cancel').onclick = () => close(null);
    back.addEventListener('keydown', e => { if (e.key === 'Escape') close(null); });
    back.querySelector('.ok').onclick = () => {
      if (!strokes.length) return close(null);
      const b = Ink.bounds(strokes), pad = 6;
      const w = Math.ceil(b.w + 2 * pad), h = Math.ceil(b.h + 2 * pad);
      const c = document.createElement('canvas');
      c.width = w * 3; c.height = h * 3;
      const g = c.getContext('2d');
      g.scale(3, 3);
      Ink.paint(g, Ink.moved(strokes, pad - b.x, pad - b.y));
      const sig = { src: c.toDataURL('image/png'), w, h };
      if (back.querySelector('.keep').checked) { try { localStorage.setItem('atlas-signature', JSON.stringify(sig)); } catch { /* too big or blocked */ } }
      close(sig);
    };
  });
}

// ---------- Saving: a standard PDF with everything in it ----------
const hex = c => { const m = /^#?([0-9a-f]{6})$/i.exec(c || ''); const n = m ? parseInt(m[1], 16) : 0; return rgb(((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255); };
/** Form values filled in on the page, by field name. */
async function formValues(src) {
  const out = new Map();
  const storage = src.doc.annotationStorage;
  if (!storage.size) return out;
  for (let i = 1; i <= src.doc.numPages; i++) {
    const annots = await (await src.doc.getPage(i)).getAnnotations();
    for (const a of annots) {
      if (!a.fieldName) continue;
      const v = storage.getRawValue(a.id);
      // Radio buttons share a name: only the one chosen says which.
      if (a.radioButton && !(v && v.value === true)) continue;
      if (v && v.value !== undefined) out.set(a.fieldName, { value: v.value, type: a.fieldType, check: a.checkBox, radio: a.radioButton, export: a.buttonValue || a.exportValue });
    }
  }
  return out;
}
function fill(form, values) {
  for (const [name, v] of values) {
    let field;
    try { field = form.getField(name); } catch { continue; }
    try {
      if (field instanceof PDFTextField) field.setText(String(v.value));
      else if (field instanceof PDFCheckBox) { if (v.value === true || v.value === 'On' || v.value === v.export) field.check(); else field.uncheck(); }
      else if (field instanceof PDFRadioGroup) { if (typeof v.value === 'string' && v.value !== 'Off') field.select(v.value); else if (v.value === true && v.export) field.select(v.export); }
      else if (field instanceof PDFDropdown || field instanceof PDFOptionList) field.select(Array.isArray(v.value) ? v.value : String(v.value));
    } catch { /* a value the field doesn't take */ }
  }
}
// Helvetica covers Western European text; other characters become "?" rather than failing.
const winAnsi = s => String(s).replace(/[^\x20-\x7e\xa0-\xff–—‘’“”•…€\n]/g, '?');

/** The PDF of the given pages (all, by default), with every change made here. */
async function build(list = pages) {
  const base = await PDFDocument.load(sources[0].bytes, { ignoreEncryption: false, updateMetadata: false });
  // Form values into each source before its pages are used.
  const libDocs = [base];
  for (let s = 1; s < sources.length; s++) libDocs.push(await PDFDocument.load(sources[s].bytes, { updateMetadata: false }));
  for (let s = 0; s < sources.length; s++) {
    const values = await formValues(sources[s]);
    if (values.size) { try { fill(libDocs[s].getForm(), values); } catch { /* no form */ } }
  }
  const originals = base.getPages();
  const out = [];
  for (const p of list) {
    if (p.src === 0) out.push(originals[p.index]);
    else { const [copy] = await base.copyPages(libDocs[p.src], [p.index]); out.push(copy); }
  }
  // The pages in their new order (the base document's pages are moved, so its form stays whole).
  while (base.getPageCount()) base.removePage(0);
  out.forEach(pg => base.addPage(pg));
  const font = await base.embedFont(StandardFonts.Helvetica);
  for (const [k, p] of list.entries()) {
    const pg = out[k];
    if (p.rot) pg.setRotation(degrees((pg.getRotation().angle + p.rot) % 360));
    // The visible area (crop box) is what pdf.js shows, so annotations are placed from it.
    const crop = pg.getCropBox(), height = crop.height, ox = crop.x, oy = crop.y;
    const X = x => x + ox, Y = y => height - y + oy;
    for (const a of p.annots) {
      if (a.type === 'highlight') for (const r of a.rects) pg.drawRectangle({ x: X(r[0]), y: Y(r[1] + r[3]), width: r[2], height: r[3], color: hex(a.color), opacity: 0.4, blendMode: BlendMode.Multiply });
      else if (a.type === 'ink') for (const st of a.strokes) {
        if (st.f) pg.drawSvgPath(st.d, { x: ox, y: height + oy, color: hex(st.c), opacity: st.o });
        else pg.drawSvgPath(st.d, { x: ox, y: height + oy, borderColor: hex(st.c), borderWidth: st.w, borderOpacity: st.o, borderLineCap: 1 });
      } else if (a.type === 'text') {
        const lines = winAnsi(a.text).split('\n');
        lines.forEach((line, li) => pg.drawText(line, { x: X(a.x), y: Y(a.y + a.size * (li + 0.85)), size: a.size, font, color: hex(a.color) }));
      } else if (a.type === 'image') {
        const img = await base.embedPng(a.src);
        pg.drawImage(img, { x: X(a.x), y: Y(a.y + a.h), width: a.w, height: a.h });
      } else if (a.type === 'note') {
        // A real PDF comment (a sticky note) that other readers show and list.
        const ctx = base.context;
        const annot = ctx.obj({ Type: 'Annot', Subtype: 'Text', Rect: [X(a.x), Y(a.y) - 20, X(a.x) + 20, Y(a.y)], Contents: PDFString.of(a.text), Name: 'Comment', C: [1, 0.85, 0.2], F: 4, T: PDFString.of(localStorage.getItem('atlas-author') || 'Isadek PDF') });
        const ref = ctx.register(annot);
        let annots = pg.node.lookup(PDFName.of('Annots'));
        if (!(annots instanceof PDFArray)) { annots = ctx.obj([]); pg.node.set(PDFName.of('Annots'), annots); }
        annots.push(ref);
      }
    }
  }
  base.setProducer('Isadek PDF');
  base.setModificationDate(new Date());
  return base.save({ useObjectStreams: true });
}

async function save(saveAs) {
  if (!pages.length) return false;
  if (sources.some(s => s.encrypted)) { UI.toast('Changes to a password-protected PDF can’t be saved yet. Use Print to save a copy.', 6000); return false; }
  let bytes;
  try { bytes = await build(); } catch (err) { UI.toast('Could not save the PDF: ' + err.message, 6000); return false; }
  let name = state.name;
  if (!/\.pdf$/i.test(name)) name = Platform.baseName(name) + '.pdf';
  try {
    const r = await Platform.save({ data: bytes, name, path: saveAs ? null : state.path, filters: [{ name: 'PDF', extensions: ['pdf'] }], saveAs: saveAs || !state.path });
    if (!r) return false;
    state.set({ name: r.name || name, path: r.path || state.path, dirty: false });
    UI.toast('Saved');
    return true;
  } catch (err) { UI.toast('Save failed: ' + err.message, 5000); return false; }
}

// ---------- Text: find, and Edit in Write ----------
const findbar = $('findbar');
let hits = [], hitAt = -1;
async function find(q) {
  hits = []; hitAt = -1;
  pagesEl.querySelectorAll('.find-hit').forEach(n => n.classList.remove('find-hit', 'cur'));
  if (!q) { $('findCount').textContent = ''; return; }
  const needle = q.toLowerCase();
  for (const p of pages) {
    const tc = await (await pdfPage(p)).getTextContent();
    tc.items.forEach((it, i) => { if (it.str && it.str.toLowerCase().includes(needle)) hits.push({ p, i }); });
  }
  $('findCount').textContent = hits.length ? `${hits.length} found` : 'Not found';
  if (hits.length) go(1);
}
async function go(d) {
  if (!hits.length) return;
  hitAt = (hitAt + d + hits.length) % hits.length;
  const h = hits[hitAt], box = boxOf(h.p);
  box.scrollIntoView({ block: 'center' });
  await renderPage(box);
  setTimeout(() => {
    pagesEl.querySelectorAll('.find-hit').forEach(n => n.classList.remove('find-hit', 'cur'));
    const span = box.querySelectorAll('.textLayer > span, .textLayer span[role="presentation"]')[h.i];
    if (span) { span.classList.add('find-hit', 'cur'); span.scrollIntoView({ block: 'center' }); }
    $('findCount').textContent = `${hitAt + 1} of ${hits.length}`;
  }, 150);
}
$('findText').addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); if (hits.length && $('findText').value === lastQ) go(e.shiftKey ? -1 : 1); else { lastQ = $('findText').value; find(lastQ); } } if (e.key === 'Escape') findbar.hidden = true; });
let lastQ = '';
$('findNext').addEventListener('click', () => go(1));
$('findPrev').addEventListener('click', () => go(-1));
$('findClose').addEventListener('click', () => { findbar.hidden = true; });

/** The text of the PDF as paragraphs and headings, for Isadek Write. */
async function toHtml() {
  const esc = s => String(s).replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
  const out = [];
  for (const p of pages) {
    const tc = await (await pdfPage(p)).getTextContent();
    // Lines: items sharing a baseline; paragraphs: lines with a normal gap between them.
    const lines = [];
    for (const it of tc.items) {
      if (!it.str) continue;
      const y = Math.round(it.transform[5]), size = Math.hypot(it.transform[2], it.transform[3]);
      const last = lines[lines.length - 1];
      if (last && Math.abs(last.y - y) < size * 0.4) { last.text += (/\s$/.test(last.text) || /^\s/.test(it.str) ? '' : ' ') + it.str; last.size = Math.max(last.size, size); }
      else lines.push({ y, size, text: it.str });
    }
    // The body text's size: the one most of the text is in.
    const bySize = new Map();
    for (const l of lines) bySize.set(Math.round(l.size), (bySize.get(Math.round(l.size)) || 0) + l.text.length);
    const body = [...bySize.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] || 12;
    let para = null;
    const flush = () => { if (para) { out.push(para.head ? `<h2>${esc(para.text.trim())}</h2>` : `<p>${esc(para.text.trim())}</p>`); para = null; } };
    lines.forEach((l, i) => {
      const head = l.size > body * 1.25;
      const gap = i ? Math.abs(lines[i - 1].y - l.y) : 0;
      if (!para || head !== para.head || gap > body * 1.8) { flush(); para = { head, text: l.text }; }
      else para.text += (para.text.endsWith('-') ? '' : ' ') + l.text;
    });
    flush();
  }
  return `<!doctype html><meta charset="utf-8"><title>${esc(Platform.baseName(state.name))}</title>` + out.join('\n');
}
$('toWrite').addEventListener('click', async () => {
  if (!pages.length) return UI.toast('Open a PDF first');
  const html = await toHtml();
  if (!/<(p|h2)>/.test(html)) return UI.toast('This PDF has no text to edit (it may be a scan)', 5000);
  await Platform.openData('write', Platform.baseName(state.name) + '.html', new TextEncoder().encode(html));
});

// ---------- Zoom and status ----------
function updateStatus() {
  const boxes = [...pagesEl.querySelectorAll('.pdf-page')];
  const top = view.getBoundingClientRect().top + 40;
  const i = Math.max(0, boxes.findIndex(b => b.getBoundingClientRect().bottom > top));
  $('stPage').textContent = pages.length ? `Page ${i + 1} of ${pages.length}` : '';
}
view.addEventListener('scroll', () => updateStatus());
const ZOOMS = [0.25, 0.5, 0.75, 1, 1.25, 1.5, 2, 3, 4];
function zoomTo(v) { zoomMode = String(v); const sel = $('zoomSel'); sel.value = [...sel.options].some(o => o.value === zoomMode) ? zoomMode : sel.value; layout(); }
$('zoomSel').addEventListener('change', e => zoomTo(e.target.value));
$('zoomIn').addEventListener('click', () => zoomTo(ZOOMS.find(z => z > scale + 0.01) || 4));
$('zoomOut').addEventListener('click', () => zoomTo([...ZOOMS].reverse().find(z => z < scale - 0.01) || 0.25));
view.addEventListener('wheel', e => { if (!(e.ctrlKey || e.metaKey)) return; e.preventDefault(); zoomTo(e.deltaY < 0 ? ZOOMS.find(z => z > scale + 0.01) || 4 : [...ZOOMS].reverse().find(z => z < scale - 0.01) || 0.25); }, { passive: false });
let resizeT = null;
addEventListener('resize', () => { clearTimeout(resizeT); resizeT = setTimeout(() => { if (zoomMode === 'fit' || zoomMode === 'page') layout(); }, 200); });

// ---------- Commands ----------
const OPEN = [{ name: 'PDF', extensions: ['pdf'] }];
Platform.onMenu(async cmd => {
  switch (cmd) {
    case 'open': {
      const f = await Platform.open(OPEN);
      if (!f) break;
      if (Platform.shouldOpenInPlace(!pages.length || !state.dirty)) { if (state.dirty && !(await UI.confirm('Open file', 'Discard the changes to this PDF?'))) break; loadFile(f); }
      else Platform.openPath(f.path);
      break;
    }
    case 'save': save(false); break;
    case 'save-as': save(true); break;
    case 'save-and-close': if (await save(false)) Platform.closeNow(); break;
    case 'version-history': Platform.versionHistory({ state, app: 'pdf', load: loadFile }); break;
    case 'open-cloud': Cloud.openDialog({ app: 'pdf', load: loadFile, isEmpty: async () => !pages.length || !state.dirty }); break;
    case 'save-cloud': Cloud.saveAsDialog({ state, serialize: () => build(), exts: ['pdf'] }); break;
    case 'share': Cloud.shareDialog({ state, serialize: () => build(), exts: ['pdf'], save: () => save(false) }); break;
    case 'print': Platform.print(); break;
    case 'undo': undo(); break;
    case 'find': findbar.hidden = false; $('findText').focus(); $('findText').select(); break;
    case 'home': Platform.desktop ? Platform.newWindow('home') : (location.href = 'index.html'); break;
  }
});
document.addEventListener('keydown', e => {
  if (e.target.closest && e.target.closest('input, textarea, select, .annotationLayer')) return;
  if (e.key === 'Escape' && tool !== 'select') setTool('select');
  if ((e.ctrlKey || e.metaKey) && !e.shiftKey && e.key.toLowerCase() === 'z' && !Platform.desktop) { e.preventDefault(); undo(); }
});

// Thumbnails for Home, and the test hooks.
Platform.preview = () => null;
window.AtlasPdf = {
  loadFile, build, toHtml,
  get pages() { return pages; },
  get sources() { return sources; },
  get scale() { return scale; },
  setTool, zoomTo,
  /** Adds an annotation to page n (1-based) — for tests. */
  annotate(n, a) { snapshot(); pages[n - 1].annots.push(a); changed(pages[n - 1]); },
  select(ids) { selected = new Set(ids.map(i => pages[i - 1].id)); thumbs.querySelectorAll('.pdf-thumb').forEach(x => x.classList.toggle('on', selected.has(+x.dataset.id))); }
};

(async () => {
  const f = await Platform.pending();
  if (f && f.error) UI.toast('Could not open file: ' + f.error, 4000);
  else if (f && !f.recovered) await loadFile(f);
  window.__pdfReady = true;
})();
