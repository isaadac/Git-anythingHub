'use strict';
// Isadek Write — word processor.
(function () {
  const doc = document.getElementById('doc');
  const page = document.getElementById('page');
  const header = document.getElementById('pageHeader');
  const footer = document.getElementById('pageFooter');
  const firstHeader = document.getElementById('firstHeader');
  const firstFooter = document.getElementById('firstFooter');
  const HF = [header, footer, firstHeader, firstFooter];
  const state = UI.docState('Isadek Write', 'Document1');
  const FILTERS = [
    { name: 'Word Document', extensions: ['docx'] },
    { name: 'Web Page', extensions: ['html', 'htm'] },
    { name: 'Markdown', extensions: ['md'] },
    { name: 'Plain Text', extensions: ['txt'] }
  ];
  const OPEN_FILTERS = [{ name: 'Documents', extensions: ['docx', 'html', 'htm', 'md', 'txt'] }, ...FILTERS];
  const PAGE_SIZES = { letter: [8.5, 11], a4: [8.27, 11.69], legal: [8.5, 14] }; // inches
  const DEFAULT_SETUP = { margin: '1in', orientation: 'portrait', size: 'letter', columns: 1, color: '#ffffff', differentFirst: false, trackChanges: false, citeStyle: 'apa', lang: '', styles: {} };
  const setup = { ...DEFAULT_SETUP };
  /** Page setup read from a file is untrusted: keep only values the ribbon itself can produce. */
  function validSetup(v) {
    const out = { ...DEFAULT_SETUP };
    if (!v || typeof v !== 'object') return out;
    if (['1in', '0.5in', '0.75in', '1.5in'].includes(v.margin)) out.margin = v.margin;
    if (['portrait', 'landscape'].includes(v.orientation)) out.orientation = v.orientation;
    if (Object.prototype.hasOwnProperty.call(PAGE_SIZES, v.size)) out.size = v.size;
    if ([1, 2, 3].includes(+v.columns)) out.columns = +v.columns;
    if (Platform.isColor(v.color)) out.color = v.color;
    out.differentFirst = v.differentFirst === true;
    out.trackChanges = v.trackChanges === true;
    if (['apa', 'mla', 'chicago', 'harvard', 'ieee'].includes(v.citeStyle)) out.citeStyle = v.citeStyle;
    out.styles = WriteExtras.cleanStyles(v.styles);
    return out;
  }

  UI.initRibbon();
  document.execCommand('styleWithCSS', false, true);
  document.execCommand('defaultParagraphSeparator', false, 'p');

  // ---------- Selection helpers ----------
  let savedRange = null;
  let extras = null; // section breaks, wrapping, equations, fields, styles, compare (write-extras.js)
  const editable = () => (HF.includes(document.activeElement) ? document.activeElement : doc);
  /** Remembers the caret in the document now (selectionchange arrives a moment after typing). */
  function captureSelection() {
    const sel = getSelection();
    if (!sel.rangeCount || !(doc.contains(sel.anchorNode) || HF.some(h => h.contains(sel.anchorNode)))) return false;
    savedRange = sel.getRangeAt(0).cloneRange();
    return true;
  }
  document.addEventListener('selectionchange', () => {
    if (captureSelection()) {
      updateToolbarState();
      updateCaretPage();
    }
  });
  function restoreSelection() {
    // A caret in the document now is newer than the remembered one.
    captureSelection();
    if (!savedRange) { doc.focus(); return; }
    const root = savedRange.startContainer.nodeType === 1 ? savedRange.startContainer : savedRange.startContainer.parentElement;
    (root.closest('[contenteditable=true]') || doc).focus();
    const sel = getSelection();
    sel.removeAllRanges();
    sel.addRange(savedRange);
  }
  function exec(cmd, val = null) {
    restoreSelection();
    document.execCommand(cmd, false, val);
    changed();
  }
  function currentBlock() {
    const sel = getSelection();
    if (!sel.rangeCount) return null;
    let n = sel.anchorNode;
    if (n && n.nodeType === 3) n = n.parentElement;
    while (n && n !== doc && !/^(P|H[1-6]|LI|BLOCKQUOTE|PRE|DIV|TD|TH)$/.test(n.tagName)) n = n.parentElement;
    return n && n !== doc ? n : null;
  }
  const closest = sel => {
    const s = getSelection();
    if (!s.rangeCount) return null;
    let n = s.anchorNode;
    if (n && n.nodeType === 3) n = n.parentElement;
    return n && doc.contains(n) ? n.closest(sel) : null;
  };

  // ---------- Ribbon: Home ----------
  document.querySelectorAll('[data-exec]').forEach(b => b.addEventListener('click', async () => {
    const cmd = b.dataset.exec;
    if (cmd === 'paste') {
      restoreSelection();
      try {
        const text = await navigator.clipboard.readText();
        document.execCommand('insertText', false, text);
        changed();
      } catch { UI.toast('Use Ctrl+V to paste'); }
      return;
    }
    exec(cmd);
  }));

  document.getElementById('fontName').addEventListener('change', e => exec('fontName', e.target.value));
  function applyFontSize(pt) {
    restoreSelection();
    // execCommand only knows sizes 1–7, so mark with 7 and then swap in the real point size.
    document.execCommand('styleWithCSS', false, false);
    document.execCommand('fontSize', false, '7');
    document.execCommand('styleWithCSS', false, true);
    editable().querySelectorAll('font[size="7"]').forEach(f => {
      const span = document.createElement('span');
      span.style.fontSize = pt + 'pt';
      span.innerHTML = f.innerHTML;
      f.replaceWith(span);
    });
    editable().querySelectorAll('span[style*="xxx-large"]').forEach(s => { s.style.fontSize = pt + 'pt'; });
    document.getElementById('fontSize').value = String(pt);
    changed();
  }
  document.getElementById('fontSize').addEventListener('change', e => applyFontSize(+e.target.value));
  const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72];
  document.getElementById('growFont').addEventListener('click', () => {
    const cur = +document.getElementById('fontSize').value;
    applyFontSize(SIZES.find(s => s > cur) || 72);
  });
  document.getElementById('shrinkFont').addEventListener('click', () => {
    const cur = +document.getElementById('fontSize').value;
    applyFontSize([...SIZES].reverse().find(s => s < cur) || 8);
  });
  function colorInput(id, cmd) {
    const input = document.getElementById(id);
    const bar = input.parentElement.querySelector('.bar');
    input.addEventListener('input', () => { bar.style.background = input.value; });
    input.addEventListener('change', () => exec(cmd, input.value));
    // Clicking the swatch area re-applies the last colour, like Office's split button.
    input.parentElement.addEventListener('contextmenu', e => { e.preventDefault(); exec(cmd, input.value); });
  }
  colorInput('foreColor', 'foreColor');
  colorInput('hiliteColor', 'hiliteColor');
  document.getElementById('shading').addEventListener('change', e => {
    restoreSelection();
    const b = currentBlock();
    if (b) { b.style.background = e.target.value; changed(); }
  });
  document.getElementById('changeCase').addEventListener('click', () => {
    restoreSelection();
    const text = getSelection().toString();
    if (!text) return;
    const next = text === text.toUpperCase() ? text.toLowerCase()
      : text === text.toLowerCase() ? text.replace(/\b\w/g, c => c.toUpperCase()) : text.toUpperCase();
    document.execCommand('insertText', false, next);
    changed();
  });
  document.getElementById('lineSpacing').addEventListener('change', e => {
    restoreSelection();
    const blocks = selectedBlocks();
    blocks.forEach(b => { b.style.lineHeight = e.target.value; });
    if (!blocks.length) doc.style.lineHeight = e.target.value;
    changed();
  });
  function selectedBlocks() {
    const sel = getSelection();
    if (!sel.rangeCount) return [];
    const range = sel.getRangeAt(0);
    const all = [...doc.querySelectorAll('p,h1,h2,h3,h4,h5,h6,li,blockquote,pre')];
    const hit = all.filter(b => range.intersectsNode(b));
    return hit.length ? hit : [currentBlock()].filter(Boolean);
  }
  document.getElementById('showMarks').addEventListener('click', e => {
    document.body.classList.toggle('show-marks');
    e.currentTarget.classList.toggle('on');
  });
  document.querySelectorAll('.style-btn').forEach(b => b.addEventListener('click', () => {
    exec('formatBlock', b.dataset.block);
    const blk = currentBlock();
    if (blk) blk.className = b.dataset.cls || '';
    refreshNav();
  }));

  // Format painter: copy the inline style at the cursor, apply to the next selection.
  let painter = null;
  document.getElementById('formatPainter').addEventListener('click', e => {
    restoreSelection();
    const n = closest('*');
    if (!n) return;
    const cs = getComputedStyle(n);
    painter = { bold: +cs.fontWeight >= 600, italic: cs.fontStyle === 'italic', underline: cs.textDecorationLine.includes('underline'),
      color: cs.color, font: cs.fontFamily.split(',')[0].replace(/"/g, ''), size: parseFloat(cs.fontSize) * 0.75 };
    e.currentTarget.classList.add('on');
    UI.toast('Select text to apply the formatting');
  });
  doc.addEventListener('mouseup', () => {
    if (!painter || getSelection().isCollapsed) return;
    const p = painter;
    painter = null;
    document.getElementById('formatPainter').classList.remove('on');
    document.execCommand('removeFormat');
    if (p.bold) document.execCommand('bold');
    if (p.italic) document.execCommand('italic');
    if (p.underline) document.execCommand('underline');
    document.execCommand('foreColor', false, p.color);
    document.execCommand('fontName', false, p.font);
    savedRange = getSelection().getRangeAt(0).cloneRange();
    applyFontSize(Math.round(p.size));
  });

  function updateToolbarState() {
    document.querySelectorAll('[data-state]').forEach(b => {
      try { b.classList.toggle('on', document.queryCommandState(b.dataset.state)); } catch { /* ignore */ }
    });
    const n = closest('*');
    if (n) {
      const cs = getComputedStyle(n);
      const font = cs.fontFamily.split(',')[0].replace(/["']/g, '').trim();
      const fsel = document.getElementById('fontName');
      if ([...fsel.options].some(o => o.value === font)) fsel.value = font;
      const pt = Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2;
      const ssel = document.getElementById('fontSize');
      if (![...ssel.options].some(o => o.value === String(pt))) {
        const o = document.createElement('option'); o.textContent = String(pt); ssel.appendChild(o);
      }
      ssel.value = String(pt);
    }
  }

  // ---------- Ribbon: Insert ----------
  document.getElementById('insPageBreak').addEventListener('click', () => {
    exec('insertHTML', '<hr class="page-break"><p><br></p>');
  });
  document.getElementById('insSectionBreak').addEventListener('click', () => extras.insertSectionBreak());
  document.getElementById('insEquation').addEventListener('click', () => extras.editEquation(null));
  document.getElementById('insField').addEventListener('click', () => {
    if (savedRange && !doc.contains(savedRange.startContainer)) return UI.toast('Fields go in the body of the document (headers and footers have Page Number)', 4000);
    extras.insertField();
  });
  document.getElementById('modifyStyle').addEventListener('click', () => { restoreSelection(); extras.modifyStyle(currentBlock()); });
  document.getElementById('compareDocs').addEventListener('click', () => compareWith());
  /** Compare: an earlier version (chosen file) against this document, side by side or as tracked changes in a new window. */
  async function compareWith() {
    const f = await Platform.open(OPEN_FILTERS);
    if (!f) return;
    let other;
    try { other = await readDocument(f); } catch (err) { return UI.toast('Could not open that file: ' + err.message, 4000); }
    const original = document.createElement('div');
    original.innerHTML = sanitize(other.html);
    const revised = document.createElement('div');
    revised.innerHTML = cleanHtmlFor(doc);
    original.querySelectorAll('hr.page-break').forEach(n => { const d = document.createElement('div'); d.style.pageBreakAfter = 'always'; n.replaceWith(d); });
    // Drawn equations and field results aren't text to compare; they are redrawn in the result.
    [original, revised].forEach(r => r.querySelectorAll('span.equation, span.fld').forEach(x => { x.textContent = ''; }));
    const n = extras.compare(original, revised, 'Compare');
    if (!n) return UI.toast('No differences in the text');
    // Side by side first; from there, the same changes can open as tracked changes.
    extras.compareView(revised, {
      leftName: f.name, rightName: state.name,
      onTracked: async () => {
        const html = toHtml(() => revised.innerHTML);
        await Platform.openData('write', `Compared — ${Platform.baseName(state.name)}.html`, new TextEncoder().encode(html));
        UI.toast(`${n} change${n === 1 ? '' : 's'} marked. Accept or reject them on the Review tab.`, 5000);
      }
    });
  }
  document.getElementById('insTable').addEventListener('click', async () => {
    const r = await UI.form('Insert Table', [
      { name: 'rows', label: 'Number of rows', type: 'number', value: 3, min: 1, max: 100 },
      { name: 'cols', label: 'Number of columns', type: 'number', value: 3, min: 1, max: 20 },
      { name: 'header', label: 'Header row', type: 'select', options: ['Yes', 'No'] }
    ], 'Insert');
    if (!r) return;
    const rows = Math.max(1, Math.min(100, +r.rows || 3));
    const cols = Math.max(1, Math.min(20, +r.cols || 3));
    let html = '<table><tbody>';
    for (let i = 0; i < rows; i++) {
      html += '<tr>';
      for (let j = 0; j < cols; j++) html += (i === 0 && r.header === 'Yes') ? '<th><br></th>' : '<td><br></td>';
      html += '</tr>';
    }
    exec('insertHTML', html + '</tbody></table><p><br></p>');
  });
  function tableOp(fn) {
    restoreSelection();
    const cell = closest('td,th');
    if (!cell) return UI.toast('Place the cursor in a table first');
    fn(cell, cell.parentElement, cell.closest('table'));
    changed();
  }
  const blankRow = tr => { const n = tr.cloneNode(true); n.querySelectorAll('td,th').forEach(c => { const td = document.createElement('td'); td.innerHTML = '<br>'; c.replaceWith(td); }); return n; };
  document.getElementById('rowAbove').addEventListener('click', () => tableOp((c, tr) => tr.before(blankRow(tr))));
  document.getElementById('rowBelow').addEventListener('click', () => tableOp((c, tr) => tr.after(blankRow(tr))));
  document.getElementById('delRow').addEventListener('click', () => tableOp((c, tr, t) => { tr.remove(); if (!t.querySelector('tr')) t.remove(); }));
  function colOp(dir) {
    tableOp((cell, tr, table) => {
      const idx = [...tr.children].indexOf(cell);
      table.querySelectorAll('tr').forEach(row => {
        const ref = row.children[idx];
        if (!ref) return;
        if (dir === 'del') { ref.remove(); return; }
        const n = document.createElement(ref.tagName.toLowerCase());
        n.innerHTML = '<br>';
        dir === 'left' ? ref.before(n) : ref.after(n);
      });
      if (!table.querySelector('td,th')) table.remove();
    });
  }
  document.getElementById('colLeft').addEventListener('click', () => colOp('left'));
  document.getElementById('colRight').addEventListener('click', () => colOp('right'));
  document.getElementById('delCol').addEventListener('click', () => colOp('del'));

  document.getElementById('insImage').addEventListener('click', async () => {
    const f = await Platform.open([{ name: 'Images', extensions: ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp', 'svg'] }]);
    if (!f) return;
    const ext = Platform.extOf(f.name);
    const mime = ext === 'svg' ? 'image/svg+xml' : ext === 'jpg' ? 'image/jpeg' : 'image/' + ext;
    const url = await Platform.bytesToDataUrl(f.data, mime);
    exec('insertHTML', `<img src="${url}" alt="${escapeHtml(f.name)}" style="width:50%">`);
  });
  document.getElementById('insShape').addEventListener('click', () => exec('insertHTML', '<div class="textbox">Text box</div><p><br></p>'));
  document.getElementById('insLink').addEventListener('click', async () => {
    restoreSelection();
    const text = getSelection().toString();
    const r = await UI.form('Insert Link', [
      { name: 'text', label: 'Text to display', value: text },
      { name: 'url', label: 'Address', value: 'https://' }
    ], 'Insert');
    if (!r || !/^(https?:|mailto:)/i.test(r.url)) return;
    exec('insertHTML', `<a href="${escapeHtml(r.url)}">${escapeHtml(r.text || r.url)}</a>`);
  });
  document.getElementById('insHr').addEventListener('click', () => exec('insertHorizontalRule'));
  document.getElementById('insDate').addEventListener('click', async () => {
    const now = new Date();
    const r = await UI.form('Date and Time', [{ name: 'fmt', label: 'Format', type: 'select', options: [
      now.toLocaleDateString(), now.toLocaleDateString(undefined, { dateStyle: 'long' }),
      now.toLocaleDateString(undefined, { dateStyle: 'full' }), now.toLocaleString(), now.toISOString().slice(0, 10)] }], 'Insert');
    if (r) exec('insertText', r.fmt);
  });
  document.getElementById('insSymbol').addEventListener('click', async () => {
    const symbols = '© ® ™ € £ ¥ ¢ § ¶ † ‡ • … – — ± × ÷ ≠ ≤ ≥ ≈ ∞ √ ∑ π Ω µ ° ½ ¼ ¾ → ← ↑ ↓ ✓ ✗ ★ ☆ ♥ α β γ δ θ λ σ φ'.split(' ');
    const r = await UI.form('Symbol', [{ name: 's', label: 'Symbol', type: 'select', options: symbols }], 'Insert');
    if (r) exec('insertText', r.s);
  });
  document.getElementById('insToc').addEventListener('click', () => toc.insert());
  document.getElementById('updToc').addEventListener('click', () => (doc.querySelector('.toc') ? toc.update() : UI.toast('Insert a table of contents first (Insert → Contents)')));
  function editHF(kind) {
    layout.showHeaderFooter(kind);
    const el = setup.differentFirst ? (kind === 'header' ? firstHeader : firstFooter) : (kind === 'header' ? header : footer);
    el.hidden = false;
    el.focus();
    el.scrollIntoView({ block: 'nearest' });
    if (setup.differentFirst && layout.pages < 2) UI.toast('The header and footer for the other pages appear from page 2', 3500);
  }
  document.getElementById('editHeader').addEventListener('click', () => editHF('header'));
  document.getElementById('editFooter').addEventListener('click', () => editHF('footer'));
  HF.forEach(el => el.addEventListener('blur', () => setTimeout(() => { if (!HF.includes(document.activeElement)) layout.hideEmptyHeaderFooter(); }, 150)));
  // Page numbers: a field that shows the page it is on (or the page count).
  const fieldHtml = kind => `<span class="fld" data-fld="${kind}" contenteditable="false">${kind === 'page' ? '1' : layout.pages}</span>`;
  function fixFields(root) { root.querySelectorAll('.fld').forEach(f => { f.contentEditable = 'false'; if (!WriteExtras.FIELDS[f.dataset.fld]) f.dataset.fld = 'page'; }); }
  document.getElementById('insPageNum').addEventListener('change', e => {
    const v = e.target.value;
    e.target.value = '';
    if (!v) return;
    if (v === 'remove') {
      HF.forEach(el => el.querySelectorAll('.fld').forEach(f => { const p = f.closest('p'); f.remove(); if (p && !p.textContent.replace(/(Page|of)/g, '').trim()) p.remove(); }));
      layout.run(); changed(); return;
    }
    const [where, align, kind] = v.split(':');
    if (where === 'here') { restoreSelection(); Platform.insertInline(fieldHtml(kind)); fixFields(editable()); layout.run(); return; }
    const targets = setup.differentFirst ? [where === 'header' ? firstHeader : firstFooter, where === 'header' ? header : footer] : [where === 'header' ? header : footer];
    const html = kind === 'nofm' ? `Page ${fieldHtml('page')} of ${fieldHtml('numpages')}` : fieldHtml('page');
    targets.forEach(el => {
      el.querySelectorAll('p.pgnum').forEach(p => p.remove());
      el.insertAdjacentHTML(where === 'header' ? 'afterbegin' : 'beforeend', `<p class="pgnum" style="text-align:${align}">${html}</p>`);
      fixFields(el);
    });
    layout.run();
    changed();
  });
  document.getElementById('differentFirst').addEventListener('change', e => {
    setup.differentFirst = e.target.checked;
    // Start the first-page header/footer from the current ones, as Word does.
    if (setup.differentFirst && !firstHeader.textContent.trim() && !firstFooter.textContent.trim()) {
      firstHeader.innerHTML = ''; firstFooter.innerHTML = '';
    }
    layout.run();
    changed();
  });

  // Image selection tools
  const imgTools = document.getElementById('imgTools');
  let selImg = null;
  doc.addEventListener('click', e => {
    doc.querySelectorAll('img.selected').forEach(i => i.classList.remove('selected'));
    if (e.target.tagName === 'IMG') {
      selImg = e.target;
      selImg.classList.add('selected');
      const r = selImg.getBoundingClientRect();
      const w = document.querySelector('.workspace').getBoundingClientRect();
      imgTools.style.left = (r.left - w.left) + 'px';
      imgTools.style.top = Math.max(0, r.top - w.top - 38) + 'px';
      imgTools.hidden = false;
    } else { selImg = null; imgTools.hidden = true; }
    const a = e.target.closest('a');
    if (a && (e.ctrlKey || e.metaKey)) {
      if (a.getAttribute('href').startsWith('#')) document.getElementById(a.getAttribute('href').slice(1))?.scrollIntoView({ behavior: 'smooth' });
      else window.open(a.href);
    }
  });
  imgTools.addEventListener('mousedown', e => e.preventDefault());
  imgTools.addEventListener('click', e => {
    const b = e.target.closest('button');
    if (!b || !selImg) return;
    if (b.dataset.w) selImg.style.width = b.dataset.w + '%';
    if (b.dataset.wrap) extras.setWrap(selImg, b.dataset.wrap);
    if (b.dataset.a) {
      extras.setWrap(selImg, 'inline');
      selImg.style.display = 'block';
      selImg.style.marginLeft = b.dataset.a === 'left' ? '0' : 'auto';
      selImg.style.marginRight = b.dataset.a === 'right' ? '0' : 'auto';
    }
    if (b.dataset.del) { selImg.remove(); selImg = null; imgTools.hidden = true; }
    if (b.dataset.alt) { altTextDialog(selImg); return; }
    changed();
  });
  /** Alt text: a description read out by screen readers (and saved in .docx), or "decorative". */
  async function altTextDialog(img) {
    if (!img) return false;
    const v = await UI.form('Alt Text', [
      { name: 'alt', label: 'How would you describe this picture to someone who can’t see it? (1–2 sentences)', type: 'textarea', value: img.classList.contains('decorative') ? '' : img.getAttribute('alt') || '' },
      { name: 'deco', label: 'Mark as decorative (screen readers skip it)', type: 'select', value: img.classList.contains('decorative') ? 'yes' : 'no', options: [{ value: 'no', label: 'No' }, { value: 'yes', label: 'Yes' }] }
    ], 'OK');
    if (!v) return false;
    if (v.deco === 'yes') { img.alt = ''; img.classList.add('decorative'); } else { img.alt = v.alt.trim().slice(0, 1000); img.classList.remove('decorative'); }
    changed();
    return true;
  }

  // ---------- Ribbon: Layout ----------
  function applySetup() {
    const [w, h] = PAGE_SIZES[setup.size];
    const land = setup.orientation === 'landscape';
    page.style.setProperty('--page-w', (land ? h : w) + 'in');
    page.style.setProperty('--page-h', (land ? w : h) + 'in');
    page.style.setProperty('--margin', setup.margin);
    page.style.setProperty('--paper', setup.color);
    doc.style.columnCount = setup.columns;
    if (extras) extras.applyStyles();
    doc.style.columnGap = '0.5in';
    let style = document.getElementById('print-page-style');
    if (!style) { style = document.createElement('style'); style.id = 'print-page-style'; document.head.appendChild(style); }
    // Margins are part of the laid-out sheets, so the printer page has none of its own.
    style.textContent = `@page { size: ${land ? h : w}in ${land ? w : h}in; margin: 0; }`;
    document.getElementById('differentFirst').checked = !!setup.differentFirst;
    if (review.tracking !== !!setup.trackChanges) setTracking(setup.trackChanges);
    document.getElementById('citeStyle').value = setup.citeStyle;
    // The proofing language: the browser's own spellchecker follows lang too.
    if (setup.lang) { doc.lang = setup.lang; HF.forEach(h => { h.lang = setup.lang; }); } else { doc.removeAttribute('lang'); HF.forEach(h => h.removeAttribute('lang')); }
    document.getElementById('stLang').textContent = Proofing.langName(setup.lang || navigator.language);
    layout.run();
    updateStatus();
  }
  [['margins', 'margin'], ['orientation', 'orientation'], ['pageSize', 'size'], ['columns', 'columns'], ['pageColor', 'color']].forEach(([id, key]) => {
    document.getElementById(id).addEventListener('change', e => { setup[key] = key === 'columns' ? +e.target.value : e.target.value; applySetup(); changed(); });
  });

  // ---------- Ribbon: Review ----------
  document.getElementById('spellToggle').addEventListener('click', e => {
    doc.spellcheck = !doc.spellcheck;
    e.currentTarget.classList.toggle('on', !doc.spellcheck);
    // Re-render so the browser shows/hides squiggles immediately.
    doc.innerHTML = doc.innerHTML;
    savedRange = null;
    UI.toast(doc.spellcheck ? 'Spell check on' : 'Spell check off');
  });
  document.getElementById('wordCount').addEventListener('click', async () => {
    const t = doc.innerText;
    const words = (t.match(/\S+/g) || []).length;
    const paras = doc.querySelectorAll('p,h1,h2,h3,li').length;
    await UI.form('Word Count', [
      { name: 'a', label: 'Pages', value: pageCount() }, { name: 'b', label: 'Words', value: words },
      { name: 'c', label: 'Characters (with spaces)', value: t.replace(/\n/g, '').length },
      { name: 'd', label: 'Characters (no spaces)', value: t.replace(/\s/g, '').length },
      { name: 'e', label: 'Paragraphs', value: paras }
    ], 'Close');
  });
  let speaking = false;
  document.getElementById('readAloud').addEventListener('click', e => {
    if (!('speechSynthesis' in window)) return UI.toast('Speech is not available on this system');
    if (speaking) { speechSynthesis.cancel(); speaking = false; e.currentTarget.classList.remove('on'); return; }
    const text = getSelection().toString() || doc.innerText;
    const u = new SpeechSynthesisUtterance(text);
    u.onend = () => { speaking = false; document.getElementById('readAloud').classList.remove('on'); };
    speechSynthesis.speak(u);
    speaking = true;
    e.currentTarget.classList.add('on');
  });
  // ---------- Review: comments and tracked changes ----------
  const commentsPane = document.getElementById('commentsPane');
  function getAuthor() {
    try { return localStorage.getItem('atlas-author') || 'Author'; } catch { return 'Author'; }
  }
  async function setAuthor() {
    const name = await UI.prompt('Your name', 'Shown on your comments and tracked changes', getAuthor() === 'Author' ? '' : getAuthor());
    if (name === null) return;
    try { localStorage.setItem('atlas-author', name.replace(/[^\w .@'-]/g, '').slice(0, 64) || 'Author'); } catch { /* ignore */ }
  }
  // ---------- Draw: a drawing made with a pen, finger or mouse (ink.js) ----------
  // It goes into the document as a picture (so Word and PDFs show it) that keeps its strokes:
  // double-click it to draw on it again.
  const INK_W = 624, INK_H = 360;
  async function drawingDialog(img) {
    let data = null;
    try { data = img && JSON.parse(img.dataset.ink); } catch { data = null; }
    const strokes = ((data && data.s) || []).map(Ink.clean).filter(Boolean);
    const w = Math.max(INK_W, (data && data.w) || 0), h = Math.max(INK_H, (data && data.h) || 0);
    const back = document.createElement('div');
    back.className = 'modal-backdrop';
    back.innerHTML = `<div class="modal ink-dialog" role="dialog" aria-label="Drawing"><h2>${img ? 'Edit Drawing' : 'Drawing'}</h2>
      <div class="ink-tools" role="toolbar" aria-label="Pens">
        <button class="btn pen on" data-pen="#1f1f1f" data-w="3" title="Black pen"><span class="ico" style="color:#1f1f1f" data-icon="pen-line"></span></button>
        <button class="btn pen" data-pen="#1565c0" data-w="3" title="Blue pen"><span class="ico" style="color:#1565c0" data-icon="pen-line"></span></button>
        <button class="btn pen" data-pen="#c62828" data-w="3" title="Red pen"><span class="ico" style="color:#c62828" data-icon="pen-line"></span></button>
        <button class="btn pen" data-pen="#2e7d32" data-w="3" title="Green pen"><span class="ico" style="color:#2e7d32" data-icon="pen-line"></span></button>
        <button class="btn pen" data-pen="#ffe600" data-w="18" data-hl="1" title="Highlighter"><span class="ico" style="color:#e0c000" data-icon="highlighter"></span></button>
        <button class="btn eraser" title="Eraser"><span class="ico" data-icon="eraser"></span></button>
        <span class="sp"></span>
        <button class="btn undo" title="Undo stroke"><span class="ico" data-icon="undo-2"></span></button>
        <button class="btn clear" title="Clear"><span class="ico" data-icon="trash-2"></span></button>
      </div>
      <div class="ink-paper"><svg class="ink-svg" viewBox="0 0 ${w} ${h}" width="${w}" height="${h}"></svg></div>
      <div class="actions"><button class="cancel">Cancel</button><button class="primary ok">${img ? 'Update' : 'Insert'}</button></div></div>`;
    const q = sel => back.querySelector(sel);
    const svg = q('.ink-svg');
    const redraw = () => { svg.innerHTML = ''; strokes.forEach(st => svg.appendChild(Ink.pathEl(st))); };
    // Strokes of a drawing being edited sit where they were (it was cropped to them, with a margin).
    redraw();
    let tool = { c: '#1f1f1f', w: 3, o: 1 };
    back.querySelectorAll('.ink-tools .pen, .ink-tools .eraser').forEach(b => b.addEventListener('click', () => {
      back.querySelectorAll('.ink-tools .on').forEach(x => x.classList.remove('on'));
      b.classList.add('on');
      tool = b.classList.contains('eraser') ? 'eraser' : { c: b.dataset.pen, w: +b.dataset.w, o: b.dataset.hl ? 0.45 : 1, hl: !!b.dataset.hl };
    }));
    q('.undo').addEventListener('click', () => { strokes.pop(); redraw(); });
    q('.clear').addEventListener('click', () => { strokes.length = 0; redraw(); });
    Ink.attach(svg, {
      tool: () => tool,
      at: e => { const r = svg.getBoundingClientRect(); return [Math.round((e.clientX - r.left) * w / r.width * 10) / 10, Math.round((e.clientY - r.top) * h / r.height * 10) / 10]; },
      strokes: () => strokes,
      add: st => { strokes.push(st); svg.appendChild(Ink.pathEl(st)); },
      erase: i => { strokes.splice(i, 1); redraw(); }
    });
    document.body.appendChild(back);
    if (window.Icons) Icons.fill(back);
    return new Promise(resolve => {
      const close = v => { back.remove(); resolve(v); };
      q('.cancel').addEventListener('click', () => close(null));
      back.addEventListener('keydown', e => { if (e.key === 'Escape') close(null); });
      q('.ok').addEventListener('click', () => {
        if (!strokes.length) { close(null); if (img) appEdit(() => img.remove()); changed(); return; }
        // Cropped to the strokes, with a margin.
        const b = Ink.bounds(strokes), pad = 12;
        const cw = Math.ceil(b.w + 2 * pad), ch = Math.ceil(b.h + 2 * pad);
        const kept = Ink.moved(strokes, pad - b.x, pad - b.y);
        const c = document.createElement('canvas');
        c.width = cw * 2; c.height = ch * 2;
        const g = c.getContext('2d');
        g.scale(2, 2);
        Ink.paint(g, kept);
        const html = `<img class="ink-drawing" alt="Drawing" width="${cw}" height="${ch}" src="${c.toDataURL('image/png')}" data-ink='${JSON.stringify({ w: cw, h: ch, s: kept })}'>`;
        close(true);
        if (img) appEdit(() => { const t = document.createElement('div'); t.innerHTML = html; img.replaceWith(t.firstChild); });
        else { restoreSelection(); document.execCommand('insertHTML', false, html); }
        changed();
      });
    });
  }
  document.getElementById('insDrawing').addEventListener('click', () => { captureSelection(); drawingDialog(null); });
  doc.addEventListener('dblclick', e => { const img = e.target.closest && e.target.closest('img.ink-drawing'); if (img && doc.contains(img)) { e.preventDefault(); drawingDialog(img); } });

  // ---------- Undo for changes the app makes (not typed) ----------
  // The browser's undo knows typing and execCommand only. Changes the app makes to whole
  // paragraphs (the Assistant's rewrites) are kept here as before/after copies: Ctrl+Z takes the
  // latest back while the document is just as it left it; otherwise the browser's undo runs.
  const appEdits = [], appRedos = [];
  const docState = () => { const c = doc.cloneNode(true); layout.clearPushes(c); return c.innerHTML; };
  function appEdit(fn) {
    const before = docState();
    fn();
    const after = docState();
    if (before === after) return false;
    appEdits.push({ before, after });
    if (appEdits.length > 20) appEdits.shift();
    appRedos.length = 0;
    return true;
  }
  function appUndo(redo) {
    const from = redo ? appRedos : appEdits, to = redo ? appEdits : appRedos, top = from[from.length - 1];
    if (!top || docState() !== (redo ? top.before : top.after)) return false;
    from.pop(); to.push(top);
    doc.innerHTML = redo ? top.after : top.before;
    changed();
    return true;
  }
  const review = WriteReview({ doc, author: getAuthor, edit: appEdit, peers: () => (live && live.status === 'live' ? live.peers.map(p => p.name) : []), changed: () => { changed(); if (!commentsPane.hidden) review.renderPane(); }, pane: commentsPane });
  document.getElementById('setAuthor').addEventListener('click', setAuthor);
  document.getElementById('addComment').addEventListener('click', async () => {
    restoreSelection();
    if (getAuthor() === 'Author') await setAuthor();
    restoreSelection();
    review.addComment();
  });
  document.getElementById('showComments').addEventListener('click', () => review.showPane(commentsPane.hidden));
  document.getElementById('closeComments').addEventListener('click', () => review.showPane(false));
  function setTracking(on) {
    review.setTracking(on);
    setup.trackChanges = !!on;
    const b = document.getElementById('trackChanges');
    b.classList.toggle('on', !!on);
    b.setAttribute('aria-pressed', String(!!on));
  }
  document.getElementById('trackChanges').addEventListener('click', async () => {
    if (!review.tracking && getAuthor() === 'Author') await setAuthor();
    setTracking(!review.tracking);
    UI.toast(review.tracking ? 'Track Changes is on: edits are marked for review' : 'Track Changes is off');
    changed();
    doc.focus();
  });
  document.getElementById('markupMode').addEventListener('change', e => {
    document.body.classList.toggle('markup-final', e.target.value === 'final');
    document.body.classList.toggle('markup-original', e.target.value === 'original');
    layout.schedule();
  });
  [['acceptChange', 'acceptAtCursor'], ['rejectChange', 'rejectAtCursor'], ['acceptAll', 'acceptAll'], ['rejectAll', 'rejectAll'], ['nextChange', 'next'], ['prevChange', 'previous']]
    .forEach(([id, fn]) => document.getElementById(id).addEventListener('click', () => { restoreSelection(); review[fn](); }));
  document.getElementById('delComment').addEventListener('click', () => {
    restoreSelection();
    const c = closest('.comment');
    if (!c) return UI.toast('Place the cursor in a comment');
    c.replaceWith(...c.childNodes);
    changed();
    if (!commentsPane.hidden) review.renderPane();
  });

  // ---------- Ribbon: View ----------
  document.getElementById('viewPrint').addEventListener('click', () => { document.body.classList.remove('view-web'); layout.run(); });
  document.getElementById('viewWeb').addEventListener('click', () => { document.body.classList.add('view-web'); layout.run(); });
  document.getElementById('viewFocus').addEventListener('click', () => {
    document.body.classList.add('view-focus');
    UI.toast('Press Esc to leave Focus mode');
  });
  document.addEventListener('keydown', e => { if (e.key === 'Escape') document.body.classList.remove('view-focus'); });
  document.getElementById('viewDark').addEventListener('click', e => { document.body.classList.toggle('dark'); e.currentTarget.classList.toggle('on'); });
  document.getElementById('toggleNav').addEventListener('click', () => {
    const nav = document.getElementById('navpane');
    nav.hidden = !nav.hidden;
    refreshNav();
  });
  function refreshNav() {
    const nav = document.getElementById('navpane');
    if (nav.hidden) return;
    const list = document.getElementById('navlist');
    list.innerHTML = '';
    doc.querySelectorAll('h1,h2,h3').forEach(h => {
      const a = document.createElement('a');
      a.textContent = h.textContent || '(empty heading)';
      a.className = 'l' + h.tagName[1];
      a.onclick = () => h.scrollIntoView({ behavior: 'smooth', block: 'start' });
      list.appendChild(a);
    });
    if (!list.children.length) list.innerHTML = '<p style="color:#605e5c">Use heading styles to build an outline.</p>';
  }

  // Zoom
  const zoom = document.getElementById('zoom');
  function setZoom(z) {
    z = Math.max(25, Math.min(200, Math.round(z)));
    zoom.value = z;
    page.style.zoom = z / 100;
    document.getElementById('zoomLabel').textContent = z + '%';
  }
  zoom.addEventListener('input', () => setZoom(+zoom.value));
  document.getElementById('zoomIn').addEventListener('click', () => setZoom(+zoom.value + 10));
  document.getElementById('zoomOut').addEventListener('click', () => setZoom(+zoom.value - 10));
  document.getElementById('zoom100').addEventListener('click', () => setZoom(100));
  document.getElementById('zoomWidth').addEventListener('click', () => {
    const avail = document.getElementById('canvas').clientWidth - 40;
    setZoom(avail / (page.offsetWidth) * 100);
  });

  // ---------- Find & replace ----------
  const findbar = document.getElementById('findbar');
  const findText = document.getElementById('findText');
  const replaceText = document.getElementById('replaceText');
  function openFind(withReplace) {
    findbar.classList.add('open');
    replaceText.style.display = withReplace ? '' : 'none';
    document.getElementById('replaceOne').style.display = withReplace ? '' : 'none';
    document.getElementById('replaceAll').style.display = withReplace ? '' : 'none';
    const s = getSelection().toString();
    if (s) findText.value = s;
    findText.focus();
    findText.select();
  }
  /** Finds the next/previous match of the find text inside the document body (across formatting). */
  function findNext(back) {
    const q = findText.value.toLowerCase();
    if (!q) return false;
    const walker = document.createTreeWalker(doc, NodeFilter.SHOW_TEXT);
    const nodes = [];
    let text = '';
    while (walker.nextNode()) { nodes.push({ n: walker.currentNode, at: text.length }); text += walker.currentNode.nodeValue; }
    const lower = text.toLowerCase();
    // Current position in the joined text.
    let pos = back ? text.length : 0;
    const sel = getSelection();
    // Search from the caret if the user clicked into the document since, else from the last match.
    const caret = document.activeElement === doc && sel.rangeCount && doc.contains(sel.anchorNode) ? sel.getRangeAt(0) : null;
    const cur = caret || (lastFound && doc.contains(lastFound.startContainer) ? lastFound : null);
    if (cur) {
      const r = cur;
      const hit = nodes.find(x => x.n === (back ? r.startContainer : r.endContainer));
      if (hit) pos = hit.at + (back ? r.startOffset : r.endOffset);
    }
    let i = back ? lower.lastIndexOf(q, pos - 1) : lower.indexOf(q, pos);
    if (i < 0) i = back ? lower.lastIndexOf(q) : lower.indexOf(q); // wrap around
    if (i < 0) { UI.toast('No matches'); return false; }
    const locate = off => { let k = nodes.length - 1; while (k > 0 && nodes[k].at > off) k--; return [nodes[k].n, off - nodes[k].at]; };
    const [sn, so] = locate(i);
    const [en, eo] = locate(i + q.length);
    const range = document.createRange();
    range.setStart(sn, so);
    range.setEnd(en, Math.min(eo, en.nodeValue.length));
    lastFound = range.cloneRange();
    if (window.CSS && CSS.highlights && window.Highlight) {
      // Highlight without moving the selection: selecting inside the document would steal focus
      // from the find box, and the key that triggered the search would then edit the text.
      CSS.highlights.set('find-match', new Highlight(range));
    } else {
      sel.removeAllRanges();
      sel.addRange(range);
      savedRange = range.cloneRange();
      findText.focus();
    }
    (sn.parentElement || doc).scrollIntoView({ block: 'center', behavior: 'smooth' });
    return true;
  }
  let lastFound = null;
  function closeFind() {
    findbar.classList.remove('open');
    if (window.CSS && CSS.highlights) CSS.highlights.delete('find-match');
    // Leave the last match selected so the user can act on it.
    if (lastFound && doc.contains(lastFound.startContainer)) {
      doc.focus();
      getSelection().removeAllRanges();
      getSelection().addRange(lastFound);
      savedRange = lastFound.cloneRange();
    }
    lastFound = null;
  }
  findText.addEventListener('keydown', e => {
    if (e.key === 'Enter') { e.preventDefault(); findNext(e.shiftKey); }
    if (e.key === 'Escape') { e.preventDefault(); closeFind(); }
  });
  findText.addEventListener('input', () => { lastFound = null; if (window.CSS && CSS.highlights) CSS.highlights.delete('find-match'); });
  document.getElementById('findNext').addEventListener('click', () => findNext(false));
  document.getElementById('findPrev').addEventListener('click', () => findNext(true));
  document.getElementById('findClose').addEventListener('click', closeFind);
  document.getElementById('replaceOne').addEventListener('click', () => {
    // Replace the match found last (the live selection is in the find box while typing there).
    if (lastFound && lastFound.toString().toLowerCase() === findText.value.toLowerCase() && doc.contains(lastFound.startContainer)) {
      doc.focus();
      const sel = getSelection();
      sel.removeAllRanges();
      sel.addRange(lastFound);
      document.execCommand('insertText', false, replaceText.value);
      // Continue after the replacement.
      lastFound = getSelection().getRangeAt(0).cloneRange();
      changed();
    }
    findNext(false);
  });
  document.getElementById('replaceAll').addEventListener('click', () => {
    if (!findText.value) return;
    const re = new RegExp(findText.value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
    let count = 0;
    const walker = document.createTreeWalker(doc, NodeFilter.SHOW_TEXT);
    const nodes = [];
    while (walker.nextNode()) nodes.push(walker.currentNode);
    nodes.forEach(n => {
      const v = n.nodeValue.replace(re, () => { count++; return replaceText.value; });
      if (v !== n.nodeValue) n.nodeValue = v;
    });
    UI.toast(`Replaced ${count} occurrence${count === 1 ? '' : 's'}`);
    if (count) changed();
  });

  // ---------- Status & change tracking ----------
  // ---------- Pages ----------
  let toc = null;
  const layout = WriteLayout({
    page, doc, bgs: document.getElementById('pageBgs'),
    hf: { header, footer, firstHeader, firstFooter },
    setup: () => {
      const [w, h] = PAGE_SIZES[setup.size];
      const land = setup.orientation === 'landscape';
      return { width: land ? h : w, height: land ? w : h, margin: parseFloat(setup.margin), differentFirst: setup.differentFirst, columns: setup.columns };
    },
    // Narrow windows (phones) and Web Layout show one continuous page.
    isPaged: () => !document.body.classList.contains('view-web') && window.innerWidth >= 900,
    onLayout: () => { updateCaretPage(); if (toc) toc.refreshPages(); if (extras) extras.updateFields(); },
    footnotes: () => notes.footnotes()
  });
  toc = WriteToc({
    doc, restoreSelection, changed: () => changed(), relayout: () => layout.run(),
    zoom: () => parseFloat(page.style.zoom) || 1, pageAt: y => layout.pageAt(y)
  });
  // ---------- References: notes, citations, bibliography ----------
  const notes = WriteNotes({ doc, setup: () => setup, restoreSelection, captureSelection, changed: () => changed() });
  extras = WriteExtras({
    doc, setup: () => setup, setStyles: st => { setup.styles = st; }, captureSelection, restoreSelection, changed: () => changed(),
    author: () => getAuthor(), fileName: () => state.name, pages: () => layout.pages,
    pageOf: el => layout.pageAt((el.getBoundingClientRect().top - doc.getBoundingClientRect().top) / (parseFloat(page.style.zoom) || 1)),
    wordCount: () => (doc.innerText.match(/\S+/g) || []).length
  });
  document.getElementById('insFootnote').addEventListener('click', () => notes.insertNote('foot'));
  document.getElementById('insEndnote').addEventListener('click', () => notes.insertNote('end'));
  document.getElementById('insCitation').addEventListener('click', () => notes.insertCitation());
  document.getElementById('insBibliography').addEventListener('click', () => notes.insertBibliography());
  document.getElementById('manageSources').addEventListener('click', () => notes.manageSources());
  document.getElementById('citeStyle').addEventListener('change', e => { setup.citeStyle = e.target.value; notes.refreshCitations(); changed(); });

  // ---------- Mailings ----------
  const merge = WriteMerge({
    doc, extraRoots: () => HF, restoreSelection, changed: () => changed(),
    onChange: () => updateMailings(), buildHtml: fn => toHtml(fn), docName: () => state.name
  });
  function updateMailings() {
    const sel = document.getElementById('mmField');
    const names = [...new Set([...merge.fields, ...merge.fieldsInDocument()])];
    sel.innerHTML = '<option value="">Insert Merge Field…</option>' + names.map(n => `<option>${escapeHtml(n)}</option>`).join('') + '<option value="__other">Other field…</option>';
    document.getElementById('mmSource').textContent = merge.count ? `${merge.count} recipients · ${merge.source}` : 'No recipient list';
    document.getElementById('mmIndex').textContent = merge.count ? `${merge.index + 1} of ${merge.count}` : '–';
    const pv = document.getElementById('mmPreview');
    pv.classList.toggle('on', merge.previewing);
    pv.setAttribute('aria-pressed', String(merge.previewing));
  }
  document.getElementById('mmRecipients').addEventListener('click', () => merge.selectRecipients());
  document.getElementById('mmField').addEventListener('change', async e => {
    let v = e.target.value;
    e.target.value = '';
    if (v === '__other') v = (await UI.prompt('Merge field', 'Field name (a column in your list)', '')) || '';
    if (v) merge.insertField(v.trim());
  });
  document.getElementById('mmPreview').addEventListener('click', () => { merge.setPreview(!merge.previewing); layout.schedule(); });
  [['mmFirst', 'first'], ['mmPrev', 'previous'], ['mmNext', 'next'], ['mmLast', 'last']].forEach(([id, fn]) => document.getElementById(id).addEventListener('click', () => { merge[fn](); layout.schedule(); }));
  document.getElementById('mmFinish').addEventListener('click', () => merge.finish());
  updateMailings();

  // ---------- Templates ----------
  async function newFromTemplate() {
    const t = await WriteTemplates.pick();
    if (!t) return;
    const empty = !state.dirty && !doc.innerText.trim() && !state.path;
    // Fill this window if it's still blank; otherwise open a new one.
    if (empty) {
      await loadFile({ name: (t.id === 'blank' ? 'Document1' : t.name) + '.docx', path: null, data: new TextEncoder().encode(WriteTemplates.toHtml(t)), recovered: true });
      state.set({ dirty: t.id !== 'blank' });
    } else WriteTemplates.open(t);
  }
  window.addEventListener('resize', () => layout.schedule());
  const pageCount = () => layout.pages;
  let caretPage = 1;
  function updateCaretPage() {
    const sel = getSelection();
    if (sel.rangeCount && doc.contains(sel.anchorNode)) {
      const r = sel.getRangeAt(0).getClientRects()[0] || (sel.anchorNode.nodeType === 1 ? sel.anchorNode : sel.anchorNode.parentElement).getBoundingClientRect();
      const z = parseFloat(page.style.zoom) || 1;
      caretPage = layout.pageAt((r.top - doc.getBoundingClientRect().top) / z);
    }
    const n = layout.pages;
    document.getElementById('stPages').textContent = `Page ${Math.min(caretPage, n)} of ${n}`;
  }
  // The word count reads the whole document, so while typing it waits for a pause.
  let wordTimer;
  function countWords() {
    clearTimeout(wordTimer);
    const words = (doc.innerText.match(/\S+/g) || []).length;
    document.getElementById('stWords').textContent = `${words.toLocaleString()} word${words === 1 ? '' : 's'}`;
  }
  function updateStatus(typing) {
    if (typing) { clearTimeout(wordTimer); wordTimer = setTimeout(countWords, 400); } else countWords();
    updateCaretPage();
  }
  let navTimer;
  function changed() {
    state.markDirty();
    if (extras) extras.renderEquations();
    notes.renumber();
    layout.schedule();
    updateStatus(true);
    clearTimeout(navTimer);
    navTimer = setTimeout(refreshNav, 400);
  }
  [doc, ...HF].forEach(el => el.addEventListener('input', changed));

  // Keyboard shortcuts the browser doesn't give contenteditable for free
  doc.addEventListener('keydown', e => {
    const mod = e.ctrlKey || e.metaKey;
    if (e.key === 'Tab') {
      e.preventDefault();
      if (closest('li')) exec(e.shiftKey ? 'outdent' : 'indent');
      else if (closest('td,th')) {
        const cells = [...closest('table').querySelectorAll('td,th')];
        const i = cells.indexOf(closest('td,th'));
        const next = cells[i + (e.shiftKey ? -1 : 1)];
        if (next) { const r = document.createRange(); r.selectNodeContents(next); r.collapse(true); getSelection().removeAllRanges(); getSelection().addRange(r); }
      } else document.execCommand('insertText', false, '\u00a0\u00a0\u00a0\u00a0');
    }
    if (mod && e.key === 'Enter') { e.preventDefault(); document.getElementById('insPageBreak').click(); }
    if (e.key === 'F9' && !mod) { e.preventDefault(); extras.updateFields(true); if (toc) toc.refreshPages(); UI.toast('Fields updated'); }
    if (mod && !e.shiftKey && e.key.toLowerCase() === 'e') { e.preventDefault(); exec('justifyCenter'); }
    if (mod && !e.shiftKey && e.key.toLowerCase() === 'l') { e.preventDefault(); exec('justifyLeft'); }
    if (mod && !e.shiftKey && e.key.toLowerCase() === 'r') { e.preventDefault(); exec('justifyRight'); }
    if (mod && !e.shiftKey && e.key.toLowerCase() === 'j') { e.preventDefault(); exec('justifyFull'); }
    if (mod && e.key.toLowerCase() === 'k') { e.preventDefault(); document.getElementById('insLink').click(); }
    if (mod && e.altKey && /^[123]$/.test(e.key)) { e.preventDefault(); exec('formatBlock', 'H' + e.key); }
    if (mod && e.key === ']') { e.preventDefault(); document.getElementById('growFont').click(); }
    if (mod && e.key === '[') { e.preventDefault(); document.getElementById('shrinkFont').click(); }
  });
  // Keep the document from ever being completely empty (so typing starts in a paragraph).
  doc.addEventListener('keyup', () => { if (!doc.firstElementChild) doc.innerHTML = '<p><br></p>'; });

  // Paste: keep formatting but strip anything unsafe (header and footer too).
  [doc, ...HF].forEach(el => el.addEventListener('paste', e => {
    if (review.tracking && el === doc) {
      // Tracked pastes are recorded as plain-text insertions.
      e.preventDefault();
      review.insertText(e.clipboardData.getData('text/plain'));
      return;
    }
    const html = e.clipboardData.getData('text/html');
    if (!html) return;
    e.preventDefault();
    document.execCommand('insertHTML', false, sanitize(html));
    fixFields(el);
    changed();
  }));

  // ---------- Import ----------
  function escapeHtml(s) { return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
  const sanitize = html => Platform.sanitizeHtml(html);
  function markdownToHtml(md) {
    const inline = s => escapeHtml(s)
      .replace(/`([^`]+)`/g, '<code>$1</code>')
      .replace(/\*\*([^*]+)\*\*/g, '<b>$1</b>').replace(/__([^_]+)__/g, '<b>$1</b>')
      .replace(/\*([^*]+)\*/g, '<i>$1</i>').replace(/(^|\W)_([^_]+)_(?=\W|$)/g, '$1<i>$2</i>')
      .replace(/~~([^~]+)~~/g, '<s>$1</s>')
      .replace(/!\[([^\]]*)\]\(([^)\s]+)\)/g, '<img alt="$1" src="$2">')
      .replace(/\[([^\]]+)\]\(((?:https?:|mailto:|#)[^)\s]*)\)/g, '<a href="$2">$1</a>');
    const lines = md.replace(/\r\n?/g, '\n').split('\n');
    let out = '', list = null, para = [], code = null;
    const flushPara = () => { if (para.length) { out += `<p>${inline(para.join(' '))}</p>`; para = []; } };
    const closeList = () => { if (list) { out += `</${list}>`; list = null; } };
    for (const line of lines) {
      if (code !== null) {
        if (/^```/.test(line)) { out += `<pre>${escapeHtml(code.join('\n'))}</pre>`; code = null; } else code.push(line);
        continue;
      }
      let m;
      if (/^```/.test(line)) { flushPara(); closeList(); code = []; }
      else if ((m = line.match(/^(#{1,6})\s+(.*)/))) { flushPara(); closeList(); const l = Math.min(m[1].length, 3); out += `<h${l}>${inline(m[2])}</h${l}>`; }
      else if ((m = line.match(/^\s*[-*+]\s+(.*)/))) { flushPara(); if (list !== 'ul') { closeList(); out += '<ul>'; list = 'ul'; } out += `<li>${inline(m[1])}</li>`; }
      else if ((m = line.match(/^\s*\d+[.)]\s+(.*)/))) { flushPara(); if (list !== 'ol') { closeList(); out += '<ol>'; list = 'ol'; } out += `<li>${inline(m[1])}</li>`; }
      else if ((m = line.match(/^>\s?(.*)/))) { flushPara(); closeList(); out += `<blockquote>${inline(m[1])}</blockquote>`; }
      else if (/^(-{3,}|\*{3,})\s*$/.test(line)) { flushPara(); closeList(); out += '<hr>'; }
      else if (!line.trim()) { flushPara(); closeList(); }
      else { closeList(); para.push(line.trim()); }
    }
    if (code !== null) out += `<pre>${escapeHtml(code.join('\n'))}</pre>`;
    flushPara(); closeList();
    return out || '<p><br></p>';
  }
  function htmlToMarkdown(root) {
    const walk = (n, ctx = {}) => {
      if (n.nodeType === 3) return n.nodeValue.replace(/\u00a0/g, ' ');
      if (n.nodeType !== 1) return '';
      if (n.classList.contains('toc-pg')) return '';
      const kids = () => [...n.childNodes].map(c => walk(c, ctx)).join('');
      switch (n.tagName) {
        case 'H1': case 'H2': case 'H3': case 'H4': return '#'.repeat(+n.tagName[1]) + ' ' + kids().trim() + '\n\n';
        case 'P': case 'DIV': return kids().trim() + '\n\n';
        case 'B': case 'STRONG': return `**${kids()}**`;
        case 'I': case 'EM': return `*${kids()}*`;
        case 'S': case 'STRIKE': return `~~${kids()}~~`;
        case 'CODE': return '`' + kids() + '`';
        case 'PRE': return '```\n' + n.innerText + '\n```\n\n';
        case 'BLOCKQUOTE': return '> ' + kids().trim() + '\n\n';
        case 'A': return `[${kids()}](${n.getAttribute('href') || ''})`;
        case 'IMG': return `![${n.alt || ''}](${n.getAttribute('src')?.startsWith('data:') ? 'image' : n.getAttribute('src')})`;
        case 'BR': return '\n';
        case 'HR': return '---\n\n';
        case 'UL': return [...n.children].map(li => '- ' + walk(li).trim()).join('\n') + '\n\n';
        case 'OL': return [...n.children].map((li, i) => `${i + 1}. ` + walk(li).trim()).join('\n') + '\n\n';
        case 'TABLE': {
          const rows = [...n.querySelectorAll('tr')].map(tr => '| ' + [...tr.children].map(c => c.innerText.trim().replace(/\|/g, '\\|')).join(' | ') + ' |');
          if (rows.length) rows.splice(1, 0, '|' + ' --- |'.repeat(n.querySelector('tr').children.length));
          return rows.join('\n') + '\n\n';
        }
        default: return kids();
      }
    };
    return walk(root).replace(/\n{3,}/g, '\n\n').trim() + '\n';
  }

  /** Reads a file into HTML with its page setup, headers/footers and sources (throws if it can't). */
  async function readDocument(f) {
    const ext = Platform.extOf(f.name);
    let html;
    let nextSetup = { ...DEFAULT_SETUP };
    let hfHtml = { header: '', footer: '', firstHeader: '', firstFooter: '' };
    let nextSources = [];
    if (ext === 'docx' && !/^\s*(<!doctype html|<html)/i.test(Platform.bytesToText(f.data.slice(0, 200)))) {
      // Page setup, headers, footers and page numbers come from the section properties (mammoth skips them).
      try {
        const sec = await readDocxSection(f.data);
        nextSetup = validSetup({ ...nextSetup, ...sec.setup });
        hfHtml = { ...hfHtml, ...sec.hf };
        nextSources = sec.sources || [];
      } catch (err) { console.info('Could not read page setup:', err); }
      // Keep tracked changes and comments, which the HTML converter would otherwise drop.
      let pre = { bytes: f.data, meta: null };
      try { pre = await review.preprocessDocx(f.data); } catch (err) { console.info('Could not read revisions:', err); }
      // Equations, fields, section breaks, wrapped pictures and styles.
      let pre2 = { bytes: pre.bytes, meta: null };
      try { pre2 = await extras.preprocessDocx(pre.bytes); } catch (err) { console.info('Could not read equations and sections:', err); }
      if (pre2.meta) {
        nextSetup.styles = pre2.meta.styles || {};
        if (pre2.meta.first) { nextSetup.orientation = pre2.meta.first.orient; nextSetup.columns = pre2.meta.first.cols; }
        nextSetup = validSetup(nextSetup);
      }
      const src = pre2.bytes;
      const res = await mammoth.convertToHtml({ arrayBuffer: src.buffer.slice(src.byteOffset, src.byteOffset + src.byteLength) }, {
        styleMap: ['p[style-name=\'Title\'] => h1.title:fresh', 'p[style-name=\'Quote\'] => blockquote:fresh', 'p[style-name=\'Intense Quote\'] => blockquote:fresh']
      });
      html = WriteToc.importHtml(notes.importMammothNotes(review.postprocessHtml(res.value, pre.meta)));
      if (pre2.meta) html = extras.postprocessHtml(html, pre2.meta);
      if (res.messages.some(m => m.type === 'warning')) console.info('Import notes:', res.messages);
    } else if (ext === 'html' || ext === 'htm' || /^\s*(<!doctype html|<html)/i.test(Platform.bytesToText(f.data.slice(0, 200)))) {
      const d = new DOMParser().parseFromString(Platform.bytesToText(f.data), 'text/html');
      const src = d.querySelector('meta[name=atlas-sources]');
      if (src) { try { nextSources = JSON.parse(src.content); } catch { /* ignore */ } }
      const saved = d.querySelector('meta[name=atlas-setup]');
      if (saved) { try { nextSetup = validSetup(JSON.parse(saved.content)); } catch { /* ignore */ } }
      for (const [key, sel] of [['header', 'header.atlas-header'], ['footer', 'footer.atlas-footer'], ['firstHeader', 'header.atlas-first-header'], ['firstFooter', 'footer.atlas-first-footer']]) {
        const el = d.querySelector(sel);
        if (el) { hfHtml[key] = el.innerHTML; el.remove(); }
      }
      html = d.querySelector('main.atlas-doc')?.innerHTML ?? d.body.innerHTML;
    } else if (ext === 'md') {
      html = markdownToHtml(Platform.bytesToText(f.data));
    } else {
      html = Platform.bytesToText(f.data).replace(/\r\n?/g, '\n').split(/\n/).map(l => `<p>${escapeHtml(l) || '<br>'}</p>`).join('');
    }
    return { html, setup: nextSetup, hf: hfHtml, sources: nextSources };
  }
  async function loadFile(f) {
    // A password-protected file asks for its password first.
    try { f = await Platform.unlockFile(f); } catch (err) { UI.toast('Could not open file: ' + err.message, 5000); return; }
    if (!f) return;
    let html, nextSetup, hfHtml, nextSources;
    try {
      ({ html, setup: nextSetup, hf: hfHtml, sources: nextSources } = await readDocument(f));
    } catch (err) {
      UI.toast('Could not open file: ' + err.message, 4000);
      return;
    }
    // Nothing from the previous document (page setup, header, footer) carries over.
    Object.assign(setup, nextSetup);
    for (const [key, el] of Object.entries({ header, footer, firstHeader, firstFooter })) { el.innerHTML = sanitize(hfHtml[key]); fixFields(el); }
    doc.innerHTML = sanitize(html) || '<p><br></p>';
    fixFields(doc);
    extras.renderEquations();
    notes.loadSources(nextSources);
    notes.renumber();
    notes.refreshCitations();
    merge.refresh();
    savedRange = null;
    applySetup();
    syncSetupControls();
    state.set({ name: f.name, path: f.path, dirty: false });
    if (!f.recovered) Platform.rememberRecent(f.name, 'write', f.path);
    extras.updateFields();
    updateStatus();
    refreshNav();
  }
  function syncSetupControls() {
    document.getElementById('margins').value = setup.margin;
    document.getElementById('orientation').value = setup.orientation;
    document.getElementById('pageSize').value = setup.size;
    document.getElementById('columns').value = String(setup.columns);
    document.getElementById('pageColor').value = setup.color;
  }

  // ---------- Export ----------
  const DOCX_STYLESHEET = `
    h1.title { font-size: 28pt; color: #000000; }
    h1 { font-size: 16pt; color: #2f5496; font-weight: normal; }
    h2 { font-size: 13pt; color: #2f5496; font-weight: normal; }
    h3 { font-size: 12pt; color: #1f3763; font-weight: normal; }
    blockquote { font-style: italic; color: #404040; }
    pre { font-family: Consolas; font-size: 10pt; }
    .textbox { border: 1px solid #555555; }
    td, th { border: 1px solid #8e8e8e; }
  `;
  /** Reads page size, margins, columns, headers/footers and page-number fields from a .docx. */
  async function readDocxSection(bytes) {
    const zip = await JSZip.loadAsync(bytes);
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const xml = async name => { const f = zip.file(name); return f ? new DOMParser().parseFromString(await f.async('string'), 'application/xml') : null; };
    const docXml = await xml('word/document.xml');
    const relsXml = await xml('word/_rels/document.xml.rels');
    if (!docXml) return { setup: {}, hf: {} };
    const rels = {};
    if (relsXml) [...relsXml.getElementsByTagName('Relationship')].forEach(r => { rels[r.getAttribute('Id')] = r.getAttribute('Target'); });
    const sects = docXml.getElementsByTagNameNS(W, 'sectPr');
    const sect = sects[sects.length - 1];
    const setup = {};
    const hf = {};
    if (!sect) return { setup, hf };
    const attr = (el, name) => el && (el.getAttributeNS(W, name) || el.getAttribute('w:' + name));
    const pgSz = sect.getElementsByTagNameNS(W, 'pgSz')[0];
    if (pgSz) {
      let w = +attr(pgSz, 'w'), h = +attr(pgSz, 'h');
      if (w > h) { [w, h] = [h, w]; setup.orientation = 'landscape'; } else setup.orientation = attr(pgSz, 'orient') === 'landscape' ? 'landscape' : 'portrait';
      const sizes = { letter: [12240, 15840], a4: [11906, 16838], legal: [12240, 20160] };
      setup.size = Object.entries(sizes).sort((a, b) => Math.hypot(a[1][0] - w, a[1][1] - h) - Math.hypot(b[1][0] - w, b[1][1] - h))[0][0];
    }
    const pgMar = sect.getElementsByTagNameNS(W, 'pgMar')[0];
    if (pgMar) {
      const avg = ['top', 'bottom', 'left', 'right'].map(k => +attr(pgMar, k) || 1440).reduce((a, b) => a + b, 0) / 4 / 1440;
      setup.margin = [0.5, 0.75, 1, 1.5].reduce((best, m) => (Math.abs(m - avg) < Math.abs(best - avg) ? m : best), 1) + 'in';
    }
    const cols = sect.getElementsByTagNameNS(W, 'cols')[0];
    if (cols && +attr(cols, 'num') > 1) setup.columns = Math.min(3, +attr(cols, 'num'));
    const titlePg = sect.getElementsByTagNameNS(W, 'titlePg')[0];
    setup.differentFirst = !!titlePg && !['0', 'false'].includes(attr(titlePg, 'val'));
    const settings = await xml('word/settings.xml');
    const tr = settings && settings.getElementsByTagNameNS(W, 'trackRevisions')[0];
    setup.trackChanges = !!tr && !['0', 'false'].includes(attr(tr, 'val'));
    const partToHtml = async target => {
      const part = await xml('word/' + target.replace(/^\/?word\//, ''));
      if (!part) return '';
      return [...part.getElementsByTagNameNS(W, 'p')].map(p => {
        const jc = attr(p.getElementsByTagNameNS(W, 'jc')[0], 'val');
        const align = { center: 'center', right: 'right', end: 'right', both: 'justify' }[jc] || '';
        let out = '', inField = false, fieldDone = false, instr = '';
        const field = kind => `<span class="fld" data-fld="${kind}">1</span>`;
        const walk = node => {
          for (const n of node.childNodes) {
            if (n.nodeType !== 1) continue;
            const ln = n.localName;
            if (ln === 'fldSimple') {
              const ins = (attr(n, 'instr') || '').trim().toUpperCase();
              out += /^NUMPAGES/.test(ins) ? field('numpages') : /^PAGE/.test(ins) ? field('page') : escapeHtml(n.textContent);
            } else if (ln === 'fldChar') {
              const t = attr(n, 'fldCharType');
              if (t === 'begin') { inField = true; fieldDone = false; instr = ''; }
              else if (t === 'separate' && inField) { const ins = instr.trim().toUpperCase(); if (/^(PAGE|NUMPAGES)/.test(ins)) { out += field(/^NUMPAGES/.test(ins) ? 'numpages' : 'page'); fieldDone = true; } }
              else if (t === 'end') { inField = false; }
            } else if (ln === 'instrText') {
              instr += n.textContent;
            } else if (ln === 't') {
              if (!(inField && fieldDone)) {
                const r = n.parentNode;
                const rPr = r && r.getElementsByTagNameNS(W, 'rPr')[0];
                let t = escapeHtml(n.textContent);
                if (rPr && rPr.getElementsByTagNameNS(W, 'b')[0]) t = `<b>${t}</b>`;
                if (rPr && rPr.getElementsByTagNameNS(W, 'i')[0]) t = `<i>${t}</i>`;
                out += t;
              }
            } else if (ln === 'tab') out += '\u00a0\u00a0\u00a0\u00a0';
            else if (ln === 'br') out += '<br>';
            else if (ln !== 'pPr' && ln !== 'rPr') walk(n);
          }
        };
        walk(p);
        return out ? `<p${align ? ` style="text-align:${align}"` : ''}>${out}</p>` : '';
      }).join('');
    };
    for (const kind of ['header', 'footer']) {
      for (const ref of sect.getElementsByTagNameNS(W, kind + 'Reference')) {
        const type = attr(ref, 'type') || 'default';
        const target = rels[ref.getAttributeNS(R, 'id') || ref.getAttribute('r:id')];
        if (!target) continue;
        const html = await partToHtml(target);
        if (type === 'default') hf[kind] = html;
        if (type === 'first') hf['first' + kind[0].toUpperCase() + kind.slice(1)] = html;
      }
    }
    const custom = await xml('docProps/custom.xml');
    let sources = [];
    if (custom) {
      const prop = name => [...custom.getElementsByTagName('property')].find(p => p.getAttribute('name') === name);
      if (prop('AtlasSources')) { try { sources = JSON.parse(prop('AtlasSources').textContent); } catch { /* ignore */ } }
      if (prop('AtlasCiteStyle')) setup.citeStyle = prop('AtlasCiteStyle').textContent.trim();
    }
    return { setup, hf, sources };
  }

  function cleanHtmlFor(el) {
    const clone = el.cloneNode(true);
    layout.clearPushes(clone);
    notes.clean(clone);
    clone.querySelectorAll('span.mergefield').forEach(m => { m.textContent = `«${m.dataset.field}»`; m.classList.remove('preview'); });
    clone.querySelectorAll('.selected').forEach(n => n.classList.remove('selected'));
    clone.querySelectorAll('[contenteditable]').forEach(n => n.removeAttribute('contenteditable'));
    clone.querySelectorAll('hr.page-break').forEach(n => { const d = document.createElement('div'); d.style.pageBreakAfter = 'always'; n.replaceWith(d); });
    return clone.innerHTML;
  }
  async function toDocx() {
    const D = window.docx;
    const [w, h] = PAGE_SIZES[setup.size];
    const twip = inch => Math.round(inch * 1440);
    const margin = twip(parseFloat(setup.margin));
    const textWidth = (setup.orientation === 'landscape' ? h : w) - 2 * parseFloat(setup.margin);
    // The document's own styles come after the defaults, so they win.
    const opts = { stylesheet: DOCX_STYLESHEET + '\n' + extras.styleCss(''), maxImageWidth: Math.round(textWidth / setup.columns * 96) };
    const pageProps = (orient, cols) => ({
      page: {
        size: { width: twip(w), height: twip(h), orientation: orient === 'landscape' ? D.PageOrientation.LANDSCAPE : D.PageOrientation.PORTRAIT },
        margin: { top: margin, right: margin, bottom: margin, left: margin }
      },
      column: cols > 1 ? { count: cols, space: 720 } : undefined
    });
    const section = {
      properties: {
        page: {
          size: { width: twip(w), height: twip(h), orientation: setup.orientation === 'landscape' ? D.PageOrientation.LANDSCAPE : D.PageOrientation.PORTRAIT },
          margin: { top: margin, right: margin, bottom: margin, left: margin }
        },
        column: setup.columns > 1 ? { count: setup.columns, space: 720 } : undefined
      },
      children: []
    };
    const body = document.createElement('div');
    body.innerHTML = cleanHtmlFor(doc);
    // Editors store a space typed at the edge of a text run as a no-break space; Word wants a normal one.
    const tw = document.createTreeWalker(body, NodeFilter.SHOW_TEXT);
    while (tw.nextNode()) tw.currentNode.nodeValue = tw.currentNode.nodeValue.replace(/(^|[^\u00a0 ])\u00a0(?=[^\u00a0 ]|$)/g, '$1 ');
    const cmts = review.markCommentsForDocx(body);
    toc.prepareForDocx(body);
    notes.prepareForDocx(body);
    merge.prepareForDocx(body);
    const objects = extras.prepareForDocx(body);
    // Each section break starts a Word section with its own orientation and columns.
    const parts = extras.splitSections(body);
    section.children = [...docxHtml.htmlToContent(parts[0].html, { ...opts, trackChanges: { author: getAuthor() } })];
    const more = parts.slice(1).map(p => ({ properties: pageProps(p.orient, p.cols), children: [...docxHtml.htmlToContent(p.html, { ...opts, trackChanges: { author: getAuthor() } })] }));
    const has = el => el.textContent.trim() !== '' || !!el.querySelector('img');
    const headers = {}, footers = {};
    if (has(header)) headers.default = new D.Header({ children: hfToDocx(header, opts) });
    if (has(footer)) footers.default = new D.Footer({ children: hfToDocx(footer, opts) });
    if (setup.differentFirst) {
      section.properties.titlePage = true;
      // An empty first-page header must still exist, or Word repeats the default one.
      headers.first = new D.Header({ children: has(firstHeader) ? hfToDocx(firstHeader, opts) : [new D.Paragraph('')] });
      footers.first = new D.Footer({ children: has(firstFooter) ? hfToDocx(firstFooter, opts) : [new D.Paragraph('')] });
    }
    if (Object.keys(headers).length) section.headers = headers;
    if (Object.keys(footers).length) section.footers = footers;
    const document_ = new D.Document({
      creator: 'Isadek Write',
      title: Platform.baseName(state.name),
      styles: { default: docxDefaultStyles() },
      background: setup.color !== '#ffffff' ? { color: setup.color.slice(1) } : undefined,
      features: setup.trackChanges ? { trackRevisions: true } : undefined,
      customProperties: [{ name: 'AtlasCiteStyle', value: setup.citeStyle }, ...(notes.sources.length ? [{ name: 'AtlasSources', value: JSON.stringify(notes.sources) }] : [])],
      comments: cmts.length ? { children: cmts.map(c => ({ id: c.id, author: c.author, initials: c.author.split(/\s+/).map(w => w[0]).join('').slice(0, 3), date: new Date(c.date), children: c.text.split('\n').map(t => new D.Paragraph(t)) })) } : undefined,
      sections: [section, ...more]
    });
    const blob = await D.Packer.toBlob(document_);
    const tabPos = Math.round(textWidth / setup.columns * 1440) - (setup.columns > 1 ? 360 : 0);
    const patched = await WriteToc.patchDocx(await review.patchDocx(new Uint8Array(await blob.arrayBuffer())), { tabPos });
    return extras.patchDocx(patched, objects);
  }
  /** Word's default and heading styles, with the document's own changes to them. */
  function docxDefaultStyles() {
    const mine = extras.docxStyles();
    const doc0 = { run: { font: 'Calibri', size: 22 }, paragraph: { spacing: { after: 160, line: 276 } } };
    if (mine.document) {
      Object.assign(doc0.run, mine.document.run);
      if (mine.document.paragraph) Object.assign(doc0.paragraph.spacing, mine.document.paragraph.spacing);
    }
    return { ...mine, document: doc0 };
  }
  /** Header/footer content as Word paragraphs, with page-number fields as real Word fields. */
  function hfToDocx(el, opts) {
    const D = window.docx;
    const blockTags = /^(P|DIV|H[1-6]|UL|OL|TABLE|BLOCKQUOTE|PRE)$/;
    const blocks = [...el.children].some(n => blockTags.test(n.tagName)) ? [...el.children] : [el];
    const out = [];
    for (const b of blocks) {
      if (!b.querySelector('.fld')) { out.push(...docxHtml.htmlToContent(b === el ? `<p>${el.innerHTML}</p>` : b.outerHTML, opts)); continue; }
      const children = [];
      let buf = '';
      const flush = () => { if (buf) { children.push(...docxHtml.htmlToRuns(buf, opts)); buf = ''; } };
      b.childNodes.forEach(n => {
        if (n.nodeType === 1 && n.classList.contains('fld')) { flush(); children.push(new D.TextRun({ children: [n.dataset.fld === 'numpages' ? D.PageNumber.TOTAL_PAGES : D.PageNumber.CURRENT] })); }
        // Plain text becomes a run as-is: converting it as HTML would trim the spaces next to fields.
        else if (n.nodeType === 3) { flush(); if (n.nodeValue) children.push(new D.TextRun(n.nodeValue.replace(/\u00a0/g, ' '))); }
        else if (n.nodeType === 1) buf += n.outerHTML;
      });
      flush();
      const align = { center: D.AlignmentType.CENTER, right: D.AlignmentType.RIGHT, justify: D.AlignmentType.JUSTIFIED }[b.style && b.style.textAlign];
      out.push(new D.Paragraph({ children, alignment: align }));
    }
    return out;
  }
  function toHtml(bodyTransform) {
    const css = [...document.styleSheets].find(s => s.href && s.href.endsWith('write.css'));
    let rules = '';
    try { rules = [...css.cssRules].map(r => r.cssText).filter(t => t.startsWith('.doc')).join('\n').replace(/\.doc\b/g, 'main.atlas-doc'); } catch { /* ignore */ }
    rules += '\n' + extras.styleCss('main.atlas-doc');
    return `<!doctype html>
<html><head><meta charset="utf-8"><title>${escapeHtml(Platform.baseName(state.name))}</title>
<meta name="generator" content="Isadek Write">
<meta name="atlas-setup" content="${escapeHtml(JSON.stringify(setup))}">
${notes.sources.length ? `<meta name="atlas-sources" content="${escapeHtml(JSON.stringify(notes.sources))}">` : ''}
<style>body{max-width:8.5in;margin:auto;padding:1in;font-family:Calibri,Arial,sans-serif;font-size:11pt;line-height:1.15;background:${setup.color}}
${rules}</style></head>
<body>${hfOut(firstHeader, 'header', 'atlas-first-header')}${hfOut(header, 'header', 'atlas-header')}<main class="atlas-doc">${bodyTransform ? bodyTransform(cleanHtmlFor(doc)) : cleanHtmlFor(doc)}</main>${hfOut(footer, 'footer', 'atlas-footer')}${hfOut(firstFooter, 'footer', 'atlas-first-footer')}</body></html>`;
  }
  function hfOut(el, tag, cls) {
    return el.textContent.trim() || el.querySelector('img') ? `<${tag} class="${cls}">${cleanHtmlFor(el)}</${tag}>` : '';
  }
  async function serialize(ext) {
    extras.updateFields();
    if (ext === 'docx') return toDocx();
    if (ext === 'html' || ext === 'htm') return toHtml();
    if (ext === 'md') {
      const clone = doc.cloneNode(true);
      extras.prepareForText(clone);
      const defs = notes.toMarkdownNotes(clone);
      return htmlToMarkdown(clone) + (defs.length ? '\n' + defs.join('\n') + '\n' : '');
    }
    return doc.innerText;
  }

  async function save(saveAs, forcedExt, quiet) {
    let name = state.name;
    if (!/\.(docx|html?|md|txt)$/i.test(name)) name = Platform.baseName(name) + '.docx';
    if (forcedExt) name = Platform.baseName(name) + '.' + forcedExt;
    // Saving to a lossy format shouldn't silently keep the user on it.
    const filters = forcedExt ? FILTERS.filter(f => f.extensions.includes(forcedExt)) : FILTERS;
    const ext = Platform.extOf(name);
    const reorder = [...filters].sort((a, b) => (b.extensions.includes(ext) ? 1 : 0) - (a.extensions.includes(ext) ? 1 : 0));
    try {
      if (saveAs || !state.path || forcedExt) {
        // Ask for the target first so we can serialize in the format the user picked.
        const target = await Platform.chooseSavePath(name, reorder);
        if (!target) return false;
        const data = await serialize(Platform.extOf(target.name) || 'docx');
        await Platform.save({ data, name: target.name, path: target.path });
        if (!forcedExt) state.set({ name: target.name, path: target.path, dirty: false });
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

  // ---------- Commands ----------
  const CLOUD_EXTS = ['docx', 'html', 'md', 'txt'];
  Platform.onMenu(async cmd => {
    switch (cmd) {
      case 'new': Platform.newWindow('write'); break;
      case 'new-template': newFromTemplate(); break;
      case 'open': {
        const f = await Platform.open(OPEN_FILTERS);
        if (!f) break;
        if (Platform.shouldOpenInPlace(!state.dirty && !doc.innerText.trim())) {
          if (state.dirty && !(await UI.confirm('Open file', 'Discard unsaved changes in this document?'))) break;
          loadFile(f);
        } else Platform.openPath(f.path);
        break;
      }
      case 'save': save(false); break;
      case 'save-as': save(true); break;
      case 'version-history': Platform.versionHistory({ state, app: 'write', load: loadFile }); break;
      case 'open-cloud': Cloud.openDialog({ app: 'write', load: loadFile, isEmpty: async () => Platform.shouldOpenInPlace(!state.dirty && !doc.innerText.trim()) && (!state.dirty || UI.confirm('Open file', 'Discard unsaved changes in this document?')) }); break;
      case 'save-cloud': Cloud.saveAsDialog({ state, serialize: ext => serialize(ext), exts: CLOUD_EXTS }); break;
      case 'share': Cloud.shareDialog({ state, serialize: ext => serialize(ext), exts: CLOUD_EXTS, save: () => save(false) }); break;
      case 'save-and-close': if (await save(false)) Platform.closeNow(); break;
      case 'export-pdf': { const restore = layout.forPrint(); try { await Platform.exportPdf(Platform.baseName(state.name) + '.pdf', setup.orientation === 'landscape'); } finally { restore(); } break; }
      case 'export-html': save(true, 'html'); break;
      case 'export-txt': save(true, 'txt'); break;
      case 'print': { extras.updateFields(); const restore = layout.forPrint(); try { await Platform.print(); } finally { restore(); } break; }
      case 'undo': if (!appUndo(false) && !liveUndo('undoLocal')) exec('undo'); break;
      case 'redo': if (!appUndo(true) && !liveUndo('redoLocal')) exec('redo'); break;
      case 'find': openFind(false); break;
      case 'replace': openFind(true); break;
      case 'home': Platform.desktop ? Platform.newWindow('home') : (location.href = 'index.html'); break;
    }
  });

  // Open a file handed over by the OS / main process or the web start page.
  (async () => {
    applySetup();
    const f = await Platform.pending();
    if (f && f.error) UI.toast('Could not open file: ' + f.error, 4000);
    else if (f && f.recovered) await loadRecovered(f.recovered);
    else if (f) await loadFile(f);
    updateStatus();
    doc.focus();
  })();

  // Printing from the browser's own menu also uses gap-free pages.
  let restorePrint = null;
  window.addEventListener('beforeprint', () => { if (!restorePrint) restorePrint = layout.forPrint(); });
  window.addEventListener('afterprint', () => { if (restorePrint) { const r = restorePrint; restorePrint = null; r(); } });

  // ---------- Crash recovery & AutoSave ----------
  const recovery = Platform.enableRecovery({
    app: 'write', state,
    snapshot: () => toHtml(),
    autosave: () => save(false, null, true),
    canAutoSave: () => !Platform.desktop && !Platform.isCloudPath(state.path) && !(Platform.canWriteBack && Platform.canWriteBack(state.path)) ? 'AutoSave needs the desktop app' : !state.path ? 'Save the document first to turn on AutoSave'
      : /\.(docx|html?)$/i.test(state.path) ? true : 'AutoSave works with .docx and .html files'
  });
  async function loadRecovered(d) {
    await loadFile({ name: 'recovered.html', path: null, data: new TextEncoder().encode(d.data), recovered: true });
    state.set({ name: d.name, path: d.path, dirty: true });
    recovery.adopt(d.draftId);
    UI.toast('Recovered your unsaved changes', 4000);
  }

  // ---------- Live editing ----------
  // Keys: b:<id> -> { p: position, h: the paragraph's HTML } for each top-level block,
  // hf:<n> -> a header or footer, setup -> page setup. Positions are strings that sort in
  // document order, so a paragraph added between two others gets a key between theirs.
  const blockId = new WeakMap(), blockById = new Map(), blockPos = new Map();
  let composing = false, livePeers = [];
  doc.addEventListener('compositionstart', () => { composing = true; });
  doc.addEventListener('compositionend', () => { composing = false; });
  const newBlockId = () => 'b' + Math.random().toString(36).slice(2, 11);
  function blockHtml(el) {
    const clone = el.cloneNode(true);
    const wrap = document.createElement('div');
    wrap.appendChild(clone);
    layout.clearPushes(wrap);
    wrap.querySelectorAll('.selected').forEach(n => n.classList.remove('selected'));
    return wrap.innerHTML;
  }
  function topBlock(node) {
    while (node && node.parentNode !== doc) node = node.parentNode;
    return node && node.nodeType === 1 ? node : null;
  }
  function textOffset(block, node, off) {
    const r = document.createRange();
    r.setStart(block, 0);
    try { r.setEnd(node, off); } catch { return 0; }
    return r.toString().length;
  }
  function pointAt(block, n) {
    const walk = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
    let t;
    while ((t = walk.nextNode())) { if (n <= t.length) return [t, n]; n -= t.length; }
    return [block, block.childNodes.length];
  }
  function caretInfo() {
    const sel = getSelection();
    if (!sel.rangeCount || !doc.contains(sel.anchorNode)) return null;
    const b = topBlock(sel.anchorNode);
    if (!b) return null;
    const fb = topBlock(sel.focusNode) || b;
    return { b: blockId.get(b), o: textOffset(b, sel.anchorNode, sel.anchorOffset), fb: blockId.get(fb), fo: textOffset(fb, sel.focusNode, sel.focusOffset) };
  }
  function liveState() {
    const out = {};
    const kids = [...doc.children].filter(el => !el.classList.contains('endnotes-gen'));
    // Keep each block's position while the order still agrees; new and moved blocks get one between their neighbours.
    const ps = kids.map(el => { let id = blockId.get(el); if (!id || blockById.get(id) !== el) { id = newBlockId(); blockId.set(el, id); blockById.set(id, el); } return blockPos.get(id) || null; });
    let last = '';
    const keep = ps.map(p => { const ok = p && p > last; if (ok) last = p; return ok; });
    for (let i = 0; i < kids.length; i++) {
      const id = blockId.get(kids[i]);
      if (!keep[i]) {
        let prev = '';
        for (let j = i - 1; j >= 0; j--) { prev = blockPos.get(blockId.get(kids[j])); break; }
        let next = '';
        for (let j = i + 1; j < kids.length; j++) if (keep[j]) { next = ps[j]; break; }
        blockPos.set(id, Collab.between(prev || '', next || ''));
      }
      out['b:' + id] = { p: blockPos.get(id), h: blockHtml(kids[i]) };
    }
    for (const id of [...blockById.keys()]) if (!blockById.get(id).isConnected || blockById.get(id).parentNode !== doc) { blockById.delete(id); blockPos.delete(id); }
    HF.forEach((el, i) => { out['hf:' + i] = el.innerHTML; });
    out.setup = { ...setup };
    out.sources = notes.sources;
    return out;
  }
  /** Where an offset in a text ends up after the text changes: moved along by what changed before it. */
  function mapOffset(before, after, off) {
    if (before === after) return off;
    let shift = 0;
    for (const h of Merge3.hunks([...before], [...after])) {
      if (h.e <= off && !(h.s === h.e && h.s === off)) shift += h.rep.length - (h.e - h.s); // wholly before the caret
      else if (h.s < off) return h.s + shift + h.rep.length; // the caret's own text changed: after the change
    }
    return Math.max(0, Math.min(after.length, off + shift));
  }
  function liveApply(changes) {
    setTimeout(() => review.checkMentions(), 0); // a comment that mentions you, from someone else
    const caret = caretInfo();
    // The text of the paragraphs the caret is in, to move it along with what others type before it.
    const textBefore = new Map();
    if (caret) [caret.b, caret.fb].forEach(id => { const el = blockById.get(id); if (el) textBefore.set(id, el.textContent); });
    let sourcesChanged = false;
    for (const [k, v] of Object.entries(changes)) {
      if (k.startsWith('b:')) {
        const id = k.slice(2);
        const el = blockById.get(id);
        if (v === null) { if (el) el.remove(); blockById.delete(id); blockPos.delete(id); continue; }
        blockPos.set(id, String(v.p));
        if (el && blockHtml(el) === v.h) continue;
        const t = document.createElement('template');
        t.innerHTML = sanitize(v.h);
        const nu = t.content.firstElementChild || document.createElement('p');
        blockId.set(nu, id);
        blockById.set(id, nu);
        if (el && el.parentNode === doc) el.replaceWith(nu); else doc.appendChild(nu);
      } else if (k.startsWith('hf:')) {
        const el = HF[+k.slice(3)];
        if (el && v !== null) el.innerHTML = sanitize(String(v));
      } else if (k === 'setup' && v) {
        for (const key of Object.keys(DEFAULT_SETUP)) if (key in v) setup[key] = v[key];
        Object.assign(setup, validSetup({ ...setup }));
        applySetup();
      } else if (k === 'sources') {
        notes.loadSources(v || []);
        sourcesChanged = true;
      }
    }
    // Put the blocks in position order.
    const order = [...doc.children].filter(el => blockId.has(el) && blockPos.has(blockId.get(el)))
      .sort((a, b) => { const pa = blockPos.get(blockId.get(a)), pb = blockPos.get(blockId.get(b)); return pa < pb ? -1 : pa > pb ? 1 : (blockId.get(a) < blockId.get(b) ? -1 : 1); });
    let ref = doc.firstElementChild;
    for (const el of order) { if (el !== ref) doc.insertBefore(el, ref); else ref = ref.nextElementSibling; }
    if (!doc.firstElementChild) doc.innerHTML = '<p><br></p>';
    // The caret stays where it was in the text.
    if (caret && caret.b && blockById.get(caret.b) && blockById.get(caret.b).isConnected) {
      const moved = (id, o) => (textBefore.has(id) && blockById.get(id) ? mapOffset(textBefore.get(id), blockById.get(id).textContent, o) : o);
      const a = pointAt(blockById.get(caret.b), moved(caret.b, caret.o));
      const fbl = blockById.get(caret.fb);
      const f = fbl && fbl.isConnected ? pointAt(fbl, moved(caret.fb, caret.fo)) : a;
      try { getSelection().setBaseAndExtent(a[0], a[1], f[0], f[1]); } catch { /* ignore */ }
    }
    if (sourcesChanged) notes.refreshCitations();
    extras.renderEquations();
    changed();
    if (!commentsPane.hidden) review.renderPane();
    drawLivePeers();
  }
  const peerLayer = document.createElement('div');
  page.appendChild(peerLayer);
  function drawLivePeers() {
    peerLayer.innerHTML = '';
    const z = parseFloat(page.style.zoom) || 1;
    const pr = page.getBoundingClientRect();
    for (const p of livePeers) {
      const c = p.presence;
      const b = c && blockById.get(c.b);
      if (!b || !b.isConnected) continue;
      const pt = pointAt(b, c.o);
      const r = document.createRange();
      try { r.setStart(pt[0], pt[1]); r.collapse(true); } catch { continue; }
      let rect = r.getClientRects()[0] || b.getBoundingClientRect();
      const f = document.createElement('div');
      f.className = 'live-flag';
      Object.assign(f.style, { left: ((rect.left - pr.left) / z) + 'px', top: ((rect.top - pr.top) / z) + 'px', width: '2px', height: (Math.max(rect.height, 14) / z) + 'px', background: p.color });
      const tag = document.createElement('span');
      tag.className = 'tag';
      tag.style.background = p.color;
      tag.textContent = p.name;
      f.appendChild(tag);
      peerLayer.appendChild(f);
    }
  }
  const live = Collab.attach({
    app: 'write', doc: state,
    dirty: () => state.dirty || !!doc.innerText.trim(),
    busy: () => composing,
    state: liveState,
    apply: liveApply,
    presence: () => caretInfo(),
    textFields: ['h'], // a paragraph's HTML merges by words when two people edit it at once
    describe(k) {
      if (k.startsWith('b:')) { const el = blockById.get(k.slice(2)); const t = el && el.isConnected ? el.textContent.replace(/\s+/g, ' ').trim() : ''; return t ? `“${t.slice(0, 40)}${t.length > 40 ? '…' : ''}”` : 'a paragraph'; }
      return k.startsWith('hf:') ? 'header or footer' : k === 'setup' ? 'page setup' : k === 'sources' ? 'sources' : null;
    },
    reveal(k) { const el = k.startsWith('b:') && blockById.get(k.slice(2)); if (el && el.isConnected) { el.scrollIntoView({ block: 'center' }); el.classList.add('flash'); setTimeout(() => el.classList.remove('flash'), 1200); } },
    follow(c) { const el = c && blockById.get(c.b); if (el && el.isConnected) { const r = el.getBoundingClientRect(); if (r.top < 120 || r.bottom > innerHeight - 40) el.scrollIntoView({ block: 'center' }); } },
    showPeers(list) { livePeers = list; drawLivePeers(); }
  });
  document.addEventListener('selectionchange', () => live.moved());
  // In a live session the browser's undo would also take back other people's edits
  // (and lose track of the page); undo then takes back only your own changes.
  function liveUndo(which) {
    if (live.status !== 'live') return false;
    if (!live[which]()) UI.toast(which === 'undoLocal' ? 'Nothing of yours to undo' : 'Nothing to redo');
    return true;
  }
  [doc, ...HF].forEach(el => el.addEventListener('beforeinput', e => {
    if (e.inputType !== 'historyUndo' && e.inputType !== 'historyRedo') return;
    const redo = e.inputType === 'historyRedo';
    if ((el === doc && appUndo(redo)) || liveUndo(redo ? 'redoLocal' : 'undoLocal')) e.preventDefault();
  }));
  [doc, ...HF].forEach(el => el.addEventListener('keydown', e => {
    if (!(e.ctrlKey || e.metaKey) || e.altKey) return;
    const k = e.key.toLowerCase();
    const which = k === 'z' ? (e.shiftKey ? 'redoLocal' : 'undoLocal') : k === 'y' && !e.shiftKey ? 'redoLocal' : null;
    if (which && ((el === doc && appUndo(which === 'redoLocal')) || liveUndo(which))) e.preventDefault();
  }));
  doc.addEventListener('input', () => setTimeout(drawLivePeers, 0));

  // ---------- Accessibility Checker ----------
  function a11yGo(el) {
    return () => {
      el.scrollIntoView({ block: 'center', behavior: 'smooth' });
      const r = document.createRange();
      r.selectNodeContents(el);
      doc.focus();
      getSelection().removeAllRanges(); getSelection().addRange(r);
    };
  }
  const a11y = A11y.checker({
    check() {
      const out = [];
      const short = el => (el.textContent || el.getAttribute('alt') || '').trim().slice(0, 40) || el.tagName.toLowerCase();
      doc.querySelectorAll('img').forEach(img => {
        if (img.classList.contains('decorative') || (img.getAttribute('alt') || '').trim()) return;
        out.push({ kind: 'error', title: 'Missing alternative text', detail: 'A picture has no description. Add alt text, or mark it as decorative.', fixLabel: 'Add alt text', fix: () => altTextDialog(img), go: () => { img.scrollIntoView({ block: 'center' }); img.click(); } });
      });
      doc.querySelectorAll('table').forEach((t, i) => {
        if (t.querySelector('th')) return;
        out.push({ kind: 'warning', title: `Table ${i + 1} has no header row`, detail: 'Screen readers use the header row to describe each cell.', fixLabel: 'Use first row as header', go: a11yGo(t),
          fix: () => { const row = t.rows[0]; if (!row) return; [...row.cells].forEach(td => { const th = document.createElement('th'); th.innerHTML = td.innerHTML; th.setAttribute('scope', 'col'); [...td.attributes].forEach(a => { if (a.name !== 'scope') th.setAttribute(a.name, a.value); }); td.replaceWith(th); }); changed(); } });
      });
      let last = 0;
      doc.querySelectorAll('h1, h2, h3, h4, h5, h6').forEach(h => {
        const lv = +h.tagName[1];
        if (!h.textContent.trim()) out.push({ kind: 'warning', title: 'Empty heading', detail: 'A heading with no text is announced but says nothing.', go: a11yGo(h), fixLabel: 'Make it a paragraph', fix: () => { const p = document.createElement('p'); p.innerHTML = h.innerHTML || '<br>'; h.replaceWith(p); changed(); } });
        else if (last && lv > last + 1) { const want = last + 1; out.push({ kind: 'warning', title: `Heading level skipped: “${short(h)}”`, detail: `Heading ${lv} follows heading ${last}. Use heading ${want} so the outline makes sense.`, go: a11yGo(h), fixLabel: `Make it Heading ${want}`, fix: () => { const n = document.createElement('h' + want); n.innerHTML = h.innerHTML; h.replaceWith(n); changed(); } }); }
        if (h.textContent.trim()) last = lv;
      });
      // Text colour against the page colour.
      const paper = setup.color || '#ffffff';
      const seen = new Set();
      doc.querySelectorAll('[style*="color"], font[color]').forEach(el => {
        const c = el.style.color || el.getAttribute('color');
        if (!c || seen.has(c) || !el.textContent.trim()) return;
        const bg = el.style.backgroundColor || paper;
        const ratio = A11y.contrast(c, bg.startsWith('rgb') || bg.startsWith('#') ? bg : paper);
        if (ratio !== null && ratio < 4.5) {
          seen.add(c);
          out.push({ kind: 'warning', title: `Hard-to-read text colour: “${short(el)}”`, detail: `Contrast ${ratio.toFixed(1)}:1; at least 4.5:1 is needed for normal text.`, go: a11yGo(el), fixLabel: 'Use automatic colour',
            fix: () => { doc.querySelectorAll('[style*="color"], font[color]').forEach(x => { if ((x.style.color || x.getAttribute('color')) === c) { x.style.color = ''; x.removeAttribute('color'); } }); changed(); } });
        }
      });
      doc.querySelectorAll('a[href]').forEach(a => {
        const t = a.textContent.trim().toLowerCase();
        if (/^(click here|here|link|this|more|read more|this link)$/.test(t) || /^https?:\/\/\S{40,}$/.test(t)) out.push({ kind: 'tip', title: `Unclear link text: “${a.textContent.trim().slice(0, 40)}”`, detail: 'Link text should say where the link goes.', go: a11yGo(a) });
      });
      const words = (doc.innerText.match(/\S+/g) || []).length;
      if (words > 400 && !doc.querySelector('h1, h2, h3')) out.push({ kind: 'tip', title: 'No headings', detail: 'Long documents are easier to navigate with headings (Home → Styles).' });
      let blanks = 0, flagged = false;
      for (const p of doc.children) {
        if (p.tagName === 'P' && !p.textContent.trim() && !p.querySelector('img, br + br')) blanks++;
        else blanks = 0;
        if (blanks === 3 && !flagged) { flagged = true; out.push({ kind: 'tip', title: 'Blank lines used for spacing', detail: 'Several empty paragraphs in a row are each read out as “blank”. Use paragraph spacing or a page break instead.', go: a11yGo(p) }); }
      }
      return out;
    }
  });
  document.getElementById('a11yCheck').addEventListener('click', () => a11y.open());

  // ---------- Proofing ----------
  const proofing = Proofing.attach({
    app: 'write',
    roots: () => [{ el: doc, live: true }],
    lang: () => setup.lang || '',
    setLang: code => { setup.lang = code; applySetup(); state.markDirty(); }
  });
  document.getElementById('proofEditor').addEventListener('click', () => proofing.open());
  document.getElementById('proofLang').addEventListener('click', () => proofing.languageDialog());
  document.getElementById('stLang').addEventListener('click', () => proofing.languageDialog());

  // ---------- Assistant ----------
  /** The section around a node: its heading and everything up to the next heading as high or higher. */
  function sectionAround(node) {
    let start = topBlock(node) || doc.firstElementChild;
    while (start && !/^H[1-6]$/.test(start.tagName) && start.previousElementSibling) start = start.previousElementSibling;
    if (!start) return [];
    const level = /^H[1-6]$/.test(start.tagName) ? +start.tagName[1] : 7;
    const blocks = [start];
    for (let n = start.nextElementSibling; n && !(/^H[1-6]$/.test(n.tagName) && +n.tagName[1] <= level); n = n.nextElementSibling) blocks.push(n);
    return blocks;
  }
  const blocksMarkdown = blocks => { const d = document.createElement('div'); blocks.forEach(b => d.appendChild(b.cloneNode(true))); return htmlToMarkdown(d); };
  function aiContext(action) {
    const r = savedRange && doc.contains(savedRange.commonAncestorContainer) ? savedRange.cloneRange() : null;
    const document_ = htmlToMarkdown(doc.cloneNode(true));
    // "This section": the heading the cursor is under, down to the next heading at its level.
    if (action && action.scope === 'section') {
      const blocks = r ? sectionAround(r.startContainer) : [];
      return { document: document_, selection: blocks.length ? blocksMarkdown(blocks) : '', hasSelection: blocks.length > 0, blocks };
    }
    const selection = r && !r.collapsed ? r.toString() : '';
    // A selection across paragraphs is rewritten paragraph by paragraph.
    let blocks = null;
    if (r && !r.collapsed) {
      const a = topBlock(r.startContainer), b = topBlock(r.endContainer);
      if (a && b && a !== b) { blocks = []; for (let n = a; n; n = n.nextElementSibling) { blocks.push(n); if (n === b) break; } }
    }
    return { document: document_, selection: blocks ? blocksMarkdown(blocks) : selection, hasSelection: !!selection.trim(), range: r, blocks };
  }
  /** Puts a rewrite in place of whole paragraphs (a section), as one step Ctrl+Z takes back. */
  function aiReplaceBlocks(text, blocks) {
    if (!blocks.every(b => doc.contains(b))) return UI.toast('The text that was rewritten has changed; ask again', 4000);
    review.replaceBlocks(blocks, sanitize(markdownToHtml(text.trim())));
    changed();
  }
  /** Puts a rewrite in place of the selection or section as tracked changes, for the author to accept or reject. */
  function aiTracked(text, ctx) {
    const html = sanitize(markdownToHtml(text.trim()));
    let done;
    if (ctx.blocks && ctx.blocks.length && ctx.blocks.every(b => doc.contains(b))) done = review.replaceBlocksTracked(ctx.blocks, html);
    else if (ctx.range && doc.contains(ctx.range.startContainer)) {
      const t = document.createElement('template');
      t.innerHTML = html;
      const kids = [...t.content.children];
      if (kids.length <= 1) done = review.replaceTextTracked(ctx.range, (kids[0] || t.content).textContent);
      else done = review.replaceBlocksTracked([topBlock(ctx.range.startContainer)], html);
    } else return UI.toast('The text that was rewritten has changed; ask again', 4000);
    if (!done) return UI.toast('No changes to suggest');
    if (!commentsPane.hidden) review.renderPane();
    UI.toast(`Suggested as tracked changes (${review.count()}). Review → Accept or Reject.`, 4000);
  }
  /** Puts a reply (Markdown) into the document at the range, or in place of it. */
  function aiInsert(text, ctx, replace) {
    let r = ctx && ctx.range && doc.contains(ctx.range.startContainer) ? ctx.range.cloneRange() : null;
    if (!r) { r = document.createRange(); r.selectNodeContents(doc); r.collapse(false); }
    if (!replace) r.collapse(false);
    const t = document.createElement('template');
    t.innerHTML = sanitize(markdownToHtml(text.trim()));
    const kids = [...t.content.children];
    doc.focus();
    const sel = getSelection();
    sel.removeAllRanges();
    sel.addRange(r);
    if (kids.length === 1 && kids[0].tagName === 'P') {
      Platform.insertInline(kids[0].innerHTML);
    } else {
      // Several paragraphs, a list or a table go after the paragraph the cursor is in.
      if (replace) r.deleteContents();
      const block = topBlock(r.endContainer) || doc.lastElementChild;
      const frag = t.content;
      const last = frag.lastChild;
      if (block) block.after(frag); else doc.appendChild(frag);
      if (block && !block.textContent.trim() && !block.querySelector('img,table,hr')) block.remove();
      const after = document.createRange();
      after.setStartAfter(last); after.collapse(true);
      sel.removeAllRanges(); sel.addRange(after);
      changed();
    }
    UI.toast(replace ? 'Replaced' : 'Inserted');
  }
  // Dictate: what is said goes in at the cursor as if typed (tracked when tracking is on).
  Assistant.dictateButton(document.getElementById('dictate'), text => {
    if (!doc.contains(getSelection().anchorNode)) restoreSelection();
    doc.focus();
    text = Assistant.spaced(doc, text);
    if (review.tracking) review.insertText(text); else document.execCommand('insertText', false, text);
  });
  Assistant.attach({
    app: 'write',
    context: aiContext,
    format: 'In Isadek Write you may use Markdown (# headings, **bold**, *italic*, - and 1. lists, tables); it becomes formatting when inserted.',
    actions: [
      { label: 'Summarize', prompt: 'Summarize this document in a short paragraph followed by its key points as a bulleted list.' },
      { label: 'Improve writing', prompt: 'Rewrite the selected text so it reads more clearly and naturally. Keep its meaning, tone and language.', selection: true },
      { label: 'Fix spelling & grammar', prompt: 'Correct the spelling, grammar and punctuation of the selected text. Change nothing else.', selection: true },
      { label: 'Make shorter', prompt: 'Rewrite the selected text to about half its length, keeping the essential points.', selection: true },
      { label: 'Make more formal', prompt: 'Rewrite the selected text in a more formal, professional tone.', selection: true },
      { label: 'Rewrite this section', scope: 'section', selection: true, prompt: 'Rewrite the section in <selection> (the part of the document the cursor is in) so it reads clearly and concisely. Keep its headings, facts and language, and keep one paragraph for each paragraph unless a paragraph is better split or merged. Reply with the whole rewritten section in Markdown.' },
      { label: 'Continue writing', prompt: 'Continue the document from where it ends with one or two paragraphs in the same style. Reply with only the new text.' },
      { label: 'Suggest a title', prompt: 'Suggest a short, fitting title for this document. Reply with only the title.' },
      { label: 'Make a presentation', prompt: 'Turn this document into a presentation outline: "# " for the title slide (with a one-line subtitle under it) and for each major part, "## " for each slide with 3 to 5 short "- " points taken from the document, and a "Notes:" line of speaker notes for each slide. Reply with only the outline.' }
    ],
    buttons: [
      { label: 'Suggest as tracked changes', selection: true, run: (text, ctx) => aiTracked(text, ctx) },
      { label: 'Replace selection', selection: true, run: (text, ctx) => (ctx.blocks && !ctx.range ? aiReplaceBlocks(text, ctx.blocks) : aiInsert(text, ctx, true)) },
      { label: 'Insert', run: text => aiInsert(text, aiContext(), false) },
      { label: 'Open as slides', run: text => (/^#{1,3}\s/m.test(text) ? SuiteTemplates.openOutline(state.name || 'Presentation', text) : UI.toast('This reply isn\'t an outline; ask for a presentation first', 4000)) }
    ]
  });

  // The start screen's thumbnail of this document: its first blocks as plain text.
  // The file's thumbnail (Finder, Quick Look, File Explorer): the top of the first page.
  Platform.previewElement = () => { const pg = document.getElementById('page'); if (pg) pg.dataset.previewAspect = '1.294'; return pg; };
  Platform.preview = () => {
    const blocks = [], MAX = 18, clip = t => t.replace(/\s+/g, ' ').trim().slice(0, 160);
    for (const el of document.getElementById('doc').children) {
      if (blocks.length >= MAX) break;
      const tag = el.tagName.toLowerCase();
      if (tag === 'ul' || tag === 'ol') { for (const li of el.children) if (blocks.length < MAX) blocks.push({ t: 'li', s: clip(li.textContent) }); continue; }
      if (tag === 'table') { blocks.push({ t: 'table', r: Math.min(el.rows.length, 6), c: Math.min((el.rows[0] && el.rows[0].cells.length) || 1, 6) }); continue; }
      if (tag === 'hr') { blocks.push({ t: 'hr' }); continue; }
      const text = clip(el.textContent);
      if (!text && el.querySelector('img, svg, canvas')) { blocks.push({ t: 'img' }); continue; }
      const t = el.classList.contains('title') ? 'title' : /^h[1-3]$/.test(tag) ? tag : 'p';
      const align = getComputedStyle(el).textAlign;
      blocks.push({ t, s: text, a: align === 'center' || align === 'right' ? align : undefined });
    }
    return { kind: 'write', blocks };
  };

  window.AtlasWrite = { loadFile, serialize, markdownToHtml, htmlToMarkdown, sanitize, get live() { return live; }, get notes() { return notes; } };
})();
