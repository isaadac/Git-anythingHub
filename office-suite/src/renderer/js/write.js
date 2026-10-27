'use strict';
// Atlas Write — word processor.
(function () {
  const doc = document.getElementById('doc');
  const page = document.getElementById('page');
  const header = document.getElementById('pageHeader');
  const footer = document.getElementById('pageFooter');
  const firstHeader = document.getElementById('firstHeader');
  const firstFooter = document.getElementById('firstFooter');
  const HF = [header, footer, firstHeader, firstFooter];
  const state = UI.docState('Atlas Write', 'Document1');
  const FILTERS = [
    { name: 'Word Document', extensions: ['docx'] },
    { name: 'Web Page', extensions: ['html', 'htm'] },
    { name: 'Markdown', extensions: ['md'] },
    { name: 'Plain Text', extensions: ['txt'] }
  ];
  const OPEN_FILTERS = [{ name: 'Documents', extensions: ['docx', 'html', 'htm', 'md', 'txt'] }, ...FILTERS];
  const PAGE_SIZES = { letter: [8.5, 11], a4: [8.27, 11.69], legal: [8.5, 14] }; // inches
  const DEFAULT_SETUP = { margin: '1in', orientation: 'portrait', size: 'letter', columns: 1, color: '#ffffff', differentFirst: false, trackChanges: false, citeStyle: 'apa' };
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
    if (['apa', 'mla', 'chicago'].includes(v.citeStyle)) out.citeStyle = v.citeStyle;
    return out;
  }

  UI.initRibbon();
  document.execCommand('styleWithCSS', false, true);
  document.execCommand('defaultParagraphSeparator', false, 'p');

  // ---------- Selection helpers ----------
  let savedRange = null;
  const editable = () => (HF.includes(document.activeElement) ? document.activeElement : doc);
  document.addEventListener('selectionchange', () => {
    const sel = getSelection();
    if (sel.rangeCount && (doc.contains(sel.anchorNode) || HF.some(h => h.contains(sel.anchorNode)))) {
      savedRange = sel.getRangeAt(0).cloneRange();
      updateToolbarState();
      updateCaretPage();
    }
  });
  function restoreSelection() {
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
  document.getElementById('insToc').addEventListener('click', () => {
    const heads = [...doc.querySelectorAll('h1:not(.title),h2,h3')];
    if (!heads.length) return UI.toast('Add some headings first (Home → Styles)');
    heads.forEach((h, i) => { if (!h.id) h.id = 'h-' + Date.now().toString(36) + i; });
    const items = heads.map(h => `<p style="margin-left:${(+h.tagName[1] - 1) * 18}pt"><a href="#${h.id}">${escapeHtml(h.textContent)}</a></p>`).join('');
    doc.querySelector('.toc')?.remove();
    doc.insertAdjacentHTML('afterbegin', `<div class="toc" contenteditable="false"><p><b>Contents</b></p>${items}</div>`);
    changed();
  });
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
  function fixFields(root) { root.querySelectorAll('.fld').forEach(f => { f.contentEditable = 'false'; if (!['page', 'numpages'].includes(f.dataset.fld)) f.dataset.fld = 'page'; }); }
  document.getElementById('insPageNum').addEventListener('change', e => {
    const v = e.target.value;
    e.target.value = '';
    if (!v) return;
    if (v === 'remove') {
      HF.forEach(el => el.querySelectorAll('.fld').forEach(f => { const p = f.closest('p'); f.remove(); if (p && !p.textContent.replace(/(Page|of)/g, '').trim()) p.remove(); }));
      layout.run(); changed(); return;
    }
    const [where, align, kind] = v.split(':');
    if (where === 'here') { exec('insertHTML', fieldHtml(kind) + '\u200b'); fixFields(editable()); layout.run(); return; }
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
    if (b.dataset.a) {
      selImg.style.display = 'block';
      selImg.style.marginLeft = b.dataset.a === 'left' ? '0' : 'auto';
      selImg.style.marginRight = b.dataset.a === 'right' ? '0' : 'auto';
    }
    if (b.dataset.del) { selImg.remove(); selImg = null; imgTools.hidden = true; }
    changed();
  });

  // ---------- Ribbon: Layout ----------
  function applySetup() {
    const [w, h] = PAGE_SIZES[setup.size];
    const land = setup.orientation === 'landscape';
    page.style.setProperty('--page-w', (land ? h : w) + 'in');
    page.style.setProperty('--page-h', (land ? w : h) + 'in');
    page.style.setProperty('--margin', setup.margin);
    page.style.setProperty('--paper', setup.color);
    doc.style.columnCount = setup.columns;
    doc.style.columnGap = '0.5in';
    let style = document.getElementById('print-page-style');
    if (!style) { style = document.createElement('style'); style.id = 'print-page-style'; document.head.appendChild(style); }
    // Margins are part of the laid-out sheets, so the printer page has none of its own.
    style.textContent = `@page { size: ${land ? h : w}in ${land ? w : h}in; margin: 0; }`;
    document.getElementById('differentFirst').checked = !!setup.differentFirst;
    if (review.tracking !== !!setup.trackChanges) setTracking(setup.trackChanges);
    document.getElementById('citeStyle').value = setup.citeStyle;
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
  const review = WriteReview({ doc, author: getAuthor, changed: () => { changed(); if (!commentsPane.hidden) review.renderPane(); }, pane: commentsPane });
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
    z = Math.max(50, Math.min(200, Math.round(z)));
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
    onLayout: () => updateCaretPage(),
    footnotes: () => notes.footnotes()
  });
  // ---------- References: notes, citations, bibliography ----------
  const notes = WriteNotes({ doc, setup: () => setup, restoreSelection, changed: () => changed() });
  document.getElementById('insFootnote').addEventListener('click', () => notes.insertNote('foot'));
  document.getElementById('insEndnote').addEventListener('click', () => notes.insertNote('end'));
  document.getElementById('insCitation').addEventListener('click', () => notes.insertCitation());
  document.getElementById('insBibliography').addEventListener('click', () => notes.insertBibliography());
  document.getElementById('manageSources').addEventListener('click', () => notes.manageSources());
  document.getElementById('citeStyle').addEventListener('change', e => { setup.citeStyle = e.target.value; notes.refreshCitations(); changed(); });
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
  function updateStatus() {
    const words = (doc.innerText.match(/\S+/g) || []).length;
    document.getElementById('stWords').textContent = `${words.toLocaleString()} word${words === 1 ? '' : 's'}`;
    updateCaretPage();
  }
  let navTimer;
  function changed() {
    state.markDirty();
    notes.renumber();
    layout.schedule();
    updateStatus();
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

  async function loadFile(f) {
    const ext = Platform.extOf(f.name);
    let html;
    let nextSetup = { ...DEFAULT_SETUP };
    let hfHtml = { header: '', footer: '', firstHeader: '', firstFooter: '' };
    let nextSources = [];
    try {
      if (ext === 'docx') {
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
        const src = pre.bytes;
        const res = await mammoth.convertToHtml({ arrayBuffer: src.buffer.slice(src.byteOffset, src.byteOffset + src.byteLength) }, {
          styleMap: ['p[style-name=\'Title\'] => h1.title:fresh', 'p[style-name=\'Quote\'] => blockquote:fresh', 'p[style-name=\'Intense Quote\'] => blockquote:fresh']
        });
        html = notes.importMammothNotes(review.postprocessHtml(res.value, pre.meta));
        if (res.messages.some(m => m.type === 'warning')) console.info('Import notes:', res.messages);
      } else if (ext === 'html' || ext === 'htm') {
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
    } catch (err) {
      UI.toast('Could not open file: ' + err.message, 4000);
      return;
    }
    // Nothing from the previous document (page setup, header, footer) carries over.
    Object.assign(setup, nextSetup);
    for (const [key, el] of Object.entries({ header, footer, firstHeader, firstFooter })) { el.innerHTML = sanitize(hfHtml[key]); fixFields(el); }
    doc.innerHTML = sanitize(html) || '<p><br></p>';
    fixFields(doc);
    notes.loadSources(nextSources);
    notes.renumber();
    notes.refreshCitations();
    savedRange = null;
    applySetup();
    syncSetupControls();
    state.set({ name: f.name, path: f.path, dirty: false });
    if (!f.recovered) Platform.rememberRecent(f.name, 'write');
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
    const opts = { stylesheet: DOCX_STYLESHEET, maxImageWidth: Math.round(textWidth / setup.columns * 96) };
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
    notes.prepareForDocx(body);
    section.children = [...docxHtml.htmlToContent(body.innerHTML, { ...opts, trackChanges: { author: getAuthor() } })];
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
      creator: 'Atlas Write',
      title: Platform.baseName(state.name),
      styles: { default: { document: { run: { font: 'Calibri', size: 22 }, paragraph: { spacing: { after: 160, line: 276 } } } } },
      background: setup.color !== '#ffffff' ? { color: setup.color.slice(1) } : undefined,
      features: setup.trackChanges ? { trackRevisions: true } : undefined,
      customProperties: [{ name: 'AtlasCiteStyle', value: setup.citeStyle }, ...(notes.sources.length ? [{ name: 'AtlasSources', value: JSON.stringify(notes.sources) }] : [])],
      comments: cmts.length ? { children: cmts.map(c => ({ id: c.id, author: c.author, initials: c.author.split(/\s+/).map(w => w[0]).join('').slice(0, 3), date: new Date(c.date), children: c.text.split('\n').map(t => new D.Paragraph(t)) })) } : undefined,
      sections: [section]
    });
    const blob = await D.Packer.toBlob(document_);
    return review.patchDocx(new Uint8Array(await blob.arrayBuffer()));
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
  function toHtml() {
    const css = [...document.styleSheets].find(s => s.href && s.href.endsWith('write.css'));
    let rules = '';
    try { rules = [...css.cssRules].map(r => r.cssText).filter(t => t.startsWith('.doc')).join('\n').replace(/\.doc\b/g, 'main.atlas-doc'); } catch { /* ignore */ }
    return `<!doctype html>
<html><head><meta charset="utf-8"><title>${escapeHtml(Platform.baseName(state.name))}</title>
<meta name="generator" content="Atlas Write">
<meta name="atlas-setup" content="${escapeHtml(JSON.stringify(setup))}">
${notes.sources.length ? `<meta name="atlas-sources" content="${escapeHtml(JSON.stringify(notes.sources))}">` : ''}
<style>body{max-width:8.5in;margin:auto;padding:1in;font-family:Calibri,Arial,sans-serif;font-size:11pt;line-height:1.15;background:${setup.color}}
${rules}</style></head>
<body>${hfOut(firstHeader, 'header', 'atlas-first-header')}${hfOut(header, 'header', 'atlas-header')}<main class="atlas-doc">${cleanHtmlFor(doc)}</main>${hfOut(footer, 'footer', 'atlas-footer')}${hfOut(firstFooter, 'footer', 'atlas-first-footer')}</body></html>`;
  }
  function hfOut(el, tag, cls) {
    return el.textContent.trim() || el.querySelector('img') ? `<${tag} class="${cls}">${cleanHtmlFor(el)}</${tag}>` : '';
  }
  async function serialize(ext) {
    if (ext === 'docx') return toDocx();
    if (ext === 'html' || ext === 'htm') return toHtml();
    if (ext === 'md') {
      const clone = doc.cloneNode(true);
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
  Platform.onMenu(async cmd => {
    switch (cmd) {
      case 'new': Platform.newWindow('write'); break;
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
      case 'save-and-close': if (await save(false)) Platform.closeNow(); break;
      case 'export-pdf': { const restore = layout.forPrint(); try { await Platform.exportPdf(Platform.baseName(state.name) + '.pdf', setup.orientation === 'landscape'); } finally { restore(); } break; }
      case 'export-html': save(true, 'html'); break;
      case 'export-txt': save(true, 'txt'); break;
      case 'print': { const restore = layout.forPrint(); try { await Platform.print(); } finally { restore(); } break; }
      case 'undo': exec('undo'); break;
      case 'redo': exec('redo'); break;
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
    canAutoSave: () => !Platform.desktop ? 'AutoSave needs the desktop app' : !state.path ? 'Save the document first to turn on AutoSave'
      : /\.(docx|html?)$/i.test(state.path) ? true : 'AutoSave works with .docx and .html files'
  });
  async function loadRecovered(d) {
    await loadFile({ name: 'recovered.html', path: null, data: new TextEncoder().encode(d.data), recovered: true });
    state.set({ name: d.name, path: d.path, dirty: true });
    recovery.adopt(d.draftId);
    UI.toast('Recovered your unsaved changes', 4000);
  }

  window.AtlasWrite = { loadFile, serialize, markdownToHtml, htmlToMarkdown, sanitize };
})();
