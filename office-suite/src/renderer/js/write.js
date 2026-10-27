'use strict';
// Atlas Write — word processor.
(function () {
  const doc = document.getElementById('doc');
  const page = document.getElementById('page');
  const header = document.getElementById('pageHeader');
  const footer = document.getElementById('pageFooter');
  const state = UI.docState('Atlas Write', 'Document1');
  const FILTERS = [
    { name: 'Word Document', extensions: ['docx'] },
    { name: 'Web Page', extensions: ['html', 'htm'] },
    { name: 'Markdown', extensions: ['md'] },
    { name: 'Plain Text', extensions: ['txt'] }
  ];
  const OPEN_FILTERS = [{ name: 'Documents', extensions: ['docx', 'html', 'htm', 'md', 'txt'] }, ...FILTERS];
  const PAGE_SIZES = { letter: [8.5, 11], a4: [8.27, 11.69], legal: [8.5, 14] }; // inches
  const setup = { margin: '1in', orientation: 'portrait', size: 'letter', columns: 1, color: '#ffffff' };

  UI.initRibbon();
  document.execCommand('styleWithCSS', false, true);
  document.execCommand('defaultParagraphSeparator', false, 'p');

  // ---------- Selection helpers ----------
  let savedRange = null;
  const editable = () => (document.activeElement === header || document.activeElement === footer) ? document.activeElement : doc;
  document.addEventListener('selectionchange', () => {
    const sel = getSelection();
    if (sel.rangeCount && (doc.contains(sel.anchorNode) || header.contains(sel.anchorNode) || footer.contains(sel.anchorNode))) {
      savedRange = sel.getRangeAt(0).cloneRange();
      updateToolbarState();
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
  document.getElementById('editHeader').addEventListener('click', () => { header.hidden = false; header.focus(); });
  document.getElementById('editFooter').addEventListener('click', () => { footer.hidden = false; footer.focus(); });

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
    page.style.background = setup.color;
    doc.style.columnCount = setup.columns;
    doc.style.columnGap = '0.5in';
    let style = document.getElementById('print-page-style');
    if (!style) { style = document.createElement('style'); style.id = 'print-page-style'; document.head.appendChild(style); }
    style.textContent = `@page { size: ${land ? h : w}in ${land ? w : h}in; margin: ${setup.margin}; }`;
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
  document.getElementById('addComment').addEventListener('click', async () => {
    restoreSelection();
    const sel = getSelection();
    if (sel.isCollapsed) return UI.toast('Select some text to comment on');
    const range = sel.getRangeAt(0).cloneRange();
    const c = await UI.prompt('New Comment', 'Comment');
    if (!c) return;
    const span = document.createElement('span');
    span.className = 'comment';
    span.title = c;
    try { range.surroundContents(span); } catch { span.appendChild(range.extractContents()); range.insertNode(span); }
    changed();
  });
  document.getElementById('delComment').addEventListener('click', () => {
    restoreSelection();
    const c = closest('.comment');
    if (!c) return UI.toast('Place the cursor in a comment');
    c.replaceWith(...c.childNodes);
    changed();
  });

  // ---------- Ribbon: View ----------
  document.getElementById('viewPrint').addEventListener('click', () => document.body.classList.remove('view-web'));
  document.getElementById('viewWeb').addEventListener('click', () => document.body.classList.add('view-web'));
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
  function findNext(back) {
    if (!findText.value) return false;
    doc.focus();
    // window.find searches the whole page; restrict hits to the document body.
    for (let i = 0; i < 2; i++) {
      const ok = window.find(findText.value, false, back, true, false, false, false);
      if (ok && doc.contains(getSelection().anchorNode)) return true;
      if (!ok) break;
    }
    UI.toast('No more matches');
    return false;
  }
  findText.addEventListener('keydown', e => { if (e.key === 'Enter') findNext(e.shiftKey); if (e.key === 'Escape') findbar.classList.remove('open'); });
  document.getElementById('findNext').addEventListener('click', () => findNext(false));
  document.getElementById('findPrev').addEventListener('click', () => findNext(true));
  document.getElementById('findClose').addEventListener('click', () => findbar.classList.remove('open'));
  document.getElementById('replaceOne').addEventListener('click', () => {
    const sel = getSelection().toString();
    if (sel.toLowerCase() === findText.value.toLowerCase() && doc.contains(getSelection().anchorNode)) {
      document.execCommand('insertText', false, replaceText.value);
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
  function pageCount() {
    const [w, h] = PAGE_SIZES[setup.size];
    const pageH = (setup.orientation === 'landscape' ? w : h) * 96;
    const margin = parseFloat(setup.margin) * 96;
    const breaks = doc.querySelectorAll('.page-break').length;
    return Math.max(1 + breaks, Math.ceil(doc.scrollHeight / (pageH - 2 * margin)));
  }
  function updateStatus() {
    const words = (doc.innerText.match(/\S+/g) || []).length;
    document.getElementById('stWords').textContent = `${words.toLocaleString()} word${words === 1 ? '' : 's'}`;
    document.getElementById('stPages').textContent = `${pageCount()} page${pageCount() === 1 ? '' : 's'}`;
  }
  let navTimer;
  function changed() {
    state.markDirty();
    updateStatus();
    clearTimeout(navTimer);
    navTimer = setTimeout(refreshNav, 400);
  }
  [doc, header, footer].forEach(el => el.addEventListener('input', changed));

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
      } else document.execCommand('insertText', false, '    ');
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

  // Paste: keep formatting but strip anything unsafe.
  doc.addEventListener('paste', e => {
    const html = e.clipboardData.getData('text/html');
    if (!html) return;
    e.preventDefault();
    document.execCommand('insertHTML', false, sanitize(html));
    changed();
  });

  // ---------- Import ----------
  function escapeHtml(s) { return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
  /** Removes scripts, event handlers and javascript: URLs from untrusted HTML. */
  function sanitize(html) {
    const d = new DOMParser().parseFromString(html, 'text/html');
    d.querySelectorAll('script,style,iframe,object,embed,link,meta,form,input,button,textarea,select,base,frame,frameset,noscript,template').forEach(n => n.remove());
    d.querySelectorAll('*').forEach(el => {
      [...el.attributes].forEach(a => {
        const n = a.name.toLowerCase();
        if (n.startsWith('on') || n === 'contenteditable' || n === 'srcdoc') el.removeAttribute(a.name);
        if ((n === 'href' || n === 'src' || n === 'xlink:href') && /^\s*(javascript|vbscript|data:text\/html)/i.test(a.value)) el.removeAttribute(a.name);
      });
    });
    return d.body.innerHTML;
  }
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
      if (n.nodeType === 3) return n.nodeValue.replace(/ /g, ' ');
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
    try {
      if (ext === 'docx') {
        const res = await mammoth.convertToHtml({ arrayBuffer: f.data.buffer.slice(f.data.byteOffset, f.data.byteOffset + f.data.byteLength) }, {
          styleMap: ['p[style-name=\'Title\'] => h1.title:fresh', 'p[style-name=\'Quote\'] => blockquote:fresh', 'p[style-name=\'Intense Quote\'] => blockquote:fresh']
        });
        html = res.value;
        if (res.messages.some(m => m.type === 'warning')) console.info('Import notes:', res.messages);
      } else if (ext === 'html' || ext === 'htm') {
        const d = new DOMParser().parseFromString(Platform.bytesToText(f.data), 'text/html');
        const saved = d.querySelector('meta[name=atlas-setup]');
        if (saved) { try { Object.assign(setup, JSON.parse(saved.content)); } catch { /* ignore */ } }
        const h = d.querySelector('header.atlas-header'), ft = d.querySelector('footer.atlas-footer');
        header.innerHTML = h ? sanitize(h.innerHTML) : ''; header.hidden = !h;
        footer.innerHTML = ft ? sanitize(ft.innerHTML) : ''; footer.hidden = !ft;
        h?.remove(); ft?.remove();
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
    doc.innerHTML = sanitize(html) || '<p><br></p>';
    applySetup();
    syncSetupControls();
    state.set({ name: f.name, path: f.path, dirty: false });
    Platform.rememberRecent(f.name, 'write');
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
    .comment { background: #fff3c4; }
    .textbox { border: 1px solid #555555; }
    td, th { border: 1px solid #8e8e8e; }
  `;
  function cleanHtmlFor(el) {
    const clone = el.cloneNode(true);
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
    const opts = { stylesheet: DOCX_STYLESHEET, maxImageWidth: Math.round((w - 2 * parseFloat(setup.margin)) * 96) };
    const section = {
      properties: {
        page: {
          size: { width: twip(w), height: twip(h), orientation: setup.orientation === 'landscape' ? D.PageOrientation.LANDSCAPE : D.PageOrientation.PORTRAIT },
          margin: { top: margin, right: margin, bottom: margin, left: margin }
        },
        column: setup.columns > 1 ? { count: setup.columns, space: 720 } : undefined
      },
      children: [...docxHtml.htmlToContent(cleanHtmlFor(doc), opts)]
    };
    if (header.innerText.trim()) section.headers = { default: new D.Header({ children: [...docxHtml.htmlToContent(header.innerHTML, opts)] }) };
    if (footer.innerText.trim()) section.footers = { default: new D.Footer({ children: [...docxHtml.htmlToContent(footer.innerHTML, opts)] }) };
    const document_ = new D.Document({
      creator: 'Atlas Write',
      title: Platform.baseName(state.name),
      styles: { default: { document: { run: { font: 'Calibri', size: 22 }, paragraph: { spacing: { after: 160, line: 276 } } } } },
      background: setup.color !== '#ffffff' ? { color: setup.color.slice(1) } : undefined,
      sections: [section]
    });
    const blob = await D.Packer.toBlob(document_);
    return new Uint8Array(await blob.arrayBuffer());
  }
  function toHtml() {
    const css = [...document.styleSheets].find(s => s.href && s.href.endsWith('write.css'));
    let rules = '';
    try { rules = [...css.cssRules].map(r => r.cssText).filter(t => t.startsWith('.doc')).join('\n').replace(/\.doc\b/g, 'main.atlas-doc'); } catch { /* ignore */ }
    return `<!doctype html>
<html><head><meta charset="utf-8"><title>${escapeHtml(Platform.baseName(state.name))}</title>
<meta name="generator" content="Atlas Write">
<meta name="atlas-setup" content="${escapeHtml(JSON.stringify(setup))}">
<style>body{max-width:8.5in;margin:auto;padding:1in;font-family:Calibri,Arial,sans-serif;font-size:11pt;line-height:1.15;background:${setup.color}}
${rules}</style></head>
<body>${header.innerText.trim() ? `<header class="atlas-header">${header.innerHTML}</header>` : ''}<main class="atlas-doc">${cleanHtmlFor(doc)}</main>${footer.innerText.trim() ? `<footer class="atlas-footer">${footer.innerHTML}</footer>` : ''}</body></html>`;
  }
  async function serialize(ext) {
    if (ext === 'docx') return toDocx();
    if (ext === 'html' || ext === 'htm') return toHtml();
    if (ext === 'md') return htmlToMarkdown(doc);
    return doc.innerText;
  }

  async function save(saveAs, forcedExt) {
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
      UI.toast('Saved');
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
      case 'export-pdf': Platform.exportPdf(Platform.baseName(state.name) + '.pdf', setup.orientation === 'landscape'); break;
      case 'export-html': save(true, 'html'); break;
      case 'export-txt': save(true, 'txt'); break;
      case 'print': Platform.print(); break;
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
    else if (f) await loadFile(f);
    updateStatus();
    doc.focus();
  })();

  window.AtlasWrite = { loadFile, serialize, markdownToHtml, htmlToMarkdown, sanitize };
})();
