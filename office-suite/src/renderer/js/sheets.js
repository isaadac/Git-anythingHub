'use strict';
// Atlas Sheets — spreadsheet.
(function () {
  const F = window.Formula;
  const { addr, indexToCol, parseRef } = F;
  const wrap = document.getElementById('gridWrap');
  const grid = document.getElementById('grid');
  const thead = grid.querySelector('thead');
  const tbody = grid.querySelector('tbody');
  const selBox = document.getElementById('selBox');
  const editor = document.getElementById('cellEditor');
  const fx = document.getElementById('fx');
  const nameBox = document.getElementById('nameBox');
  const state = UI.docState('Atlas Sheets', 'Book1');
  const DEFAULT_COL_W = 80;
  const SAVE_FILTERS = [
    { name: 'Excel Workbook', extensions: ['xlsx'] },
    { name: 'OpenDocument Spreadsheet', extensions: ['ods'] },
    { name: 'CSV (current sheet)', extensions: ['csv'] },
    { name: 'Excel 97-2003 Workbook', extensions: ['xls'] }
  ];
  const OPEN_FILTERS = [{ name: 'Spreadsheets', extensions: ['xlsx', 'xls', 'xlsm', 'ods', 'csv', 'tsv'] }];

  UI.initRibbon();

  // ---------- Model ----------
  const newSheet = name => ({ name, cells: {}, rows: 100, cols: 26, colWidths: {}, merges: [], charts: [], hiddenRows: [] });
  let book = { sheets: [newSheet('Sheet1')], active: 0 };
  const sheet = () => book.sheets[book.active];
  const cellAt = (r, c, s = sheet()) => s.cells[addr(r, c)];
  const rawAt = (r, c, s = sheet()) => (s.cells[addr(r, c)] || {}).v ?? '';

  const engine = new F.Engine({
    getRaw: (name, r, c) => {
      const s = book.sheets.find(x => x.name === name);
      return s ? rawAt(r, c, s) : '';
    },
    sheetNames: () => book.sheets.map(s => s.name)
  });
  const valueAt = (r, c) => engine.value(sheet().name, r, c);

  function setCell(r, c, props, s = sheet()) {
    const key = addr(r, c);
    const cur = s.cells[key] || {};
    const next = { ...cur, ...props };
    if (props.s) next.s = { ...(cur.s || {}), ...props.s };
    if ((next.v === undefined || next.v === '') && (!next.s || !Object.keys(next.s).length)) delete s.cells[key];
    else s.cells[key] = next;
    if (r >= s.rows - 1) s.rows = r + 20;
    if (c >= s.cols - 1) s.cols = c + 5;
  }

  // ---------- Undo ----------
  const undoStack = [], redoStack = [];
  function snapshot() {
    undoStack.push(JSON.stringify(book));
    if (undoStack.length > 100) undoStack.shift();
    redoStack.length = 0;
  }
  function restore(json) {
    book = JSON.parse(json);
    engine.astCache.clear();
    buildGrid();
    changed(false);
  }
  function undo() { if (!undoStack.length) return; redoStack.push(JSON.stringify(book)); restore(undoStack.pop()); }
  function redo() { if (!redoStack.length) return; undoStack.push(JSON.stringify(book)); restore(redoStack.pop()); }

  // ---------- Selection ----------
  let active = { r: 0, c: 0 };
  let anchor = { r: 0, c: 0 };
  const selRange = () => ({ r1: Math.min(active.r, anchor.r), c1: Math.min(active.c, anchor.c), r2: Math.max(active.r, anchor.r), c2: Math.max(active.c, anchor.c) });
  const rangeText = g => g.r1 === g.r2 && g.c1 === g.c2 ? addr(g.r1, g.c1) : `${addr(g.r1, g.c1)}:${addr(g.r2, g.c2)}`;
  function forRange(g, fn) { for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) fn(r, c); }
  const td = (r, c) => tbody.rows[r] && tbody.rows[r].cells[c + 1];
  function mergeAt(r, c) { return sheet().merges.find(m => r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2); }

  function select(r, c, extend = false) {
    r = Math.max(0, r); c = Math.max(0, c);
    ensureSize(r, c);
    const m = mergeAt(r, c);
    if (m && !extend) { r = m.r1; c = m.c1; }
    active = { r, c };
    if (!extend) anchor = { r, c };
    drawSelection();
    scrollIntoView(r, c);
  }
  function drawSelection() {
    tbody.querySelectorAll('td.sel').forEach(t => t.classList.remove('sel'));
    thead.querySelectorAll('th.hl').forEach(t => t.classList.remove('hl'));
    tbody.querySelectorAll('th.hl').forEach(t => t.classList.remove('hl'));
    let g = selRange();
    // Grow the box to cover merged cells.
    sheet().merges.forEach(m => { if (m.r1 <= g.r2 && m.r2 >= g.r1 && m.c1 <= g.c2 && m.c2 >= g.c1) g = { r1: Math.min(g.r1, m.r1), c1: Math.min(g.c1, m.c1), r2: Math.max(g.r2, m.r2), c2: Math.max(g.c2, m.c2) }; });
    if (g.r1 !== g.r2 || g.c1 !== g.c2) forRange(g, (r, c) => { const t = td(r, c); if (t && !(r === active.r && c === active.c)) t.classList.add('sel'); });
    for (let c = g.c1; c <= g.c2; c++) thead.rows[0].cells[c + 1]?.classList.add('hl');
    for (let r = g.r1; r <= g.r2; r++) tbody.rows[r]?.cells[0].classList.add('hl');
    const a = td(g.r1, g.c1), b = td(g.r2, g.c2);
    if (a && b) {
      const wr = wrap.getBoundingClientRect();
      const ra = a.getBoundingClientRect(), rb = b.getBoundingClientRect();
      Object.assign(selBox.style, {
        display: 'block',
        left: (ra.left - wr.left + wrap.scrollLeft - 1) + 'px',
        top: (ra.top - wr.top + wrap.scrollTop - 1) + 'px',
        width: (rb.right - ra.left + 1) + 'px',
        height: (rb.bottom - ra.top + 1) + 'px'
      });
    }
    nameBox.value = rangeText(selRange());
    if (document.activeElement !== fx) fx.value = rawAt(active.r, active.c);
    syncToolbar();
    updateAggregate();
  }
  function scrollIntoView(r, c) {
    const t = td(r, c);
    if (!t) return;
    const wr = wrap.getBoundingClientRect();
    const tr = t.getBoundingClientRect();
    const headH = thead.getBoundingClientRect().height;
    const headW = tbody.rows[0].cells[0].getBoundingClientRect().width;
    if (tr.bottom > wr.bottom - 14) wrap.scrollTop += tr.bottom - wr.bottom + 16;
    if (tr.top < wr.top + headH) wrap.scrollTop -= wr.top + headH - tr.top;
    if (tr.right > wr.right - 14) wrap.scrollLeft += tr.right - wr.right + 16;
    if (tr.left < wr.left + headW) wrap.scrollLeft -= wr.left + headW - tr.left;
  }

  // ---------- Rendering ----------
  function colWidth(c) { return sheet().colWidths[c] || DEFAULT_COL_W; }
  function buildGrid() {
    const s = sheet();
    let colgroup = grid.querySelector('colgroup');
    if (colgroup) colgroup.remove();
    colgroup = document.createElement('colgroup');
    colgroup.innerHTML = '<col style="width:46px">' + Array.from({ length: s.cols }, (_, c) => `<col style="width:${colWidth(c)}px">`).join('');
    grid.prepend(colgroup);
    thead.innerHTML = '<tr><th class="corner" title="Select all"></th>' + Array.from({ length: s.cols }, (_, c) => `<th data-c="${c}">${indexToCol(c)}<span class="resizer" data-c="${c}"></span></th>`).join('') + '</tr>';
    const rows = [];
    for (let r = 0; r < s.rows; r++) {
      rows.push(`<tr${s.hiddenRows.includes(r) ? ' class="filtered"' : ''}><th data-r="${r}">${r + 1}</th>` + '<td></td>'.repeat(s.cols) + '</tr>');
    }
    tbody.innerHTML = rows.join('');
    s.merges.forEach(m => {
      const t = td(m.r1, m.c1);
      if (!t) return;
      t.colSpan = m.c2 - m.c1 + 1;
      t.rowSpan = m.r2 - m.r1 + 1;
      forRange(m, (r, c) => { if (r !== m.r1 || c !== m.c1) td(r, c)?.classList.add('hidden-merged'); });
    });
    grid.style.width = (46 + Array.from({ length: s.cols }, (_, c) => colWidth(c)).reduce((a, b) => a + b, 0)) + 'px';
    renderAll();
    renderTabs();
    renderCharts();
    drawSelection();
  }
  function ensureSize(r, c) {
    const s = sheet();
    if (r < s.rows - 1 && c < s.cols - 1) return;
    s.rows = Math.max(s.rows, r + 50);
    s.cols = Math.max(s.cols, c + 5);
    buildGrid();
  }

  let showFormulas = false;
  function renderCell(r, c) {
    const t = td(r, c);
    if (!t) return;
    const cell = cellAt(r, c);
    const st = (cell && cell.s) || {};
    let v = cell ? valueAt(r, c) : '';
    let cls = '';
    let text;
    if (showFormulas && cell && typeof cell.v === 'string' && cell.v[0] === '=') { text = cell.v; }
    else if (v instanceof F.FormulaError) { text = v.error; cls = 'err'; }
    else if (typeof v === 'number') { text = st.fmt === '@' ? String(cell.v) : F.formatValue(v, st.fmt || 'General'); cls = st.fmt === '@' ? '' : 'num'; }
    else if (typeof v === 'boolean') { text = v ? 'TRUE' : 'FALSE'; cls = 'bool'; }
    else text = v;
    t.textContent = text;
    t.className = cls + (st.wrap ? ' wrap' : '') + (t.classList.contains('hidden-merged') ? ' hidden-merged' : '');
    const style = t.style;
    style.fontWeight = st.b ? '700' : '';
    style.fontStyle = st.i ? 'italic' : '';
    style.textDecoration = st.u ? 'underline' : '';
    style.color = st.color || '';
    style.backgroundColor = st.bg || '';
    style.textAlign = st.align || '';
    style.fontFamily = st.font || '';
    style.fontSize = st.size ? st.size + 'pt' : '';
    const bd = st.border || {};
    style.borderTop = bd.t ? '1px solid #000' : '';
    style.borderLeft = bd.l ? '1px solid #000' : '';
    style.borderRight = bd.r ? '1px solid #000' : '';
    style.borderBottom = bd.b ? '1px solid #000' : '';
    if (sheet().hiddenRows.length && r === 0 && sheet().filterCol !== undefined && c === sheet().filterCol) t.textContent += ' ⏷';
  }
  function renderAll() {
    engine.invalidate();
    const s = sheet();
    // Only cells that have content or style need work; clear the rest cheaply.
    tbody.querySelectorAll('td').forEach(t => { if (t.textContent || t.getAttribute('style')) { t.textContent = ''; t.removeAttribute('style'); t.className = t.classList.contains('hidden-merged') ? 'hidden-merged' : ''; } });
    Object.keys(s.cells).forEach(k => { const p = parseRef(k); renderCell(p.row, p.col); });
  }
  function changed(markDirty = true) {
    renderAll();
    renderCharts();
    drawSelection();
    if (markDirty) state.markDirty();
  }

  function renderTabs() {
    const box = document.getElementById('sheetTabs');
    box.innerHTML = '';
    book.sheets.forEach((s, i) => {
      const t = document.createElement('div');
      t.className = 'tab' + (i === book.active ? ' active' : '');
      t.textContent = s.name;
      t.onclick = () => { if (i !== book.active) { commitEdit(); book.active = i; active = anchor = { r: 0, c: 0 }; buildGrid(); } };
      t.ondblclick = () => renameSheet(i);
      t.oncontextmenu = async e => {
        e.preventDefault();
        const r = await UI.form(`Sheet "${s.name}"`, [{ name: 'a', label: 'Action', type: 'select', options: ['Rename', 'Duplicate', 'Delete', 'Move left', 'Move right'] }]);
        if (!r) return;
        if (r.a === 'Rename') renameSheet(i);
        if (r.a === 'Duplicate') { snapshot(); const copy = JSON.parse(JSON.stringify(s)); copy.name = uniqueName(s.name + ' (2)'); book.sheets.splice(i + 1, 0, copy); book.active = i + 1; buildGrid(); state.markDirty(); }
        if (r.a === 'Delete') deleteSheet(i);
        if (r.a === 'Move left' && i > 0) { snapshot(); book.sheets.splice(i - 1, 0, book.sheets.splice(i, 1)[0]); book.active = i - 1; renderTabs(); state.markDirty(); }
        if (r.a === 'Move right' && i < book.sheets.length - 1) { snapshot(); book.sheets.splice(i + 1, 0, book.sheets.splice(i, 1)[0]); book.active = i + 1; renderTabs(); state.markDirty(); }
      };
      box.appendChild(t);
    });
  }
  function uniqueName(base) {
    let n = base, i = 2;
    while (book.sheets.some(s => s.name.toLowerCase() === n.toLowerCase())) n = base.replace(/ \(\d+\)$/, '') + ` (${i++})`;
    return n;
  }
  async function renameSheet(i) {
    const name = await UI.prompt('Rename Sheet', 'Sheet name', book.sheets[i].name);
    if (!name || name === book.sheets[i].name) return;
    if (/[\\/?*[\]:]/.test(name) || name.length > 31) return UI.toast('Sheet names can\'t contain \\ / ? * [ ] : and must be 31 characters or fewer', 4000);
    if (book.sheets.some((s, j) => j !== i && s.name.toLowerCase() === name.toLowerCase())) return UI.toast('That name is already taken');
    snapshot();
    const old = book.sheets[i].name;
    book.sheets[i].name = name;
    // Update references to the renamed sheet in every formula.
    const quote = n => /^[A-Za-z_][\w.]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`;
    const re = new RegExp(`(^|[^\\w.'])(${old.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}|'${old.replace(/'/g, "''").replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}')!`, 'gi');
    book.sheets.forEach(s => Object.values(s.cells).forEach(cell => {
      if (typeof cell.v === 'string' && cell.v[0] === '=') cell.v = cell.v.replace(re, (m, p) => p + quote(name) + '!');
    }));
    renderTabs();
    changed();
  }
  async function deleteSheet(i) {
    if (book.sheets.length === 1) return UI.toast('A workbook must contain at least one sheet');
    if (!(await UI.confirm('Delete Sheet', `Delete "${book.sheets[i].name}"? This can be undone with Ctrl+Z.`))) return;
    snapshot();
    book.sheets.splice(i, 1);
    book.active = Math.min(book.active, book.sheets.length - 1);
    buildGrid();
    state.markDirty();
  }
  function addSheet() {
    snapshot();
    let n = book.sheets.length + 1;
    while (book.sheets.some(s => s.name === 'Sheet' + n)) n++;
    book.sheets.push(newSheet('Sheet' + n));
    book.active = book.sheets.length - 1;
    active = anchor = { r: 0, c: 0 };
    buildGrid();
    state.markDirty();
  }
  document.getElementById('addSheet').addEventListener('click', addSheet);
  document.getElementById('insSheet').addEventListener('click', addSheet);

  // ---------- Editing ----------
  let editing = false;
  let editFromFx = false;
  function startEdit(initial) {
    const t = td(active.r, active.c);
    if (!t) return;
    const wr = wrap.getBoundingClientRect();
    const r = t.getBoundingClientRect();
    editing = true;
    editor.hidden = false;
    Object.assign(editor.style, {
      left: (r.left - wr.left + wrap.scrollLeft - 1) + 'px',
      top: (r.top - wr.top + wrap.scrollTop - 1) + 'px',
      minWidth: (r.width + 2) + 'px', width: (r.width + 2) + 'px', height: (r.height + 2) + 'px'
    });
    const st = (cellAt(active.r, active.c) || {}).s || {};
    editor.style.fontWeight = st.b ? '700' : '';
    editor.style.fontStyle = st.i ? 'italic' : '';
    editor.value = initial !== undefined ? initial : rawAt(active.r, active.c);
    fx.value = editor.value;
    editor.focus();
    editor.setSelectionRange(editor.value.length, editor.value.length);
    autoSizeEditor();
    document.getElementById('stMode').textContent = initial !== undefined ? 'Enter' : 'Edit';
  }
  function autoSizeEditor() {
    editor.style.width = 'auto';
    editor.style.width = Math.max(parseFloat(editor.style.minWidth), Math.min(600, editor.scrollWidth + 8)) + 'px';
  }
  function commitEdit(moveR = 0, moveC = 0) {
    if (!editing) return;
    editing = false;
    editor.hidden = true;
    hideHint();
    let v = editor.value;
    // Close unbalanced parentheses like Excel does.
    if (v[0] === '=') { const open = (v.match(/\(/g) || []).length - (v.match(/\)/g) || []).length; if (open > 0) v += ')'.repeat(open); v = v.replace(/^=\s*/, '='); }
    if (v !== rawAt(active.r, active.c)) {
      snapshot();
      setCell(active.r, active.c, { v });
      autoFormat(active.r, active.c, v);
      ensureSize(active.r, active.c);
      changed();
    }
    document.getElementById('stMode').textContent = 'Ready';
    wrap.focus();
    if (moveR || moveC) select(active.r + moveR, active.c + moveC);
    else drawSelection();
  }
  /** Gives typed dates, percentages and currency a matching number format. */
  function autoFormat(r, c, v) {
    const st = (cellAt(r, c) || {}).s || {};
    if (st.fmt || typeof v !== 'string') return;
    const s = v.trim();
    if (/^\d+(\.\d+)?%$/.test(s)) setCell(r, c, { s: { fmt: s.includes('.') ? '0.00%' : '0%' } });
    else if (/^-?\$[\d,]+(\.\d+)?$/.test(s)) setCell(r, c, { s: { fmt: '$#,##0.00' } });
    else if (/^\d{1,4}[-/]\d{1,2}[-/]\d{1,4}$/.test(s)) {
      const parts = s.split(/[-/]/).map(Number);
      const d = parts[0] > 31 ? new Date(parts[0], parts[1] - 1, parts[2]) : new Date(parts[2] < 100 ? 2000 + parts[2] : parts[2], parts[0] - 1, parts[1]);
      if (!isNaN(d)) setCell(r, c, { v: String(F.dateToSerial(d)), s: { fmt: 'm/d/yyyy' } });
    }
  }
  function cancelEdit() {
    editing = false;
    editor.hidden = true;
    hideHint();
    fx.value = rawAt(active.r, active.c);
    document.getElementById('stMode').textContent = 'Ready';
    wrap.focus();
  }
  editor.addEventListener('input', () => { fx.value = editor.value; autoSizeEditor(); showHint(editor); });
  editor.addEventListener('keydown', e => {
    e.stopPropagation(); // the editor sits inside the grid; don't let the grid handle these keys too
    if (hintKey(e)) return;
    if (e.key === 'Enter' && e.altKey) { e.preventDefault(); const p = editor.selectionStart; editor.setRangeText('\n', p, p, 'end'); return; }
    if (e.key === 'Enter') { e.preventDefault(); commitEdit(e.shiftKey ? -1 : 1, 0); }
    else if (e.key === 'Tab') { e.preventDefault(); commitEdit(0, e.shiftKey ? -1 : 1); }
    else if (e.key === 'Escape') { e.preventDefault(); cancelEdit(); }
    else if (/^Arrow/.test(e.key) && document.getElementById('stMode').textContent === 'Enter' && !isRefInsertPoint()) {
      e.preventDefault();
      const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key];
      commitEdit(...d);
    }
  });
  fx.addEventListener('focus', () => { editFromFx = true; });
  fx.addEventListener('input', () => { if (!editing) startEdit(fx.value); editor.value = fx.value; fx.focus(); showHint(fx); });
  fx.addEventListener('keydown', e => {
    if (hintKey(e)) return;
    if (e.key === 'Enter') { e.preventDefault(); editor.value = fx.value; if (!editing) { editing = true; } commitEdit(1, 0); }
    if (e.key === 'Escape') { e.preventDefault(); cancelEdit(); }
  });
  fx.addEventListener('blur', () => { editFromFx = false; });
  nameBox.addEventListener('keydown', e => {
    if (e.key !== 'Enter') return;
    e.preventDefault();
    const [a, b] = nameBox.value.trim().toUpperCase().split(':');
    const p = parseRef(a), q = b ? parseRef(b) : null;
    if (!p) return UI.toast('Enter a reference like B7 or A1:C5');
    select(p.row, p.col);
    if (q) { anchor = { r: p.row, c: p.col }; active = { r: q.row, c: q.col }; drawSelection(); }
    wrap.focus();
  });
  const isRefInsertPoint = () => {
    const v = editor.value.slice(0, editor.selectionStart);
    return v[0] === '=' && /[=(,+\-*/^&<>:]$/.test(v);
  };

  // Function autocomplete
  const hint = document.getElementById('fnHint');
  let hintItems = [], hintIdx = 0, hintTarget = null;
  function showHint(input) {
    const pos = input.selectionStart;
    const before = input.value.slice(0, pos);
    const m = before[0] === '=' && /([A-Za-z][A-Za-z0-9.]*)$/.exec(before);
    if (!m) return hideHint();
    const word = m[1].toUpperCase();
    hintItems = F.FUNCTIONS.filter(f => f.startsWith(word)).slice(0, 12);
    if (!hintItems.length || (hintItems.length === 1 && hintItems[0] === word)) return hideHint();
    hintIdx = 0;
    hintTarget = { input, start: pos - m[1].length, end: pos };
    hint.innerHTML = hintItems.map((f, i) => `<div class="${i === 0 ? 'active' : ''}" data-f="${f}">${f}</div>`).join('');
    const ir = input.getBoundingClientRect(), ws = document.querySelector('.workspace').getBoundingClientRect();
    hint.style.left = (ir.left - ws.left) + 'px';
    hint.style.top = Math.max(0, ir.bottom - ws.top + 2) + 'px';
    hint.hidden = false;
  }
  function hideHint() { hint.hidden = true; hintTarget = null; }
  function applyHint(f) {
    const { input, start, end } = hintTarget;
    input.value = input.value.slice(0, start) + f + '(' + input.value.slice(end);
    input.setSelectionRange(start + f.length + 1, start + f.length + 1);
    if (input === editor) fx.value = editor.value; else editor.value = fx.value;
    hideHint();
  }
  function hintKey(e) {
    if (hint.hidden) return false;
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault();
      hintIdx = (hintIdx + (e.key === 'ArrowDown' ? 1 : -1) + hintItems.length) % hintItems.length;
      [...hint.children].forEach((d, i) => d.classList.toggle('active', i === hintIdx));
      return true;
    }
    if (e.key === 'Tab') { e.preventDefault(); applyHint(hintItems[hintIdx]); return true; }
    if (e.key === 'Escape') { hideHint(); e.preventDefault(); return true; }
    return false;
  }
  hint.addEventListener('mousedown', e => { e.preventDefault(); const d = e.target.closest('[data-f]'); if (d) applyHint(d.dataset.f); });

  // ---------- Mouse ----------
  let dragging = false, refDrag = null;
  function cellFromEvent(e) {
    const t = e.target.closest('td');
    if (!t || !tbody.contains(t)) return null;
    return { r: t.parentElement.rowIndex - 1, c: t.cellIndex - 1 };
  }
  tbody.addEventListener('mousedown', e => {
    if (e.button !== 0) return;
    const rowHead = e.target.closest('tbody th');
    if (rowHead) {
      commitEdit();
      const r = +rowHead.dataset.r;
      if (e.shiftKey) { active = { r, c: sheet().cols - 1 }; anchor = { r: anchor.r, c: 0 }; }
      else { anchor = { r, c: 0 }; active = { r, c: sheet().cols - 1 }; }
      drawSelection();
      return;
    }
    const p = cellFromEvent(e);
    if (!p) return;
    if (editing && editor.value[0] === '=' && isRefInsertPoint()) {
      // Clicking while typing a formula inserts a reference.
      e.preventDefault();
      const pos = editor.selectionStart;
      const ref = addr(p.r, p.c);
      editor.setRangeText(ref, pos, pos, 'end');
      refDrag = { start: pos, from: p };
      fx.value = editor.value;
      return;
    }
    commitEdit();
    if (e.shiftKey) { active = p; drawSelection(); }
    else select(p.r, p.c);
    dragging = true;
    wrap.focus();
    e.preventDefault();
  });
  tbody.addEventListener('mousemove', e => {
    const p = cellFromEvent(e);
    if (!p) return;
    if (refDrag) {
      const a = refDrag.from;
      const text = (a.r === p.r && a.c === p.c) ? addr(a.r, a.c) : `${addr(Math.min(a.r, p.r), Math.min(a.c, p.c))}:${addr(Math.max(a.r, p.r), Math.max(a.c, p.c))}`;
      editor.setRangeText(text, refDrag.start, editor.selectionStart, 'end');
      fx.value = editor.value;
      return;
    }
    if (dragging && (p.r !== active.r || p.c !== active.c)) { active = p; drawSelection(); }
  });
  document.addEventListener('mouseup', () => { dragging = false; if (refDrag) { refDrag = null; editor.focus(); } });
  tbody.addEventListener('dblclick', e => { if (cellFromEvent(e)) startEdit(); });
  thead.addEventListener('mousedown', e => {
    if (e.target.classList.contains('resizer')) return startResize(e);
    const th = e.target.closest('th');
    if (!th) return;
    commitEdit();
    if (th.classList.contains('corner')) { anchor = { r: 0, c: 0 }; active = { r: sheet().rows - 1, c: sheet().cols - 1 }; drawSelection(); return; }
    const c = +th.dataset.c;
    if (e.shiftKey) { active = { r: sheet().rows - 1, c }; anchor = { r: 0, c: anchor.c }; }
    else { anchor = { r: 0, c }; active = { r: sheet().rows - 1, c }; }
    drawSelection();
  });
  thead.addEventListener('dblclick', e => {
    if (!e.target.classList.contains('resizer')) return;
    autoFitColumn(+e.target.dataset.c);
  });
  function startResize(e) {
    e.preventDefault();
    const c = +e.target.dataset.c;
    const startX = e.clientX, startW = colWidth(c);
    const col = grid.querySelectorAll('colgroup col')[c + 1];
    const move = ev => { const w = Math.max(20, startW + ev.clientX - startX); col.style.width = w + 'px'; sheet().colWidths[c] = w; };
    const up = () => { document.removeEventListener('mousemove', move); document.removeEventListener('mouseup', up); buildGrid(); state.markDirty(); };
    document.addEventListener('mousemove', move);
    document.addEventListener('mouseup', up);
  }
  function autoFitColumn(c) {
    const ctx = document.createElement('canvas').getContext('2d');
    ctx.font = '11pt Calibri, Arial';
    let w = 40;
    for (let r = 0; r < sheet().rows; r++) { const t = td(r, c); if (t && t.textContent) w = Math.max(w, ctx.measureText(t.textContent).width + 14); }
    sheet().colWidths[c] = Math.min(500, Math.ceil(w));
    buildGrid();
    state.markDirty();
  }
  wrap.addEventListener('scroll', () => {
    // Grow the sheet as the user scrolls near the end, like an endless grid.
    const s = sheet();
    if (wrap.scrollTop + wrap.clientHeight > wrap.scrollHeight - 200 && s.rows < 100000) { s.rows += 100; buildGrid(); }
    if (wrap.scrollLeft + wrap.clientWidth > wrap.scrollWidth - 200 && s.cols < 16384) { s.cols += 10; buildGrid(); }
  });

  // ---------- Keyboard ----------
  wrap.addEventListener('keydown', e => {
    if (editing) return;
    const mod = e.ctrlKey || e.metaKey;
    const k = e.key;
    const moves = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (moves[k]) {
      e.preventDefault();
      let [dr, dc] = moves[k];
      if (mod) {
        // Jump to the edge of the data region.
        let r = active.r, c = active.c;
        const filled = (r2, c2) => rawAt(r2, c2) !== '';
        const startFilled = filled(r, c) && filled(r + dr, c + dc);
        do { r += dr; c += dc; } while (r >= 0 && c >= 0 && r < sheet().rows && c < sheet().cols && (startFilled ? filled(r + dr, c + dc) : !filled(r, c)));
        r = Math.max(0, Math.min(sheet().rows - 1, r)); c = Math.max(0, Math.min(sheet().cols - 1, c));
        if (e.shiftKey) { active = { r, c }; drawSelection(); scrollIntoView(r, c); } else select(r, c);
        return;
      }
      // Skip over hidden (filtered) rows.
      let nr = active.r + dr;
      while (dr && sheet().hiddenRows.includes(nr)) nr += dr;
      const m = mergeAt(active.r, active.c);
      if (m && !e.shiftKey) { if (dr > 0) nr = m.r2 + 1; if (dc > 0) dc = m.c2 - active.c + 1; }
      if (e.shiftKey) { active = { r: Math.max(0, nr), c: Math.max(0, active.c + dc) }; ensureSize(active.r, active.c); drawSelection(); scrollIntoView(active.r, active.c); }
      else select(nr, active.c + dc);
      return;
    }
    if (k === 'Enter') { e.preventDefault(); moveInSelection(e.shiftKey ? -1 : 1, 0); return; }
    if (k === 'Tab') { e.preventDefault(); moveInSelection(0, e.shiftKey ? -1 : 1); return; }
    if (k === 'F2') { e.preventDefault(); startEdit(); return; }
    if (k === 'Home') { e.preventDefault(); select(mod ? 0 : active.r, 0); return; }
    if (k === 'PageDown' || k === 'PageUp') { e.preventDefault(); select(active.r + (k === 'PageDown' ? 20 : -20), active.c); return; }
    if (k === 'Delete' || k === 'Backspace') { e.preventDefault(); clearRange('contents'); return; }
    if (mod && k.toLowerCase() === 'a') { e.preventDefault(); anchor = { r: 0, c: 0 }; active = { r: sheet().rows - 1, c: sheet().cols - 1 }; drawSelection(); return; }
    if (mod && k.toLowerCase() === 'b') { e.preventDefault(); toggleStyle('b'); return; }
    if (mod && k.toLowerCase() === 'i') { e.preventDefault(); toggleStyle('i'); return; }
    if (mod && k.toLowerCase() === 'u') { e.preventDefault(); toggleStyle('u'); return; }
    if (mod && k.toLowerCase() === 'd') { e.preventDefault(); fill('down'); return; }
    if (mod && k.toLowerCase() === 'r') { e.preventDefault(); fill('right'); return; }
    if (mod && k.toLowerCase() === 'z' && !Platform.desktop) { e.preventDefault(); e.shiftKey ? redo() : undo(); return; }
    if (mod && k.toLowerCase() === 'y' && !Platform.desktop) { e.preventDefault(); redo(); return; }
    if (mod && k === ';') { e.preventDefault(); snapshot(); setCell(active.r, active.c, { v: String(Math.floor(F.dateToSerial(new Date()))), s: { fmt: 'm/d/yyyy' } }); changed(); return; }
    if (e.altKey && k === '=') { e.preventDefault(); autoSum(); return; }
    if (k.length === 1 && !mod && !e.altKey) { e.preventDefault(); startEdit(k); }
  });
  function moveInSelection(dr, dc) {
    const g = selRange();
    if (g.r1 === g.r2 && g.c1 === g.c2) return select(active.r + dr, active.c + dc);
    let r = active.r + dr, c = active.c + dc;
    if (c > g.c2) { c = g.c1; r++; } if (c < g.c1) { c = g.c2; r--; }
    if (r > g.r2) { r = g.r1; c = dc ? c : c + 1; } if (r < g.r1) { r = g.r2; }
    if (c > g.c2) c = g.c1;
    active = { r, c };
    drawSelection();
  }

  // ---------- Clipboard ----------
  let clip = null;
  function rangeToTsv(g) {
    const lines = [];
    for (let r = g.r1; r <= g.r2; r++) {
      const row = [];
      for (let c = g.c1; c <= g.c2; c++) {
        const v = valueAt(r, c);
        const st = (cellAt(r, c) || {}).s || {};
        let t = v instanceof F.FormulaError ? v.error : typeof v === 'number' ? F.formatValue(v, st.fmt || 'General') : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v);
        if (/[\t\n"]/.test(t)) t = '"' + t.replace(/"/g, '""') + '"';
        row.push(t);
      }
      lines.push(row.join('\t'));
    }
    return lines.join('\n');
  }
  function copy(cut) {
    const g = selRange();
    const cells = [];
    for (let r = g.r1; r <= g.r2; r++) {
      const row = [];
      for (let c = g.c1; c <= g.c2; c++) row.push(JSON.parse(JSON.stringify(cellAt(r, c) || {})));
      cells.push(row);
    }
    const text = rangeToTsv(g);
    clip = { g, cells, text, cut, sheet: book.active };
    return text;
  }
  function parseDelimited(text, delim) {
    const rows = [];
    let row = [], cur = '', q = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (q) {
        if (ch === '"' && text[i + 1] === '"') { cur += '"'; i++; }
        else if (ch === '"') q = false;
        else cur += ch;
      } else if (ch === '"' && cur === '') q = true;
      else if (ch === delim) { row.push(cur); cur = ''; }
      else if (ch === '\n' || ch === '\r') { if (ch === '\r' && text[i + 1] === '\n') i++; row.push(cur); rows.push(row); row = []; cur = ''; }
      else cur += ch;
    }
    if (cur !== '' || row.length) { row.push(cur); rows.push(row); }
    return rows;
  }
  function paste(text) {
    snapshot();
    const r0 = active.r, c0 = active.c;
    if (clip && text.replace(/\r/g, '') === clip.text.replace(/\r/g, '')) {
      const h = clip.cells.length, w = clip.cells[0].length;
      clip.cells.forEach((row, i) => row.forEach((cell, j) => {
        let v = cell.v ?? '';
        if (typeof v === 'string' && v[0] === '=' && !clip.cut) v = F.shiftFormula(v, r0 - clip.g.r1, c0 - clip.g.c1);
        setCell(r0 + i, c0 + j, { v, s: cell.s ? { ...cell.s } : undefined });
        if (!cell.s) { const cur = cellAt(r0 + i, c0 + j); if (cur) delete cur.s; }
      }));
      if (clip.cut) {
        const src = book.sheets[clip.sheet];
        forRange(clip.g, (r, c) => {
          const inside = book.active === clip.sheet && r >= r0 && r < r0 + clip.cells.length && c >= c0 && c < c0 + clip.cells[0].length;
          if (!inside) delete src.cells[addr(r, c)];
        });
        clip = null;
      }
      anchor = { r: r0 + h - 1, c: c0 + w - 1 };
    } else {
      const rows = parseDelimited(text.replace(/\n$/, ''), text.includes('\t') ? '\t' : ',');
      rows.forEach((row, i) => row.forEach((v, j) => setCell(r0 + i, c0 + j, { v })));
      if (rows.length) { anchor = { r: r0 + rows.length - 1, c: c0 + Math.max(...rows.map(r => r.length)) - 1 }; }
    }
    ensureSize(anchor.r, anchor.c);
    buildGrid();
    state.markDirty();
  }
  document.addEventListener('copy', e => { if (editing || document.activeElement === fx || document.activeElement === nameBox || isModalOpen()) return; e.preventDefault(); e.clipboardData.setData('text/plain', copy(false)); UI.toast('Copied'); });
  document.addEventListener('cut', e => { if (editing || document.activeElement === fx || document.activeElement === nameBox || isModalOpen()) return; e.preventDefault(); e.clipboardData.setData('text/plain', copy(true)); });
  document.addEventListener('paste', e => { if (editing || document.activeElement === fx || document.activeElement === nameBox || isModalOpen()) return; e.preventDefault(); paste(e.clipboardData.getData('text/plain')); });
  const isModalOpen = () => !!document.querySelector('.modal-backdrop') || document.getElementById('findbar').classList.contains('open') && document.activeElement.closest('.findbar');

  // ---------- Formatting ----------
  function styleRange(fn) {
    snapshot();
    forRange(selRange(), (r, c) => {
      const cur = { ...((cellAt(r, c) || {}).s || {}) };
      const next = fn(cur, r, c) || cur;
      Object.keys(next).forEach(k => { if (next[k] === undefined || next[k] === false || next[k] === '') delete next[k]; });
      const cell = cellAt(r, c);
      if (cell) { cell.s = next; if (!Object.keys(next).length) delete cell.s; if (cell.v === undefined || cell.v === '') { if (!cell.s) delete sheet().cells[addr(r, c)]; } }
      else if (Object.keys(next).length) sheet().cells[addr(r, c)] = { v: '', s: next };
    });
    changed();
  }
  function toggleStyle(k) {
    const on = !((cellAt(active.r, active.c) || {}).s || {})[k];
    styleRange(s => { s[k] = on; });
  }
  document.querySelectorAll('[data-style]').forEach(b => b.addEventListener('click', () => toggleStyle(b.dataset.style)));
  document.querySelectorAll('[data-align]').forEach(b => b.addEventListener('click', () => {
    const a = b.dataset.align;
    const cur = ((cellAt(active.r, active.c) || {}).s || {}).align;
    styleRange(s => { s.align = cur === a ? '' : a; });
  }));
  document.getElementById('fontName').addEventListener('change', e => { styleRange(s => { s.font = e.target.value === 'Calibri' ? '' : e.target.value; }); wrap.focus(); });
  document.getElementById('fontSize').addEventListener('change', e => { styleRange(s => { s.size = e.target.value === '11' ? '' : +e.target.value; }); wrap.focus(); });
  function colorPicker(id, key) {
    const input = document.getElementById(id);
    const bar = input.parentElement.querySelector('.bar');
    input.addEventListener('input', () => { bar.style.background = input.value; });
    input.addEventListener('change', () => { styleRange(s => { s[key] = input.value; }); wrap.focus(); });
  }
  colorPicker('fillColor', 'bg');
  colorPicker('textColor', 'color');
  document.getElementById('borders').addEventListener('change', e => {
    const kind = e.target.value;
    e.target.value = '';
    if (!kind) return;
    const g = selRange();
    styleRange((s, r, c) => {
      const b = { ...(s.border || {}) };
      if (kind === 'none') return { ...s, border: undefined };
      if (kind === 'all') Object.assign(b, { t: 1, l: 1, r: 1, b: 1 });
      if (kind === 'outer') { if (r === g.r1) b.t = 1; if (r === g.r2) b.b = 1; if (c === g.c1) b.l = 1; if (c === g.c2) b.r = 1; }
      if (kind === 'bottom' && r === g.r2) b.b = 1;
      s.border = Object.keys(b).length ? b : undefined;
    });
    wrap.focus();
  });
  document.getElementById('numFmt').addEventListener('change', e => { styleRange(s => { s.fmt = e.target.value === 'General' ? '' : e.target.value; }); wrap.focus(); });
  document.querySelectorAll('[data-fmt]').forEach(b => b.addEventListener('click', () => styleRange(s => { s.fmt = b.dataset.fmt; })));
  function changeDecimals(delta) {
    styleRange((s, r, c) => {
      let f = s.fmt;
      if (!f || f === 'General') {
        const v = valueAt(r, c);
        if (typeof v !== 'number') return;
        const d = (F.formatGeneral(v).split('.')[1] || '').length;
        f = d ? '0.' + '0'.repeat(d) : '0';
      }
      const m = /(0)(\.0+)?(?!.*0)/.exec(f);
      if (!m) return;
      const decimals = Math.max(0, (m[2] ? m[2].length - 1 : 0) + delta);
      s.fmt = f.slice(0, m.index) + '0' + (decimals ? '.' + '0'.repeat(decimals) : '') + f.slice(m.index + m[0].length);
    });
  }
  document.getElementById('incDec').addEventListener('click', () => changeDecimals(1));
  document.getElementById('decDec').addEventListener('click', () => changeDecimals(-1));
  document.getElementById('merge').addEventListener('click', () => {
    const g = selRange();
    snapshot();
    const s = sheet();
    const existing = s.merges.filter(m => m.r1 <= g.r2 && m.r2 >= g.r1 && m.c1 <= g.c2 && m.c2 >= g.c1);
    if (existing.length) s.merges = s.merges.filter(m => !existing.includes(m));
    else if (g.r1 !== g.r2 || g.c1 !== g.c2) {
      // Keep only the top-left value, as Excel does.
      forRange(g, (r, c) => { if ((r !== g.r1 || c !== g.c1) && cellAt(r, c)) delete s.cells[addr(r, c)]; });
      s.merges.push({ ...g });
      setCell(g.r1, g.c1, { s: { align: 'center' } });
      anchor = active = { r: g.r1, c: g.c1 };
    }
    buildGrid();
    state.markDirty();
  });
  function syncToolbar() {
    const st = (cellAt(active.r, active.c) || {}).s || {};
    document.querySelectorAll('[data-style]').forEach(b => b.classList.toggle('on', !!st[b.dataset.style]));
    document.querySelectorAll('[data-align]').forEach(b => b.classList.toggle('on', st.align === b.dataset.align));
    document.getElementById('fontName').value = st.font || 'Calibri';
    document.getElementById('fontSize').value = String(st.size || 11);
    const nf = document.getElementById('numFmt');
    nf.value = [...nf.options].some(o => o.value === st.fmt) ? st.fmt : 'General';
    document.getElementById('merge').classList.toggle('on', !!mergeAt(active.r, active.c));
  }

  // ---------- Cells: insert/delete ----------
  function shiftStructure(axis, index, count) {
    snapshot();
    const s = sheet();
    const moved = {};
    Object.entries(s.cells).forEach(([k, cell]) => {
      const p = parseRef(k);
      const v = axis === 'row' ? p.row : p.col;
      if (count < 0 && v >= index && v < index - count) return;
      const nv = v >= index ? v + count : v;
      moved[axis === 'row' ? addr(nv, p.col) : addr(p.row, nv)] = cell;
    });
    s.cells = moved;
    // Fix formulas on every sheet that point at this one.
    book.sheets.forEach(other => Object.values(other.cells).forEach(cell => {
      if (typeof cell.v === 'string' && cell.v[0] === '=') cell.v = F.adjustForInsert(cell.v, axis, index, count, s.name, other.name);
    }));
    if (axis === 'row') { s.rows = Math.max(1, s.rows + count); s.hiddenRows = []; }
    else {
      s.cols = Math.max(1, s.cols + count);
      const w = {};
      Object.entries(s.colWidths).forEach(([c, width]) => { c = +c; if (count < 0 && c >= index && c < index - count) return; w[c >= index ? c + count : c] = width; });
      s.colWidths = w;
    }
    s.merges = s.merges.map(m => {
      const k1 = axis === 'row' ? 'r1' : 'c1', k2 = axis === 'row' ? 'r2' : 'c2';
      const n = { ...m };
      if (count < 0 && m[k1] >= index && m[k2] < index - count) return null;
      if (m[k1] >= index) n[k1] = Math.max(index, m[k1] + count);
      if (m[k2] >= index) n[k2] = Math.max(n[k1], m[k2] + count);
      return n;
    }).filter(Boolean);
    buildGrid();
    state.markDirty();
  }
  document.getElementById('insRow').addEventListener('click', () => { const g = selRange(); shiftStructure('row', g.r1, g.r2 - g.r1 + 1 > 1000 ? 1 : g.r2 - g.r1 + 1); });
  document.getElementById('insCol').addEventListener('click', () => { const g = selRange(); shiftStructure('col', g.c1, g.c2 - g.c1 + 1 > 100 ? 1 : g.c2 - g.c1 + 1); });
  document.getElementById('delRow').addEventListener('click', () => { const g = selRange(); shiftStructure('row', g.r1, -(g.r2 - g.r1 + 1)); select(g.r1, active.c); });
  document.getElementById('delCol').addEventListener('click', () => { const g = selRange(); shiftStructure('col', g.c1, -(g.c2 - g.c1 + 1)); select(active.r, g.c1); });

  // ---------- Editing group ----------
  function clearRange(what) {
    snapshot();
    forRange(selRange(), (r, c) => {
      const cell = cellAt(r, c);
      if (!cell) return;
      if (what === 'all') delete sheet().cells[addr(r, c)];
      else if (what === 'formats') { delete cell.s; if (!cell.v) delete sheet().cells[addr(r, c)]; }
      else { cell.v = ''; if (!cell.s) delete sheet().cells[addr(r, c)]; }
    });
    changed();
  }
  document.getElementById('clear').addEventListener('change', e => { const v = e.target.value; e.target.value = ''; if (v) clearRange(v); wrap.focus(); });
  function fill(dir) {
    const g = selRange();
    snapshot();
    if (dir === 'down') {
      const r0 = g.r1 === g.r2 ? g.r1 - 1 : g.r1;
      if (r0 < 0) return;
      for (let c = g.c1; c <= g.c2; c++) {
        const src = cellAt(r0, c) || {};
        for (let r = r0 + 1; r <= g.r2; r++) {
          const v = typeof src.v === 'string' && src.v[0] === '=' ? F.shiftFormula(src.v, r - r0, 0) : src.v ?? '';
          setCell(r, c, { v, s: src.s ? { ...src.s } : undefined });
        }
      }
    } else {
      const c0 = g.c1 === g.c2 ? g.c1 - 1 : g.c1;
      if (c0 < 0) return;
      for (let r = g.r1; r <= g.r2; r++) {
        const src = cellAt(r, c0) || {};
        for (let c = c0 + 1; c <= g.c2; c++) {
          const v = typeof src.v === 'string' && src.v[0] === '=' ? F.shiftFormula(src.v, 0, c - c0) : src.v ?? '';
          setCell(r, c, { v, s: src.s ? { ...src.s } : undefined });
        }
      }
    }
    changed();
  }
  document.getElementById('fillDown').addEventListener('click', () => fill('down'));
  document.getElementById('fillRight').addEventListener('click', () => fill('right'));

  function autoSum() {
    const g = selRange();
    snapshot();
    if (g.r1 !== g.r2 || g.c1 !== g.c2) {
      // Sum each selected column into the row below the selection.
      for (let c = g.c1; c <= g.c2; c++) setCell(g.r2 + 1, c, { v: `=SUM(${addr(g.r1, c)}:${addr(g.r2, c)})` });
      changed();
      select(g.r2 + 1, g.c1);
      return;
    }
    // Single cell: look up for a block of numbers, else left.
    let r = active.r - 1;
    while (r >= 0 && typeof valueAt(r, active.c) === 'number') r--;
    if (r < active.r - 1) { setCell(active.r, active.c, { v: `=SUM(${addr(r + 1, active.c)}:${addr(active.r - 1, active.c)})` }); changed(); return; }
    let c = active.c - 1;
    while (c >= 0 && typeof valueAt(active.r, c) === 'number') c--;
    if (c < active.c - 1) { setCell(active.r, active.c, { v: `=SUM(${addr(active.r, c + 1)}:${addr(active.r, active.c - 1)})` }); changed(); return; }
    undoStack.pop();
    startEdit('=SUM(');
  }
  document.getElementById('autoSum').addEventListener('click', autoSum);
  document.getElementById('autoSum2').addEventListener('click', autoSum);

  function dataRegion() {
    // Use the selection if it is a block, else the contiguous data block around the active cell.
    const g = selRange();
    if (g.r1 !== g.r2 && g.r2 - g.r1 < sheet().rows - 1) return g;
    const s = sheet();
    let maxR = 0, maxC = 0;
    Object.keys(s.cells).forEach(k => { const p = parseRef(k); if (s.cells[k].v !== '' && s.cells[k].v !== undefined) { maxR = Math.max(maxR, p.row); maxC = Math.max(maxC, p.col); } });
    return { r1: 0, c1: 0, r2: maxR, c2: maxC };
  }
  async function sortData(desc) {
    const g = dataRegion();
    const col = active.c;
    let hasHeader = false;
    if (g.r1 === 0 && g.r2 > 0) {
      const top = valueAt(0, col), next = valueAt(1, col);
      hasHeader = typeof top === 'string' && typeof next === 'number';
    }
    snapshot();
    const start = hasHeader ? g.r1 + 1 : g.r1;
    const rows = [];
    for (let r = start; r <= g.r2; r++) {
      const row = [];
      for (let c = g.c1; c <= g.c2; c++) row.push(cellAt(r, c) ? JSON.parse(JSON.stringify(cellAt(r, c))) : null);
      rows.push({ key: valueAt(r, col), row, r });
    }
    rows.sort((a, b) => {
      const x = a.key, y = b.key;
      const blank = v => v === '' || v === null || v === undefined;
      if (blank(x) && blank(y)) return 0;
      if (blank(x)) return 1;
      if (blank(y)) return -1;
      let d = (typeof x === 'number' && typeof y === 'number') ? x - y : (typeof x === 'number' ? -1 : typeof y === 'number' ? 1 : String(x).localeCompare(String(y), undefined, { numeric: true, sensitivity: 'base' }));
      return desc ? -d : d;
    });
    rows.forEach((item, i) => item.row.forEach((cell, j) => {
      const r = start + i, c = g.c1 + j;
      if (!cell) { delete sheet().cells[addr(r, c)]; return; }
      if (typeof cell.v === 'string' && cell.v[0] === '=') cell.v = F.shiftFormula(cell.v, r - item.r, 0);
      sheet().cells[addr(r, c)] = cell;
    }));
    changed();
    UI.toast(`Sorted by column ${indexToCol(col)}${hasHeader ? ' (header kept)' : ''}`);
  }
  ['sortAsc', 'sortAsc2'].forEach(id => document.getElementById(id).addEventListener('click', () => sortData(false)));
  ['sortDesc', 'sortDesc2'].forEach(id => document.getElementById(id).addEventListener('click', () => sortData(true)));

  document.getElementById('filter').addEventListener('click', async () => {
    const g = dataRegion();
    const col = active.c;
    const values = new Set();
    for (let r = g.r1 + 1; r <= g.r2; r++) { const v = valueAt(r, col); values.add(v === '' ? '(Blanks)' : String(v)); }
    const r = await UI.form(`Filter column ${indexToCol(col)}`, [
      { name: 'mode', label: 'Show rows where the value', type: 'select', options: ['equals', 'contains', 'is greater than', 'is less than', 'does not equal'] },
      { name: 'v', label: `Value (e.g. ${[...values].slice(0, 3).join(', ')})`, value: '' }
    ], 'Filter');
    if (!r) return;
    const hidden = [];
    for (let row = g.r1 + 1; row <= g.r2; row++) {
      const v = valueAt(row, col);
      const s = v === '' ? '(Blanks)' : String(v);
      const n = parseFloat(r.v);
      const keep = { equals: s.toLowerCase() === r.v.toLowerCase(), contains: s.toLowerCase().includes(r.v.toLowerCase()),
        'is greater than': typeof v === 'number' && v > n, 'is less than': typeof v === 'number' && v < n, 'does not equal': s.toLowerCase() !== r.v.toLowerCase() }[r.mode];
      if (!keep) hidden.push(row);
    }
    sheet().hiddenRows = hidden;
    sheet().filterCol = col;
    buildGrid();
    UI.toast(`${hidden.length} row${hidden.length === 1 ? '' : 's'} hidden`);
  });
  document.getElementById('clearFilter').addEventListener('click', () => { sheet().hiddenRows = []; delete sheet().filterCol; buildGrid(); });
  document.getElementById('dedupe').addEventListener('click', () => {
    const g = dataRegion();
    snapshot();
    const seen = new Set();
    const keep = [];
    for (let r = g.r1; r <= g.r2; r++) {
      const key = JSON.stringify(Array.from({ length: g.c2 - g.c1 + 1 }, (_, j) => String(valueAt(r, g.c1 + j))));
      if (seen.has(key)) continue;
      seen.add(key);
      keep.push(Array.from({ length: g.c2 - g.c1 + 1 }, (_, j) => cellAt(r, g.c1 + j)));
    }
    const removed = g.r2 - g.r1 + 1 - keep.length;
    forRange(g, (r, c) => delete sheet().cells[addr(r, c)]);
    keep.forEach((row, i) => row.forEach((cell, j) => { if (cell) sheet().cells[addr(g.r1 + i, g.c1 + j)] = cell; }));
    changed();
    UI.toast(`${removed} duplicate row${removed === 1 ? '' : 's'} removed`);
  });
  document.getElementById('textToCols').addEventListener('click', async () => {
    const r = await UI.form('Text to Columns', [{ name: 'd', label: 'Delimiter', type: 'select', options: [{ value: ',', label: 'Comma' }, { value: '\t', label: 'Tab' }, { value: ';', label: 'Semicolon' }, { value: ' ', label: 'Space' }, { value: '|', label: 'Pipe' }] }], 'Split');
    if (!r) return;
    const g = selRange();
    snapshot();
    for (let row = g.r1; row <= g.r2; row++) {
      const v = String(rawAt(row, g.c1));
      if (!v || v[0] === '=') continue;
      v.split(r.d).forEach((part, j) => setCell(row, g.c1 + j, { v: part.trim() }));
    }
    buildGrid();
    state.markDirty();
  });

  // ---------- Functions dialog ----------
  async function insertFunction() {
    const r = await UI.form('Insert Function', [{ name: 'f', label: `Function (${F.FUNCTIONS.length} available)`, type: 'select', options: F.FUNCTIONS }], 'Insert');
    if (r) { wrap.focus(); startEdit('=' + r.f + '('); }
  }
  document.getElementById('insFunction').addEventListener('click', insertFunction);
  document.getElementById('insFunction2').addEventListener('click', insertFunction);
  document.getElementById('showFormulas').addEventListener('click', e => { showFormulas = !showFormulas; e.currentTarget.classList.toggle('on', showFormulas); renderAll(); });
  document.getElementById('recalc').addEventListener('click', () => { changed(false); UI.toast('Recalculated'); });

  // ---------- View ----------
  document.getElementById('freezeRow').addEventListener('click', e => { sheet().freezeRow = !sheet().freezeRow; document.body.classList.toggle('freeze-row', sheet().freezeRow); e.currentTarget.classList.toggle('on', sheet().freezeRow); });
  document.getElementById('freezeCol').addEventListener('click', e => { sheet().freezeCol = !sheet().freezeCol; document.body.classList.toggle('freeze-col', sheet().freezeCol); e.currentTarget.classList.toggle('on', sheet().freezeCol); });
  document.getElementById('gridlines').addEventListener('click', e => { document.body.classList.toggle('no-grid'); e.currentTarget.classList.toggle('on', document.body.classList.contains('no-grid')); });
  const zoom = document.getElementById('zoom');
  function setZoom(z) {
    z = Math.max(50, Math.min(200, Math.round(z)));
    zoom.value = z;
    grid.style.zoom = z / 100;
    document.getElementById('zoomLabel').textContent = z + '%';
    drawSelection();
    renderCharts();
  }
  zoom.addEventListener('input', () => setZoom(+zoom.value));
  document.getElementById('zoomIn').addEventListener('click', () => setZoom(+zoom.value + 10));
  document.getElementById('zoomOut').addEventListener('click', () => setZoom(+zoom.value - 10));
  document.getElementById('zoom100').addEventListener('click', () => setZoom(100));

  // ---------- Status bar aggregate ----------
  function updateAggregate() {
    const g = selRange();
    const el = document.getElementById('stAgg');
    if ((g.r1 === g.r2 && g.c1 === g.c2) || (g.r2 - g.r1 + 1) * (g.c2 - g.c1 + 1) > 200000) { el.textContent = ''; return; }
    let sum = 0, count = 0, numCount = 0;
    const s = sheet();
    Object.keys(s.cells).forEach(k => {
      const p = parseRef(k);
      if (p.row < g.r1 || p.row > g.r2 || p.col < g.c1 || p.col > g.c2) return;
      const v = valueAt(p.row, p.col);
      if (v === '') return;
      count++;
      if (typeof v === 'number') { sum += v; numCount++; }
    });
    el.textContent = numCount ? `Average: ${F.formatGeneral(sum / numCount)}    Count: ${count}    Sum: ${F.formatGeneral(sum)}` : (count ? `Count: ${count}` : '');
  }

  // ---------- Find & replace ----------
  const findbar = document.getElementById('findbar');
  const findText = document.getElementById('findText');
  document.getElementById('findClose').addEventListener('click', () => { findbar.classList.remove('open'); wrap.focus(); });
  function findNext() {
    const q = findText.value.toLowerCase();
    if (!q) return;
    const keys = Object.keys(sheet().cells).map(parseRef).sort((a, b) => a.row - b.row || a.col - b.col);
    const after = keys.filter(p => p.row > active.r || (p.row === active.r && p.col > active.c));
    const hit = [...after, ...keys].find(p => { const v = valueAt(p.row, p.col); return String(v instanceof F.FormulaError ? v.error : v).toLowerCase().includes(q) || String(rawAt(p.row, p.col)).toLowerCase().includes(q); });
    if (hit) select(hit.row, hit.col); else UI.toast('No matches');
  }
  document.getElementById('findNext').addEventListener('click', findNext);
  findText.addEventListener('keydown', e => { if (e.key === 'Enter') findNext(); if (e.key === 'Escape') findbar.classList.remove('open'); });
  document.getElementById('replaceAll').addEventListener('click', () => {
    const q = findText.value;
    if (!q) return;
    snapshot();
    const re = new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
    let n = 0;
    Object.values(sheet().cells).forEach(cell => {
      if (typeof cell.v !== 'string' || cell.v[0] === '=') return;
      const nv = cell.v.replace(re, () => { n++; return document.getElementById('replaceText').value; });
      cell.v = nv;
    });
    changed();
    UI.toast(`Replaced ${n} occurrence${n === 1 ? '' : 's'}`);
  });

  // ---------- Charts ----------
  const PALETTE = ['#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47', '#264478', '#9e480e'];
  function insertChart(type) {
    const g = selRange();
    if (g.r1 === g.r2 && g.c1 === g.c2) return UI.toast('Select the data for the chart first (include labels)');
    snapshot();
    const t = td(g.r1, g.c2 + 1) || td(g.r1, g.c2);
    const wr = wrap.getBoundingClientRect(), tr = t.getBoundingClientRect();
    sheet().charts.push({ type, range: rangeText(g), x: tr.left - wr.left + wrap.scrollLeft + 20, y: tr.top - wr.top + wrap.scrollTop, w: 440, h: 280, title: '' });
    renderCharts();
    state.markDirty();
  }
  document.querySelectorAll('[data-chart]').forEach(b => b.addEventListener('click', () => insertChart(b.dataset.chart)));
  function chartData(ch) {
    const [a, b] = ch.range.split(':');
    const p = parseRef(a), q = parseRef(b || a);
    if (!p || !q) return null;
    const rows = [];
    for (let r = p.row; r <= q.row; r++) { const row = []; for (let c = p.col; c <= q.col; c++) row.push(valueAt(r, c)); rows.push(row); }
    const hasRowLabels = rows.slice(1).some(r => typeof r[0] === 'string' && r[0] !== '') || (rows.length > 1 && rows.every(r => typeof r[0] !== 'number'));
    const hasHeader = rows[0].slice(hasRowLabels ? 1 : 0).some(v => typeof v === 'string' && v !== '');
    const body = hasHeader ? rows.slice(1) : rows;
    const labels = body.map((r, i) => hasRowLabels ? String(r[0]) : String(i + 1));
    const seriesCount = rows[0].length - (hasRowLabels ? 1 : 0);
    const series = [];
    for (let s = 0; s < seriesCount; s++) {
      const ci = s + (hasRowLabels ? 1 : 0);
      series.push({ name: hasHeader ? String(rows[0][ci]) : 'Series ' + (s + 1), values: body.map(r => (typeof r[ci] === 'number' ? r[ci] : 0)) });
    }
    return { labels, series };
  }
  function drawChart(canvas, ch) {
    const dpr = window.devicePixelRatio || 1;
    const W = ch.w, H = ch.h - 22;
    canvas.width = W * dpr; canvas.height = H * dpr;
    canvas.style.width = W + 'px'; canvas.style.height = H + 'px';
    const ctx = canvas.getContext('2d');
    ctx.scale(dpr, dpr);
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, W, H);
    const data = chartData(ch);
    if (!data || !data.series.length) return;
    ctx.font = '12px Calibri, Arial, sans-serif';
    const title = ch.title || (data.series.length === 1 ? data.series[0].name : 'Chart');
    ctx.fillStyle = '#404040'; ctx.textAlign = 'center'; ctx.font = '600 14px Calibri, Arial';
    ctx.fillText(title, W / 2, 20);
    ctx.font = '11px Calibri, Arial';
    if (ch.type === 'pie') {
      const vals = data.series[0].values.map(v => Math.max(0, v));
      const total = vals.reduce((a, b) => a + b, 0) || 1;
      const cx = W * 0.4, cy = H / 2 + 10, rad = Math.min(W * 0.35, H / 2 - 30);
      let a0 = -Math.PI / 2;
      vals.forEach((v, i) => {
        const a1 = a0 + v / total * Math.PI * 2;
        ctx.beginPath(); ctx.moveTo(cx, cy); ctx.arc(cx, cy, rad, a0, a1); ctx.closePath();
        ctx.fillStyle = PALETTE[i % PALETTE.length]; ctx.fill();
        ctx.strokeStyle = '#fff'; ctx.lineWidth = 1.5; ctx.stroke();
        a0 = a1;
      });
      ctx.textAlign = 'left';
      data.labels.forEach((l, i) => {
        const y = 44 + i * 18;
        if (y > H - 10) return;
        ctx.fillStyle = PALETTE[i % PALETTE.length]; ctx.fillRect(W * 0.78, y - 9, 10, 10);
        ctx.fillStyle = '#404040'; ctx.fillText(`${l} (${Math.round(vals[i] / total * 100)}%)`, W * 0.78 + 14, y);
      });
      return;
    }
    const all = data.series.flatMap(s => s.values);
    let max = Math.max(0, ...all), min = Math.min(0, ...all);
    if (max === min) max = min + 1;
    const step = niceStep((max - min) / 5);
    max = Math.ceil(max / step) * step; min = Math.floor(min / step) * step;
    const legendH = data.series.length > 1 ? 20 : 0;
    const L = 48, R = W - 14, T = 34, B = H - 26 - legendH;
    const horizontal = ch.type === 'hbar';
    ctx.strokeStyle = '#e1e1e1'; ctx.fillStyle = '#595959'; ctx.lineWidth = 1;
    for (let v = min; v <= max + 1e-9; v += step) {
      if (horizontal) {
        const x = L + (v - min) / (max - min) * (R - L);
        ctx.beginPath(); ctx.moveTo(x, T); ctx.lineTo(x, B); ctx.stroke();
        ctx.textAlign = 'center'; ctx.fillText(F.formatGeneral(+v.toFixed(10)), x, B + 14);
      } else {
        const y = B - (v - min) / (max - min) * (B - T);
        ctx.beginPath(); ctx.moveTo(L, y); ctx.lineTo(R, y); ctx.stroke();
        ctx.textAlign = 'right'; ctx.fillText(F.formatGeneral(+v.toFixed(10)), L - 6, y + 4);
      }
    }
    const n = data.labels.length;
    const band = (horizontal ? (B - T) : (R - L)) / n;
    const zero = horizontal ? L + (0 - min) / (max - min) * (R - L) : B - (0 - min) / (max - min) * (B - T);
    data.labels.forEach((l, i) => {
      ctx.fillStyle = '#595959';
      if (horizontal) { ctx.textAlign = 'right'; ctx.fillText(l.slice(0, 8), L - 4, T + band * i + band / 2 + 4); }
      else { ctx.textAlign = 'center'; ctx.fillText(l.slice(0, Math.max(3, Math.floor(band / 6))), L + band * i + band / 2, B + 14); }
    });
    data.series.forEach((s, si) => {
      ctx.fillStyle = ctx.strokeStyle = PALETTE[si % PALETTE.length];
      if (ch.type === 'line') {
        ctx.lineWidth = 2.25; ctx.beginPath();
        s.values.forEach((v, i) => { const x = L + band * i + band / 2, y = B - (v - min) / (max - min) * (B - T); i ? ctx.lineTo(x, y) : ctx.moveTo(x, y); });
        ctx.stroke();
        s.values.forEach((v, i) => { const x = L + band * i + band / 2, y = B - (v - min) / (max - min) * (B - T); ctx.beginPath(); ctx.arc(x, y, 3, 0, Math.PI * 2); ctx.fill(); });
      } else {
        const bw = band * 0.7 / data.series.length;
        s.values.forEach((v, i) => {
          const off = band * 0.15 + bw * si;
          if (horizontal) {
            const x = L + (v - min) / (max - min) * (R - L);
            ctx.fillRect(Math.min(zero, x), T + band * i + off, Math.abs(x - zero), bw - 1);
          } else {
            const y = B - (v - min) / (max - min) * (B - T);
            ctx.fillRect(L + band * i + off, Math.min(zero, y), bw - 1, Math.abs(zero - y));
          }
        });
      }
    });
    if (legendH) {
      let x = L;
      ctx.textAlign = 'left';
      data.series.forEach((s, si) => {
        ctx.fillStyle = PALETTE[si % PALETTE.length]; ctx.fillRect(x, H - 16, 10, 10);
        ctx.fillStyle = '#404040'; ctx.fillText(s.name, x + 14, H - 7);
        x += ctx.measureText(s.name).width + 32;
      });
    }
  }
  function niceStep(raw) {
    const p = Math.pow(10, Math.floor(Math.log10(raw || 1)));
    const f = raw / p;
    return (f <= 1 ? 1 : f <= 2 ? 2 : f <= 5 ? 5 : 10) * p;
  }
  function renderCharts() {
    const layer = document.getElementById('chartsLayer');
    layer.innerHTML = '';
    (sheet().charts || []).forEach((ch, i) => {
      const box = document.createElement('div');
      box.className = 'chart';
      Object.assign(box.style, { left: ch.x + 'px', top: ch.y + 'px', width: ch.w + 'px', height: ch.h + 'px' });
      box.innerHTML = `<div class="bar"><button data-a="title" title="Edit title and data">✎</button><button data-a="type" title="Change chart type">⇄</button><button data-a="png" title="Save as image">⤓</button><button data-a="del" title="Delete chart">✕</button></div><canvas></canvas><div class="grip"></div>`;
      layer.appendChild(box);
      drawChart(box.querySelector('canvas'), ch);
      const barEl = box.querySelector('.bar');
      barEl.addEventListener('mousedown', e => {
        if (e.target.tagName === 'BUTTON') return;
        e.preventDefault();
        const sx = e.clientX, sy = e.clientY, ox = ch.x, oy = ch.y;
        const mv = ev => { ch.x = Math.max(0, ox + ev.clientX - sx); ch.y = Math.max(0, oy + ev.clientY - sy); box.style.left = ch.x + 'px'; box.style.top = ch.y + 'px'; };
        const up = () => { document.removeEventListener('mousemove', mv); document.removeEventListener('mouseup', up); state.markDirty(); };
        document.addEventListener('mousemove', mv); document.addEventListener('mouseup', up);
      });
      box.querySelector('.grip').addEventListener('mousedown', e => {
        e.preventDefault();
        const sx = e.clientX, sy = e.clientY, ow = ch.w, oh = ch.h;
        const mv = ev => { ch.w = Math.max(200, ow + ev.clientX - sx); ch.h = Math.max(150, oh + ev.clientY - sy); box.style.width = ch.w + 'px'; box.style.height = ch.h + 'px'; drawChart(box.querySelector('canvas'), ch); };
        const up = () => { document.removeEventListener('mousemove', mv); document.removeEventListener('mouseup', up); state.markDirty(); };
        document.addEventListener('mousemove', mv); document.addEventListener('mouseup', up);
      });
      barEl.addEventListener('click', async e => {
        const a = e.target.dataset.a;
        if (a === 'del') { snapshot(); sheet().charts.splice(i, 1); renderCharts(); state.markDirty(); }
        if (a === 'type') { const order = ['bar', 'line', 'pie', 'hbar']; ch.type = order[(order.indexOf(ch.type) + 1) % order.length]; drawChart(box.querySelector('canvas'), ch); state.markDirty(); }
        if (a === 'title') {
          const r = await UI.form('Chart', [{ name: 'title', label: 'Title', value: ch.title }, { name: 'range', label: 'Data range', value: ch.range }]);
          if (r) { ch.title = r.title; if (/^[A-Z]+\d+(:[A-Z]+\d+)?$/i.test(r.range.trim())) ch.range = r.range.trim().toUpperCase(); drawChart(box.querySelector('canvas'), ch); state.markDirty(); }
        }
        if (a === 'png') {
          const blob = await new Promise(res => box.querySelector('canvas').toBlob(res));
          await Platform.save({ data: new Uint8Array(await blob.arrayBuffer()), name: (ch.title || 'Chart') + '.png', filters: [{ name: 'PNG image', extensions: ['png'] }], saveAs: true });
        }
      });
    });
  }

  // ---------- File I/O ----------
  function workbookFromXlsx(bytes, name) {
    const ext = Platform.extOf(name);
    let wb;
    if (ext === 'csv' || ext === 'tsv') {
      const text = Platform.bytesToText(bytes).replace(/^\ufeff/, '');
      const s = newSheet(Platform.baseName(name).slice(0, 31) || 'Sheet1');
      parseDelimited(text, ext === 'tsv' ? '\t' : (text.split('\n')[0].includes(';') && !text.split('\n')[0].includes(',') ? ';' : ',')).forEach((row, r) => row.forEach((v, c) => { if (v !== '') setCell(r, c, { v }, s); }));
      return { sheets: [s], active: 0 };
    }
    wb = XLSX.read(bytes, { type: 'array', cellFormula: true, cellNF: true, cellStyles: true });
    let meta = null;
    try { meta = wb.Custprops && wb.Custprops.AtlasMeta ? JSON.parse(wb.Custprops.AtlasMeta) : null; } catch { meta = null; }
    const sheets = wb.SheetNames.map((sn, si) => {
      const ws = wb.Sheets[sn];
      const s = newSheet(sn);
      const range = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
      for (let r = range.s.r; r <= range.e.r; r++) {
        for (let c = range.s.c; c <= range.e.c; c++) {
          const cell = ws[XLSX.utils.encode_cell({ r, c })];
          if (!cell) continue;
          let v;
          if (cell.f) v = '=' + cell.f;
          else if (cell.t === 'n') v = String(cell.v);
          else if (cell.t === 'b') v = cell.v ? 'TRUE' : 'FALSE';
          else if (cell.t === 'e') v = cell.w || '#N/A';
          else if (cell.t === 'd') v = String(F.dateToSerial(cell.v));
          else {
            v = cell.v === undefined ? '' : String(cell.v);
            if (v && (v[0] === '=' || typeof F.parseLiteral(v) !== 'string')) v = "'" + v;
          }
          const st = {};
          if (cell.z && cell.z !== 'General') st.fmt = String(cell.z);
          if (v !== '' || Object.keys(st).length) s.cells[addr(r, c)] = Object.keys(st).length ? { v, s: st } : { v };
        }
      }
      s.rows = Math.max(100, range.e.r + 20);
      s.cols = Math.max(26, range.e.c + 5);
      (ws['!cols'] || []).forEach((col, c) => { if (col && (col.wpx || col.wch)) s.colWidths[c] = Math.round(col.wpx || col.wch * 7 + 5); });
      (ws['!merges'] || []).forEach(m => s.merges.push({ r1: m.s.r, c1: m.s.c, r2: m.e.r, c2: m.e.c }));
      const ms = meta && meta.sheets && meta.sheets[si];
      if (ms && ms.name === sn) {
        Object.entries(ms.styles || {}).forEach(([k, st]) => { s.cells[k] = { v: s.cells[k]?.v ?? '', s: { ...(s.cells[k]?.s || {}), ...st } }; });
        s.charts = ms.charts || [];
        if (ms.colWidths) s.colWidths = ms.colWidths;
      }
      return s;
    });
    return { sheets, active: Math.min(meta?.active || 0, sheets.length - 1) };
  }

  function toWorkbook() {
    const wb = XLSX.utils.book_new();
    const meta = { sheets: [], active: book.active };
    book.sheets.forEach(s => {
      const ws = {};
      let maxR = 0, maxC = 0;
      const styles = {};
      Object.entries(s.cells).forEach(([k, cell]) => {
        const p = parseRef(k);
        const st = cell.s || {};
        const visual = { ...st };
        delete visual.fmt;
        if (Object.keys(visual).length) styles[k] = visual;
        if (cell.v === '' || cell.v === undefined) { if (st.fmt) { ws[k] = { t: 'z', z: st.fmt }; } return; }
        maxR = Math.max(maxR, p.row); maxC = Math.max(maxC, p.col);
        const v = engine.value(s.name, p.row, p.col);
        const out = {};
        if (typeof cell.v === 'string' && cell.v[0] === '=') out.f = cell.v.slice(1);
        if (v instanceof F.FormulaError) { out.t = 'e'; out.v = { '#NULL!': 0, '#DIV/0!': 7, '#VALUE!': 15, '#REF!': 23, '#NAME?': 29, '#NUM!': 36, '#N/A': 42 }[v.error] ?? 15; out.w = v.error; }
        else if (typeof v === 'number') { out.t = 'n'; out.v = v; }
        else if (typeof v === 'boolean') { out.t = 'b'; out.v = v; }
        else { out.t = 's'; out.v = String(v); }
        if (st.fmt) out.z = st.fmt;
        ws[k] = out;
      });
      ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxR, c: maxC } });
      const cols = [];
      Object.entries(s.colWidths).forEach(([c, w]) => { cols[+c] = { wpx: w }; });
      if (cols.length) { for (let i = 0; i < cols.length; i++) if (!cols[i]) cols[i] = { wpx: DEFAULT_COL_W }; ws['!cols'] = cols; }
      if (s.merges.length) ws['!merges'] = s.merges.map(m => ({ s: { r: m.r1, c: m.c1 }, e: { r: m.r2, c: m.c2 } }));
      XLSX.utils.book_append_sheet(wb, ws, s.name);
      meta.sheets.push({ name: s.name, styles, charts: s.charts, colWidths: s.colWidths });
    });
    // Cell styling and charts live in a custom document property so Atlas can restore them;
    // other apps ignore it and still get values, formulas, number formats, widths and merges.
    wb.Custprops = { AtlasMeta: JSON.stringify(meta) };
    return wb;
  }
  function serialize(ext) {
    if (ext === 'csv') {
      const s = sheet();
      let maxR = 0, maxC = 0;
      Object.keys(s.cells).forEach(k => { const p = parseRef(k); if (s.cells[k].v !== '') { maxR = Math.max(maxR, p.row); maxC = Math.max(maxC, p.col); } });
      const lines = [];
      for (let r = 0; r <= maxR; r++) {
        const row = [];
        for (let c = 0; c <= maxC; c++) {
          const v = valueAt(r, c);
          const st = (cellAt(r, c) || {}).s || {};
          let t = v instanceof F.FormulaError ? v.error : typeof v === 'number' ? (st.fmt && /[ymd]/i.test(st.fmt) ? F.formatValue(v, st.fmt) : String(v)) : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v);
          if (/[",\n]/.test(t)) t = '"' + t.replace(/"/g, '""') + '"';
          row.push(t);
        }
        lines.push(row.join(','));
      }
      return lines.join('\r\n') + '\r\n';
    }
    const bookType = { xlsx: 'xlsx', ods: 'ods', xls: 'biff8' }[ext] || 'xlsx';
    return new Uint8Array(XLSX.write(toWorkbook(), { type: 'array', bookType, compression: true }));
  }

  async function loadFile(f) {
    try {
      book = workbookFromXlsx(f.data, f.name);
    } catch (err) {
      UI.toast('Could not open file: ' + err.message, 4000);
      return;
    }
    undoStack.length = redoStack.length = 0;
    engine.astCache.clear();
    active = anchor = { r: 0, c: 0 };
    state.set({ name: f.name, path: f.path, dirty: false });
    Platform.rememberRecent(f.name, 'sheets');
    buildGrid();
  }
  async function save(saveAs, forcedExt) {
    commitEdit();
    let name = state.name;
    if (!/\.(xlsx|ods|csv|xls)$/i.test(name)) name = Platform.baseName(name) + '.xlsx';
    if (forcedExt) name = Platform.baseName(name) + '.' + forcedExt;
    const ext = Platform.extOf(name);
    const filters = forcedExt ? SAVE_FILTERS.filter(f => f.extensions.includes(forcedExt)) : [...SAVE_FILTERS].sort((a, b) => (b.extensions.includes(ext) ? 1 : 0) - (a.extensions.includes(ext) ? 1 : 0));
    try {
      if (saveAs || !state.path || forcedExt) {
        const target = await Platform.chooseSavePath(name, filters);
        if (!target) return false;
        const text = Platform.extOf(target.name);
        if (text === 'csv' && book.sheets.length > 1 && !forcedExt) UI.toast('CSV keeps only the current sheet', 3500);
        await Platform.save({ data: serialize(text || 'xlsx'), name: target.name, path: target.path });
        if (!forcedExt) state.set({ name: target.name, path: target.path, dirty: false });
      } else {
        await Platform.save({ data: serialize(ext), name, path: state.path });
        state.set({ dirty: false });
      }
      UI.toast('Saved');
      return true;
    } catch (err) {
      UI.toast('Save failed: ' + err.message, 4000);
      return false;
    }
  }
  function printView() {
    const s = sheet();
    let maxR = 0, maxC = 0;
    Object.keys(s.cells).forEach(k => { const p = parseRef(k); maxR = Math.max(maxR, p.row); maxC = Math.max(maxC, p.col); });
    const t = document.createElement('table');
    for (let r = 0; r <= maxR; r++) {
      const tr = t.insertRow();
      for (let c = 0; c <= maxC; c++) {
        const src = td(r, c);
        const cell = tr.insertCell();
        if (src) { cell.textContent = src.textContent; cell.style.cssText = src.style.cssText; cell.className = src.className; if (src.classList.contains('num')) cell.style.textAlign = cell.style.textAlign || 'right'; }
      }
    }
    const area = document.getElementById('printArea');
    area.innerHTML = '';
    area.appendChild(t);
  }
  window.addEventListener('beforeprint', printView);

  Platform.onMenu(async cmd => {
    switch (cmd) {
      case 'new': Platform.newWindow('sheets'); break;
      case 'open': {
        const f = await Platform.open(OPEN_FILTERS);
        if (!f) break;
        const empty = !state.dirty && book.sheets.length === 1 && !Object.keys(sheet().cells).length;
        if (Platform.shouldOpenInPlace(empty)) {
          if (state.dirty && !(await UI.confirm('Open file', 'Discard unsaved changes in this workbook?'))) break;
          loadFile(f);
        } else Platform.openPath(f.path);
        break;
      }
      case 'save': save(false); break;
      case 'save-as': save(true); break;
      case 'save-and-close': if (await save(false)) Platform.closeNow(); break;
      case 'export-csv': save(true, 'csv'); break;
      case 'export-pdf': printView(); Platform.exportPdf(Platform.baseName(state.name) + '.pdf', true); break;
      case 'print': printView(); Platform.print(); break;
      case 'undo': if (!editing) undo(); else document.execCommand('undo'); break;
      case 'redo': if (!editing) redo(); else document.execCommand('redo'); break;
      case 'find': case 'replace': findbar.classList.add('open'); findText.focus(); findText.select(); break;
      case 'copy-btn': document.execCommand('copy') || UI.toast('Use Ctrl+C'); break;
      case 'cut-btn': document.execCommand('cut') || UI.toast('Use Ctrl+X'); break;
      case 'paste-btn': try { paste(await navigator.clipboard.readText()); } catch { UI.toast('Use Ctrl+V to paste'); } break;
      case 'home': Platform.desktop ? Platform.newWindow('home') : (location.href = 'index.html'); break;
    }
  });

  window.addEventListener('resize', () => { drawSelection(); });
  (async () => {
    buildGrid();
    wrap.focus();
    const f = await Platform.pending();
    if (f && f.error) UI.toast('Could not open file: ' + f.error, 4000);
    else if (f) await loadFile(f);
  })();

  window.AtlasSheets = { loadFile, serialize, get book() { return book; }, select, setCell: (ref, v) => { const p = parseRef(ref); setCell(p.row, p.col, { v }); changed(); }, valueOf: ref => { const p = parseRef(ref); return valueAt(p.row, p.col); } };
})();
