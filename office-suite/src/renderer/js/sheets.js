'use strict';
// Atlas Sheets — spreadsheet.
// The grid is virtual: only the rows and columns on screen exist in the DOM, and every
// position (selection box, editor, scrolling) is computed from row heights and column widths.
(function () {
  const F = window.Formula;
  const { addr, indexToCol, parseRef } = F;
  const wrap = document.getElementById('gridWrap');
  const inner = document.getElementById('gridInner');
  const grid = document.getElementById('grid');
  const thead = grid.querySelector('thead');
  const tbody = grid.querySelector('tbody');
  const selBox = document.getElementById('selBox');
  const editor = document.getElementById('cellEditor');
  const fx = document.getElementById('fx');
  const nameBox = document.getElementById('nameBox');
  const state = UI.docState('Atlas Sheets', 'Book1');
  const DEFAULT_COL_W = 80;
  const ROW_H = 22;          // default row height (px)
  const HEAD_H = 22;         // column header height
  const RH_W = 46;           // row header width
  const MAX_ROWS = 1048576;  // Excel's limits
  const MAX_COLS = 16384;
  const ROW_BUFFER = 20, COL_BUFFER = 4;
  const SAVE_FILTERS = [
    { name: 'Excel Workbook', extensions: ['xlsx'] },
    { name: 'OpenDocument Spreadsheet', extensions: ['ods'] },
    { name: 'CSV (current sheet)', extensions: ['csv'] },
    { name: 'Excel 97-2003 Workbook', extensions: ['xls'] }
  ];
  const OPEN_FILTERS = [{ name: 'Spreadsheets', extensions: ['xlsx', 'xls', 'xlsm', 'ods', 'csv', 'tsv'] }];
  const CELL_KEY = /^[A-Z]{1,3}[1-9]\d{0,6}$/;

  UI.initRibbon();

  // ---------- Model ----------
  const newSheet = name => ({ name, cells: {}, rows: 100, cols: 26, colWidths: {}, rowHeights: {}, merges: [], charts: [], hiddenRows: [], cf: [], pivots: [], dv: [] });
  let book = { sheets: [newSheet('Sheet1')], active: 0 };
  const sheet = () => book.sheets[book.active];
  const cellAt = (r, c, s = sheet()) => s.cells[addr(r, c)];
  const rawAt = (r, c, s = sheet()) => (s.cells[addr(r, c)] || {}).v ?? '';

  const engine = new F.Engine({
    getRaw: (name, r, c) => {
      const s = book.sheets.find(x => x.name === name);
      return s ? rawAt(r, c, s) : '';
    },
    sheetNames: () => book.sheets.map(s => s.name),
    formulasInRange: (name, r1, c1, r2, c2) => {
      const s = book.sheets.find(x => x.name === name);
      if (!s) return [];
      const out = [];
      const area = (r2 - r1 + 1) * (c2 - c1 + 1);
      const keys = Object.keys(s.cells);
      if (area < keys.length) {
        for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) { const v = rawAt(r, c, s); if (typeof v === 'string' && v[0] === '=') out.push([r, c]); }
      } else {
        for (const k of keys) {
          const v = s.cells[k].v;
          if (typeof v !== 'string' || v[0] !== '=') continue;
          const p = parseRef(k);
          if (p.row >= r1 && p.row <= r2 && p.col >= c1 && p.col <= c2) out.push([p.row, p.col]);
        }
      }
      return out;
    }
  });
  const valueAt = (r, c) => engine.value(sheet().name, r, c);
  const cf = SheetsCF({ engine, sheet: () => sheet() });
  const dv = SheetsDV({ engine, sheet: () => sheet(), sheetByName: n => sheetByName(n) });

  function setCell(r, c, props, s = sheet()) {
    const key = addr(r, c);
    const cur = s.cells[key] || {};
    const next = { ...cur, ...props };
    if (props.s) next.s = { ...(cur.s || {}), ...props.s };
    if ((next.v === undefined || next.v === '') && (!next.s || !Object.keys(next.s).length) && !next.n) delete s.cells[key];
    else s.cells[key] = next;
    if (r >= s.rows - 1) s.rows = Math.min(MAX_ROWS, r + 20);
    if (c >= s.cols - 1) s.cols = Math.min(MAX_COLS, c + 5);
  }
  /** Largest row/column that holds anything. */
  function usedBounds(s = sheet()) {
    let r = 0, c = 0;
    for (const k of Object.keys(s.cells)) { const p = parseRef(k); if (p.row > r) r = p.row; if (p.col > c) c = p.col; }
    return { r, c };
  }
  /** A selection clipped to the used area when it is huge (select all, whole columns). */
  function opRange(g = selRange()) {
    if ((g.r2 - g.r1 + 1) * (g.c2 - g.c1 + 1) <= 100000) return g;
    const u = usedBounds();
    return { r1: g.r1, c1: g.c1, r2: Math.max(g.r1, Math.min(g.r2, u.r)), c2: Math.max(g.c1, Math.min(g.c2, u.c)) };
  }

  // ---------- Undo ----------
  const undoStack = [], redoStack = [];
  // With live editing, each step remembers how many remote changes had arrived, so undo
  // takes back only your own changes (see Collab.replaySince).
  const undoMarks = [], redoMarks = [];
  let live = null, peers = [], peerLayer = null;
  const liveMark = () => (live ? live.mark() : 0);
  let undoBytes = 0;
  const UNDO_BUDGET = 150e6; // characters of history kept
  function snapshot() {
    const json = JSON.stringify(book);
    undoStack.push(json);
    undoMarks.push(liveMark());
    undoBytes += json.length;
    // Big workbooks keep fewer steps so history can't exhaust memory.
    while (undoStack.length > 100 || (undoBytes > UNDO_BUDGET && undoStack.length > 1)) { undoBytes -= undoStack.shift().length; undoMarks.shift(); }
    redoStack.length = redoMarks.length = 0;
  }
  function restore(json) {
    book = JSON.parse(json);
    engine.astCache.clear();
    buildGrid();
    changed(false);
  }
  function undo() {
    if (!undoStack.length) return;
    redoStack.push(JSON.stringify(book)); redoMarks.push(liveMark());
    const j = undoStack.pop(), m = undoMarks.pop();
    undoBytes -= j.length;
    restore(j);
    if (live) live.replaySince(m);
  }
  function redo() {
    if (!redoStack.length) return;
    const j = JSON.stringify(book);
    undoStack.push(j); undoMarks.push(liveMark());
    undoBytes += j.length;
    const m = redoMarks.pop();
    restore(redoStack.pop());
    if (live) live.replaySince(m);
  }

  // ---------- Geometry ----------
  let colX = [0];                         // colX[c] = left edge of column c (data area)
  let rowSpecials = [], rowPrefix = [0];  // rows whose height isn't ROW_H, with cumulative deltas
  let hiddenSet = new Set();
  let rowHMap = new Map();
  const zoomLevel = () => +(inner.style.zoom || 1);
  function colWidth(c) { return sheet().colWidths[c] || DEFAULT_COL_W; }
  function rebuildCols() {
    const s = sheet();
    colX = new Array(s.cols + 1);
    colX[0] = 0;
    for (let c = 0; c < s.cols; c++) colX[c + 1] = colX[c] + colWidth(c);
  }
  function rebuildRows() {
    const s = sheet();
    hiddenSet = new Set(s.hiddenRows);
    rowHMap = new Map();
    for (const [r, h] of Object.entries(s.rowHeights || {})) if (+h !== ROW_H) rowHMap.set(+r, +h);
    hiddenSet.forEach(r => rowHMap.set(r, 0));
    rowSpecials = [...rowHMap.keys()].sort((a, b) => a - b);
    rowPrefix = [0];
    rowSpecials.forEach((r, i) => { rowPrefix[i + 1] = rowPrefix[i] + rowHMap.get(r) - ROW_H; });
  }
  const rowH = r => (rowHMap.has(r) ? rowHMap.get(r) : ROW_H);
  function rowTop(r) {
    let lo = 0, hi = rowSpecials.length;
    while (lo < hi) { const m = (lo + hi) >> 1; if (rowSpecials[m] < r) lo = m + 1; else hi = m; }
    return r * ROW_H + rowPrefix[lo];
  }
  function rowAtY(y) {
    let lo = 0, hi = sheet().rows - 1;
    while (lo < hi) { const m = (lo + hi + 1) >> 1; if (rowTop(m) <= y) lo = m; else hi = m - 1; }
    return lo;
  }
  function colAtX(x) {
    let lo = 0, hi = sheet().cols - 1;
    while (lo < hi) { const m = (lo + hi + 1) >> 1; if (colX[m] <= x) lo = m; else hi = m - 1; }
    return lo;
  }
  /** Box of a cell range in grid coordinates (unzoomed). */
  function rangeBox(g) {
    return { x: RH_W + colX[g.c1], y: HEAD_H + rowTop(g.r1), w: colX[g.c2 + 1] - colX[g.c1], h: rowTop(g.r2 + 1) - rowTop(g.r1) };
  }

  // ---------- Selection ----------
  let active = { r: 0, c: 0 };
  let anchor = { r: 0, c: 0 };
  const selRange = () => ({ r1: Math.min(active.r, anchor.r), c1: Math.min(active.c, anchor.c), r2: Math.max(active.r, anchor.r), c2: Math.max(active.c, anchor.c) });
  const rangeText = g => g.r1 === g.r2 && g.c1 === g.c2 ? addr(g.r1, g.c1) : `${addr(g.r1, g.c1)}:${addr(g.r2, g.c2)}`;
  function forRange(g, fn) { for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) fn(r, c); }
  function mergeAt(r, c) { return sheet().merges.find(m => r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2); }
  function expandByMerges(g) {
    let out = { ...g };
    sheet().merges.forEach(m => { if (m.r1 <= out.r2 && m.r2 >= out.r1 && m.c1 <= out.c2 && m.c2 >= out.c1) out = { r1: Math.min(out.r1, m.r1), c1: Math.min(out.c1, m.c1), r2: Math.max(out.r2, m.r2), c2: Math.max(out.c2, m.c2) }; });
    return out;
  }

  function select(r, c, extend = false) {
    r = Math.max(0, Math.min(MAX_ROWS - 1, r)); c = Math.max(0, Math.min(MAX_COLS - 1, c));
    ensureSize(r, c);
    const m = mergeAt(r, c);
    if (m && !extend) { r = m.r1; c = m.c1; }
    active = { r, c };
    if (!extend) anchor = { r, c };
    scrollIntoView(r, c);
    drawSelection();
  }
  let paintedSel = [];
  /** Highlights the selection on the cells currently in the DOM and positions the box. */
  function paintSelection() {
    paintedSel.forEach(el => el.classList.remove('sel', 'hl'));
    paintedSel = [];
    const g = expandByMerges(selRange());
    const multi = g.r1 !== g.r2 || g.c1 !== g.c2;
    for (const [r, tr] of rowEls) {
      if (r < g.r1 || r > g.r2) continue;
      paintedSel.push(tr.cells[0]); tr.cells[0].classList.add('hl');
      if (!multi) continue;
      for (const [c, idx] of colPos) {
        if (c < g.c1 || c > g.c2 || (r === active.r && c === active.c)) continue;
        const t = tr.cells[idx];
        if (t) { t.classList.add('sel'); paintedSel.push(t); }
      }
    }
    const hdr = thead.rows[0];
    for (const [c, idx] of colPos) if (c >= g.c1 && c <= g.c2 && hdr.cells[idx]) { hdr.cells[idx].classList.add('hl'); paintedSel.push(hdr.cells[idx]); }
    const b = rangeBox(g);
    Object.assign(selBox.style, { display: 'block', left: (b.x - 1) + 'px', top: (b.y - 1) + 'px', width: (b.w + 1) + 'px', height: (b.h + 1) + 'px' });
  }
  function drawSelection() {
    paintSelection();
    nameBox.value = rangeText(selRange());
    const sp = !cellAt(active.r, active.c) && spillBoxes.find(b => active.r >= b.row && active.r < b.row + b.h && active.c >= b.col && active.c < b.col + b.w);
    if (document.activeElement !== fx) fx.value = sp ? rawAt(sp.row, sp.col) : rawAt(active.r, active.c);
    fx.classList.toggle('ghost', !!sp);
    drawSpillBox();
    syncToolbar();
    updateAggregate();
    updatePivotPane();
    updateDvUi();
    updateProtectUi();
    drawPeers();
    if (live) live.moved();
  }
  /** Excel's blue outline around the spill range of the selected cell. */
  const spillBox = document.createElement('div');
  spillBox.className = 'spill-box';
  spillBox.hidden = true;
  document.getElementById('gridInner').appendChild(spillBox);
  function drawSpillBox() {
    const sp = spillBoxes.find(b => active.r >= b.row && active.r < b.row + b.h && active.c >= b.col && active.c < b.col + b.w);
    spillBox.hidden = !sp;
    if (!sp) return;
    const b = rangeBox({ r1: sp.row, c1: sp.col, r2: sp.row + sp.h - 1, c2: sp.col + sp.w - 1 });
    Object.assign(spillBox.style, { left: b.x + 'px', top: b.y + 'px', width: b.w + 'px', height: b.h + 'px' });
  }
  function scrollIntoView(r, c) {
    const z = zoomLevel();
    const s = sheet();
    const top = HEAD_H + (s.freezeRow ? rowH(0) : 0);
    const left = RH_W + (s.freezeCol ? colWidth(0) : 0);
    const y0 = HEAD_H + rowTop(r), y1 = y0 + rowH(r);
    const x0 = RH_W + colX[c], x1 = RH_W + colX[c + 1];
    const viewTop = wrap.scrollTop / z, viewH = wrap.clientHeight / z;
    const viewLeft = wrap.scrollLeft / z, viewW = wrap.clientWidth / z;
    if (!(s.freezeRow && r === 0)) {
      if (y1 > viewTop + viewH) wrap.scrollTop = (y1 - viewH + 2) * z;
      else if (y0 < viewTop + top) wrap.scrollTop = Math.max(0, (y0 - top) * z);
    }
    if (!(s.freezeCol && c === 0)) {
      if (x1 > viewLeft + viewW) wrap.scrollLeft = (x1 - viewW + 2) * z;
      else if (x0 < viewLeft + left) wrap.scrollLeft = Math.max(0, (x0 - left) * z);
    }
    renderWindow(false);
  }

  // ---------- Rendering ----------
  let win = null;                 // rendered window {r0, r1, c0, c1}
  let rowEls = new Map();         // row -> <tr>
  let colPos = new Map();         // col -> cell index within a row
  let showFormulas = false;

  function neededWindow() {
    const z = zoomLevel();
    const s = sheet();
    const y = wrap.scrollTop / z - HEAD_H, x = wrap.scrollLeft / z - RH_W;
    const r0 = rowAtY(Math.max(0, y)), r1 = rowAtY(Math.max(0, y + wrap.clientHeight / z));
    const c0 = colAtX(Math.max(0, x)), c1 = colAtX(Math.max(0, x + wrap.clientWidth / z));
    return { r0, r1: Math.min(s.rows - 1, r1), c0, c1: Math.min(s.cols - 1, c1) };
  }
  /** Re-renders the visible rows/columns when the view moved outside what is rendered. */
  function renderWindow(force) {
    const s = sheet();
    const need = neededWindow();
    if (!force && win && need.r0 >= win.r0 && need.r1 <= win.r1 && need.c0 >= win.c0 && need.c1 <= win.c1) return;
    win = {
      r0: Math.max(0, need.r0 - ROW_BUFFER), r1: Math.min(s.rows - 1, need.r1 + ROW_BUFFER),
      c0: Math.max(0, need.c0 - COL_BUFFER), c1: Math.min(s.cols - 1, need.c1 + COL_BUFFER)
    };
    // Column layout: [row header] [frozen col 0] [left spacer] [cols c0..c1] [right spacer]
    const items = [];
    let gapFrom = 0;
    if (s.freezeCol && win.c0 > 0) { items.push({ c: 0 }); gapFrom = 1; }
    if (colX[win.c0] - colX[gapFrom] > 0) items.push({ sp: colX[win.c0] - colX[gapFrom] });
    for (let c = win.c0; c <= win.c1; c++) items.push({ c });
    if (colX[s.cols] - colX[win.c1 + 1] > 0) items.push({ sp: colX[s.cols] - colX[win.c1 + 1] });
    colPos = new Map();
    items.forEach((it, i) => { if (it.c !== undefined) colPos.set(it.c, i + 1); });
    const fc = c => (s.freezeCol && c === 0 ? ' class="fc"' : '');
    const cellsHtml = items.map(it => (it.sp !== undefined ? '<td class="hs"></td>' : `<td data-c="${it.c}"${fc(it.c)}></td>`)).join('');

    let colgroup = grid.querySelector('colgroup');
    if (colgroup) colgroup.remove();
    colgroup = document.createElement('colgroup');
    colgroup.innerHTML = `<col style="width:${RH_W}px">` + items.map(it => `<col style="width:${it.sp !== undefined ? it.sp : colWidth(it.c)}px">`).join('');
    grid.prepend(colgroup);
    grid.style.width = (RH_W + colX[s.cols]) + 'px';

    let headHtml = '<tr class="hdr"><th class="corner" title="Select all"></th>' + items.map(it => (it.sp !== undefined
      ? '<th class="hs"></th>'
      : `<th data-c="${it.c}"${fc(it.c)}>${indexToCol(it.c)}<span class="resizer" data-c="${it.c}"></span></th>`)).join('') + '</tr>';
    const firstBodyRow = s.freezeRow ? Math.max(1, win.r0) : win.r0;
    if (s.freezeRow) headHtml += `<tr class="frozen" data-r="0" style="height:${rowH(0)}px"><th>1</th>${cellsHtml}</tr>`;
    thead.innerHTML = headHtml;

    const span = items.length + 1;
    const rowsHtml = [];
    const topGap = rowTop(firstBodyRow) - rowTop(s.freezeRow ? 1 : 0);
    if (topGap > 0) rowsHtml.push(`<tr class="sp"><td colspan="${span}" style="height:${topGap}px"></td></tr>`);
    for (let r = firstBodyRow; r <= win.r1; r++) {
      if (hiddenSet.has(r)) continue;
      rowsHtml.push(`<tr data-r="${r}" style="height:${rowH(r)}px"><th>${r + 1}</th>${cellsHtml}</tr>`);
    }
    const bottomGap = rowTop(s.rows) - rowTop(win.r1 + 1);
    if (bottomGap > 0) rowsHtml.push(`<tr class="sp"><td colspan="${span}" style="height:${bottomGap}px"></td></tr>`);
    tbody.innerHTML = rowsHtml.join('');

    rowEls = new Map();
    for (const tr of [...thead.rows, ...tbody.rows]) if (tr.dataset.r !== undefined) rowEls.set(+tr.dataset.r, tr);
    applyMerges();
    renderVisibleCells();
    paintSelection();
  }
  const tdAt = (r, c) => { const tr = rowEls.get(r); const i = colPos.get(c); return tr && i !== undefined ? tr.cells[i] : null; };
  /** Where a merge is drawn: its top-left cell, or the first rendered cell inside it. */
  function mergeDrawCell(m) {
    let r = m.r1;
    while (r <= m.r2 && !rowEls.has(r)) r++;
    let c = m.c1;
    while (c <= m.c2 && !colPos.has(c)) c++;
    return r <= m.r2 && c <= m.c2 ? { r, c } : null;
  }
  function applyMerges() {
    sheet().merges.forEach(m => {
      const at = mergeDrawCell(m);
      if (!at) return;
      let rows = 0, cols = 0;
      for (let r = at.r; r <= m.r2; r++) if (rowEls.has(r)) rows++;
      for (let c = at.c; c <= m.c2; c++) if (colPos.has(c)) cols++;
      const t = tdAt(at.r, at.c);
      t.rowSpan = rows; t.colSpan = cols;
      for (let r = at.r; r <= m.r2; r++) for (let c = at.c; c <= m.c2; c++) if (r !== at.r || c !== at.c) tdAt(r, c)?.classList.add('hidden-merged');
    });
  }

  /** The text and alignment class a cell shows. */
  function cellDisplay(r, c, s = sheet()) {
    const cell = cellAt(r, c, s);
    if (!cell) {
      // Part of a formula's spilled result?
      const v = engine.value(s.name, r, c);
      if (v === '') return { text: '', cls: '' };
      if (v instanceof F.FormulaError) return { text: v.error, cls: 'err' };
      if (typeof v === 'number') return { text: F.formatValue(v, 'General'), cls: 'num' };
      if (typeof v === 'boolean') return { text: v ? 'TRUE' : 'FALSE', cls: 'bool' };
      return { text: String(v), cls: '' };
    }
    const st = cell.s || {};
    const v = s === sheet() ? valueAt(r, c) : engine.value(s.name, r, c);
    if (showFormulas && typeof cell.v === 'string' && cell.v[0] === '=') return { text: cell.v, cls: '' };
    if (v instanceof F.FormulaError) return { text: v.error, cls: 'err' };
    if (typeof v === 'number') return st.fmt === '@' ? { text: String(cell.v), cls: '' } : { text: F.formatValue(v, st.fmt || 'General'), cls: 'num' };
    if (typeof v === 'boolean') return { text: v ? 'TRUE' : 'FALSE', cls: 'bool' };
    return { text: String(v), cls: '' };
  }
  let paintedCells = [];
  let spillBoxes = [];
  function renderCell(r, c) {
    const m = mergeAt(r, c);
    let t;
    if (m) {
      if (r !== m.r1 || c !== m.c1) return;
      const at = mergeDrawCell(m);
      t = at && tdAt(at.r, at.c);
    } else t = tdAt(r, c);
    if (!t) return;
    const cell = cellAt(r, c);
    const st = (cell && cell.s) || {};
    const { text, cls } = cellDisplay(r, c);
    const fmt = cf.format(r, c, valueAt(r, c));
    t.textContent = '';
    if (text !== '') {
      const d = document.createElement('div');
      d.className = 'cv';
      d.textContent = text;
      if (fmt && fmt.icon) { const ic = document.createElement('span'); ic.className = 'cf-icon'; ic.textContent = fmt.icon[0]; ic.style.color = fmt.icon[1]; d.prepend(ic); }
      // Rows never grow past their height, so the computed geometry stays exact.
      let h = 0;
      if (m) { for (let rr = m.r1; rr <= m.r2; rr++) h += rowH(rr); } else h = rowH(r);
      d.style.maxHeight = Math.max(0, h - 1) + 'px';
      t.appendChild(d);
    }
    t.className = [cls, st.wrap ? 'wrap' : '', s_fc(c), cell && cell.n ? 'has-note' : ''].filter(Boolean).join(' ');
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
    if (fmt) applyCf(style, fmt);
    if (circleInvalid) { const rule = dv.ruleAt(r, c); if (rule && !dv.check(rule, r, c, valueAt(r, c))) t.classList.add('dv-bad'); }
    const s = sheet();
    if (s.hiddenRows.length && r === 0 && s.filterCol === c) t.firstChild ? (t.firstChild.textContent += ' ⏷') : (t.textContent = '⏷');
    paintedCells.push(t);
  }
  /** Conditional formatting drawn over the cell's own style. */
  function applyCf(style, fmt) {
    if (fmt.bg) style.backgroundColor = fmt.bg;
    if (fmt.color) style.color = fmt.color;
    if (fmt.b) style.fontWeight = '700';
    if (fmt.i) style.fontStyle = 'italic';
    if (fmt.bar) style.backgroundImage = `linear-gradient(90deg, ${fmt.bar.color} 0%, ${fmt.bar.color}55 ${fmt.bar.pct}%, transparent ${fmt.bar.pct}%)`;
  }
  const s_fc = c => (sheet().freezeCol && c === 0 ? 'fc' : '');
  function renderVisibleCells() {
    engine.invalidate();
    cf.begin();
    paintedCells.forEach(t => { if (t.isConnected) { t.textContent = ''; t.removeAttribute('style'); t.className = t.classList.contains('hidden-merged') ? 'hidden-merged' : s_fc(+t.dataset.c); } });
    paintedCells = [];
    const s = sheet();
    for (const r of rowEls.keys()) {
      for (const c of colPos.keys()) if (s.cells[addr(r, c)]) renderCell(r, c);
    }
    // Cells that formulas spill into.
    spillBoxes = engine.spillsOn(s.name);
    for (const sp of spillBoxes) {
      for (let r = sp.row; r < sp.row + sp.h; r++) {
        if (!rowEls.has(r)) continue;
        for (let c = sp.col; c < sp.col + sp.w; c++) if ((r !== sp.row || c !== sp.col) && colPos.has(c)) renderCell(r, c);
      }
    }
    // Merged cells whose top-left is scrolled away still show their value.
    s.merges.forEach(m => { if (!rowEls.has(m.r1) || !colPos.has(m.c1)) renderCell(m.r1, m.c1); });
  }
  function buildGrid() {
    const s = sheet();
    s.rowHeights = s.rowHeights || {};
    s.hiddenRows = s.hiddenRows || [];
    rebuildCols();
    rebuildRows();
    document.getElementById('freezeRow').classList.toggle('on', !!s.freezeRow);
    document.getElementById('freezeCol').classList.toggle('on', !!s.freezeCol);
    renderWindow(true);
    renderTabs();
    renderCharts();
    drawSelection();
  }
  function ensureSize(r, c) {
    const s = sheet();
    if (r < s.rows - 1 && c < s.cols - 1) return;
    s.rows = Math.min(MAX_ROWS, Math.max(s.rows, r + 50));
    s.cols = Math.min(MAX_COLS, Math.max(s.cols, c + 5));
    rebuildCols();
    rebuildRows();
    renderWindow(true);
  }
  function changed(markDirty = true) {
    renderVisibleCells();
    drawNotes();
    renderCharts();
    drawSelection();
    if (markDirty) state.markDirty();
  }

  // Row heights: wrapped text and large fonts make a row taller, as in Excel.
  const measurer = document.createElement('div');
  Object.assign(measurer.style, { position: 'absolute', visibility: 'hidden', left: '-10000px', top: '0', whiteSpace: 'pre-wrap', wordBreak: 'break-word', lineHeight: '1.25', fontFamily: 'Calibri, Carlito, Arial, sans-serif', fontSize: '11pt', padding: '0 3px', boxSizing: 'border-box' });
  document.body.appendChild(measurer);
  function neededHeight(r, c, cell) {
    const st = cell.s || {};
    const size = st.size || 11;
    let h = size > 11 ? Math.ceil(size * 1.3333 * 1.2) + 4 : ROW_H;
    if (st.wrap && cell.v !== '' && cell.v !== undefined) {
      measurer.style.width = colWidth(c) + 'px';
      measurer.style.fontSize = size + 'pt';
      measurer.style.fontFamily = st.font || 'Calibri, Carlito, Arial, sans-serif';
      measurer.textContent = cellDisplay(r, c).text;
      h = Math.max(h, measurer.offsetHeight + 4);
    }
    return Math.min(545, h);
  }
  /** Recomputes the height of the given rows (or of every row with tall content when rows is null). */
  function fitRows(rows) {
    const s = sheet();
    const need = new Map();
    for (const [k, cell] of Object.entries(s.cells)) {
      const st = cell.s;
      if (!st || !(st.wrap || st.size > 11)) continue;
      const p = parseRef(k);
      if (rows && !rows.has(p.row)) continue;
      if (mergeAt(p.row, p.col)) continue;
      need.set(p.row, Math.max(need.get(p.row) || ROW_H, neededHeight(p.row, p.col, cell)));
    }
    if (rows) rows.forEach(r => { if (!need.has(r)) delete s.rowHeights[r]; });
    need.forEach((h, r) => { if (h > ROW_H) s.rowHeights[r] = h; else delete s.rowHeights[r]; });
    rebuildRows();
    renderWindow(true);
  }
  const rowsOf = g => { const set = new Set(); for (let r = g.r1; r <= g.r2; r++) set.add(r); return set; };

  function renderTabs() {
    const box = document.getElementById('sheetTabs');
    box.innerHTML = '';
    book.sheets.forEach((s, i) => {
      const t = document.createElement('div');
      t.className = 'tab' + (i === book.active ? ' active' : '');
      t.textContent = s.name;
      t.onclick = () => { if (i !== book.active) { commitEdit(); book.active = i; active = anchor = { r: 0, c: 0 }; wrap.scrollTop = wrap.scrollLeft = 0; buildGrid(); } };
      t.ondblclick = () => renameSheet(i);
      t.oncontextmenu = async e => {
        e.preventDefault();
        const r = await UI.form(`Sheet "${s.name}"`, [{ name: 'a', label: 'Action', type: 'select', options: ['Rename', 'Duplicate', 'Delete', 'Move left', 'Move right'] }]);
        if (!r) return;
        if (r.a === 'Rename') renameSheet(i);
        if (r.a === 'Duplicate') { snapshot(); const copy = JSON.parse(JSON.stringify(s)); copy.name = uniqueName(s.name.slice(0, 26) + ' (2)'); book.sheets.splice(i + 1, 0, copy); book.active = i + 1; buildGrid(); state.markDirty(); }
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
  /** Sheet names Excel accepts: no \ / ? * [ ] :, 1–31 characters. */
  const cleanSheetName = n => String(n || '').replace(/[\\/?*[\]:]/g, '_').replace(/^'+|'+$/g, '').slice(0, 31).trim();
  async function renameSheet(i) {
    const name = await UI.prompt('Rename Sheet', 'Sheet name', book.sheets[i].name);
    if (!name || name === book.sheets[i].name) return;
    if (/[\\/?*[\]:]/.test(name) || name.length > 31 || /^'|'$/.test(name)) return UI.toast('Sheet names can\'t contain \\ / ? * [ ] :, start or end with \', or be longer than 31 characters', 4000);
    if (book.sheets.some((s, j) => j !== i && s.name.toLowerCase() === name.toLowerCase())) return UI.toast('That name is already taken');
    snapshot();
    const old = book.sheets[i].name;
    book.sheets[i].name = name;
    // Update references to the renamed sheet in every formula.
    const quote = n => /^[A-Za-z_][\w.]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`;
    const esc = t => t.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const re = new RegExp(`(^|[^\\w.'])(${esc(old)}|'${esc(old.replace(/'/g, "''"))}')!`, 'gi');
    book.sheets.forEach(s => Object.values(s.cells).forEach(cell => {
      if (typeof cell.v === 'string' && cell.v[0] === '=') cell.v = cell.v.replace(re, (m, p) => p + quote(name) + '!');
    }));
    book.sheets.forEach(s => (s.pivots || []).forEach(p => { if (p.src.sheet === old) p.src.sheet = name; }));
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
    wrap.scrollTop = wrap.scrollLeft = 0;
    buildGrid();
    state.markDirty();
  }
  document.getElementById('addSheet').addEventListener('click', addSheet);
  document.getElementById('insSheet').addEventListener('click', addSheet);

  // ---------- Editing ----------
  let editing = false;
  function startEdit(initial) {
    if (!guard({ r1: active.r, c1: active.c, r2: active.r, c2: active.c })) { fx.value = rawAt(active.r, active.c); return; }
    scrollIntoView(active.r, active.c);
    const m = mergeAt(active.r, active.c);
    const b = rangeBox(m || { r1: active.r, c1: active.c, r2: active.r, c2: active.c });
    editing = true;
    editor.hidden = false;
    Object.assign(editor.style, {
      left: (b.x - 1) + 'px', top: (b.y - 1) + 'px',
      minWidth: (b.w + 2) + 'px', width: (b.w + 2) + 'px', height: Math.max(b.h + 2, 24) + 'px'
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
      const before = cellAt(active.r, active.c);
      const prev = before ? JSON.parse(JSON.stringify(before)) : null;
      snapshot();
      setCell(active.r, active.c, { v });
      autoFormat(active.r, active.c, v);
      // A line break typed with Alt+Enter turns on wrapping, as in Excel.
      if (v.includes('\n') && v[0] !== '=') setCell(active.r, active.c, { s: { wrap: true } });
      ensureSize(active.r, active.c);
      fitRows(new Set([active.r]));
      changed();
      const rule = dv.ruleAt(active.r, active.c);
      if (rule && rule.error.show !== false && !dv.check(rule, active.r, active.c, valueAt(active.r, active.c))) {
        document.getElementById('stMode').textContent = 'Ready';
        refuseValue(rule, active.r, active.c, prev, editor.value, moveR, moveC);
        return;
      }
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
    if (e.key === 'Enter' && e.altKey) { e.preventDefault(); const p = editor.selectionStart; editor.setRangeText('\n', p, editor.selectionEnd, 'end'); fx.value = editor.value; return; }
    if (e.key === 'Enter') { e.preventDefault(); commitEdit(e.shiftKey ? -1 : 1, 0); }
    else if (e.key === 'Tab') { e.preventDefault(); commitEdit(0, e.shiftKey ? -1 : 1); }
    else if (e.key === 'Escape') { e.preventDefault(); cancelEdit(); }
    else if (/^Arrow/.test(e.key) && document.getElementById('stMode').textContent === 'Enter' && !isRefInsertPoint()) {
      e.preventDefault();
      const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key];
      commitEdit(...d);
    }
  });
  fx.addEventListener('input', () => { if (!editing) startEdit(fx.value); editor.value = fx.value; fx.focus(); showHint(fx); });
  fx.addEventListener('keydown', e => {
    if (hintKey(e)) return;
    if (e.key === 'Enter') { e.preventDefault(); editor.value = fx.value; editing = true; commitEdit(1, 0); }
    if (e.key === 'Escape') { e.preventDefault(); cancelEdit(); }
  });
  nameBox.addEventListener('keydown', e => {
    if (e.key !== 'Enter') return;
    e.preventDefault();
    const [a, b] = nameBox.value.trim().toUpperCase().split(':');
    const p = parseRef(a || ''), q = b ? parseRef(b) : null;
    if (!p || p.row >= MAX_ROWS || p.col >= MAX_COLS || (q && (q.row >= MAX_ROWS || q.col >= MAX_COLS))) return UI.toast('Enter a reference like B7 or A1:C5');
    select(p.row, p.col);
    if (q) { anchor = { r: p.row, c: p.col }; active = { r: q.row, c: q.col }; ensureSize(q.row, q.col); drawSelection(); }
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
  /** What the pointer is on: a cell, a row header, a column header or the corner. */
  function hit(e) {
    const el = e.target.closest('td,th');
    if (!el || !grid.contains(el)) return null;
    const tr = el.parentElement;
    if (tr.classList.contains('hdr')) {
      if (el.classList.contains('corner')) return { kind: 'corner' };
      if (el.dataset.c !== undefined) return { kind: 'col', c: +el.dataset.c };
      return null;
    }
    if (tr.dataset.r === undefined) return null;
    const r = +tr.dataset.r;
    if (el.tagName === 'TH') return { kind: 'row', r };
    if (el.dataset.c === undefined) return null;
    return { kind: 'cell', r, c: +el.dataset.c };
  }
  grid.addEventListener('mousedown', e => {
    if (e.button !== 0) return;
    if (e.target.classList.contains('resizer')) return startResize(e);
    const h = hit(e);
    if (!h) return;
    const s = sheet();
    if (h.kind === 'corner') { commitEdit(); anchor = { r: 0, c: 0 }; active = { r: s.rows - 1, c: s.cols - 1 }; drawSelection(); return; }
    if (h.kind === 'col') {
      commitEdit();
      if (e.shiftKey) { active = { r: s.rows - 1, c: h.c }; anchor = { r: 0, c: anchor.c }; }
      else { anchor = { r: 0, c: h.c }; active = { r: s.rows - 1, c: h.c }; }
      drawSelection();
      return;
    }
    if (h.kind === 'row') {
      commitEdit();
      if (e.shiftKey) { active = { r: h.r, c: s.cols - 1 }; anchor = { r: anchor.r, c: 0 }; }
      else { anchor = { r: h.r, c: 0 }; active = { r: h.r, c: s.cols - 1 }; }
      drawSelection();
      return;
    }
    if (editing && editor.value[0] === '=' && isRefInsertPoint()) {
      // Clicking while typing a formula inserts a reference.
      e.preventDefault();
      const pos = editor.selectionStart;
      editor.setRangeText(addr(h.r, h.c), pos, pos, 'end');
      refDrag = { start: pos, from: h };
      fx.value = editor.value;
      return;
    }
    commitEdit();
    if (e.shiftKey) { active = { r: h.r, c: h.c }; drawSelection(); }
    else select(h.r, h.c);
    dragging = true;
    wrap.focus();
    e.preventDefault();
  });
  grid.addEventListener('mousemove', e => {
    if (!refDrag && !dragging) return;
    const h = hit(e);
    if (!h || h.kind !== 'cell') return;
    if (refDrag) {
      const a = refDrag.from;
      const text = (a.r === h.r && a.c === h.c) ? addr(a.r, a.c) : `${addr(Math.min(a.r, h.r), Math.min(a.c, h.c))}:${addr(Math.max(a.r, h.r), Math.max(a.c, h.c))}`;
      editor.setRangeText(text, refDrag.start, editor.selectionStart, 'end');
      fx.value = editor.value;
      return;
    }
    if (h.r !== active.r || h.c !== active.c) { active = { r: h.r, c: h.c }; drawSelection(); }
  });
  document.addEventListener('mouseup', () => { dragging = false; if (refDrag) { refDrag = null; editor.focus(); } });
  grid.addEventListener('dblclick', e => {
    if (e.target.classList.contains('resizer')) return autoFitColumn(+e.target.dataset.c);
    const h = hit(e);
    if (h && h.kind === 'cell') startEdit();
  });
  function startResize(e) {
    e.preventDefault();
    const c = +e.target.dataset.c;
    const startX = e.clientX, startW = colWidth(c);
    const z = zoomLevel();
    const col = grid.querySelectorAll('colgroup col')[colPos.get(c)];
    let w = startW;
    const move = ev => { w = Math.max(20, Math.round(startW + (ev.clientX - startX) / z)); if (col) col.style.width = w + 'px'; };
    const up = () => {
      document.removeEventListener('mousemove', move);
      document.removeEventListener('mouseup', up);
      if (w === startW) return;
      snapshot();
      sheet().colWidths[c] = w;
      fitRows(null);
      buildGrid();
      state.markDirty();
    };
    document.addEventListener('mousemove', move);
    document.addEventListener('mouseup', up);
  }
  const measureCtx = document.createElement('canvas').getContext('2d');
  /** Width a cell's text needs, in px. */
  function textWidth(text, st = {}) {
    measureCtx.font = `${st.b ? 'bold ' : ''}${st.size || 11}pt ${st.font || 'Calibri, Carlito, Arial'}`;
    return measureCtx.measureText(text).width + 14;
  }
  function autoFitColumn(c) {
    let w = 40;
    for (const k of Object.keys(sheet().cells)) {
      const p = parseRef(k);
      if (p.col !== c || mergeAt(p.row, p.col)) continue;
      const text = cellDisplay(p.row, c).text;
      if (text) w = Math.max(w, textWidth(text, sheet().cells[k].s));
    }
    snapshot();
    sheet().colWidths[c] = Math.min(500, Math.ceil(w));
    buildGrid();
    state.markDirty();
  }
  let scrollQueued = false;
  wrap.addEventListener('scroll', () => {
    if (scrollQueued) return;
    scrollQueued = true;
    requestAnimationFrame(() => {
      scrollQueued = false;
      // Grow the sheet as the user scrolls near the end, like an endless grid.
      const s = sheet();
      let grew = false;
      if (wrap.scrollTop + wrap.clientHeight > wrap.scrollHeight - 300 && s.rows < MAX_ROWS) { s.rows = Math.min(MAX_ROWS, s.rows + 200); rebuildRows(); grew = true; }
      if (wrap.scrollLeft + wrap.clientWidth > wrap.scrollWidth - 300 && s.cols < MAX_COLS) { s.cols = Math.min(MAX_COLS, s.cols + 10); rebuildCols(); grew = true; }
      renderWindow(grew);
    });
  });

  // ---------- Keyboard ----------
  wrap.addEventListener('keydown', e => {
    if (editing) return;
    const mod = e.ctrlKey || e.metaKey;
    const k = e.key;
    if (!dvList.hidden && dvListKey(e)) return;
    if (e.altKey && k === 'ArrowDown' && !dvArrow.hidden) { e.preventDefault(); openDvList(); return; }
    const s = sheet();
    const moves = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (moves[k]) {
      e.preventDefault();
      let [dr, dc] = moves[k];
      if (mod) {
        // Jump to the edge of the data region.
        let r = active.r, c = active.c;
        const filled = (r2, c2) => rawAt(r2, c2) !== '';
        const u = usedBounds();
        const limR = Math.max(u.r + 1, active.r), limC = Math.max(u.c + 1, active.c);
        const startFilled = filled(r, c) && filled(r + dr, c + dc);
        do { r += dr; c += dc; } while (r >= 0 && c >= 0 && r <= limR && c <= limC && (startFilled ? filled(r + dr, c + dc) : !filled(r, c)));
        r = Math.max(0, Math.min(limR, r)); c = Math.max(0, Math.min(limC, c));
        if (e.shiftKey) { active = { r, c }; ensureSize(r, c); scrollIntoView(r, c); drawSelection(); } else select(r, c);
        return;
      }
      // Skip over hidden (filtered) rows.
      let nr = active.r + dr;
      while (dr && hiddenSet.has(nr)) nr += dr;
      const m = mergeAt(active.r, active.c);
      if (m && !e.shiftKey) { if (dr > 0) nr = m.r2 + 1; if (dc > 0) dc = m.c2 - active.c + 1; }
      if (e.shiftKey) {
        active = { r: Math.max(0, Math.min(MAX_ROWS - 1, nr)), c: Math.max(0, Math.min(MAX_COLS - 1, active.c + dc)) };
        ensureSize(active.r, active.c); scrollIntoView(active.r, active.c); drawSelection();
      } else select(nr, active.c + dc);
      return;
    }
    if (k === 'Enter') { e.preventDefault(); moveInSelection(e.shiftKey ? -1 : 1, 0); return; }
    if (k === 'Tab') { e.preventDefault(); moveInSelection(0, e.shiftKey ? -1 : 1); return; }
    if (k === 'F2' && e.shiftKey) { e.preventDefault(); editNote(); return; }
    if (k === 'F2') { e.preventDefault(); startEdit(); return; }
    if (k === 'Home') { e.preventDefault(); select(mod ? 0 : active.r, 0); return; }
    if (k === 'End' && mod) { e.preventDefault(); const u = usedBounds(); select(u.r, u.c); return; }
    if (k === 'PageDown' || k === 'PageUp') {
      e.preventDefault();
      const page = Math.max(1, Math.floor(wrap.clientHeight / zoomLevel() / ROW_H) - 1);
      select(active.r + (k === 'PageDown' ? page : -page), active.c);
      return;
    }
    if (k === 'Delete' || k === 'Backspace') { e.preventDefault(); clearRange('contents'); return; }
    if (mod && k.toLowerCase() === 'a') { e.preventDefault(); anchor = { r: 0, c: 0 }; active = { r: s.rows - 1, c: s.cols - 1 }; drawSelection(); return; }
    if (mod && k.toLowerCase() === 'b') { e.preventDefault(); toggleStyle('b'); return; }
    if (mod && k.toLowerCase() === 'i') { e.preventDefault(); toggleStyle('i'); return; }
    if (mod && k.toLowerCase() === 'u') { e.preventDefault(); toggleStyle('u'); return; }
    if (mod && k.toLowerCase() === 'd') { e.preventDefault(); fill('down'); return; }
    if (mod && k.toLowerCase() === 'r') { e.preventDefault(); fill('right'); return; }
    if (mod && k.toLowerCase() === 'z' && !Platform.desktop) { e.preventDefault(); e.shiftKey ? redo() : undo(); return; }
    if (mod && k.toLowerCase() === 'y' && !Platform.desktop) { e.preventDefault(); redo(); return; }
    if (mod && k === ';') { e.preventDefault(); if (!guard({ r1: active.r, c1: active.c, r2: active.r, c2: active.c })) return; snapshot(); setCell(active.r, active.c, { v: String(Math.floor(F.dateToSerial(new Date()))), s: { fmt: 'm/d/yyyy' } }); changed(); return; }
    if (e.altKey && k === '=') { e.preventDefault(); autoSum(); return; }
    if (k.length === 1 && !mod && !e.altKey) { e.preventDefault(); startEdit(k); }
  });
  // Input methods (Chinese, Japanese, Korean…) compose text; start editing so it lands in the cell.
  wrap.addEventListener('compositionstart', () => { if (!editing) { startEdit(''); } });
  function moveInSelection(dr, dc) {
    const g = selRange();
    if (g.r1 === g.r2 && g.c1 === g.c2) return select(active.r + dr, active.c + dc);
    let r = active.r + dr, c = active.c + dc;
    if (c > g.c2) { c = g.c1; r++; } if (c < g.c1) { c = g.c2; r--; }
    if (r > g.r2) { r = g.r1; c = dc ? c : c + 1; } if (r < g.r1) { r = g.r2; }
    if (c > g.c2) c = g.c1;
    active = { r, c };
    scrollIntoView(r, c);
    drawSelection();
  }

  // ---------- Clipboard ----------
  let clip = null;
  function rangeToTsv(g) {
    const lines = [];
    for (let r = g.r1; r <= g.r2; r++) {
      const row = [];
      for (let c = g.c1; c <= g.c2; c++) {
        let t = cellDisplay(r, c).text;
        if (/[\t\n"]/.test(t)) t = '"' + t.replace(/"/g, '""') + '"';
        row.push(t);
      }
      lines.push(row.join('\t'));
    }
    return lines.join('\n');
  }
  function copy(cut) {
    const g = opRange();
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
    if (!text) return;
    if (!guard({ r1: active.r, c1: active.c, r2: active.r + text.split('\n').length - 1, c2: active.c + Math.max(...text.split('\n').map(l => l.split('\t').length)) - 1 })) return;
    snapshot();
    const r0 = active.r, c0 = active.c;
    if (clip && text.replace(/\r/g, '') === clip.text.replace(/\r/g, '')) {
      const h = clip.cells.length, w = clip.cells[0].length;
      clip.cells.forEach((row, i) => row.forEach((cell, j) => {
        let v = cell.v ?? '';
        if (typeof v === 'string' && v[0] === '=' && !clip.cut) v = F.shiftFormula(v, r0 - clip.g.r1, c0 - clip.g.c1);
        setCell(r0 + i, c0 + j, { v, s: cell.s ? { ...cell.s } : undefined });
        if (!cell.s) { const cur = cellAt(r0 + i, c0 + j); if (cur) delete cur.s; if (cur && (cur.v === '' || cur.v === undefined)) delete sheet().cells[addr(r0 + i, c0 + j)]; }
      }));
      if (clip.cut) {
        const src = book.sheets[clip.sheet];
        forRange(clip.g, (r, c) => {
          const inside = book.active === clip.sheet && r >= r0 && r < r0 + h && c >= c0 && c < c0 + w;
          if (!inside) delete src.cells[addr(r, c)];
        });
        clip = null;
      }
      anchor = { r: r0 + h - 1, c: c0 + w - 1 };
    } else {
      const rows = parseDelimited(text.replace(/\r?\n$/, ''), text.includes('\t') ? '\t' : ',');
      rows.forEach((row, i) => row.forEach((v, j) => setCell(r0 + i, c0 + j, { v })));
      if (rows.length) { anchor = { r: r0 + rows.length - 1, c: c0 + Math.max(...rows.map(r => r.length)) - 1 }; }
    }
    ensureSize(anchor.r, anchor.c);
    fitRows(rowsOf({ r1: Math.min(r0, anchor.r), r2: Math.max(r0, anchor.r) }));
    buildGrid();
    state.markDirty();
  }
  const ownsClipboard = () => !(editing || document.activeElement === fx || document.activeElement === nameBox || isModalOpen());
  document.addEventListener('copy', e => { if (!ownsClipboard()) return; e.preventDefault(); e.clipboardData.setData('text/plain', copy(false)); UI.toast('Copied'); });
  document.addEventListener('cut', e => { if (!ownsClipboard()) return; e.preventDefault(); e.clipboardData.setData('text/plain', copy(true)); });
  document.addEventListener('paste', e => { if (!ownsClipboard()) return; e.preventDefault(); paste(e.clipboardData.getData('text/plain')); });
  const isModalOpen = () => !!document.querySelector('.modal-backdrop') || !!document.activeElement.closest('.findbar');

  // ---------- Formatting ----------
  function styleRange(fn) {
    if (!guard(opRange(), 'formatCells')) return;
    snapshot();
    const g = opRange();
    forRange(g, (r, c) => {
      const cur = { ...((cellAt(r, c) || {}).s || {}) };
      const next = fn(cur, r, c) || cur;
      Object.keys(next).forEach(k => { if (next[k] === undefined || next[k] === false || next[k] === '') delete next[k]; });
      const cell = cellAt(r, c);
      if (cell) { cell.s = next; if (!Object.keys(next).length) delete cell.s; if ((cell.v === undefined || cell.v === '') && !cell.s) delete sheet().cells[addr(r, c)]; }
      else if (Object.keys(next).length) sheet().cells[addr(r, c)] = { v: '', s: next };
    });
    fitRows(rowsOf(g));
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
    const g = opRange();
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
    const g = opRange();
    if (!guardSheet()) return;
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
    if (!guardSheet((count > 0 ? 'insert' : 'delete') + (axis === 'row' ? 'Rows' : 'Columns'))) return;
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
    const shiftKeys = obj => {
      const out = {};
      Object.entries(obj || {}).forEach(([k, val]) => { k = +k; if (count < 0 && k >= index && k < index - count) return; out[k >= index ? k + count : k] = val; });
      return out;
    };
    if (axis === 'row') {
      s.rows = Math.max(1, Math.min(MAX_ROWS, s.rows + count));
      s.rowHeights = shiftKeys(s.rowHeights);
      s.hiddenRows = [];
      delete s.filterCol;
    } else {
      s.cols = Math.max(1, Math.min(MAX_COLS, s.cols + count));
      s.colWidths = shiftKeys(s.colWidths);
    }
    s.merges = s.merges.map(m => {
      const k1 = axis === 'row' ? 'r1' : 'c1', k2 = axis === 'row' ? 'r2' : 'c2';
      const n = { ...m };
      if (count < 0 && m[k1] >= index && m[k2] < index - count) return null;
      if (m[k1] >= index) n[k1] = Math.max(index, m[k1] + count);
      if (m[k2] >= index) n[k2] = Math.max(n[k1], m[k2] + count);
      return n;
    }).filter(Boolean);
    s.cf = SheetsCF.shift(s.cf, axis, index, count, s.name);
    s.dv = SheetsCF.shift(s.dv, axis, index, count, s.name);
    shiftPivots(s, axis, index, count);
    buildGrid();
    state.markDirty();
  }
  document.getElementById('insRow').addEventListener('click', () => { const g = selRange(); shiftStructure('row', g.r1, Math.min(1000, g.r2 - g.r1 + 1)); });
  document.getElementById('insCol').addEventListener('click', () => { const g = selRange(); shiftStructure('col', g.c1, Math.min(100, g.c2 - g.c1 + 1)); });
  document.getElementById('delRow').addEventListener('click', () => { const g = selRange(); shiftStructure('row', g.r1, -(g.r2 - g.r1 + 1)); select(g.r1, active.c); });
  document.getElementById('delCol').addEventListener('click', () => { const g = selRange(); shiftStructure('col', g.c1, -(g.c2 - g.c1 + 1)); select(active.r, g.c1); });

  // ---------- Conditional formatting ----------
  const CF_STYLE_OPTIONS = [
    { value: 'red', label: 'Light red fill with dark red text' }, { value: 'yellow', label: 'Yellow fill with dark yellow text' },
    { value: 'green', label: 'Green fill with dark green text' }, { value: 'redFill', label: 'Light red fill' },
    { value: 'redText', label: 'Red text' }, { value: 'bold', label: 'Bold' }, { value: 'custom', label: 'Custom format…' }
  ];
  const styleField = { name: 'style', label: 'with', type: 'select', options: CF_STYLE_OPTIONS };
  async function cfStyle(key) {
    if (key !== 'custom') return { ...SheetsCF.PRESETS[key] };
    const r = await UI.form('Format Cells', [
      { name: 'bg', label: 'Fill colour', type: 'color', value: '#ffc7ce' }, { name: 'color', label: 'Font colour', type: 'color', value: '#9c0006' },
      { name: 'b', label: 'Font style', type: 'select', options: [{ value: '', label: 'Regular' }, { value: 'b', label: 'Bold' }, { value: 'i', label: 'Italic' }, { value: 'bi', label: 'Bold italic' }] }
    ]);
    if (!r) return null;
    return { bg: r.bg, color: r.color, ...(r.b.includes('b') ? { b: true } : {}), ...(r.b.includes('i') ? { i: true } : {}) };
  }
  function addCfRule(rule) {
    const g = selRange();
    const clean = SheetsCF.clean({ ...rule, ref: SheetsCF.refOf(g) });
    if (!clean) { UI.toast('That rule could not be added'); return; }
    snapshot();
    const s = sheet();
    s.cf = [clean, ...(s.cf || [])]; // a new rule takes priority, as in Excel
    changed();
  }
  async function conditionalFormat(cmd) {
    commitEdit();
    if (!guardSheet('formatCells')) return;
    const [kind, arg] = cmd.split(':');
    const ops = { gt: ['Greater Than', 'greaterThan', 'Format cells that are GREATER THAN'], lt: ['Less Than', 'lessThan', 'Format cells that are LESS THAN'], eq: ['Equal To', 'equal', 'Format cells that are EQUAL TO'] };
    if (ops[kind]) {
      const [title, op, label] = ops[kind];
      const r = await UI.form(title, [{ name: 'v', label, value: '' }, styleField]);
      if (!r || r.v.trim() === '') return;
      const style = await cfStyle(r.style); if (!style) return;
      return addCfRule({ type: 'cellIs', op, v1: r.v.trim(), style });
    }
    switch (kind) {
      case 'between': {
        const r = await UI.form('Between', [{ name: 'a', label: 'Format cells that are BETWEEN', value: '' }, { name: 'b', label: 'and', value: '' }, styleField]);
        if (!r || r.a.trim() === '' || r.b.trim() === '') return;
        const style = await cfStyle(r.style); if (!style) return;
        return addCfRule({ type: 'cellIs', op: 'between', v1: r.a.trim(), v2: r.b.trim(), style });
      }
      case 'text': {
        const r = await UI.form('Text That Contains', [{ name: 'op', label: 'Format cells whose text', type: 'select', options: [{ value: 'containsText', label: 'contains' }, { value: 'notContainsText', label: 'does not contain' }, { value: 'beginsWith', label: 'begins with' }, { value: 'endsWith', label: 'ends with' }] }, { name: 't', label: 'the text', value: '' }, styleField]);
        if (!r || !r.t) return;
        const style = await cfStyle(r.style); if (!style) return;
        return addCfRule({ type: r.op, text: r.t, style });
      }
      case 'dup': {
        const r = await UI.form('Duplicate Values', [{ name: 'k', label: 'Format cells that contain', type: 'select', options: [{ value: 'duplicateValues', label: 'Duplicate values' }, { value: 'uniqueValues', label: 'Unique values' }] }, styleField]);
        if (!r) return;
        const style = await cfStyle(r.style); if (!style) return;
        return addCfRule({ type: r.k, style });
      }
      case 'top': case 'toppct': case 'bottom': case 'bottompct': {
        const pct = kind.endsWith('pct'), bottom = kind.startsWith('bottom');
        const r = await UI.form(`${bottom ? 'Bottom' : 'Top'} ${pct ? '10%' : '10 Items'}`, [{ name: 'n', label: `Format cells that rank in the ${bottom ? 'BOTTOM' : 'TOP'}${pct ? ' (%)' : ''}`, type: 'number', min: 1, max: pct ? 100 : 1000, value: '10' }, styleField]);
        if (!r) return;
        const style = await cfStyle(r.style); if (!style) return;
        return addCfRule({ type: 'top10', rank: +r.n || 10, percent: pct, bottom, style });
      }
      case 'above': case 'below': {
        const r = await UI.form(kind === 'above' ? 'Above Average' : 'Below Average', [{ ...styleField, label: `Format cells that are ${kind.toUpperCase()} AVERAGE with` }]);
        if (!r) return;
        const style = await cfStyle(r.style); if (!style) return;
        return addCfRule({ type: 'aboveAverage', below: kind === 'below', style });
      }
      case 'bar': return addCfRule({ type: 'dataBar', color: arg });
      case 'scale': return addCfRule({ type: 'colorScale', colors: arg.split(',') });
      case 'icons': return addCfRule({ type: 'iconSet', set: arg });
      case 'formula': {
        const g = selRange();
        const r = await UI.form('New Formatting Rule', [{ name: 'f', label: `Format values where this formula is true (written for ${addr(g.r1, g.c1)}, the first selected cell)`, value: '=' }, styleField]);
        if (!r || r.f.replace(/^=/, '').trim() === '') return;
        try { F.parse(r.f.replace(/^=/, '')); } catch { UI.toast('That formula has an error'); return; }
        const style = await cfStyle(r.style); if (!style) return;
        return addCfRule({ type: 'expression', formula: r.f, style });
      }
      case 'clear-sel': case 'clear-sheet': {
        const s = sheet();
        if (!s.cf || !s.cf.length) { UI.toast('There are no rules to clear'); return; }
        snapshot();
        if (kind === 'clear-sheet') s.cf = [];
        else {
          const g = selRange();
          s.cf = s.cf.filter(rule => { const a = SheetsCF.area(rule.ref); return a.r2 < g.r1 || a.r1 > g.r2 || a.c2 < g.c1 || a.c1 > g.c2; });
        }
        changed();
        return;
      }
      case 'manage': return manageCfRules();
    }
  }
  function manageCfRules() {
    commitEdit();
    const s = sheet();
    let rules = (s.cf || []).map(r => ({ ...r }));
    const back = document.createElement('div');
    back.className = 'modal-backdrop';
    back.innerHTML = '<div class="modal cf-manage" role="dialog" aria-label="Conditional Formatting Rules Manager"><h2>Conditional Formatting Rules</h2><div class="cf-list"></div><div class="actions"><button class="cancel">Cancel</button><button class="primary ok">OK</button></div></div>';
    const list = back.querySelector('.cf-list');
    const draw = () => {
      list.innerHTML = '';
      if (!rules.length) { list.innerHTML = '<div class="cf-empty">No rules on this sheet. Use Conditional Formatting to add one.</div>'; return; }
      rules.forEach((r, i) => {
        const row = document.createElement('div');
        row.className = 'cf-row';
        row.innerHTML = '<div><span class="cf-swatch">AaBb</span><span class="cf-desc"></span></div><input aria-label="Applies to"><div><button data-a="up" title="Move up">▲</button> <button data-a="down" title="Move down">▼</button> <button data-a="del" title="Delete rule">✕</button></div>';
        row.querySelector('.cf-desc').textContent = SheetsCF.describe(r);
        const sw = row.querySelector('.cf-swatch');
        if (r.style) Object.assign(sw.style, { background: r.style.bg || '', color: r.style.color || '', fontWeight: r.style.b ? '700' : '', fontStyle: r.style.i ? 'italic' : '' });
        else if (r.type === 'colorScale') sw.style.background = `linear-gradient(90deg, ${r.colors.join(',')})`;
        else if (r.type === 'dataBar') sw.style.background = `linear-gradient(90deg, ${r.color} 60%, transparent 60%)`;
        else sw.textContent = SheetsCF.ICON_SETS[r.set].map(x => x[0]).join('');
        const input = row.querySelector('input');
        input.value = r.ref;
        input.addEventListener('change', () => { r.ref = input.value.trim().toUpperCase().replace(/\$/g, ''); });
        row.querySelectorAll('button').forEach(b => b.addEventListener('click', () => {
          const a = b.dataset.a;
          if (a === 'del') rules.splice(i, 1);
          else { const j = a === 'up' ? i - 1 : i + 1; if (j < 0 || j >= rules.length) return; [rules[i], rules[j]] = [rules[j], rules[i]]; }
          draw();
        }));
        list.appendChild(row);
      });
    };
    draw();
    const close = () => back.remove();
    back.querySelector('.cancel').addEventListener('click', close);
    back.addEventListener('keydown', e => { if (e.key === 'Escape') close(); });
    back.querySelector('.ok').addEventListener('click', () => {
      const cleaned = rules.map(SheetsCF.clean);
      if (cleaned.some(x => !x)) { UI.toast('Check the “Applies to” ranges (for example B2:B20)', 3500); return; }
      close();
      snapshot();
      s.cf = cleaned;
      changed();
    });
    document.body.appendChild(back);
    back.querySelector('.ok').focus();
  }
  document.getElementById('condFmt').addEventListener('change', e => { const v = e.target.value; e.target.value = ''; if (v) conditionalFormat(v); });
  document.getElementById('cfManage').addEventListener('click', manageCfRules);

  // ---------- Notes (cell comments) ----------
  const notePop = document.createElement('div');
  notePop.className = 'note-pop';
  notePop.hidden = true;
  inner.appendChild(notePop);
  const noteLayer = document.createElement('div');
  noteLayer.className = 'note-layer';
  inner.appendChild(noteLayer);
  let showAllNotes = false, noteHover = null;
  const author = () => { try { return localStorage.getItem('atlas-author') || 'Author'; } catch { return 'Author'; } };
  function placeNote(el, r, c, n) {
    el.textContent = '';
    const who = document.createElement('b'); who.textContent = (n.author || 'Author') + ':';
    const body = document.createElement('div'); body.textContent = n.text;
    el.append(who, body);
    const b = rangeBox({ r1: r, c1: c, r2: r, c2: c });
    Object.assign(el.style, { left: (b.x + b.w + 8) + 'px', top: Math.max(0, b.y - 4) + 'px' });
  }
  function drawNotes() {
    noteLayer.textContent = '';
    if (!showAllNotes) return;
    const s = sheet();
    for (const [k, cell] of Object.entries(s.cells)) {
      if (!cell.n) continue;
      const p = parseRef(k);
      if (!rowEls.has(p.row) || !colPos.has(p.col)) continue;
      const el = document.createElement('div'); el.className = 'note-pop';
      placeNote(el, p.row, p.col, cell.n);
      noteLayer.appendChild(el);
    }
  }
  grid.addEventListener('mouseover', e => {
    const h = hit(e);
    const cell = h && h.kind === 'cell' && cellAt(h.r, h.c);
    if (!cell || !cell.n || showAllNotes) { if (noteHover) { notePop.hidden = true; noteHover = null; } return; }
    if (noteHover === addr(h.r, h.c)) return;
    noteHover = addr(h.r, h.c);
    placeNote(notePop, h.r, h.c, cell.n);
    notePop.hidden = false;
  });
  grid.addEventListener('mouseleave', () => { notePop.hidden = true; noteHover = null; });
  async function editNote() {
    commitEdit();
    if (!guardSheet('objects')) return;
    const cur = (cellAt(active.r, active.c) || {}).n;
    const r = await UI.form(cur ? 'Edit Note' : 'New Note', [{ name: 't', label: `Note on ${addr(active.r, active.c)} (${author()})`, type: 'textarea', value: cur ? cur.text : '' }], 'Save');
    if (!r) return;
    snapshot();
    const k = addr(active.r, active.c);
    if (!r.t.trim()) { if (sheet().cells[k]) { delete sheet().cells[k].n; setCell(active.r, active.c, {}); } }
    else setCell(active.r, active.c, { n: { text: r.t.slice(0, 32767), author: cur ? cur.author : author(), date: new Date().toISOString() } });
    changed();
  }
  function deleteNotes() {
    if (!guardSheet('objects')) return;
    const g = opRange();
    let n = 0;
    snapshot();
    forRange(g, (r, c) => { const cell = cellAt(r, c); if (cell && cell.n) { delete cell.n; setCell(r, c, {}); n++; } });
    changed();
    UI.toast(n ? `Deleted ${n} note${n === 1 ? '' : 's'}` : 'No notes in the selection');
  }
  function gotoNote(dir) {
    const keys = Object.keys(sheet().cells).filter(k => sheet().cells[k].n).map(parseRef).sort((a, b) => a.row - b.row || a.col - b.col);
    if (!keys.length) return UI.toast('There are no notes on this sheet');
    const cur = active.r * 20000 + active.c;
    const list = dir > 0 ? keys : keys.slice().reverse();
    const next = list.find(p => (dir > 0 ? p.row * 20000 + p.col > cur : p.row * 20000 + p.col < cur)) || list[0];
    select(next.row, next.col);
    placeNote(notePop, next.row, next.col, cellAt(next.row, next.col).n);
    notePop.hidden = false; noteHover = addr(next.row, next.col);
  }
  document.getElementById('noteNew').addEventListener('click', editNote);
  document.getElementById('noteDel').addEventListener('click', deleteNotes);
  document.getElementById('notePrev').addEventListener('click', () => gotoNote(-1));
  document.getElementById('noteNext').addEventListener('click', () => gotoNote(1));
  document.getElementById('noteShowAll').addEventListener('click', e => { showAllNotes = !showAllNotes; e.currentTarget.classList.toggle('on', showAllNotes); notePop.hidden = true; drawNotes(); });

  // ---------- Sheet protection ----------
  const PROTECTED_MSG = 'The cell or chart you\'re trying to change is on a protected sheet. To make a change, unprotect the sheet (Review > Unprotect Sheet).';
  const isLocked = (r, c) => !(((cellAt(r, c) || {}).s || {}).unlocked);
  /** Whether cells in g may change on this sheet; `allow` names an action the protection may permit. */
  function guard(g, allow) {
    const p = sheet().protect;
    if (!p) return true;
    if (allow && p.allow[allow]) return true;
    const area = (g.r2 - g.r1 + 1) * (g.c2 - g.c1 + 1);
    let ok = true;
    if (area > 200000) ok = false;
    else for (let r = g.r1; r <= g.r2 && ok; r++) for (let c = g.c1; c <= g.c2; c++) if (isLocked(r, c)) { ok = false; break; }
    if (!ok) UI.toast(PROTECTED_MSG, 4500);
    return ok;
  }
  function guardSheet(allow) {
    const p = sheet().protect;
    if (!p || (allow && p.allow[allow])) return true;
    UI.toast(PROTECTED_MSG, 4500);
    return false;
  }
  async function protectSheet() {
    commitEdit();
    const s = sheet();
    if (s.protect) {
      const needs = s.protect.hash || s.protect.algo;
      const r = needs ? await UI.form('Unprotect Sheet', [{ name: 'pw', label: 'Password', type: 'password' }], 'Unprotect') : {};
      if (!r) return;
      if (needs && !(await SheetsProtect.check(s.protect, r.pw))) { UI.toast('The password you supplied is not correct.', 3500); return; }
      snapshot();
      delete s.protect;
      UI.toast('Sheet unprotected');
    } else {
      const yesNo = [{ value: '', label: 'No' }, { value: '1', label: 'Yes' }];
      const r = await UI.form('Protect Sheet', [
        { name: 'pw', label: 'Password to unprotect sheet (optional)', type: 'password' },
        { name: 'pw2', label: 'Reenter password', type: 'password' },
        { name: 'formatCells', label: 'Allow formatting cells', type: 'select', options: yesNo },
        { name: 'insertRows', label: 'Allow inserting and deleting rows', type: 'select', options: yesNo },
        { name: 'insertColumns', label: 'Allow inserting and deleting columns', type: 'select', options: yesNo },
        { name: 'sort', label: 'Allow sorting', type: 'select', options: yesNo }
      ], 'Protect');
      if (!r) return;
      if (r.pw !== r.pw2) { UI.toast('The passwords don\'t match.', 3500); return; }
      snapshot();
      s.protect = SheetsProtect.clean({
        hash: SheetsProtect.legacyHash(r.pw),
        allow: { formatCells: !!r.formatCells, insertRows: !!r.insertRows, deleteRows: !!r.insertRows, insertColumns: !!r.insertColumns, deleteColumns: !!r.insertColumns, sort: !!r.sort }
      });
      UI.toast('Sheet protected. Only unlocked cells can be changed.', 3500);
    }
    updateProtectUi();
    state.markDirty();
  }
  function updateProtectUi() {
    const p = !!sheet().protect;
    const b = document.getElementById('protectSheet');
    b.classList.toggle('on', p);
    b.querySelector('.lbl').textContent = p ? 'Unprotect Sheet' : 'Protect Sheet';
    document.body.classList.toggle('sheet-protected', p);
    const st = (cellAt(active.r, active.c) || {}).s || {};
    document.getElementById('lockCell').classList.toggle('on', !st.unlocked);
  }
  document.getElementById('protectSheet').addEventListener('click', protectSheet);
  document.getElementById('lockCell').addEventListener('click', () => {
    if (!guardSheet()) return;
    const g = opRange();
    const lock = !isLocked(active.r, active.c);
    snapshot();
    forRange(g, (r, c) => {
      const cell = cellAt(r, c);
      if (lock) { if (cell && cell.s) { delete cell.s.unlocked; if (!Object.keys(cell.s).length) delete cell.s; setCell(r, c, {}); } }
      else setCell(r, c, { s: { unlocked: true } });
    });
    changed();
    UI.toast(lock ? 'Cells locked' : 'Cells unlocked: they can be edited when the sheet is protected');
  });

  // ---------- Data validation ----------
  const dvArrow = document.getElementById('dvArrow');
  const dvList = document.getElementById('dvList');
  const dvPrompt = document.getElementById('dvPrompt');
  let circleInvalid = false;
  /** A typed value broke the cell's rule: Stop refuses it, Warning and Information ask. */
  async function refuseValue(rule, r, c, prev, typed, moveR, moveC) {
    const undoRefused = () => {
      const k = addr(r, c);
      if (prev) sheet().cells[k] = prev; else delete sheet().cells[k];
      const j = undoStack.pop(); if (j) undoBytes -= j.length;
      changed(false);
    };
    const msg = rule.error.text || SheetsDV.defaultError(rule);
    const title = rule.error.title || 'Atlas Sheets';
    const reopen = () => { select(r, c); startEdit(typed); };
    if (rule.error.style === 'stop') {
      undoRefused();
      const i = await UI.choose(title, msg, ['Retry', 'Cancel'], '⛔');
      if (i === 0) reopen(); else select(r, c);
      return;
    }
    if (rule.error.style === 'warning') {
      const i = await UI.choose(title, msg + '\n\nContinue?', ['Yes', 'No', 'Cancel'], '⚠️');
      if (i === 0) { select(r + moveR, c + moveC); return; }
      undoRefused();
      if (i === 1) reopen(); else select(r, c);
      return;
    }
    const i = await UI.choose(title, msg, ['OK', 'Cancel'], 'ℹ️');
    if (i === 0) { select(r + moveR, c + moveC); return; }
    undoRefused();
    select(r, c);
  }
  /** Shows the dropdown arrow and input message for the active cell. */
  function updateDvUi() {
    dvList.hidden = true;
    const rule = dv.ruleAt(active.r, active.c);
    const b = rangeBox({ r1: active.r, c1: active.c, r2: active.r, c2: active.c });
    const showArrow = !!rule && rule.type === 'list' && rule.dropdown;
    dvArrow.hidden = !showArrow;
    if (showArrow) Object.assign(dvArrow.style, { left: (b.x + b.w + 1) + 'px', top: (b.y + Math.max(0, b.h - 21)) + 'px' });
    const p = rule && rule.prompt;
    dvPrompt.hidden = !p;
    if (p) {
      dvPrompt.querySelector('b').textContent = p.title;
      dvPrompt.querySelector('span').textContent = p.text;
      Object.assign(dvPrompt.style, { left: (b.x + Math.min(b.w, 40)) + 'px', top: (b.y + b.h + 6) + 'px' });
    }
  }
  function openDvList() {
    const rule = dv.ruleAt(active.r, active.c);
    if (!rule || rule.type !== 'list') return;
    const items = dv.listItems(rule, active.r, active.c);
    dvList.textContent = '';
    const cur = String(valueAt(active.r, active.c)).toLowerCase();
    items.forEach(it => {
      const d = document.createElement('div');
      d.setAttribute('role', 'option');
      d.textContent = it;
      if (it.toLowerCase() === cur) d.className = 'on';
      d.addEventListener('mousedown', e => { e.preventDefault(); pickDv(it); });
      dvList.appendChild(d);
    });
    if (!items.length) { const d = document.createElement('div'); d.textContent = '(no items)'; d.style.color = '#888'; dvList.appendChild(d); }
    const b = rangeBox({ r1: active.r, c1: active.c, r2: active.r, c2: active.c });
    Object.assign(dvList.style, { left: b.x + 'px', top: (b.y + b.h + 1) + 'px', minWidth: (b.w + 19) + 'px' });
    dvList.hidden = false;
    dvPrompt.hidden = true;
    (dvList.querySelector('.on') || dvList.firstChild).scrollIntoView({ block: 'nearest' });
  }
  function pickDv(text) {
    dvList.hidden = true;
    if (!guard({ r1: active.r, c1: active.c, r2: active.r, c2: active.c })) return;
    snapshot();
    setCell(active.r, active.c, { v: typeof F.parseLiteral(text) !== 'string' && !/^(TRUE|FALSE)$/.test(text) ? text : (text[0] === '=' ? "'" + text : text) });
    changed();
    wrap.focus();
  }
  function dvListKey(e) {
    const opts = [...dvList.children];
    let i = opts.findIndex(o => o.classList.contains('on'));
    if (e.key === 'Escape') { dvList.hidden = true; e.preventDefault(); return true; }
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault();
      i = Math.max(0, Math.min(opts.length - 1, i + (e.key === 'ArrowDown' ? 1 : -1)));
      opts.forEach((o, j) => o.classList.toggle('on', j === i));
      opts[i] && opts[i].scrollIntoView({ block: 'nearest' });
      return true;
    }
    if (e.key === 'Enter' || e.key === 'Tab') { e.preventDefault(); if (opts[i] && opts[i].getAttribute('role')) pickDv(opts[i].textContent); else dvList.hidden = true; return true; }
    return false;
  }
  dvArrow.addEventListener('mousedown', e => { e.preventDefault(); e.stopPropagation(); if (dvList.hidden) openDvList(); else dvList.hidden = true; });
  document.addEventListener('mousedown', e => { if (!dvList.hidden && !dvList.contains(e.target) && e.target !== dvArrow) dvList.hidden = true; });

  function validationDialog() {
    commitEdit();
    if (!guardSheet()) return;
    const g = selRange();
    const ref = SheetsCF.refOf(g);
    const cur = dv.ruleAt(active.r, active.c) || { type: 'any', allowBlank: true, dropdown: true, error: { style: 'stop' } };
    const back = document.createElement('div');
    back.className = 'modal-backdrop';
    back.innerHTML = `<div class="modal dv-dialog" role="dialog" aria-label="Data Validation"><h2>Data Validation</h2>
      <div class="dv-tabs"><button data-p="0" class="on">Settings</button><button data-p="1">Input Message</button><button data-p="2">Error Alert</button></div>
      <div class="dv-page" data-p="0">
        <label>Allow<select name="type"><option value="any">Any value</option><option value="whole">Whole number</option><option value="decimal">Decimal</option><option value="list">List</option><option value="date">Date</option><option value="time">Time</option><option value="textLength">Text length</option><option value="custom">Custom</option></select></label>
        <label data-for="op">Data<select name="op">${SheetsDV.OPS.map(o => `<option value="${o}">${SheetsDV.OP_WORDS[o]}</option>`).join('')}</select></label>
        <label data-for="v1"><span>Minimum</span><input name="v1" spellcheck="false"></label>
        <label data-for="v2"><span>Maximum</span><input name="v2" spellcheck="false"></label>
        <label data-for="src">Source<input name="src" placeholder="Yes,No,Maybe  or  =$E$1:$E$9" spellcheck="false"></label>
        <label data-for="formula">Formula<input name="formula" placeholder="=LEN(A1)<=10" spellcheck="false"></label>
        <label class="chk" data-for="blank"><input type="checkbox" name="blank"> Ignore blank</label>
        <label class="chk" data-for="dropdown"><input type="checkbox" name="dropdown"> In-cell dropdown</label>
        <small style="color:var(--muted)">Applies to ${ref}</small>
      </div>
      <div class="dv-page" data-p="1" hidden>
        <label class="chk"><input type="checkbox" name="showPrompt"> Show input message when cell is selected</label>
        <label>Title<input name="ptitle" maxlength="32"></label>
        <label>Input message<textarea name="ptext" maxlength="255"></textarea></label>
      </div>
      <div class="dv-page" data-p="2" hidden>
        <label class="chk"><input type="checkbox" name="showError"> Show error alert after invalid data is entered</label>
        <label>Style<select name="estyle"><option value="stop">Stop</option><option value="warning">Warning</option><option value="information">Information</option></select></label>
        <label>Title<input name="etitle" maxlength="32"></label>
        <label>Error message<textarea name="etext" maxlength="255"></textarea></label>
      </div>
      <div class="actions"><button class="left clear">Clear All</button><button class="cancel">Cancel</button><button class="primary ok">OK</button></div></div>`;
    const q = n => back.querySelector(`[name="${n}"]`);
    q('type').value = cur.type;
    q('op').value = cur.op || 'between';
    if (cur.type === 'list') q('src').value = cur.v1;
    else if (cur.type === 'custom') q('formula').value = cur.v1;
    else { q('v1').value = cur.v1 || ''; q('v2').value = cur.v2 || ''; }
    q('blank').checked = cur.allowBlank !== false;
    q('dropdown').checked = cur.dropdown !== false;
    q('showPrompt').checked = !!cur.prompt;
    q('ptitle').value = cur.prompt ? cur.prompt.title : '';
    q('ptext').value = cur.prompt ? cur.prompt.text : '';
    q('showError').checked = cur.error.show !== false;
    q('estyle').value = cur.error.style;
    q('etitle').value = cur.error.title || '';
    q('etext').value = cur.error.text || '';
    const sync = () => {
      const t = q('type').value, op = q('op').value, between = /etween/.test(op);
      const show = { op: !['any', 'list', 'custom'].includes(t), v1: !['any', 'list', 'custom'].includes(t), v2: !['any', 'list', 'custom'].includes(t) && between, src: t === 'list', formula: t === 'custom', blank: t !== 'any', dropdown: t === 'list' };
      back.querySelectorAll('[data-for]').forEach(l => { l.hidden = !show[l.dataset.for]; });
      const unit = t === 'date' ? ' (date, e.g. 2024-12-31)' : t === 'time' ? ' (time, e.g. 17:30)' : '';
      back.querySelector('[data-for="v1"] span').textContent = (between ? 'Minimum' : 'Value') + unit;
      back.querySelector('[data-for="v2"] span').textContent = 'Maximum' + unit;
    };
    q('type').addEventListener('change', sync); q('op').addEventListener('change', sync);
    sync();
    back.querySelectorAll('.dv-tabs button').forEach(b => b.addEventListener('click', () => {
      back.querySelectorAll('.dv-tabs button').forEach(x => x.classList.toggle('on', x === b));
      back.querySelectorAll('.dv-page').forEach(pg => { pg.hidden = pg.dataset.p !== b.dataset.p; });
    }));
    const close = () => back.remove();
    const inside = rule => { const a = SheetsCF.area(rule.ref); return a.r1 >= g.r1 && a.r2 <= g.r2 && a.c1 >= g.c1 && a.c2 <= g.c2; };
    back.querySelector('.cancel').addEventListener('click', close);
    back.addEventListener('keydown', e => { if (e.key === 'Escape') close(); });
    back.querySelector('.clear').addEventListener('click', () => {
      close();
      snapshot();
      sheet().dv = (sheet().dv || []).filter(r => !inside(r));
      changed();
    });
    back.querySelector('.ok').addEventListener('click', () => {
      const t = q('type').value;
      const others = (sheet().dv || []).filter(r => !inside(r));
      if (t === 'any') { close(); snapshot(); sheet().dv = others; changed(); return; }
      const rule = SheetsDV.clean({
        ref, type: t, op: q('op').value,
        v1: t === 'list' ? q('src').value.trim() : t === 'custom' ? q('formula').value.trim() : q('v1').value.trim(),
        v2: q('v2').value.trim(), allowBlank: q('blank').checked, dropdown: q('dropdown').checked,
        prompt: q('showPrompt').checked ? { title: q('ptitle').value, text: q('ptext').value } : null,
        error: { style: q('estyle').value, title: q('etitle').value, text: q('etext').value, show: q('showError').checked }
      });
      if (!rule) { UI.toast(t === 'list' ? 'Enter the list items or a range for Source' : 'Fill in the values for this rule', 3500); return; }
      close();
      snapshot();
      sheet().dv = [rule, ...others]; // the new rule wins where it overlaps an older one
      changed();
    });
    document.body.appendChild(back);
    q('type').focus();
  }
  document.getElementById('dataValidation').addEventListener('click', validationDialog);
  document.getElementById('circleInvalid').addEventListener('click', e => {
    circleInvalid = !circleInvalid;
    e.currentTarget.classList.toggle('on', circleInvalid);
    renderVisibleCells();
    if (circleInvalid) {
      let n = 0;
      (sheet().dv || []).forEach(rule => { const a = SheetsCF.area(rule.ref); const u = usedBounds(); for (let r = a.r1; r <= Math.min(a.r2, u.r); r++) for (let c = a.c1; c <= Math.min(a.c2, u.c); c++) if (cellAt(r, c) && dv.ruleAt(r, c) === rule && !dv.check(rule, r, c, valueAt(r, c))) n++; });
      UI.toast(n ? `${n} cell${n === 1 ? '' : 's'} with invalid data` : 'No invalid data found');
    }
  });

  // ---------- PivotTables ----------
  const pivotPane = document.getElementById('pivotPane');
  const PIVOT_STYLE = {
    hdr: { b: true, bg: '#ddebf7', border: { b: 1 } }, total: { b: true, bg: '#ddebf7', border: { t: 1 } }, sub: { b: true }, label: {}, val: {}
  };
  const sheetByName = name => book.sheets.find(x => x.name.toLowerCase() === String(name).toLowerCase());
  /** The source range as a table of computed values; the first row holds the field names. */
  function pivotTable(p) {
    const src = sheetByName(p.src.sheet);
    if (!src) return null;
    const g = SheetsCF.area(p.src.ref);
    const u = usedBounds(src);
    const r2 = Math.min(g.r2, u.r), c2 = Math.min(g.c2, u.c);
    const at = (r, c) => { const v = engine.value(src.name, r, c); return v instanceof F.FormulaError ? v.error : v; };
    const headers = [];
    for (let c = g.c1; c <= c2; c++) { const h = at(g.r1, c); headers.push(h === '' ? 'Column' + (c - g.c1 + 1) : String(h)); }
    const rows = [];
    for (let r = g.r1 + 1; r <= r2; r++) { const row = []; for (let c = g.c1; c <= c2; c++) row.push(at(r, c)); rows.push(row); }
    return { headers, rows };
  }
  const rawFor = v => (typeof v === 'number' ? String(v) : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : (v && (v[0] === '=' || typeof F.parseLiteral(v) !== 'string') ? "'" + v : String(v)));
  /** Recomputes a pivot and rewrites the cells it occupies. */
  function renderPivot(p, host) {
    engine.invalidate();
    const table = pivotTable(p);
    const res = table ? SheetsPivot.compute(table, p) : { cells: [], error: `The source sheet "${p.src.sheet}" no longer exists.` };
    if (p.extent) for (let r = p.r; r <= p.extent.r2; r++) for (let c = p.c; c <= p.extent.c2; c++) delete host.cells[addr(r, c)];
    let cells = res.cells;
    const empty = !p.rows.length && !p.cols.length && !p.vals.length;
    if (empty || res.error) cells = [[{ v: res.error || `${p.name}: choose fields in the PivotTable Fields pane`, k: 'note' }]];
    cells.forEach((row, i) => row.forEach((cell, j) => {
      const st = cell.k === 'note' ? { i: true, color: '#605e5c' } : { ...PIVOT_STYLE[cell.k] };
      if (typeof cell.v === 'number' && !Number.isInteger(cell.v)) st.fmt = '#,##0.00';
      if (cell.v === '' && !Object.keys(st).length) return;
      setCell(p.r + i, p.c + j, { v: cell.v === '' ? '' : rawFor(cell.v), s: st }, host);
    }));
    // Widen (never narrow) the columns so labels and totals show in full.
    if (!empty && !res.error) {
      const need = {};
      cells.forEach(row => row.forEach((cell, j) => {
        if (cell.v === '') return;
        const text = typeof cell.v === 'number' ? F.formatValue(cell.v, Number.isInteger(cell.v) ? 'General' : '#,##0.00') : String(cell.v);
        need[j] = Math.max(need[j] || 0, textWidth(text, PIVOT_STYLE[cell.k]));
      }));
      Object.entries(need).forEach(([j, w]) => { const c = p.c + +j; host.colWidths[c] = Math.max(host.colWidths[c] || DEFAULT_COL_W, Math.min(400, Math.ceil(w))); });
    }
    p.extent = { r2: p.r + Math.max(0, cells.length - 1), c2: p.c + Math.max(0, ...cells.map(r => r.length - 1)) };
    return res;
  }
  function refreshPivots(quiet) {
    let n = 0;
    snapshot();
    book.sheets.forEach(s => (s.pivots || []).forEach(p => { const r = renderPivot(p, s); n++; if (r.error && !quiet) UI.toast(r.error, 4000); }));
    buildGrid(); // column widths may have changed
    state.markDirty();
    if (!quiet) UI.toast(n ? `Refreshed ${n} PivotTable${n === 1 ? '' : 's'}` : 'There are no PivotTables in this workbook');
  }
  const pivotAt = (r, c, s = sheet()) => (s.pivots || []).find(p => r >= p.r && c >= p.c && p.extent && r <= p.extent.r2 && c <= p.extent.c2);
  async function insertPivot() {
    commitEdit();
    if (!guardSheet('pivotTables')) return;
    const g = dataRegion();
    if (g.r2 <= g.r1) { UI.toast('Select a range with a header row and at least one row of data'); return; }
    const srcRef = `${sheet().name}!${addr(g.r1, g.c1)}:${addr(g.r2, g.c2)}`;
    const r = await UI.form('Create PivotTable', [
      { name: 'src', label: 'Table/Range', value: srcRef },
      { name: 'where', label: 'Place the PivotTable on', type: 'select', options: [{ value: 'new', label: 'New worksheet' }, { value: 'here', label: 'Existing worksheet (at the cell below)' }] },
      { name: 'at', label: 'Location (existing worksheet)', value: addr(0, Math.min(MAX_COLS - 1, g.c2 + 2)) }
    ], 'Create');
    if (!r) return;
    const m = /^(?:(.+)!)?\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)$/i.exec(r.src.trim());
    const srcSheet = m && (m[1] ? sheetByName(m[1].replace(/^'|'$/g, '').replace(/''/g, "'")) : sheet());
    if (!m || !srcSheet) { UI.toast('Enter the source as Sheet1!A1:D20'); return; }
    const ref = (m[2] + m[3] + ':' + m[4] + m[5]).toUpperCase();
    const ga = SheetsCF.area(ref);
    if (ga.r2 <= ga.r1) { UI.toast('The source needs a header row and at least one row of data'); return; }
    snapshot();
    let host, at;
    if (r.where === 'new') {
      host = newSheet(uniqueName('PivotTable'));
      book.sheets.splice(book.active + 1, 0, host);
      at = { r: 2, c: 0 };
    } else {
      host = sheet();
      const p = parseRef(r.at.trim().toUpperCase().replace(/\$/g, ''));
      if (!p) { UI.toast('Enter a cell such as H1 for the location'); return; }
      at = { r: p.row, c: p.col };
    }
    host.pivots = host.pivots || [];
    let n = 1;
    while (book.sheets.some(s => (s.pivots || []).some(x => x.name === 'PivotTable' + n))) n++;
    const p = SheetsPivot.clean({ name: 'PivotTable' + n, src: { sheet: srcSheet.name, ref: ref.includes(':') ? ref : ref + ':' + ref }, r: at.r, c: at.c, rows: [], cols: [], vals: [], filters: [] });
    host.pivots.push(p);
    renderPivot(p, host);
    book.active = book.sheets.indexOf(host);
    active = anchor = { r: at.r, c: at.c };
    buildGrid();
    changed();
  }
  let paneFor = null;
  function updatePivotPane(force) {
    const p = pivotAt(active.r, active.c);
    if (!p) { if (!pivotPane.hidden) { pivotPane.hidden = true; paneFor = null; } return; }
    if (!force && paneFor === p.id && !pivotPane.hidden) return;
    paneFor = p.id;
    pivotPane.hidden = false;
    drawPivotPane(p);
  }
  function drawPivotPane(p) {
    const table = pivotTable(p);
    const fields = table ? table.headers : [];
    pivotPane.textContent = '';
    const el = (tag, cls, text) => { const e = document.createElement(tag); if (cls) e.className = cls; if (text !== undefined) e.textContent = text; return e; };
    const head = el('div', 'pv-head');
    head.append(el('b', '', 'PivotTable Fields'), el('span', 'pv-name', p.name));
    const close = el('button', 'btn pv-close', '✕'); close.title = 'Close'; close.onclick = () => { pivotPane.hidden = true; };
    head.appendChild(close);
    pivotPane.appendChild(head);
    pivotPane.appendChild(el('div', 'pv-src', `Source: ${p.src.sheet}!${p.src.ref}`));
    const used = new Set([...p.rows, ...p.cols, ...p.vals.map(v => v.field), ...p.filters.map(f => f.field)]);
    const isNumeric = i => { let n = 0, t = 0; for (const r of table.rows) { const v = r[i]; if (v === '' || v === null) continue; t++; if (typeof v === 'number') n++; } return t > 0 && n / t >= 0.6; };
    const apply = fn => { snapshot(); fn(); renderPivot(p, sheet()); buildGrid(); state.markDirty(); updatePivotPane(true); };
    const list = el('div', 'pv-fields');
    fields.forEach((f, i) => {
      const lab = el('label', 'pv-field');
      const cb = el('input'); cb.type = 'checkbox'; cb.checked = used.has(f);
      cb.addEventListener('change', () => apply(() => {
        if (cb.checked) { if (isNumeric(i)) p.vals.push({ field: f, agg: 'sum' }); else p.rows.push(f); }
        else { p.rows = p.rows.filter(x => x !== f); p.cols = p.cols.filter(x => x !== f); p.vals = p.vals.filter(x => x.field !== f); p.filters = p.filters.filter(x => x.field !== f); }
      }));
      lab.append(cb, el('span', '', f));
      list.appendChild(lab);
    });
    pivotPane.appendChild(list);
    const AREAS = [['filters', 'Filters'], ['cols', 'Columns'], ['rows', 'Rows'], ['vals', 'Values']];
    const nameOf = (area, item) => (area === 'vals' || area === 'filters' ? item.field : item);
    const moveTo = (area, i, dest) => apply(() => {
      const item = p[area].splice(i, 1)[0];
      const f = nameOf(area, item);
      if (dest === 'remove') return;
      if (dest === 'up' || dest === 'down') { p[area].splice(Math.max(0, Math.min(p[area].length, i + (dest === 'up' ? -1 : 1))), 0, item); return; }
      p[dest].push(dest === 'vals' ? { field: f, agg: 'count' } : dest === 'filters' ? { field: f, value: null } : f);
    });
    const areas = el('div', 'pv-areas');
    AREAS.forEach(([area, title]) => {
      const box = el('div', 'pv-area');
      box.appendChild(el('div', 'pv-area-title', title));
      if (!p[area].length) box.appendChild(el('div', 'pv-empty', 'Tick a field or move one here'));
      p[area].forEach((item, i) => {
        const row = el('div', 'pv-item');
        row.appendChild(el('span', 'pv-item-name', area === 'vals' ? SheetsPivot.valName(item) : nameOf(area, item)));
        if (area === 'vals') {
          const agg = el('select', 'ctl'); agg.title = 'Summarize values by';
          SheetsPivot.AGGS.forEach(a => { const o = el('option', '', SheetsPivot.AGG_LABEL[a]); o.value = a; agg.appendChild(o); });
          agg.value = item.agg;
          agg.addEventListener('change', () => apply(() => { item.agg = agg.value; }));
          row.appendChild(agg);
        }
        if (area === 'filters') {
          const i2 = fields.indexOf(item.field);
          const valuesSel = el('select', 'ctl'); valuesSel.title = 'Show';
          const all = el('option', '', '(All)'); all.value = '\u0000all'; valuesSel.appendChild(all);
          const seen = new Map();
          if (table && i2 >= 0) table.rows.forEach(r => { const v = SheetsPivot.label(r[i2]); if (!seen.has(String(v)) && seen.size < 500) seen.set(String(v), v); });
          [...seen.keys()].sort((a, b) => a.localeCompare(b, undefined, { numeric: true })).forEach(k => { const o = el('option', '', k); o.value = k; valuesSel.appendChild(o); });
          valuesSel.value = item.value === null || item.value === undefined ? '\u0000all' : String(item.value);
          valuesSel.addEventListener('change', () => apply(() => { item.value = valuesSel.value === '\u0000all' ? null : (typeof seen.get(valuesSel.value) === 'number' ? seen.get(valuesSel.value) : valuesSel.value); }));
          row.appendChild(valuesSel);
        }
        const mv = el('select', 'ctl pv-move'); mv.title = 'Move';
        [['', '⋯'], ['up', 'Move up'], ['down', 'Move down'], ...AREAS.filter(a => a[0] !== area).map(a => [a[0], 'Move to ' + a[1]]), ['remove', 'Remove field']]
          .forEach(([v, t]) => { const o = el('option', '', t); o.value = v; mv.appendChild(o); });
        mv.addEventListener('change', () => { if (mv.value) moveTo(area, i, mv.value); });
        row.appendChild(mv);
        box.appendChild(row);
      });
      areas.appendChild(box);
    });
    pivotPane.appendChild(areas);
    const actions = el('div', 'pv-actions');
    const refresh = el('button', 'btn', '⟳ Refresh'); refresh.onclick = () => apply(() => {});
    const del = el('button', 'btn', '🗑 Delete PivotTable');
    del.onclick = () => {
      snapshot();
      const s = sheet();
      if (p.extent) for (let r = p.r; r <= p.extent.r2; r++) for (let c = p.c; c <= p.extent.c2; c++) delete s.cells[addr(r, c)];
      s.pivots = s.pivots.filter(x => x !== p);
      pivotPane.hidden = true; paneFor = null;
      changed();
    };
    actions.append(refresh, del);
    pivotPane.appendChild(actions);
  }
  /** Keeps pivots and their sources in place when rows/columns move. */
  function shiftPivots(s, axis, index, count) {
    const k = axis === 'row' ? 'r' : 'c', k2 = axis === 'row' ? 'r2' : 'c2';
    (s.pivots || []).forEach(p => {
      if (p[k] >= index) { p[k] = Math.max(index, p[k] + count); if (p.extent) p.extent[k2] = Math.max(p[k], p.extent[k2] + count); }
    });
    book.sheets.forEach(h => (h.pivots || []).forEach(p => {
      if (p.src.sheet.toLowerCase() !== s.name.toLowerCase()) return;
      const moved = SheetsCF.shift([{ ref: p.src.ref }], axis, index, count, s.name)[0];
      if (moved) p.src.ref = moved.ref.includes(':') ? moved.ref : moved.ref + ':' + moved.ref;
    }));
  }
  document.getElementById('insPivot').addEventListener('click', insertPivot);
  document.getElementById('insPivot2').addEventListener('click', insertPivot);
  document.getElementById('refreshAll').addEventListener('click', () => { commitEdit(); refreshPivots(false); });

  // ---------- Editing group ----------
  function clearRange(what) {
    if (!guard(opRange(), what === 'formats' ? 'formatCells' : null)) return;
    snapshot();
    const g = opRange();
    forRange(g, (r, c) => {
      const cell = cellAt(r, c);
      if (!cell) return;
      if (what === 'all') delete sheet().cells[addr(r, c)];
      else if (what === 'formats') { delete cell.s; if (!cell.v && !cell.n) delete sheet().cells[addr(r, c)]; }
      else { cell.v = ''; if (!cell.s && !cell.n) delete sheet().cells[addr(r, c)]; }
    });
    if (what !== 'contents') fitRows(rowsOf(g));
    changed();
  }
  document.getElementById('clear').addEventListener('change', e => { const v = e.target.value; e.target.value = ''; if (v) clearRange(v); wrap.focus(); });
  // ---------- Fill handle (Auto Fill) ----------
  const fillHandle = document.createElement('div');
  fillHandle.className = 'fill-handle';
  fillHandle.title = 'Drag to fill a series (Ctrl: copy instead); double-click to fill down';
  selBox.appendChild(fillHandle);
  const fillPreview = document.createElement('div');
  fillPreview.className = 'fill-preview';
  fillPreview.hidden = true;
  inner.appendChild(fillPreview);
  let fillDrag = null;
  const isDateFmt = f => !!f && /[dmy]/i.test(f.replace(/"[^"]*"|\[[^\]]*\]/g, ''));
  function cellAtPoint(e) {
    const b = inner.getBoundingClientRect(), z = zoomLevel();
    return { r: rowAtY(Math.max(0, (e.clientY - b.top) / z - HEAD_H)), c: colAtX(Math.max(0, (e.clientX - b.left) / z - RH_W)) };
  }
  fillHandle.addEventListener('mousedown', e => {
    if (e.button !== 0) return;
    e.preventDefault(); e.stopPropagation();
    commitEdit();
    fillDrag = { g: opRange(), target: null };
  });
  document.addEventListener('mousemove', e => {
    if (!fillDrag) return;
    const p = cellAtPoint(e), g = fillDrag.g;
    // The direction is the one the pointer has moved furthest outside the selection.
    const d = { down: p.r - g.r2, up: g.r1 - p.r, right: p.c - g.c2, left: g.c1 - p.c };
    const dir = Object.keys(d).reduce((a, b) => (d[b] > d[a] ? b : a));
    let t = null;
    if (d[dir] > 0) t = dir === 'down' ? { ...g, r2: p.r } : dir === 'up' ? { ...g, r1: p.r } : dir === 'right' ? { ...g, c2: p.c } : { ...g, c1: p.c };
    if (t) t.dir = dir;
    fillDrag.target = t;
    const b = rangeBox(t || g);
    Object.assign(fillPreview.style, { left: b.x + 'px', top: b.y + 'px', width: b.w + 'px', height: b.h + 'px' });
    fillPreview.hidden = !t;
  });
  document.addEventListener('mouseup', e => {
    if (!fillDrag) return;
    const { g, target } = fillDrag;
    fillDrag = null;
    fillPreview.hidden = true;
    if (target) autoFill(g, target, e.ctrlKey || e.metaKey);
  });
  fillHandle.addEventListener('dblclick', e => {
    e.preventDefault(); e.stopPropagation();
    // Fill down as far as the data in the column next to the selection goes.
    const g = opRange(), s = sheet();
    const filled = c => c >= 0 && rawAt(g.r2 + 1, c) !== '';
    const col = filled(g.c1 - 1) ? g.c1 - 1 : filled(g.c2 + 1) ? g.c2 + 1 : -1;
    if (col < 0) return;
    let r = g.r2 + 1;
    while (r + 1 < s.rows && rawAt(r + 1, col) !== '') r++;
    autoFill(g, { ...g, r2: r, dir: 'down' }, false);
  });
  /** Extends the source range g over t, continuing series and shifting formulas. */
  function autoFill(g, t, ctrl) {
    if (!guard(t)) return;
    snapshot();
    const s = sheet();
    const vertical = t.dir === 'down' || t.dir === 'up', reverse = t.dir === 'up' || t.dir === 'left';
    const n = vertical ? (t.r2 - t.r1) - (g.r2 - g.r1) : (t.c2 - t.c1) - (g.c2 - g.c1);
    const lines = vertical ? [g.c1, g.c2] : [g.r1, g.r2];
    for (let line = lines[0]; line <= lines[1]; line++) {
      const cellsOf = vertical ? Array.from({ length: g.r2 - g.r1 + 1 }, (_, i) => [g.r1 + i, line]) : Array.from({ length: g.c2 - g.c1 + 1 }, (_, i) => [line, g.c1 + i]);
      const src = cellsOf.map(([r, c]) => { const cell = cellAt(r, c) || {}; return { raw: cell.v ?? '', isDate: isDateFmt(cell.s && cell.s.fmt), s: cell.s }; });
      const k = src.length;
      const out = SheetsFill.extend(src, n, { reverse, increment: ctrl && k === 1, copy: ctrl && k > 1 });
      out.forEach((val, i) => {
        const pos = reverse ? -1 - i : k + i;
        const from = ((pos % k) + k) % k;
        const [r, c] = vertical ? [reverse ? g.r1 - 1 - i : g.r2 + 1 + i, line] : [line, reverse ? g.c1 - 1 - i : g.c2 + 1 + i];
        const [sr, sc] = cellsOf[typeof val === 'object' ? val.from : from];
        const v = typeof val === 'object' ? F.shiftFormula(val.formula, r - sr, c - sc) : val;
        const key = addr(r, c);
        delete s.cells[key];
        setCell(r, c, { v, s: src[from].s ? JSON.parse(JSON.stringify(src[from].s)) : undefined });
      });
    }
    anchor = { r: t.r1, c: t.c1 };
    active = { r: t.r2, c: t.c2 };
    ensureSize(t.r2, t.c2);
    changed();
  }
  async function seriesDialog() {
    commitEdit();
    const g = opRange();
    if (!guard(g)) return;
    const r = await UI.form('Series', [
      { name: 'dir', label: 'Series in', type: 'select', options: [{ value: 'cols', label: 'Columns' }, { value: 'rows', label: 'Rows' }] },
      { name: 'type', label: 'Type', type: 'select', options: [{ value: 'linear', label: 'Linear' }, { value: 'growth', label: 'Growth' }, { value: 'date', label: 'Date' }] },
      { name: 'unit', label: 'Date unit', type: 'select', options: [{ value: 'day', label: 'Day' }, { value: 'weekday', label: 'Weekday' }, { value: 'month', label: 'Month' }, { value: 'year', label: 'Year' }] },
      { name: 'step', label: 'Step value', value: '1' },
      { name: 'stop', label: 'Stop value (optional)', value: '' }
    ]);
    if (!r) return;
    const step = Number(r.step), stop = r.stop.trim() === '' ? null : Number(r.stop);
    if (!isFinite(step) || (stop !== null && !isFinite(stop))) { UI.toast('Enter numbers for the step and stop values'); return; }
    const start = valueAt(g.r1, g.c1);
    if (typeof start !== 'number') { UI.toast('Put the first value of the series in the first selected cell'); return; }
    const down = r.dir === 'cols';
    const single = g.r1 === g.r2 && g.c1 === g.c2;
    const len = single ? (stop === null ? 1 : 10000) : down ? g.r2 - g.r1 + 1 : g.c2 - g.c1 + 1;
    const dates = {
      add: (v, n, unit) => {
        const d = F.serialToDate(v);
        if (unit === 'day') return v + n;
        if (unit === 'weekday') { let x = v, left = Math.abs(n); const st = n < 0 ? -1 : 1; while (left > 0) { x += st; const wd = F.serialToDate(x).getUTCDay(); if (wd !== 0 && wd !== 6) left--; } return x; }
        const y = d.getUTCFullYear() + (unit === 'year' ? n : 0), m = d.getUTCMonth() + (unit === 'month' ? n : 0);
        return F.dateToSerial(new Date(y, m, Math.min(d.getUTCDate(), new Date(y, m + 1, 0).getDate())));
      }
    };
    const vals = SheetsFill.series(start, len, { type: r.type, step, stop, unit: r.unit }, dates);
    snapshot();
    const lines = down ? [g.c1, g.c2] : [g.r1, g.r2];
    for (let line = lines[0]; line <= lines[1]; line++) {
      const first = down ? valueAt(g.r1, line) : valueAt(line, g.c1);
      if (typeof first !== 'number') continue;
      const own = SheetsFill.series(first, vals.length, { type: r.type, step, stop, unit: r.unit }, dates);
      const st = (cellAt(down ? g.r1 : line, down ? line : g.c1) || {}).s;
      own.forEach((v, i) => { if (i) setCell(down ? g.r1 + i : line, down ? line : g.c1 + i, { v: String(v), s: st ? { ...st } : undefined }); });
    }
    ensureSize(g.r1 + (down ? vals.length : 1), g.c1 + (down ? 1 : vals.length));
    changed();
  }
  document.getElementById('fillSeries').addEventListener('click', seriesDialog);

  function fill(dir) {
    const g = opRange();
    if (dir === 'down' ? (g.r1 === g.r2 && g.r1 === 0) : (g.c1 === g.c2 && g.c1 === 0)) return;
    snapshot();
    if (dir === 'down') {
      const r0 = g.r1 === g.r2 ? g.r1 - 1 : g.r1;
      for (let c = g.c1; c <= g.c2; c++) {
        const src = cellAt(r0, c) || {};
        for (let r = r0 + 1; r <= g.r2; r++) {
          const v = typeof src.v === 'string' && src.v[0] === '=' ? F.shiftFormula(src.v, r - r0, 0) : src.v ?? '';
          setCell(r, c, { v, s: src.s ? { ...src.s } : undefined });
        }
      }
    } else {
      const c0 = g.c1 === g.c2 ? g.c1 - 1 : g.c1;
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
    if (!guard({ r1: g.r1, c1: g.c1, r2: g.r2 + 1, c2: g.c2 })) return;
    if (g.r1 !== g.r2 || g.c1 !== g.c2) {
      const o = opRange(g);
      snapshot();
      // Sum each selected column into the row below the selection.
      for (let c = o.c1; c <= o.c2; c++) setCell(o.r2 + 1, c, { v: `=SUM(${addr(o.r1, c)}:${addr(o.r2, c)})` });
      ensureSize(o.r2 + 1, o.c2);
      changed();
      select(o.r2 + 1, o.c1);
      return;
    }
    // Single cell: look up for a block of numbers, else left.
    let r = active.r - 1;
    while (r >= 0 && typeof valueAt(r, active.c) === 'number') r--;
    if (r < active.r - 1) { snapshot(); setCell(active.r, active.c, { v: `=SUM(${addr(r + 1, active.c)}:${addr(active.r - 1, active.c)})` }); changed(); return; }
    let c = active.c - 1;
    while (c >= 0 && typeof valueAt(active.r, c) === 'number') c--;
    if (c < active.c - 1) { snapshot(); setCell(active.r, active.c, { v: `=SUM(${addr(active.r, c + 1)}:${addr(active.r, active.c - 1)})` }); changed(); return; }
    startEdit('=SUM(');
  }
  document.getElementById('autoSum').addEventListener('click', autoSum);
  document.getElementById('autoSum2').addEventListener('click', autoSum);

  function dataRegion() {
    // Use the selection if it is a block, else the used area of the sheet.
    const g = selRange();
    if (g.r1 !== g.r2 && g.r2 - g.r1 < sheet().rows - 1) return opRange(g);
    const s = sheet();
    let maxR = 0, maxC = 0;
    Object.keys(s.cells).forEach(k => { const p = parseRef(k); if (s.cells[k].v !== '' && s.cells[k].v !== undefined) { maxR = Math.max(maxR, p.row); maxC = Math.max(maxC, p.col); } });
    return { r1: 0, c1: 0, r2: maxR, c2: maxC };
  }
  async function sortData(desc) {
    const g = dataRegion();
    if (!guard(g, 'sort')) return;
    const col = active.c;
    let hasHeader = false;
    if (g.r1 === 0 && g.r2 > 0) {
      const top = valueAt(0, col), next = valueAt(1, col);
      hasHeader = typeof top === 'string' && top !== '' && typeof next === 'number';
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
      const d = (typeof x === 'number' && typeof y === 'number') ? x - y : (typeof x === 'number' ? -1 : typeof y === 'number' ? 1 : String(x).localeCompare(String(y), undefined, { numeric: true, sensitivity: 'base' }));
      return desc ? -d : d;
    });
    rows.forEach((item, i) => item.row.forEach((cell, j) => {
      const r = start + i, c = g.c1 + j;
      if (!cell) { delete sheet().cells[addr(r, c)]; return; }
      if (typeof cell.v === 'string' && cell.v[0] === '=') cell.v = F.shiftFormula(cell.v, r - item.r, 0);
      sheet().cells[addr(r, c)] = cell;
    }));
    fitRows(rowsOf(g));
    changed();
    UI.toast(`Sorted by column ${indexToCol(col)}${hasHeader ? ' (header kept)' : ''}`);
  }
  ['sortAsc', 'sortAsc2'].forEach(id => document.getElementById(id).addEventListener('click', () => sortData(false)));
  ['sortDesc', 'sortDesc2'].forEach(id => document.getElementById(id).addEventListener('click', () => sortData(true)));

  document.getElementById('filter').addEventListener('click', async () => {
    const g = dataRegion();
    const col = active.c;
    const values = new Set();
    for (let r = g.r1 + 1; r <= g.r2 && values.size < 4; r++) { const v = valueAt(r, col); values.add(v === '' ? '(Blanks)' : String(v)); }
    const r = await UI.form(`Filter column ${indexToCol(col)}`, [
      { name: 'mode', label: 'Show rows where the value', type: 'select', options: ['equals', 'contains', 'is greater than', 'is less than', 'does not equal'] },
      { name: 'v', label: `Value (e.g. ${[...values].slice(0, 3).join(', ')})`, value: '' }
    ], 'Filter');
    if (!r) return;
    const hidden = [];
    const q = r.v.toLowerCase(), n = parseFloat(r.v);
    for (let row = g.r1 + 1; row <= g.r2; row++) {
      const v = valueAt(row, col);
      const s = (v === '' ? '(Blanks)' : String(v instanceof F.FormulaError ? v.error : v)).toLowerCase();
      const keep = { equals: s === q, contains: s.includes(q),
        'is greater than': typeof v === 'number' && v > n, 'is less than': typeof v === 'number' && v < n, 'does not equal': s !== q }[r.mode];
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
    if (!guard(g)) return;
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
    if (!guardSheet()) return;
    const r = await UI.form('Text to Columns', [{ name: 'd', label: 'Delimiter', type: 'select', options: [{ value: ',', label: 'Comma' }, { value: '\t', label: 'Tab' }, { value: ';', label: 'Semicolon' }, { value: ' ', label: 'Space' }, { value: '|', label: 'Pipe' }] }], 'Split');
    if (!r) return;
    const g = opRange();
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
  document.getElementById('showFormulas').addEventListener('click', e => { showFormulas = !showFormulas; e.currentTarget.classList.toggle('on', showFormulas); renderVisibleCells(); });
  document.getElementById('recalc').addEventListener('click', () => { changed(false); UI.toast('Recalculated'); });

  // ---------- View ----------
  document.getElementById('freezeRow').addEventListener('click', () => { sheet().freezeRow = !sheet().freezeRow; buildGrid(); state.markDirty(); });
  document.getElementById('freezeCol').addEventListener('click', () => { sheet().freezeCol = !sheet().freezeCol; buildGrid(); state.markDirty(); });
  document.getElementById('gridlines').addEventListener('click', e => { document.body.classList.toggle('no-grid'); e.currentTarget.classList.toggle('on', document.body.classList.contains('no-grid')); });
  const zoom = document.getElementById('zoom');
  function setZoom(z) {
    z = Math.max(50, Math.min(200, Math.round(z)));
    zoom.value = z;
    inner.style.zoom = z / 100;
    document.getElementById('zoomLabel').textContent = z + '%';
    renderWindow(true);
    if (editing) startEdit(editor.value);
  }
  zoom.addEventListener('input', () => setZoom(+zoom.value));
  document.getElementById('zoomIn').addEventListener('click', () => setZoom(+zoom.value + 10));
  document.getElementById('zoomOut').addEventListener('click', () => setZoom(+zoom.value - 10));
  document.getElementById('zoom100').addEventListener('click', () => setZoom(100));

  // ---------- Status bar aggregate ----------
  function updateAggregate() {
    const g = selRange();
    const el = document.getElementById('stAgg');
    if (g.r1 === g.r2 && g.c1 === g.c2) { el.textContent = ''; return; }
    let sum = 0, count = 0, numCount = 0;
    const add = v => { if (v === '') return; count++; if (typeof v === 'number') { sum += v; numCount++; } };
    const s = sheet();
    const keys = Object.keys(s.cells);
    // Walk whichever is smaller: the selection or the filled cells.
    if ((g.r2 - g.r1 + 1) * (g.c2 - g.c1 + 1) <= keys.length) forRange(g, (r, c) => { if (s.cells[addr(r, c)]) add(valueAt(r, c)); });
    else keys.forEach(k => { const p = parseRef(k); if (p.row >= g.r1 && p.row <= g.r2 && p.col >= g.c1 && p.col <= g.c2) add(valueAt(p.row, p.col)); });
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
    const found = [...after, ...keys].find(p => cellDisplay(p.row, p.col).text.toLowerCase().includes(q) || String(rawAt(p.row, p.col)).toLowerCase().includes(q));
    if (found) select(found.row, found.col); else UI.toast('No matches');
  }
  document.getElementById('findNext').addEventListener('click', findNext);
  findText.addEventListener('keydown', e => { if (e.key === 'Enter') findNext(); if (e.key === 'Escape') findbar.classList.remove('open'); });
  document.getElementById('replaceAll').addEventListener('click', () => {
    const q = findText.value;
    if (!q) return;
    snapshot();
    const re = new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
    const replacement = document.getElementById('replaceText').value;
    let n = 0;
    Object.values(sheet().cells).forEach(cell => {
      if (typeof cell.v !== 'string' || cell.v[0] === '=') return;
      cell.v = cell.v.replace(re, () => { n++; return replacement; });
    });
    changed();
    UI.toast(`Replaced ${n} occurrence${n === 1 ? '' : 's'}`);
  });

  // ---------- Charts ----------
  const PALETTE = ['#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47', '#264478', '#9e480e'];
  const CHART_TYPES = ['bar', 'line', 'pie', 'hbar', 'area', 'scatter', 'doughnut', 'combo'];
  function insertChart(type) {
    const g = opRange();
    if (!guardSheet('objects')) return;
    if (g.r1 === g.r2 && g.c1 === g.c2) return UI.toast('Select the data for the chart first (include labels)');
    if ((g.r2 - g.r1 + 1) * (g.c2 - g.c1 + 1) > 20000) return UI.toast('Select a smaller range for the chart');
    snapshot();
    const b = rangeBox(g);
    sheet().charts.push({ type, range: rangeText(g), x: b.x + b.w + 20, y: b.y, w: 440, h: 280, title: '' });
    renderCharts();
    state.markDirty();
  }
  document.querySelectorAll('[data-chart]').forEach(b => b.addEventListener('click', () => insertChart(b.dataset.chart)));
  const chartSheet = (ch, host = sheet()) => (ch.sheet && sheetByName(ch.sheet)) || host;
  /** The chart's labels and series from its range, with where each part is (for Excel references). */
  function chartData(ch, host) {
    const src = chartSheet(ch, host);
    const [a, b] = ch.range.split(':');
    const p = parseRef(a), q = parseRef(b || a);
    if (!p || !q || (q.row - p.row + 1) * (q.col - p.col + 1) > 20000) return null;
    const rows = [];
    for (let r = p.row; r <= q.row; r++) { const row = []; for (let c = p.col; c <= q.col; c++) row.push(engine.value(src.name, r, c)); rows.push(row); }
    const scatter = ch.type === 'scatter';
    const hasRowLabels = scatter ? rows[0].length > 1 : rows.slice(1).some(r => typeof r[0] === 'string' && r[0] !== '') || (rows.length > 1 && rows.every(r => typeof r[0] !== 'number'));
    const hasHeader = rows[0].slice(hasRowLabels ? 1 : 0).some(v => typeof v === 'string' && v !== '');
    const body = hasHeader ? rows.slice(1) : rows;
    const labels = body.map((r, i) => hasRowLabels ? String(r[0]) : String(i + 1));
    // Like Excel, a scatter chart whose x values aren't numbers plots them as 1, 2, 3…
    const numericX = hasRowLabels && body.every(r => typeof r[0] === 'number');
    const xs = body.map((r, i) => (numericX ? r[0] : i + 1));
    const seriesCount = rows[0].length - (hasRowLabels ? 1 : 0);
    const r1 = p.row + (hasHeader ? 1 : 0), r2 = q.row;
    const ref = (c1, ra, c2, rb) => `${quoteSheet(src.name)}!$${F.indexToCol(c1)}$${ra + 1}` + (ra === rb && c1 === c2 ? '' : `:$${F.indexToCol(c2)}$${rb + 1}`);
    const series = [];
    for (let k = 0; k < seriesCount; k++) {
      const ci = k + (hasRowLabels ? 1 : 0), col = p.col + ci;
      series.push({
        name: hasHeader ? String(rows[0][ci]) : 'Series ' + (k + 1),
        values: body.map(r => (typeof r[ci] === 'number' ? r[ci] : 0)),
        nameRef: hasHeader ? ref(col, p.row, col, p.row) : null,
        valRef: ref(col, r1, col, r2)
      });
    }
    return { labels, xs, series, catRef: hasRowLabels ? ref(p.col, r1, p.col, r2) : null, catNumeric: numericX };
  }
  const quoteSheet = n => (/^[A-Za-z_][\w.]*$/.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`);
  function drawChart(canvas, ch) {
    const dpr = window.devicePixelRatio || 1;
    const W = ch.w, H = ch.h - 22;
    canvas.width = W * dpr; canvas.height = H * dpr;
    canvas.style.width = W + 'px'; canvas.style.height = H + 'px';
    const ctx = canvas.getContext('2d');
    ctx.scale(dpr, dpr);
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, W, H);
    const data = chartData(ch);
    if (!data || !data.series.length || !data.labels.length) return;
    const title = ch.title || (data.series.length === 1 ? data.series[0].name : 'Chart');
    ctx.fillStyle = '#404040'; ctx.textAlign = 'center'; ctx.font = '600 14px Calibri, Arial';
    ctx.fillText(title, W / 2, 20);
    ctx.font = '11px Calibri, Arial';
    if (ch.type === 'pie' || ch.type === 'doughnut') {
      const vals = data.series[0].values.map(v => Math.max(0, v));
      const total = vals.reduce((a, b) => a + b, 0) || 1;
      const cx = W * 0.4, cy = H / 2 + 10, rad = Math.max(10, Math.min(W * 0.35, H / 2 - 30));
      let a0 = -Math.PI / 2;
      vals.forEach((v, i) => {
        const a1 = a0 + v / total * Math.PI * 2;
        ctx.beginPath(); ctx.moveTo(cx, cy); ctx.arc(cx, cy, rad, a0, a1); ctx.closePath();
        ctx.fillStyle = PALETTE[i % PALETTE.length]; ctx.fill();
        ctx.strokeStyle = '#fff'; ctx.lineWidth = 1.5; ctx.stroke();
        a0 = a1;
      });
      if (ch.type === 'doughnut') { ctx.beginPath(); ctx.arc(cx, cy, rad * 0.5, 0, Math.PI * 2); ctx.fillStyle = '#fff'; ctx.fill(); }
      ctx.textAlign = 'left';
      data.labels.forEach((l, i) => {
        const y = 44 + i * 18;
        if (y > H - 10) return;
        ctx.fillStyle = PALETTE[i % PALETTE.length]; ctx.fillRect(W * 0.78, y - 9, 10, 10);
        ctx.fillStyle = '#404040'; ctx.fillText(`${l} (${Math.round(vals[i] / total * 100)}%)`, W * 0.78 + 14, y);
      });
      return;
    }
    if (ch.type === 'scatter') return drawScatter(ctx, data, W, H);
    const all = data.series.flatMap(s => s.values);
    let max = Math.max(0, ...all), min = Math.min(0, ...all);
    if (max === min) max = min + 1;
    const step = niceStep((max - min) / 5);
    max = Math.ceil(max / step) * step; min = Math.floor(min / step) * step;
    const legendH = data.series.length > 1 ? 20 : 0;
    const L = 48, R = W - 14, T = 34, B = H - 26 - legendH;
    const horizontal = ch.type === 'hbar';
    ctx.strokeStyle = '#e1e1e1'; ctx.fillStyle = '#595959'; ctx.lineWidth = 1;
    for (let v = min; v <= max + step / 1e6; v += step) {
      if (horizontal) {
        const x = L + (v - min) / (max - min) * (R - L);
        ctx.beginPath(); ctx.moveTo(x, T); ctx.lineTo(x, B); ctx.stroke();
        ctx.textAlign = 'center'; ctx.fillText(F.formatGeneral(+v.toPrecision(10)), x, B + 14);
      } else {
        const y = B - (v - min) / (max - min) * (B - T);
        ctx.beginPath(); ctx.moveTo(L, y); ctx.lineTo(R, y); ctx.stroke();
        ctx.textAlign = 'right'; ctx.fillText(F.formatGeneral(+v.toPrecision(10)), L - 6, y + 4);
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
      const asLine = ch.type === 'line' || (ch.type === 'combo' && si === data.series.length - 1 && data.series.length > 1);
      if (ch.type === 'area') {
        const pts = s.values.map((v, i) => [L + band * i + band / 2, B - (v - min) / (max - min) * (B - T)]);
        ctx.globalAlpha = 0.45;
        ctx.beginPath(); ctx.moveTo(pts[0][0], zero); pts.forEach(([x, y]) => ctx.lineTo(x, y)); ctx.lineTo(pts[pts.length - 1][0], zero); ctx.closePath(); ctx.fill();
        ctx.globalAlpha = 1; ctx.lineWidth = 2; ctx.beginPath(); pts.forEach(([x, y], i) => (i ? ctx.lineTo(x, y) : ctx.moveTo(x, y))); ctx.stroke();
      } else if (asLine) {
        ctx.lineWidth = 2.25; ctx.beginPath();
        s.values.forEach((v, i) => { const x = L + band * i + band / 2, y = B - (v - min) / (max - min) * (B - T); i ? ctx.lineTo(x, y) : ctx.moveTo(x, y); });
        ctx.stroke();
        s.values.forEach((v, i) => { const x = L + band * i + band / 2, y = B - (v - min) / (max - min) * (B - T); ctx.beginPath(); ctx.arc(x, y, 3, 0, Math.PI * 2); ctx.fill(); });
      } else {
        const nBars = ch.type === 'combo' && data.series.length > 1 ? data.series.length - 1 : data.series.length;
        const bw = band * 0.7 / nBars;
        s.values.forEach((v, i) => {
          const off = band * 0.15 + bw * si;
          if (horizontal) {
            const x = L + (v - min) / (max - min) * (R - L);
            ctx.fillRect(Math.min(zero, x), T + band * i + off, Math.abs(x - zero), Math.max(1, bw - 1));
          } else {
            const y = B - (v - min) / (max - min) * (B - T);
            ctx.fillRect(L + band * i + off, Math.min(zero, y), Math.max(1, bw - 1), Math.abs(zero - y));
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
  /** Scatter: numeric x and y axes, a point per row. */
  function drawScatter(ctx, data, W, H) {
    const pts = data.series.map(s => s.values.map((y, i) => [data.xs[i], y]).filter(([x]) => typeof x === 'number'));
    const xsAll = pts.flat().map(p => p[0]), ysAll = pts.flat().map(p => p[1]);
    if (!xsAll.length) return;
    const scale = (lo, hi) => { if (lo === hi) { lo -= 1; hi += 1; } const st = niceStep((hi - lo) / 5); return [Math.floor(lo / st) * st, Math.ceil(hi / st) * st, st]; };
    const [x0, x1, xs] = scale(Math.min(...xsAll), Math.max(...xsAll)), [y0, y1, ys] = scale(Math.min(0, ...ysAll), Math.max(...ysAll));
    const legendH = data.series.length > 1 ? 20 : 0;
    const L = 48, R = W - 14, T = 34, B = H - 26 - legendH;
    const X = v => L + (v - x0) / (x1 - x0) * (R - L), Y = v => B - (v - y0) / (y1 - y0) * (B - T);
    ctx.strokeStyle = '#e1e1e1'; ctx.fillStyle = '#595959'; ctx.lineWidth = 1;
    for (let v = y0; v <= y1 + ys / 1e6; v += ys) { ctx.beginPath(); ctx.moveTo(L, Y(v)); ctx.lineTo(R, Y(v)); ctx.stroke(); ctx.textAlign = 'right'; ctx.fillText(F.formatGeneral(+v.toPrecision(10)), L - 6, Y(v) + 4); }
    for (let v = x0; v <= x1 + xs / 1e6; v += xs) { ctx.textAlign = 'center'; ctx.fillText(F.formatGeneral(+v.toPrecision(10)), X(v), B + 14); }
    pts.forEach((list, si) => { ctx.fillStyle = PALETTE[si % PALETTE.length]; list.forEach(([x, y]) => { ctx.beginPath(); ctx.arc(X(x), Y(y), 3.5, 0, Math.PI * 2); ctx.fill(); }); });
    if (legendH) { let x = L; ctx.textAlign = 'left'; data.series.forEach((s, si) => { ctx.fillStyle = PALETTE[si % PALETTE.length]; ctx.fillRect(x, H - 16, 10, 10); ctx.fillStyle = '#404040'; ctx.fillText(s.name, x + 14, H - 7); x += ctx.measureText(s.name).width + 32; }); }
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
        const sx = e.clientX, sy = e.clientY, ox = ch.x, oy = ch.y, z = zoomLevel();
        let moved = false;
        const mv = ev => {
          if (!moved) { snapshot(); moved = true; }
          ch.x = Math.max(0, Math.round(ox + (ev.clientX - sx) / z)); ch.y = Math.max(0, Math.round(oy + (ev.clientY - sy) / z));
          box.style.left = ch.x + 'px'; box.style.top = ch.y + 'px';
        };
        const up = () => { document.removeEventListener('mousemove', mv); document.removeEventListener('mouseup', up); if (moved) state.markDirty(); };
        document.addEventListener('mousemove', mv); document.addEventListener('mouseup', up);
      });
      box.querySelector('.grip').addEventListener('mousedown', e => {
        e.preventDefault();
        const sx = e.clientX, sy = e.clientY, ow = ch.w, oh = ch.h, z = zoomLevel();
        snapshot();
        const mv = ev => { ch.w = Math.max(200, Math.round(ow + (ev.clientX - sx) / z)); ch.h = Math.max(150, Math.round(oh + (ev.clientY - sy) / z)); box.style.width = ch.w + 'px'; box.style.height = ch.h + 'px'; drawChart(box.querySelector('canvas'), ch); };
        const up = () => { document.removeEventListener('mousemove', mv); document.removeEventListener('mouseup', up); state.markDirty(); };
        document.addEventListener('mousemove', mv); document.addEventListener('mouseup', up);
      });
      barEl.addEventListener('click', async e => {
        const a = e.target.dataset.a;
        if (a === 'del') { snapshot(); sheet().charts.splice(i, 1); renderCharts(); state.markDirty(); }
        if (a === 'type') { snapshot(); ch.type = CHART_TYPES[(CHART_TYPES.indexOf(ch.type) + 1) % CHART_TYPES.length]; drawChart(box.querySelector('canvas'), ch); state.markDirty(); }
        if (a === 'title') {
          const r = await UI.form('Chart', [{ name: 'title', label: 'Title', value: ch.title }, { name: 'range', label: 'Data range', value: ch.range }]);
          if (r) {
            snapshot();
            ch.title = r.title.slice(0, 200);
            if (/^[A-Z]{1,3}\d+(:[A-Z]{1,3}\d+)?$/i.test(r.range.trim())) ch.range = r.range.trim().toUpperCase();
            drawChart(box.querySelector('canvas'), ch);
            state.markDirty();
          }
        }
        if (a === 'png') {
          const blob = await new Promise(res => box.querySelector('canvas').toBlob(res));
          const name = (ch.title || 'Chart').replace(/[\\/:*?"<>|]/g, '_') + '.png';
          await Platform.save({ data: new Uint8Array(await blob.arrayBuffer()), name, filters: [{ name: 'PNG image', extensions: ['png'] }], saveAs: true });
        }
      });
    });
  }

  // ---------- File I/O ----------
  /** Styling and charts saved by Atlas inside an .xlsx are untrusted input: keep only known, well-typed values. */
  function cleanStyle(st) {
    if (!st || typeof st !== 'object') return null;
    const out = {};
    ['b', 'i', 'u', 'wrap', 'unlocked'].forEach(k => { if (st[k] === true) out[k] = true; });
    if (['left', 'center', 'right'].includes(st.align)) out.align = st.align;
    if (Platform.isColor(st.color)) out.color = st.color;
    if (Platform.isColor(st.bg)) out.bg = st.bg;
    if (typeof st.font === 'string' && /^[\w .-]{1,64}$/.test(st.font)) out.font = st.font;
    if (typeof st.size === 'number' && st.size >= 4 && st.size <= 96) out.size = st.size;
    if (st.border && typeof st.border === 'object') {
      const b = {};
      ['t', 'l', 'r', 'b'].forEach(k => { if (st.border[k]) b[k] = 1; });
      if (Object.keys(b).length) out.border = b;
    }
    return Object.keys(out).length ? out : null;
  }
  function cleanChart(ch) {
    if (!ch || typeof ch !== 'object' || !CHART_TYPES.includes(ch.type) || typeof ch.range !== 'string' || !/^[A-Z]{1,3}\d+(:[A-Z]{1,3}\d+)?$/.test(ch.range)) return null;
    const n = (v, d, lo, hi) => (typeof v === 'number' && isFinite(v) ? Math.max(lo, Math.min(hi, v)) : d);
    const out = { type: ch.type, range: ch.range, x: n(ch.x, 100, 0, 1e7), y: n(ch.y, 100, 0, 1e8), w: n(ch.w, 440, 200, 4000), h: n(ch.h, 280, 150, 4000), title: typeof ch.title === 'string' ? ch.title.slice(0, 200) : '' };
    if (typeof ch.sheet === 'string' && ch.sheet.length <= 31) out.sheet = ch.sheet;
    return out;
  }
  const cleanSizes = (obj, lo, hi) => {
    const out = {};
    if (obj && typeof obj === 'object') Object.entries(obj).forEach(([k, v]) => { if (/^\d{1,7}$/.test(k) && typeof v === 'number' && v >= lo && v <= hi) out[+k] = Math.round(v); });
    return out;
  };

  /**
   * Parses an Excel/ODS file in a throwaway worker with a time limit, so a hostile file can't
   * pollute this page's objects or freeze it (SheetJS advisories CVE-2023-30533, CVE-2024-22363).
   */
  function parseSpreadsheet(bytes) {
    return new Promise((resolve, reject) => {
      let worker;
      try { worker = new Worker('js/xlsx-worker.js'); } catch (err) { reject(err); return; }
      const timer = setTimeout(() => { worker.terminate(); reject(new Error('the file took too long to read')); }, 30000);
      worker.onmessage = e => {
        clearTimeout(timer);
        worker.terminate();
        if (e.data && e.data.ok) resolve(e.data); else reject(new Error((e.data && e.data.error) || 'unreadable file'));
      };
      worker.onerror = e => { clearTimeout(timer); worker.terminate(); reject(new Error(e.message || 'unreadable file')); };
      const copy = bytes.slice();
      worker.postMessage(copy, [copy.buffer]);
    });
  }

  async function workbookFromFile(bytes, name) {
    const ext = Platform.extOf(name);
    if (ext === 'csv' || ext === 'tsv') {
      const text = Platform.bytesToText(bytes).replace(/^﻿/, '');
      const s = newSheet(cleanSheetName(Platform.baseName(name)) || 'Sheet1');
      const first = text.slice(0, text.search(/\r?\n|$/));
      const delim = ext === 'tsv' ? '\t' : (first.includes(';') && !first.includes(',') ? ';' : ',');
      parseDelimited(text, delim).forEach((row, r) => row.forEach((v, c) => {
        if (v === '' || r >= MAX_ROWS || c >= MAX_COLS) return;
        s.cells[addr(r, c)] = { v };
      }));
      const u = usedBounds(s);
      s.rows = Math.max(100, u.r + 20); s.cols = Math.max(26, u.c + 5);
      return { sheets: [s], active: 0 };
    }
    const wb = await parseSpreadsheet(bytes);
    let meta = null;
    try { meta = typeof wb.AtlasMeta === 'string' ? JSON.parse(wb.AtlasMeta) : null; } catch { meta = null; }
    const usedNames = new Set();
    const sheets = wb.SheetNames.map((sn, si) => {
      const ws = wb.Sheets[si] || {};
      let name = cleanSheetName(sn) || 'Sheet' + (si + 1);
      while (usedNames.has(name.toLowerCase())) name = name.slice(0, 27) + '_' + si;
      usedNames.add(name.toLowerCase());
      const s = newSheet(name);
      // Walk the cells that exist rather than the declared range, which can be the whole sheet.
      let maxR = 0, maxC = 0;
      for (const k of Object.keys(ws)) {
        if (k[0] === '!') continue;
        const cell = ws[k];
        const p = parseRef(k);
        if (!p || !cell || p.row >= MAX_ROWS || p.col >= MAX_COLS) continue;
        let v;
        // Cells inside another cell's array formula hold its results, which spill again here.
        if (cell.F && String(cell.F).split(':')[0].toUpperCase() !== k.toUpperCase()) v = '';
        else if (cell.f) v = '=' + F.fromExcelFormula(cell.f);
        else if (cell.t === 'n') v = String(cell.v);
        else if (cell.t === 'b') v = cell.v ? 'TRUE' : 'FALSE';
        else if (cell.t === 'e') v = cell.w || '#N/A';
        else if (cell.t === 'd' || cell.isDate) v = String(F.dateToSerial(new Date(cell.v)));
        else if (cell.t === 'z') v = '';
        else {
          v = cell.v === undefined ? '' : String(cell.v);
          if (v && (v[0] === '=' || typeof F.parseLiteral(v) !== 'string')) v = "'" + v;
        }
        const st = {};
        if (cell.z && cell.z !== 'General') st.fmt = String(cell.z).slice(0, 200);
        const note = Array.isArray(cell.c) && cell.c.length ? { text: cell.c.map(x => String(x.t || '')).join('\n').slice(0, 32767), author: String(cell.c[0].a || '').slice(0, 100) } : null;
        if (v === '' && !Object.keys(st).length && !note) continue;
        s.cells[addr(p.row, p.col)] = { v, ...(Object.keys(st).length ? { s: st } : {}), ...(note ? { n: note } : {}) };
        if (p.row > maxR) maxR = p.row;
        if (p.col > maxC) maxC = p.col;
      }
      s.rows = Math.min(MAX_ROWS, Math.max(100, maxR + 20));
      s.cols = Math.min(MAX_COLS, Math.max(26, maxC + 5));
      (ws['!cols'] || []).slice(0, MAX_COLS).forEach((col, c) => { if (col && (col.wpx || col.wch)) s.colWidths[c] = Math.max(10, Math.min(2000, Math.round(col.wpx || col.wch * 7 + 5))); });
      (ws['!rows'] || []).slice(0, MAX_ROWS).forEach((row, r) => {
        if (!row) return;
        if (row.hidden) s.hiddenRows.push(r);
        const h = row.hpx || (row.hpt && row.hpt * 4 / 3);
        if (h && Math.abs(h - ROW_H) > 2) s.rowHeights[r] = Math.max(4, Math.min(545, Math.round(h)));
      });
      (ws['!merges'] || []).forEach(m => { if (m.e.r < MAX_ROWS && m.e.c < MAX_COLS) s.merges.push({ r1: m.s.r, c1: m.s.c, r2: m.e.r, c2: m.e.c }); });
      const ms = meta && Array.isArray(meta.sheets) ? meta.sheets[si] : null;
      if (ms && typeof ms === 'object' && ms.name === sn) {
        if (ms.styles && typeof ms.styles === 'object') {
          Object.entries(ms.styles).forEach(([k, st]) => {
            const clean = CELL_KEY.test(k) && cleanStyle(st);
            if (!clean) return;
            s.cells[k] = { v: s.cells[k]?.v ?? '', s: { ...(s.cells[k]?.s || {}), ...clean } };
          });
        }
        s.charts = (Array.isArray(ms.charts) ? ms.charts : []).map(cleanChart).filter(Boolean).slice(0, 100);
        if (ms.colWidths) s.colWidths = cleanSizes(ms.colWidths, 10, 2000);
        if (ms.freezeRow === true) s.freezeRow = true;
        if (ms.freezeCol === true) s.freezeCol = true;
      }
      if (ms && ms.name === sn && Array.isArray(ms.pivots)) s.pivots = ms.pivots.map(SheetsPivot.clean).filter(Boolean).slice(0, 50);
      // Atlas's own copy of the rules is exact; otherwise read Excel's.
      const ex = wb.Extras && Array.isArray(wb.Extras.sheets) ? wb.Extras.sheets[si] : null;
      if (ms && ms.name === sn && Array.isArray(ms.cf)) s.cf = ms.cf.map(SheetsCF.clean).filter(Boolean).slice(0, 500);
      else if (ex && typeof ex.cf === 'string' && ex.cf) {
        try { s.cf = SheetsCF.fromXlsx(ex.cf, typeof wb.Extras.dxfs === 'string' ? wb.Extras.dxfs : ''); } catch (err) { console.info('Conditional formatting skipped:', err); }
      }
      if (!(ms && ms.name === sn) && ex && ex.drawing) {
        try { s.charts = chartsFromExcel(ex, s); } catch (err) { console.info('Charts skipped:', err); }
      }
      // Protection: Atlas's copy, else Excel's (with its unlocked cells).
      const prot = ms && ms.name === sn && ms.protect ? SheetsProtect.clean(ms.protect) : ex && ex.protection ? SheetsProtect.fromXml(ex.protection) : null;
      if (prot) s.protect = prot;
      if (!(ms && ms.name === sn) && ex && Array.isArray(ex.unlocked)) ex.unlocked.slice(0, 100000).forEach(k => { if (CELL_KEY.test(k)) s.cells[k] = { v: s.cells[k]?.v ?? '', ...(s.cells[k] || {}), s: { ...((s.cells[k] || {}).s || {}), unlocked: true } }; });
      if (ms && ms.name === sn && Array.isArray(ms.dv)) s.dv = ms.dv.map(SheetsDV.clean).filter(Boolean).slice(0, 500);
      else if (ex && typeof ex.dv === 'string' && ex.dv) {
        try { s.dv = SheetsDV.fromXlsx(ex.dv); } catch (err) { console.info('Data validation skipped:', err); }
      }
      return s;
    });
    if (!sheets.length) sheets.push(newSheet('Sheet1'));
    const act = meta && Number.isInteger(meta.active) ? meta.active : 0;
    return { sheets, active: Math.max(0, Math.min(act, sheets.length - 1)) };
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
        const note = cell.n ? [{ a: cell.n.author || 'Author', t: cell.n.text }] : null;
        if (note) note.hidden = true;
        if (cell.v === '' || cell.v === undefined) {
          if (st.fmt || note) { ws[k] = { t: 'z', ...(st.fmt ? { z: st.fmt } : {}), ...(note ? { c: note } : {}) }; maxR = Math.max(maxR, p.row); maxC = Math.max(maxC, p.col); }
          return;
        }
        maxR = Math.max(maxR, p.row); maxC = Math.max(maxC, p.col);
        const v = engine.value(s.name, p.row, p.col);
        const out = {};
        if (typeof cell.v === 'string' && cell.v[0] === '=') {
          out.f = F.toExcelFormula(cell.v.slice(1));
          const sp = engine.spillAt(s.name, p.row, p.col);
          // A spilling formula is a dynamic array formula over its range.
          if (sp && sp.row === p.row && sp.col === p.col && (sp.h > 1 || sp.w > 1)) { out.F = `${k}:${addr(p.row + sp.h - 1, p.col + sp.w - 1)}`; out.D = true; }
        }
        if (v instanceof F.FormulaError) { out.t = 'e'; out.v = { '#NULL!': 0, '#DIV/0!': 7, '#VALUE!': 15, '#REF!': 23, '#NAME?': 29, '#NUM!': 36, '#N/A': 42 }[v.error] ?? 15; out.w = v.error; }
        else if (typeof v === 'number') { out.t = 'n'; out.v = v; }
        else if (typeof v === 'boolean') { out.t = 'b'; out.v = v; }
        else { out.t = 's'; out.v = String(v); }
        if (st.fmt) out.z = st.fmt;
        if (note) out.c = note;
        ws[k] = out;
      });
      // Values a formula spilled into are stored too, as Excel does.
      for (const sp of engine.spillsOn(s.name)) {
        for (let r = sp.row; r < sp.row + sp.h; r++) for (let c = sp.col; c < sp.col + sp.w; c++) {
          if (r === sp.row && c === sp.col) continue;
          const v = engine.value(s.name, r, c), k = addr(r, c);
          if (v === '' || ws[k]) continue;
          ws[k] = v instanceof F.FormulaError ? { t: 'e', v: 15, w: v.error } : typeof v === 'number' ? { t: 'n', v } : typeof v === 'boolean' ? { t: 'b', v } : { t: 's', v: String(v) };
          maxR = Math.max(maxR, r); maxC = Math.max(maxC, c);
        }
      }
      ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxR, c: maxC } });
      const cols = [];
      Object.entries(s.colWidths).forEach(([c, w]) => { cols[+c] = { wpx: w }; });
      if (cols.length) { for (let i = 0; i < cols.length; i++) if (!cols[i]) cols[i] = { wpx: DEFAULT_COL_W }; ws['!cols'] = cols; }
      const rows = [];
      Object.entries(s.rowHeights || {}).forEach(([r, h]) => { rows[+r] = { hpx: h }; });
      (s.hiddenRows || []).forEach(r => { rows[r] = { ...(rows[r] || {}), hidden: true }; });
      if (rows.length) ws['!rows'] = rows;
      if (s.merges.length) ws['!merges'] = s.merges.map(m => ({ s: { r: m.r1, c: m.c1 }, e: { r: m.r2, c: m.c2 } }));
      XLSX.utils.book_append_sheet(wb, ws, s.name);
      meta.sheets.push({ name: s.name, protect: s.protect || null, styles, charts: s.charts, colWidths: s.colWidths, freezeRow: !!s.freezeRow, freezeCol: !!s.freezeCol, cf: s.cf || [], pivots: s.pivots || [], dv: s.dv || [] });
    });
    // Cell styling and charts live in a custom document property so Atlas can restore them;
    // other apps ignore it and still get values, formulas, number formats, widths and merges.
    wb.Custprops = { AtlasMeta: JSON.stringify(meta) };
    return wb;
  }
  // Chart positions are grid pixels (including the headers); Excel anchors them to cells plus EMU offsets.
  const EMU = 9525;
  const colLeft = (s, c) => { let x = 0; for (let i = 0; i < c; i++) x += s.colWidths[i] || DEFAULT_COL_W; return x; };
  const rowTopOf = (s, r) => { let y = r * ROW_H; for (const [k, h] of Object.entries(s.rowHeights || {})) if (+k < r) y += h - ROW_H; return y; };
  function cellOfPx(s, px, axis) {
    let i = 0, start = 0;
    for (;;) {
      const size = axis === 'col' ? s.colWidths[i] || DEFAULT_COL_W : (s.rowHeights || {})[i] || ROW_H;
      if (start + size > px || i > (axis === 'col' ? 16383 : 1048575)) return { i, off: Math.round((px - start) * EMU) };
      start += size; i++;
    }
  }
  /** Excel drawing and chart parts for a sheet's charts. */
  function chartParts(s) {
    if (!s.charts || !s.charts.length) return null;
    const anchors = [], parts = [];
    s.charts.forEach((ch, i) => {
      const data = chartData(ch, s);
      if (!data || !data.series.length) return;
      const series = (ch.type === 'pie' || ch.type === 'doughnut' ? data.series.slice(0, 1) : data.series).map((x, k) => ({
        ...x, color: PALETTE[k % PALETTE.length], catRef: data.catRef, cats: ch.type === 'scatter' ? data.xs : data.labels, catNumeric: ch.type === 'scatter' || data.catNumeric
      }));
      parts.push(SheetsCharts.toXml({ type: ch.type, title: ch.title || (series.length === 1 && ch.type !== 'scatter' ? series[0].name : ''), series, palette: PALETTE }));
      const at = (x, y) => { const c = cellOfPx(s, Math.max(0, x - RH_W), 'col'), r = cellOfPx(s, Math.max(0, y - HEAD_H), 'row'); return { col: c.i, colOff: c.off, row: r.i, rowOff: r.off }; };
      anchors.push({ name: ch.title || 'Chart ' + (i + 1), from: at(ch.x, ch.y), to: at(ch.x + ch.w, ch.y + ch.h) });
    });
    return parts.length ? { drawing: SheetsCharts.drawingXml(anchors), parts } : null;
  }
  /** Charts read from an Excel sheet's drawing, as Atlas charts over the same data. */
  function chartsFromExcel(ex, s) {
    if (!ex || !ex.drawing || !Array.isArray(ex.charts)) return [];
    const out = [];
    SheetsCharts.anchorsFromDrawing(ex.drawing).slice(0, 50).forEach(a => {
      const part = ex.charts.find(c => c && c.rid === a.rid);
      const spec = part && typeof part.xml === 'string' && SheetsCharts.fromXml(part.xml);
      if (!spec || !spec.series.length) return;
      // The data range is the box around every reference (names, categories, values) on one sheet.
      let sheetName = null, box = null, ok = true;
      for (const sr of spec.series) for (const ref of [sr.nameRef, sr.catRef, sr.valRef]) {
        if (!ref) continue;
        const m = /^(?:(?:'((?:[^']|'')+)'|([^!]+))!)?\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$/i.exec(ref.replace(/^\(|\)$/g, ''));
        if (!m) { ok = false; continue; }
        const sn = m[1] ? m[1].replace(/''/g, "'") : m[2] || s.name;
        if (sheetName && sheetName.toLowerCase() !== sn.toLowerCase()) { ok = false; continue; }
        sheetName = sn;
        const g = SheetsCF.area((m[3] + m[4] + (m[5] ? ':' + m[5] + m[6] : '')).toUpperCase());
        box = box ? { r1: Math.min(box.r1, g.r1), c1: Math.min(box.c1, g.c1), r2: Math.max(box.r2, g.r2), c2: Math.max(box.c2, g.c2) } : g;
      }
      if (!ok || !box) return;
      const x = RH_W + colLeft(s, a.from.col) + a.from.colOff / EMU, y = HEAD_H + rowTopOf(s, a.from.row) + a.from.rowOff / EMU;
      const w = a.to ? RH_W + colLeft(s, a.to.col) + a.to.colOff / EMU - x : a.ext ? a.ext.cx / EMU : 440;
      const h = a.to ? HEAD_H + rowTopOf(s, a.to.row) + a.to.rowOff / EMU - y : a.ext ? a.ext.cy / EMU : 280;
      const ch = cleanChart({ type: spec.type, range: SheetsCF.refOf(box), x: Math.round(x), y: Math.round(y), w: Math.round(w), h: Math.round(h), title: spec.title, ...(sheetName.toLowerCase() !== s.name.toLowerCase() ? { sheet: sheetName } : {}) });
      if (ch) out.push(ch);
    });
    return out;
  }
  async function serialize(ext) {
    if (ext === 'csv') {
      const u = usedBounds();
      const lines = [];
      for (let r = 0; r <= u.r; r++) {
        const row = [];
        for (let c = 0; c <= u.c; c++) {
          const v = valueAt(r, c);
          const st = (cellAt(r, c) || {}).s || {};
          let t = v instanceof F.FormulaError ? v.error : typeof v === 'number' ? (st.fmt && /[ymd]/i.test(st.fmt) ? F.formatValue(v, st.fmt) : String(v)) : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v);
          if (/[",\r\n]/.test(t)) t = '"' + t.replace(/"/g, '""') + '"';
          row.push(t);
        }
        lines.push(row.join(','));
      }
      return lines.join('\r\n') + '\r\n';
    }
    const bookType = { xlsx: 'xlsx', ods: 'ods', xls: 'biff8' }[ext] || 'xlsx';
    const bytes = new Uint8Array(XLSX.write(toWorkbook(), { type: 'array', bookType, compression: true }));
    if (bookType !== 'xlsx') return bytes;
    // What SheetJS can't write goes into the XML afterwards.
    const dxfs = [];
    const sheets = book.sheets.map(s => ({ conditionalFormatting: SheetsCF.toXlsx(s.cf, dxfs), dataValidations: SheetsDV.toXlsx(s.dv), charts: chartParts(s),
      sheetProtection: SheetsProtect.toXml(s.protect),
      unlocked: Object.keys(s.cells).filter(k => s.cells[k].s && s.cells[k].s.unlocked) }));
    return XlsxPatch.patch(bytes, { sheets, dxfs });
  }

  async function loadFile(f) {
    let next;
    try {
      next = await workbookFromFile(f.data, f.name);
    } catch (err) {
      UI.toast('Could not open file: ' + err.message, 4000);
      return;
    }
    book = next;
    undoStack.length = redoStack.length = undoMarks.length = redoMarks.length = 0;
    undoBytes = 0;
    engine.astCache.clear();
    active = anchor = { r: 0, c: 0 };
    wrap.scrollTop = wrap.scrollLeft = 0;
    state.set({ name: f.name, path: f.path, dirty: false });
    Platform.rememberRecent(f.name, 'sheets');
    // Size rows for wrapped text and big fonts on every sheet.
    const keep = book.active;
    book.sheets.forEach((s, i) => { book.active = i; rebuildCols(); fitRowsFor(s); });
    book.active = keep;
    buildGrid();
  }
  /** Like fitRows(null) but keeps heights that came from the file. */
  function fitRowsFor(s) {
    for (const [k, cell] of Object.entries(s.cells)) {
      const st = cell.s;
      if (!st || !(st.wrap || st.size > 11)) continue;
      const p = parseRef(k);
      if (s.merges.some(m => p.row >= m.r1 && p.row <= m.r2 && p.col >= m.c1 && p.col <= m.c2)) continue;
      const h = neededHeight(p.row, p.col, cell);
      if (h > (s.rowHeights[p.row] || ROW_H)) s.rowHeights[p.row] = h;
    }
  }
  async function save(saveAs, forcedExt, quiet) {
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
        await Platform.save({ data: await serialize(text || 'xlsx'), name: target.name, path: target.path });
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
  function printView() {
    const s = sheet();
    const u = usedBounds();
    const t = document.createElement('table');
    engine.invalidate();
    cf.begin();
    const cap = Math.min(u.r, 20000); // printing is for reports, not whole databases
    for (let r = 0; r <= cap; r++) {
      if (hiddenSet.has(r)) continue;
      const tr = t.insertRow();
      for (let c = 0; c <= u.c; c++) {
        const td = tr.insertCell();
        const cell = s.cells[addr(r, c)];
        if (!cell) continue;
        const st = cell.s || {};
        const { text, cls } = cellDisplay(r, c);
        const fmt = cf.format(r, c, valueAt(r, c));
        td.textContent = (fmt && fmt.icon ? fmt.icon[0] + ' ' : '') + text;
        Object.assign(td.style, {
          fontWeight: st.b ? '700' : '', fontStyle: st.i ? 'italic' : '', textDecoration: st.u ? 'underline' : '',
          color: st.color || '', backgroundColor: st.bg || '', fontFamily: st.font || '', fontSize: st.size ? st.size + 'pt' : '',
          textAlign: st.align || (cls === 'num' ? 'right' : cls ? 'center' : ''), whiteSpace: st.wrap ? 'pre-wrap' : 'nowrap'
        });
        if (fmt) applyCf(td.style, fmt);
      }
    }
    const area = document.getElementById('printArea');
    area.innerHTML = '';
    area.appendChild(t);
  }
  window.addEventListener('beforeprint', printView);

  const CLOUD_EXTS = ['xlsx', 'ods', 'csv'];
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
      case 'open-cloud': Cloud.openDialog({ app: 'sheets', load: loadFile, isEmpty: async () => Platform.shouldOpenInPlace(!state.dirty && book.sheets.length === 1 && !Object.keys(sheet().cells).length) && (!state.dirty || UI.confirm('Open file', 'Discard unsaved changes in this workbook?')) }); break;
      case 'save-cloud': Cloud.saveAsDialog({ state, serialize: ext => serialize(ext), exts: CLOUD_EXTS, before: () => commitEdit() }); break;
      case 'share': Cloud.shareDialog({ state, serialize: ext => serialize(ext), exts: CLOUD_EXTS, before: () => commitEdit(), save: () => save(false) }); break;
      case 'save-and-close': if (await save(false)) Platform.closeNow(); break;
      case 'export-csv': save(true, 'csv'); break;
      case 'export-pdf': printView(); Platform.exportPdf(Platform.baseName(state.name) + '.pdf', true); break;
      case 'print': printView(); Platform.print(); break;
      case 'undo': if (!editing) undo(); else document.execCommand('undo'); break;
      case 'redo': if (!editing) redo(); else document.execCommand('redo'); break;
      case 'find': case 'replace': findbar.classList.add('open'); findText.focus(); findText.select(); break;
      case 'copy-btn': wrap.focus(); document.execCommand('copy') || UI.toast('Use Ctrl+C'); break;
      case 'cut-btn': wrap.focus(); document.execCommand('cut') || UI.toast('Use Ctrl+X'); break;
      case 'paste-btn': try { const text = await navigator.clipboard.readText(); wrap.focus(); paste(text); } catch { UI.toast('Use Ctrl+V to paste'); } break;
      case 'home': Platform.desktop ? Platform.newWindow('home') : (location.href = 'index.html'); break;
    }
  });

  window.addEventListener('resize', () => { renderWindow(false); paintSelection(); });
  (async () => {
    buildGrid();
    wrap.focus();
    const f = await Platform.pending();
    if (f && f.error) UI.toast('Could not open file: ' + f.error, 4000);
    else if (f && f.recovered) loadRecovered(f.recovered);
    else if (f) await loadFile(f);
  })();

  // ---------- Crash recovery & AutoSave ----------
  const recovery = Platform.enableRecovery({
    app: 'sheets', state,
    snapshot: () => JSON.stringify(book),
    autosave: () => save(false, null, true),
    canAutoSave: () => !Platform.desktop && !Platform.isCloudPath(state.path) ? 'AutoSave needs the desktop app' : !state.path ? 'Save the workbook first to turn on AutoSave'
      : /\.(xlsx|ods)$/i.test(state.path) ? true : 'AutoSave works with .xlsx and .ods files'
  });
  function loadRecovered(d) {
    let next;
    try { next = JSON.parse(d.data); } catch { return UI.toast('The recovered workbook could not be read', 4000); }
    if (!next || !Array.isArray(next.sheets) || !next.sheets.length) return;
    book = next;
    book.sheets.forEach(s => {
      s.cf = (Array.isArray(s.cf) ? s.cf : []).map(SheetsCF.clean).filter(Boolean);
      s.pivots = (Array.isArray(s.pivots) ? s.pivots : []).map(SheetsPivot.clean).filter(Boolean);
      s.dv = (Array.isArray(s.dv) ? s.dv : []).map(SheetsDV.clean).filter(Boolean);
    });
    book.active = Math.max(0, Math.min(book.active | 0, book.sheets.length - 1));
    undoStack.length = redoStack.length = undoMarks.length = redoMarks.length = 0;
    undoBytes = 0;
    engine.astCache.clear();
    active = anchor = { r: 0, c: 0 };
    buildGrid();
    state.set({ name: d.name, path: d.path, dirty: true });
    recovery.adopt(d.draftId);
    UI.toast('Recovered your unsaved changes', 4000);
  }

  // ---------- Live editing ----------
  // Keys: sh:<sheet id> -> the sheet's settings and position, c:<sheet id>:<A1> -> a cell.
  const sheetUid = s => s.uid || (s.uid = 's' + Math.random().toString(36).slice(2, 12));
  const SHEET_LOCAL = new Set(['cells', 'rows', 'cols', 'uid']);
  const orphanCells = new Map(); // cells that arrived before their sheet
  function drawPeers() {
    if (!peerLayer) { if (!peers.length) return; peerLayer = document.createElement('div'); inner.appendChild(peerLayer); }
    peerLayer.innerHTML = '';
    const u = sheet() && sheet().uid;
    for (const p of peers) {
      const g = p.presence;
      if (!g || g.sheet !== u || !(g.r2 < sheet().rows && g.c2 < sheet().cols)) continue;
      const b = rangeBox(g);
      const f = document.createElement('div');
      f.className = 'live-flag';
      Object.assign(f.style, { left: (b.x - 1) + 'px', top: (b.y - 1) + 'px', width: (b.w - 2) + 'px', height: (b.h - 2) + 'px', border: `2px solid ${p.color}` });
      const tag = document.createElement('span');
      tag.className = 'tag';
      tag.style.background = p.color;
      tag.textContent = p.name;
      f.appendChild(tag);
      peerLayer.appendChild(f);
    }
  }
  function growFor(s, ref) {
    const p = parseRef(ref);
    if (!p) return;
    if (p.row >= s.rows - 1) s.rows = Math.min(MAX_ROWS, p.row + 20);
    if (p.col >= s.cols - 1) s.cols = Math.min(MAX_COLS, p.col + 5);
  }
  live = Collab.attach({
    app: 'sheets', doc: state,
    dirty: () => state.dirty || Object.keys(sheet().cells).length > 0 || book.sheets.length > 1,
    busy: () => editing,
    state() {
      const out = {};
      book.sheets.forEach((s, i) => {
        const u = sheetUid(s);
        const props = { pos: i };
        for (const [k, v] of Object.entries(s)) if (!SHEET_LOCAL.has(k)) props[k] = v;
        out['sh:' + u] = props;
        for (const [ref, cell] of Object.entries(s.cells)) out[`c:${u}:${ref}`] = cell;
      });
      return out;
    },
    apply(changes) {
      const activeUid = sheetUid(sheet());
      const byUid = new Map(book.sheets.map((s, i) => { s.pos = i; return [sheetUid(s), s]; }));
      let structural = false;
      for (const [k, v] of Object.entries(changes)) {
        if (k.startsWith('sh:')) {
          structural = true;
          const u = k.slice(3);
          if (v === null) { byUid.delete(u); continue; }
          let s = byUid.get(u);
          if (!s) {
            s = { ...newSheet(''), uid: u, cells: orphanCells.get(u) || {} };
            orphanCells.delete(u);
            byUid.set(u, s);
            Object.keys(s.cells).forEach(ref => growFor(s, ref));
          }
          for (const key of Object.keys(s)) if (!SHEET_LOCAL.has(key)) delete s[key];
          Object.assign(s, JSON.parse(JSON.stringify(v)));
        } else if (k.startsWith('c:')) {
          const i = k.indexOf(':', 2);
          const u = k.slice(2, i), ref = k.slice(i + 1);
          if (!CELL_KEY.test(ref)) continue;
          const s = byUid.get(u);
          let cells = s && s.cells;
          if (!cells) { if (!orphanCells.has(u)) orphanCells.set(u, {}); cells = orphanCells.get(u); }
          if (v === null) delete cells[ref]; else { cells[ref] = v; if (s) growFor(s, ref); }
        }
      }
      if (structural) {
        const list = [...byUid.values()].sort((a, b) => (a.pos - b.pos) || (a.uid < b.uid ? -1 : 1));
        if (list.length) book.sheets = list;
      }
      book.sheets.forEach(s => { delete s.pos; });
      const ai = book.sheets.findIndex(s => s.uid === activeUid);
      book.active = ai >= 0 ? ai : Math.min(book.active, book.sheets.length - 1);
      engine.astCache.clear();
      if (structural) buildGrid(); else changed();
      if (structural) state.markDirty();
    },
    presence: () => ({ sheet: sheetUid(sheet()), ...selRange() }),
    showPeers(list) { peers = list; drawPeers(); }
  });

  // ---------- Proofing ----------
  let proofLang = '';
  const proofing = Proofing.attach({
    app: 'sheets',
    roots: () => {
      commitEdit();
      const out = [];
      book.sheets.forEach((s, si) => {
        for (const [ref, cell] of Object.entries(s.cells)) {
          if (typeof cell.v !== 'string' || !cell.v.trim() || cell.v[0] === '=' || !/\p{L}{2}/u.test(cell.v)) continue;
          const el = document.createElement('div');
          el.textContent = cell.v;
          out.push({
            el,
            commit: () => { snapshot(); setCell(parseRef(ref).row, parseRef(ref).col, { v: el.textContent }, s); changed(); },
            focus: () => { if (book.active !== si) { book.active = si; buildGrid(); } const p = parseRef(ref); select(p.row, p.col); }
          });
        }
      });
      return out;
    },
    lang: () => proofLang,
    setLang: code => { proofLang = code; }
  });
  document.getElementById('proofEditor').addEventListener('click', () => proofing.open());
  document.getElementById('proofLang').addEventListener('click', () => proofing.languageDialog());

  // ---------- Assistant ----------
  function rangeTsv(g, limit) {
    const lines = [];
    let n = 0;
    for (let r = g.r1; r <= g.r2 && n < limit; r++) {
      const row = [];
      for (let c = g.c1; c <= g.c2; c++, n++) {
        const raw = rawAt(r, c);
        const shown = raw === '' ? '' : String(raw)[0] === '=' ? `${raw} → ${valueAt(r, c)}` : raw;
        row.push(String(shown).replace(/[\t\n]/g, ' '));
      }
      lines.push(`${r + 1}\t${row.join('\t')}`);
    }
    const head = [];
    for (let c = g.c1; c <= g.c2; c++) head.push(indexToCol(c));
    return `\t${head.join('\t')}\n${lines.join('\n')}`;
  }
  function aiContext() {
    commitEdit();
    const u = usedBounds();
    const all = { r1: 0, c1: 0, r2: u.r, c2: u.c };
    const g = selRange();
    const multi = g.r1 !== g.r2 || g.c1 !== g.c2;
    return {
      document: `Sheet "${sheet().name}" (other sheets: ${book.sheets.filter(x => x !== sheet()).map(x => x.name).join(', ') || 'none'}). ` +
        `Cells are tab-separated with column letters across the top and row numbers down the side; formulas are shown as "=FORMULA → value".\n` + rangeTsv(all, 20000),
      selection: `${rangeText(g)}${multi ? '\n' + rangeTsv(opRange(g), 5000) : ''}`,
      extra: `The active cell is ${addr(active.r, active.c)}.`,
      hasSelection: true
    };
  }
  /** Writes a reply into the sheet from the active cell: one value or formula, or a table. */
  function aiPut(text) {
    let lines = text.trim().replace(/^```\w*\n?|\n?```$/g, '').split('\n').map(l => l.trimEnd()).filter(l => l.trim());
    // A formula with an explanation after it: just the formula.
    if (lines.length && lines[0][0] === '=') lines = [lines[0]];
    let rows;
    if (lines.length && lines.every(l => /^\s*\|/.test(l))) {
      rows = lines.filter(l => !/^\s*\|?\s*:?-{3,}/.test(l)).map(l => l.replace(/^\s*\||\|\s*$/g, '').split('|').map(x => x.trim()));
    } else if (lines.some(l => l.includes('\t'))) rows = lines.map(l => l.split('\t'));
    else if (lines.length > 1 && lines.every(l => l.includes(',') && l[0] !== '=')) rows = lines.map(l => l.split(',').map(x => x.trim()));
    else rows = lines.length === 1 ? [[lines[0]]] : lines.map(l => [l]);
    if (!rows.length) return;
    if (rows.length * Math.max(...rows.map(r => r.length)) > 100000) return UI.toast('That is too much to paste at once', 4000);
    snapshot();
    rows.forEach((row, i) => row.forEach((v, j) => {
      const r = active.r + i, c = active.c + j;
      setCell(r, c, { v: v.replace(/\*\*/g, '') });
      autoFormat(r, c, v);
      ensureSize(r, c);
    }));
    anchor = { r: active.r + rows.length - 1, c: active.c + Math.max(...rows.map(r => r.length)) - 1 };
    changed();
    drawSelection();
    UI.toast(rows.length === 1 && rows[0].length === 1 ? 'Put in ' + addr(active.r, active.c) : 'Pasted into ' + rangeText(selRange()));
  }
  Assistant.attach({
    app: 'sheets',
    context: aiContext,
    format: 'In Atlas Sheets, give a formula as one line starting with = using Excel syntax and A1 references (Atlas supports Excel\'s functions, including dynamic arrays such as FILTER, SORT, UNIQUE and XLOOKUP). Give a table as a Markdown table or tab-separated lines; it is pasted starting at the active cell.',
    actions: [
      { label: 'Write a formula', prompt: 'Write a formula for the active cell that does what the selection and nearby headers suggest. If it is unclear, give your best formula and explain it in one short sentence after a blank line.' },
      { label: 'Explain this formula', prompt: 'Explain what the formula in the active cell does, step by step, in plain language.' },
      { label: 'Analyze the selection', prompt: 'Analyze the selected data: give the main patterns, totals or averages that matter, outliers, and anything that looks wrong. Be brief.' },
      { label: 'Find problems', prompt: 'Check this sheet for likely mistakes: broken or inconsistent formulas, hard-coded numbers that should be formulas, mismatched ranges, typos in labels. List each with its cell.' },
      { label: 'Suggest a chart', prompt: 'Which chart would show the selected data best, and why? Answer in two sentences.' }
    ],
    buttons: [{ label: 'Put in sheet', run: text => aiPut(text) }]
  });

  window.AtlasSheets = {
    loadFile, serialize, select, refreshPivots, get live() { return live; },
    get book() { return book; },
    setCell: (ref, v) => { const p = parseRef(ref); setCell(p.row, p.col, { v }); ensureSize(p.row, p.col); changed(); },
    valueOf: ref => { const p = parseRef(ref); return valueAt(p.row, p.col); },
    domStats: () => ({ rows: tbody.rows.length, cells: grid.querySelectorAll('td').length })
  };
})();
