'use strict';
// Isadek Sheets — tables, named ranges, What-If Analysis (Goal Seek, Solver), sparklines,
// slicers, scripts (Automate) and data from the web: the dialogs and on-sheet parts.
// sheets.js passes the workbook and the helpers these need (ctx), and calls back in to draw.
(function () {
  const T = () => window.SheetsTables;
  const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  function SheetsData(ctx) {
    const F = window.Formula;
    const { addr, parseRef, indexToCol } = F;
    const book = () => ctx.book();
    const sheet = () => ctx.sheet();
    const textAt = (s, r, c) => String(ctx.rawAt(r, c, s) ?? '').replace(/^'/, '');
    const shown = (s, r, c) => {
      const v = ctx.engine.value(s.name, r, c);
      if (v instanceof F.FormulaError) return v.error;
      const st = ((s.cells[addr(r, c)] || {}).s) || {};
      return typeof v === 'number' ? F.formatValue(v, st.fmt || 'General') : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v);
    };
    const parseRange = text => {
      const t = String(text || '').trim().replace(/\$/g, '');
      let sheetName = null, body = t;
      const bang = t.lastIndexOf('!');
      if (bang > 0) { sheetName = t.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'"); body = t.slice(bang + 1); }
      const [a, b] = body.toUpperCase().split(':');
      const p = parseRef(a || ''), q = b ? parseRef(b) : p;
      if (!p || !q) return null;
      return { sheet: sheetName, r1: Math.min(p.row, q.row), c1: Math.min(p.col, q.col), r2: Math.max(p.row, q.row), c2: Math.max(p.col, q.col) };
    };
    const rangeText = g => addr(g.r1, g.c1) + (g.r1 === g.r2 && g.c1 === g.c2 ? '' : ':' + addr(g.r2, g.c2));
    const quoteSheet = n => (/^[A-Za-z_][\w.]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`);

    // =====================================================================================
    // Tables
    // =====================================================================================
    const allTables = () => book().sheets.flatMap(s => s.tables || []);
    const tableAt = (r, c, s = sheet()) => T().tableAt(s.tables, r, c);
    const usedName = n => allTables().some(t => t.name.toLowerCase() === n.toLowerCase()) || (book().names || []).some(x => x.name.toLowerCase() === n.toLowerCase());
    function newTableName() { let i = 1; while (usedName('Table' + i)) i++; return 'Table' + i; }
    const headerNames = (t, s = sheet()) => T().headerNames(t, (r, c) => textAt(s, r, c));
    const overlaps = (a, b) => !(a.r2 < b.r1 || b.r2 < a.r1 || a.c2 < b.c1 || b.c2 < a.c1);

    /** The block of filled cells around the active cell (what Ctrl+A selects in Excel). */
    function currentRegion() {
      const g = ctx.selRange();
      if (g.r1 !== g.r2 || g.c1 !== g.c2) return g;
      const s = sheet(), filled = (r, c) => r >= 0 && c >= 0 && ctx.rawAt(r, c) !== '';
      let { r1, c1, r2, c2 } = g;
      if (!filled(r1, c1)) return g;
      for (let grew = true; grew;) {
        grew = false;
        const rowHas = r => { for (let c = c1 - 1; c <= c2 + 1; c++) if (filled(r, c)) return true; return false; };
        const colHas = c => { for (let r = r1 - 1; r <= r2 + 1; r++) if (filled(r, c)) return true; return false; };
        if (r1 > 0 && rowHas(r1 - 1)) { r1--; grew = true; }
        if (r2 < s.rows - 1 && rowHas(r2 + 1)) { r2++; grew = true; }
        if (c1 > 0 && colHas(c1 - 1)) { c1--; grew = true; }
        if (c2 < s.cols - 1 && colHas(c2 + 1)) { c2++; grew = true; }
        if (r2 - r1 > 100000) break;
      }
      return { r1, c1, r2, c2 };
    }

    async function insertTable() {
      ctx.commitEdit();
      const g0 = currentRegion();
      const r = await UI.form('Create Table', [
        { name: 'ref', label: 'Where is the data for your table?', value: rangeText(g0) },
        { name: 'header', label: 'My table has headers', type: 'select', options: [{ value: 'yes', label: 'Yes — the first row is the column names' }, { value: 'no', label: 'No — add a header row' }] },
        { name: 'style', label: 'Style', type: 'select', options: Object.keys(T().STYLES).map(k => ({ value: k, label: k[0].toUpperCase() + k.slice(1) })) }
      ], 'Create');
      if (!r) return;
      const g = parseRange(r.ref);
      if (!g || g.sheet) return UI.toast('Enter a range on this sheet, such as A1:D20', 4000);
      const s = sheet();
      if ((s.tables || []).some(t => overlaps(t, g))) return UI.toast('A table can\'t overlap another table', 4000);
      if (s.merges.some(m => overlaps(m, g))) return UI.toast('A table can\'t contain merged cells', 4000);
      if (!ctx.guard(g)) return;
      ctx.snapshot();
      if (r.header === 'no') {
        // A new header row above the data, as Excel does.
        ctx.shiftStructure('row', g.r1, 1, true);
        g.r2++;
      }
      const t = { name: newTableName(), r1: g.r1, c1: g.c1, r2: Math.max(g.r2, g.r1 + 1), c2: g.c2, totals: false, banded: true, style: r.style, filter: {} };
      // Headers are text and unique.
      headerNames(t).forEach((n, i) => { if (textAt(s, t.r1, t.c1 + i) !== n) ctx.setCell(t.r1, t.c1 + i, { v: n }); });
      s.tables = [...(s.tables || []), t];
      ctx.changed();
      UI.toast(`${t.name} created. Formulas can use ${t.name}[${headerNames(t)[0]}].`, 4000);
    }

    async function tableDesign() {
      const s = sheet(), a = ctx.active();
      const t = tableAt(a.r, a.c);
      if (!t) return UI.toast('Select a cell in a table first (Insert → Table makes one)', 4000);
      const r = await UI.form(`Table Design — ${t.name}`, [
        { name: 'name', label: 'Table name', value: t.name },
        { name: 'ref', label: 'Range (resize the table)', value: rangeText(t) },
        { name: 'style', label: 'Style', type: 'select', options: Object.keys(T().STYLES).map(k => ({ value: k, label: k[0].toUpperCase() + k.slice(1) })), value: t.style },
        { name: 'banded', label: 'Banded rows', type: 'select', options: [{ value: 'yes', label: 'Yes' }, { value: 'no', label: 'No' }], value: t.banded ? 'yes' : 'no' },
        { name: 'totals', label: 'Total row', type: 'select', options: [{ value: 'no', label: 'No' }, { value: 'yes', label: 'Yes' }], value: t.totals ? 'yes' : 'no' }
      ], 'OK');
      if (!r) return;
      const name = r.name.trim();
      if (name !== t.name && (!T().validName(name) || usedName(name))) return UI.toast('Table names start with a letter, have no spaces, and must be unused', 4000);
      const g = parseRange(r.ref);
      if (!g || g.sheet || g.r2 <= g.r1) return UI.toast('Enter a range of at least two rows, such as A1:D20', 4000);
      if ((s.tables || []).some(x => x !== t && overlaps(x, g))) return UI.toast('A table can\'t overlap another table', 4000);
      ctx.snapshot();
      if (name !== t.name) renameTableRefs(t.name, name);
      Object.assign(t, { name, r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2, style: r.style, banded: r.banded === 'yes' });
      setTotals(t, r.totals === 'yes');
      headerNames(t).forEach((n, i) => { if (textAt(s, t.r1, t.c1 + i) !== n) ctx.setCell(t.r1, t.c1 + i, { v: n }); });
      applyFilters(s);
      ctx.changed();
    }
    function setTotals(t, on) {
      if (on === t.totals) return;
      const s = sheet();
      const names = headerNames(t);
      if (on) {
        const below = t.r2 + 1;
        let free = true;
        for (let c = t.c1; c <= t.c2; c++) if (ctx.rawAt(below, c) !== '') free = false;
        if (!free) ctx.shiftStructure('row', below, 1, true);
        t.r2 = below;
        t.totals = true;
        ctx.setCell(t.r2, t.c1, { v: 'Total', s: { b: true } });
        // Sum the last column that holds numbers.
        for (let c = t.c2; c > t.c1; c--) {
          let numeric = false;
          for (let r = t.r1 + 1; r < t.r2 && !numeric; r++) numeric = typeof ctx.engine.value(s.name, r, c) === 'number';
          if (numeric) { ctx.setCell(t.r2, c, { v: `=SUBTOTAL(109,${t.name}[${names[c - t.c1]}])` }); break; }
        }
      } else {
        for (let c = t.c1; c <= t.c2; c++) ctx.setCell(t.r2, c, { v: '' });
        t.r2--;
        t.totals = false;
      }
    }
    function convertToRange() {
      const s = sheet(), a = ctx.active();
      const t = tableAt(a.r, a.c);
      if (!t) return UI.toast('Select a cell in a table first', 3000);
      ctx.snapshot();
      // Formulas keep working: table references become ordinary ranges.
      replaceStructured(t);
      // The table's look stays, as Excel does.
      for (let r = t.r1; r <= t.r2; r++) for (let c = t.c1; c <= t.c2; c++) {
        const ts = T().styleAt([t], r, c);
        const own = ((s.cells[addr(r, c)] || {}).s) || {};
        const st = {};
        if (ts.bg && !own.bg) st.bg = ts.bg;
        if (ts.color && !own.color) st.color = ts.color;
        if (ts.b && own.b === undefined) st.b = true;
        if (Object.keys(st).length) ctx.setCell(r, c, { s: st });
      }
      s.tables = s.tables.filter(x => x !== t);
      s.slicers = (s.slicers || []).filter(x => x.table !== t.name);
      applyFilters(s);
      ctx.changed();
      UI.toast(`${t.name} is now a normal range`);
    }
    /** Rewrites Table[…] references in every formula (table renamed or removed). */
    function rewriteFormulas(fn) {
      book().sheets.forEach(s => Object.values(s.cells).forEach(cell => {
        if (typeof cell.v === 'string' && cell.v[0] === '=' && cell.v.includes('[')) { const n = fn(cell.v, s); if (n !== cell.v) cell.v = n; }
      }));
    }
    function renameTableRefs(oldName, newName) {
      const re = new RegExp(`(^|[^\\w.])${oldName.replace(/[.]/g, '\\.')}(?=\\[)`, 'gi');
      rewriteFormulas(f => f.replace(re, (m, p) => p + newName));
      book().sheets.forEach(s => (s.slicers || []).forEach(x => { if (x.table === oldName) x.table = newName; }));
    }
    function replaceStructured(t) {
      const s0 = book().sheets.find(s => (s.tables || []).includes(t));
      rewriteFormulas((f, s) => f.replace(/((?:[A-Za-z_][\w.]*)?\[(?:[^\[\]']|'.|\[(?:[^\[\]']|'.)*\])*\])/g, m => {
        let spec;
        try { spec = F.parseStructured(m); } catch { return m; }
        if (spec.table ? spec.table.toLowerCase() !== t.name.toLowerCase() : false) return m;
        if (spec.thisRow) {
          // [@Col] in a row: that row's cell, relative.
          const g = T().resolve({ sheets: [s0] }, { ...spec, thisRow: false, parts: new Set(['#data']) }, s0.name, null, (sh, r, c) => textAt(sh, r, c));
          return g ? (s !== s0 ? quoteSheet(s0.name) + '!' : '') + `INDEX(${addr(g.r1, g.c1)}:${addr(g.r2, g.c2)},ROW()-${g.r1})` : m;
        }
        const g = T().resolve({ sheets: [s0] }, spec, s0.name, null, (sh, r, c) => textAt(sh, r, c));
        if (!g) return m;
        return (s !== s0 ? quoteSheet(s0.name) + '!' : '') + '$' + rangeText(g).replace(/([A-Z]+)(\d+)/g, '$1$$$2').replace(':', ':$');
      }));
    }

    /** Rows hidden by tables' filters, merged with the sheet's other hidden rows. */
    function applyFilters(s = sheet()) {
      const inData = r => (s.tables || []).some(t => r > t.r1 && r <= (t.totals ? t.r2 - 1 : t.r2));
      const keep = (s.hiddenRows || []).filter(r => !inData(r));
      const add = (s.tables || []).flatMap(t => T().hiddenRows(t, (r, c) => shown(s, r, c)));
      s.hiddenRows = [...new Set([...keep, ...add])].sort((a, b) => a - b);
    }

    function sortTable(t, col, desc) {
      const s = sheet();
      const start = t.r1 + 1, end = t.totals ? t.r2 - 1 : t.r2;
      ctx.snapshot();
      const rows = [];
      for (let r = start; r <= end; r++) {
        const row = [];
        for (let c = t.c1; c <= t.c2; c++) row.push(s.cells[addr(r, c)] ? JSON.parse(JSON.stringify(s.cells[addr(r, c)])) : null);
        rows.push({ key: ctx.engine.value(s.name, r, col), row, r });
      }
      rows.sort((a, b) => {
        const x = a.key, y = b.key, blank = v => v === '' || v === null || v === undefined;
        if (blank(x) && blank(y)) return 0;
        if (blank(x)) return 1;
        if (blank(y)) return -1;
        const d = typeof x === 'number' && typeof y === 'number' ? x - y : typeof x === 'number' ? -1 : typeof y === 'number' ? 1 : String(x).localeCompare(String(y), undefined, { numeric: true, sensitivity: 'base' });
        return desc ? -d : d;
      });
      rows.forEach((item, i) => item.row.forEach((cell, j) => {
        const r = start + i, c = t.c1 + j;
        if (!cell) { delete s.cells[addr(r, c)]; return; }
        if (typeof cell.v === 'string' && cell.v[0] === '=') cell.v = F.shiftFormula(cell.v, r - item.r, 0);
        s.cells[addr(r, c)] = cell;
      }));
      ctx.touched(s);
      applyFilters(s);
      ctx.changed();
    }

    /** The filter and sort menu of a table column. */
    function openFilterMenu(t, c, at) {
      document.querySelectorAll('.tbl-menu').forEach(m => m.remove());
      const s = sheet();
      const k = String(c - t.c1);
      const end = t.totals ? t.r2 - 1 : t.r2;
      const values = [...new Set(Array.from({ length: end - t.r1 }, (_, i) => shown(s, t.r1 + 1 + i, c)))].sort((a, b) => a.localeCompare(b, undefined, { numeric: true })).slice(0, 2000);
      const allowed = t.filter[k] ? new Set(t.filter[k]) : null;
      const m = document.createElement('div');
      m.className = 'tbl-menu';
      m.setAttribute('role', 'dialog');
      m.setAttribute('aria-label', 'Filter ' + headerNames(t)[c - t.c1]);
      m.innerHTML = '<button data-a="asc">A↓ Sort A to Z</button><button data-a="desc">Z↓ Sort Z to A</button><hr><input type="search" placeholder="Search" aria-label="Search values"><div class="vals"></div><div class="row"><button data-a="clear">Clear filter</button><span class="sp"></span><button data-a="cancel">Cancel</button><button class="primary" data-a="ok">OK</button></div>';
      const box = m.querySelector('.vals');
      const items = [['(Select All)', null], ...values.map(v => [v === '' ? '(Blanks)' : v, v])].map(([label, v]) => {
        const l = document.createElement('label');
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.checked = v === null ? !allowed : !allowed || allowed.has(v);
        l.append(cb, ' ' + label);
        box.appendChild(l);
        return { l, cb, v, label };
      });
      const all = items[0];
      all.cb.addEventListener('change', () => items.slice(1).forEach(i => { if (!i.l.hidden) i.cb.checked = all.cb.checked; }));
      m.querySelector('input[type=search]').addEventListener('input', e => { const q = e.target.value.toLowerCase(); items.slice(1).forEach(i => { i.l.hidden = !!q && !i.label.toLowerCase().includes(q); }); });
      const close = () => { m.remove(); document.removeEventListener('mousedown', outside, true); };
      const outside = e => { if (!m.contains(e.target)) close(); };
      m.addEventListener('click', e => {
        const a = e.target.dataset && e.target.dataset.a;
        if (!a) return;
        if (a === 'cancel') return close();
        close();
        if (a === 'asc' || a === 'desc') return sortTable(t, c, a === 'desc');
        ctx.snapshot();
        if (a === 'clear') delete t.filter[k];
        else {
          const chosen = items.slice(1).filter(i => i.cb.checked && !i.l.hidden).map(i => i.v);
          if (chosen.length === values.length) delete t.filter[k]; else t.filter[k] = chosen;
        }
        applyFilters(s);
        ctx.changed();
      });
      m.addEventListener('keydown', e => { if (e.key === 'Escape') close(); e.stopPropagation(); });
      document.body.appendChild(m);
      const r = at.getBoundingClientRect();
      m.style.left = Math.min(innerWidth - m.offsetWidth - 8, r.left) + 'px';
      m.style.top = Math.min(innerHeight - m.offsetHeight - 8, r.bottom + 2) + 'px';
      setTimeout(() => document.addEventListener('mousedown', outside, true));
      m.querySelector('[data-a=ok]').focus();
    }

    /**
     * After a cell is typed in: a row or column typed next to a table joins it, a formula typed
     * into an empty table column fills the whole column, and renamed headers carry into formulas.
     */
    function afterEdit(r, c, raw, before) {
      const s = sheet();
      let touched = false;
      for (const t of s.tables || []) {
        // A row typed just under the table (with no total row) joins it; calculated columns continue.
        if (raw !== '' && !t.totals && r === t.r2 + 1 && c >= t.c1 && c <= t.c2) {
          t.r2++;
          for (let cc = t.c1; cc <= t.c2; cc++) {
            const above = ctx.rawAt(t.r2 - 1, cc);
            if (cc !== c && ctx.rawAt(t.r2, cc) === '' && typeof above === 'string' && above[0] === '=') ctx.setCell(t.r2, cc, { v: F.shiftFormula(above, 1, 0) });
          }
          touched = true;
          continue;
        }
        // A column typed just right of the table joins it.
        if (raw !== '' && c === t.c2 + 1 && r >= t.r1 && r <= t.r2 && !(s.tables || []).some(x => x !== t && overlaps(x, { r1: t.r1, r2: t.r2, c1: c, c2: c }))) {
          t.c2++;
          const n = headerNames(t)[c - t.c1];
          if (r !== t.r1 && textAt(s, t.r1, c) !== n) ctx.setCell(t.r1, c, { v: n });
          touched = true;
        }
        if (c < t.c1 || c > t.c2) continue;
        const end = t.totals ? t.r2 - 1 : t.r2;
        // A formula typed into an otherwise empty column fills the whole column.
        if (r > t.r1 && r <= end && typeof raw === 'string' && raw[0] === '=') {
          let empty = true;
          for (let rr = t.r1 + 1; rr <= end && empty; rr++) if (rr !== r && ctx.rawAt(rr, c) !== '') empty = false;
          if (empty && end > t.r1 + 1) {
            for (let rr = t.r1 + 1; rr <= end; rr++) if (rr !== r) ctx.setCell(rr, c, { v: F.shiftFormula(raw, rr - r, 0) });
            touched = true;
          }
        } else if (r === t.r1) {
          // A header: never empty, and formulas follow a renamed column.
          const oldName = String(before ?? '').replace(/^'/, '').trim();
          const newName = headerNames(t)[c - t.c1];
          if (raw === '') ctx.setCell(r, c, { v: newName });
          if (oldName && oldName.toLowerCase() !== newName.toLowerCase()) {
            const e = oldName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
            const inside = new RegExp(`\\[(@?)\\[?${e}\\]?\\]`, 'gi');
            const mentions = new RegExp(`(^|[^\\w.])${t.name}\\[`, 'i');
            const name = /^[A-Za-z_][\w.]*$/.test(newName) ? newName : '[' + newName + ']';
            rewriteFormulas((f, sh) => (mentions.test(f) || sh === s ? f.replace(inside, (m, at) => `[${at}${name}]`) : f));
          }
          touched = true;
        }
      }
      return touched;
    }

    /** How a table draws a cell, and whether it shows a filter button there. */
    const tableStyle = (r, c) => T().styleAt(sheet().tables, r, c);

    // =====================================================================================
    // Named ranges
    // =====================================================================================
    function resolveName(name) {
      const n = (book().names || []).find(x => x.name.toUpperCase() === name);
      if (n) return n.ref;
      // A bare table name means its data.
      for (const s of book().sheets) {
        const t = (s.tables || []).find(x => x.name.toUpperCase() === name);
        if (t) return `${quoteSheet(s.name)}!${addr(t.r1 + 1, t.c1)}:${addr(t.totals ? t.r2 - 1 : t.r2, t.c2)}`;
      }
      return null;
    }
    const absRef = (s, g) => `${quoteSheet(s.name)}!$${indexToCol(g.c1)}$${g.r1 + 1}` + (g.r1 === g.r2 && g.c1 === g.c2 ? '' : `:$${indexToCol(g.c2)}$${g.r2 + 1}`);
    function addName(name, ref) {
      if (!T().validName(name)) { UI.toast('Names start with a letter or _, have no spaces, and can\'t look like a cell (A1)', 4500); return false; }
      if (usedName(name)) { UI.toast(`"${name}" is already used`, 3000); return false; }
      book().names = [...(book().names || []), { name, ref: String(ref).replace(/^=/, '') }];
      return true;
    }
    /** Name box: a name goes to its range; a new name names the selection (as in Excel). */
    function nameBoxGo(text) {
      const name = text.trim();
      const ref = resolveName(name.toUpperCase());
      if (ref) {
        const g = parseRange(ref);
        if (!g) { UI.toast(`${name} is a value (=${ref}), not a range`); return true; }
        if (g.sheet) { const i = book().sheets.findIndex(s => s.name.toLowerCase() === g.sheet.toLowerCase()); if (i >= 0 && i !== book().active) ctx.activateSheet(i); }
        ctx.selectRange(g);
        return true;
      }
      if (!T().validName(name)) return false;
      ctx.snapshot();
      if (addName(name, absRef(sheet(), ctx.selRange()))) { ctx.changed(); UI.toast(`Named ${rangeText(ctx.selRange())} "${name}"`); }
      return true;
    }
    async function defineName() {
      const r = await UI.form('New Name', [
        { name: 'name', label: 'Name', value: '' },
        { name: 'ref', label: 'Refers to (a range, a constant like 0.07, or a formula)', value: absRef(sheet(), ctx.selRange()) }
      ], 'OK');
      if (!r) return;
      ctx.snapshot();
      if (addName(r.name.trim(), r.ref.trim())) ctx.changed();
    }
    function nameManager() {
      ctx.commitEdit();
      let names = (book().names || []).map(n => ({ ...n }));
      const back = document.createElement('div');
      back.className = 'modal-backdrop';
      back.innerHTML = '<div class="modal name-manage" role="dialog" aria-label="Name Manager"><h2>Name Manager</h2><div class="nm-list"></div><div class="actions"><button class="new">New…</button><span class="sp"></span><button class="cancel">Cancel</button><button class="primary ok">OK</button></div></div>';
      const list = back.querySelector('.nm-list');
      const valueOf = ref => {
        try { const v = ctx.engine.evalNode(F.parse(String(ref).replace(/\$/g, '')), sheet().name); return v instanceof F.Range ? `{${v.flat().slice(0, 6).map(x => (x instanceof F.FormulaError ? x.error : x)).join(', ')}${v.flat().length > 6 ? ', …' : ''}}` : v instanceof F.FormulaError ? v.error : String(v); } catch (e) { return e instanceof F.FormulaError ? e.error : '#ERROR!'; }
      };
      const draw = () => {
        list.innerHTML = '';
        const tables = allTables().map(t => ({ name: t.name, ref: '(table)', table: true }));
        if (!names.length && !tables.length) { list.innerHTML = '<div class="cf-empty">No names yet. Type a name in the Name Box to name the selection, or click New.</div>'; return; }
        names.forEach((n, i) => {
          const row = document.createElement('div');
          row.className = 'nm-row';
          row.innerHTML = '<b></b><input aria-label="Refers to"><span class="val"></span><button title="Delete name"><span class="ico" data-icon="trash-2"></span></button>';
          row.querySelector('b').textContent = n.name;
          const input = row.querySelector('input');
          input.value = n.ref;
          row.querySelector('.val').textContent = valueOf(n.ref);
          input.addEventListener('change', () => { n.ref = input.value.trim().replace(/^=/, ''); row.querySelector('.val').textContent = valueOf(n.ref); });
          row.querySelector('button').addEventListener('click', () => { names.splice(i, 1); draw(); });
          list.appendChild(row);
        });
        tables.forEach(t => { const row = document.createElement('div'); row.className = 'nm-row table'; row.innerHTML = '<b></b><span class="muted">Table</span>'; row.querySelector('b').textContent = t.name; list.appendChild(row); });
      };
      draw();
      const close = () => back.remove();
      back.querySelector('.cancel').addEventListener('click', close);
      back.addEventListener('keydown', e => { if (e.key === 'Escape') close(); });
      back.querySelector('.new').addEventListener('click', async () => {
        const r = await UI.form('New Name', [{ name: 'name', label: 'Name' }, { name: 'ref', label: 'Refers to', value: absRef(sheet(), ctx.selRange()) }], 'Add');
        if (!r) return;
        const name = r.name.trim();
        if (!T().validName(name) || names.some(n => n.name.toLowerCase() === name.toLowerCase()) || allTables().some(t => t.name.toLowerCase() === name.toLowerCase())) return UI.toast('That name can\'t be used', 3000);
        names.push({ name, ref: r.ref.trim().replace(/^=/, '') });
        draw();
      });
      back.querySelector('.ok').addEventListener('click', () => {
        close();
        ctx.snapshot();
        book().names = names.map(T().cleanName).filter(Boolean);
        ctx.changed();
      });
      document.body.appendChild(back);
      back.querySelector('.ok').focus();
    }
    /** Names follow their cells when rows or columns are inserted or deleted. */
    function shiftNames(axis, index, count, sheetName) {
      book().names = (book().names || []).map(n => {
        const f = F.adjustForInsert('=' + n.ref, axis, index, count, sheetName, '\u0000');
        return { ...n, ref: f.slice(1) };
      });
    }

    // =====================================================================================
    // What-If Analysis
    // =====================================================================================
    const cellRef = (text, s = sheet()) => { const g = parseRange(text); if (!g || g.r1 !== g.r2 || g.c1 !== g.c2) return null; return { s: g.sheet ? ctx.sheetByName(g.sheet) : s, r: g.r1, c: g.c1 }; };
    async function goalSeek() {
      ctx.commitEdit();
      const a = ctx.active();
      const r = await UI.form('Goal Seek', [
        { name: 'set', label: 'Set cell (a formula)', value: addr(a.r, a.c) },
        { name: 'to', label: 'To value', type: 'number', value: '' },
        { name: 'by', label: 'By changing cell (a number)', value: '' }
      ], 'OK');
      if (!r) return;
      const target = cellRef(r.set), change = cellRef(r.by);
      if (!target || !target.s) return UI.toast('Set cell must be a single cell, such as B10', 4000);
      if (!change || !change.s) return UI.toast('By changing cell must be a single cell, such as B2', 4000);
      const raw0 = String(ctx.rawAt(target.r, target.c, target.s));
      if (raw0[0] !== '=') return UI.toast('The set cell must contain a formula', 4000);
      const raw1 = ctx.rawAt(change.r, change.c, change.s);
      if (typeof raw1 === 'string' && raw1[0] === '=') return UI.toast('The changing cell must contain a value, not a formula', 4000);
      const goal = +r.to;
      if (!Number.isFinite(goal) || r.to === '') return UI.toast('Enter the value to reach', 3000);
      const key = addr(change.r, change.c);
      const orig = change.s.cells[key] ? JSON.parse(JSON.stringify(change.s.cells[key])) : null;
      const put = x => { change.s.cells[key] = { ...(orig || {}), v: String(x) }; ctx.engine.invalidate(); };
      const f = x => { put(x); const v = ctx.engine.value(target.s.name, target.r, target.c); return typeof v === 'number' ? v : NaN; };
      const start = +ctx.engine.value(change.s.name, change.r, change.c) || 0;
      const res = SheetsWhatIf.goalSeek(f, start, goal);
      // Put the original back; Keep applies the answer as one undoable step.
      if (orig) change.s.cells[key] = orig; else delete change.s.cells[key];
      ctx.touched(change.s);
      ctx.engine.invalidate();
      const fmtN = n => (Number.isFinite(n) ? +n.toPrecision(10) : '—');
      const msg = res.found ? `Goal Seek found a solution: ${r.by.toUpperCase()} = ${fmtN(res.x)} makes ${r.set.toUpperCase()} = ${fmtN(res.value)}.` : `Goal Seek may not have found a solution. Closest: ${r.by.toUpperCase()} = ${fmtN(res.x)} gives ${fmtN(res.value)}.`;
      const i = await UI.choose('Goal Seek Status', msg, ['Keep', 'Cancel']);
      if (i !== 0) { ctx.changed(false); return; }
      ctx.snapshot();
      ctx.setCell(change.r, change.c, { v: String(+res.x.toPrecision(15)) }, change.s);
      ctx.changed();
    }

    async function solver() {
      ctx.commitEdit();
      const s = sheet(), prev = s.solver || {};
      const r = await UI.form('Solver', [
        { name: 'obj', label: 'Set objective (a formula cell)', value: prev.obj || addr(ctx.active().r, ctx.active().c) },
        { name: 'goal', label: 'To', type: 'select', options: [{ value: 'max', label: 'Max' }, { value: 'min', label: 'Min' }, { value: 'value', label: 'Value of' }], value: prev.goal || 'max' },
        { name: 'target', label: 'Value (for "Value of")', type: 'number', value: prev.target ?? 0 },
        { name: 'vars', label: 'By changing variable cells (e.g. B2:B4)', value: prev.vars || '' },
        { name: 'cons', label: 'Subject to the constraints — one per line, e.g.  B5 <= 100   B2:B4 >= 1   C7 = D7   B2:B4 int', type: 'textarea', value: prev.cons || '' },
        { name: 'nonneg', label: 'Make unconstrained variables non-negative', type: 'select', options: [{ value: 'yes', label: 'Yes' }, { value: 'no', label: 'No' }], value: prev.nonneg === false ? 'no' : 'yes' }
      ], 'Solve');
      if (!r) return;
      const obj = cellRef(r.obj);
      if (!obj || !obj.s || String(ctx.rawAt(obj.r, obj.c, obj.s))[0] !== '=') return UI.toast('The objective must be a single cell with a formula', 4000);
      const vg = parseRange(r.vars);
      if (!vg) return UI.toast('Enter the variable cells, such as B2:B4', 4000);
      const vs = vg.sheet ? ctx.sheetByName(vg.sheet) : s;
      const vars = [];
      for (let rr = vg.r1; rr <= vg.r2; rr++) for (let cc = vg.c1; cc <= vg.c2; cc++) {
        const raw = ctx.rawAt(rr, cc, vs);
        if (typeof raw === 'string' && raw[0] === '=') return UI.toast(`${addr(rr, cc)} has a formula; variable cells must hold values`, 4000);
        vars.push({ r: rr, c: cc });
      }
      if (vars.length > 200) return UI.toast('Solver works with up to 200 variable cells', 4000);
      // Constraints: "<range> <op> <number|cell|range>" or "<range> int".
      const ints = new Set(), cons = [];
      const varIndex = (sh, rr, cc) => (sh === vs ? vars.findIndex(v => v.r === rr && v.c === cc) : -1);
      for (const line of r.cons.split(/\n/).map(x => x.trim()).filter(Boolean)) {
        const mi = /^(.+?)\s+int(eger)?$/i.exec(line);
        if (mi) {
          const g = parseRange(mi[1]);
          if (!g) return UI.toast(`Can't read the constraint "${line}"`, 4000);
          const sh = g.sheet ? ctx.sheetByName(g.sheet) : s;
          for (let rr = g.r1; rr <= g.r2; rr++) for (let cc = g.c1; cc <= g.c2; cc++) { const i = varIndex(sh, rr, cc); if (i >= 0) ints.add(i); }
          continue;
        }
        const m = /^(.+?)\s*(<=|>=|=|<|>)\s*(.+)$/.exec(line);
        const g = m && parseRange(m[1]);
        if (!g) return UI.toast(`Can't read the constraint "${line}" (use e.g. B5 <= 100)`, 4000);
        const op = m[2] === '<' ? '<=' : m[2] === '>' ? '>=' : m[2];
        const rhsG = /^-?[\d.]+(e-?\d+)?$/i.test(m[3].trim()) ? null : parseRange(m[3]);
        if (!rhsG && !Number.isFinite(+m[3])) return UI.toast(`Can't read the right side of "${line}"`, 4000);
        const lsh = g.sheet ? ctx.sheetByName(g.sheet) : s, rsh = rhsG ? (rhsG.sheet ? ctx.sheetByName(rhsG.sheet) : s) : null;
        let k = 0;
        for (let rr = g.r1; rr <= g.r2; rr++) for (let cc = g.c1; cc <= g.c2; cc++, k++) {
          const rr2 = rhsG ? rhsG.r1 + (rhsG.r2 > rhsG.r1 ? rr - g.r1 : 0) : 0, cc2 = rhsG ? rhsG.c1 + (rhsG.c2 > rhsG.c1 ? cc - g.c1 : 0) : 0;
          cons.push({ l: [lsh, rr, cc], op, rhs: rhsG ? [rsh, rr2, cc2] : +m[3] });
        }
      }
      s.solver = { obj: r.obj, goal: r.goal, target: +r.target || 0, vars: r.vars, cons: r.cons, nonneg: r.nonneg === 'yes' };
      const saved = vars.map(v => (vs.cells[addr(v.r, v.c)] ? JSON.parse(JSON.stringify(vs.cells[addr(v.r, v.c)])) : null));
      const num = v => (typeof v === 'number' ? v : typeof v === 'boolean' ? +v : v === '' ? 0 : NaN);
      const evaluate = x => {
        vars.forEach((v, i) => { const k = addr(v.r, v.c); vs.cells[k] = { ...(saved[i] || {}), v: String(x[i]) }; });
        ctx.engine.invalidate();
        return {
          obj: num(ctx.engine.value(obj.s.name, obj.r, obj.c)),
          cons: cons.map(c => ({ lhs: num(ctx.engine.value(c.l[0].name, c.l[1], c.l[2])), op: c.op, rhs: Array.isArray(c.rhs) ? num(ctx.engine.value(c.rhs[0].name, c.rhs[1], c.rhs[2])) : c.rhs }))
        };
      };
      UI.toast('Solving…', 1500);
      await new Promise(res => setTimeout(res, 30));
      const x0 = vars.map((v, i) => num(ctx.engine.value(vs.name, v.r, v.c)) || 0);
      const res = SheetsWhatIf.solve({ x0, evaluate, goal: r.goal, target: +r.target || 0, nonneg: r.nonneg === 'yes', integers: ints });
      // Put the originals back; Keep applies the solution as one undoable step.
      vars.forEach((v, i) => { const k = addr(v.r, v.c); if (saved[i]) vs.cells[k] = saved[i]; else delete vs.cells[k]; });
      ctx.touched(vs);
      ctx.engine.invalidate();
      const objText = Number.isFinite(res.obj) ? +res.obj.toPrecision(10) : '—';
      const msg = res.ok ? `Solver found a solution. All constraints are satisfied. Objective: ${objText}.` : `Solver could not find a solution that satisfies all the constraints. Best found: objective ${objText} (constraints off by ${res.violation.toPrecision(3)}).`;
      const i = await UI.choose('Solver Results', msg, ['Keep Solver Solution', 'Restore Original Values']);
      if (i !== 0) { ctx.changed(); return; }
      ctx.snapshot();
      vars.forEach((v, j) => ctx.setCell(v.r, v.c, { v: String(+res.x[j].toPrecision(15)) }, vs));
      ctx.changed();
    }

    // =====================================================================================
    // Sparklines
    // =====================================================================================
    async function insertSparklines(type) {
      ctx.commitEdit();
      const g0 = ctx.selRange();
      const r = await UI.form(`Insert ${type === 'winloss' ? 'Win/Loss' : type[0].toUpperCase() + type.slice(1)} Sparklines`, [
        { name: 'src', label: 'Data range (one row per sparkline, e.g. B2:M5)', value: rangeText(g0) },
        { name: 'at', label: 'Location range (one cell per row of data, e.g. N2:N5)', value: addr(g0.r1, g0.c2 + 1) + (g0.r2 > g0.r1 ? ':' + addr(g0.r2, g0.c2 + 1) : '') },
        { name: 'color', label: 'Colour', type: 'color', value: '#376092' }
      ], 'OK');
      if (!r) return;
      const src = parseRange(r.src), at = parseRange(r.at);
      if (!src || !at || at.sheet) return UI.toast('Enter ranges such as B2:M5 and N2:N5', 4000);
      const cells = [];
      for (let rr = at.r1; rr <= at.r2; rr++) for (let cc = at.c1; cc <= at.c2; cc++) cells.push([rr, cc]);
      const rows = src.r2 - src.r1 + 1, cols = src.c2 - src.c1 + 1;
      const prefix = src.sheet ? quoteSheet(src.sheet) + '!' : '';
      let items;
      if (cells.length === 1) items = [[cells[0], rangeText(src)]];
      else if (cells.length === rows) items = cells.map((cell, i) => [cell, rangeText({ r1: src.r1 + i, c1: src.c1, r2: src.r1 + i, c2: src.c2 })]);
      else if (cells.length === cols) items = cells.map((cell, i) => [cell, rangeText({ r1: src.r1, c1: src.c1 + i, r2: src.r2, c2: src.c1 + i })]);
      else return UI.toast(`The location needs ${rows} cell${rows === 1 ? '' : 's'} (one per row of data)`, 4000);
      const s = sheet();
      ctx.snapshot();
      const list = (s.sparklines || []).filter(sp => !items.some(([[rr, cc]]) => sp.cell === addr(rr, cc)));
      items.forEach(([[rr, cc], ref]) => { const sp = SheetsSpark.clean({ type, src: prefix + ref, cell: addr(rr, cc), color: r.color }); if (sp) list.push(sp); });
      s.sparklines = list;
      ctx.changed();
    }
    function clearSparklines() {
      const s = sheet(), g = ctx.selRange();
      const keep = (s.sparklines || []).filter(sp => { const p = parseRef(sp.cell); return !(p.row >= g.r1 && p.row <= g.r2 && p.col >= g.c1 && p.col <= g.c2); });
      if (keep.length === (s.sparklines || []).length) return UI.toast('No sparklines in the selection');
      ctx.snapshot();
      s.sparklines = keep;
      ctx.changed();
    }
    const sparkMap = new WeakMap();
    function sparkAt(r, c) {
      const s = sheet();
      if (!s.sparklines || !s.sparklines.length) return null;
      let m = sparkMap.get(s.sparklines);
      if (!m) { m = new Map(s.sparklines.map(sp => [sp.cell, sp])); sparkMap.set(s.sparklines, m); }
      const sp = m.get(addr(r, c));
      if (!sp) return null;
      let values = [];
      try { const v = ctx.engine.evalNode(F.parse(sp.src), s.name); values = v instanceof F.Range ? v.flat() : [v]; } catch { values = []; }
      return SheetsSpark.svg(sp, values, ctx.colWidth(c), ctx.rowH(r));
    }
    const sparkCells = () => (sheet().sparklines || []).map(sp => { const p = parseRef(sp.cell); return [p.row, p.col]; });

    // =====================================================================================
    // Slicers
    // =====================================================================================
    async function insertSlicer() {
      const a = ctx.active();
      const t = tableAt(a.r, a.c);
      if (!t) return UI.toast('Select a cell in a table first: slicers filter a table', 4000);
      const names = headerNames(t);
      const r = await UI.form('Insert Slicer', [{ name: 'col', label: 'Filter by column', type: 'select', options: names, value: names[a.c - t.c1] }], 'OK');
      if (!r) return;
      const s = sheet();
      const box = ctx.rangeBox({ r1: t.r1, c1: t.c2 + 1, r2: t.r1, c2: t.c2 + 1 });
      ctx.snapshot();
      const n = (s.slicers || []).length;
      s.slicers = [...(s.slicers || []), T().cleanSlicer({ table: t.name, col: r.col, x: box.x + 20 + n * 24, y: box.y + n * 24, w: 180, h: 220 })];
      ctx.changed();
    }
    function renderSlicers(layer) {
      layer.innerHTML = '';
      const s = sheet();
      (s.slicers || []).forEach((sl, i) => {
        const t = (s.tables || []).find(x => x.name === sl.table);
        const box = document.createElement('div');
        box.className = 'slicer';
        Object.assign(box.style, { left: sl.x + 'px', top: sl.y + 'px', width: sl.w + 'px', height: sl.h + 'px' });
        box.innerHTML = '<div class="hd"><span class="t"></span><button data-a="clear" title="Clear filter"><span class="ico" data-icon="filter-x"></span></button><button data-a="del" title="Remove slicer"><span class="ico" data-icon="x"></span></button></div><div class="items"></div><div class="grip"></div>';
        box.querySelector('.t').textContent = sl.col;
        layer.appendChild(box);
        if (!t) { box.querySelector('.items').textContent = `Table ${sl.table} no longer exists`; }
        const names = t ? headerNames(t) : [];
        const ci = names.findIndex(n => n.toLowerCase() === sl.col.toLowerCase());
        if (t && ci >= 0) {
          const c = t.c1 + ci, k = String(ci), end = t.totals ? t.r2 - 1 : t.r2;
          const values = [...new Set(Array.from({ length: end - t.r1 }, (_, j) => shown(s, t.r1 + 1 + j, c)))].sort((a, b) => a.localeCompare(b, undefined, { numeric: true })).slice(0, 500);
          const allowed = t.filter[k] ? new Set(t.filter[k]) : null;
          values.forEach(v => {
            const b = document.createElement('button');
            b.textContent = v === '' ? '(blank)' : v;
            b.setAttribute('aria-pressed', String(!allowed || allowed.has(v)));
            b.className = !allowed || allowed.has(v) ? 'on' : '';
            b.addEventListener('click', e => {
              ctx.snapshot();
              const cur = t.filter[k] ? new Set(t.filter[k]) : null;
              if (e.ctrlKey || e.metaKey) {
                const set = cur || new Set(values);
                if (set.has(v)) set.delete(v); else set.add(v);
                if (set.size === values.length || !set.size) delete t.filter[k]; else t.filter[k] = [...set];
              } else if (cur && cur.size === 1 && cur.has(v)) delete t.filter[k];
              else t.filter[k] = [v];
              applyFilters(s);
              ctx.changed();
            });
            box.querySelector('.items').appendChild(b);
          });
          box.querySelector('[data-a=clear]').disabled = !allowed;
        }
        box.querySelector('.hd').addEventListener('mousedown', e => {
          if (e.target.tagName === 'BUTTON') return;
          e.preventDefault();
          const sx = e.clientX, sy = e.clientY, ox = sl.x, oy = sl.y, z = ctx.zoomLevel();
          const mv = ev => { sl.x = Math.max(0, Math.round(ox + (ev.clientX - sx) / z)); sl.y = Math.max(0, Math.round(oy + (ev.clientY - sy) / z)); box.style.left = sl.x + 'px'; box.style.top = sl.y + 'px'; };
          const up = () => { removeEventListener('mousemove', mv); removeEventListener('mouseup', up); ctx.state.markDirty(); };
          addEventListener('mousemove', mv); addEventListener('mouseup', up);
        });
        box.querySelector('.grip').addEventListener('mousedown', e => {
          e.preventDefault();
          const sx = e.clientX, sy = e.clientY, ow = sl.w, oh = sl.h, z = ctx.zoomLevel();
          const mv = ev => { sl.w = Math.max(120, Math.round(ow + (ev.clientX - sx) / z)); sl.h = Math.max(80, Math.round(oh + (ev.clientY - sy) / z)); box.style.width = sl.w + 'px'; box.style.height = sl.h + 'px'; };
          const up = () => { removeEventListener('mousemove', mv); removeEventListener('mouseup', up); ctx.state.markDirty(); };
          addEventListener('mousemove', mv); addEventListener('mouseup', up);
        });
        box.querySelector('[data-a=clear]').addEventListener('click', () => { if (!t) return; ctx.snapshot(); delete t.filter[String(ci)]; applyFilters(s); ctx.changed(); });
        box.querySelector('[data-a=del]').addEventListener('click', () => { ctx.snapshot(); s.slicers.splice(i, 1); ctx.changed(); });
      });
    }

    // =====================================================================================
    // Scripts (Automate)
    // =====================================================================================
    let pane = null, curScript = 0;
    const trusted = new Set(); // scripts written here, or confirmed before running
    const EXAMPLE = `// Runs on this workbook. Examples:
//   const sheet = workbook.getActiveWorksheet();
//   sheet.getRange("A1").setValue("Hello");
//   const values = sheet.getRange("A1:B5").getValues();
//   sheet.getRange("C1").setFormula("=SUM(A1:A5)");
//   sheet.getRange("A1:C1").getFormat().setBold(true);
//   console.log(sheet.getRange("C1").getValue());
function main(workbook) {
  const sheet = workbook.getActiveWorksheet();
  const used = sheet.getUsedRange();
  console.log("Used range: " + used.getAddress());
}
`;
    function scriptsPane(show = !pane || pane.hidden) {
      if (!pane) {
        pane = document.createElement('aside');
        pane.className = 'script-pane';
        pane.setAttribute('aria-label', 'Scripts');
        pane.innerHTML = `<header><b>Scripts</b><span class="sp"></span><button class="x" title="Close"><span class="ico" data-icon="x"></span></button></header>
          <div class="row"><select aria-label="Script"></select><button data-a="new" title="New script"><span class="ico" data-icon="plus"></span></button><button data-a="ren" title="Rename"><span class="ico" data-icon="pencil"></span></button><button data-a="del" title="Delete script">🗑</button></div>
          <textarea spellcheck="false" aria-label="Script code"></textarea>
          <div class="row"><button class="primary" data-a="run">▶ Run</button><span class="sp"></span><span class="hint">Ctrl+Enter runs</span></div>
          <pre class="out" aria-live="polite"></pre>`;
        document.body.appendChild(pane);
        pane.querySelector('.x').addEventListener('click', () => scriptsPane(false));
        const ta = pane.querySelector('textarea');
        ta.addEventListener('input', () => { const sc = (book().scripts || [])[curScript]; if (sc) { sc.code = ta.value; trusted.add(sc.code); ctx.state.markDirty(); } });
        ta.addEventListener('keydown', e => {
          e.stopPropagation();
          if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); runScript(); }
          if (e.key === 'Tab') { e.preventDefault(); document.execCommand('insertText', false, '  '); }
        });
        pane.querySelector('select').addEventListener('change', e => { curScript = +e.target.value; drawPane(); });
        pane.addEventListener('click', async e => {
          const a = e.target.dataset && e.target.dataset.a;
          if (a === 'new') return newScript();
          if (a === 'run') return runScript();
          const sc = (book().scripts || [])[curScript];
          if (!sc) return;
          if (a === 'ren') { const n = await UI.prompt('Rename Script', 'Name', sc.name); if (n) { sc.name = n.trim().slice(0, 60); ctx.state.markDirty(); drawPane(); } }
          if (a === 'del' && await UI.confirm('Delete Script', `Delete "${sc.name}"?`)) { book().scripts.splice(curScript, 1); curScript = 0; ctx.state.markDirty(); drawPane(); }
        });
      }
      pane.hidden = !show;
      document.body.classList.toggle('script-open', show);
      if (show) { if (!(book().scripts || []).length) newScript(); else drawPane(); }
    }
    function drawPane() {
      const list = book().scripts || [];
      const sel = pane.querySelector('select');
      sel.innerHTML = list.map((sc, i) => `<option value="${i}">${esc(sc.name)}</option>`).join('');
      curScript = Math.min(curScript, Math.max(0, list.length - 1));
      sel.value = String(curScript);
      const ta = pane.querySelector('textarea');
      ta.value = list[curScript] ? list[curScript].code : '';
      ta.disabled = !list.length;
    }
    function newScript() {
      const list = book().scripts = book().scripts || [];
      let i = list.length + 1;
      while (list.some(sc => sc.name === 'Script ' + i)) i++;
      list.push({ name: 'Script ' + i, code: EXAMPLE });
      trusted.add(EXAMPLE);
      curScript = list.length - 1;
      ctx.state.markDirty();
      if (!pane || pane.hidden) scriptsPane(true); else drawPane();
      pane.querySelector('textarea').focus();
    }
    async function runScript() {
      const sc = (book().scripts || [])[curScript];
      if (!sc) return;
      const out = pane.querySelector('.out');
      if (!trusted.has(sc.code)) {
        const ok = await UI.confirm('Run Script', `"${sc.name}" came with this workbook. A script can change anything in the workbook. Run it only if you trust where the workbook came from.`);
        if (!ok) return;
        trusted.add(sc.code);
      }
      ctx.commitEdit();
      out.textContent = 'Running…';
      out.className = 'out';
      const g = ctx.selRange();
      const data = { sheets: book().sheets.map(s => ({ name: s.name, cells: Object.fromEntries(Object.entries(s.cells).map(([k, c]) => [k, c.v ?? '']).filter(([, v]) => v !== '')) })), names: book().names || [] };
      const worker = new Worker('js/sheets-script-worker.js');
      const res = await new Promise(resolve => {
        const timer = setTimeout(() => { worker.terminate(); resolve({ ok: false, error: 'The script ran for more than 15 seconds and was stopped', log: [] }); }, 15000);
        worker.onmessage = e => { clearTimeout(timer); worker.terminate(); resolve(e.data); };
        worker.onerror = e => { clearTimeout(timer); worker.terminate(); resolve({ ok: false, error: e.message || 'The script failed', log: [] }); };
        worker.postMessage({ code: sc.code, book: data, active: sheet().name, selection: rangeText(g) });
      });
      const lines = [...(res.log || [])];
      if (!res.ok) { lines.push(`Error${res.line ? ` (line ${Math.max(1, res.line - 2)})` : ''}: ${res.error}`); out.className = 'out err'; }
      if (res.ok && res.ops.length) {
        ctx.snapshot();
        let added = false;
        for (const op of res.ops) {
          if (op.t === 'addSheet') { ctx.addSheetNamed(op.name); added = true; continue; }
          const s = ctx.sheetByName(op.sheet);
          if (!s || !Number.isInteger(op.r) || !Number.isInteger(op.c) || op.r < 0 || op.c < 0 || op.r >= 1048576 || op.c >= 16384) continue;
          if (op.t === 'set') ctx.setCell(op.r, op.c, { v: String(op.v) }, s);
          else if (op.t === 'style') { const st = ctx.cleanStyle(op.s); if (st) ctx.setCell(op.r, op.c, { s: st }, s); }
        }
        if (added) ctx.buildGrid(); else ctx.changed();
        lines.push(`Done: ${res.ops.length} change${res.ops.length === 1 ? '' : 's'} (Ctrl+Z undoes them all).`);
      } else if (res.ok) lines.push('Done.');
      out.textContent = lines.join('\n');
    }

    // =====================================================================================
    // Data from the web
    // =====================================================================================
    async function fromWeb() {
      ctx.commitEdit();
      const r = await UI.form('From Web', [
        { name: 'url', label: 'Address of a CSV or JSON file, or a web page with tables', value: 'https://' },
        { name: 'dest', label: 'Put the data', type: 'select', options: [{ value: 'new', label: 'In a new sheet' }, { value: 'here', label: 'At the selected cell' }] }
      ], 'Get Data');
      if (!r) return;
      UI.toast('Getting data…', 2000);
      let parsed;
      try { const d = await SheetsWeb.download(r.url.trim()); parsed = SheetsWeb.parse(d.bytes, d.contentType, r.url); } catch (err) { return UI.toast(err.message, 6000); }
      let index = 0;
      if (parsed.tables.length > 1) {
        const pick = await UI.form('Choose a Table', [{ name: 't', label: `${parsed.tables.length} tables on the page`, type: 'select', options: parsed.tables.map((t, i) => ({ value: String(i), label: `${t.name} — ${t.rows.length} rows × ${Math.max(...t.rows.map(x => x.length))} columns` })) }], 'Load');
        if (!pick) return;
        index = +pick.t;
      }
      const rows = parsed.tables[index].rows;
      if (!rows.length) return UI.toast('No data found there', 4000);
      ctx.snapshot();
      let s = sheet(), at = { r: ctx.active().r, c: ctx.active().c };
      if (r.dest === 'new') { s = ctx.addSheetNamed(ctx.uniqueName(webSheetName(r.url))); at = { r: 0, c: 0 }; }
      const q = { url: r.url.trim(), table: index, r: at.r, c: at.c, h: 0, w: 0 };
      writeRows(s, q, rows);
      s.queries = [...(s.queries || []).filter(x => !(x.r === q.r && x.c === q.c)), q];
      trustedUrls.add(q.url);
      if (r.dest === 'new') ctx.buildGrid(); else ctx.changed();
      UI.toast(`${rows.length} row${rows.length === 1 ? '' : 's'} loaded. Data → Refresh All gets the latest.`, 4000);
    }
    const webSheetName = url => { try { return (new URL(url).hostname.replace(/^www\./, '') || 'Web').slice(0, 28); } catch { return 'Web'; } };
    /** Writes data for a query, clearing what an earlier refresh put there. */
    function writeRows(s, q, rows) {
      for (let r = 0; r < q.h; r++) for (let c = 0; c < q.w; c++) ctx.setCell(q.r + r, q.c + c, { v: '' }, s);
      let w = 0;
      rows.forEach((row, i) => {
        w = Math.max(w, row.length);
        row.forEach((v, j) => {
          if (v === '' || v === null || v === undefined) return;
          // Text that looks like a formula stays text: downloaded data never becomes a formula.
          const raw = typeof v === 'number' ? String(v) : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v)[0] === '=' ? "'" + v : String(v);
          ctx.setCell(q.r + i, q.c + j, { v: raw }, s);
        });
      });
      for (let j = 0; j < w; j++) ctx.setCell(q.r, q.c + j, { s: { b: true } }, s);
      q.h = rows.length; q.w = w;
    }
    // Web addresses set up here, or already allowed: a workbook from someone else (a file, or a
    // live session) doesn't fetch from the web until its addresses are allowed, as in Excel.
    const trustedUrls = new Set();
    async function refreshQueries() {
      const list = book().sheets.flatMap(s => (s.queries || []).map(q => [s, q]));
      if (!list.length) return 0;
      const unknown = [...new Set(list.map(([, q]) => q.url).filter(u => !trustedUrls.has(u)))];
      if (unknown.length) {
        const hosts = [...new Set(unknown.map(u => { try { return new URL(u).host; } catch { return u; } }))];
        const ok = await UI.confirm('Data from the Web', `This workbook gets data from ${hosts.slice(0, 5).join(', ')}${hosts.length > 5 ? ` and ${hosts.length - 5} more` : ''}. Refresh it only if you trust where the workbook came from.`);
        if (!ok) return 0;
        unknown.forEach(u => trustedUrls.add(u));
      }
      let n = 0, failed = 0;
      for (const [s, q] of list) {
        try {
          const d = await SheetsWeb.download(q.url);
          const p = SheetsWeb.parse(d.bytes, d.contentType, q.url);
          const t = p.tables[Math.min(q.table || 0, p.tables.length - 1)];
          if (!n) ctx.snapshot();
          writeRows(s, q, t.rows);
          n++;
        } catch { failed++; }
      }
      if (n) ctx.changed();
      UI.toast(`${n} web quer${n === 1 ? 'y' : 'ies'} refreshed${failed ? `, ${failed} failed` : ''}`, 3500);
      return n;
    }

    // =====================================================================================
    // Clean & Transform (Power Query–style): steps on a copy of a range, loaded to a new sheet
    // =====================================================================================
    const refOfArea = g => `${F.indexToCol(g.c1)}${g.r1 + 1}:${F.indexToCol(g.c2)}${g.r2 + 1}`;
    /** The values of a range on a sheet, as a table (the first row is the header). */
    function readTable(srcSheet, g) {
      const rows = [];
      for (let r = g.r1; r <= g.r2 && rows.length <= 1000000; r++) {
        const row = [];
        for (let c = g.c1; c <= g.c2; c++) { const v = ctx.engine.value(srcSheet.name, r, c); row.push(v && typeof v === 'object' && v.error ? v.error : v); }
        rows.push(row);
      }
      return rows;
    }
    // icons: wand-sparkles scissors-line-dashed text-cursor-input arrow-down-wide-narrow list-filter copy-minus combine arrow-down-to-line rotate-3d table-2 trash-2 rows-3 group case-sensitive replace eraser hash
    async function cleanData() {
      ctx.commitEdit();
      const g0 = currentRegion();
      if (g0.r1 === g0.r2 && g0.c1 === g0.c2) return UI.toast('Select the data to clean (with its header row)', 4000);
      const src = sheet(), g = { ...g0, r2: Math.min(g0.r2, g0.r1 + 1000000) };
      const source = readTable(src, g);
      const steps = [];
      const back = document.createElement('div');
      back.className = 'modal-backdrop';
      back.innerHTML = `<div class="modal clean-editor" role="dialog" aria-label="Clean and Transform">
        <h2>Clean &amp; Transform</h2>
        <p class="hint"></p>
        <div class="ce-tools" role="toolbar" aria-label="Transform"></div>
        <div class="ce-body"><div class="ce-preview" tabindex="0"><table></table></div>
        <aside class="ce-steps"><b>Applied steps</b><ol></ol></aside></div>
        <div class="actions"><button class="cancel">Cancel</button><button class="primary load">Load to new sheet</button></div></div>`;
      const q = sel => back.querySelector(sel);
      q('.hint').textContent = `From ${src.name}!${refOfArea(g)}. Each step works on the result of the one before; Data → Refresh All runs them again on the latest data.`;
      let result = SheetsClean.run(source, steps);
      const head = () => result[0] || [];
      function draw() {
        result = SheetsClean.run(source, steps);
        const t = q('.ce-preview table');
        t.innerHTML = '';
        const thead = document.createElement('thead'), tr = document.createElement('tr');
        head().forEach(h => { const th = document.createElement('th'); th.textContent = h; tr.appendChild(th); });
        thead.appendChild(tr); t.appendChild(thead);
        const tb = document.createElement('tbody');
        result.slice(1, 101).forEach(row => { const r = document.createElement('tr'); row.forEach(v => { const td = document.createElement('td'); td.textContent = v === '' || v === null || v === undefined ? '' : String(v); if (typeof v === 'number') td.className = 'num'; r.appendChild(td); }); tb.appendChild(r); });
        t.appendChild(tb);
        const ol = q('.ce-steps ol');
        ol.innerHTML = '';
        const li0 = document.createElement('li'); li0.textContent = `Source (${source.length - 1} rows)`; li0.className = 'src'; ol.appendChild(li0);
        steps.forEach((st, i) => {
          const li = document.createElement('li');
          const span = document.createElement('span'); span.textContent = SheetsClean.describe(st);
          const x = document.createElement('button'); x.title = 'Remove this step'; x.setAttribute('aria-label', 'Remove step: ' + span.textContent); x.innerHTML = '<span class="ico" data-icon="trash-2"></span>';
          x.addEventListener('click', () => { steps.splice(i, 1); draw(); });
          li.append(span, x); ol.appendChild(li);
        });
        q('.load').textContent = `Load ${Math.max(0, result.length - 1).toLocaleString()} rows to a new sheet`;
      }
      const colField = (name, label, multi) => ({ name, label, type: 'select', options: (multi ? [{ value: '*', label: 'All columns' }] : []).concat(head().map(h => ({ value: h, label: h }))) });
      const pickCols = v => (v === '*' ? head().slice() : [v]);
      const TOOLS = [
        ['Trim spaces', 'eraser', async () => { const r = await UI.form('Trim Spaces', [colField('c', 'Column', true)]); return r && [{ op: 'trim', cols: pickCols(r.c) }, { op: 'clean', cols: pickCols(r.c) }]; }],
        ['Change case', 'case-sensitive', async () => { const r = await UI.form('Change Case', [colField('c', 'Column', true), { name: 'to', label: 'To', type: 'select', options: [{ value: 'proper', label: 'Capitalize Each Word' }, { value: 'upper', label: 'UPPERCASE' }, { value: 'lower', label: 'lowercase' }] }]); return r && { op: 'case', cols: pickCols(r.c), to: r.to }; }],
        ['Replace values', 'replace', async () => { const r = await UI.form('Replace Values', [colField('c', 'Column', true), { name: 'find', label: 'Find' }, { name: 'with', label: 'Replace with' }, { name: 'whole', label: 'Match', type: 'select', options: [{ value: '', label: 'Anywhere in the cell' }, { value: '1', label: 'The whole cell' }] }]); return r && r.find !== '' && { op: 'replace', cols: pickCols(r.c), find: r.find, with: r.with, whole: r.whole === '1' }; }],
        ['Change type', 'hash', async () => { const r = await UI.form('Change Type', [colField('c', 'Column', true), { name: 'to', label: 'To', type: 'select', options: ['number', 'date', 'text'] }]); return r && { op: 'type', cols: pickCols(r.c), to: r.to }; }],
        ['Split column', 'scissors-line-dashed', async () => { const r = await UI.form('Split Column', [colField('c', 'Column'), { name: 'by', label: 'By (a delimiter, such as , or a space)', value: ',' }]); return r && r.by !== '' && { op: 'split', col: r.c, by: r.by === 'space' ? ' ' : r.by }; }],
        ['Merge columns', 'combine', async () => { const r = await UI.form('Merge Columns', [colField('a', 'First column'), colField('b', 'Then'), { name: 'sep', label: 'With between them', value: ' ' }, { name: 'name', label: 'New column name', value: 'Merged' }]); return r && r.a !== r.b && { op: 'merge', cols: [r.a, r.b], sep: r.sep, name: r.name }; }],
        ['Rename', 'text-cursor-input', async () => { const r = await UI.form('Rename Column', [colField('c', 'Column'), { name: 'to', label: 'New name' }]); return r && r.to.trim() && { op: 'rename', col: r.c, to: r.to.trim() }; }],
        ['Remove columns', 'table-2', async () => { const r = await UI.form('Remove Column', [colField('c', 'Column')]); return r && { op: 'remove', cols: [r.c] }; }],
        ['Fill down', 'arrow-down-to-line', async () => { const r = await UI.form('Fill Down', [colField('c', 'Column', true)]); return r && { op: 'fillDown', cols: pickCols(r.c) }; }],
        ['Filter rows', 'list-filter', async () => { const r = await UI.form('Keep Rows', [colField('c', 'Where column'), { name: 'cond', label: 'Is', type: 'select', options: SheetsClean.CONDITIONS }, { name: 'v', label: 'Value' }]); return r && { op: 'filter', col: r.c, cond: r.cond, value: r.v }; }],
        ['Remove blank rows', 'rows-3', async () => ({ op: 'removeBlankRows' })],
        ['Remove duplicates', 'copy-minus', async () => { const r = await UI.form('Remove Duplicates', [colField('c', 'Compare', true)]); return r && { op: 'dedupe', cols: r.c === '*' ? [] : [r.c] }; }],
        ['Sort', 'arrow-down-wide-narrow', async () => { const r = await UI.form('Sort', [colField('c', 'By'), { name: 'd', label: 'Order', type: 'select', options: [{ value: '', label: 'Ascending (A to Z, small to large)' }, { value: '1', label: 'Descending' }] }]); return r && { op: 'sort', col: r.c, desc: r.d === '1' }; }],
        ['Group by', 'group', async () => { const r = await UI.form('Group By', [colField('by', 'Group by'), { name: 'agg', label: 'Calculate', type: 'select', options: ['count', 'sum', 'average', 'min', 'max', 'distinct'] }, colField('c', 'Of column')]); return r && { op: 'group', by: [r.by], agg: r.agg, col: r.agg === 'count' ? '' : r.c }; }],
        ['Unpivot', 'rotate-3d', async () => { const r = await UI.form('Unpivot Columns', [colField('a', 'From column'), colField('b', 'To column')]); if (!r) return null; const h = head(), i = h.indexOf(r.a), j = h.indexOf(r.b); return i >= 0 && j >= i && { op: 'unpivot', cols: h.slice(i, j + 1) }; }]
      ];
      const tools = q('.ce-tools');
      for (const [label, icon, make] of TOOLS) {
        const b = document.createElement('button');
        b.className = 'btn';
        b.innerHTML = `<span class="ico" data-icon="${icon}"></span>`;
        b.append(label);
        b.addEventListener('click', async () => {
          const st = await make();
          if (!st) return;
          (Array.isArray(st) ? st : [st]).forEach(x => steps.push(x));
          draw();
        });
        tools.appendChild(b);
      }
      draw();
      document.body.appendChild(back);
      const close = () => back.remove();
      q('.cancel').addEventListener('click', close);
      back.addEventListener('keydown', e => { if (e.key === 'Escape' && !document.querySelectorAll('.modal-backdrop')[1]) close(); });
      q('.load').addEventListener('click', () => {
        close();
        ctx.snapshot();
        const dest = ctx.addSheetNamed(ctx.uniqueName(('Cleaned ' + src.name).slice(0, 31)));
        const query = { src: { sheet: src.name, ref: refOfArea(g) }, steps: steps.slice(), r: 0, c: 0, h: 0, w: 0 };
        writeRows(dest, query, result);
        dest.cleans = [query];
        ctx.buildGrid();
        UI.toast(`${result.length - 1} rows loaded. Data → Refresh All runs the steps again.`, 4000);
      });
      q('.ce-tools .btn').focus();
    }
    /** Runs every Clean & Transform query again on its source's current data. */
    function refreshCleans() {
      let n = 0;
      for (const s of book().sheets) for (const qy of s.cleans || []) {
        const src = book().sheets.find(x => x.name === qy.src.sheet);
        const a = src && SheetsCF.area(qy.src.ref);
        if (!a) continue;
        // Rows added right below the data are part of it.
        const rowFilled = r => { for (let c = a.c1; c <= a.c2; c++) if (ctx.rawAt(r, c, src) !== '') return true; return false; };
        while (a.r2 < 1048575 && rowFilled(a.r2 + 1)) a.r2++;
        if (!n) ctx.snapshot();
        qy.src.ref = refOfArea(a);
        writeRows(s, qy, SheetsClean.run(readTable(src, a), qy.steps));
        n++;
      }
      if (n) ctx.changed();
      return n;
    }

    // =====================================================================================
    // Saving and opening
    // =====================================================================================
    /** Formula text as Excel stores it, given where the cell is. */
    function excelFormulaOpts(s, r, c) {
      const names = new Set([...(book().names || []).map(n => n.name.toUpperCase()), ...allTables().map(t => t.name.toUpperCase())]);
      const t = tableAt(r, c, s);
      return { names, table: t ? t.name : '' };
    }
    /** Table parts for one sheet: [{ xml, name }] (ids are numbered across the workbook). */
    let tableId = 0;
    function tablesXml(s, reset) {
      if (reset) tableId = 0;
      return (s.tables || []).map(t => T().toXml(t, ++tableId, headerNames(t, s), c => String(ctx.rawAt(t.r2, c, s))));
    }

    return {
      // tables
      insertTable, tableDesign, convertToRange, openFilterMenu, applyFilters, afterEdit, tableStyle, tableAt, headerNames, currentRegion, setTotals,
      // names
      resolveName, nameBoxGo, defineName, nameManager, shiftNames,
      resolveTable: (spec, sheetName, cur) => T().resolve(book(), spec, sheetName, cur, textAt),
      // what-if
      goalSeek, solver,
      // sparklines and slicers
      insertSparklines, clearSparklines, sparkAt, sparkCells, insertSlicer, renderSlicers,
      // scripts and web data
      scriptsPane, newScript, fromWeb, refreshQueries, cleanData, refreshCleans,
      // files
      excelFormulaOpts, tablesXml
    };
  }

  window.SheetsData = SheetsData;
})();
