'use strict';
// Isadek Sheets — tables (Excel's "Format as Table" / ListObjects), named ranges and slicers.
//
// A table on a sheet: { name, r1, c1, r2, c2, totals, banded, style, filter: { colIndex: [values shown] } }
// Row r1 is the header row (the column names); with `totals` the last row is the total row.
// Formulas refer to tables by name: Table1[Sales], Table1[#All], [@Sales] (this row) …
// A workbook name: { name, ref } where ref is what it refers to, e.g. "Sheet1!$B$2:$B$9" or "0.07".
// A slicer: { table, col, x, y, w, h } — buttons that filter one table column.
(function () {
  const STYLES = {
    blue: { header: '#4472c4', band: '#d9e1f2', excel: 'TableStyleMedium2' },
    orange: { header: '#ed7d31', band: '#fbe2d5', excel: 'TableStyleMedium3' },
    gray: { header: '#7f7f7f', band: '#ededed', excel: 'TableStyleMedium4' },
    gold: { header: '#bf8f00', band: '#fff2cc', excel: 'TableStyleMedium5' },
    green: { header: '#70ad47', band: '#e2efda', excel: 'TableStyleMedium7' }
  };
  const EXCEL_STYLE = Object.fromEntries(Object.entries(STYLES).map(([k, v]) => [v.excel, k]));
  const NAME_RE = /^[A-Za-z_\\][\w.]{0,254}$/;
  const isCellLike = n => /^[A-Za-z]{1,3}\d+$/.test(n) || /^R\d*C\d*$/i.test(n);
  const validName = n => typeof n === 'string' && NAME_RE.test(n) && !isCellLike(n) && !/^(TRUE|FALSE)$/i.test(n);

  function cleanTable(t) {
    if (!t || typeof t !== 'object' || !validName(t.name)) return null;
    const n = k => Math.max(0, Math.min(k.startsWith('r') ? 1048575 : 16383, Math.floor(+t[k] || 0)));
    const out = { name: t.name, r1: n('r1'), c1: n('c1'), r2: n('r2'), c2: n('c2'), totals: t.totals === true, banded: t.banded !== false, style: STYLES[t.style] ? t.style : 'blue', filter: {} };
    if (out.r2 <= out.r1 || out.c2 < out.c1) return null;
    if (t.filter && typeof t.filter === 'object') {
      for (const [k, v] of Object.entries(t.filter)) if (/^\d{1,5}$/.test(k) && Array.isArray(v)) out.filter[k] = v.slice(0, 10000).map(String);
    }
    return out;
  }
  function cleanName(x) {
    if (!x || typeof x !== 'object' || !validName(x.name) || typeof x.ref !== 'string') return null;
    const ref = x.ref.trim().replace(/^=/, '').slice(0, 2000);
    return ref ? { name: x.name, ref } : null;
  }
  function cleanSlicer(x) {
    if (!x || typeof x !== 'object' || !validName(x.table) || typeof x.col !== 'string') return null;
    const n = (v, lo, hi, d) => (Number.isFinite(+v) ? Math.max(lo, Math.min(hi, Math.round(+v))) : d);
    return { table: x.table, col: x.col.slice(0, 255), x: n(x.x, 0, 1e6, 100), y: n(x.y, 0, 1e7, 60), w: n(x.w, 120, 800, 180), h: n(x.h, 80, 1200, 220) };
  }

  /** Column names: the header cells' text, made unique and never empty (as Excel does). */
  function headerNames(t, textAt) {
    const seen = new Set();
    const out = [];
    for (let c = t.c1; c <= t.c2; c++) {
      let n = String(textAt(t.r1, c) ?? '').trim() || 'Column' + (c - t.c1 + 1);
      let k = n, i = 2;
      while (seen.has(k.toLowerCase())) k = n + i++;
      seen.add(k.toLowerCase());
      out.push(k);
    }
    return out;
  }

  /**
   * The cells a structured reference covers: { sheet, r1, c1, r2, c2 } or null (#REF!).
   * book: { sheets: [{ name, tables }] }; spec from Formula.parseStructured; cur: the cell computing it.
   */
  function resolve(book, spec, sheetName, cur, textAt) {
    let t = null, s = null;
    if (spec.table) {
      for (const sh of book.sheets) { const f = (sh.tables || []).find(x => x.name.toLowerCase() === spec.table.toLowerCase()); if (f) { t = f; s = sh; break; } }
    } else if (cur) {
      s = book.sheets.find(x => x.name === cur.sheet) || null;
      t = s && (s.tables || []).find(x => cur.row >= x.r1 && cur.row <= x.r2 && cur.col >= x.c1 && cur.col <= x.c2);
    }
    if (!t) return null;
    let c1 = t.c1, c2 = t.c2;
    if (spec.cols) {
      const names = headerNames(t, (r, c) => textAt(s, r, c)).map(n => n.toLowerCase());
      const a = names.indexOf(spec.cols[0].toLowerCase()), b = names.indexOf(spec.cols[1].toLowerCase());
      if (a < 0 || b < 0) return null;
      c1 = t.c1 + Math.min(a, b); c2 = t.c1 + Math.max(a, b);
    }
    const dataStart = t.r1 + 1, dataEnd = t.totals ? t.r2 - 1 : t.r2;
    let r1, r2;
    if (spec.thisRow) {
      if (!cur || cur.sheet !== s.name || cur.row < dataStart || cur.row > dataEnd) return null;
      r1 = r2 = cur.row;
    } else {
      const p = spec.parts;
      if (p.has('#all')) { r1 = t.r1; r2 = t.r2; }
      else {
        const hd = p.has('#headers'), tot = p.has('#totals'), data = p.has('#data') || (!hd && !tot);
        if (tot && !t.totals) return null;
        r1 = hd ? t.r1 : data ? dataStart : t.r2;
        r2 = tot ? t.r2 : data ? dataEnd : t.r1;
      }
    }
    if (r2 < r1) return null;
    return { sheet: s.name, r1, c1, r2, c2 };
  }

  /** How a table draws a cell (under the cell's own formatting), or null outside tables. */
  function styleAt(tables, r, c) {
    for (const t of tables || []) {
      if (r < t.r1 || r > t.r2 || c < t.c1 || c > t.c2) continue;
      const st = STYLES[t.style] || STYLES.blue;
      if (r === t.r1) return { b: true, color: '#ffffff', bg: st.header, header: true, table: t };
      if (t.totals && r === t.r2) return { b: true, borderTop: st.header, table: t };
      return { bg: t.banded && (r - t.r1) % 2 === 1 ? st.band : '', table: t };
    }
    return null;
  }
  const tableAt = (tables, r, c) => (tables || []).find(t => r >= t.r1 && r <= t.r2 && c >= t.c1 && c <= t.c2) || null;

  /** Rows the table's filters hide. valueText(r, c) gives a cell's shown text. */
  function hiddenRows(t, valueText) {
    const out = [];
    const cols = Object.entries(t.filter || {});
    if (!cols.length) return out;
    const end = t.totals ? t.r2 - 1 : t.r2;
    for (let r = t.r1 + 1; r <= end; r++) {
      for (const [k, allowed] of cols) {
        if (!allowed.includes(valueText(r, t.c1 + +k))) { out.push(r); break; }
      }
    }
    return out;
  }

  /** Moves tables, names' references and slicers when rows or columns are inserted or deleted. */
  function shift(tables, axis, index, count) {
    const k1 = axis === 'row' ? 'r1' : 'c1', k2 = axis === 'row' ? 'r2' : 'c2';
    return (tables || []).map(t => {
      const n = { ...t };
      if (count < 0 && t[k1] >= index && t[k2] < index - count) return null; // deleted whole
      if (t[k1] >= index) n[k1] = Math.max(index, t[k1] + count);
      if (t[k2] >= index) n[k2] = Math.max(n[k1] + (axis === 'row' ? 1 : 0), t[k2] + count);
      return n;
    }).filter(Boolean);
  }

  // ---------------- Excel ----------------
  const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const A1 = (r, c) => window.Formula.addr(r, c);
  const TOTAL_FN = { SUM: 'sum', AVERAGE: 'average', COUNT: 'count', COUNTA: 'countNums', MAX: 'max', MIN: 'min' };
  /** The table part (xl/tables/tableN.xml). names: the column names; totalsAt(c): the total row's formula there. */
  function toXml(t, id, names, totalsAt) {
    const ref = `${A1(t.r1, t.c1)}:${A1(t.r2, t.c2)}`;
    const filterRef = `${A1(t.r1, t.c1)}:${A1(t.totals ? t.r2 - 1 : t.r2, t.c2)}`;
    const filters = Object.entries(t.filter || {}).map(([k, vals]) => `<filterColumn colId="${+k}"><filters${vals.includes('') ? ' blank="1"' : ''}>${vals.filter(v => v !== '').map(v => `<filter val="${esc(v)}"/>`).join('')}</filters></filterColumn>`).join('');
    const cols = names.map((n, i) => {
      let extra = '';
      if (t.totals) {
        const f = totalsAt ? totalsAt(t.c1 + i) : '';
        const m = /^=SUBTOTAL\(1(0\d),/i.exec(f || '');
        const fn = m ? { '01': 'average', '02': 'count', '03': 'countNums', '04': 'max', '05': 'min', '09': 'sum' }[m[1]] : null;
        if (fn) extra = ` totalsRowFunction="${fn}"`;
        else if (i === 0) extra = ' totalsRowLabel="Total"';
      }
      return `<tableColumn id="${i + 1}" name="${esc(n)}"${extra}/>`;
    }).join('');
    const st = STYLES[t.style] || STYLES.blue;
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="${id}" name="${esc(t.name)}" displayName="${esc(t.name)}" ref="${ref}"${t.totals ? ' totalsRowCount="1"' : ''}>` +
      `<autoFilter ref="${filterRef}">${filters}</autoFilter><tableColumns count="${names.length}">${cols}</tableColumns>` +
      `<tableStyleInfo name="${st.excel}" showFirstColumn="0" showLastColumn="0" showRowStripes="${t.banded ? 1 : 0}" showColumnStripes="0"/></table>`;
  }
  /** A table part from a file. */
  function fromXml(xml) {
    const d = new DOMParser().parseFromString(xml, 'application/xml');
    const el = d.documentElement;
    if (!el || el.localName !== 'table') return null;
    const ref = (el.getAttribute('ref') || '').split(':');
    const p1 = window.Formula.parseRef(ref[0] || ''), p2 = window.Formula.parseRef(ref[1] || ref[0] || '');
    if (!p1 || !p2) return null;
    if (el.getAttribute('headerRowCount') === '0') return null; // tables without a header row aren't supported
    const si = el.getElementsByTagName('tableStyleInfo')[0];
    const t = { name: el.getAttribute('displayName') || el.getAttribute('name'), r1: p1.row, c1: p1.col, r2: p2.row, c2: p2.col,
      totals: +(el.getAttribute('totalsRowCount') || 0) > 0,
      banded: !si || si.getAttribute('showRowStripes') !== '0', style: (si && EXCEL_STYLE[si.getAttribute('name')]) || 'blue', filter: {} };
    for (const fc of el.getElementsByTagName('filterColumn')) {
      const fl = fc.getElementsByTagName('filters')[0];
      const vals = [...fc.getElementsByTagName('filter')].map(f => f.getAttribute('val') || '');
      if (fl && fl.getAttribute('blank') === '1') vals.push('');
      if (vals.length) t.filter[fc.getAttribute('colId')] = vals;
    }
    return cleanTable(t);
  }

  window.SheetsTables = { STYLES, TOTAL_FN, cleanTable, cleanName, cleanSlicer, validName, headerNames, resolve, styleAt, tableAt, hiddenRows, shift, toXml, fromXml };
})();
