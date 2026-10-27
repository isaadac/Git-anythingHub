'use strict';
// Atlas Sheets — pivot tables.
// A pivot is defined on the sheet that shows it:
//   { id, src: { sheet, ref }, r, c, rows: [field], cols: [field], vals: [{ field, agg }],
//     filters: [{ field, value }], extent: { r2, c2 } }
// Fields are the header names in the source's first row. compute() is pure: it turns the
// source table into the grid of cells the sheet shows (values only, styled by kind).
(function (root) {
  const AGGS = ['sum', 'count', 'average', 'max', 'min'];
  const AGG_LABEL = { sum: 'Sum', count: 'Count', average: 'Average', max: 'Max', min: 'Min' };
  const BLANK = '(blank)';
  const num = v => typeof v === 'number' && isFinite(v);
  const label = v => (v === '' || v === null || v === undefined ? BLANK : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : v);
  // Numbers before text, text case-insensitively, (blank) last — as Excel sorts row labels.
  const cmp = (a, b) => {
    if (a === BLANK || b === BLANK) return a === b ? 0 : a === BLANK ? 1 : -1;
    if (num(a) && num(b)) return a - b;
    if (num(a)) return -1;
    if (num(b)) return 1;
    return String(a).localeCompare(String(b), undefined, { sensitivity: 'base', numeric: true });
  };
  const cmpKeys = (x, y) => { for (let i = 0; i < x.length; i++) { const d = cmp(x[i], y[i]); if (d) return d; } return 0; };
  const valName = v => `${AGG_LABEL[v.agg]} of ${v.field}`;

  function acc() { return { n: 0, nums: 0, sum: 0, min: Infinity, max: -Infinity }; }
  function add(a, v) {
    if (v === '' || v === null || v === undefined) return;
    a.n++;
    if (num(v)) { a.nums++; a.sum += v; if (v < a.min) a.min = v; if (v > a.max) a.max = v; }
  }
  function result(a, agg) {
    if (!a) return '';
    switch (agg) {
      case 'count': return a.n;
      case 'sum': return a.n ? a.sum : '';
      case 'average': return a.nums ? a.sum / a.nums : '';
      case 'max': return a.nums ? a.max : '';
      case 'min': return a.nums ? a.min : '';
    }
    return '';
  }

  /**
   * table: { headers: [name], rows: [[value]] } with computed values (number | string | boolean | '').
   * Returns { cells: [[{ v, k }]], error } where k is 'hdr' | 'label' | 'sub' | 'total' | 'val'.
   */
  function compute(table, def) {
    const idx = name => table.headers.findIndex(h => String(h) === String(name));
    const rows = (def.rows || []).map(idx), cols = (def.cols || []).map(idx);
    const vals = (def.vals || []).map(v => ({ ...v, i: idx(v.field) }));
    const filters = (def.filters || []).map(f => ({ ...f, i: idx(f.field) }));
    if ([...rows, ...cols, ...vals.map(v => v.i), ...filters.map(f => f.i)].some(i => i < 0)) return { cells: [], error: 'A field in this PivotTable is no longer in the source data. Refresh after fixing the headers, or change the fields.' };
    const recs = table.rows.filter(r => r.some(v => v !== '' && v !== null && v !== undefined))
      .filter(r => filters.every(f => f.value === null || f.value === undefined || String(label(r[f.i])) === String(f.value)));
    const R = rows.length, V = vals.length;
    // Accumulators per (row prefix, column key or total).
    const accs = new Map();
    const rowKeys = new Map(), colKeys = new Map();
    const K = a => JSON.stringify(a);
    for (const r of recs) {
      const rk = rows.map(i => label(r[i])), ck = cols.map(i => label(r[i]));
      rowKeys.set(K(rk), rk);
      if (cols.length) colKeys.set(K(ck), ck);
      for (let p = 0; p <= R; p++) {
        for (const c of cols.length ? [K(ck), '*'] : ['*']) {
          const key = K(rk.slice(0, p)) + '|' + c;
          let list = accs.get(key);
          if (!list) { list = vals.map(acc); accs.set(key, list); }
          vals.forEach((v, j) => add(list[j], r[v.i]));
        }
      }
    }
    const rk = [...rowKeys.values()].sort(cmpKeys), ck = [...colKeys.values()].sort(cmpKeys);
    const colGroups = cols.length ? ck.map(k => ({ key: K(k), text: k.map(String).join(' · ') })) : [];
    const valueCols = [];
    if (cols.length) {
      colGroups.forEach(g => vals.forEach((v, j) => valueCols.push({ c: g.key, j, head: V > 1 ? `${g.text} · ${valName(v)}` : g.text })));
      vals.forEach((v, j) => valueCols.push({ c: '*', j, head: V > 1 ? `Total ${valName(v)}` : 'Grand Total', total: true }));
    } else vals.forEach((v, j) => valueCols.push({ c: '*', j, head: valName(v) }));
    const cells = [];
    const valueRow = (prefix, k) => valueCols.map(vc => {
      const list = accs.get(K(prefix) + '|' + vc.c);
      return { v: result(list && list[vc.j], vals[vc.j].agg), k: vc.total && k === 'val' ? 'sub' : k };
    });
    const pad = R || 1;
    if (cols.length && V === 1) cells.push(Array(pad + valueCols.length).fill(0).map((_, i) => ({ v: i === 0 ? valName(vals[0]) : i === pad ? 'Column Labels' : '', k: 'hdr' })));
    const head = R ? rows.map(i => ({ v: String(table.headers[i]), k: 'hdr' })) : [{ v: '', k: 'hdr' }];
    cells.push([...head, ...valueCols.map(vc => ({ v: vc.head, k: 'hdr' }))]);
    if (!V) {
      rk.forEach(key => cells.push(key.map(v => ({ v, k: 'label' }))));
      return { cells };
    }
    // Body: rows in order; an outer label shows on its group's first row, and each outer group ends with a subtotal.
    let prev = null;
    // With no row fields there is only the grand total row.
    (R ? rk : []).forEach((key, n) => {
      const row = key.map((v, i) => ({ v: prev && prev.slice(0, i + 1).every((x, j) => cmp(x, key[j]) === 0) ? '' : v, k: 'label' }));
      cells.push([...row, ...valueRow(key, 'val')]);
      prev = key;
      const next = rk[n + 1];
      for (let p = R - 1; p >= 1; p--) {
        if (next && next.slice(0, p).every((x, j) => cmp(x, key[j]) === 0)) break;
        const lab = Array(pad).fill(0).map((_, i) => ({ v: i === p - 1 ? `${key[p - 1]} Total` : '', k: 'sub' }));
        cells.push([...lab, ...valueRow(key.slice(0, p), 'sub')]);
      }
    });
    cells.push([{ v: 'Grand Total', k: 'total' }, ...Array(pad - 1).fill(0).map(() => ({ v: '', k: 'total' })), ...valueRow([], 'total')]);
    return { cells };
  }

  /** A definition from a file or the user, with only known fields; null if unusable. */
  function clean(p) {
    const REF = /^[A-Z]{1,3}[1-9]\d{0,6}:[A-Z]{1,3}[1-9]\d{0,6}$/;
    if (!p || typeof p !== 'object' || !p.src || typeof p.src.sheet !== 'string' || typeof p.src.ref !== 'string' || !REF.test(p.src.ref)) return null;
    const int = (v, max) => (Number.isInteger(v) && v >= 0 && v < max ? v : null);
    const r = int(p.r, 1048576), c = int(p.c, 16384);
    if (r === null || c === null) return null;
    const field = f => (typeof f === 'string' && f.length <= 255 ? f : null);
    const list = (a, fn) => (Array.isArray(a) ? a.slice(0, 50).map(fn).filter(x => x !== null) : []);
    const out = {
      id: typeof p.id === 'string' && /^[\w-]{1,40}$/.test(p.id) ? p.id : 'pv' + Math.random().toString(36).slice(2, 10),
      name: typeof p.name === 'string' ? p.name.slice(0, 100) : 'PivotTable',
      src: { sheet: p.src.sheet.slice(0, 31), ref: p.src.ref }, r, c,
      rows: list(p.rows, field), cols: list(p.cols, field),
      vals: list(p.vals, v => (v && field(v.field) !== null && AGGS.includes(v.agg) ? { field: v.field, agg: v.agg } : null)),
      filters: list(p.filters, f => (f && field(f.field) !== null ? { field: f.field, value: typeof f.value === 'string' || typeof f.value === 'number' ? f.value : null } : null))
    };
    if (p.extent && int(p.extent.r2, 1048576) !== null && int(p.extent.c2, 16384) !== null) out.extent = { r2: p.extent.r2, c2: p.extent.c2 };
    return out;
  }

  const api = { compute, clean, AGGS, AGG_LABEL, valName, label, BLANK };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.SheetsPivot = api;
})(typeof window !== 'undefined' ? window : globalThis);
