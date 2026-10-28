'use strict';
// Runs a workbook script (Automate) away from the page. The script works on a copy of the
// workbook with the same formula engine, so it reads recalculated values after it writes;
// what it changes is sent back as a list of edits for the app to apply (one undo step).
// The script gets no network, storage or page access.
self.window = self;
importScripts('formula.js');
const F = self.Formula;
// Nothing the script could use to reach outside.
for (const k of ['fetch', 'XMLHttpRequest', 'WebSocket', 'EventSource', 'importScripts', 'indexedDB', 'caches', 'Worker', 'SharedWorker', 'BroadcastChannel', 'navigator', 'location']) {
  try { Object.defineProperty(self, k, { value: undefined, configurable: false, writable: false }); } catch { /* not configurable */ }
}

self.onmessage = e => {
  const { code, book, active, selection } = e.data;
  const sheets = book.sheets.map(s => ({ name: s.name, cells: new Map(Object.entries(s.cells)) }));
  const ops = [], log = [];
  const byName = n => sheets.find(s => s.name.toLowerCase() === String(n).toLowerCase());
  const engine = new F.Engine({
    getRaw: (name, r, c) => { const s = byName(name); return s ? s.cells.get(F.addr(r, c)) ?? '' : ''; },
    sheetNames: () => sheets.map(s => s.name),
    resolveName: n => (book.names || []).find(x => x.name.toUpperCase() === n)?.ref ?? null
  });
  const out = v => (v instanceof F.FormulaError ? v.error : v);
  const parseAddr = (s, a) => {
    const [x, y] = String(a).replace(/\$/g, '').toUpperCase().split(':');
    const p = F.parseRef(x), q = y ? F.parseRef(y) : p;
    if (!p || !q) throw new Error(`"${a}" is not a cell or range address`);
    return { s, r1: Math.min(p.row, q.row), c1: Math.min(p.col, q.col), r2: Math.max(p.row, q.row), c2: Math.max(p.col, q.col) };
  };
  const write = (s, r, c, raw) => {
    raw = raw === null || raw === undefined ? '' : typeof raw === 'boolean' ? (raw ? 'TRUE' : 'FALSE') : String(raw);
    if (raw === '') s.cells.delete(F.addr(r, c)); else s.cells.set(F.addr(r, c), raw);
    ops.push({ t: 'set', sheet: s.name, r, c, v: raw });
    engine.invalidate();
  };
  function Range(g) {
    const each = fn => { for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) fn(r, c); };
    const grid = fn => Array.from({ length: g.r2 - g.r1 + 1 }, (_, i) => Array.from({ length: g.c2 - g.c1 + 1 }, (_, j) => fn(g.r1 + i, g.c1 + j)));
    const style = props => each((r, c) => ops.push({ t: 'style', sheet: g.s.name, r, c, s: props }));
    const range = {
      getAddress: () => F.addr(g.r1, g.c1) + (g.r1 === g.r2 && g.c1 === g.c2 ? '' : ':' + F.addr(g.r2, g.c2)),
      getRowCount: () => g.r2 - g.r1 + 1,
      getColumnCount: () => g.c2 - g.c1 + 1,
      getValue: () => out(engine.value(g.s.name, g.r1, g.c1)),
      getValues: () => grid((r, c) => out(engine.value(g.s.name, r, c))),
      getFormula: () => g.s.cells.get(F.addr(g.r1, g.c1)) ?? '',
      getFormulas: () => grid((r, c) => g.s.cells.get(F.addr(r, c)) ?? ''),
      setValue: v => { each((r, c) => write(g.s, r, c, typeof v === 'string' && v[0] === '=' ? "'" + v : v)); return range; },
      setFormula: f => { each((r, c) => write(g.s, r, c, String(f).startsWith('=') ? F.shiftFormula(String(f), r - g.r1, c - g.c1) : f)); return range; },
      setValues: rows => {
        if (!Array.isArray(rows) || !rows.every(Array.isArray)) throw new Error('setValues needs rows: [[...], [...]]');
        rows.forEach((row, i) => row.forEach((v, j) => write(g.s, g.r1 + i, g.c1 + j, v)));
        return range;
      },
      setFormulas: rows => { rows.forEach((row, i) => row.forEach((v, j) => write(g.s, g.r1 + i, g.c1 + j, v))); return range; },
      clear: () => { each((r, c) => write(g.s, r, c, '')); return range; },
      getCell: (r, c) => Range({ s: g.s, r1: g.r1 + r, c1: g.c1 + c, r2: g.r1 + r, c2: g.c1 + c }),
      getOffsetRange: (dr, dc) => Range({ s: g.s, r1: g.r1 + dr, c1: g.c1 + dc, r2: g.r2 + dr, c2: g.c2 + dc }),
      getResizedRange: (dr, dc) => Range({ ...g, r2: g.r2 + dr, c2: g.c2 + dc }),
      getFormat: () => ({
        setBold: b => style({ b: !!b }), setItalic: b => style({ i: !!b }),
        setFill: color => style({ bg: String(color) }), setFontColor: color => style({ color: String(color) }),
        setNumberFormat: f => style({ fmt: String(f) }), setHorizontalAlignment: a => style({ align: String(a).toLowerCase() })
      }),
      getWorksheet: () => Sheet(g.s)
    };
    return range;
  }
  function Sheet(s) {
    return {
      getName: () => s.name,
      getRange: a => Range(parseAddr(s, a)),
      getCell: (r, c) => Range({ s, r1: r, c1: c, r2: r, c2: c }),
      getUsedRange: () => {
        let r2 = 0, c2 = 0;
        for (const k of s.cells.keys()) { const p = F.parseRef(k); if (p) { r2 = Math.max(r2, p.row); c2 = Math.max(c2, p.col); } }
        return Range({ s, r1: 0, c1: 0, r2, c2 });
      }
    };
  }
  const workbook = {
    getActiveWorksheet: () => Sheet(byName(active)),
    getWorksheet: n => { const s = byName(n); return s ? Sheet(s) : undefined; },
    getWorksheets: () => sheets.map(Sheet),
    addWorksheet: n => {
      const name = String(n || 'Sheet' + (sheets.length + 1)).slice(0, 31);
      if (byName(name)) throw new Error(`There is already a sheet called ${name}`);
      const s = { name, cells: new Map() };
      sheets.push(s);
      ops.push({ t: 'addSheet', name });
      return Sheet(s);
    },
    getSelectedRange: () => Range(parseAddr(byName(active), selection || 'A1'))
  };
  const consoleLike = { log: (...a) => log.push(a.map(x => (typeof x === 'string' ? x : JSON.stringify(x))).join(' ')) };
  consoleLike.info = consoleLike.warn = consoleLike.error = consoleLike.log;
  try {
    // A script is either the body of main(workbook) or defines function main(workbook) { … }.
    const run = new Function('workbook', 'console', `"use strict";\n${code}\n;if (typeof main === 'function') return main(workbook);`);
    const result = run(workbook, consoleLike);
    if (result !== undefined) log.push('→ ' + (typeof result === 'string' ? result : JSON.stringify(result)));
    self.postMessage({ ok: true, ops: ops.slice(0, 1e6), log: log.slice(0, 1000) });
  } catch (err) {
    self.postMessage({ ok: false, error: String((err && err.message) || err), line: ((err && err.stack) || '').match(/<anonymous>:(\d+)/)?.[1], ops: [], log });
  }
};
