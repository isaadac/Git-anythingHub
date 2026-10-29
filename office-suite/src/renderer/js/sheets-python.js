'use strict';
// Python in cells (=PY): runs each cell's code in a worker (sheets-python-worker.js) and keeps the
// results, keyed by the code and the values it read, so a cell is only run again when one of those
// changes. While code runs the cell shows #BUSY!, then the sheet recalculates with the result.
// A workbook opened from a file doesn't run its Python until the person allows it (#BLOCKED!).
(function (root) {
  function SheetsPython(opts) {
    const F = opts.Formula;
    const MAX = 500, LOAD_MS = 60000, RUN_MS = 15000;
    const results = new Map(); // key -> { value } | { error }
    const running = new Map(); // id -> { key, timer }
    let worker = null, loaded = false, seq = 0, blocked = false, onChange = () => {};

    function start() {
      worker = new Worker(opts.url || 'js/sheets-python-worker.js', { type: 'module' });
      loaded = false;
      worker.onmessage = e => {
        const job = running.get(e.data.id);
        if (!job) return;
        clearTimeout(job.timer);
        running.delete(e.data.id);
        loaded = true;
        keep(job.key, e.data.error !== undefined ? { error: e.data.error } : { value: e.data.value });
        if (!running.size) onChange();
      };
      worker.onerror = e => {
        e.preventDefault();
        fail('Python could not start' + (e.message ? ': ' + e.message : ''));
      };
    }
    /** Ends every job with the same error (the worker died or hung) and starts afresh next time. */
    function fail(msg) {
      for (const job of running.values()) { clearTimeout(job.timer); keep(job.key, { error: msg }); }
      running.clear();
      if (worker) worker.terminate();
      worker = null;
      onChange();
    }
    function keep(key, r) {
      results.delete(key);
      results.set(key, r);
      if (results.size > MAX) results.delete(results.keys().next().value);
    }

    /** The PY function's runner (Engine book.python): a result, or #BUSY! while it is worked out. */
    function run(code, input, ctx) {
      if (blocked) throw F.ERR.BLOCKED;
      const key = code + '\u0000' + JSON.stringify(input);
      const r = results.get(key);
      const cur = ctx && ctx.engine && ctx.engine.cur;
      const at = cur ? cur.sheet + '!' + F.addr(cur.row, cur.col) : '';
      if (r) {
        if (r.error !== undefined) { errors.set(at, r.error); throw F.ERR.CALC; }
        errors.delete(at);
        return r.value;
      }
      if ([...running.values()].some(j => j.key === key)) throw F.ERR.BUSY;
      if (!worker) start();
      const id = ++seq;
      const timer = setTimeout(() => fail(loaded ? 'The Python code took too long (more than 15 seconds)' : 'Python took too long to start'), loaded ? RUN_MS : LOAD_MS);
      running.set(id, { key, timer });
      worker.postMessage({ id, code, ...input });
      throw F.ERR.BUSY;
    }
    const errors = new Map(); // "Sheet1!A1" -> the error its code raised

    return {
      run,
      /** The error the code in a cell raised, to show with its #CALC! (null when it ran fine). */
      errorAt: (sheet, row, col) => errors.get(sheet + '!' + F.addr(row, col)) || null,
      get busy() { return running.size > 0; },
      get blocked() { return blocked; },
      set blocked(v) { blocked = !!v; },
      onChange(f) { onChange = f; },
      /** Whether any cell of the book has a Python formula. */
      hasPython: book => book.sheets.some(s => Object.values(s.cells).some(c => typeof c.v === 'string' && c.v[0] === '=' && /\bPY\s*\(/i.test(c.v))),
      stop() { if (worker) worker.terminate(); worker = null; running.clear(); }
    };
  }
  if (typeof module === 'object' && module.exports) module.exports = SheetsPython;
  else root.SheetsPython = SheetsPython;
})(typeof window !== 'undefined' ? window : globalThis);
