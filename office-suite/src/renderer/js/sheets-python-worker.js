// Runs the Python of =PY(...) cells, away from the page, with Pyodide (CPython built for the web).
// It is loaded once, on the first Python cell; it can read only the values the cell hands it, and
// has no network, storage or page access.
//
// Pyodide's files (all bundled, nothing is downloaded): node_modules/pyodide/pyodide.mjs,
// node_modules/pyodide/pyodide.asm.mjs, node_modules/pyodide/pyodide.asm.wasm,
// node_modules/pyodide/python_stdlib.zip, node_modules/pyodide/pyodide-lock.json
import { loadPyodide } from '../../../node_modules/pyodide/pyodide.mjs';

const ready = (async () => {
  const py = await loadPyodide({ indexURL: new URL('../../../node_modules/pyodide/', import.meta.url).href, fullStdLib: false });
  // Now nothing the code could use to reach outside: not on the worker's global object, nor on
  // the prototypes behind it (Python can reach those through js).
  const OUTSIDE = ['fetch', 'XMLHttpRequest', 'WebSocket', 'WebSocketStream', 'WebTransport', 'EventSource', 'importScripts', 'indexedDB', 'caches', 'Worker', 'SharedWorker', 'BroadcastChannel', 'Request', 'navigator'];
  for (let o = self; o && o !== Object.prototype; o = Object.getPrototypeOf(o)) {
    for (const k of OUTSIDE) {
      if (o !== self && !Object.prototype.hasOwnProperty.call(o, k)) continue;
      try { Object.defineProperty(o, k, { value: undefined, configurable: false, writable: false }); } catch { try { delete o[k]; } catch { /* nothing more to do */ } }
    }
  }
  py.runPython(`
import sys, js
for m in ('js', 'pyodide.http', 'pyodide_js', 'micropip'):
    sys.modules[m] = None
del js
def _isadek_xl(ref, _refs):
    if ref not in _refs:
        raise NameError(f'xl("{ref}") is not a cell or range')
    return _refs[ref]
`);
  return py;
})();

/** A Python value as cells: a number, text or truth value, or rows of them. */
function toCells(py, v) {
  if (v === undefined || v === null) return null;
  if (typeof v === 'number' || typeof v === 'string' || typeof v === 'boolean') return v;
  if (typeof v === 'bigint') return Number(v);
  if (v instanceof Map) return [...v.entries()].map(([k, x]) => [toCells(py, k), toCells(py, x)]);
  if (Array.isArray(v)) return v.map(x => (Array.isArray(x) ? x.map(y => toCells(py, y)) : toCells(py, x)));
  return String(v);
}

self.onmessage = async e => {
  const { id, code, refs, args } = e.data;
  try {
    const py = await ready;
    const ns = py.globals.get('dict')();
    const toPy = v => py.toPy(v);
    ns.set('_refs', toPy(refs));
    ns.set('args', toPy(args));
    ns.set('_isadek_xl', py.globals.get('_isadek_xl'));
    py.runPython('xl = lambda ref: _isadek_xl(ref, _refs)', { globals: ns });
    let out = py.runPython(code, { globals: ns });
    // A pandas-like value, a set or a tuple comes back as a list; anything else as its text.
    let v = out;
    if (out && typeof out === 'object' && typeof out.toJs === 'function') {
      const type = out.type;
      v = ['list', 'tuple', 'dict', 'set', 'range', 'frozenset'].includes(type) ? out.toJs({ dict_converter: m => new Map(m), depth: 2 }) : String(out);
      if (v instanceof Set) v = [...v];
      out.destroy();
    }
    ns.destroy();
    self.postMessage({ id, value: toCells(py, v) });
  } catch (err) {
    const msg = String(err && err.message || err).trim().split('\n').filter(Boolean);
    self.postMessage({ id, error: msg[msg.length - 1] || 'Python error' });
  }
};
