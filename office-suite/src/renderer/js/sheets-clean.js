'use strict';
// Isadek Sheets — Clean & Transform: a table (a header row, then rows) put through a list of
// steps, as in Excel's Power Query. Each step is plain data ({ op, ... }), so a query is saved with
// the workbook and Refresh runs it again on the source's current data. Pure functions only.
(function (root) {
  const MAX_ROWS = 1000000;
  const isBlank = v => v === '' || v === null || v === undefined;
  const text = v => (isBlank(v) ? '' : String(v));
  const num = v => { if (typeof v === 'number') return v; const t = text(v).trim().replace(/,/g, ''); return t !== '' && /^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?%?$/i.test(t) ? (t.endsWith('%') ? parseFloat(t) / 100 : parseFloat(t)) : NaN; };
  const proper = s => s.toLowerCase().replace(/(^|[^\p{L}'])(\p{L})/gu, (m, a, b) => a + b.toUpperCase());

  /** Column indexes from names (or indexes), in the current header. */
  function cols(head, list) {
    return (Array.isArray(list) ? list : [list]).map(c => (typeof c === 'number' ? c : head.indexOf(c))).filter(i => i >= 0 && i < head.length);
  }
  const mapCols = (t, list, f) => { const ix = cols(t.head, list); return { head: t.head, rows: t.rows.map(r => { const o = r.slice(); ix.forEach(i => { o[i] = f(o[i]); }); return o; }) }; };
  const uniqueName = (head, name) => { let n = name || 'Column', k = 2; while (head.includes(n)) n = `${name} ${k++}`; return n; };

  function compare(v, op, x) {
    const a = num(v), b = num(x), bothNum = !isNaN(a) && !isNaN(b);
    const s = text(v).toLowerCase(), t = text(x).toLowerCase();
    switch (op) {
      case 'equals': return bothNum ? a === b : s === t;
      case 'not equals': return bothNum ? a !== b : s !== t;
      case 'contains': return s.includes(t);
      case 'does not contain': return !s.includes(t);
      case 'begins with': return s.startsWith(t);
      case 'ends with': return s.endsWith(t);
      case 'greater than': return bothNum ? a > b : s > t;
      case 'less than': return bothNum ? a < b : s < t;
      case 'is blank': return isBlank(v) || text(v).trim() === '';
      case 'is not blank': return !(isBlank(v) || text(v).trim() === '');
      default: return true;
    }
  }

  const STEPS = {
    // Text
    trim: (t, s) => mapCols(t, s.cols, v => (typeof v === 'string' ? v.replace(/\s+/g, ' ').trim() : v)),
    clean: (t, s) => mapCols(t, s.cols, v => (typeof v === 'string' ? v.replace(/[\u0000-\u001f\u007f-\u009f​﻿]/g, '') : v)),
    case: (t, s) => mapCols(t, s.cols, v => (typeof v !== 'string' ? v : s.to === 'upper' ? v.toUpperCase() : s.to === 'lower' ? v.toLowerCase() : proper(v))),
    replace: (t, s) => mapCols(t, s.cols, v => {
      if (s.whole) return text(v) === text(s.find) ? s.with : v;
      if (typeof v !== 'string' || !s.find) return v;
      return v.split(s.find).join(s.with || '');
    }),
    prefix: (t, s) => mapCols(t, s.cols, v => (isBlank(v) ? v : text(s.before) + text(v) + text(s.after))),
    // Columns
    remove: (t, s) => { const ix = new Set(cols(t.head, s.cols)); return { head: t.head.filter((_, i) => !ix.has(i)), rows: t.rows.map(r => r.filter((_, i) => !ix.has(i))) }; },
    keep: (t, s) => { const ix = cols(t.head, s.cols); return { head: ix.map(i => t.head[i]), rows: t.rows.map(r => ix.map(i => r[i])) }; },
    rename: (t, s) => { const [i] = cols(t.head, s.col); if (i === undefined) return t; const head = t.head.slice(); head[i] = uniqueName(head.filter((_, j) => j !== i), text(s.to) || head[i]); return { head, rows: t.rows }; },
    split: (t, s) => {
      const [i] = cols(t.head, s.col);
      if (i === undefined || !s.by) return t;
      const parts = t.rows.map(r => text(r[i]).split(s.by));
      const n = Math.min(50, Math.max(1, ...parts.map(p => p.length)));
      const names = [];
      for (let k = 0; k < n; k++) names.push(uniqueName([...t.head, ...names], `${t.head[i]}.${k + 1}`));
      return {
        head: [...t.head.slice(0, i), ...names, ...t.head.slice(i + 1)],
        rows: t.rows.map((r, ri) => [...r.slice(0, i), ...Array.from({ length: n }, (_, k) => { const p = parts[ri][k]; return p === undefined ? '' : (s.trim === false ? p : p.trim()); }).map(p => (isNaN(num(p)) || p === '' ? p : num(p))), ...r.slice(i + 1)])
      };
    },
    merge: (t, s) => {
      const ix = cols(t.head, s.cols);
      if (ix.length < 2) return t;
      const name = uniqueName(t.head.filter((_, i) => !ix.includes(i)), text(s.name) || t.head[ix[0]]);
      const at = ix[0];
      const keepIx = t.head.map((_, i) => i).filter(i => !ix.includes(i) || i === at);
      return {
        head: keepIx.map(i => (i === at ? name : t.head[i])),
        rows: t.rows.map(r => keepIx.map(i => (i === at ? ix.map(j => text(r[j])).filter(x => x !== '' || s.keepBlanks).join(s.sep ?? ' ') : r[i])))
      };
    },
    type: (t, s) => mapCols(t, s.cols, v => {
      if (s.to === 'text') return text(v);
      if (s.to === 'number') { const n = num(v); return isNaN(n) ? (isBlank(v) ? '' : v) : n; }
      if (s.to === 'date') {
        if (typeof v === 'number') return v;
        const d = new Date(text(v));
        return isNaN(d) ? v : Math.round((Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) - Date.UTC(1899, 11, 30)) / 86400000);
      }
      return v;
    }),
    fillDown: (t, s) => { const ix = cols(t.head, s.cols), last = {}; return { head: t.head, rows: t.rows.map(r => { const o = r.slice(); ix.forEach(i => { if (isBlank(o[i]) || text(o[i]).trim() === '') o[i] = last[i] ?? o[i]; else last[i] = o[i]; }); return o; }) }; },
    // Rows
    filter: (t, s) => { const [i] = cols(t.head, s.col); if (i === undefined) return t; return { head: t.head, rows: t.rows.filter(r => compare(r[i], s.cond, s.value)) }; },
    removeBlankRows: t => ({ head: t.head, rows: t.rows.filter(r => r.some(v => !(isBlank(v) || text(v).trim() === ''))) }),
    dedupe: (t, s) => {
      const ix = s.cols && s.cols.length ? cols(t.head, s.cols) : t.head.map((_, i) => i);
      const seen = new Set();
      return { head: t.head, rows: t.rows.filter(r => { const k = JSON.stringify(ix.map(i => text(r[i]).trim().toLowerCase())); if (seen.has(k)) return false; seen.add(k); return true; }) };
    },
    sort: (t, s) => {
      const [i] = cols(t.head, s.col);
      if (i === undefined) return t;
      const dir = s.desc ? -1 : 1;
      return { head: t.head, rows: t.rows.map((r, k) => [r, k]).sort(([a, ka], [b, kb]) => {
        const x = a[i], y = b[i];
        if (isBlank(x) !== isBlank(y)) return isBlank(x) ? 1 : -1;
        const nx = num(x), ny = num(y);
        const d = !isNaN(nx) && !isNaN(ny) ? nx - ny : text(x).localeCompare(text(y), undefined, { numeric: true, sensitivity: 'base' });
        return d * dir || ka - kb;
      }).map(([r]) => r) };
    },
    top: (t, s) => ({ head: t.head, rows: t.rows.slice(0, Math.max(0, +s.n || 0)) }),
    // Shape
    group: (t, s) => {
      const by = cols(t.head, s.by), [vi] = cols(t.head, s.col);
      const agg = s.agg || 'count';
      const groups = new Map();
      for (const r of t.rows) {
        const key = JSON.stringify(by.map(i => text(r[i])));
        if (!groups.has(key)) groups.set(key, { keys: by.map(i => r[i]), vals: [] });
        if (vi !== undefined) groups.get(key).vals.push(r[vi]);
        else groups.get(key).vals.push(1);
      }
      const f = vals => {
        const ns = vals.map(num).filter(n => !isNaN(n));
        switch (agg) {
          case 'sum': return ns.reduce((a, b) => a + b, 0);
          case 'average': return ns.length ? ns.reduce((a, b) => a + b, 0) / ns.length : '';
          case 'min': return ns.length ? Math.min(...ns) : '';
          case 'max': return ns.length ? Math.max(...ns) : '';
          case 'distinct': return new Set(vals.map(text)).size;
          default: return vals.length;
        }
      };
      const label = agg === 'count' ? 'Count' : `${agg[0].toUpperCase()}${agg.slice(1)} of ${vi !== undefined ? t.head[vi] : 'rows'}`;
      return { head: [...by.map(i => t.head[i]), label], rows: [...groups.values()].map(g => [...g.keys, f(g.vals)]) };
    },
    unpivot: (t, s) => {
      // The chosen columns become rows: one row per original row and column, as Attribute and Value.
      const ix = cols(t.head, s.cols), keepIx = t.head.map((_, i) => i).filter(i => !ix.includes(i));
      const rows = [];
      for (const r of t.rows) for (const i of ix) { if (!isBlank(r[i])) rows.push([...keepIx.map(k => r[k]), t.head[i], r[i]]); if (rows.length >= MAX_ROWS) break; }
      return { head: [...keepIx.map(i => t.head[i]), uniqueName(keepIx.map(i => t.head[i]), 'Attribute'), uniqueName(keepIx.map(i => t.head[i]), 'Value')], rows };
    }
  };

  /** What each step does, in words, for the Applied Steps list. */
  function describe(s) {
    const c = x => (Array.isArray(x) ? x.join(', ') : String(x));
    switch (s.op) {
      case 'trim': return `Trimmed spaces in ${c(s.cols)}`;
      case 'clean': return `Removed hidden characters in ${c(s.cols)}`;
      case 'case': return `${{ upper: 'UPPERCASE', lower: 'lowercase', proper: 'Capitalized Each Word' }[s.to] || 'Changed case'}: ${c(s.cols)}`;
      case 'replace': return `Replaced “${s.find}” with “${s.with || ''}” in ${c(s.cols)}`;
      case 'prefix': return `Added text to ${c(s.cols)}`;
      case 'remove': return `Removed ${c(s.cols)}`;
      case 'keep': return `Kept only ${c(s.cols)}`;
      case 'rename': return `Renamed ${s.col} to ${s.to}`;
      case 'split': return `Split ${s.col} by “${s.by}”`;
      case 'merge': return `Merged ${c(s.cols)} into ${s.name || 'one column'}`;
      case 'type': return `Changed ${c(s.cols)} to ${s.to}`;
      case 'fillDown': return `Filled down ${c(s.cols)}`;
      case 'filter': return `Kept rows where ${s.col} ${s.cond}${/blank/.test(s.cond) ? '' : ` “${s.value}”`}`;
      case 'removeBlankRows': return 'Removed blank rows';
      case 'dedupe': return s.cols && s.cols.length ? `Removed duplicates by ${c(s.cols)}` : 'Removed duplicate rows';
      case 'sort': return `Sorted by ${s.col} ${s.desc ? 'descending' : 'ascending'}`;
      case 'top': return `Kept the first ${s.n} rows`;
      case 'group': return `Grouped by ${c(s.by)} (${s.agg || 'count'}${s.col ? ' of ' + s.col : ''})`;
      case 'unpivot': return `Unpivoted ${c(s.cols)}`;
      default: return s.op;
    }
  }

  /** Runs the steps on a table ([header, ...rows]); returns [header, ...rows]. Bad steps are skipped. */
  function run(table, steps) {
    let t = { head: (table[0] || []).map((h, i) => text(h) || `Column ${i + 1}`), rows: table.slice(1, MAX_ROWS + 1).map(r => r.slice()) };
    const width = t.head.length;
    t.rows = t.rows.map(r => { while (r.length < width) r.push(''); return r.slice(0, width); });
    for (const s of steps || []) {
      const f = s && STEPS[s.op];
      if (!f) continue;
      try { t = f(t, s); } catch { /* a step that no longer fits the data is skipped */ }
    }
    return [t.head, ...t.rows];
  }

  /** A step from a file, rebuilt from known fields. */
  function cleanStep(s) {
    if (!s || typeof s !== 'object' || !STEPS[s.op]) return null;
    const str = v => (typeof v === 'string' ? v.slice(0, 500) : typeof v === 'number' && isFinite(v) ? v : '');
    const list = v => (Array.isArray(v) ? v.slice(0, 200).map(str).filter(x => x !== '') : []);
    const out = { op: s.op };
    for (const k of ['cols', 'by']) if (k in s) out[k] = Array.isArray(s[k]) ? list(s[k]) : str(s[k]);
    for (const k of ['col', 'to', 'find', 'with', 'before', 'after', 'name', 'sep', 'cond', 'value', 'agg']) if (k in s) out[k] = str(s[k]);
    for (const k of ['whole', 'desc', 'keepBlanks', 'trim']) if (typeof s[k] === 'boolean') out[k] = s[k];
    if ('n' in s) out.n = Math.max(0, Math.min(MAX_ROWS, +s.n || 0));
    return out;
  }

  const api = { run, describe, cleanStep, STEPS: Object.keys(STEPS), CONDITIONS: ['equals', 'not equals', 'contains', 'does not contain', 'begins with', 'ends with', 'greater than', 'less than', 'is blank', 'is not blank'] };
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.SheetsClean = api;
})(typeof window !== 'undefined' ? window : globalThis);
