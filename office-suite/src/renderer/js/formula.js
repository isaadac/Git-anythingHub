'use strict';
// Isadek Sheets formula engine: tokenizer, parser, evaluator and ~80 Excel-compatible functions.
// Pure JavaScript with no DOM access, so it runs in the app and under `node --test`.
(function (root) {
  // ---------- Cell addressing ----------
  function colToIndex(col) {
    let n = 0;
    for (const ch of col.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  }
  const COL_NAMES = []; // column letters, worked out once each (addresses are built constantly)
  function indexToCol(i) {
    const known = COL_NAMES[i];
    if (known !== undefined) return known;
    let s = '', n = i + 1;
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    if (i >= 0 && i < 16384) COL_NAMES[i] = s;
    return s;
  }
  function parseRef(ref) {
    const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(ref);
    if (!m) return null;
    return { col: colToIndex(m[1]), row: +m[2] - 1 };
  }
  const addr = (row, col) => indexToCol(col) + (row + 1);

  // ---------- Errors ----------
  class FormulaError {
    constructor(code) { this.error = code; }
    toString() { return this.error; }
  }
  const ERR = {
    DIV0: new FormulaError('#DIV/0!'), VALUE: new FormulaError('#VALUE!'), REF: new FormulaError('#REF!'),
    NAME: new FormulaError('#NAME?'), NA: new FormulaError('#N/A'), NUM: new FormulaError('#NUM!'),
    CIRC: new FormulaError('#CIRC!'), ERROR: new FormulaError('#ERROR!'), SPILL: new FormulaError('#SPILL!'), CALC: new FormulaError('#CALC!'),
    NULL: new FormulaError('#NULL!')
  };
  const isErr = v => v instanceof FormulaError;

  // ---------- Tokenizer ----------
  const TOKEN_RE = /\s*(?:(\d+\.?\d*(?:[eE][+-]?\d+)?|\.\d+(?:[eE][+-]?\d+)?)|("(?:[^"]|"")*")|((?:'(?:[^']|'')+'|[A-Za-z_][\w.]*(?::[A-Za-z_][\w.]*)?)!\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)|(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)(?![\w(\[])|((?:[A-Za-z_\\][\w.\\]*)?\[(?:[^\[\]']|'.|\[(?:[^\[\]']|'.)*\])*\])|([A-Za-z_\\][\w.\\]*)|(<>|<=|>=|[-+*/^&=<>%(),:;]))/y;

  const CP1252 = '€\u0081‚ƒ„…†‡ˆ‰Š‹Œ\u008DŽ\u008F\u0090‘’“”•–—˜™š›œ\u009DžŸ';
  const ERR_LIT_RE = /^\s*(?:(?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?(#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|SPILL!|CALC!|GETTING_DATA))/i;
  function tokenize(src) {
    const tokens = [];
    TOKEN_RE.lastIndex = 0;
    let pos = 0;
    while (pos < src.length) {
      if (/^\s*$/.test(src.slice(pos))) break;
      const er = ERR_LIT_RE.exec(src.slice(pos));
      if (er) { pos += er[0].length; tokens.push({ t: 'err', v: er[1].toUpperCase() }); continue; }
      TOKEN_RE.lastIndex = pos;
      const m = TOKEN_RE.exec(src);
      if (!m) throw ERR.ERROR;
      pos = TOKEN_RE.lastIndex;
      if (m[1] !== undefined) tokens.push({ t: 'num', v: parseFloat(m[1]) });
      else if (m[2] !== undefined) tokens.push({ t: 'str', v: m[2].slice(1, -1).replace(/""/g, '"') });
      else if (m[3] !== undefined || m[4] !== undefined) {
        // A1# refers to the whole range a formula in A1 spills into.
        if (src[pos] === '#') { pos++; tokens.push({ t: 'ref', v: m[3] || m[4], spill: true }); }
        else tokens.push({ t: 'ref', v: m[3] || m[4] });
      }
      else if (m[5] !== undefined) tokens.push({ t: 'sref', v: m[5] });
      else if (m[6] !== undefined) {
        const up = m[6].toUpperCase();
        if (up === 'TRUE' || up === 'FALSE') tokens.push({ t: 'bool', v: up === 'TRUE' });
        else tokens.push({ t: 'name', v: up });
      } else tokens.push({ t: 'op', v: m[7] === ';' ? ',' : m[7] });
    }
    return tokens;
  }

  // ---------- Parser (precedence climbing) ----------
  const BIN = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 };

  function parse(src) {
    const tokens = tokenize(src);
    let i = 0;
    const peek = () => tokens[i];
    const next = () => tokens[i++];
    const expectOp = v => { const t = next(); if (!t || t.t !== 'op' || t.v !== v) throw ERR.ERROR; };

    function primary() {
      const t = next();
      if (!t) throw ERR.ERROR;
      if (t.t === 'num') return { k: 'num', v: t.v };
      if (t.t === 'str') return { k: 'str', v: t.v };
      if (t.t === 'bool') {
        // TRUE() and FALSE() are functions too.
        if (peek() && peek().t === 'op' && peek().v === '(' && tokens[i + 1] && tokens[i + 1].t === 'op' && tokens[i + 1].v === ')') i += 2;
        return { k: 'bool', v: t.v };
      }
      if (t.t === 'err') return { k: 'err', v: t.v };
      if (t.t === 'ref') {
        const n = refNode(t.v);
        if (!t.spill) return n;
        if (n.k !== 'ref') throw ERR.ERROR;
        return { ...n, k: 'spill' };
      }
      if (t.t === 'sref') return { k: 'tref', ...parseStructured(t.v) };
      if (t.t === 'name') {
        if (peek() && peek().t === 'op' && peek().v === '(') {
          next();
          const args = [];
          if (!(peek() && peek().t === 'op' && peek().v === ')')) {
            do {
              // Allow empty arguments like IF(A1,,1)
              if (peek() && peek().t === 'op' && (peek().v === ',' || peek().v === ')')) args.push({ k: 'empty' });
              else args.push(expr(0));
            } while (peek() && peek().t === 'op' && peek().v === ',' && next());
          }
          expectOp(')');
          return { k: 'fn', name: t.v, args };
        }
        return { k: 'name', v: t.v };
      }
      if (t.t === 'op' && t.v === '(') { const e = expr(0); expectOp(')'); return e; }
      if (t.t === 'op' && (t.v === '-' || t.v === '+')) return { k: 'neg', sign: t.v, e: expr(6) };
      throw ERR.ERROR;
    }
    function postfix() {
      let e = primary();
      while (peek() && peek().t === 'op' && peek().v === '%') { next(); e = { k: 'pct', e }; }
      return e;
    }
    function expr(minPrec) {
      let left = postfix();
      for (;;) {
        const t = peek();
        if (!t || t.t !== 'op' || !(t.v in BIN)) break;
        const prec = BIN[t.v];
        if (prec < minPrec) break;
        next();
        const right = expr(t.v === '^' ? prec : prec + 1);
        left = { k: 'bin', op: t.v, l: left, r: right };
      }
      return left;
    }
    const ast = expr(0);
    if (i < tokens.length) throw ERR.ERROR;
    return ast;
  }

  /**
   * A table (structured) reference: Table1[Sales], Table1[#All], Table1[[#Headers],[Sales]],
   * Table1[[Q1]:[Q4]], Table1[@Sales], [@Sales] (the table the formula is in), Table1[[#This Row],[Sales]].
   * → { table, parts: Set of '#all' '#data' '#headers' '#totals', cols: [from, to] | null, thisRow }
   */
  function parseStructured(text) {
    const at = text.indexOf('[');
    const table = at > 0 ? text.slice(0, at) : null;
    let inner = text.slice(at + 1, -1).trim();
    const unesc = s => s.replace(/'(.)/g, '$1').trim();
    const out = { table, parts: new Set(), cols: null, thisRow: false };
    if (inner.startsWith('@')) { out.thisRow = true; inner = inner.slice(1).trim(); if (!inner) return out; if (inner[0] !== '[') { out.cols = [unesc(inner), unesc(inner)]; return out; } }
    if (!inner.includes('[')) {
      if (!inner) { out.parts.add('#data'); return out; }
      if (inner[0] === '#') { const p = inner.toLowerCase(); if (p === '#this row') out.thisRow = true; else out.parts.add(p); return out; }
      out.cols = [unesc(inner), unesc(inner)];
      return out;
    }
    // [[#Headers],[Sales]:[Cost]] — items separated by commas at the top level.
    const items = [];
    let depth = 0, cur = '';
    for (let i = 0; i < inner.length; i++) {
      const ch = inner[i];
      if (ch === "'") { cur += ch + (inner[i + 1] || ''); i++; continue; }
      if (ch === '[') depth++;
      if (ch === ']') depth--;
      if (ch === ',' && depth === 0) { items.push(cur.trim()); cur = ''; continue; }
      cur += ch;
    }
    if (cur.trim()) items.push(cur.trim());
    for (const it of items) {
      const m = it.match(/^\[((?:[^\]']|'.)*)\](?:\s*:\s*\[((?:[^\]']|'.)*)\])?$/);
      if (!m) throw ERR.REF;
      if (m[1][0] === '#' && !m[2]) { const p = m[1].toLowerCase(); if (p === '#this row') out.thisRow = true; else out.parts.add(p); continue; }
      out.cols = [unesc(m[1]), unesc(m[2] !== undefined ? m[2] : m[1])];
    }
    return out;
  }

  function refNode(text) {
    let sheet = null, body = text;
    const bang = text.lastIndexOf('!');
    if (bang >= 0) {
      sheet = text.slice(0, bang);
      if (sheet.startsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'");
      body = text.slice(bang + 1);
    }
    const [a, b] = body.split(':');
    const s = parseRef(a);
    const e = b ? parseRef(b) : s;
    // Sheet1:Sheet4!C1 — the same cells on every sheet from Sheet1 to Sheet4.
    if (sheet && sheet.includes(':')) {
      const [s1, s2] = sheet.split(':');
      return { k: 'ref3d', s1, s2, r1: Math.min(s.row, e.row), c1: Math.min(s.col, e.col), r2: Math.max(s.row, e.row), c2: Math.max(s.col, e.col) };
    }
    if (!b) return { k: 'ref', sheet, row: s.row, col: s.col };
    return { k: 'range', sheet, r1: Math.min(s.row, e.row), c1: Math.min(s.col, e.col), r2: Math.max(s.row, e.row), c2: Math.max(s.col, e.col) };
  }

  // ---------- Value helpers ----------
  const NOT_LITERAL = Symbol('not a literal');
  const FORMULA_CELL = Symbol('formula cell'); // in a column of values (book.columnValues): compute it
  class Range {
    constructor(rows) { this.rows = rows; }
    flat() { return this.rows.flat(); }
    get h() { return this.rows.length; }
    get w() { return this.rows.length ? this.rows[0].length : 0; }
  }
  /** Implicit intersection: the value of a one-column (or one-row) range in the given row (column). */
  function intersect(v, row, col) {
    if (!isMulti(v)) return v.rows[0] ? v.rows[0][0] : '';
    if (!v.at) return v.rows[0][0];
    if (v.w === 1) { const i = row - v.at.r1; return i >= 0 && i < v.h ? v.rows[i][0] : ERR.VALUE; }
    if (v.h === 1) { const j = col - v.at.c1; return j >= 0 && j < v.w ? v.rows[0][j] : ERR.VALUE; }
    return ERR.VALUE;
  }
  const isMulti = v => v instanceof Range && (v.rows.length > 1 || (v.rows[0] && v.rows[0].length > 1));
  /**
   * Applies f to matching elements of the arguments, as Excel does when a scalar operation gets
   * arrays: a 1-wide or 1-tall argument is repeated, a missing position is #N/A, and an error
   * stays in its own element.
   */
  function broadcast(args, f) {
    const dims = args.map(a => (a instanceof Range ? [a.h, a.w] : [1, 1]));
    const R = Math.max(...dims.map(d => d[0])), C = Math.max(...dims.map(d => d[1]));
    const pick = (a, d, r, c) => {
      if (!(a instanceof Range)) return a;
      const rr = d[0] === 1 ? 0 : r, cc = d[1] === 1 ? 0 : c;
      return rr < d[0] && cc < d[1] ? a.rows[rr][cc] : ERR.NA;
    };
    const rows = [];
    for (let r = 0; r < R; r++) {
      const row = [];
      for (let c = 0; c < C; c++) {
        const xs = args.map((a, i) => pick(a, dims[i], r, c));
        try { const v = f(xs); row.push(v instanceof Range ? v.flat()[0] : v); } catch (e) { if (isErr(e)) row.push(e); else throw e; }
      }
      rows.push(row);
    }
    return new Range(rows);
  }
  const isBlank = v => v === null || v === undefined || v === '';
  function toNum(v) {
    if (isErr(v)) throw v;
    if (v instanceof Range) { const f = v.flat(); if (f.length === 1) return toNum(f[0]); throw ERR.VALUE; }
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    if (isBlank(v)) return 0;
    const s = String(v).trim();
    if (/%$/.test(s) && !isNaN(parseFloat(s))) return parseFloat(s) / 100;
    const n = Number(s.replace(/[$,]/g, ''));
    if (s !== '' && !isNaN(n)) return n;
    throw ERR.VALUE;
  }
  function toStr(v) {
    if (isErr(v)) throw v;
    if (v instanceof Range) { const f = v.flat(); return toStr(f[0]); }
    if (isBlank(v)) return '';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'number') return formatGeneral(v);
    return String(v);
  }
  function toBool(v) {
    if (isErr(v)) throw v;
    if (v instanceof Range) return toBool(v.flat()[0]);
    if (typeof v === 'boolean') return v;
    if (typeof v === 'number') return v !== 0;
    if (isBlank(v)) return false;
    const s = String(v).toUpperCase();
    if (s === 'TRUE') return true;
    if (s === 'FALSE') return false;
    throw ERR.VALUE;
  }
  function formatGeneral(n) {
    if (!isFinite(n)) return String(n);
    if (Number.isInteger(n)) return String(n);
    const r = parseFloat(n.toPrecision(12));
    return Math.abs(r) >= 1e11 || (Math.abs(r) < 1e-9 && r !== 0) ? r.toExponential(5).replace(/\.?0+e/, 'E') : String(r);
  }
  /** Flattens arguments into a list of values, expanding ranges. */
  function flatten(args) {
    const out = [];
    for (const a of args) {
      if (a instanceof Range) out.push(...a.flat());
      else out.push(a);
    }
    return out;
  }
  /** Year, month and day of a serial number as Excel gives them (0 is 0 January 1900, 60 is 29 February 1900). */
  function excelYMD(serial) {
    if (serial < 0) throw ERR.NUM;
    const s = Math.floor(serial);
    if (s === 0) return [1900, 1, 0];
    if (s === 60) return [1900, 2, 29];
    const d = serialToDate(s);
    return [d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()];
  }
  /** Σ f(x, y) over the pairs of two same-sized arrays where both are numbers (SUMX2MY2 and kin). */
  function pairSum(a, f) {
    const xs = a[0] instanceof Range ? a[0].flat() : [a[0]], ys = a[1] instanceof Range ? a[1].flat() : [a[1]];
    if (xs.length !== ys.length) throw ERR.NA;
    let s = 0, n = 0;
    for (let i = 0; i < xs.length; i++) {
      if (isErr(xs[i])) throw xs[i];
      if (isErr(ys[i])) throw ys[i];
      if (typeof xs[i] === 'number' && typeof ys[i] === 'number') { s += f(xs[i], ys[i]); n++; }
    }
    if (!n) throw ERR.DIV0;
    return s;
  }
  /** The logical values AND, OR and XOR look at (Excel's rules: ranges skip text and blanks). */
  function logicals(args) {
    const out = [];
    for (const a of args) {
      if (a instanceof Range) {
        for (const v of a.flat()) { if (isErr(v)) throw v; if (typeof v === 'boolean') out.push(v); else if (typeof v === 'number') out.push(v !== 0); }
      } else if (isErr(a)) throw a;
      else if (typeof a === 'boolean') out.push(a);
      else if (typeof a === 'number') out.push(a !== 0);
      else if (typeof a === 'string' && a !== '') { const u = a.trim().toUpperCase(); if (u === 'TRUE' || u === 'FALSE') out.push(u === 'TRUE'); else throw ERR.VALUE; }
    }
    if (!out.length) throw ERR.VALUE;
    return out;
  }
  /** Numbers from args: from ranges only real numbers count; direct args are coerced (Excel rules). */
  function nums(args) {
    const out = [];
    for (const a of args) {
      if (a instanceof Range) { for (const v of a.flat()) { if (isErr(v)) throw v; if (typeof v === 'number') out.push(v); } }
      else if (!isBlank(a)) out.push(toNum(a));
    }
    return out;
  }

  // Excel date serials (1900 system; serial 1 = 1900-01-01, with the 1900 leap-year quirk).
  const EPOCH = Date.UTC(1899, 11, 30);
  // Excel counts 29 February 1900 (a day that didn't exist, kept from Lotus 1-2-3), so its serial
  // numbers before 1 March 1900 are one less than the calendar gives.
  const dateToSerial = d => { const s = (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds()) - EPOCH) / 86400000; return s < 61 && s >= 1 ? s - 1 : s; };
  const serialToDate = s => new Date(EPOCH + Math.round((s < 60 && s >= 1 ? s + 1 : s) * 86400000));

  /** Builds an Excel-style criteria matcher, e.g. ">5", "<>x", "a*", 10. */
  function criteria(c) {
    if (typeof c === 'number' || typeof c === 'boolean') return v => v === c || (typeof v === 'string' && v !== '' && Number(v) === c);
    const s = toStr(c);
    const m = /^(<=|>=|<>|<|>|=)?(.*)$/.exec(s);
    const op = m[1] || '=';
    const rhs = m[2];
    const rhsNum = rhs.trim() !== '' && !isNaN(Number(rhs)) ? Number(rhs) : null;
    const wild = new RegExp('^' + rhs.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/~\*/g, '\u0001').replace(/~\?/g, '\u0002').replace(/\*/g, '.*').replace(/\?/g, '.').replace(/\u0001/g, '\\*').replace(/\u0002/g, '\\?') + '$', 'i');
    return v => {
      if (isErr(v)) return false;
      if (rhsNum !== null && (typeof v === 'number' || (typeof v === 'string' && v.trim() !== '' && !isNaN(Number(v))))) {
        const n = Number(v);
        return { '=': n === rhsNum, '<>': n !== rhsNum, '<': n < rhsNum, '>': n > rhsNum, '<=': n <= rhsNum, '>=': n >= rhsNum }[op];
      }
      if (op === '=') return rhs === '' ? isBlank(v) : wild.test(toStr(v));
      if (op === '<>') return rhs === '' ? !isBlank(v) : !wild.test(toStr(v));
      if (rhsNum !== null) return false;
      const a = toStr(v).toLowerCase(), b = rhs.toLowerCase();
      return { '<': a < b, '>': a > b, '<=': a <= b, '>=': a >= b }[op];
    };
  }
  const roundTo = (n, d) => { const f = Math.pow(10, d); return Math.round((n + (n >= 0 ? 1 : -1) * Number.EPSILON * Math.abs(n)) * f) / f; };
  const compare = (a, b) => {
    if (typeof a === 'number' && typeof b === 'number') return a - b;
    return toStr(a).toLowerCase().localeCompare(toStr(b).toLowerCase());
  };
  const eqLoose = (a, b) => (typeof a === 'string' && typeof b === 'string') ? a.toLowerCase() === b.toLowerCase() : a === b || (!isBlank(a) && !isBlank(b) && String(a).toLowerCase() === String(b).toLowerCase());
  const asRange = v => v instanceof Range ? v : new Range([[v]]);
  function ifs(args, cb) {
    for (let i = 0; i < args.length; i += 2) {
      const r = asRange(args[i]).flat();
      const m = criteria(args[i + 1]);
      cb(r, m);
    }
  }

  // ---------- Functions ----------
  // Lazy functions receive (argNodes, evalNode) so they can short-circuit.
  const LAZY = {
    // Only the chosen value is worked out (an error elsewhere doesn't matter).
    CHOOSE: (a, ev) => { const i = Math.trunc(toNum(scalar(ev(a[0])))); if (i < 1 || i >= a.length) throw ERR.VALUE; return ev(a[i]); },
    IF: (a, ev) => {
      const c = ev(a[0]);
      if (isMulti(c)) {
        const t = a[1] && a[1].k !== 'empty' ? ev(a[1]) : a[1] ? 0 : true, f = a[2] && a[2].k !== 'empty' ? ev(a[2]) : a[2] ? 0 : false;
        return broadcast([c, t, f], ([x, y, z]) => (toBool(x) ? y : z));
      }
      return toBool(c) ? (a[1] ? ev(a[1]) : true) : (a[2] ? ev(a[2]) : false);
    },
    IFERROR: (a, ev) => { try { const v = ev(a[0]); if (isErr(v)) return ev(a[1]); if (v instanceof Range && isErr(v.flat()[0])) return ev(a[1]); return v; } catch (e) { if (isErr(e)) return ev(a[1]); throw e; } },
    IFNA: (a, ev) => { try { const v = ev(a[0]); return v === ERR.NA ? ev(a[1]) : v; } catch (e) { if (e === ERR.NA) return ev(a[1]); throw e; } },
    IFS: (a, ev) => { for (let i = 0; i < a.length; i += 2) if (toBool(ev(a[i]))) return ev(a[i + 1]); throw ERR.NA; },
    ISERROR: (a, ev) => { try { const v = ev(a[0]); return isErr(v) || (v instanceof Range && isErr(v.flat()[0])); } catch (e) { if (isErr(e)) return true; throw e; } },
    ISERR: (a, ev) => { try { const v = ev(a[0]); return isErr(v) && v !== ERR.NA; } catch (e) { if (isErr(e)) return e !== ERR.NA; throw e; } },
    ISNA: (a, ev) => { try { return ev(a[0]) === ERR.NA; } catch (e) { return e === ERR.NA; } },
    SWITCH: (a, ev) => {
      const v = ev(a[0]);
      let i = 1;
      for (; i + 1 < a.length; i += 2) if (eqLoose(scalar(v), scalar(ev(a[i])))) return ev(a[i + 1]);
      if (i < a.length) return ev(a[i]);
      throw ERR.NA;
    }
  };
  const scalar = v => (v instanceof Range ? v.flat()[0] : v);

  const FN = {
    // Math
    SUM: a => nums(a).reduce((s, n) => s + n, 0),
    PRODUCT: a => { const n = nums(a); return n.length ? n.reduce((s, x) => s * x, 1) : 0; },
    AVERAGE: a => { const n = nums(a); if (!n.length) throw ERR.DIV0; return n.reduce((s, x) => s + x, 0) / n.length; },
    MIN: a => { const n = nums(a); return n.length ? Math.min(...n) : 0; },
    MAX: a => { const n = nums(a); return n.length ? Math.max(...n) : 0; },
    MEDIAN: a => { const n = nums(a).sort((x, y) => x - y); if (!n.length) throw ERR.NUM; const m = n.length >> 1; return n.length % 2 ? n[m] : (n[m - 1] + n[m]) / 2; },
    MODE: a => { const c = new Map(); let best = null, bc = 1; for (const n of nums(a)) { const k = (c.get(n) || 0) + 1; c.set(n, k); if (k > bc) { bc = k; best = n; } } if (best === null) throw ERR.NA; return best; },
    STDEV: a => { const n = nums(a); if (n.length < 2) throw ERR.DIV0; const m = n.reduce((s, x) => s + x, 0) / n.length; return Math.sqrt(n.reduce((s, x) => s + (x - m) ** 2, 0) / (n.length - 1)); },
    STDEVP: a => { const n = nums(a); if (!n.length) throw ERR.DIV0; const m = n.reduce((s, x) => s + x, 0) / n.length; return Math.sqrt(n.reduce((s, x) => s + (x - m) ** 2, 0) / n.length); },
    VAR: a => { const n = nums(a); if (n.length < 2) throw ERR.DIV0; const m = n.reduce((s, x) => s + x, 0) / n.length; return n.reduce((s, x) => s + (x - m) ** 2, 0) / (n.length - 1); },
    LARGE: a => { const n = nums([a[0]]).sort((x, y) => y - x); const k = toNum(a[1]); if (k < 1 || k > n.length) throw ERR.NUM; return n[k - 1]; },
    SMALL: a => { const n = nums([a[0]]).sort((x, y) => x - y); const k = toNum(a[1]); if (k < 1 || k > n.length) throw ERR.NUM; return n[k - 1]; },
    RANK: a => { const v = toNum(a[0]); const n = nums([a[1]]); const asc = a[2] !== undefined && toNum(a[2]) !== 0; if (!n.includes(v)) throw ERR.NA; return 1 + n.filter(x => (asc ? x < v : x > v)).length; },
    COUNT: a => flatten(a).filter(v => typeof v === 'number').length,
    COUNTA: a => flatten(a).filter(v => !isBlank(v)).length,
    COUNTBLANK: a => flatten(a).filter(isBlank).length,
    ABS: a => Math.abs(toNum(a[0])),
    SQRT: a => { const n = toNum(a[0]); if (n < 0) throw ERR.NUM; return Math.sqrt(n); },
    POWER: a => Math.pow(toNum(a[0]), toNum(a[1])),
    EXP: a => Math.exp(toNum(a[0])),
    LN: a => { const n = toNum(a[0]); if (n <= 0) throw ERR.NUM; return Math.log(n); },
    LOG: a => { const n = toNum(a[0]); const b = a[1] === undefined ? 10 : toNum(a[1]); if (n <= 0) throw ERR.NUM; return Math.log(n) / Math.log(b); },
    LOG10: a => Math.log10(toNum(a[0])),
    MOD: a => { const n = toNum(a[0]), d = toNum(a[1]); if (d === 0) throw ERR.DIV0; return n - d * Math.floor(n / d); },
    INT: a => Math.floor(toNum(a[0])),
    TRUNC: a => { const d = a[1] === undefined ? 0 : toNum(a[1]); const f = 10 ** d; return Math.trunc(toNum(a[0]) * f) / f; },
    ROUND: a => roundTo(toNum(a[0]), a[1] === undefined ? 0 : Math.trunc(toNum(a[1]))),
    ROUNDUP: a => { const d = a[1] === undefined ? 0 : Math.trunc(toNum(a[1])); const f = 10 ** d; const n = toNum(a[0]); return Math.sign(n) * Math.ceil(Math.abs(n) * f - 1e-9) / f; },
    ROUNDDOWN: a => { const d = a[1] === undefined ? 0 : Math.trunc(toNum(a[1])); const f = 10 ** d; const n = toNum(a[0]); return Math.sign(n) * Math.floor(Math.abs(n) * f + 1e-9) / f; },
    CEILING: a => { const s = a[1] === undefined ? 1 : toNum(a[1]); return s === 0 ? 0 : Math.ceil(toNum(a[0]) / s) * s; },
    FLOOR: a => { const s = a[1] === undefined ? 1 : toNum(a[1]); if (s === 0) { if (toNum(a[0]) === 0) return 0; throw ERR.DIV0; } return Math.floor(toNum(a[0]) / s) * s; },
    SIGN: a => Math.sign(toNum(a[0])),
    PI: () => Math.PI,
    SIN: a => Math.sin(toNum(a[0])), COS: a => Math.cos(toNum(a[0])), TAN: a => Math.tan(toNum(a[0])),
    RADIANS: a => toNum(a[0]) * Math.PI / 180, DEGREES: a => toNum(a[0]) * 180 / Math.PI,
    RAND: () => Math.random(),
    RANDBETWEEN: a => { const lo = Math.ceil(toNum(a[0])), hi = Math.floor(toNum(a[1])); return lo + Math.floor(Math.random() * (hi - lo + 1)); },
    SUMPRODUCT: a => {
      const arrs = a.map(x => asRange(x).flat());
      const len = arrs[0].length;
      if (arrs.some(r => r.length !== len)) throw ERR.VALUE;
      let s = 0;
      for (let i = 0; i < len; i++) s += arrs.reduce((p, r) => p * (typeof r[i] === 'number' ? r[i] : 0), 1);
      return s;
    },
    SUMIF: a => { const r = asRange(a[0]).flat(); const s = a[2] ? asRange(a[2]).flat() : r; const m = criteria(a[1]); let t = 0; r.forEach((v, i) => { if (m(v) && typeof s[i] === 'number') t += s[i]; }); return t; },
    COUNTIF: a => asRange(a[0]).flat().filter(criteria(a[1])).length,
    AVERAGEIF: a => { const r = asRange(a[0]).flat(); const s = a[2] ? asRange(a[2]).flat() : r; const m = criteria(a[1]); const v = []; r.forEach((x, i) => { if (m(x) && typeof s[i] === 'number') v.push(s[i]); }); if (!v.length) throw ERR.DIV0; return v.reduce((p, x) => p + x, 0) / v.length; },
    SUMIFS: a => { const s = asRange(a[0]).flat(); const ok = s.map(() => true); ifs(a.slice(1), (r, m) => r.forEach((v, i) => { if (!m(v)) ok[i] = false; })); return s.reduce((t, v, i) => t + (ok[i] && typeof v === 'number' ? v : 0), 0); },
    COUNTIFS: a => { const n = asRange(a[0]).flat().length; const ok = new Array(n).fill(true); ifs(a, (r, m) => r.forEach((v, i) => { if (!m(v)) ok[i] = false; })); return ok.filter(Boolean).length; },
    AVERAGEIFS: a => { const s = asRange(a[0]).flat(); const ok = s.map(() => true); ifs(a.slice(1), (r, m) => r.forEach((v, i) => { if (!m(v)) ok[i] = false; })); const v = s.filter((x, i) => ok[i] && typeof x === 'number'); if (!v.length) throw ERR.DIV0; return v.reduce((p, x) => p + x, 0) / v.length; },
    MAXIFS: a => { const s = asRange(a[0]).flat(); const ok = s.map(() => true); ifs(a.slice(1), (r, m) => r.forEach((v, i) => { if (!m(v)) ok[i] = false; })); const v = s.filter((x, i) => ok[i] && typeof x === 'number'); return v.length ? Math.max(...v) : 0; },
    MINIFS: a => { const s = asRange(a[0]).flat(); const ok = s.map(() => true); ifs(a.slice(1), (r, m) => r.forEach((v, i) => { if (!m(v)) ok[i] = false; })); const v = s.filter((x, i) => ok[i] && typeof x === 'number'); return v.length ? Math.min(...v) : 0; },

    // Logic
    AND: a => logicals(a).every(Boolean),
    OR: a => logicals(a).some(Boolean),
    XOR: a => logicals(a).filter(Boolean).length % 2 === 1,
    NOT: a => !toBool(a[0]),
    TRUE: () => true, FALSE: () => false,
    NA: () => { throw ERR.NA; },
    ISBLANK: a => isBlank(scalar(a[0])),
    ISNUMBER: a => typeof scalar(a[0]) === 'number',
    ISTEXT: a => typeof scalar(a[0]) === 'string' && scalar(a[0]) !== '',
    ISLOGICAL: a => typeof scalar(a[0]) === 'boolean',
    ISEVEN: a => Math.floor(Math.abs(toNum(a[0]))) % 2 === 0,
    ISODD: a => Math.floor(Math.abs(toNum(a[0]))) % 2 === 1,

    // Text
    LEN: a => toStr(a[0]).length,
    LEFT: a => toStr(a[0]).slice(0, a[1] === undefined ? 1 : toNum(a[1])),
    RIGHT: a => { const s = toStr(a[0]); const n = a[1] === undefined ? 1 : toNum(a[1]); return n <= 0 ? '' : s.slice(-n); },
    MID: a => toStr(a[0]).substr(toNum(a[1]) - 1, toNum(a[2])),
    UPPER: a => toStr(a[0]).toUpperCase(),
    LOWER: a => toStr(a[0]).toLowerCase(),
    PROPER: a => toStr(a[0]).toLowerCase().replace(/(^|[^a-z])([a-z])/g, (m, p, c) => p + c.toUpperCase()),
    TRIM: a => toStr(a[0]).trim().replace(/ {2,}/g, ' '),
    CONCAT: a => flatten(a).map(toStr).join(''),
    CONCATENATE: a => flatten(a).map(toStr).join(''),
    TEXTJOIN: a => { const d = toStr(a[0]); const skip = toBool(a[1]); return flatten(a.slice(2)).filter(v => !(skip && isBlank(v))).map(toStr).join(d); },
    REPT: a => toStr(a[0]).repeat(Math.max(0, toNum(a[1]))),
    SUBSTITUTE: a => {
      const s = toStr(a[0]), f = toStr(a[1]), r = toStr(a[2]);
      if (!f) return s;
      if (a[3] === undefined) return s.split(f).join(r);
      const n = Math.trunc(toNum(a[3])); let idx = -1;
      if (n < 1) throw ERR.VALUE;
      for (let k = 0; k < n; k++) { idx = s.indexOf(f, idx + 1); if (idx < 0) return s; }
      return s.slice(0, idx) + r + s.slice(idx + f.length);
    },
    REPLACE: a => { const s = toStr(a[0]); const st = Math.trunc(toNum(a[1])) - 1, n = Math.trunc(toNum(a[2])); if (st < 0 || n < 0) throw ERR.VALUE; return s.slice(0, st) + toStr(a[3]) + s.slice(st + n); },
    FIND: a => { const i = toStr(a[1]).indexOf(toStr(a[0]), (a[2] === undefined ? 1 : toNum(a[2])) - 1); if (i < 0) throw ERR.VALUE; return i + 1; },
    SEARCH: a => { const i = toStr(a[1]).toLowerCase().indexOf(toStr(a[0]).toLowerCase(), (a[2] === undefined ? 1 : toNum(a[2])) - 1); if (i < 0) throw ERR.VALUE; return i + 1; },
    EXACT: a => toStr(a[0]) === toStr(a[1]),
    VALUE: a => toNum(a[0]),
    T: a => (typeof scalar(a[0]) === 'string' ? scalar(a[0]) : ''),
    N: a => (typeof scalar(a[0]) === 'number' ? scalar(a[0]) : typeof scalar(a[0]) === 'boolean' ? +scalar(a[0]) : 0),
    // Character n (1–255) of the Windows character set, as Excel on Windows.
    CHAR: a => { const n = Math.trunc(toNum(a[0])); if (n < 1 || n > 255) throw ERR.VALUE; return n >= 128 && n < 160 ? CP1252[n - 128] : String.fromCharCode(n); },
    CODE: a => toStr(a[0]).charCodeAt(0) || 0,
    TEXT: a => formatValue(toNum(a[0]), toStr(a[1])),
    FIXED: a => toNum(a[0]).toLocaleString('en-US', { minimumFractionDigits: a[1] === undefined ? 2 : toNum(a[1]), maximumFractionDigits: a[1] === undefined ? 2 : toNum(a[1]), useGrouping: !(a[2] && toBool(a[2])) }),
    DOLLAR: a => {
      const d = a[1] === undefined ? 2 : Math.trunc(toNum(a[1]));
      const n = roundTo(toNum(a[0]), d);
      const body = formatValue(Math.abs(n), '$#,##0' + (d > 0 ? '.' + '0'.repeat(Math.min(d, 127)) : ''));
      return n < 0 ? '(' + body + ')' : body;
    },

    // Date & time
    TODAY: () => Math.floor(dateToSerial(new Date())),
    NOW: () => dateToSerial(new Date()),
    DATE: a => {
      // Years 0–1899 count from 1900; months and days may run over (DATE(2024, 14, 0)).
      let y = Math.trunc(toNum(a[0]));
      if (y >= 0 && y < 1900) y += 1900;
      if (y < 0 || y > 9999) throw ERR.NUM;
      const first = new Date(y, Math.trunc(toNum(a[1])) - 1, 1);
      const v = dateToSerial(first) + Math.trunc(toNum(a[2])) - 1;
      if (v < 0) throw ERR.NUM;
      return v;
    },
    TIME: a => (toNum(a[0]) * 3600 + toNum(a[1]) * 60 + toNum(a[2])) / 86400,
    YEAR: a => excelYMD(toNum(a[0]))[0],
    MONTH: a => excelYMD(toNum(a[0]))[1],
    DAY: a => excelYMD(toNum(a[0]))[2],
    HOUR: a => serialToDate(toNum(a[0])).getUTCHours(),
    MINUTE: a => serialToDate(toNum(a[0])).getUTCMinutes(),
    SECOND: a => serialToDate(toNum(a[0])).getUTCSeconds(),
    WEEKDAY: a => { const d = serialToDate(toNum(a[0])).getUTCDay(); const t = a[1] === undefined ? 1 : toNum(a[1]); return t === 2 ? ((d + 6) % 7) + 1 : t === 3 ? (d + 6) % 7 : d + 1; },
    DAYS: a => toNum(a[0]) - toNum(a[1]),
    EDATE: a => { const d = serialToDate(toNum(a[0])); return dateToSerial(new Date(d.getUTCFullYear(), d.getUTCMonth() + toNum(a[1]), d.getUTCDate())); },
    EOMONTH: a => { const d = serialToDate(toNum(a[0])); return dateToSerial(new Date(d.getUTCFullYear(), d.getUTCMonth() + toNum(a[1]) + 1, 0)); },
    DATEVALUE: a => { const t = Date.parse(toStr(a[0])); if (isNaN(t)) throw ERR.VALUE; return Math.floor(dateToSerial(new Date(t))); },

    // Lookup
    VLOOKUP: a => lookup(a, false),
    HLOOKUP: a => lookup(a, true),
    INDEX: a => {
      const r = asRange(a[0]).rows;
      let row = toNum(a[1]), col = a[2] === undefined ? 1 : toNum(a[2]);
      if (r.length === 1 && a[2] === undefined) { col = row; row = 1; }
      if (row < 1 || col < 1 || row > r.length || col > r[0].length) throw ERR.REF;
      return r[row - 1][col - 1];
    },
    MATCH: a => {
      const v = scalar(a[0]);
      const list = asRange(a[1]).flat();
      const type = a[2] === undefined ? 1 : toNum(a[2]);
      if (type === 0) {
        const m = typeof v === 'string' ? criteria(v) : (x => eqLoose(x, v));
        const i = list.findIndex(x => m(x));
        if (i < 0) throw ERR.NA;
        return i + 1;
      }
      let best = -1;
      list.forEach((x, i) => { if (!isBlank(x) && (type > 0 ? compare(x, v) <= 0 : compare(x, v) >= 0)) best = i; });
      if (best < 0) throw ERR.NA;
      return best + 1;
    },
    XLOOKUP: a => {
      const v = scalar(a[0]);
      const look = asRange(a[1]).flat();
      const ret = asRange(a[2]);
      const i = look.findIndex(x => eqLoose(x, v));
      if (i < 0) { if (a[3] !== undefined && a[3] !== null) return a[3]; throw ERR.NA; }
      const rows = ret.rows;
      return rows.length === 1 ? rows[0][i] : rows[i][0];
    },

    ROWS: a => asRange(a[0]).rows.length,
    COLUMNS: a => asRange(a[0]).rows[0].length,

    // Finance
    PMT: a => { const r = toNum(a[0]), n = toNum(a[1]), pv = toNum(a[2]), fv = a[3] === undefined ? 0 : toNum(a[3]), t = a[4] === undefined ? 0 : toNum(a[4]); if (r === 0) return -(pv + fv) / n; const p = Math.pow(1 + r, n); return -(r * (pv * p + fv)) / ((1 + r * t) * (p - 1)); },
    FV: a => { const r = toNum(a[0]), n = toNum(a[1]), pmt = toNum(a[2]), pv = a[3] === undefined ? 0 : toNum(a[3]), t = a[4] === undefined ? 0 : toNum(a[4]); if (r === 0) return -(pv + pmt * n); const p = Math.pow(1 + r, n); return -(pv * p + pmt * (1 + r * t) * (p - 1) / r); },
    PV: a => { const r = toNum(a[0]), n = toNum(a[1]), pmt = toNum(a[2]), fv = a[3] === undefined ? 0 : toNum(a[3]), t = a[4] === undefined ? 0 : toNum(a[4]); if (r === 0) return -(fv + pmt * n); const p = Math.pow(1 + r, n); return -(fv + pmt * (1 + r * t) * (p - 1) / r) / p; },
    NPV: a => { const r = toNum(a[0]); return nums(a.slice(1)).reduce((s, v, i) => s + v / Math.pow(1 + r, i + 1), 0); }
  };

  // ---------- More functions: dynamic arrays, references, statistics, finance, text, dates ----------
  const rowsOf = v => (v instanceof Range ? v.rows : [[v]]);
  const optN = (v, d) => (v === undefined || v === '' || v === null ? d : toNum(v));
  const optB = (v, d) => (v === undefined || v === '' || v === null ? d : toBool(v));
  const arrOut = rows => { if (!rows.length || !rows[0].length) throw ERR.CALC; return new Range(rows); };
  const transpose = rows => (rows.length ? rows[0].map((_, c) => rows.map(r => r[c])) : []);
  // Excel's sort order: numbers, text (case-insensitive), logicals, errors; blanks last.
  const rankType = v => (isBlank(v) ? 4 : typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : 3);
  const sortCmp = (a, b) => {
    const ta = rankType(a), tb = rankType(b);
    if (ta !== tb) return ta - tb;
    if (ta === 0) return a - b;
    if (ta === 1) return a.localeCompare(b, undefined, { sensitivity: 'base' });
    if (ta === 2) return +a - +b;
    return 0;
  };
  const keyOf = v => (typeof v === 'string' ? 's' + v.toLowerCase() : isErr(v) ? 'e' + v.error : typeof v + ':' + v);
  const pairs = (xs, ys) => {
    const a = asRange(xs).flat(), b = asRange(ys).flat();
    if (a.length !== b.length) throw ERR.NA;
    const out = [];
    a.forEach((x, i) => { if (isErr(x)) throw x; if (isErr(b[i])) throw b[i]; if (typeof x === 'number' && typeof b[i] === 'number') out.push([x, b[i]]); });
    return out;
  };
  const mean = n => n.reduce((s, x) => s + x, 0) / n.length;
  const sqDev = n => { const m = mean(n); return n.reduce((s, x) => s + (x - m) ** 2, 0); };
  const variance = (n, sample) => { if (n.length < (sample ? 2 : 1)) throw ERR.DIV0; return sqDev(n) / (n.length - (sample ? 1 : 0)); };
  function percentile(n, p, exc) {
    n = n.slice().sort((x, y) => x - y);
    if (!n.length) throw ERR.NUM;
    let rank;
    if (exc) { rank = p * (n.length + 1) - 1; if (rank < 0 || rank > n.length - 1) throw ERR.NUM; }
    else { if (p < 0 || p > 1) throw ERR.NUM; rank = p * (n.length - 1); }
    const lo = Math.floor(rank), f = rank - lo;
    return lo + 1 < n.length ? n[lo] + f * (n[lo + 1] - n[lo]) : n[lo];
  }
  // Normal distribution: erfc (W. J. Cody's rational approximations, ~1e-15) and Acklam's inverse refined by Newton.
  function erfc(x) {
    const z = Math.abs(x), t = 1 / (1 + 0.5 * z);
    const r = t * Math.exp(-z * z - 1.26551223 + t * (1.00002368 + t * (0.37409196 + t * (0.09678418 + t * (-0.18628806 + t * (0.27886807 + t * (-1.13520398 + t * (1.48851587 + t * (-0.82215223 + t * 0.17087277)))))))));
    return x >= 0 ? r : 2 - r;
  }
  function normCdf(z) {
    // Series for small |z|, continued fraction tail otherwise: accurate to ~1e-14.
    if (Math.abs(z) < 3) {
      let sum = z, term = z;
      for (let i = 1; i < 200; i++) { term *= z * z / (2 * i + 1); sum += term; if (Math.abs(term) < 1e-17 * Math.abs(sum)) break; }
      return 0.5 + sum * Math.exp(-z * z / 2) / Math.sqrt(2 * Math.PI);
    }
    return 0.5 * erfc(-z / Math.SQRT2);
  }
  const normPdf = z => Math.exp(-z * z / 2) / Math.sqrt(2 * Math.PI);
  function normInv(p) {
    if (p <= 0 || p >= 1) throw ERR.NUM;
    const a = [-39.69683028665376, 220.9460984245205, -275.9285104469687, 138.357751867269, -30.66479806614716, 2.506628277459239];
    const b = [-54.47609879822406, 161.5858368580409, -155.6989798598866, 66.80131188771972, -13.28068155288572];
    const c = [-0.007784894002430293, -0.3223964580411365, -2.400758277161838, -2.549732539343734, 4.374664141464968, 2.938163982698783];
    const d = [0.007784695709041462, 0.3224671290700398, 2.445134137142996, 3.754408661907416];
    let x;
    if (p < 0.02425) { const q = Math.sqrt(-2 * Math.log(p)); x = (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1); }
    else if (p > 1 - 0.02425) { const q = Math.sqrt(-2 * Math.log(1 - p)); x = -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1); }
    else { const q = p - 0.5, r = q * q; x = (((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q / (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1); }
    for (let i = 0; i < 3; i++) x -= (normCdf(x) - p) / normPdf(x);
    return x;
  }
  function lgamma(x) {
    const g = [76.18009172947146, -86.50532032941677, 24.01409824083091, -1.231739572450155, 0.1208650973866179e-2, -0.5395239384953e-5];
    let y = x, tmp = x + 5.5; tmp -= (x + 0.5) * Math.log(tmp);
    let ser = 1.000000000190015;
    for (const k of g) ser += k / ++y;
    return -tmp + Math.log(2.5066282746310005 * ser / x);
  }
  const lnComb = (n, k) => lgamma(n + 1) - lgamma(k + 1) - lgamma(n - k + 1);
  function newton(f, guess) {
    let x = guess;
    for (let i = 0; i < 100; i++) {
      const y = f(x), h = 1e-7 * Math.max(1, Math.abs(x)), d = (f(x + h) - y) / h;
      if (!isFinite(y) || !isFinite(d) || d === 0) break;
      const nx = x - y / d;
      if (Math.abs(nx - x) < 1e-12) return nx;
      x = nx;
    }
    if (Math.abs(f(x)) < 1e-7) return x;
    throw ERR.NUM;
  }
  const ipmt = (rate, per, nper, pv, fv, type) => {
    const pmt = FN.PMT([rate, nper, pv, fv, type]);
    let interest;
    if (per === 1) interest = type === 1 ? 0 : -pv;
    else interest = type === 1 ? FN.FV([rate, per - 2, pmt, pv, 1]) - pmt : FN.FV([rate, per - 1, pmt, pv, 0]);
    return interest * rate;
  };
  const ymd = s => { const d = serialToDate(s); return [d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()]; };
  const isWeekend = (s, mask) => mask[(serialToDate(s).getUTCDay() + 6) % 7] === '1'; // mask: Mon..Sun
  function weekendMask(w) {
    if (w === undefined || w === '' || w === null) return '0000011';
    if (typeof w === 'string' && /^[01]{7}$/.test(w)) return w;
    const n = toNum(w);
    const m = { 1: '0000011', 2: '1000001', 3: '1100000', 4: '0110000', 5: '0011000', 6: '0001100', 7: '0000110', 11: '0000001', 12: '1000000', 13: '0100000', 14: '0010000', 15: '0001000', 16: '0000100', 17: '0000010' }[n];
    if (!m) throw ERR.NUM;
    return m;
  }
  const holidays = h => new Set(h === undefined ? [] : nums([h]).map(Math.floor));
  function networkdays(start, end, mask, hol) {
    start = Math.floor(start); end = Math.floor(end);
    const sign = end < start ? -1 : 1;
    if (sign < 0) [start, end] = [end, start];
    let n = 0;
    for (let s = start; s <= end; s++) if (!isWeekend(s, mask) && !hol.has(s)) n++;
    return n * sign;
  }
  function workday(start, days, mask, hol) {
    let s = Math.floor(start);
    const step = days < 0 ? -1 : 1;
    if (mask === '1111111') throw ERR.VALUE;
    for (let left = Math.abs(Math.trunc(days)); left > 0;) { s += step; if (!isWeekend(s, mask) && !hol.has(s)) left--; }
    return s;
  }
  const ROMAN = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];
  function toBase(n, radix, bits, places) {
    n = Math.trunc(toNum(n));
    const lim = 2 ** (bits - 1);
    if (n < -lim || n >= lim) throw ERR.NUM;
    if (n < 0) return (2 ** bits + n).toString(radix).toUpperCase();
    const s = n.toString(radix).toUpperCase();
    if (places === undefined) return s;
    const p = toNum(places);
    if (p < s.length) throw ERR.NUM;
    return s.padStart(p, '0');
  }
  function fromBase(text, radix, bits) {
    const s = toStr(text).trim();
    if (s.length > 10 || !new RegExp(`^[${'0123456789ABCDEF'.slice(0, radix)}]*$`, 'i').test(s)) throw ERR.NUM;
    const n = parseInt(s || '0', radix);
    return s.length === 10 && n >= 2 ** (bits - 1) ? n - 2 ** bits : n;
  }
  function textSplitAt(text, delims, instance, caseless, matchEnd) {
    const t = caseless ? text.toLowerCase() : text;
    const ds = delims.map(d => (caseless ? d.toLowerCase() : d)).filter(d => d !== '');
    if (!ds.length) return instance > 0 ? [0, 0] : [text.length, text.length];
    const hits = [];
    for (let i = 0; i < t.length; i++) for (const d of ds) if (t.startsWith(d, i)) { hits.push([i, i + d.length]); break; }
    if (matchEnd) hits.push([t.length, t.length]);
    const h = instance > 0 ? hits[instance - 1] : hits[hits.length + instance];
    return h || null;
  }
  function xmatch(v, list, mode, search) {
    v = scalar(v);
    const idx = [...list.keys()];
    if (search < 0) idx.reverse();
    if (mode === 2) { const m = criteria(typeof v === 'string' ? v : toStr(v)); const i = idx.find(k => m(list[k])); if (i === undefined) throw ERR.NA; return i; }
    let best = -1;
    for (const k of idx) {
      const x = list[k];
      if (isBlank(x) && !isBlank(v)) continue;
      const c = sortCmp(x, v);
      if (rankType(x) === rankType(v) && c === 0) return k;
      if (mode === -1 && rankType(x) === rankType(v) && c < 0 && (best < 0 || sortCmp(x, list[best]) > 0)) best = k;
      if (mode === 1 && rankType(x) === rankType(v) && c > 0 && (best < 0 || sortCmp(x, list[best]) < 0)) best = k;
    }
    if (best < 0) throw ERR.NA;
    return best;
  }
  function minverse(m) {
    const n = m.length;
    if (!n || m.some(r => r.length !== n)) throw ERR.VALUE;
    const a = m.map((r, i) => [...r.map(toNum), ...Array.from({ length: n }, (_, j) => (i === j ? 1 : 0))]);
    let det = 1;
    for (let c = 0; c < n; c++) {
      let p = c;
      for (let r = c + 1; r < n; r++) if (Math.abs(a[r][c]) > Math.abs(a[p][c])) p = r;
      if (Math.abs(a[p][c]) < 1e-15) return { inv: null, det: 0 };
      if (p !== c) { [a[p], a[c]] = [a[c], a[p]]; det = -det; }
      const pv = a[c][c]; det *= pv;
      for (let j = 0; j < 2 * n; j++) a[c][j] /= pv;
      for (let r = 0; r < n; r++) if (r !== c) { const f = a[r][c]; if (f) for (let j = 0; j < 2 * n; j++) a[r][j] -= f * a[c][j]; }
    }
    return { inv: a.map(r => r.slice(n)), det };
  }

  /** Functions that take single values; given arrays they work per element. */
  const LIFT = new Set(('ABS SQRT POWER EXP LN LOG LOG10 MOD INT TRUNC ROUND ROUNDUP ROUNDDOWN CEILING FLOOR SIGN SIN COS TAN RADIANS DEGREES ' +
    'LEN LEFT RIGHT MID UPPER LOWER PROPER TRIM REPT SUBSTITUTE REPLACE FIND SEARCH EXACT VALUE CHAR CODE TEXT FIXED DOLLAR T N ' +
    'DATE TIME YEAR MONTH DAY HOUR MINUTE SECOND WEEKDAY DAYS EDATE EOMONTH DATEVALUE ISBLANK ISNUMBER ISTEXT ISLOGICAL ISEVEN ISODD NOT ' +
    'MROUND QUOTIENT FACT COMBIN PERMUT ODD EVEN ASIN ACOS ATAN ATAN2 SINH COSH TANH ASINH ACOSH ATANH SQRTPI CEILING.MATH FLOOR.MATH ' +
    'CEILING.PRECISE FLOOR.PRECISE ROMAN ARABIC BASE DECIMAL DEC2BIN BIN2DEC DEC2HEX HEX2DEC DEC2OCT OCT2DEC CLEAN UNICHAR UNICODE ' +
    'NUMBERVALUE TEXTBEFORE TEXTAFTER DATEDIF WEEKNUM ISOWEEKNUM YEARFRAC TIMEVALUE ISNONTEXT STANDARDIZE NORM.DIST NORMDIST NORM.S.DIST ' +
    'NORMSDIST NORM.INV NORMINV NORM.S.INV NORMSINV EXPON.DIST POISSON.DIST BINOM.DIST ADDRESS EFFECT NOMINAL SLN SYD').split(' '));
  /** Functions given an error as a value to look at, not one to pass on. */
  const TESTS_ERRORS = new Set(['ISNUMBER', 'ISTEXT', 'ISBLANK', 'ISLOGICAL', 'ISNONTEXT', 'ISREF', 'ISFORMULA']);
  /** Lazily evaluated functions whose arguments a formula from an older Excel intersects. */
  const LEGACY_LAZY = new Set(['ISERROR', 'ISERR', 'ISNA', 'IF', 'IFERROR', 'IFNA', 'CHOOSE']);
  /** Functions that treat a reference to a single cell like a range: its text is skipped. */
  const REF_AS_RANGE = new Set(('SUM PRODUCT AVERAGE MIN MAX MEDIAN MODE STDEV STDEVP VAR VARP STDEV.S STDEV.P VAR.S VAR.P ' +
    'GEOMEAN HARMEAN DEVSQ AVEDEV SKEW KURT SUMSQ AND OR XOR COUNT').split(' '));
  /** Functions whose result can be an array (they spill). */
  const ARRAY_FNS = new Set(('FILTER SORT SORTBY UNIQUE SEQUENCE RANDARRAY TRANSPOSE TOCOL TOROW VSTACK HSTACK TAKE DROP CHOOSECOLS CHOOSEROWS ' +
    'WRAPROWS WRAPCOLS EXPAND TEXTSPLIT MMULT MINVERSE FREQUENCY MODE.MULT OFFSET INDIRECT INDEX XLOOKUP ROW COLUMN').split(' '));
  function maySpill(n) {
    if (!n || typeof n !== 'object') return false;
    switch (n.k) {
      case 'range': case 'spill': case 'name': return true;
      case 'tref': return !n.thisRow;
      case 'bin': return maySpill(n.l) || maySpill(n.r);
      case 'neg': case 'pct': return maySpill(n.e);
      case 'fn':
        if (ARRAY_FNS.has(n.name)) return true;
        if (LIFT.has(n.name) || LAZY[n.name]) return n.args.some(maySpill);
        return false;
    }
    return false;
  }

  Object.assign(LAZY, {
    LET: (a, ev, ctx) => {
      if (a.length < 3 || a.length % 2 === 0) throw ERR.VALUE;
      const scope = new Map();
      ctx.engine.scopes.push(scope);
      try {
        for (let i = 0; i < a.length - 1; i += 2) { if (a[i].k !== 'name') throw ERR.VALUE; scope.set(a[i].v, ev(a[i + 1])); }
        return ev(a[a.length - 1]);
      } finally { ctx.engine.scopes.pop(); }
    },
    INDIRECT: (a, ev, ctx) => {
      const text = toStr(scalar(ev(a[0]))).trim();
      if (a[1] && a[1].k !== 'empty' && !toBool(ev(a[1]))) throw ERR.REF; // R1C1 style isn't supported
      if (!/^(?:(?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$/.test(text)) throw ERR.REF;
      return ctx.engine.evalNode(refNode(text.replace(/\$/g, '')), ctx.sheet);
    },
    OFFSET: (a, ev, ctx) => {
      const base = a[0];
      if (!base || (base.k !== 'ref' && base.k !== 'range')) throw ERR.VALUE;
      const r1 = base.k === 'ref' ? base.row : base.r1, c1 = base.k === 'ref' ? base.col : base.c1;
      const h0 = base.k === 'ref' ? 1 : base.r2 - base.r1 + 1, w0 = base.k === 'ref' ? 1 : base.c2 - base.c1 + 1;
      const dr = toNum(scalar(ev(a[1]))), dc = toNum(scalar(ev(a[2])));
      const h = a[3] && a[3].k !== 'empty' ? toNum(scalar(ev(a[3]))) : h0, w = a[4] && a[4].k !== 'empty' ? toNum(scalar(ev(a[4]))) : w0;
      const r = r1 + Math.trunc(dr), c = c1 + Math.trunc(dc);
      if (r < 0 || c < 0 || h < 1 || w < 1 || r + h > 1048576 || c + w > 16384) throw ERR.REF;
      const node = h === 1 && w === 1 ? { k: 'ref', sheet: base.sheet, row: r, col: c } : { k: 'range', sheet: base.sheet, r1: r, c1: c, r2: r + Math.trunc(h) - 1, c2: c + Math.trunc(w) - 1 };
      return ctx.engine.evalNode(node, ctx.sheet);
    },
    ROW: (a, ev, ctx) => {
      const n = a[0];
      if (!n || n.k === 'empty') { if (!ctx.engine.cur) throw ERR.VALUE; return ctx.engine.cur.row + 1; }
      if (n.k === 'ref' || n.k === 'spill') return n.row + 1;
      if (n.k === 'range') return n.r1 === n.r2 ? n.r1 + 1 : new Range(Array.from({ length: n.r2 - n.r1 + 1 }, (_, i) => [n.r1 + i + 1]));
      throw ERR.VALUE;
    },
    COLUMN: (a, ev, ctx) => {
      const n = a[0];
      if (!n || n.k === 'empty') { if (!ctx.engine.cur) throw ERR.VALUE; return ctx.engine.cur.col + 1; }
      if (n.k === 'ref' || n.k === 'spill') return n.col + 1;
      if (n.k === 'range') return n.c1 === n.c2 ? n.c1 + 1 : new Range([Array.from({ length: n.c2 - n.c1 + 1 }, (_, i) => n.c1 + i + 1)]);
      throw ERR.VALUE;
    },
    ISREF: a => !!a[0] && ['ref', 'range', 'spill'].includes(a[0].k),
    ISFORMULA: (a, ev, ctx) => {
      if (!a[0] || a[0].k !== 'ref') throw ERR.VALUE;
      const raw = ctx.engine.book.getRaw(ctx.engine.resolveSheet(a[0].sheet, ctx.sheet), a[0].row, a[0].col);
      return typeof raw === 'string' && raw[0] === '=' && raw.length > 1;
    },
    FORMULATEXT: (a, ev, ctx) => {
      if (!a[0] || a[0].k !== 'ref') throw ERR.VALUE;
      const raw = ctx.engine.book.getRaw(ctx.engine.resolveSheet(a[0].sheet, ctx.sheet), a[0].row, a[0].col);
      if (typeof raw !== 'string' || raw[0] !== '=') throw ERR.NA;
      return raw;
    },
    'ERROR.TYPE': (a, ev) => {
      let v;
      try { v = scalar(ev(a[0])); } catch (e) { if (!isErr(e)) throw e; v = e; }
      const code = isErr(v) && { '#NULL!': 1, '#DIV/0!': 2, '#VALUE!': 3, '#REF!': 4, '#NAME?': 5, '#NUM!': 6, '#N/A': 7, '#SPILL!': 9, '#CALC!': 14 }[v.error];
      if (!code) throw ERR.NA;
      return code;
    },
    TYPE: (a, ev) => {
      let v;
      try { v = ev(a[0]); } catch (e) { if (isErr(e)) return 16; throw e; }
      if (isMulti(v)) return 64;
      v = scalar(v);
      return isErr(v) ? 16 : typeof v === 'number' || isBlank(v) ? 1 : typeof v === 'string' ? 2 : typeof v === 'boolean' ? 4 : 1;
    }
  });

  Object.assign(FN, {
    // Dynamic arrays
    FILTER: a => {
      const rows = rowsOf(a[0]), inc = rowsOf(a[1]);
      const ok = v => { if (isErr(v)) throw v; return toBool(v); };
      let out;
      if (inc.length === rows.length && inc[0].length === 1) out = rows.filter((_, i) => ok(inc[i][0]));
      else if (inc.length === 1 && inc[0].length === rows[0].length) out = rows.map(r => r.filter((_, j) => ok(inc[0][j])));
      else throw ERR.VALUE;
      if (!out.length || !out[0].length) { if (a.length > 2) return a[2]; throw ERR.CALC; }
      return new Range(out);
    },
    SORT: a => {
      let rows = rowsOf(a[0]);
      const byCol = optB(a[3], false);
      if (byCol) rows = transpose(rows);
      const idx = optN(a[1], 1) - 1, order = optN(a[2], 1);
      if (idx < 0 || idx >= rows[0].length || (order !== 1 && order !== -1)) throw ERR.VALUE;
      const sorted = rows.map((r, i) => [r, i]).sort((x, y) => sortCmp(x[0][idx], y[0][idx]) * order || x[1] - y[1]).map(x => x[0]);
      return new Range(byCol ? transpose(sorted) : sorted);
    },
    SORTBY: a => {
      const rows = rowsOf(a[0]);
      const keys = [];
      for (let i = 1; i < a.length; i += 2) {
        const by = rowsOf(a[i]), order = optN(a[i + 1], 1);
        if (by.length === rows.length && by[0].length === 1) keys.push({ v: by.map(r => r[0]), order, col: false });
        else if (by.length === 1 && by[0].length === rows[0].length) keys.push({ v: by[0], order, col: true });
        else throw ERR.VALUE;
      }
      if (!keys.length) throw ERR.VALUE;
      const byCol = keys[0].col;
      if (keys.some(k => k.col !== byCol)) throw ERR.VALUE;
      const items = byCol ? transpose(rows) : rows;
      const sorted = items.map((r, i) => [r, i]).sort((x, y) => { for (const k of keys) { const d = sortCmp(k.v[x[1]], k.v[y[1]]) * k.order; if (d) return d; } return x[1] - y[1]; }).map(x => x[0]);
      return new Range(byCol ? transpose(sorted) : sorted);
    },
    UNIQUE: a => {
      const byCol = optB(a[1], false), once = optB(a[2], false);
      const items = byCol ? transpose(rowsOf(a[0])) : rowsOf(a[0]);
      const count = new Map();
      items.forEach(r => { const k = r.map(keyOf).join('\u0001'); count.set(k, (count.get(k) || 0) + 1); });
      const seen = new Set();
      const out = items.filter(r => { const k = r.map(keyOf).join('\u0001'); if (seen.has(k)) return false; seen.add(k); return !once || count.get(k) === 1; });
      if (!out.length) throw ERR.CALC;
      return new Range(byCol ? transpose(out) : out);
    },
    SEQUENCE: a => {
      const r = Math.trunc(toNum(a[0])), c = Math.trunc(optN(a[1], 1)), start = optN(a[2], 1), step = optN(a[3], 1);
      if (r < 1 || c < 1) throw ERR.CALC;
      if (r * c > 1048576) throw ERR.NUM;
      return new Range(Array.from({ length: r }, (_, i) => Array.from({ length: c }, (_, j) => start + (i * c + j) * step)));
    },
    RANDARRAY: a => {
      const r = Math.trunc(optN(a[0], 1)), c = Math.trunc(optN(a[1], 1)), lo = optN(a[2], 0), hi = optN(a[3], 1), int = optB(a[4], false);
      if (r < 1 || c < 1) throw ERR.CALC;
      if (hi < lo) throw ERR.VALUE;
      return new Range(Array.from({ length: r }, () => Array.from({ length: c }, () => (int ? Math.ceil(lo) + Math.floor(Math.random() * (Math.floor(hi) - Math.ceil(lo) + 1)) : lo + Math.random() * (hi - lo)))));
    },
    TRANSPOSE: a => new Range(transpose(rowsOf(a[0]))),
    TOCOL: a => {
      const ignore = optN(a[1], 0), byCol = optB(a[2], false);
      const rows = byCol ? transpose(rowsOf(a[0])) : rowsOf(a[0]);
      const flat = rows.flat().filter(v => !((ignore & 1) && isBlank(v)) && !((ignore & 2) && isErr(v)));
      return arrOut(flat.map(v => [v]));
    },
    TOROW: a => { const col = FN.TOCOL(a); return new Range([col.flat()]); },
    VSTACK: a => { const parts = a.map(rowsOf); const w = Math.max(...parts.map(p => p[0].length)); return arrOut(parts.flatMap(p => p.map(r => [...r, ...Array(w - r.length).fill(ERR.NA)]))); },
    HSTACK: a => { const parts = a.map(rowsOf); const h = Math.max(...parts.map(p => p.length)); return arrOut(Array.from({ length: h }, (_, i) => parts.flatMap(p => p[i] || Array(p[0].length).fill(ERR.NA)))); },
    TAKE: a => {
      let rows = rowsOf(a[0]);
      const r = a[1] === undefined || a[1] === '' ? rows.length : Math.trunc(toNum(a[1])), c = a[2] === undefined || a[2] === '' ? rows[0].length : Math.trunc(toNum(a[2]));
      rows = r >= 0 ? rows.slice(0, r) : rows.slice(r);
      return arrOut(rows.map(row => (c >= 0 ? row.slice(0, c) : row.slice(c))));
    },
    DROP: a => {
      let rows = rowsOf(a[0]);
      const r = Math.trunc(optN(a[1], 0)), c = Math.trunc(optN(a[2], 0));
      rows = r >= 0 ? rows.slice(r) : rows.slice(0, r);
      return arrOut(rows.map(row => (c >= 0 ? row.slice(c) : row.slice(0, c))));
    },
    CHOOSEROWS: a => { const rows = rowsOf(a[0]); return arrOut(nums(a.slice(1)).map(n => { const i = n > 0 ? n - 1 : rows.length + n; if (!rows[i] || n === 0) throw ERR.VALUE; return rows[i]; })); },
    CHOOSECOLS: a => { const cols = transpose(rowsOf(a[0])); return arrOut(transpose(nums(a.slice(1)).map(n => { const i = n > 0 ? n - 1 : cols.length + n; if (!cols[i] || n === 0) throw ERR.VALUE; return cols[i]; }))); },
    WRAPROWS: a => {
      const v = rowsOf(a[0]); if (v.length > 1 && v[0].length > 1) throw ERR.VALUE;
      const flat = v.flat(), n = Math.trunc(toNum(a[1])), pad = a.length > 2 ? a[2] : ERR.NA;
      if (n < 1) throw ERR.NUM;
      const out = [];
      for (let i = 0; i < flat.length; i += n) { const row = flat.slice(i, i + n); while (row.length < n) row.push(pad); out.push(row); }
      return arrOut(out);
    },
    WRAPCOLS: a => new Range(transpose(FN.WRAPROWS(a).rows)),
    EXPAND: a => {
      const rows = rowsOf(a[0]);
      const r = a[1] === undefined || a[1] === '' ? rows.length : toNum(a[1]), c = a[2] === undefined || a[2] === '' ? rows[0].length : toNum(a[2]);
      const pad = a.length > 3 ? a[3] : ERR.NA;
      if (r < rows.length || c < rows[0].length) throw ERR.VALUE;
      return new Range(Array.from({ length: r }, (_, i) => Array.from({ length: c }, (_, j) => (rows[i] && j < rows[i].length ? rows[i][j] : pad))));
    },
    TEXTSPLIT: a => {
      const text = toStr(a[0]);
      const cd = a[1] === undefined || a[1] === '' ? [] : asRange(a[1]).flat().map(toStr);
      const rd = a[2] === undefined || a[2] === '' ? [] : asRange(a[2]).flat().map(toStr);
      const ignore = optB(a[3], false), caseless = optN(a[4], 0) === 1, pad = a.length > 5 ? a[5] : ERR.NA;
      const split = (s, ds) => {
        if (!ds.length) return [s];
        const out = []; let cur = '', i = 0;
        const t = caseless ? s.toLowerCase() : s, dd = ds.map(d => (caseless ? d.toLowerCase() : d)).filter(Boolean);
        while (i < s.length) { const d = dd.find(x => t.startsWith(x, i)); if (d) { out.push(cur); cur = ''; i += d.length; } else cur += s[i++]; }
        out.push(cur);
        return ignore ? out.filter(x => x !== '') : out;
      };
      const rows = split(text, rd).map(line => split(line, cd));
      const w = Math.max(...rows.map(r => r.length));
      return arrOut(rows.map(r => [...r, ...Array(w - r.length).fill(pad)]));
    },
    MMULT: a => {
      const x = rowsOf(a[0]).map(r => r.map(toNum)), y = rowsOf(a[1]).map(r => r.map(toNum));
      if (x[0].length !== y.length) throw ERR.VALUE;
      return new Range(x.map(r => y[0].map((_, j) => r.reduce((s, v, k) => s + v * y[k][j], 0))));
    },
    MINVERSE: a => { const { inv } = minverse(rowsOf(a[0])); if (!inv) throw ERR.NUM; return new Range(inv); },
    MDETERM: a => minverse(rowsOf(a[0])).det,
    FREQUENCY: a => {
      const data = nums([a[0]]), bins = nums([a[1]]).sort((x, y) => x - y);
      const out = new Array(bins.length + 1).fill(0);
      for (const v of data) { let i = bins.findIndex(b => v <= b); if (i < 0) i = bins.length; out[i]++; }
      return new Range(out.map(v => [v]));
    },
    'MODE.MULT': a => {
      const c = new Map(); let best = 1;
      for (const n of nums(a)) { const k = (c.get(n) || 0) + 1; c.set(n, k); best = Math.max(best, k); }
      const out = [...c].filter(([, k]) => k === best && best > 1).map(([n]) => [n]);
      if (!out.length) throw ERR.NA;
      return new Range(out);
    },
    XMATCH: a => xmatch(a[0], asRange(a[1]).flat(), optN(a[2], 0), optN(a[3], 1)) + 1,
    XLOOKUP: a => {
      const look = asRange(a[1]), ret = asRange(a[2]);
      const vertical = look.w === 1;
      let i;
      try { i = xmatch(a[0], look.flat(), optN(a[4], 0), optN(a[5], 1)); } catch (e) {
        if (e === ERR.NA && a.length > 3 && a[3] !== undefined && a[3] !== '') return a[3];
        throw e;
      }
      if (vertical) { if (ret.h !== look.h) throw ERR.VALUE; return ret.w === 1 ? ret.rows[i][0] : new Range([ret.rows[i]]); }
      if (ret.w !== look.w) throw ERR.VALUE;
      return ret.h === 1 ? ret.rows[0][i] : new Range(ret.rows.map(r => [r[i]]));
    },
    INDEX: a => {
      const r = asRange(a[0]).rows;
      let row = toNum(a[1]), col = a[2] === undefined ? 1 : toNum(a[2]);
      if (r.length === 1 && a[2] === undefined) { col = row; row = 1; }
      if (row < 0 || col < 0 || row > r.length || col > r[0].length) throw ERR.REF;
      if (row === 0 && col === 0) return new Range(r);
      if (row === 0) return new Range(r.map(x => [x[col - 1]]));
      if (col === 0) return new Range([r[row - 1]]);
      return r[row - 1][col - 1];
    },
    LOOKUP: a => {
      const v = scalar(a[0]), look = asRange(a[1]).flat(), res = a[2] !== undefined ? asRange(a[2]).flat() : look;
      let hit = -1;
      look.forEach((x, i) => { if (!isBlank(x) && rankType(x) === rankType(v) && sortCmp(x, v) <= 0) hit = i; });
      if (hit < 0 || hit >= res.length) throw ERR.NA;
      return res[hit];
    },
    ADDRESS: a => {
      const r = Math.trunc(toNum(a[0])), c = Math.trunc(toNum(a[1])), abs = optN(a[2], 1), a1 = optB(a[3], true);
      if (r < 1 || c < 1 || c > 16384 || r > 1048576) throw ERR.VALUE;
      const sheet = a[4] !== undefined && a[4] !== '' ? toStr(a[4]) : '';
      const pre = sheet ? (/^[A-Za-z_][\w.]*$/.test(sheet) ? sheet : `'${sheet.replace(/'/g, "''")}'`) + '!' : '';
      if (!a1) return pre + `R${abs === 1 || abs === 2 ? r : `[${r}]`}C${abs === 1 || abs === 3 ? c : `[${c}]`}`;
      return pre + (abs === 1 || abs === 3 ? '$' : '') + indexToCol(c - 1) + (abs === 1 || abs === 2 ? '$' : '') + r;
    },
    SUBTOTAL: a => {
      const f = Math.trunc(toNum(a[0])) % 100;
      const map = { 1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN', 6: 'PRODUCT', 7: 'STDEV', 8: 'STDEVP', 9: 'SUM', 10: 'VAR', 11: 'VARP' };
      if (!map[f]) throw ERR.VALUE;
      return FN[map[f]](a.slice(1));
    },
    AGGREGATE: a => {
      const f = Math.trunc(toNum(a[0])), opt = Math.trunc(toNum(a[1]));
      const skipErr = [2, 3, 6, 7].includes(opt);
      const clean = x => (x instanceof Range ? new Range(x.rows.map(r => r.map(v => (skipErr && isErr(v) ? '' : v)))) : x);
      const names = { 1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN', 6: 'PRODUCT', 7: 'STDEV.S', 8: 'STDEV.P', 9: 'SUM', 10: 'VAR.S', 11: 'VAR.P', 12: 'MEDIAN', 13: 'MODE.SNGL', 14: 'LARGE', 15: 'SMALL', 16: 'PERCENTILE.INC', 17: 'QUARTILE.INC', 18: 'PERCENTILE.EXC', 19: 'QUARTILE.EXC' };
      if (!names[f]) throw ERR.VALUE;
      return f <= 13 ? FN[names[f]](a.slice(2).map(clean)) : FN[names[f]]([clean(a[2]), a[3]]);
    },

    // Statistics
    'STDEV.S': a => Math.sqrt(variance(nums(a), true)), 'STDEV.P': a => Math.sqrt(variance(nums(a), false)),
    'VAR.S': a => variance(nums(a), true), 'VAR.P': a => variance(nums(a), false), VARP: a => variance(nums(a), false),
    AVERAGEA: a => { const v = flatten(a).filter(x => !isBlank(x)).map(x => (typeof x === 'number' ? x : typeof x === 'boolean' ? +x : 0)); if (!v.length) throw ERR.DIV0; return mean(v); },
    MAXA: a => { const v = flatten(a).filter(x => !isBlank(x)).map(x => (typeof x === 'number' ? x : typeof x === 'boolean' ? +x : 0)); return v.length ? Math.max(...v) : 0; },
    MINA: a => { const v = flatten(a).filter(x => !isBlank(x)).map(x => (typeof x === 'number' ? x : typeof x === 'boolean' ? +x : 0)); return v.length ? Math.min(...v) : 0; },
    'MODE.SNGL': a => FN.MODE(a),
    'RANK.EQ': a => FN.RANK(a),
    'RANK.AVG': a => { const v = toNum(a[0]), n = nums([a[1]]), asc = optN(a[2], 0) !== 0; const k = n.filter(x => x === v).length; if (!k) throw ERR.NA; return 1 + n.filter(x => (asc ? x < v : x > v)).length + (k - 1) / 2; },
    PERCENTILE: a => percentile(nums([a[0]]), toNum(a[1]), false),
    'PERCENTILE.INC': a => percentile(nums([a[0]]), toNum(a[1]), false),
    'PERCENTILE.EXC': a => percentile(nums([a[0]]), toNum(a[1]), true),
    QUARTILE: a => { const q = Math.trunc(toNum(a[1])); if (q < 0 || q > 4) throw ERR.NUM; return percentile(nums([a[0]]), q / 4, false); },
    'QUARTILE.INC': a => FN.QUARTILE(a),
    'QUARTILE.EXC': a => { const q = Math.trunc(toNum(a[1])); if (q < 1 || q > 3) throw ERR.NUM; return percentile(nums([a[0]]), q / 4, true); },
    'PERCENTRANK.INC': a => {
      const n = nums([a[0]]).sort((x, y) => x - y), x = toNum(a[1]), sig = Math.trunc(optN(a[2], 3));
      if (!n.length || x < n[0] || x > n[n.length - 1] || sig < 1) throw ERR.NA;
      let r;
      const lo = n.filter(v => v < x).length;
      if (n.includes(x)) r = lo / (n.length - 1);
      else { const i = lo - 1; r = (i + (x - n[i]) / (n[i + 1] - n[i])) / (n.length - 1); }
      const f = 10 ** sig;
      return Math.floor(r * f + 1e-9) / f;
    },
    PERCENTRANK: a => FN['PERCENTRANK.INC'](a),
    GEOMEAN: a => { const n = nums(a); if (!n.length || n.some(x => x <= 0)) throw ERR.NUM; return Math.exp(n.reduce((s, x) => s + Math.log(x), 0) / n.length); },
    HARMEAN: a => { const n = nums(a); if (!n.length || n.some(x => x <= 0)) throw ERR.NUM; return n.length / n.reduce((s, x) => s + 1 / x, 0); },
    TRIMMEAN: a => { const n = nums([a[0]]).sort((x, y) => x - y), p = toNum(a[1]); if (p < 0 || p >= 1) throw ERR.NUM; const k = Math.floor(n.length * p / 2); const m = n.slice(k, n.length - k); if (!m.length) throw ERR.NUM; return mean(m); },
    DEVSQ: a => sqDev(nums(a)),
    AVEDEV: a => { const n = nums(a); if (!n.length) throw ERR.NUM; const m = mean(n); return mean(n.map(x => Math.abs(x - m))); },
    SKEW: a => { const n = nums(a), k = n.length; if (k < 3) throw ERR.DIV0; const m = mean(n), s = Math.sqrt(variance(n, true)); if (!s) throw ERR.DIV0; return k / ((k - 1) * (k - 2)) * n.reduce((t, x) => t + ((x - m) / s) ** 3, 0); },
    KURT: a => { const n = nums(a), k = n.length; if (k < 4) throw ERR.DIV0; const m = mean(n), s = Math.sqrt(variance(n, true)); if (!s) throw ERR.DIV0; return k * (k + 1) / ((k - 1) * (k - 2) * (k - 3)) * n.reduce((t, x) => t + ((x - m) / s) ** 4, 0) - 3 * (k - 1) ** 2 / ((k - 2) * (k - 3)); },
    CORREL: a => { const p = pairs(a[0], a[1]); if (p.length < 2) throw ERR.DIV0; const mx = mean(p.map(q => q[0])), my = mean(p.map(q => q[1])); let sxy = 0, sxx = 0, syy = 0; for (const [x, y] of p) { sxy += (x - mx) * (y - my); sxx += (x - mx) ** 2; syy += (y - my) ** 2; } if (!sxx || !syy) throw ERR.DIV0; return sxy / Math.sqrt(sxx * syy); },
    PEARSON: a => FN.CORREL(a),
    RSQ: a => FN.CORREL(a) ** 2,
    'COVARIANCE.P': a => { const p = pairs(a[0], a[1]); if (!p.length) throw ERR.DIV0; const mx = mean(p.map(q => q[0])), my = mean(p.map(q => q[1])); return p.reduce((s, [x, y]) => s + (x - mx) * (y - my), 0) / p.length; },
    'COVARIANCE.S': a => { const p = pairs(a[0], a[1]); if (p.length < 2) throw ERR.DIV0; const mx = mean(p.map(q => q[0])), my = mean(p.map(q => q[1])); return p.reduce((s, [x, y]) => s + (x - mx) * (y - my), 0) / (p.length - 1); },
    COVAR: a => FN['COVARIANCE.P'](a),
    SLOPE: a => { const p = pairs(a[1], a[0]); const mx = mean(p.map(q => q[0])), my = mean(p.map(q => q[1])); let sxy = 0, sxx = 0; for (const [x, y] of p) { sxy += (x - mx) * (y - my); sxx += (x - mx) ** 2; } if (!sxx) throw ERR.DIV0; return sxy / sxx; },
    INTERCEPT: a => { const p = pairs(a[1], a[0]); if (!p.length) throw ERR.DIV0; return mean(p.map(q => q[1])) - FN.SLOPE(a) * mean(p.map(q => q[0])); },
    FORECAST: a => FN.INTERCEPT([a[1], a[2]]) + FN.SLOPE([a[1], a[2]]) * toNum(a[0]),
    'FORECAST.LINEAR': a => FN.FORECAST(a),
    STANDARDIZE: a => { const sd = toNum(a[2]); if (sd <= 0) throw ERR.NUM; return (toNum(a[0]) - toNum(a[1])) / sd; },
    'NORM.DIST': a => { const x = toNum(a[0]), m = toNum(a[1]), sd = toNum(a[2]); if (sd <= 0) throw ERR.NUM; const z = (x - m) / sd; return toBool(a[3]) ? normCdf(z) : normPdf(z) / sd; },
    NORMDIST: a => FN['NORM.DIST'](a),
    'NORM.S.DIST': a => { const z = toNum(a[0]); return a[1] === undefined || toBool(a[1]) ? normCdf(z) : normPdf(z); },
    NORMSDIST: a => normCdf(toNum(a[0])),
    'NORM.INV': a => { const sd = toNum(a[2]); if (sd <= 0) throw ERR.NUM; return toNum(a[1]) + sd * normInv(toNum(a[0])); },
    NORMINV: a => FN['NORM.INV'](a),
    'NORM.S.INV': a => normInv(toNum(a[0])),
    NORMSINV: a => normInv(toNum(a[0])),
    'CONFIDENCE.NORM': a => { const alpha = toNum(a[0]), sd = toNum(a[1]), n = toNum(a[2]); if (alpha <= 0 || alpha >= 1 || sd <= 0 || n < 1) throw ERR.NUM; return normInv(1 - alpha / 2) * sd / Math.sqrt(n); },
    CONFIDENCE: a => FN['CONFIDENCE.NORM'](a),
    'EXPON.DIST': a => { const x = toNum(a[0]), l = toNum(a[1]); if (x < 0 || l <= 0) throw ERR.NUM; return toBool(a[2]) ? 1 - Math.exp(-l * x) : l * Math.exp(-l * x); },
    POISSON: a => FN['POISSON.DIST'](a),
    DAYS360: a => {
      // 30-day months: US (NASD) method, or European with TRUE.
      const [y1, m1, d1a] = excelYMD(toNum(a[0])), [y2, m2, d2a] = excelYMD(toNum(a[1]));
      let d1 = d1a, d2 = d2a;
      if (a[2] !== undefined && toBool(a[2])) { if (d1 === 31) d1 = 30; if (d2 === 31) d2 = 30; }
      else {
        const lastFeb = (y, m, d) => m === 2 && d === new Date(Date.UTC(y, 2, 0)).getUTCDate();
        if (lastFeb(y1, m1, d1) && lastFeb(y2, m2, d2)) d2 = 30;
        if (lastFeb(y1, m1, d1)) d1 = 30;
        if (d2 === 31 && d1 >= 30) d2 = 30;
        if (d1 === 31) d1 = 30;
      }
      return (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1);
    },
    'POISSON.DIST': a => {
      const x = Math.trunc(toNum(a[0])), m = toNum(a[1]);
      if (x < 0 || m < 0) throw ERR.NUM;
      const pmf = k => Math.exp(k * Math.log(m) - m - lgamma(k + 1));
      if (!toBool(a[2])) return m === 0 ? (x === 0 ? 1 : 0) : pmf(x);
      let s = 0; for (let k = 0; k <= x; k++) s += m === 0 ? (k === 0 ? 1 : 0) : pmf(k); return Math.min(1, s);
    },
    'BINOM.DIST': a => {
      const k = Math.trunc(toNum(a[0])), n = Math.trunc(toNum(a[1])), p = toNum(a[2]);
      if (k < 0 || k > n || p < 0 || p > 1) throw ERR.NUM;
      const pmf = j => (p === 0 ? (j === 0 ? 1 : 0) : p === 1 ? (j === n ? 1 : 0) : Math.exp(lnComb(n, j) + j * Math.log(p) + (n - j) * Math.log(1 - p)));
      if (!toBool(a[3])) return pmf(k);
      let s = 0; for (let j = 0; j <= k; j++) s += pmf(j); return Math.min(1, s);
    },

    // Finance
    IRR: a => { const v = nums([a[0]]); if (!v.some(x => x > 0) || !v.some(x => x < 0)) throw ERR.NUM; return newton(r => v.reduce((s, x, i) => s + x / Math.pow(1 + r, i), 0), optN(a[1], 0.1)); },
    NPER: a => {
      const r = toNum(a[0]), pmt = toNum(a[1]), pv = toNum(a[2]), fv = optN(a[3], 0), t = optN(a[4], 0);
      if (r === 0) { if (pmt === 0) throw ERR.NUM; return -(pv + fv) / pmt; }
      const num = pmt * (1 + r * t) - fv * r, den = pmt * (1 + r * t) + pv * r;
      if (num / den <= 0) throw ERR.NUM;
      return Math.log(num / den) / Math.log(1 + r);
    },
    RATE: a => {
      const n = toNum(a[0]), pmt = toNum(a[1]), pv = toNum(a[2]), fv = optN(a[3], 0), t = optN(a[4], 0);
      return newton(r => (Math.abs(r) < 1e-12 ? pv + pmt * n + fv : pv * Math.pow(1 + r, n) + pmt * (1 + r * t) * (Math.pow(1 + r, n) - 1) / r + fv), optN(a[5], 0.1));
    },
    IPMT: a => { const per = toNum(a[1]), n = toNum(a[2]); if (per < 1 || per > n) throw ERR.NUM; return ipmt(toNum(a[0]), per, n, toNum(a[3]), optN(a[4], 0), optN(a[5], 0)); },
    PPMT: a => { const r = toNum(a[0]), per = toNum(a[1]), n = toNum(a[2]), pv = toNum(a[3]), fv = optN(a[4], 0), t = optN(a[5], 0); if (per < 1 || per > n) throw ERR.NUM; return FN.PMT([r, n, pv, fv, t]) - ipmt(r, per, n, pv, fv, t); },
    CUMIPMT: a => { const [r, n, pv, s, e, t] = a.map(toNum); if (r <= 0 || n <= 0 || pv <= 0 || s < 1 || e < s || e > n || (t !== 0 && t !== 1)) throw ERR.NUM; let sum = 0; for (let p = Math.ceil(s); p <= Math.floor(e); p++) sum += ipmt(r, p, n, pv, 0, t); return sum; },
    CUMPRINC: a => { const [r, n, pv, s, e, t] = a.map(toNum); if (r <= 0 || n <= 0 || pv <= 0 || s < 1 || e < s || e > n || (t !== 0 && t !== 1)) throw ERR.NUM; const pmt = FN.PMT([r, n, pv, 0, t]); let sum = 0; for (let p = Math.ceil(s); p <= Math.floor(e); p++) sum += pmt - ipmt(r, p, n, pv, 0, t); return sum; },
    XNPV: a => { const r = toNum(a[0]), v = nums([a[1]]), d = nums([a[2]]); if (v.length !== d.length) throw ERR.NUM; return v.reduce((s, x, i) => s + x / Math.pow(1 + r, (d[i] - d[0]) / 365), 0); },
    XIRR: a => { const v = nums([a[0]]), d = nums([a[1]]); if (v.length !== d.length || !v.some(x => x > 0) || !v.some(x => x < 0)) throw ERR.NUM; return newton(r => v.reduce((s, x, i) => s + x / Math.pow(1 + r, (d[i] - d[0]) / 365), 0), optN(a[2], 0.1)); },
    SLN: a => { const life = toNum(a[2]); if (!life) throw ERR.DIV0; return (toNum(a[0]) - toNum(a[1])) / life; },
    SYD: a => { const [cost, salv, life, per] = a.map(toNum); if (life <= 0 || per < 1 || per > life) throw ERR.NUM; return (cost - salv) * (life - per + 1) * 2 / (life * (life + 1)); },
    DDB: a => {
      const cost = toNum(a[0]), salv = toNum(a[1]), life = toNum(a[2]), per = toNum(a[3]), f = optN(a[4], 2);
      if (life <= 0 || per <= 0 || per > life || f <= 0 || cost < 0 || salv < 0) throw ERR.NUM;
      let acc = 0, dep = 0;
      for (let p = 1; p <= per; p++) { dep = Math.max(0, Math.min((cost - acc) * f / life, cost - salv - acc)); acc += dep; }
      return dep;
    },
    EFFECT: a => { const r = toNum(a[0]), n = Math.trunc(toNum(a[1])); if (r <= 0 || n < 1) throw ERR.NUM; return Math.pow(1 + r / n, n) - 1; },
    NOMINAL: a => { const r = toNum(a[0]), n = Math.trunc(toNum(a[1])); if (r <= 0 || n < 1) throw ERR.NUM; return n * (Math.pow(1 + r, 1 / n) - 1); },

    // Math
    SUMSQ: a => nums(a).reduce((s, x) => s + x * x, 0),
    SUMX2MY2: a => pairSum(a, (x, y) => x * x - y * y),
    SUMX2PY2: a => pairSum(a, (x, y) => x * x + y * y),
    SUMXMY2: a => pairSum(a, (x, y) => (x - y) * (x - y)),
    MROUND: a => { const n = toNum(a[0]), m = toNum(a[1]); if (m === 0) return 0; if (Math.sign(n) * Math.sign(m) < 0) throw ERR.NUM; return roundTo(Math.round(n / m + Number.EPSILON * Math.sign(n / m)) * m, 10); },
    QUOTIENT: a => { const d = toNum(a[1]); if (d === 0) throw ERR.DIV0; return Math.trunc(toNum(a[0]) / d); },
    GCD: a => { const n = nums(a).map(Math.trunc); if (n.some(x => x < 0)) throw ERR.NUM; const g = (x, y) => (y ? g(y, x % y) : x); return n.reduce(g, 0); },
    LCM: a => { const n = nums(a).map(Math.trunc); if (n.some(x => x < 0)) throw ERR.NUM; const g = (x, y) => (y ? g(y, x % y) : x); return n.reduce((l, x) => (l === 0 || x === 0 ? 0 : l / g(l, x) * x), 1); },
    FACT: a => { const n = Math.trunc(toNum(a[0])); if (n < 0) throw ERR.NUM; if (n > 170) throw ERR.NUM; let f = 1; for (let i = 2; i <= n; i++) f *= i; return f; },
    COMBIN: a => { const n = Math.trunc(toNum(a[0])), k = Math.trunc(toNum(a[1])); if (n < 0 || k < 0 || k > n) throw ERR.NUM; return Math.round(Math.exp(lnComb(n, k))); },
    PERMUT: a => { const n = Math.trunc(toNum(a[0])), k = Math.trunc(toNum(a[1])); if (n < 0 || k < 0 || k > n) throw ERR.NUM; let p = 1; for (let i = 0; i < k; i++) p *= n - i; return p; },
    ODD: a => { const n = toNum(a[0]), s = n < 0 ? -1 : 1; let v = Math.ceil(Math.abs(n)); if (v % 2 === 0) v++; return s * v; },
    EVEN: a => { const n = toNum(a[0]), s = n < 0 ? -1 : 1; let v = Math.ceil(Math.abs(n)); if (v % 2) v++; return s * v; },
    ASIN: a => { const x = toNum(a[0]); if (x < -1 || x > 1) throw ERR.NUM; return Math.asin(x); },
    ACOS: a => { const x = toNum(a[0]); if (x < -1 || x > 1) throw ERR.NUM; return Math.acos(x); },
    ATAN: a => Math.atan(toNum(a[0])),
    ATAN2: a => { const x = toNum(a[0]), y = toNum(a[1]); if (!x && !y) throw ERR.DIV0; return Math.atan2(y, x); },
    SINH: a => Math.sinh(toNum(a[0])), COSH: a => Math.cosh(toNum(a[0])), TANH: a => Math.tanh(toNum(a[0])),
    ASINH: a => Math.asinh(toNum(a[0])), ACOSH: a => { const x = toNum(a[0]); if (x < 1) throw ERR.NUM; return Math.acosh(x); }, ATANH: a => { const x = toNum(a[0]); if (Math.abs(x) >= 1) throw ERR.NUM; return Math.atanh(x); },
    SQRTPI: a => { const x = toNum(a[0]); if (x < 0) throw ERR.NUM; return Math.sqrt(x * Math.PI); },
    'CEILING.MATH': a => { const n = toNum(a[0]), s = Math.abs(optN(a[1], 1)), m = optN(a[2], 0); if (!s) return 0; return n >= 0 || !m ? Math.ceil(n / s) * s : -Math.ceil(-n / s) * s; },
    'FLOOR.MATH': a => { const n = toNum(a[0]), s = Math.abs(optN(a[1], 1)), m = optN(a[2], 0); if (!s) return 0; return n >= 0 || !m ? Math.floor(n / s) * s : -Math.floor(-n / s) * s; },
    'CEILING.PRECISE': a => { const s = Math.abs(optN(a[1], 1)); return s ? Math.ceil(toNum(a[0]) / s) * s : 0; },
    'FLOOR.PRECISE': a => { const s = Math.abs(optN(a[1], 1)); return s ? Math.floor(toNum(a[0]) / s) * s : 0; },
    ROMAN: a => { let n = Math.trunc(toNum(a[0])); if (n < 0 || n > 3999) throw ERR.VALUE; let s = ''; for (const [v, r] of ROMAN) while (n >= v) { s += r; n -= v; } return s; },
    ARABIC: a => {
      const s = toStr(a[0]).trim().toUpperCase(), neg = s[0] === '-', t = neg ? s.slice(1) : s;
      if (!/^[MDCLXVI]*$/.test(t)) throw ERR.VALUE;
      const val = { M: 1000, D: 500, C: 100, L: 50, X: 10, V: 5, I: 1 };
      let n = 0; for (let i = 0; i < t.length; i++) { const v = val[t[i]], nx = val[t[i + 1]] || 0; n += v < nx ? -v : v; }
      return neg ? -n : n;
    },
    BASE: a => { const n = Math.trunc(toNum(a[0])), r = Math.trunc(toNum(a[1])); if (n < 0 || r < 2 || r > 36) throw ERR.NUM; return n.toString(r).toUpperCase().padStart(Math.trunc(optN(a[2], 0)), '0'); },
    DECIMAL: a => { const r = Math.trunc(toNum(a[1])), s = toStr(a[0]).trim(); if (r < 2 || r > 36 || !s) throw ERR.NUM; const n = parseInt(s, r); if (isNaN(n) || n.toString(r).toUpperCase() !== s.toUpperCase().replace(/^0+(?=.)/, '')) throw ERR.NUM; return n; },
    DEC2BIN: a => toBase(a[0], 2, 10, a[1]), DEC2OCT: a => toBase(a[0], 8, 30, a[1]), DEC2HEX: a => toBase(a[0], 16, 40, a[1]),
    BIN2DEC: a => fromBase(a[0], 2, 10), OCT2DEC: a => fromBase(a[0], 8, 30), HEX2DEC: a => fromBase(a[0], 16, 40),

    // Text
    TEXTBEFORE: a => {
      const t = toStr(a[0]), ds = asRange(a[1]).flat().map(toStr), inst = Math.trunc(optN(a[2], 1));
      if (inst === 0 || Math.abs(inst) > t.length + 1) throw ERR.VALUE;
      const h = textSplitAt(t, ds, inst, optN(a[3], 0) === 1, optN(a[4], 0) === 1);
      if (!h) { if (a.length > 5) return a[5]; throw ERR.NA; }
      return t.slice(0, h[0]);
    },
    TEXTAFTER: a => {
      const t = toStr(a[0]), ds = asRange(a[1]).flat().map(toStr), inst = Math.trunc(optN(a[2], 1));
      if (inst === 0 || Math.abs(inst) > t.length + 1) throw ERR.VALUE;
      const h = textSplitAt(t, ds, inst, optN(a[3], 0) === 1, optN(a[4], 0) === 1);
      if (!h) { if (a.length > 5) return a[5]; throw ERR.NA; }
      return t.slice(h[1]);
    },
    CLEAN: a => toStr(a[0]).replace(/[\u0000-\u001f]/g, ''),
    UNICHAR: a => { const n = Math.trunc(toNum(a[0])); if (n < 1 || n > 0x10ffff) throw ERR.VALUE; return String.fromCodePoint(n); },
    UNICODE: a => { const s = toStr(a[0]); if (!s) throw ERR.VALUE; return s.codePointAt(0); },
    NUMBERVALUE: a => {
      const dec = a[1] === undefined ? '.' : toStr(a[1])[0] || '.', grp = a[2] === undefined ? ',' : toStr(a[2])[0] || ',';
      let s = toStr(a[0]).replace(/\s/g, '');
      if (!s) return 0;
      const pct = (s.match(/%+$/) || [''])[0].length;
      s = s.replace(/%+$/, '').split(grp).join('').replace(dec, '.');
      const n = Number(s);
      if (s === '' || isNaN(n)) throw ERR.VALUE;
      return n / 100 ** pct;
    },

    // Date & time
    NETWORKDAYS: a => networkdays(toNum(a[0]), toNum(a[1]), '0000011', holidays(a[2])),
    'NETWORKDAYS.INTL': a => networkdays(toNum(a[0]), toNum(a[1]), weekendMask(a[2]), holidays(a[3])),
    WORKDAY: a => workday(toNum(a[0]), toNum(a[1]), '0000011', holidays(a[2])),
    'WORKDAY.INTL': a => workday(toNum(a[0]), toNum(a[1]), weekendMask(a[2]), holidays(a[3])),
    DATEDIF: a => {
      const s = Math.floor(toNum(a[0])), e = Math.floor(toNum(a[1])), u = toStr(a[2]).toUpperCase();
      if (s > e) throw ERR.NUM;
      const [y1, m1, d1] = ymd(s), [y2, m2, d2] = ymd(e);
      let months = (y2 - y1) * 12 + (m2 - m1); if (d2 < d1) months--;
      switch (u) {
        case 'D': return e - s;
        case 'M': return months;
        case 'Y': return Math.floor(months / 12);
        case 'YM': return months % 12;
        case 'MD': { if (d2 >= d1) return d2 - d1; const prevMonthDays = new Date(Date.UTC(y2, m2 - 1, 0)).getUTCDate(); return prevMonthDays - d1 + d2; }
        case 'YD': { let start = Date.UTC(y2, m1 - 1, d1); if (start > Date.UTC(y2, m2 - 1, d2)) start = Date.UTC(y2 - 1, m1 - 1, d1); return Math.round((Date.UTC(y2, m2 - 1, d2) - start) / 86400000); }
      }
      throw ERR.NUM;
    },
    WEEKNUM: a => {
      const s = Math.floor(toNum(a[0])), t = optN(a[1], 1);
      if (t === 21) return FN.ISOWEEKNUM([s]);
      const start = { 1: 0, 17: 0, 2: 1, 11: 1, 12: 2, 13: 3, 14: 4, 15: 5, 16: 6 }[t];
      if (start === undefined) throw ERR.NUM;
      const [y] = ymd(s);
      const jan1 = Math.round((Date.UTC(y, 0, 1) - EPOCH) / 86400000);
      const off = (serialToDate(jan1).getUTCDay() - start + 7) % 7;
      return Math.floor((s - jan1 + off) / 7) + 1;
    },
    ISOWEEKNUM: a => {
      const d = serialToDate(Math.floor(toNum(a[0])));
      const t = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
      t.setUTCDate(t.getUTCDate() + 4 - (t.getUTCDay() || 7));
      return Math.ceil(((t - Date.UTC(t.getUTCFullYear(), 0, 1)) / 86400000 + 1) / 7);
    },
    YEARFRAC: a => {
      let s = Math.floor(toNum(a[0])), e = Math.floor(toNum(a[1]));
      const basis = Math.trunc(optN(a[2], 0));
      if (s > e) [s, e] = [e, s];
      let [y1, m1, d1] = ymd(s), [y2, m2, d2] = ymd(e);
      switch (basis) {
        case 0: { // US (NASD) 30/360
          const lastFeb = (y, m, d) => m === 2 && d === new Date(Date.UTC(y, 2, 0)).getUTCDate();
          if (lastFeb(y1, m1, d1) && lastFeb(y2, m2, d2)) d2 = 30;
          if (lastFeb(y1, m1, d1)) d1 = 30;
          if (d2 === 31 && d1 >= 30) d2 = 30;
          if (d1 === 31) d1 = 30;
          return ((y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)) / 360;
        }
        case 1: {
          const leap = y => (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
          if (y1 === y2) return (e - s) / (leap(y1) ? 366 : 365);
          const oneYear = y2 === y1 + 1 && (m2 < m1 || (m2 === m1 && d2 <= d1));
          if (oneYear) { const feb29 = (leap(y1) && (m1 < 3)) || (leap(y2) && (m2 > 2 || (m2 === 2 && d2 === 29))); return (e - s) / (feb29 ? 366 : 365); }
          let days = 0; for (let y = y1; y <= y2; y++) days += leap(y) ? 366 : 365;
          return (e - s) / (days / (y2 - y1 + 1));
        }
        case 2: return (e - s) / 360;
        case 3: return (e - s) / 365;
        case 4: { if (d1 === 31) d1 = 30; if (d2 === 31) d2 = 30; return ((y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)) / 360; }
      }
      throw ERR.NUM;
    },
    TIMEVALUE: a => {
      const m = /^\s*(\d{1,2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?\s*(am|pm)?\s*$/i.exec(toStr(a[0]));
      if (!m) throw ERR.VALUE;
      let h = +m[1];
      if (m[4]) { if (h > 12) throw ERR.VALUE; h = h % 12 + (m[4].toLowerCase() === 'pm' ? 12 : 0); }
      return ((h * 3600 + +m[2] * 60 + +(m[3] || 0)) / 86400) % 1;
    },
    ISNONTEXT: a => typeof scalar(a[0]) !== 'string' || scalar(a[0]) === '',
    HYPERLINK: a => (a[1] !== undefined ? scalar(a[1]) : toStr(a[0]))
  });

  // ---------- Excel file names for newer functions ----------
  // Functions added after Excel 2007 are stored with an _xlfn. prefix (FILTER and SORT also
  // _xlws.), and LET's names with _xlpm.; without them Excel shows #NAME?.
  const XLFN = new Set(('AGGREGATE ARABIC BASE BITAND BITLSHIFT BITOR BITRSHIFT BITXOR CEILING.MATH CEILING.PRECISE CHOOSECOLS CHOOSEROWS COMBINA ' +
    'CONCAT CONFIDENCE.NORM CONFIDENCE.T COT COTH COVARIANCE.P COVARIANCE.S CSC CSCH DAYS DECIMAL DROP ERF.PRECISE ERFC.PRECISE EXPAND EXPON.DIST ' +
    'FILTER FLOOR.MATH FLOOR.PRECISE FORECAST.LINEAR FORMULATEXT HSTACK IFNA IFS ISFORMULA ISOWEEKNUM LET MAXIFS MINIFS MODE.MULT MODE.SNGL ' +
    'NETWORKDAYS.INTL NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV NUMBERVALUE PERCENTILE.EXC PERCENTILE.INC PERCENTRANK.INC POISSON.DIST QUARTILE.EXC ' +
    'QUARTILE.INC RANDARRAY RANK.AVG RANK.EQ SEC SECH SEQUENCE SHEET SHEETS SORT SORTBY STDEV.P STDEV.S SWITCH TAKE TEXTAFTER TEXTBEFORE TEXTJOIN ' +
    'TEXTSPLIT TOCOL TOROW UNICHAR UNICODE UNIQUE VAR.P VAR.S VSTACK WORKDAY.INTL WRAPCOLS WRAPROWS XLOOKUP XMATCH XOR BINOM.DIST').split(' '));
  const XLWS = new Set(['FILTER', 'SORT']);
  /** Rewrites the identifiers of a formula (without '='), leaving strings and quoted sheet names alone. */
  function mapIdentifiers(f, fn) {
    let out = '', i = 0;
    while (i < f.length) {
      const ch = f[i];
      if (ch === '"' || ch === "'") { let j = i + 1; while (j < f.length && !(f[j] === ch && f[j + 1] !== ch)) j += f[j] === ch ? 2 : 1; out += f.slice(i, j + 1); i = j + 1; continue; }
      if (ch === '[') {
        let j = i, depth = 0;
        for (; j < f.length; j++) { if (f[j] === "'") { j++; continue; } if (f[j] === '[') depth++; else if (f[j] === ']' && --depth === 0) break; }
        out += fn(f.slice(i, j + 1), 'spec');
        i = j + 1;
        continue;
      }
      const m = /^[A-Za-z_][\w.]*/.exec(f.slice(i));
      if (m && !/[\w.$]/.test(f[i - 1] || '')) {
        const rest = f.slice(i + m[0].length);
        const kind = /^\s*\(/.test(rest) ? 'fn' : /^!/.test(rest) ? 'sheet' : /^\[/.test(rest) ? 'table' : /^[A-Za-z]{1,3}\d+$/.test(m[0]) || /^\$?\d/.test(rest) || /^(TRUE|FALSE)$/i.test(m[0]) ? 'other' : 'name';
        out += fn(m[0], kind);
        i += m[0].length;
        continue;
      }
      out += ch; i++;
    }
    return out;
  }
  /**
   * A formula as Excel stores it. opts.names: workbook and table names (kept as they are; other
   * names are LET's); opts.table: the table the cell is in, for [@Col] and [Col].
   */
  const toExcelFormula = (f, opts = {}) => {
    let prevTable = null;
    return mapIdentifiers(f, (id, kind) => {
      const up = id.toUpperCase();
      if (kind === 'table') { prevTable = id; return id; }
      if (kind === 'spec') {
        const t = prevTable; prevTable = null;
        // Excel writes "this row" in full, and always with the table's name.
        const m = /^\[@\s*(\[?)([\s\S]*?)\]?\]$/.exec(id);
        const name = t ? '' : opts.table || '';
        if (m) return name + (m[2] ? `[[#This Row],[${m[2]}]]` : '[#This Row]');
        return name + id;
      }
      prevTable = null;
      if (kind === 'fn' && XLFN.has(up)) return '_xlfn.' + (XLWS.has(up) ? '_xlws.' : '') + id;
      if (kind === 'name' && !(opts.names && opts.names.has(up))) return '_xlpm.' + id;
      return id;
    });
  };
  const fromExcelFormula = f => mapIdentifiers(f, (id, kind) => (kind === 'spec'
    ? id.replace(/^\[\[#This Row\],\s*\[((?:[^\]']|'.)*)\]\]$/i, '[@[$1]]').replace(/^\[@\[((?:[A-Za-z_][\w.]*))\]\]$/, '[@$1]').replace(/^\[#This Row\]$/i, '[@]')
    : id.replace(/^_xlfn\./i, '').replace(/^_xlws\./i, '').replace(/^_xlpm\./i, '')));

  function lookup(a, horizontal) {
    const v = scalar(a[0]);
    let rows = asRange(a[1]).rows;
    if (horizontal) rows = rows[0].map((_, c) => rows.map(r => r[c]));
    const idx = toNum(a[2]);
    const approx = a[3] === undefined ? true : toBool(a[3]);
    if (idx < 1 || idx > rows[0].length) throw ERR.REF;
    let hit = -1;
    if (!approx) {
      const m = typeof v === 'string' ? criteria(v) : (x => eqLoose(x, v));
      hit = rows.findIndex(r => m(r[0]));
    } else {
      rows.forEach((r, i) => { if (!isBlank(r[0]) && compare(r[0], v) <= 0) hit = i; });
    }
    if (hit < 0) throw ERR.NA;
    return rows[hit][idx - 1];
  }

  // ---------- Number formatting ----------
  // Excel's number format codes: up to four sections (positive;negative;zero;text) with
  // conditions, digit placeholders (0 # ?), grouping and scaling commas, percent, scientific,
  // fractions, literals ("text", \c, _x, $ - + ( ) : and other symbols), and dates and times.

  /** General as Excel shows a number: at most 11 characters (after SheetJS's SSF, Apache-2.0). */
  function generalText(v) {
    if (!isFinite(v)) return String(v);
    if (v === 0) return '0';
    const strip = o => (o.indexOf('.') === -1 ? o : o.replace(/(?:\.0*|(\.\d*[1-9])0+)$/, '$1'));
    const normExp = o => (o.indexOf('E') === -1 ? o : o.replace(/(?:\.0*|(\.\d*[1-9])0+)[Ee]/, '$1E').replace(/(E[+-])(\d)$/, '$10$2'));
    const V = Math.floor(Math.log(Math.abs(v)) * Math.LOG10E);
    const w = v < 0 ? 12 : 11;
    let o;
    if (V >= -4 && V <= -1) o = v.toPrecision(10 + V);
    else if (Math.abs(V) <= 9) {
      o = strip(v.toFixed(12));
      if (o.length > w) { o = v.toPrecision(10); if (o.length > w) o = v.toExponential(5); }
    } else if (V === 10) o = v.toFixed(10).substr(0, 12);
    else { o = strip(v.toFixed(11)); if (o.length > w || o === '0' || o === '-0') o = v.toPrecision(6); }
    return strip(normExp(o.toUpperCase()));
  }

  const fmtCache = new Map();
  /** A format code split into sections of tokens. */
  function compileFormat(fmt) {
    let c = fmtCache.get(fmt);
    if (c) return c;
    const sections = [];
    let toks = [], cond = null, i = 0;
    const lit = t => toks.push({ k: 'lit', v: t });
    while (i < fmt.length) {
      const ch = fmt[i];
      if (ch === ';') { sections.push({ toks, cond }); toks = []; cond = null; i++; continue; }
      if (ch === '"') { const j = fmt.indexOf('"', i + 1); lit(fmt.slice(i + 1, j < 0 ? fmt.length : j)); i = j < 0 ? fmt.length : j + 1; continue; }
      if (ch === '\\') { lit(fmt[i + 1] || ''); i += 2; continue; }
      if (ch === '_') { lit(' '); i += 2; continue; }
      if (ch === '*') { i += 2; continue; } // repeat-to-fill: nothing in text
      if (ch === '[') {
        const j = fmt.indexOf(']', i);
        const body = fmt.slice(i + 1, j < 0 ? fmt.length : j);
        i = j < 0 ? fmt.length : j + 1;
        let m;
        if ((m = /^(<=|>=|<>|<|>|=)\s*(-?[\d.]+(?:[eE][+-]?\d+)?)$/.exec(body))) cond = { op: m[1], v: +m[2] };
        else if ((m = /^\$([^-\]]*)/.exec(body))) { if (m[1]) lit(m[1]); }
        else if (/^(h+|m+|s+)$/i.test(body)) toks.push({ k: 'date', v: body.toLowerCase(), elapsed: true });
        continue; // colours and locale codes show nothing
      }
      if (ch === '0' || ch === '#' || ch === '?') { toks.push({ k: 'dig', v: ch }); i++; continue; }
      if (ch === '.') { toks.push({ k: 'dot' }); i++; continue; }
      if (ch === ',') { toks.push({ k: 'comma' }); i++; continue; }
      if (ch === '%') { toks.push({ k: 'pct' }); i++; continue; }
      if (ch === '/') { toks.push({ k: 'slash' }); i++; continue; }
      if (ch === '@') { toks.push({ k: 'text' }); i++; continue; }
      if ((ch === 'E' || ch === 'e') && (fmt[i + 1] === '+' || fmt[i + 1] === '-')) { toks.push({ k: 'exp', sign: fmt[i + 1] }); i += 2; continue; }
      let m;
      if ((m = /^(AM\/PM|A\/P)/i.exec(fmt.slice(i)))) { toks.push({ k: 'ampm', v: m[1] }); i += m[1].length; continue; }
      if ((m = /^(y+|m+|d+|h+|s+)/i.exec(fmt.slice(i)))) { toks.push({ k: 'date', v: m[1].toLowerCase() }); i += m[1].length; continue; }
      if (/^General/i.test(fmt.slice(i))) { toks.push({ k: 'general' }); i += 7; continue; }
      lit(ch); i++;
    }
    sections.push({ toks, cond });
    for (const s of sections) {
      s.isDate = s.toks.some(t => t.k === 'date' || t.k === 'ampm');
      if (s.isDate) {
        // Digits after seconds are fractions of a second; "m" next to h or s is minutes.
        const ts = s.toks;
        for (let a = 0; a < ts.length; a++) {
          if (ts[a].k === 'dot' && ts[a + 1] && ts[a + 1].k === 'dig' && ts[a + 1].v === '0') {
            let n = 0; while (ts[a + 1 + n] && ts[a + 1 + n].k === 'dig' && ts[a + 1 + n].v === '0') n++;
            ts.splice(a, n + 1, { k: 'frac', n });
          }
        }
        const dates = ts.filter(t => t.k === 'date');
        dates.forEach((t, j) => {
          if (t.v[0] !== 'm' || t.v.length > 2 || t.elapsed) return;
          const prev = dates[j - 1], next = dates[j + 1];
          if ((prev && prev.v[0] === 'h') || (next && next.v[0] === 's')) t.minute = true;
        });
        s.frac = (ts.find(t => t.k === 'frac') || { n: 0 }).n;
      }
    }
    c = sections;
    if (fmtCache.size > 2000) fmtCache.clear();
    fmtCache.set(fmt, c);
    return c;
  }

  /** Formats a value with an Excel number format code. */
  function formatValue(v, fmt) {
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v !== 'number') {
      if (typeof v === 'string' && fmt) {
        const secs = compileFormat(fmt);
        const s = secs.length >= 4 ? secs[3] : secs.find(x => x.toks.some(t => t.k === 'text'));
        if (s) return s.toks.map(t => (t.k === 'text' ? v : t.k === 'lit' ? t.v : '')).join('');
      }
      return v;
    }
    if (!fmt || /^general$/i.test(fmt) || fmt === '@') return generalText(v);
    const secs = compileFormat(fmt);
    const test = (c, x) => ({ '<': x < c.v, '>': x > c.v, '=': x === c.v, '<=': x <= c.v, '>=': x >= c.v, '<>': x !== c.v }[c.op]);
    let s, n = v, sign = '';
    if (secs.some(x => x.cond)) {
      s = secs.find(x => x.cond && test(x.cond, v)) || secs.find(x => !x.cond) || secs[secs.length - 1];
      if (v < 0 && !(s.cond && s.cond.v <= 0 && (s.cond.op === '<' || s.cond.op === '<='))) sign = '-';
      n = Math.abs(v);
    } else if (v < 0 && secs.length > 1 && secs[1].toks.length) { s = secs[1]; n = -v; }
    else if (v === 0 && secs.length > 2) s = secs[2];
    else { s = secs[0]; if (v < 0) { sign = '-'; n = -v; } }
    if (s.toks.some(t => t.k === 'general')) return sign + s.toks.map(t => (t.k === 'general' ? generalText(n) : t.k === 'lit' ? t.v : '')).join('');
    if (s.isDate) return formatDate(sign ? -n : n, s);
    return sign + formatNumber(n, s.toks);
  }

  /** Rounds half away from zero, as Excel does with the decimal digits it shows. */
  function roundShown(x, d) {
    if (d > 15) d = 15;
    // Excel works to 15 significant digits: 0.035 is 0.035, not 0.034999999999999996.
    x = Number(x.toPrecision(15));
    const r = Math.round(Number(x + 'e' + d));
    return isFinite(r) ? Number(r + 'e-' + d) : x;
  }
  function digitsOf(x) {
    // Plain digits, also for numbers JavaScript would print with an exponent.
    if (!isFinite(x)) return String(x);
    if (x < 1e21) return Math.floor(x).toFixed(0);
    return BigInt(Math.floor(x)).toString();
  }

  function formatNumber(n, toks) {
    const t = toks.filter(x => x.k !== 'lit' || x.v !== '');
    const pct = t.filter(x => x.k === 'pct').length;
    n *= Math.pow(100, pct);
    if (t.some(x => x.k === 'slash') && t.some(x => x.k === 'dig')) return formatFraction(n, t);
    const expAt = t.findIndex(x => x.k === 'exp');
    const digAt = t.map((x, i) => (x.k === 'dig' ? i : -1)).filter(i => i >= 0);
    if (!digAt.length) return t.map(x => (x.k === 'lit' ? x.v : x.k === 'pct' ? '%' : x.k === 'text' ? generalText(n) : '')).join('');
    const end = expAt >= 0 ? expAt : t.length;
    const dot = t.findIndex((x, i) => x.k === 'dot' && i < end);
    const lastDig = Math.max(...digAt.filter(i => i < end));
    // Commas: between digits before the point group thousands; after the last digit (or just
    // before the point) each divides by 1000; anywhere else they are just commas.
    const role = new Map();
    t.forEach((x, i) => {
      if (x.k !== 'comma' || i >= end) return;
      const before = digAt.some(j => j < i && (dot < 0 || j < dot));
      let j = i + 1; while (t[j] && t[j].k === 'comma') j++;
      const scale = (j > lastDig && i > lastDig - (dot >= 0 && dot < lastDig ? 0 : 0) && digAt.every(d => d < i)) || (t[j] && t[j].k === 'dot');
      if (before && scale) role.set(i, 'scale');
      else if (before && (dot < 0 || i < dot) && digAt.some(d => d > i && (dot < 0 || d < dot))) role.set(i, 'group');
      else role.set(i, 'lit');
    });
    const scales = [...role.values()].filter(r => r === 'scale').length;
    n /= Math.pow(1000, scales);
    const group = [...role.values()].includes('group');
    const intToks = [], decToks = [];
    t.forEach((x, i) => { if (x.k === 'dig' && i < end) (dot >= 0 && i > dot ? decToks : intToks).push(i); });
    let expText = '';
    if (expAt >= 0) {
      const expDigits = t.slice(expAt + 1).filter(x => x.k === 'dig');
      let e = 0;
      if (n !== 0) {
        e = Math.floor(Math.log10(n));
        const k = intToks.length || 1;
        if (k > 1 && intToks.some(i => t[i].v === '#')) e = Math.floor(e / k) * k;
        else e -= k - 1;
        n /= Math.pow(10, e);
        if (roundShown(n, decToks.length) >= Math.pow(10, k > 1 && intToks.some(i => t[i].v === '#') ? k : k)) { n /= 10; e += 1; }
      }
      const ed = String(Math.abs(e)).padStart(expDigits.filter(x => x.v === '0').length || 1, '0');
      expText = 'E' + (e < 0 ? '-' : t[expAt].sign === '+' ? '+' : '') + ed;
    }
    const r = roundShown(n, decToks.length);
    const whole = Math.floor(r + 1e-9 * Math.max(1, r) * 0) ;
    let fracDigits = decToks.length ? r.toFixed(decToks.length).split('.')[1] || '' : '';
    let intDigits = digitsOf(r + (decToks.length ? 0 : 0));
    if (decToks.length && r.toFixed(decToks.length).split('.')[0] !== intDigits) intDigits = r.toFixed(decToks.length).split('.')[0];
    if (intDigits === '0') intDigits = '';
    void whole;
    // Integer positions, right to left: each placeholder takes one digit, the first takes the rest.
    const ints = intToks.length ? intToks : [];
    const cells = new Map();
    const k = ints.length;
    const ds = intDigits.split('');
    ints.forEach((ti, j) => {
      const fromRight = k - 1 - j;
      const have = fromRight < ds.length;
      let items = [];
      if (j === 0 && ds.length > k) items = ds.slice(0, ds.length - k).map(c => ({ ch: c, real: true }));
      if (have) items.push({ ch: ds[ds.length - 1 - fromRight], real: true });
      else { const p = t[ti].v; items.push(p === '0' ? { ch: '0', real: true } : p === '?' ? { ch: ' ', real: false, space: true } : { ch: '', real: false }); }
      cells.set(ti, items);
    });
    if (group) {
      const all = ints.flatMap(ti => cells.get(ti));
      for (let p = all.length - 3; p > 0; p -= 3) {
        const left = all[p - 1];
        all[p - 1] = { ...left, ch: left.ch + (left.real ? ',' : left.space ? ' ' : '') };
      }
      // Put the grouped positions back into their placeholders.
      let at = 0;
      ints.forEach(ti => { const len = cells.get(ti).length; cells.set(ti, all.slice(at, at + len)); at += len; });
    }
    // Decimals: trailing zeros in # and ? places are dropped (? leaves a space).
    const dec = decToks.map((ti, j) => ({ ti, d: fracDigits[j] || '0' }));
    for (let j = dec.length - 1; j >= 0; j--) {
      const p = t[dec[j].ti].v;
      if (dec[j].d !== '0' || p === '0') break;
      dec[j].d = p === '?' ? ' ' : '';
    }
    const decMap = new Map(dec.map(x => [x.ti, x.d]));
    let out = '';
    let placedInt = intToks.length > 0;
    t.forEach((x, i) => {
      if (i >= end && i !== expAt) {
        if (x.k === 'dig') {
          // Exponent digits were written with the E.
          return;
        }
      }
      if (x.k === 'dig') {
        if (cells.has(i)) out += cells.get(i).map(c => c.ch).join('');
        else if (decMap.has(i)) out += decMap.get(i);
        return;
      }
      if (x.k === 'dot' && i < end) {
        if (!placedInt) { out += intDigits; placedInt = true; }
        out += '.'; return;
      }
      if (i === expAt) { out += expText; return; }
      if (x.k === 'comma') { if (role.get(i) === 'lit' || i >= end) out += ','; return; }
      if (x.k === 'pct') { out += '%'; return; }
      if (x.k === 'lit') { out += x.v; return; }
      if (x.k === 'text') { out += generalText(n); return; }
      if (x.k === 'slash') out += '/';
    });
    if (!placedInt && intDigits) out = intDigits + out;
    return out;
  }

  function formatFraction(n, t) {
    const slash = t.findIndex(x => x.k === 'slash');
    let ne = slash - 1; while (ne >= 0 && t[ne].k !== 'dig') ne--;
    let ns = ne; while (ns > 0 && t[ns - 1].k === 'dig') ns--;
    let ds = slash + 1; while (ds < t.length && t[ds].k !== 'dig' && !(t[ds].k === 'lit' && /^\d+$/.test(t[ds].v))) ds++;
    let de = ds; while (de + 1 < t.length && t[de + 1].k === 'dig') de++;
    const numToks = ne >= 0 ? t.slice(ns, ne + 1).map(x => x.v) : [];
    const intIdx = [];
    for (let i = 0; i < ns; i++) if (t[i].k === 'dig') intIdx.push(i);
    // A fixed denominator ("# ?/8") is written as digits.
    let fixed = null, denToks = [];
    if (t[ds] && t[ds].k === 'lit') { let s = ''; let j = ds; while (t[j] && ((t[j].k === 'lit' && /^\d+$/.test(t[j].v)) || (t[j].k === 'dig' && t[j].v === '0'))) { s += t[j].k === 'lit' ? t[j].v : '0'; j++; } fixed = +s; de = j - 1; }
    else denToks = t.slice(ds, de + 1).filter(x => x.k === 'dig').map(x => x.v);
    const hasInt = intIdx.length > 0;
    let whole = hasInt ? Math.floor(n) : 0;
    const frac = n - whole;
    let num, den;
    if (fixed) { den = fixed; num = Math.round(frac * den); }
    else {
      const max = Math.pow(10, denToks.length || 1) - 1;
      let best = [Math.round(frac), 1], err = Math.abs(frac - best[0]);
      for (let d = 1; d <= max && err > 1e-12; d++) { const nn = Math.round(frac * d), e2 = Math.abs(frac - nn / d); if (e2 < err - 1e-12) { err = e2; best = [nn, d]; } }
      [num, den] = best;
    }
    if (hasInt && num === den && den) { whole += 1; num = 0; }
    const fill = (toks, digits, right) => {
      // Placeholders for a number: right-aligned (integer, numerator) or left-aligned (denominator).
      const s = String(digits);
      if (s.length >= toks.length) return s;
      const pad = toks.slice(0, toks.length - s.length);
      const padText = pad.map(p => (p === '0' ? '0' : p === '?' ? ' ' : '')).join('');
      if (right) return padText + s;
      if (toks.includes('0')) return s.padStart(toks.length, '0');
      return s + toks.slice(s.length).map(p => (p === '?' ? ' ' : p === '0' ? '0' : '')).join('');
    };
    const numHasZero = numToks.includes('0');
    const hideFrac = num === 0 && !numHasZero && hasInt;
    const hideInt = hasInt && whole === 0 && !intIdx.some(i => t[i].v === '0') && !hideFrac;
    const qmark = numToks.includes('?') || denToks.includes('?');
    let out = '';
    // The integer part, with its literals filled from the right.
    if (hasInt) {
      const ints = intIdx.map(i => t[i].v);
      const ds2 = whole === 0 ? (hideInt ? [] : ['0']) : String(whole).split('');
      const k = ints.length;
      const parts = new Map();
      intIdx.forEach((ti, j) => {
        const fromRight = k - 1 - j;
        let s = '';
        if (j === 0 && ds2.length > k) s = ds2.slice(0, ds2.length - k).join('');
        if (fromRight < ds2.length) s += ds2[ds2.length - 1 - fromRight];
        else s += ints[j] === '0' ? '0' : ints[j] === '?' ? ' ' : '';
        parts.set(ti, s);
      });
      for (let i = 0; i < ns; i++) {
        if (parts.has(i)) out += parts.get(i);
        else if (t[i].k === 'lit') {
          const between = i > intIdx[0] || false;
          const afterInt = i > intIdx[intIdx.length - 1];
          if (afterInt && hideInt) out += numToks.includes('?') || intIdx.some(j => t[j].v === '?') ? ' '.repeat(t[i].v.length) : '';
          else out += t[i].v;
          void between;
        }
      }
    } else for (let i = 0; i < ns; i++) if (t[i].k === 'lit') out += t[i].v;
    // The fraction (numerator, what is around the slash, denominator), or as many spaces.
    let fracText = '';
    if (ne >= 0) {
      fracText += fill(numToks, num, true);
      for (let i = ne + 1; i < (fixed ? ds : ds); i++) fracText += t[i].k === 'slash' ? '/' : t[i].k === 'lit' ? t[i].v : '';
      fracText += fixed ? String(fixed) : fill(denToks, den, false);
    }
    if (hideFrac) {
      // Nothing to show after the whole number: spaces if the denominator holds its place.
      let around = '';
      for (let i = hasInt ? intIdx[intIdx.length - 1] + 1 : 0; i < ns; i++) if (t[i].k === 'lit') around += t[i].v;
      if (qmark) { out = out.slice(0, out.length - around.length) + ' '.repeat(around.length + fracText.length); }
      else out = out.slice(0, out.length - around.length);
      if (!hasInt) out += '0';
    } else out += fracText;
    for (let i = de + 1; i < t.length; i++) out += t[i].k === 'lit' ? t[i].v : t[i].k === 'pct' ? '%' : '';
    return out;
  }

  function formatDate(serial, s) {
    if (serial < 0 || !isFinite(serial)) return '#'.repeat(10);
    const p = s.frac || 0;
    // Rounded to what is shown: whole seconds, or the fractions of a second the format has.
    const unit = Math.pow(10, p);
    let total = Math.round(serial * 86400 * unit) / unit;
    const day = Math.floor(total / 86400);
    let secs = total - day * 86400;
    const d = serialToDate(day);
    const Y = d.getUTCFullYear(), M = d.getUTCMonth(), D = d.getUTCDate(), dow = d.getUTCDay();
    const h = Math.floor(secs / 3600), mi = Math.floor((secs % 3600) / 60), sWhole = Math.floor(secs % 60);
    const fracS = secs - Math.floor(secs);
    const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
    const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    const ampm = s.toks.some(t => t.k === 'ampm');
    const p2 = x => String(x).padStart(2, '0');
    let out = '';
    for (const t of s.toks) {
      if (t.k === 'lit') { out += t.v; continue; }
      if (t.k === 'frac') { out += '.' + fracS.toFixed(t.n).slice(2); continue; }
      if (t.k === 'ampm') {
        const pm = h >= 12;
        if (t.v.length > 3) out += pm ? 'PM' : 'AM';
        else { const c = pm ? t.v[2] : t.v[0]; out += c; }
        continue;
      }
      if (t.k === 'dig') { out += t.v === '0' ? '0' : ''; continue; }
      if (t.k === 'dot') { out += '.'; continue; }
      if (t.k === 'comma') { out += ','; continue; }
      if (t.k === 'slash') { out += '/'; continue; }
      if (t.k === 'pct') { out += '%'; continue; }
      if (t.k !== 'date') continue;
      const v = t.v, c = v[0], len = v.length;
      if (t.elapsed) {
        const tot = c === 'h' ? Math.floor(total / 3600) : c === 'm' ? Math.floor(total / 60) : Math.floor(total);
        out += String(tot).padStart(len, '0');
        continue;
      }
      if (c === 'y') out += len <= 2 ? p2(Y % 100) : String(Y);
      else if (c === 'm' && t.minute) out += len === 1 ? String(mi) : p2(mi);
      else if (c === 'm') out += len === 1 ? String(M + 1) : len === 2 ? p2(M + 1) : len === 3 ? months[M].slice(0, 3) : len === 5 ? months[M][0] : months[M];
      else if (c === 'd') out += len === 1 ? String(D) : len === 2 ? p2(D) : len === 3 ? days[dow].slice(0, 3) : days[dow];
      else if (c === 'h') { const hh = ampm ? ((h + 11) % 12) + 1 : h; out += len === 1 ? String(hh) : p2(hh); }
      else if (c === 's') out += len === 1 ? String(sWhole) : p2(sWhole);
    }
    return out;
  }

  // ---------- Workbook evaluator ----------
  /**
   * Evaluates cells for a workbook-like object:
   *   getRaw(sheetName, row, col) -> raw input string ('' if empty)
   *   sheetNames() -> list of sheet names
   *   formulasInRange(sheetName, r1, c1, r2, c2) -> [[row, col], ...] (optional, speeds up big ranges)
   */
  class Engine {
    constructor(book) {
      this.book = book;
      this.cache = new Map();
      this.astCache = new Map();
      this.stack = new Set();
      this.scopes = [];             // LET names
      this.cur = null;              // the cell being computed, for ROW() and COLUMN()
      this.spills = new Map();      // anchor key -> { sheet, row, col, rows }
      this.cover = new Map();       // spilled cell key -> anchor key
      this.cands = new Map();       // sheet -> formula cells that may spill
      this.maySpillCache = new Map();
    }
    invalidate() { this.cache.clear(); this.spills.clear(); this.cover.clear(); this.cands.clear(); }
    ast(raw) {
      let ast = this.astCache.get(raw);
      if (!ast) { ast = parse(raw.slice(1)); this.astCache.set(raw, ast); }
      return ast;
    }
    /** Formula cells on a sheet whose result could be an array that spills. */
    candidates(sheet) {
      let list = this.cands.get(sheet);
      if (list) return list;
      list = [];
      const cells = this.book.formulasInRange ? this.book.formulasInRange(sheet, 0, 0, 1048575, 16383) : [];
      for (const [r, c] of cells) {
        const raw = this.book.getRaw(sheet, r, c);
        let may = this.maySpillCache.get(raw);
        if (may === undefined) { try { may = maySpill(this.ast(raw)); } catch { may = false; } this.maySpillCache.set(raw, may); }
        if (may) list.push([r, c]);
      }
      this.cands.set(sheet, list);
      return list;
    }
    /** If (row, col) is empty, computes the formulas that could spill into it. */
    spillLookup(sheet, row, col) {
      const key = sheet + '\u0000' + row + ',' + col;
      for (const [r, c] of this.candidates(sheet)) {
        if (r > row || c > col || (r === row && c === col)) continue;
        if (this.stack.has(sheet + '\u0000' + r + ',' + c)) continue;
        this.value(sheet, r, c);
        if (this.cache.has(key)) return;
      }
    }
    /** Writes an array result over the cells below and right of its formula, or #SPILL! if they aren't free. */
    spill(sheet, row, col, range) {
      const rows = range.rows, h = rows.length, w = h ? rows[0].length : 0;
      if (!h || !w) return ERR.CALC;
      if (row + h > 1048576 || col + w > 16384) return ERR.SPILL;
      const anchor = sheet + '\u0000' + row + ',' + col;
      const norm = v => (v === undefined || v === null ? '' : v);
      for (let i = 0; i < h; i++) for (let j = 0; j < w; j++) {
        if (!i && !j) continue;
        const raw = this.book.getRaw(sheet, row + i, col + j);
        if (raw !== '' && raw !== undefined && raw !== null) return ERR.SPILL;
        const k = sheet + '\u0000' + (row + i) + ',' + (col + j);
        if (this.cover.has(k) && this.cover.get(k) !== anchor) return ERR.SPILL;
      }
      for (let i = 0; i < h; i++) for (let j = 0; j < w; j++) {
        if (!i && !j) continue;
        const k = sheet + '\u0000' + (row + i) + ',' + (col + j);
        this.cover.set(k, anchor);
        this.cache.set(k, norm(rows[i][j]));
      }
      this.spills.set(anchor, { sheet, row, col, rows });
      return norm(rows[0][0]);
    }
    /** The spill range that (row, col) belongs to: { row, col, h, w } of its anchor, or null. */
    spillAt(sheet, row, col) {
      this.value(sheet, row, col);
      const key = sheet + '\u0000' + row + ',' + col;
      const a = this.spills.get(key) || this.spills.get(this.cover.get(key));
      return a ? { row: a.row, col: a.col, h: a.rows.length, w: a.rows[0].length } : null;
    }
    /** Every spill range on a sheet (computing the formulas that may spill). */
    spillsOn(sheet) {
      for (const [r, c] of this.candidates(sheet)) this.value(sheet, r, c);
      return [...this.spills.values()].filter(s => s.sheet === sheet).map(s => ({ row: s.row, col: s.col, h: s.rows.length, w: s.rows[0].length }));
    }
    /** Returns the computed value of a cell: number | string | boolean | FormulaError | ''. */
    value(sheet, row, col) {
      const key = sheet + '\u0000' + row + ',' + col;
      if (this.cache.has(key)) return this.cache.get(key);
      // A constant is its own value: no caching, no dependency walk (a SUM over a million numbers
      // reads a million of these).
      const lit = this.literal(sheet, row, col);
      if (lit !== NOT_LITERAL) return lit;
      if (this.stack.has(key)) return ERR.CIRC;
      if (!this.stack.size) this.warm(sheet, row, col);
      if (this.cache.has(key)) return this.cache.get(key);
      const raw = this.book.getRaw(sheet, row, col);
      if (raw === '' || raw === undefined || raw === null) {
        this.spillLookup(sheet, row, col);
        if (this.cache.has(key)) return this.cache.get(key);
        this.cache.set(key, '');
        return '';
      }
      this.stack.add(key);
      let v;
      const wasLegacy = this.legacy;
      this.legacy = !!(this.book.isLegacy && this.book.isLegacy(sheet, row, col));
      try { v = this.compute(sheet, raw, { row, col, spill: true }); } finally { this.stack.delete(key); this.legacy = wasLegacy; }
      // A formula from before dynamic arrays (an older Excel) takes the one value in its own row
      // or column, as Excel's implicit intersection does, instead of spilling.
      if (v instanceof Range && this.book.isLegacy && this.book.isLegacy(sheet, row, col)) v = intersect(v, row, col);
      if (v instanceof Range) v = this.spill(sheet, row, col, v);
      this.cache.set(key, v);
      return v;
    }
    /** A non-formula, non-empty cell's value, or NOT_LITERAL. */
    literal(sheet, row, col) {
      const raw = this.book.getRaw(sheet, row, col);
      if (raw === '' || raw === undefined || raw === null) return NOT_LITERAL;
      if (typeof raw !== 'string') return raw;
      if (raw[0] === '=' && raw.length > 1) return NOT_LITERAL;
      return parseLiteral(raw);
    }
    /**
     * Evaluates a cell's precedents bottom-up with an explicit stack, so long chains
     * (e.g. a running total down 50,000 rows) don't overflow the JavaScript call stack.
     */
    warm(sheet, row, col) {
      const onPath = new Set();
      const frames = [{ sheet, row, col, key: sheet + '\u0000' + row + ',' + col, deps: null, i: 0 }];
      onPath.add(frames[0].key);
      while (frames.length) {
        const f = frames[frames.length - 1];
        if (!f.deps) f.deps = this.precedents(f.sheet, f.row, f.col);
        if (f.i < f.deps.length) {
          const [s2, r2, c2] = f.deps[f.i++];
          const k2 = s2 + '\u0000' + r2 + ',' + c2;
          if (this.cache.has(k2) || onPath.has(k2)) continue;
          onPath.add(k2);
          frames.push({ sheet: s2, row: r2, col: c2, key: k2, deps: null, i: 0 });
          continue;
        }
        frames.pop();
        onPath.delete(f.key);
        if (frames.length) {
          // Compute now: its precedents are cached, so this recursion stays shallow.
          this.stack.add('\u0001warm');
          try { this.value(f.sheet, f.row, f.col); } finally { this.stack.delete('\u0001warm'); }
        }
      }
    }
    /** The sheets from a to b, as they are ordered in the workbook (a 3D reference). */
    sheetSpan(a, b) {
      const names = this.book.sheetNames ? this.book.sheetNames() : [];
      const lower = names.map(x => x.toLowerCase());
      let i = lower.indexOf(String(a).toLowerCase()), j = lower.indexOf(String(b).toLowerCase());
      if (i < 0 || j < 0) throw ERR.REF;
      if (i > j) [i, j] = [j, i];
      return names.slice(i, j + 1);
    }
    /** Cells a formula cell refers to directly (only non-empty formula cells matter for depth). */
    precedents(sheet, row, col) {
      const raw = this.book.getRaw(sheet, row, col);
      if (typeof raw !== 'string' || raw[0] !== '=' || raw.length < 2) return [];
      let ast = this.astCache.get(raw);
      if (!ast) { try { ast = parse(raw.slice(1)); } catch { return []; } this.astCache.set(raw, ast); }
      const out = [];
      const visit = n => {
        if (!n || typeof n !== 'object') return;
        if (n.k === 'ref3d') {
          let span = [];
          try { span = this.sheetSpan(n.s1, n.s2); } catch { return; }
          for (const s of span) visit({ k: 'range', sheet: s, r1: n.r1, c1: n.c1, r2: n.r2, c2: n.c2 });
          return;
        }
        if (n.k === 'ref' || n.k === 'range' || n.k === 'spill') {
          let s;
          try { s = this.resolveSheet(n.sheet, sheet); } catch { return; }
          if (n.k === 'ref' || n.k === 'spill') { out.push([s, n.row, n.col]); return; }
          if (this.book.formulasInRange) {
            for (const [r, c] of this.book.formulasInRange(s, n.r1, n.c1, n.r2, n.c2)) out.push([s, r, c]);
            return;
          }
          const area = (n.r2 - n.r1 + 1) * (n.c2 - n.c1 + 1);
          if (area > 200000) return; // huge ranges are evaluated on demand
          for (let r = n.r1; r <= n.r2; r++) for (let c = n.c1; c <= n.c2; c++) {
            const v = this.book.getRaw(s, r, c);
            if (typeof v === 'string' && v[0] === '=') out.push([s, r, c]);
          }
          return;
        }
        if (n.k === 'name' && this.book.resolveName) {
          const t = this.book.resolveName(n.v, sheet);
          if (t && !(this.nameSeen || (this.nameSeen = new Set())).has(n.v)) {
            this.nameSeen.add(n.v);
            try { visit(this.ast('=' + String(t).replace(/^=/, '').replace(/\$/g, ''))); } catch { /* not a formula */ } finally { this.nameSeen.delete(n.v); }
          }
          return;
        }
        if (n.k === 'tref' && this.book.tableRef) {
          let g = null;
          try { g = this.book.tableRef(n, sheet, { sheet, row, col }); } catch { g = null; }
          if (g) visit({ k: 'range', sheet: g.sheet, r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2 });
          return;
        }
        if (n.e) visit(n.e);
        if (n.l) visit(n.l);
        if (n.r) visit(n.r);
        if (n.args) n.args.forEach(visit);
      };
      visit(ast);
      return out;
    }
    /**
     * Evaluates raw cell input. `at` is the cell it is for ({ row, col }), which ROW() and COLUMN()
     * use; with at.spill an array result is returned whole so the caller can spill it.
     */
    compute(sheet, raw, at) {
      if (raw === null || raw === undefined || raw === '') return '';
      if (typeof raw !== 'string') return raw;
      if (raw[0] === '=' && raw.length > 1) {
        const prev = this.cur;
        this.cur = at ? { sheet, row: at.row, col: at.col } : null;
        try {
          const ast = this.ast(raw);
          let v = this.evalNode(ast, sheet);
          if (v instanceof Range) {
            if (at && at.spill && isMulti(v)) return v;
            v = v.flat()[0];
          }
          if (typeof v === 'number' && !isFinite(v)) return ERR.NUM;
          // Like Excel, =A1 pointing at an empty cell shows 0.
          if (v === undefined || v === null || (v === '' && (ast.k === 'ref' || ast.k === 'range'))) return 0;
          return v;
        } catch (e) {
          if (isErr(e)) return e;
          return ERR.ERROR;
        } finally { this.cur = prev; }
      }
      return parseLiteral(raw);
    }
    evalNode(n, sheet) {
      switch (n.k) {
        case 'num': case 'str': case 'bool': return n.v;
        case 'empty': return '';
        case 'name': {
          for (let i = this.scopes.length - 1; i >= 0; i--) if (this.scopes[i].has(n.v)) return this.scopes[i].get(n.v);
          return this.definedName(n.v, sheet);
        }
        case 'tref': {
          if (!this.book.tableRef) throw ERR.REF;
          const g = this.book.tableRef(n, sheet, this.cur);
          if (!g) throw ERR.REF;
          if (g.r1 === g.r2 && g.c1 === g.c2) return this.value(g.sheet, g.r1, g.c1);
          return this.evalNode({ k: 'range', sheet: g.sheet, r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2 }, sheet);
        }
        case 'spill': {
          const s = this.resolveSheet(n.sheet, sheet);
          const v = this.value(s, n.row, n.col);
          const sp = this.spills.get(s + '\u0000' + n.row + ',' + n.col);
          if (sp) return new Range(sp.rows);
          if (isErr(v)) throw v;
          throw ERR.REF;
        }
        case 'ref': {
          const s = this.resolveSheet(n.sheet, sheet);
          const v = this.value(s, n.row, n.col);
          return v;
        }
        case 'err': throw Object.values(ERR).find(e => e.error === n.v) || ERR.ERROR;
        case 'ref3d': {
          const rows = [];
          for (const s of this.sheetSpan(n.s1, n.s2)) {
            for (let r = n.r1; r <= n.r2; r++) {
              const row = [];
              for (let c = n.c1; c <= n.c2; c++) row.push(this.value(s, r, c));
              rows.push(row);
            }
          }
          return new Range(rows);
        }
        case 'range': {
          const s = this.resolveSheet(n.sheet, sheet);
          const rows = [];
          // Tall ranges read whole columns of values at once, when the workbook keeps them.
          const cols = n.r2 - n.r1 >= 63 && this.book.columnValues ? [] : null;
          if (cols) for (let c = n.c1; c <= n.c2; c++) cols.push(this.book.columnValues(s, c));
          const noSpills = cols && !this.candidates(s).length;
          for (let r = n.r1; r <= n.r2; r++) {
            const row = [];
            for (let c = n.c1; c <= n.c2; c++) {
              const col = cols && cols[c - n.c1];
              if (col) {
                const x = col[r];
                row.push(x === undefined ? (noSpills ? '' : this.value(s, r, c)) : x === FORMULA_CELL ? this.value(s, r, c) : x);
              } else { const lit = this.literal(s, r, c); row.push(lit !== NOT_LITERAL ? lit : this.value(s, r, c)); }
            }
            rows.push(row);
          }
          const out = new Range(rows);
          out.at = { r1: n.r1, c1: n.c1 };
          return out;
        }
        case 'neg': {
          const x = this.evalNode(n.e, sheet);
          if (isMulti(x)) return broadcast([x], ([v]) => (n.sign === '-' ? -toNum(v) : toNum(v)));
          const v = toNum(x); return n.sign === '-' ? -v : v;
        }
        case 'pct': {
          const x = this.evalNode(n.e, sheet);
          if (isMulti(x)) return broadcast([x], ([v]) => toNum(v) / 100);
          return toNum(x) / 100;
        }
        case 'bin': {
          let l = this.evalNode(n.l, sheet), r = this.evalNode(n.r, sheet);
          if (this.legacy && this.cur) { if (l instanceof Range) l = intersect(l, this.cur.row, this.cur.col); if (r instanceof Range) r = intersect(r, this.cur.row, this.cur.col); }
          return this.binary(n.op, l, r);
        }
        case 'fn': {
          if (LAZY[n.name]) {
            const legacyArg = this.legacy && this.cur && LEGACY_LAZY.has(n.name);
            return LAZY[n.name](n.args, a => { const v = this.evalNode(a, sheet); return legacyArg && v instanceof Range ? intersect(v, this.cur.row, this.cur.col) : v; }, { engine: this, sheet });
          }
          const fn = FN[n.name];
          if (!fn) throw ERR.NAME;
          let args = n.args.map(a => {
            const v = this.evalNode(a, sheet);
            return a.k === 'ref' && REF_AS_RANGE.has(n.name) ? new Range([[v]]) : v;
          });
          // A formula from an older Excel gives a scalar function the one value of a range in its own
          // row or column (implicit intersection).
          if (this.legacy && this.cur && (LIFT.has(n.name) || n.name === 'CONCATENATE')) args = args.map(v => (v instanceof Range ? intersect(v, this.cur.row, this.cur.col) : v));
          // Errors in direct (non-range) arguments propagate (except to the functions that test for them).
          if (!TESTS_ERRORS.has(n.name)) for (const a of args) if (isErr(a)) throw a;
          // A function of single values given arrays works element by element (and spills).
          if (LIFT.has(n.name) && args.some(isMulti)) return broadcast(args, xs => { for (const x of xs) if (isErr(x)) throw x; return fn(xs); });
          const v = fn(args);
          if (isErr(v)) throw v;
          return v;
        }
      }
      throw ERR.ERROR;
    }
    /** A workbook name (Name Manager): what it refers to, evaluated where the formula is. */
    definedName(name, sheet) {
      const refersTo = this.book.resolveName ? this.book.resolveName(name, sheet) : null;
      if (refersTo === null || refersTo === undefined || refersTo === '') throw ERR.NAME;
      if (!this.nameStack) this.nameStack = new Set();
      if (this.nameStack.has(name)) throw ERR.CIRC;
      this.nameStack.add(name);
      try {
        const text = String(refersTo).replace(/^=/, '');
        const ast = this.ast('=' + text.replace(/\$/g, ''));
        return this.evalNode(ast, sheet);
      } finally { this.nameStack.delete(name); }
    }
    resolveSheet(name, current) {
      if (!name) return current;
      const found = this.book.sheetNames().find(s => s.toLowerCase() === name.toLowerCase());
      if (!found) throw ERR.REF;
      return found;
    }
    binary(op, l, r) {
      if (isMulti(l) || isMulti(r)) return broadcast([l, r], ([a, b]) => this.binary(op, a, b));
      l = scalar(l); r = scalar(r);
      if (isErr(l)) throw l;
      if (isErr(r)) throw r;
      switch (op) {
        case '+': return toNum(l) + toNum(r);
        case '-': return toNum(l) - toNum(r);
        case '*': return toNum(l) * toNum(r);
        case '/': { const d = toNum(r); if (d === 0) throw ERR.DIV0; return toNum(l) / d; }
        case '^': return Math.pow(toNum(l), toNum(r));
        case '&': return toStr(l) + toStr(r);
        default: {
          const a = isBlank(l) ? (typeof r === 'string' ? '' : 0) : l;
          const b = isBlank(r) ? (typeof l === 'string' ? '' : 0) : r;
          let c;
          if (typeof a === 'number' && typeof b === 'number') c = a - b;
          else if (typeof a === 'boolean' || typeof b === 'boolean') c = (typeof a === typeof b) ? (+a - +b) : (typeof a === 'boolean' ? 1 : -1);
          else if (typeof a === 'number') c = -1;
          else if (typeof b === 'number') c = 1;
          else c = String(a).toLowerCase().localeCompare(String(b).toLowerCase());
          return { '=': c === 0, '<>': c !== 0, '<': c < 0, '>': c > 0, '<=': c <= 0, '>=': c >= 0 }[op];
        }
      }
    }
  }

  /** Turns typed input into a value: numbers, percentages, booleans, currency, otherwise text. */
  function parseLiteral(raw) {
    const s = raw.trim();
    if (s === '') return '';
    if (s[0] === '#') { const e = Object.values(ERR).find(x => x.error === s.toUpperCase()); if (e && e !== ERR.ERROR && e !== ERR.CIRC) return e; }
    if (/^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/.test(s)) return parseFloat(s);
    if (/^[-+]?\$?\d{1,3}(,\d{3})+(\.\d+)?$/.test(s) || /^[-+]?\$\d+(\.\d+)?$/.test(s)) return parseFloat(s.replace(/[$,]/g, ''));
    if (/^[-+]?(\d+\.?\d*|\.\d+)%$/.test(s)) return parseFloat(s) / 100;
    const up = s.toUpperCase();
    if (up === 'TRUE') return true;
    if (up === 'FALSE') return false;
    if (raw[0] === "'") return raw.slice(1);
    return raw;
  }

  // ---------- Reference rewriting (copy/paste, insert/delete rows & columns) ----------
  /**
   * Rewrites every cell reference in a formula. fn(ref) receives
   * {row, col, absRow, absCol, sheet} and returns a new one or null for #REF!.
   */
  function rewriteRefs(formula, fn) {
    if (!formula || formula[0] !== '=') return formula;
    let out = '';
    let i = 1;
    const src = formula;
    const refRe = /((?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?::(\$?)([A-Za-z]{1,3})(\$?)(\d+))?(?![\w(!])/y;
    out = '=';
    while (i < src.length) {
      const ch = src[i];
      if (ch === '"') { // copy strings verbatim
        let j = i + 1;
        while (j < src.length && !(src[j] === '"' && src[j + 1] !== '"')) j += src[j] === '"' ? 2 : 1;
        out += src.slice(i, j + 1);
        i = j + 1;
        continue;
      }
      if (ch === '[') { // table references: column names are not cell references
        let j = i, depth = 0;
        for (; j < src.length; j++) { if (src[j] === "'") { j++; continue; } if (src[j] === '[') depth++; else if (src[j] === ']' && --depth === 0) break; }
        out += src.slice(i, j + 1);
        i = j + 1;
        continue;
      }
      const prev = src[i - 1];
      if (/[A-Za-z_$']/.test(ch) && !/[\w.]/.test(prev || '')) {
        refRe.lastIndex = i;
        const m = refRe.exec(src);
        if (m) {
          const sheet = m[1] ? m[1].slice(0, -1) : null;
          const conv = (absC, col, absR, row, which) => fn({ sheet, absCol: !!absC, col: colToIndex(col), absRow: !!absR, row: +row - 1 }, which);
          const a = conv(m[2], m[3], m[4], m[5], m[7] ? 'start' : 'single');
          const b = m[7] ? conv(m[6], m[7], m[8], m[9], 'end') : null;
          if (!a || (m[7] && !b)) out += '#REF!';
          else {
            const fmt = r => (r.absCol ? '$' : '') + indexToCol(r.col) + (r.absRow ? '$' : '') + (r.row + 1);
            out += (m[1] || '') + fmt(a) + (b ? ':' + fmt(b) : '');
          }
          i = refRe.lastIndex;
          continue;
        }
        // Function or name: copy the identifier whole so we don't match refs inside it (e.g. LOG10).
        const w = /[A-Za-z_][\w.]*/y; w.lastIndex = i;
        const wm = w.exec(src);
        if (wm) { out += wm[0]; i = w.lastIndex; continue; }
      }
      out += ch;
      i++;
    }
    return out;
  }
  /** Shift relative references by (dRow, dCol), as when copying a formula. */
  function shiftFormula(formula, dRow, dCol) {
    return rewriteRefs(formula, r => {
      const row = r.absRow ? r.row : r.row + dRow;
      const col = r.absCol ? r.col : r.col + dCol;
      if (row < 0 || col < 0) return null;
      return { ...r, row, col };
    });
  }
  /** Adjust references after inserting (count>0) or deleting (count<0) rows/cols at index. */
  function adjustForInsert(formula, axis, index, count, onSheet, currentSheet) {
    return rewriteRefs(formula, (r, which) => {
      const sameSheet = r.sheet ? r.sheet.replace(/^'|'$/g, '').toLowerCase() === onSheet.toLowerCase() : currentSheet === onSheet;
      if (!sameSheet) return r;
      const key = axis === 'row' ? 'row' : 'col';
      const v = r[key];
      if (count < 0 && v >= index && v < index - count) {
        // A deleted range endpoint shrinks the range, as in Excel; a lone deleted cell is #REF!.
        if (which === 'start') return { ...r, [key]: index };
        if (which === 'end') return index > 0 ? { ...r, [key]: index - 1 } : null;
        return null;
      }
      if (v >= index) return { ...r, [key]: v + count };
      return r;
    });
  }

  const api = { Engine, FormulaError, ERR, FORMULA_CELL, parse, tokenize, parseRef, addr, colToIndex, indexToCol, formatValue, formatGeneral, generalText, parseLiteral, shiftFormula, adjustForInsert, rewriteRefs, parseStructured, dateToSerial, serialToDate, toExcelFormula, fromExcelFormula, Range, FUNCTIONS: [...Object.keys(FN), ...Object.keys(LAZY)].sort() };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.Formula = api;
})(typeof window !== 'undefined' ? window : globalThis);
