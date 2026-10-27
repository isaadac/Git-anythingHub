'use strict';
// Atlas Sheets formula engine: tokenizer, parser, evaluator and ~80 Excel-compatible functions.
// Pure JavaScript with no DOM access, so it runs in the app and under `node --test`.
(function (root) {
  // ---------- Cell addressing ----------
  function colToIndex(col) {
    let n = 0;
    for (const ch of col.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  }
  function indexToCol(i) {
    let s = '';
    i += 1;
    while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); }
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
    CIRC: new FormulaError('#CIRC!'), ERROR: new FormulaError('#ERROR!')
  };
  const isErr = v => v instanceof FormulaError;

  // ---------- Tokenizer ----------
  const TOKEN_RE = /\s*(?:(\d+\.?\d*(?:[eE][+-]?\d+)?|\.\d+(?:[eE][+-]?\d+)?)|("(?:[^"]|"")*")|((?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)|(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)(?![\w(])|([A-Za-z_][\w.]*)|(<>|<=|>=|[-+*/^&=<>%(),:;]))/y;

  function tokenize(src) {
    const tokens = [];
    TOKEN_RE.lastIndex = 0;
    let pos = 0;
    while (pos < src.length) {
      if (/^\s*$/.test(src.slice(pos))) break;
      TOKEN_RE.lastIndex = pos;
      const m = TOKEN_RE.exec(src);
      if (!m) throw ERR.ERROR;
      pos = TOKEN_RE.lastIndex;
      if (m[1] !== undefined) tokens.push({ t: 'num', v: parseFloat(m[1]) });
      else if (m[2] !== undefined) tokens.push({ t: 'str', v: m[2].slice(1, -1).replace(/""/g, '"') });
      else if (m[3] !== undefined || m[4] !== undefined) tokens.push({ t: 'ref', v: m[3] || m[4] });
      else if (m[5] !== undefined) {
        const up = m[5].toUpperCase();
        if (up === 'TRUE' || up === 'FALSE') tokens.push({ t: 'bool', v: up === 'TRUE' });
        else tokens.push({ t: 'name', v: up });
      } else tokens.push({ t: 'op', v: m[6] === ';' ? ',' : m[6] });
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
      if (t.t === 'bool') return { k: 'bool', v: t.v };
      if (t.t === 'ref') return refNode(t.v);
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
    if (!b) return { k: 'ref', sheet, row: s.row, col: s.col };
    const e = parseRef(b);
    return { k: 'range', sheet, r1: Math.min(s.row, e.row), c1: Math.min(s.col, e.col), r2: Math.max(s.row, e.row), c2: Math.max(s.col, e.col) };
  }

  // ---------- Value helpers ----------
  class Range {
    constructor(rows) { this.rows = rows; }
    flat() { return this.rows.flat(); }
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
  const dateToSerial = d => (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds()) - EPOCH) / 86400000;
  const serialToDate = s => new Date(EPOCH + Math.round(s * 86400000));

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
    IF: (a, ev) => (toBool(ev(a[0])) ? (a[1] ? ev(a[1]) : true) : (a[2] ? ev(a[2]) : false)),
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
    PRODUCT: a => nums(a).reduce((s, n) => s * n, 1),
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
    ROUND: a => roundTo(toNum(a[0]), a[1] === undefined ? 0 : toNum(a[1])),
    ROUNDUP: a => { const d = a[1] === undefined ? 0 : toNum(a[1]); const f = 10 ** d; const n = toNum(a[0]); return Math.sign(n) * Math.ceil(Math.abs(n) * f - 1e-9) / f; },
    ROUNDDOWN: a => { const d = a[1] === undefined ? 0 : toNum(a[1]); const f = 10 ** d; const n = toNum(a[0]); return Math.sign(n) * Math.floor(Math.abs(n) * f + 1e-9) / f; },
    CEILING: a => { const s = a[1] === undefined ? 1 : toNum(a[1]); return s === 0 ? 0 : Math.ceil(toNum(a[0]) / s) * s; },
    FLOOR: a => { const s = a[1] === undefined ? 1 : toNum(a[1]); if (s === 0) throw ERR.DIV0; return Math.floor(toNum(a[0]) / s) * s; },
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
    AND: a => flatten(a).filter(v => !isBlank(v)).every(toBool),
    OR: a => flatten(a).filter(v => !isBlank(v)).some(toBool),
    XOR: a => flatten(a).filter(v => !isBlank(v)).filter(toBool).length % 2 === 1,
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
      const n = toNum(a[3]); let idx = -1;
      for (let k = 0; k < n; k++) { idx = s.indexOf(f, idx + 1); if (idx < 0) return s; }
      return s.slice(0, idx) + r + s.slice(idx + f.length);
    },
    REPLACE: a => { const s = toStr(a[0]); const st = toNum(a[1]) - 1; return s.slice(0, st) + toStr(a[3]) + s.slice(st + toNum(a[2])); },
    FIND: a => { const i = toStr(a[1]).indexOf(toStr(a[0]), (a[2] === undefined ? 1 : toNum(a[2])) - 1); if (i < 0) throw ERR.VALUE; return i + 1; },
    SEARCH: a => { const i = toStr(a[1]).toLowerCase().indexOf(toStr(a[0]).toLowerCase(), (a[2] === undefined ? 1 : toNum(a[2])) - 1); if (i < 0) throw ERR.VALUE; return i + 1; },
    EXACT: a => toStr(a[0]) === toStr(a[1]),
    VALUE: a => toNum(a[0]),
    T: a => (typeof scalar(a[0]) === 'string' ? scalar(a[0]) : ''),
    N: a => (typeof scalar(a[0]) === 'number' ? scalar(a[0]) : typeof scalar(a[0]) === 'boolean' ? +scalar(a[0]) : 0),
    CHAR: a => String.fromCharCode(toNum(a[0])),
    CODE: a => toStr(a[0]).charCodeAt(0) || 0,
    TEXT: a => formatValue(toNum(a[0]), toStr(a[1])),
    FIXED: a => toNum(a[0]).toLocaleString('en-US', { minimumFractionDigits: a[1] === undefined ? 2 : toNum(a[1]), maximumFractionDigits: a[1] === undefined ? 2 : toNum(a[1]), useGrouping: !(a[2] && toBool(a[2])) }),
    DOLLAR: a => formatValue(toNum(a[0]), '$#,##0.00'),

    // Date & time
    TODAY: () => Math.floor(dateToSerial(new Date())),
    NOW: () => dateToSerial(new Date()),
    DATE: a => dateToSerial(new Date(toNum(a[0]), toNum(a[1]) - 1, toNum(a[2]))),
    TIME: a => (toNum(a[0]) * 3600 + toNum(a[1]) * 60 + toNum(a[2])) / 86400,
    YEAR: a => serialToDate(toNum(a[0])).getUTCFullYear(),
    MONTH: a => serialToDate(toNum(a[0])).getUTCMonth() + 1,
    DAY: a => serialToDate(toNum(a[0])).getUTCDate(),
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
    CHOOSE: a => { const i = toNum(a[0]); if (i < 1 || i >= a.length) throw ERR.VALUE; return a[i]; },
    ROWS: a => asRange(a[0]).rows.length,
    COLUMNS: a => asRange(a[0]).rows[0].length,

    // Finance
    PMT: a => { const r = toNum(a[0]), n = toNum(a[1]), pv = toNum(a[2]), fv = a[3] === undefined ? 0 : toNum(a[3]), t = a[4] === undefined ? 0 : toNum(a[4]); if (r === 0) return -(pv + fv) / n; const p = Math.pow(1 + r, n); return -(r * (pv * p + fv)) / ((1 + r * t) * (p - 1)); },
    FV: a => { const r = toNum(a[0]), n = toNum(a[1]), pmt = toNum(a[2]), pv = a[3] === undefined ? 0 : toNum(a[3]), t = a[4] === undefined ? 0 : toNum(a[4]); if (r === 0) return -(pv + pmt * n); const p = Math.pow(1 + r, n); return -(pv * p + pmt * (1 + r * t) * (p - 1) / r); },
    PV: a => { const r = toNum(a[0]), n = toNum(a[1]), pmt = toNum(a[2]), fv = a[3] === undefined ? 0 : toNum(a[3]), t = a[4] === undefined ? 0 : toNum(a[4]); if (r === 0) return -(fv + pmt * n); const p = Math.pow(1 + r, n); return -(fv + pmt * (1 + r * t) * (p - 1) / r) / p; },
    NPV: a => { const r = toNum(a[0]); return nums(a.slice(1)).reduce((s, v, i) => s + v / Math.pow(1 + r, i + 1), 0); }
  };

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
  /** Formats a number with a (subset of) Excel number format codes. */
  function formatValue(v, fmt) {
    if (typeof v !== 'number' || !fmt || fmt === 'General') return typeof v === 'number' ? formatGeneral(v) : v;
    if (fmt === '@') return String(v);
    const sections = fmt.split(';');
    let f = sections[0];
    let n = v;
    if (v < 0 && sections[1]) { f = sections[1]; n = -v; }
    else if (v === 0 && sections[2]) f = sections[2];
    f = f.replace(/\[[^\]]*\]/g, '');
    if (/[ymdhs]/i.test(f.replace(/"[^"]*"/g, '')) && !/[0#]/.test(f)) return formatDate(n, f);
    const sci = /([0#.]+)E([+-])(0+)/i.exec(f.replace(/"[^"]*"/g, ''));
    if (sci) {
      // Scientific, e.g. 0.00E+00
      const decimals = (sci[1].split('.')[1] || '').length;
      const [mant, exp] = Math.abs(n).toExponential(decimals).split('e');
      const e = +exp;
      const expText = (e < 0 ? '-' : sci[2] === '+' ? '+' : '') + String(Math.abs(e)).padStart(sci[3].length, '0');
      return (n < 0 && !sections[1] ? '-' : '') + mant + 'E' + expText;
    }
    const pct = f.includes('%');
    if (pct) n *= 100;
    const m = /([#0,]*)(?:\.([0#]+))?/.exec(f.replace(/"[^"]*"/g, '').replace(/[^#0,.%]/g, m2 => m2 === '%' ? '%' : ' ').trim().split(' ').find(p => /[0#]/.test(p)) || '0');
    const decimals = m && m[2] ? m[2].length : 0;
    const grouping = m && m[1] && m[1].includes(',');
    let body = Math.abs(n).toLocaleString('en-US', { minimumFractionDigits: decimals, maximumFractionDigits: decimals, useGrouping: !!grouping });
    if (m && m[1] && /^0+/.test(m[1].replace(/,/g, '')) ) {
      const minInt = m[1].replace(/[#,]/g, '').length;
      const [ip, dp] = body.split('.');
      if (ip.replace(/,/g, '').length < minInt) body = ip.padStart(minInt, '0') + (dp ? '.' + dp : '');
    }
    const prefix = (f.match(/^[^#0]*/)[0] || '').replace(/"/g, '').replace(/\\/g, '').replace(/[_*]./g, '');
    const suffix = (f.match(/[^#0%]*%?[^#0]*$/)[0] || '').replace(/"/g, '').replace(/\\/g, '').replace(/[_*]./g, '');
    const sign = n < 0 && !sections[1] ? '-' : '';
    return sign + prefix.replace('-', '') + body + suffix.replace(/^[,.]+/, '');
  }
  function formatDate(serial, f) {
    const d = serialToDate(serial);
    const Y = d.getUTCFullYear(), M = d.getUTCMonth(), D = d.getUTCDate(), h = d.getUTCHours(), mi = d.getUTCMinutes(), s = d.getUTCSeconds();
    const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
    const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    const ampm = /AM\/PM/i.test(f);
    const p2 = x => String(x).padStart(2, '0');
    return f.replace(/"([^"]*)"|yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|AM\/PM/gi, (tok, lit, off, str) => {
      if (lit !== undefined) return lit;
      const t = tok.toLowerCase();
      // "m" after h or before s means minutes.
      const isMinute = (t === 'mm' || t === 'm') && (/h[^ymd]*$/i.test(str.slice(0, off)) || /^[^ymd]*s/i.test(str.slice(off + tok.length)));
      if (isMinute) return t === 'mm' ? p2(mi) : String(mi);
      switch (t) {
        case 'yyyy': return String(Y);
        case 'yy': return p2(Y % 100);
        case 'mmmm': return months[M];
        case 'mmm': return months[M].slice(0, 3);
        case 'mm': return p2(M + 1);
        case 'm': return String(M + 1);
        case 'dddd': return days[d.getUTCDay()];
        case 'ddd': return days[d.getUTCDay()].slice(0, 3);
        case 'dd': return p2(D);
        case 'd': return String(D);
        case 'hh': return p2(ampm ? ((h + 11) % 12) + 1 : h);
        case 'h': return String(ampm ? ((h + 11) % 12) + 1 : h);
        case 'ss': return p2(s);
        case 's': return String(s);
        case 'am/pm': return h < 12 ? 'AM' : 'PM';
        default: return tok;
      }
    });
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
    }
    invalidate() { this.cache.clear(); }
    /** Returns the computed value of a cell: number | string | boolean | FormulaError | ''. */
    value(sheet, row, col) {
      const key = sheet + '\u0000' + row + ',' + col;
      if (this.cache.has(key)) return this.cache.get(key);
      if (this.stack.has(key)) return ERR.CIRC;
      if (!this.stack.size) this.warm(sheet, row, col);
      if (this.cache.has(key)) return this.cache.get(key);
      this.stack.add(key);
      let v;
      try { v = this.compute(sheet, this.book.getRaw(sheet, row, col)); } finally { this.stack.delete(key); }
      this.cache.set(key, v);
      return v;
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
    /** Cells a formula cell refers to directly (only non-empty formula cells matter for depth). */
    precedents(sheet, row, col) {
      const raw = this.book.getRaw(sheet, row, col);
      if (typeof raw !== 'string' || raw[0] !== '=' || raw.length < 2) return [];
      let ast = this.astCache.get(raw);
      if (!ast) { try { ast = parse(raw.slice(1)); } catch { return []; } this.astCache.set(raw, ast); }
      const out = [];
      const visit = n => {
        if (!n || typeof n !== 'object') return;
        if (n.k === 'ref' || n.k === 'range') {
          let s;
          try { s = this.resolveSheet(n.sheet, sheet); } catch { return; }
          if (n.k === 'ref') { out.push([s, n.row, n.col]); return; }
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
        if (n.e) visit(n.e);
        if (n.l) visit(n.l);
        if (n.r) visit(n.r);
        if (n.args) n.args.forEach(visit);
      };
      visit(ast);
      return out;
    }
    compute(sheet, raw) {
      if (raw === null || raw === undefined || raw === '') return '';
      if (typeof raw !== 'string') return raw;
      if (raw[0] === '=' && raw.length > 1) {
        try {
          let ast = this.astCache.get(raw);
          if (!ast) { ast = parse(raw.slice(1)); this.astCache.set(raw, ast); }
          let v = this.evalNode(ast, sheet);
          if (v instanceof Range) v = v.flat()[0];
          if (typeof v === 'number' && !isFinite(v)) return ERR.NUM;
          // Like Excel, =A1 pointing at an empty cell shows 0.
          if (v === undefined || v === null || (v === '' && (ast.k === 'ref' || ast.k === 'range'))) return 0;
          return v;
        } catch (e) {
          if (isErr(e)) return e;
          return ERR.ERROR;
        }
      }
      return parseLiteral(raw);
    }
    evalNode(n, sheet) {
      switch (n.k) {
        case 'num': case 'str': case 'bool': return n.v;
        case 'empty': return '';
        case 'name': throw ERR.NAME;
        case 'ref': {
          const s = this.resolveSheet(n.sheet, sheet);
          const v = this.value(s, n.row, n.col);
          return v;
        }
        case 'range': {
          const s = this.resolveSheet(n.sheet, sheet);
          const rows = [];
          for (let r = n.r1; r <= n.r2; r++) {
            const row = [];
            for (let c = n.c1; c <= n.c2; c++) row.push(this.value(s, r, c));
            rows.push(row);
          }
          return new Range(rows);
        }
        case 'neg': { const v = toNum(this.evalNode(n.e, sheet)); return n.sign === '-' ? -v : v; }
        case 'pct': return toNum(this.evalNode(n.e, sheet)) / 100;
        case 'bin': return this.binary(n.op, this.evalNode(n.l, sheet), this.evalNode(n.r, sheet));
        case 'fn': {
          if (LAZY[n.name]) return LAZY[n.name](n.args, a => this.evalNode(a, sheet));
          const fn = FN[n.name];
          if (!fn) throw ERR.NAME;
          const args = n.args.map(a => this.evalNode(a, sheet));
          // Errors in direct (non-range) arguments propagate.
          for (const a of args) if (isErr(a)) throw a;
          const v = fn(args);
          if (isErr(v)) throw v;
          return v;
        }
      }
      throw ERR.ERROR;
    }
    resolveSheet(name, current) {
      if (!name) return current;
      const found = this.book.sheetNames().find(s => s.toLowerCase() === name.toLowerCase());
      if (!found) throw ERR.REF;
      return found;
    }
    binary(op, l, r) {
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

  const api = { Engine, FormulaError, ERR, parse, tokenize, parseRef, addr, colToIndex, indexToCol, formatValue, formatGeneral, parseLiteral, shiftFormula, adjustForInsert, rewriteRefs, dateToSerial, serialToDate, FUNCTIONS: [...Object.keys(FN), ...Object.keys(LAZY)].sort() };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.Formula = api;
})(typeof window !== 'undefined' ? window : globalThis);
