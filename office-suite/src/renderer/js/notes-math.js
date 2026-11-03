'use strict';
// Isadek Notes — math help: reads a typed (or recognised) expression or equation and works it out.
//
//   NotesMath.analyze('2x + 3 = 11')  → { kind: 'solve', variable: 'x', solutions: [...], steps: [...] }
//   NotesMath.analyze('12 * (3 + 4)') → { kind: 'value', value: 84, text: '84' }
//   NotesMath.analyze('y = x^2 - 4')  → { kind: 'graph', fn, label }
//
// It understands + − × ÷ ^ and parentheses, implicit multiplication (2x, 3(x + 1), (x + 1)(x − 1)),
// powers written ² ³, √, π, and the functions sin cos tan asin acos atan sqrt ln log abs exp.
// Linear and quadratic equations are solved exactly, with the steps; anything else numerically.
(function (root) {
  const FUNCS = { sin: Math.sin, cos: Math.cos, tan: Math.tan, asin: Math.asin, acos: Math.acos, atan: Math.atan, sqrt: Math.sqrt, ln: Math.log, log: Math.log10, abs: Math.abs, exp: Math.exp };
  const CONSTS = { pi: Math.PI, e: Math.E };
  const FUNC_NAMES = Object.keys(FUNCS).sort((a, b) => b.length - a.length);
  const MAX_LEN = 500;

  // ---------------- Reading ----------------
  function normalize(s) {
    return String(s).slice(0, MAX_LEN)
      .replace(/[−–—]/g, '-').replace(/[×·∙⋅]/g, '*').replace(/÷/g, '/').replace(/π/g, 'pi')
      .replace(/√/g, 'sqrt').replace(/²/g, '^2').replace(/³/g, '^3').replace(/\*\*/g, '^')
      .replace(/(\d),(\d{3})(?!\d)/g, '$1$2').trim();
  }
  function tokenize(s) {
    const out = [];
    let i = 0;
    while (i < s.length) {
      const c = s[i];
      if (/\s/.test(c)) { i++; continue; }
      const num = /^(\d+\.?\d*|\.\d+)(e[+-]?\d+)?/i.exec(s.slice(i));
      if (num) { out.push({ t: 'n', v: parseFloat(num[0]) }); i += num[0].length; continue; }
      const word = /^[a-z]+/i.exec(s.slice(i));
      if (word) {
        let w = word[0];
        i += w.length;
        // Function names and constants, then single-letter variables (xy is x times y).
        // A lone "e" is Euler's number; inside a word it is a letter.
        if (w.toLowerCase() === 'e') { out.push({ t: 'c', v: 'e' }); continue; }
        while (w) {
          const lw = w.toLowerCase();
          const f = FUNC_NAMES.find(k => lw.startsWith(k));
          if (f) { out.push({ t: 'f', v: f }); w = w.slice(f.length); }
          else if (lw.startsWith('pi')) { out.push({ t: 'c', v: 'pi' }); w = w.slice(2); }
          else { out.push({ t: 'v', v: w[0] }); w = w.slice(1); }
        }
        continue;
      }
      if ('+-*/^()=,'.includes(c)) { out.push({ t: c }); i++; continue; }
      throw new Error(`“${c}” isn’t something this can read`);
    }
    return out;
  }
  /** A tree: { n } | { v } | { op, a, b } | { neg } | { fn, a }. */
  function parse(s) {
    const toks = tokenize(s);
    let p = 0;
    const peek = () => toks[p], next = () => toks[p++];
    const startsFactor = t => t && (t.t === 'n' || t.t === 'v' || t.t === 'c' || t.t === 'f' || t.t === '(');
    function expr() {
      let a = term();
      while (peek() && (peek().t === '+' || peek().t === '-')) { const op = next().t; a = { op, a, b: term() }; }
      return a;
    }
    function term() {
      let a = unary();
      for (;;) {
        const t = peek();
        if (t && (t.t === '*' || t.t === '/')) { next(); a = { op: t.t, a, b: unary() }; }
        else if (startsFactor(t)) a = { op: '*', a, b: power() }; // implicit: 2x, 3(x+1), (a)(b)
        else return a;
      }
    }
    function unary() {
      if (peek() && peek().t === '-') { next(); return { neg: unary() }; }
      if (peek() && peek().t === '+') { next(); return unary(); }
      return power();
    }
    function power() {
      const base = atom();
      if (peek() && peek().t === '^') { next(); return { op: '^', a: base, b: unary() }; }
      return base;
    }
    function atom() {
      const t = next();
      if (!t) throw new Error('The expression ends too soon');
      if (t.t === 'n') return { n: t.v };
      if (t.t === 'v') return { v: t.v };
      if (t.t === 'c') return { n: CONSTS[t.v], c: t.v };
      if (t.t === 'f') {
        if (peek() && peek().t === '(') { next(); const a = expr(); if (!peek() || next().t !== ')') throw new Error('A bracket isn’t closed'); return { fn: t.v, a }; }
        return { fn: t.v, a: power() };
      }
      if (t.t === '(') { const a = expr(); if (!peek() || next().t !== ')') throw new Error('A bracket isn’t closed'); return a; }
      throw new Error(`“${t.t}” is in an unexpected place`);
    }
    const tree = expr();
    if (p < toks.length) throw new Error(`“${toks[p].t === 'n' ? toks[p].v : toks[p].v || toks[p].t}” is in an unexpected place`);
    return tree;
  }
  function evaluate(t, env = {}) {
    if (t.n !== undefined) return t.n;
    if (t.v !== undefined) { if (!(t.v in env)) throw new Error('unknown ' + t.v); return env[t.v]; }
    if (t.neg) return -evaluate(t.neg, env);
    if (t.fn) return FUNCS[t.fn](evaluate(t.a, env));
    const a = evaluate(t.a, env), b = evaluate(t.b, env);
    switch (t.op) { case '+': return a + b; case '-': return a - b; case '*': return a * b; case '/': return a / b; default: return Math.pow(a, b); }
  }
  function variables(t, set = new Set()) {
    if (t.v !== undefined) set.add(t.v);
    for (const k of ['a', 'b', 'neg']) if (t[k]) variables(t[k], set);
    return set;
  }

  // ---------------- Writing numbers ----------------
  function fraction(x) {
    if (!isFinite(x) || Number.isInteger(x)) return null;
    let h1 = 1, h0 = 0, k1 = 0, k0 = 1, b = x;
    for (let i = 0; i < 20; i++) {
      const a = Math.floor(b);
      [h1, h0] = [a * h1 + h0, h1];
      [k1, k0] = [a * k1 + k0, k1];
      if (k1 > 1000) return null;
      if (Math.abs(x - h1 / k1) < 1e-9 * Math.max(1, Math.abs(x))) return k1 === 1 ? null : `${h1}/${k1}`;
      b = 1 / (b - a);
    }
    return null;
  }
  function fmt(x) {
    if (Number.isNaN(x)) return 'undefined';
    if (!isFinite(x)) return x > 0 ? '∞' : '−∞';
    if (Math.abs(x) < 1e-12) return '0';
    const r = Math.round(x);
    if (Math.abs(x - r) < 1e-9 * Math.max(1, Math.abs(x))) return String(r).replace('-', '−');
    const s = Math.abs(x) >= 1e12 || Math.abs(x) < 1e-6 ? x.toExponential(6).replace(/\.?0+e/, 'e') : String(+x.toPrecision(10));
    return s.replace('-', '−');
  }
  /** A number, with its fraction when it is a simple one: "0.75 (3/4)". */
  const fmtExact = x => { const f = fraction(x); return f ? `${f.replace('-', '−')} ≈ ${fmt(x)}` : fmt(x); };
  const coef = (a, v, first) => {
    const s = Math.abs(a) === 1 ? v : fmt(Math.abs(a)) + v;
    return first ? (a < 0 ? '−' + s : s) : (a < 0 ? ' − ' + s : ' + ' + s);
  };
  const constant = (b, first) => (first ? fmt(b) : b < 0 ? ' − ' + fmt(-b) : ' + ' + fmt(b));
  /** ax² + bx + c written out. */
  function poly(c2, c1, c0, v) {
    let s = '';
    if (Math.abs(c2) > 1e-12) s += coef(c2, v + '²', !s);
    if (Math.abs(c1) > 1e-12) s += coef(c1, v, !s);
    if (Math.abs(c0) > 1e-12 || !s) s += constant(c0, !s);
    return s;
  }

  // ---------------- Solving ----------------
  /** The coefficients [c0, c1, c2] when f is a polynomial of degree ≤ 2 in v, else null. */
  function quadratic(f) {
    const y0 = f(0), y1 = f(1), ym = f(-1);
    if (![y0, y1, ym].every(isFinite)) return null;
    const c0 = y0, c2 = (y1 + ym) / 2 - y0, c1 = (y1 - ym) / 2;
    for (const x of [2, -3, 0.5, 7, -11.25, 100]) {
      const want = c0 + c1 * x + c2 * x * x, got = f(x);
      if (!isFinite(got) || Math.abs(got - want) > 1e-7 * Math.max(1, Math.abs(want), Math.abs(got))) return null;
    }
    const clean = n => (Math.abs(n) < 1e-11 ? 0 : Math.abs(n - Math.round(n)) < 1e-9 ? Math.round(n) : n);
    return [clean(c0), clean(c1), clean(c2)];
  }
  /** Roots of f found numerically between −1000 and 1000, the ones nearest 0 first (up to 6 each way). */
  function numericRoots(f) {
    const found = [];
    for (const dir of [1, -1]) {
      const roots = [];
      let px = 0, py = f(0);
      if (py === 0) roots.push(0);
      for (let i = 1; i <= 3000 && roots.length < 6; i++) {
        const x = dir * (Math.pow(10, i / 1000) - 1), y = f(x);
        if (isFinite(py) && isFinite(y)) {
          if (y === 0) roots.push(x);
          else if (py * y < 0) {
            let a = px, b = x, fa = py;
            for (let k = 0; k < 80; k++) { const m = (a + b) / 2, fm = f(m); if (fa * fm <= 0) b = m; else { a = m; fa = fm; } }
            const r = (a + b) / 2;
            if (Math.abs(f(r)) < 1e-6 * Math.max(1, Math.abs(py), Math.abs(y))) roots.push(r); // not a pole
          }
        }
        px = x; py = y;
      }
      found.push(...roots);
    }
    const tidy = r => (Math.abs(r - Math.round(r)) < 1e-9 ? Math.round(r) : r);
    return found.map(tidy).sort((m, n) => m - n).filter((r, i, all) => !i || Math.abs(r - all[i - 1]) > 1e-7);
  }
  function solve(lhs, rhs, v, source) {
    const L = x => evaluate(lhs, { [v]: x }), R = x => evaluate(rhs, { [v]: x });
    const f = x => L(x) - R(x);
    const ql = quadratic(L), qr = quadratic(R), q = quadratic(f);
    const steps = [source];
    if (q && q[2] === 0 && q[1] !== 0 && ql && qr && ql[2] === 0 && qr[2] === 0) {
      // Linear: a1·v + b1 = a2·v + b2.
      const [b1, a1] = ql, [b2, a2] = qr;
      steps.push(`${poly(0, a1, b1, v)} = ${poly(0, a2, b2, v)}`);
      if (a2 !== 0) steps.push(`Subtract ${coef(a2, v, true).replace(/^−/, '−')} from both sides: ${poly(0, a1 - a2, b1, v)} = ${fmt(b2)}`);
      if (b1 !== 0) steps.push(`${b1 > 0 ? 'Subtract ' + fmt(b1) : 'Add ' + fmt(-b1)} ${b1 > 0 ? 'from' : 'to'} both sides: ${poly(0, a1 - a2, 0, v)} = ${fmt(b2 - b1)}`);
      const x = (b2 - b1) / (a1 - a2);
      if (a1 - a2 !== 1) steps.push(`Divide both sides by ${fmt(a1 - a2)}: ${v} = ${fmtExact(x)}`);
      else if (steps[steps.length - 1].indexOf(`${v} = `) < 0) steps.push(`${v} = ${fmtExact(x)}`);
      return { kind: 'solve', variable: v, method: 'linear', solutions: [x], text: `${v} = ${fmtExact(x)}`, steps };
    }
    if (q && q[2] !== 0) {
      const [c, b, a] = q, D = b * b - 4 * a * c;
      steps.push(`Bring everything to one side: ${poly(a, b, c, v)} = 0`);
      steps.push(`This is a quadratic equation a${v}² + b${v} + c = 0 with a = ${fmt(a)}, b = ${fmt(b)}, c = ${fmt(c)}`);
      steps.push(`Discriminant: b² − 4ac = ${fmt(b * b)} − ${fmt(4 * a * c).replace(/^−/, '(−') + (4 * a * c < 0 ? ')' : '')} = ${fmt(D)}`);
      steps.push(`Quadratic formula: ${v} = (−b ± √(b² − 4ac)) / 2a = (${fmt(-b)} ± √${fmt(D)}) / ${fmt(2 * a)}`);
      if (D < -1e-12) {
        const re = -b / (2 * a), im = Math.sqrt(-D) / Math.abs(2 * a);
        const text = `${v} = ${fmt(re)} ± ${fmt(im)}i`;
        steps.push(`The discriminant is negative, so there are no real solutions: ${text}`);
        return { kind: 'solve', variable: v, method: 'quadratic', solutions: [], complex: [[re, im], [re, -im]], text, steps };
      }
      const r = Math.sqrt(Math.max(0, D));
      const xs = [...new Set([(-b - r) / (2 * a), (-b + r) / (2 * a)].map(x => +x.toPrecision(15)))].sort((m, n) => m - n);
      const text = xs.map(x => `${v} = ${fmtExact(x)}`).join(' or ');
      steps.push(xs.length === 1 ? `The discriminant is 0, so there is one solution: ${text}` : text);
      return { kind: 'solve', variable: v, method: 'quadratic', solutions: xs, text, steps };
    }
    if (q && q[1] === 0 && q[2] === 0) {
      const always = Math.abs(q[0]) < 1e-12;
      return { kind: 'solve', variable: v, method: 'none', solutions: [], text: always ? `True for every ${v}` : 'No solution', steps: [...steps, always ? `Both sides are always equal` : `${fmt(q[0])} = 0 is never true`] };
    }
    const xs = numericRoots(f);
    steps.push(`Solved numerically: looked for where ${source.replace('=', '−').replace(/^(.*)−(.*)$/, '($1) − ($2)')} crosses zero between −1000 and 1000`);
    const text = xs.length ? xs.map(x => `${v} ≈ ${fmt(x)}`).join(', ') : 'No real solution found between −1000 and 1000';
    steps.push(text);
    return { kind: 'solve', variable: v, method: 'numeric', solutions: xs, text, steps };
  }

  /** Works out what an expression or equation is and does it. Throws Error with a message to show. */
  function analyze(input) {
    let src = normalize(input).replace(/[.;:?\s]+$/, '');
    // "2 + 3 =" asks for the value.
    if (/=$/.test(src) && src.indexOf('=') === src.length - 1) src = src.slice(0, -1).trim();
    if (!src) throw new Error('Type or select an expression or equation, like 2x + 3 = 11');
    const sides = src.split('=');
    if (sides.length > 2) throw new Error('Only one = sign, please');
    const label = src.replace(/\*/g, '·').replace(/-/g, '−');
    if (sides.length === 1) {
      const t = parse(sides[0]);
      const vars = [...variables(t)];
      if (!vars.length) { const value = evaluate(t); return { kind: 'value', value, text: fmtExact(value), steps: [label, `= ${fmtExact(value)}`] }; }
      if (vars.length === 1) {
        const v = vars[0], fn = x => evaluate(t, { [v]: x });
        return { kind: 'graph', variable: v, fn, label: `y = ${label}`, text: `Graph of y = ${label}`, roots: () => solve(t, { n: 0 }, v, `${label} = 0`) };
      }
      throw new Error(`There are ${vars.length} letters (${vars.join(', ')}); this works with one`);
    }
    const lhs = parse(sides[0]), rhs = parse(sides[1]);
    const vl = variables(lhs), vr = variables(rhs), vars = [...new Set([...vl, ...vr])];
    if (!vars.length) {
      const a = evaluate(lhs), b = evaluate(rhs), ok = Math.abs(a - b) < 1e-9 * Math.max(1, Math.abs(a), Math.abs(b));
      return { kind: 'check', value: ok, text: ok ? 'True' : `False: the left side is ${fmt(a)}, the right side ${fmt(b)}`, steps: [label, `${fmt(a)} ${ok ? '=' : '≠'} ${fmt(b)}`] };
    }
    // y = f(x): a graph.
    if (vars.length === 2 && lhs.v === 'y' && !vr.has('y')) {
      const v = vars.find(n => n !== 'y'), fn = x => evaluate(rhs, { [v]: x });
      return { kind: 'graph', variable: v, fn, label, text: `Graph of ${label}`, roots: () => solve(rhs, { n: 0 }, v, `${sides[1].trim()} = 0`) };
    }
    if (vars.length === 1 && lhs.v === 'y' && vars[0] === 'y') return { kind: 'value', value: evaluate(rhs), text: `y = ${fmtExact(evaluate(rhs))}`, steps: [label, `y = ${fmtExact(evaluate(rhs))}`] };
    if (vars.length > 1) throw new Error(`There are ${vars.length} letters (${vars.join(', ')}); this solves for one`);
    return solve(lhs, rhs, vars[0], label);
  }

  // ---------------- Graphs ----------------
  /** A graph of fn as an SVG string (width × height), x from x0 to x1, y chosen to fit. */
  function graphSvg(fn, label, { x0 = -10, x1 = 10, width = 560, height = 360, dark = false } = {}) {
    const N = 600, pts = [];
    for (let i = 0; i <= N; i++) { const x = x0 + (x1 - x0) * i / N; let y; try { y = fn(x); } catch { y = NaN; } pts.push([x, isFinite(y) ? y : NaN]); }
    const ys = pts.map(p => p[1]).filter(isFinite).sort((a, b) => a - b);
    let y0 = ys.length ? ys[Math.floor(ys.length * 0.02)] : -10, y1 = ys.length ? ys[Math.ceil(ys.length * 0.98) - 1] : 10;
    if (y0 > 0) y0 = Math.min(0, y0); if (y1 < 0) y1 = Math.max(0, y1);
    if (y1 - y0 < 1e-9) { y0 -= 1; y1 += 1; }
    const pad = (y1 - y0) * 0.08; y0 -= pad; y1 += pad;
    const M = 36, W = width - M - 12, H = height - 28 - 24;
    const X = x => M + (x - x0) / (x1 - x0) * W, Y = y => 28 + (y1 - y) / (y1 - y0) * H;
    const step = r => { const raw = r / 8, p = Math.pow(10, Math.floor(Math.log10(raw))); return [1, 2, 5, 10].map(m => m * p).find(s => s >= raw); };
    const sx = step(x1 - x0), sy = step(y1 - y0);
    const ink = dark ? '#e6e6e6' : '#1f1f1f', grid = dark ? '#3a3a3a' : '#e3e3e3', axis = dark ? '#9a9a9a' : '#666', line = dark ? '#7ab8ff' : '#1565c0';
    let g = '';
    for (let x = Math.ceil(x0 / sx) * sx; x <= x1 + 1e-9; x += sx) g += `<line x1="${X(x).toFixed(1)}" y1="28" x2="${X(x).toFixed(1)}" y2="${28 + H}" stroke="${grid}"/><text x="${X(x).toFixed(1)}" y="${Math.min(28 + H, Math.max(28, Y(0))) + 14}" font-size="11" text-anchor="middle" fill="${axis}">${fmt(+x.toPrecision(10))}</text>`;
    for (let y = Math.ceil(y0 / sy) * sy; y <= y1 + 1e-9; y += sy) g += `<line x1="${M}" y1="${Y(y).toFixed(1)}" x2="${M + W}" y2="${Y(y).toFixed(1)}" stroke="${grid}"/>${Math.abs(y) > sy / 2 ? `<text x="${Math.min(M + W, Math.max(M, X(0))) - 4}" y="${(Y(y) + 4).toFixed(1)}" font-size="11" text-anchor="end" fill="${axis}">${fmt(+y.toPrecision(10))}</text>` : ''}`;
    if (y0 <= 0 && y1 >= 0) g += `<line x1="${M}" y1="${Y(0).toFixed(1)}" x2="${M + W}" y2="${Y(0).toFixed(1)}" stroke="${axis}" stroke-width="1.2"/>`;
    if (x0 <= 0 && x1 >= 0) g += `<line x1="${X(0).toFixed(1)}" y1="28" x2="${X(0).toFixed(1)}" y2="${28 + H}" stroke="${axis}" stroke-width="1.2"/>`;
    let d = '', pen = false;
    for (const [x, y] of pts) {
      if (!isFinite(y) || y < y0 - (y1 - y0) * 2 || y > y1 + (y1 - y0) * 2) { pen = false; continue; }
      d += `${pen ? 'L' : 'M'}${X(x).toFixed(1)} ${Y(Math.max(y0 - (y1 - y0), Math.min(y1 + (y1 - y0), y))).toFixed(1)}`;
      pen = true;
    }
    const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" font-family="Segoe UI, Arial, sans-serif"><rect width="100%" height="100%" fill="${dark ? '#1e1e1e' : '#fff'}"/><defs><clipPath id="plot"><rect x="${M}" y="28" width="${W}" height="${H}"/></clipPath></defs>${g}<path d="${d}" fill="none" stroke="${line}" stroke-width="2.2" stroke-linejoin="round" clip-path="url(#plot)"/><text x="${M}" y="18" font-size="13" font-weight="600" fill="${ink}">${esc(label)}</text></svg>`;
  }

  const api = { analyze, parse, evaluate, graphSvg, fmt, fraction, normalize };
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.NotesMath = api;
})(typeof window !== 'undefined' ? window : globalThis);
