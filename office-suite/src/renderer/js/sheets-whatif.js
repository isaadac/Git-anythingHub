'use strict';
// Isadek Sheets — What-If Analysis: Goal Seek and Solver.
// Both work on the real workbook: they change the input cells, let the formulas recalculate and
// read the result, so any formula (and any chain of formulas) can be the target.
(function () {
  const finite = v => typeof v === 'number' && isFinite(v);

  /**
   * Goal Seek: the x for which f(x) = target. f(x) returns the formula's value with the changing
   * cell set to x (or a non-number when the formula fails there).
   * → { found, x, value, iterations }
   */
  function goalSeek(f, start, target, { maxIter = 200, tol = 1e-7 } = {}) {
    const g = x => { const v = f(x); return finite(v) ? v - target : NaN; };
    const scale = Math.max(1, Math.abs(target));
    const done = (x, gx, it) => ({ found: Math.abs(gx) <= tol * scale, x, value: gx + target, iterations: it });
    let x0 = finite(start) ? start : 0, g0 = g(x0);
    if (Math.abs(g0) <= tol * scale) return done(x0, g0, 0);
    // Secant steps first: fast for smooth formulas.
    let x1 = x0 === 0 ? 0.01 : x0 * 1.01, g1 = g(x1), it = 0;
    for (; it < maxIter / 2 && finite(g0) && finite(g1); it++) {
      if (Math.abs(g1) <= tol * scale) return done(x1, g1, it);
      if (g1 === g0) break;
      const x2 = x1 - g1 * (x1 - x0) / (g1 - g0);
      if (!finite(x2)) break;
      x0 = x1; g0 = g1; x1 = x2; g1 = g(x2);
    }
    // Otherwise find a sign change around the start and narrow it down (bisection).
    let a = finite(start) ? start : 0, ga = g(a), b = null, gb = NaN;
    for (let step = 1, k = 0; k < 60 && b === null; k++, step *= 2) {
      for (const cand of [a + step, a - step]) {
        const gc = g(cand);
        if (finite(gc) && finite(ga) && Math.sign(gc) !== Math.sign(ga)) { b = cand; gb = gc; break; }
      }
    }
    if (b === null) return done(x1, finite(g1) ? g1 : ga, it);
    for (; it < maxIter; it++) {
      const m = (a + b) / 2, gm = g(m);
      if (!finite(gm)) break;
      if (Math.abs(gm) <= tol * scale || Math.abs(b - a) < 1e-12 * Math.max(1, Math.abs(m))) return done(m, gm, it);
      if (Math.sign(gm) === Math.sign(ga)) { a = m; ga = gm; } else { b = m; gb = gm; }
    }
    return done((a + b) / 2, g((a + b) / 2), it);
  }

  /**
   * Solver: set the variable cells so the objective is as large / small as possible (or equals
   * a value), keeping the constraints. Nelder–Mead on a penalised objective, with the penalty
   * raised until the constraints hold; integer variables are rounded and improved locally.
   *   evaluate(x) → { obj, cons: [{ lhs, op: '<=' | '>=' | '=', rhs }] }
   * → { ok, x, obj, violation, evaluations }
   */
  function solve({ x0, evaluate, goal = 'max', target = 0, nonneg = true, integers = new Set(), timeLimit = 4000, maxEval = 20000 }) {
    const n = x0.length;
    const t0 = Date.now();
    let evals = 0;
    const project = x => x.map((v, i) => { let y = nonneg ? Math.max(0, v) : v; if (integers.has(i)) y = Math.round(y); return y; });
    const measure = x => {
      evals++;
      const r = evaluate(x);
      const obj = finite(r.obj) ? r.obj : NaN;
      let viol = 0;
      for (const c of r.cons || []) {
        if (!finite(c.lhs) || !finite(c.rhs)) { viol += 1e6; continue; }
        const d = c.op === '<=' ? c.lhs - c.rhs : c.op === '>=' ? c.rhs - c.lhs : Math.abs(c.lhs - c.rhs);
        if (d > 0) viol += d * d;
      }
      return { obj, viol };
    };
    const score = (m, mu) => {
      if (!finite(m.obj)) return Infinity;
      const base = goal === 'max' ? -m.obj : goal === 'min' ? m.obj : (m.obj - target) ** 2;
      return base + mu * m.viol;
    };
    const out = () => evals >= maxEval || Date.now() - t0 > timeLimit;
    function nelderMead(start, mu, step) {
      const f = x => score(measure(project(x)), mu);
      let simplex = [start.slice()];
      for (let i = 0; i < n; i++) { const p = start.slice(); p[i] += step[i]; simplex.push(p); }
      let vals = simplex.map(f);
      for (let iter = 0; iter < 400 * n && !out(); iter++) {
        const order = vals.map((v, i) => i).sort((a, b) => vals[a] - vals[b]);
        simplex = order.map(i => simplex[i]); vals = order.map(i => vals[i]);
        const spread = Math.abs(vals[n] - vals[0]);
        const size = Math.max(...simplex.slice(1).map(p => Math.max(...p.map((v, i) => Math.abs(v - simplex[0][i])))));
        if (spread < 1e-12 * (1 + Math.abs(vals[0])) && size < 1e-9 * (1 + Math.max(...simplex[0].map(Math.abs)))) break;
        const c = Array(n).fill(0);
        for (let j = 0; j < n; j++) for (let i = 0; i < n; i++) c[i] += simplex[j][i] / n;
        const pt = k => c.map((v, i) => v + k * (simplex[n][i] - v));
        const xr = pt(-1), fr = f(xr);
        if (fr < vals[0]) { const xe = pt(-2), fe = f(xe); if (fe < fr) { simplex[n] = xe; vals[n] = fe; } else { simplex[n] = xr; vals[n] = fr; } }
        else if (fr < vals[n - 1]) { simplex[n] = xr; vals[n] = fr; }
        else {
          const xc = fr < vals[n] ? pt(-0.5) : pt(0.5), fc = f(xc);
          if (fc < Math.min(fr, vals[n])) { simplex[n] = xc; vals[n] = fc; }
          else for (let j = 1; j <= n; j++) { simplex[j] = simplex[j].map((v, i) => simplex[0][i] + 0.5 * (v - simplex[0][i])); vals[j] = f(simplex[j]); }
        }
      }
      const best = vals.indexOf(Math.min(...vals));
      return project(simplex[best]);
    }
    let x = project(x0.map(v => (finite(v) ? v : 0)));
    let mu = 10;
    for (let round = 0; round < 8 && !out(); round++, mu *= 100) {
      const step = x.map(v => Math.max(1, Math.abs(v) * 0.25) / (round + 1));
      // A few restarts, each from the best point so far.
      for (let k = 0; k < 3 && !out(); k++) x = nelderMead(x, mu, step.map(s => s / (k + 1)));
      if (measure(x).viol < 1e-10) break;
    }
    // Integers: try nudging each by ±1 while that improves things.
    if (integers.size) {
      let cur = measure(x), improved = true;
      while (improved && !out()) {
        improved = false;
        for (const i of integers) for (const d of [1, -1]) {
          const y = x.slice(); y[i] += d;
          const p = project(y), m = measure(p);
          if (score(m, 1e9) < score(cur, 1e9)) { x = p; cur = m; improved = true; }
        }
      }
    }
    // Tidy: a value near a round number is usually that number. Each variable takes the
    // roundest nearby value that keeps the constraints and the objective.
    let cur = measure(x);
    const tolObj = 1e-5 * (1 + Math.abs(cur.obj || 0));
    for (let pass = 0; pass < 2; pass++) for (let i = 0; i < n; i++) {
      if (integers.has(i)) continue;
      for (let d = 0; d <= 6; d++) {
        const y = x.slice(); y[i] = Math.round(x[i] * 10 ** d) / 10 ** d;
        if (y[i] === x[i]) break;
        const m = measure(project(y));
        const worse = goal === 'max' ? cur.obj - m.obj : goal === 'min' ? m.obj - cur.obj : Math.abs(m.obj - target) - Math.abs(cur.obj - target);
        if (m.viol <= Math.max(cur.viol, 1e-12) && worse <= tolObj) { x = project(y); cur = m; break; }
      }
    }
    const m = measure(x);
    return { ok: m.viol < 1e-6 * (1 + Math.abs(m.obj || 0)) && finite(m.obj), x, obj: m.obj, violation: Math.sqrt(m.viol), evaluations: evals };
  }

  const api = { goalSeek, solve };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else window.SheetsWhatIf = api;
})();
