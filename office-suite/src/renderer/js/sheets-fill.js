'use strict';
// Atlas Sheets — Auto Fill series, as Excel's fill handle extends a selection.
// extend(source, n, opts) returns the raw values for the n cells after the source cells
// (or before them with opts.reverse, nearest first). Formulas are returned with an offset
// ({ formula, from }) so the caller can shift their references.
(function (root) {
  const DAYS = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
  const MONTHS = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
  const LISTS = [DAYS, MONTHS];
  const num = raw => {
    const s = String(raw).trim();
    return s !== '' && /^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/.test(s) ? parseFloat(s) : null;
  };
  const clean = x => parseFloat(x.toPrecision(15));
  /** Least-squares line through (0, y0), (1, y1)…: returns f(i). */
  function trend(ys) {
    const n = ys.length, mx = (n - 1) / 2, my = ys.reduce((a, b) => a + b, 0) / n;
    let sxy = 0, sxx = 0;
    ys.forEach((y, i) => { sxy += (i - mx) * (y - my); sxx += (i - mx) ** 2; });
    const slope = sxx ? sxy / sxx : 0;
    return i => clean(my + slope * (i - mx));
  }
  function nameIn(s) {
    const t = String(s).trim().toLowerCase();
    for (const list of LISTS) {
      const i = list.indexOf(t);
      if (i >= 0) return { list, i, short: false };
      const j = list.findIndex(x => x.slice(0, 3) === t);
      if (j >= 0 && t.length === 3) return { list, i: j, short: true };
    }
    return null;
  }
  const styleLike = (sample, word) => (sample === sample.toUpperCase() ? word.toUpperCase() : sample === sample.toLowerCase() ? word : word[0].toUpperCase() + word.slice(1));

  /**
   * source: [{ raw, isDate }] in fill order. Returns n entries: a raw string, or
   * { formula, from } where from is the index of the source cell it repeats.
   */
  function extend(source, n, opts = {}) {
    const k = source.length;
    const at = i => (opts.reverse ? -1 - i : k + i); // position of the i-th new cell on the source's axis
    const cycle = i => { const p = ((at(i) % k) + k) % k; return { p, raw: source[p].raw }; };
    const copyAll = () => Array.from({ length: n }, (_, i) => { const { p, raw } = cycle(i); return typeof raw === 'string' && raw[0] === '=' ? { formula: raw, from: p } : raw; });
    if (!k) return [];
    // Blanks keep their places; the other cells form the series in the slots between them.
    const blank = c => c.raw === '' || c.raw === undefined || c.raw === null;
    if (source.some(blank) && !source.every(blank) && !opts.copy) {
      const filled = source.filter(c => !blank(c)), idx = source.map((c, i) => (blank(c) ? -1 : i)).filter(i => i >= 0);
      const slots = Array.from({ length: n }, (_, i) => ((at(i) % k) + k) % k);
      const need = slots.filter(p => !blank(source[p])).length;
      const vals = extend(filled, need, opts);
      let j = 0;
      return slots.map(p => {
        if (blank(source[p])) return '';
        const v = vals[j++];
        return v && typeof v === 'object' ? { formula: v.formula, from: idx[v.from] } : v;
      });
    }
    if (source.some(c => typeof c.raw === 'string' && c.raw[0] === '=')) return copyAll();
    // Numbers (and dates): a linear trend; one number is copied (Ctrl reverses that), one date counts up by a day.
    const nums = source.map(c => num(c.raw));
    if (nums.every(x => x !== null)) {
      if (opts.copy) return copyAll();
      if (k === 1 && !source[0].isDate && !opts.increment) return copyAll();
      const f = k === 1 ? (i => clean(nums[0] + i)) : trend(nums);
      return Array.from({ length: n }, (_, i) => String(f(at(i))));
    }
    if (opts.copy) return copyAll();
    // Day and month names, full or short, keep their step and capitalisation.
    const names = source.map(c => nameIn(c.raw));
    if (names.every(x => x && x.list === names[0].list && x.short === names[0].short)) {
      const L = names[0].list.length;
      let step = 1;
      if (k > 1) { step = ((names[1].i - names[0].i) % L + L) % L || L; if (names.some((x, i) => i && ((x.i - names[i - 1].i) % L + L) % L !== step % L)) step = null; }
      if (step !== null) {
        return Array.from({ length: n }, (_, i) => {
          const pos = ((names[0].i + step * at(i)) % L + L) % L;
          const word = names[0].short ? names[0].list[pos].slice(0, 3) : names[0].list[pos];
          return styleLike(String(source[0].raw).trim(), word);
        });
      }
    }
    // Text ending in a number ("Item 1", "Q3", "Week 07"): the number counts on.
    const parts = source.map(c => /^(.*?)(\d+)$/.exec(String(c.raw)));
    if (parts.every(m => m && m[1] === parts[0][1])) {
      const ys = parts.map(m => +m[2]);
      const f = k === 1 ? (i => ys[0] + i) : trend(ys);
      const width = parts[0][2].length;
      return Array.from({ length: n }, (_, i) => {
        const v = Math.round(f(at(i)));
        return parts[0][1] + (v < 0 ? '-' + String(-v).padStart(width, '0') : String(v).padStart(width, '0'));
      });
    }
    return copyAll();
  }

  /** Fill > Series: a linear or growth series from start, or a date series by day/weekday/month/year. */
  function series(start, n, { type = 'linear', step = 1, stop = null, unit = 'day' } = {}, dateHelpers) {
    const out = [];
    let v = start;
    for (let i = 0; i < n; i++) {
      // Dates count from the start, so Jan 31 by month gives Feb 29, Mar 31.
      if (type === 'date' && dateHelpers) v = i ? dateHelpers.add(start, step * i, unit) : start;
      if (stop !== null && (step >= 0 ? v > stop : v < stop)) break;
      out.push(clean(v));
      if (type === 'growth') v *= step;
      else if (type !== 'date') v += step;
    }
    return out;
  }

  const api = { extend, series };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.SheetsFill = api;
})(typeof window !== 'undefined' ? window : globalThis);
