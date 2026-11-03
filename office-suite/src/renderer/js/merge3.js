'use strict';
// Three-way merge of two edits to the same value, for live editing: when someone else's change to
// a paragraph, cell or shape arrives while yours to it is still on its way, both are kept.
//
// merge3(base, ours, theirs, textFields) works on JSON values. Objects merge field by field; the
// fields named in textFields (a paragraph's HTML, a text box's) merge by characters and tags, so two
// people typing in different parts of one paragraph both keep their words. Where both changed the
// very same characters (or typed at the very same spot), or any other value (a cell, a size), the
// change the server put later wins, as it would have anyway. So both people's apps work out the
// same merge. oursEarlier says whose change the server put first.
(function (root) {
  const enc = v => JSON.stringify(v === undefined ? null : v);
  const same = (a, b) => enc(a) === enc(b);
  // Tags and entities whole, everything else character by character (people type mid-word).
  const TOKENS = /<[^>]*>|&#?\w+;|[\s\S]/g;
  const tokens = s => s.match(TOKENS) || [];

  /** Where x differs from base: [{ s, e, rep }] — base tokens s..e-1 became rep. */
  function hunks(base, x) {
    const n = base.length, m = x.length;
    // Common ends first (most edits touch one spot), then the middle by longest common subsequence.
    let pre = 0;
    while (pre < n && pre < m && base[pre] === x[pre]) pre++;
    let suf = 0;
    while (suf < n - pre && suf < m - pre && base[n - 1 - suf] === x[m - 1 - suf]) suf++;
    const A = base.slice(pre, n - suf), B = x.slice(pre, m - suf);
    if (!A.length && !B.length) return [];
    if (A.length * B.length > 4e6) return [{ s: pre, e: n - suf, rep: B }];
    const L = Array.from({ length: A.length + 1 }, () => new Uint32Array(B.length + 1));
    for (let i = A.length - 1; i >= 0; i--) for (let j = B.length - 1; j >= 0; j--) L[i][j] = A[i] === B[j] ? L[i + 1][j + 1] + 1 : Math.max(L[i + 1][j], L[i][j + 1]);
    const out = [];
    let i = 0, j = 0, cur = null;
    const open = () => { if (!cur) cur = { s: pre + i, e: pre + i, rep: [] }; };
    const close = () => { if (cur) { out.push(cur); cur = null; } };
    while (i < A.length || j < B.length) {
      if (i < A.length && j < B.length && A[i] === B[j]) { close(); i++; j++; }
      else if (j < B.length && (i >= A.length || L[i][j + 1] >= L[i + 1][j])) { open(); cur.rep.push(B[j++]); }
      else { open(); i++; cur.e = pre + i; }
    }
    close();
    return out;
  }

  /** Merges two edits of a text. */
  function mergeText(base, ours, theirs, oursEarlier = false) {
    if (ours === theirs) return ours;
    if (ours === base) return theirs;
    if (theirs === base) return ours;
    const b = tokens(base), ho = hunks(b, tokens(ours)), ht = hunks(b, tokens(theirs));
    // Walk both lists of changes along the original; overlapping changes are a conflict (ours wins).
    // In order along the original; at one spot insertions come before replacements, theirs first.
    const all = [...ho.map(h => ({ ...h, who: 'o' })), ...ht.map(h => ({ ...h, who: 't' }))]
      .sort((x, y) => x.s - y.s || (x.e - x.s) - (y.e - y.s) || ((x.who === 'o') === oursEarlier ? -1 : 1));
    const groups = [];
    for (const h of all) {
      const g = groups[groups.length - 1];
      // Changes conflict when their ranges overlap, or when both insert at the same spot.
      const both = g && g.s === g.e && h.s === h.e && h.s === g.s;
      if (g && (h.s < g.e || both)) { g.e = Math.max(g.e, h.e); g.items.push(h); }
      else groups.push({ s: h.s, e: h.e, items: [h] });
    }
    const out = [];
    let pos = 0;
    for (const g of groups) {
      out.push(...b.slice(pos, g.s));
      const o = g.items.filter(h => h.who === 'o'), t = g.items.filter(h => h.who === 't');
      if (!t.length || !o.length) {
        // Only one side changed here.
        let p = g.s;
        for (const h of g.items) { out.push(...b.slice(p, h.s), ...h.rep); p = h.e; }
        out.push(...b.slice(p, g.e));
      } else {
        // Both changed the same characters, or typed at the very same spot: the later change. (So
        // text is never doubled, even when one side's starting point was out of date.)
        const win = oursEarlier ? t : o;
        let p = g.s;
        for (const h of win) { out.push(...b.slice(p, h.s), ...h.rep); p = h.e; }
        out.push(...b.slice(p, g.e));
      }
      pos = g.e;
    }
    out.push(...b.slice(pos));
    return out.join('');
  }

  /** Merges two edits of a JSON value (undefined or null means "not there"). */
  function merge3(base, ours, theirs, textFields, field = '', oursEarlier = false) {
    if (same(ours, theirs)) return ours;
    if (same(base, ours)) return theirs;
    if (same(base, theirs)) return ours;
    if (typeof ours === 'string' && typeof theirs === 'string' && typeof base === 'string' && textFields && textFields.includes(field)) return mergeText(base, ours, theirs, oursEarlier);
    const obj = v => v && typeof v === 'object' && !Array.isArray(v);
    if (obj(ours) && obj(theirs)) {
      const b = obj(base) ? base : {}, out = {};
      for (const k of new Set([...Object.keys(ours), ...Object.keys(theirs)])) {
        const v = merge3(b[k], ours[k], theirs[k], textFields, k, oursEarlier);
        if (v !== undefined && v !== null) out[k] = v;
      }
      return out;
    }
    // Changed differently on both sides (or deleted on one): the later change.
    return oursEarlier ? theirs : ours;
  }

  const api = { merge3, mergeText, tokens, hunks };
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.Merge3 = api;
})(typeof window !== 'undefined' ? window : globalThis);
