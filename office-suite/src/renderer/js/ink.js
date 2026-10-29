'use strict';
// Isadek — ink: drawing with a pen, a finger or the mouse, for Notes, Slides and Write.
//
// A stroke is { c: colour, w: width, o: opacity, d: SVG path, f?: 1 }. A pen that reports pressure
// draws a filled outline that is wider where it was pressed harder (f: 1); otherwise the path is a
// smoothed line drawn with width w. Coordinates are whatever the app's surface uses.
//
// Ink.attach(layer, opts) handles the pointer: opts.tool() says what to do ({ c, w, o } to draw,
// 'eraser', or null), opts.at(e) turns an event into [x, y] on the surface, opts.strokes() lists
// the strokes there, opts.add(stroke) keeps a new one and opts.erase(i) removes one. Once a pen has
// been used, touches scroll instead of drawing (a palm on the screen leaves no marks), unless
// Ink.touchDraws is set.
(function (root) {
  const r1 = n => Math.round(n * 10) / 10;
  const NS = 'http://www.w3.org/2000/svg';

  /** A smoothed path through points [[x, y], …]: curves through the midpoints. */
  function smoothPath(pts) {
    if (!pts.length) return '';
    if (pts.length === 1) return `M${r1(pts[0][0])} ${r1(pts[0][1])} L${r1(pts[0][0] + 0.1)} ${r1(pts[0][1] + 0.1)}`;
    let d = `M${r1(pts[0][0])} ${r1(pts[0][1])}`;
    for (let i = 1; i < pts.length - 1; i++) {
      const [x, y] = pts[i], [nx, ny] = pts[i + 1];
      d += ` Q${r1(x)} ${r1(y)} ${r1((x + nx) / 2)} ${r1((y + ny) / 2)}`;
    }
    const last = pts[pts.length - 1];
    return d + ` L${r1(last[0])} ${r1(last[1])}`;
  }

  /**
   * A filled outline around points [[x, y, pressure], …], half as wide as w at no pressure and as
   * wide as w at full, with round ends.
   */
  function outlinePath(pts, w) {
    // Steady the pressure (pens report it jumpily) and drop points too close together.
    const P = [];
    let pr = pts[0] ? pts[0][2] || 0.5 : 0.5;
    for (const p of pts) {
      pr = pr * 0.6 + (p[2] || 0.5) * 0.4;
      const last = P[P.length - 1];
      if (!last || Math.hypot(p[0] - last[0], p[1] - last[1]) >= 0.8) P.push([p[0], p[1], pr]);
    }
    if (P.length < 2) {
      const [x, y, p] = P[0] || pts[0], r = w * (0.5 + 0.5 * (p || 0.5)) / 2;
      return `M${r1(x - r)} ${r1(y)} Q${r1(x - r)} ${r1(y - r)} ${r1(x)} ${r1(y - r)} Q${r1(x + r)} ${r1(y - r)} ${r1(x + r)} ${r1(y)} Q${r1(x + r)} ${r1(y + r)} ${r1(x)} ${r1(y + r)} Q${r1(x - r)} ${r1(y + r)} ${r1(x - r)} ${r1(y)} Z`;
    }
    const left = [], right = [];
    for (let i = 0; i < P.length; i++) {
      const a = P[Math.max(0, i - 1)], b = P[Math.min(P.length - 1, i + 1)];
      let dx = b[0] - a[0], dy = b[1] - a[1];
      const len = Math.hypot(dx, dy) || 1;
      dx /= len; dy /= len;
      const r = w * (0.5 + 0.5 * P[i][2]) / 2;
      left.push([P[i][0] - dy * r, P[i][1] + dx * r]);
      right.push([P[i][0] + dy * r, P[i][1] - dx * r]);
    }
    const curve = list => list.slice(1).map((p, i) => { const q = list[i]; return ` Q${r1(q[0])} ${r1(q[1])} ${r1((q[0] + p[0]) / 2)} ${r1((q[1] + p[1]) / 2)}`; }).join('');
    const end = P[P.length - 1], start = P[0];
    const rEnd = w * (0.5 + 0.5 * end[2]) / 2, rStart = w * (0.5 + 0.5 * start[2]) / 2;
    const rl = right[right.length - 1], ll = left[left.length - 1];
    let d = `M${r1(left[0][0])} ${r1(left[0][1])}` + curve(left) + ` L${r1(ll[0])} ${r1(ll[1])}`;
    // Round cap at the end: a curve out past the last point to the other side.
    const ex = end[0] + (end[0] - P[P.length - 2][0]) / (Math.hypot(end[0] - P[P.length - 2][0], end[1] - P[P.length - 2][1]) || 1) * rEnd * 1.3;
    const ey = end[1] + (end[1] - P[P.length - 2][1]) / (Math.hypot(end[0] - P[P.length - 2][0], end[1] - P[P.length - 2][1]) || 1) * rEnd * 1.3;
    d += ` Q${r1(ex)} ${r1(ey)} ${r1(rl[0])} ${r1(rl[1])}`;
    const rr = right.slice().reverse();
    d += curve(rr) + ` L${r1(rr[rr.length - 1][0])} ${r1(rr[rr.length - 1][1])}`;
    const sx = start[0] - (P[1][0] - start[0]) / (Math.hypot(P[1][0] - start[0], P[1][1] - start[1]) || 1) * rStart * 1.3;
    const sy = start[1] - (P[1][1] - start[1]) / (Math.hypot(P[1][0] - start[0], P[1][1] - start[1]) || 1) * rStart * 1.3;
    return d + ` Q${r1(sx)} ${r1(sy)} ${r1(left[0][0])} ${r1(left[0][1])} Z`;
  }

  /** Accepts only well-formed strokes (they come from files). */
  function clean(s) {
    if (!s || typeof s.d !== 'string' || !/^[MLQZ0-9 .,-]{1,200000}$/.test(s.d)) return null;
    const c = typeof s.c === 'string' && /^#[0-9a-f]{3,8}$/i.test(s.c) ? s.c : '#1f1f1f';
    const out = { c, w: Math.max(0.5, Math.min(60, +s.w || 2)), o: Math.max(0.1, Math.min(1, +s.o || 1)), d: s.d };
    if (s.f === 1 || s.f === true) out.f = 1;
    return out;
  }

  /** The points of a path (its corners and curve ends). */
  const pointsOf = d => (d.match(/-?[\d.]+ -?[\d.]+/g) || []).map(pt => pt.split(' ').map(Number));

  /** The stroke near (x, y), as an index, or -1. */
  function hit(strokes, x, y, slack = 8) {
    for (let k = strokes.length - 1; k >= 0; k--) {
      const s = strokes[k], pts = pointsOf(s.d), r = slack + (s.f ? 0 : s.w / 2);
      // Distance to each segment (fast strokes have points far apart).
      const near = pts.some((a, i) => {
        const b = pts[i + 1] || a;
        const dx = b[0] - a[0], dy = b[1] - a[1], len = dx * dx + dy * dy;
        const t = len ? Math.max(0, Math.min(1, ((x - a[0]) * dx + (y - a[1]) * dy) / len)) : 0;
        return Math.hypot(a[0] + t * dx - x, a[1] + t * dy - y) < r;
      });
      if (near) return k;
    }
    return -1;
  }

  /** The box around strokes: { x, y, w, h }. */
  function bounds(strokes) {
    let x1 = Infinity, y1 = Infinity, x2 = -Infinity, y2 = -Infinity;
    for (const s of strokes) for (const [x, y] of pointsOf(s.d)) { const r = s.f ? 0 : s.w / 2; x1 = Math.min(x1, x - r); y1 = Math.min(y1, y - r); x2 = Math.max(x2, x + r); y2 = Math.max(y2, y + r); }
    return x1 === Infinity ? { x: 0, y: 0, w: 0, h: 0 } : { x: x1, y: y1, w: x2 - x1, h: y2 - y1 };
  }
  /** Strokes moved by (dx, dy). */
  const moved = (strokes, dx, dy) => strokes.map(s => ({ ...s, d: s.d.replace(/(-?[\d.]+) (-?[\d.]+)/g, (m, x, y) => `${r1(+x + dx)} ${r1(+y + dy)}`) }));

  /** An SVG path element for a stroke. */
  function pathEl(s) {
    const p = document.createElementNS(NS, 'path');
    p.setAttribute('d', s.d);
    if (s.f) { p.setAttribute('fill', s.c); p.setAttribute('stroke', 'none'); if (s.o < 1) p.setAttribute('fill-opacity', s.o); }
    else { p.setAttribute('fill', 'none'); p.setAttribute('stroke', s.c); p.setAttribute('stroke-width', s.w); if (s.o < 1) p.setAttribute('stroke-opacity', s.o); }
    p.setAttribute('stroke-linecap', 'round');
    p.setAttribute('stroke-linejoin', 'round');
    return p;
  }
  /** Strokes as SVG markup (for files and exports). */
  const svgMarkup = (strokes, w, h) => `<svg xmlns="${NS}" viewBox="0 0 ${r1(w)} ${r1(h)}" width="${r1(w)}" height="${r1(h)}">${strokes.map(s => (s.f
    ? `<path d="${s.d}" fill="${s.c}"${s.o < 1 ? ` fill-opacity="${s.o}"` : ''}/>`
    : `<path d="${s.d}" fill="none" stroke="${s.c}" stroke-width="${s.w}"${s.o < 1 ? ` stroke-opacity="${s.o}"` : ''} stroke-linecap="round" stroke-linejoin="round"/>`)).join('')}</svg>`;
  /** Strokes drawn into a canvas context (for PNG pictures). */
  function paint(ctx, strokes) {
    for (const s of strokes) {
      const p = new Path2D(s.d);
      ctx.save();
      ctx.globalAlpha = s.o;
      if (s.f) { ctx.fillStyle = s.c; ctx.fill(p); } else { ctx.strokeStyle = s.c; ctx.lineWidth = s.w; ctx.lineCap = ctx.lineJoin = 'round'; ctx.stroke(p); }
      ctx.restore();
    }
  }

  let penSeen = false;
  const api = {
    smoothPath, outlinePath, clean, hit, bounds, moved, pathEl, svgMarkup, paint, pointsOf,
    touchDraws: false,
    get penSeen() { return penSeen; },
    attach(layer, o) {
      layer.addEventListener('pointerdown', e => {
        const tool = o.tool();
        if (!tool) return;
        if (e.pointerType === 'pen') penSeen = true;
        // Palm rejection: with a pen in use, a touch scrolls.
        if (e.pointerType === 'touch' && penSeen && !api.touchDraws) return;
        if (e.button > 0 && e.pointerType === 'mouse') return;
        e.preventDefault();
        e.stopPropagation();
        try { layer.setPointerCapture(e.pointerId); } catch { /* not capturable */ }
        const id = e.pointerId;
        if (tool === 'eraser') {
          const rub = ev => { const [x, y] = o.at(ev); const i = hit(o.strokes(), x, y, o.slack || 8); if (i >= 0) o.erase(i); };
          rub(e);
          const move = ev => { if (ev.pointerId === id) rub(ev); };
          const up = ev => { if (ev.pointerId !== id) return; layer.removeEventListener('pointermove', move); layer.removeEventListener('pointerup', up); layer.removeEventListener('pointercancel', up); };
          layer.addEventListener('pointermove', move);
          layer.addEventListener('pointerup', up);
          layer.addEventListener('pointercancel', up);
          return;
        }
        // Pressure counts only from a pen (mice report 0.5, touch 0 or 1).
        const pressured = e.pointerType === 'pen' && !tool.hl;
        const pts = [[...o.at(e), e.pressure || 0.5]];
        const live = pathEl({ ...tool, d: smoothPath(pts), f: 0 });
        (o.liveLayer ? o.liveLayer() : layer).appendChild(live);
        const make = () => (pressured ? { c: tool.c, w: tool.w * 1.6, o: tool.o, d: outlinePath(pts, tool.w * 1.6), f: 1 } : { c: tool.c, w: tool.w, o: tool.o, d: smoothPath(pts) });
        const move = ev => {
          if (ev.pointerId !== id) return;
          // Every point the device reported since the last event, for smooth fast strokes.
          const evs = ev.getCoalescedEvents ? ev.getCoalescedEvents() : [ev];
          for (const c of (evs.length ? evs : [ev])) {
            const [x, y] = o.at(c), last = pts[pts.length - 1];
            if (Math.hypot(x - last[0], y - last[1]) < 1) continue;
            pts.push([x, y, c.pressure || 0.5]);
          }
          const s = make();
          live.setAttribute('d', s.d);
          if (s.f) { live.setAttribute('fill', s.c); live.setAttribute('stroke', 'none'); if (s.o < 1) live.setAttribute('fill-opacity', s.o); }
        };
        const up = ev => {
          if (ev.pointerId !== id) return;
          layer.removeEventListener('pointermove', move);
          layer.removeEventListener('pointerup', up);
          layer.removeEventListener('pointercancel', up);
          live.remove();
          o.add(make());
        };
        layer.addEventListener('pointermove', move);
        layer.addEventListener('pointerup', up);
        layer.addEventListener('pointercancel', up);
      });
    }
  };
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.Ink = api;
})(typeof window !== 'undefined' ? window : globalThis);
