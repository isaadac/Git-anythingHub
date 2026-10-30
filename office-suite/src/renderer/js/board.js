'use strict';
// Isadek Whiteboard: an endless board for ideas and diagrams — sticky notes, text, shapes,
// frames, connectors, pictures and drawing — with diagram helpers (+ handles, mind maps with Tab
// and Enter, Tidy Up), templates, live editing and export as a picture, SVG or PDF.
//
// The board is { bg, items: [...] }; an item is { id, type, x, y, w, h, … } in board units:
//   sticky { color, text } · text { text, size, color } · shape { shape, fill, line, text, size }
//   frame { title, fill } · image { src } · ink { d, c, sw, o, f, bx, by, bw, bh }
//   connector { from: { id }, to: { id }, x1, y1, x2, y2, kind, arrow, color }
// Everything is drawn as SVG (text as <text> wrapped here), so an export is the same picture.
(function () {
  const $ = id => document.getElementById(id);
  const NS = 'http://www.w3.org/2000/svg';
  const state = UI.docState('Isadek Whiteboard', 'Board.aboard');
  const svg = $('board'), world = $('world'), itemsG = $('items'), overlay = $('overlay'), peersG = $('peers'), editor = $('editor');
  UI.initRibbon();

  // ---------- The model ----------
  const SHAPES = ['rect', 'round', 'ellipse', 'diamond', 'triangle', 'parallelogram', 'cylinder', 'hexagon'];
  const STICKY = ['#fff176', '#ffcc80', '#f8bbd0', '#c5e1a5', '#b3e5fc', '#e1bee7', '#eeeeee', '#ffffff'];
  const INK = ['#1f1f1f', '#e53935', '#1e88e5', '#43a047', '#fb8c00', '#8e24aa'];
  const SHAPE_FILL = '#e8f0fe', SHAPE_LINE = '#1a73e8';
  const LIMITS = { items: 20000, text: 5000, path: 200000, image: 15e6 };
  let board = { bg: 'dots', items: [] };
  let uid = 0;
  const newId = () => 'b' + Date.now().toString(36) + (uid++).toString(36) + Math.random().toString(36).slice(2, 5);
  const ID_OK = /^[\w-]{1,60}$/;
  const color = c => (typeof c === 'string' && /^#[0-9a-f]{6}$/i.test(c) ? c.toLowerCase() : null);
  const num = (v, d, lo = -1e7, hi = 1e7) => (typeof v === 'number' && isFinite(v) ? Math.min(hi, Math.max(lo, v)) : d);
  const str = (v, n) => (typeof v === 'string' ? v.replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '').slice(0, n) : '');
  const oneOf = (v, list, d) => (list.includes(v) ? v : d);
  const byId = id => board.items.find(i => i.id === id);

  /** An item from a file or another person: known fields only, within limits. */
  function cleanItem(it, keepId = true) {
    if (!it || typeof it !== 'object') return null;
    const type = oneOf(it.type, ['sticky', 'text', 'shape', 'frame', 'image', 'ink', 'connector'], null);
    if (!type) return null;
    const out = { id: keepId && typeof it.id === 'string' && ID_OK.test(it.id) ? it.id : newId(), type, x: num(it.x, 0), y: num(it.y, 0), w: num(it.w, 160, 1, 1e6), h: num(it.h, 100, 1, 1e6) };
    if (type === 'sticky') { out.color = color(it.color) || STICKY[0]; out.text = str(it.text, LIMITS.text); }
    if (type === 'text') { out.text = str(it.text, LIMITS.text); out.size = num(it.size, 20, 6, 400); out.color = color(it.color) || INK[0]; if (it.bold === true) out.bold = true; }
    if (type === 'shape') { out.shape = oneOf(it.shape, SHAPES, 'rect'); out.fill = color(it.fill) || SHAPE_FILL; out.line = color(it.line) || SHAPE_LINE; out.text = str(it.text, LIMITS.text); out.size = num(it.size, 16, 6, 200); }
    if (type === 'frame') { out.title = str(it.title, 200) || 'Frame'; out.fill = color(it.fill) || '#ffffff'; }
    if (type === 'image') { if (typeof it.src !== 'string' || !/^data:image\/(png|jpe?g|gif|webp);base64,[a-z0-9+/=]+$/i.test(it.src) || it.src.length > LIMITS.image) return null; out.src = it.src; }
    if (type === 'ink') {
      if (typeof it.d !== 'string' || !/^[MLQCZ0-9.\s,eE+-]+$/i.test(it.d) || it.d.length > LIMITS.path) return null;
      Object.assign(out, { d: it.d, c: color(it.c) || INK[0], sw: num(it.sw, 3, 0.5, 200), o: num(it.o, 1, 0.05, 1), bx: num(it.bx, out.x), by: num(it.by, out.y), bw: num(it.bw, out.w, 1, 1e6), bh: num(it.bh, out.h, 1, 1e6) });
      if (it.f === 1) out.f = 1;
    }
    if (type === 'connector') {
      for (const k of ['from', 'to']) if (it[k] && typeof it[k] === 'object' && typeof it[k].id === 'string' && ID_OK.test(it[k].id)) out[k] = { id: it[k].id };
      ['x1', 'y1', 'x2', 'y2'].forEach(k => { out[k] = num(it[k], 0); });
      out.kind = oneOf(it.kind, ['elbow', 'straight', 'curve'], 'elbow');
      out.arrow = oneOf(it.arrow, ['end', 'both', 'none'], 'end');
      out.color = color(it.color) || '#5f6368';
    }
    return out;
  }
  function cleanBoard(b) {
    if (!b || typeof b !== 'object' || !Array.isArray(b.items)) throw new Error('Not a whiteboard');
    const items = b.items.slice(0, LIMITS.items).map(i => cleanItem(i)).filter(Boolean);
    const ids = new Set(items.map(i => i.id));
    // Connectors whose ends are gone keep their last points.
    items.forEach(i => { if (i.type === 'connector') ['from', 'to'].forEach(k => { if (i[k] && !ids.has(i[k].id)) delete i[k]; }); });
    return { bg: oneOf(b.bg, ['dots', 'grid', 'plain'], 'dots'), items };
  }

  // ---------- Undo ----------
  const undoStack = [], redoStack = [];
  function snapshot() { undoStack.push(JSON.stringify(board)); if (undoStack.length > 100) undoStack.shift(); redoStack.length = 0; }
  function undo() { if (!undoStack.length) return; redoStack.push(JSON.stringify(board)); board = JSON.parse(undoStack.pop()); sel.clear(); changed(); }
  function redo() { if (!redoStack.length) return; undoStack.push(JSON.stringify(board)); board = JSON.parse(redoStack.pop()); sel.clear(); changed(); }

  // ---------- The view: pan and zoom ----------
  const view = { x: 0, y: 0, k: 1 };
  const applyView = () => {
    world.setAttribute('transform', `translate(${view.x} ${view.y}) scale(${view.k})`);
    for (const p of svg.querySelectorAll('pattern')) p.setAttribute('patternTransform', `translate(${view.x} ${view.y}) scale(${view.k})`);
    $('stZoom').textContent = Math.round(view.k * 100) + '%';
    if (editing) placeEditor();
    if (peersNow.length) drawPeers(peersNow);
  };
  const toBoard = (cx, cy) => { const r = svg.getBoundingClientRect(); return [(cx - r.left - view.x) / view.k, (cy - r.top - view.y) / view.k]; };
  function zoomAt(k, cx, cy) {
    k = Math.min(8, Math.max(0.05, k));
    const r = svg.getBoundingClientRect();
    const px = cx - r.left, py = cy - r.top;
    view.x = px - (px - view.x) * (k / view.k);
    view.y = py - (py - view.y) * (k / view.k);
    view.k = k;
    applyView();
  }
  const center = () => { const r = svg.getBoundingClientRect(); return toBoard(r.left + r.width / 2, r.top + r.height / 2); };
  function bounds(list = board.items) {
    let x1 = Infinity, y1 = Infinity, x2 = -Infinity, y2 = -Infinity;
    for (const i of list) {
      const b = boxOf(i);
      x1 = Math.min(x1, b.x); y1 = Math.min(y1, b.y); x2 = Math.max(x2, b.x + b.w); y2 = Math.max(y2, b.y + b.h);
    }
    return isFinite(x1) ? { x: x1, y: y1, w: x2 - x1, h: y2 - y1 } : null;
  }
  function zoomFit() {
    const b = bounds();
    const r = svg.getBoundingClientRect();
    if (!b) { view.k = 1; view.x = r.width / 2; view.y = r.height / 2; return applyView(); }
    const k = Math.min(2, Math.max(0.05, Math.min((r.width - 80) / b.w, (r.height - 80) / b.h)));
    view.k = k;
    view.x = (r.width - b.w * k) / 2 - b.x * k;
    view.y = (r.height - b.h * k) / 2 - b.y * k;
    applyView();
  }

  // ---------- Text: wrapped with the same measurements everywhere ----------
  const measureCtx = document.createElement('canvas').getContext('2d');
  const FONT = 'Inter, "Segoe UI", system-ui, -apple-system, Roboto, sans-serif';
  function wrap(text, width, size, bold) {
    measureCtx.font = `${bold ? 600 : 400} ${size}px ${FONT}`;
    const out = [];
    for (const para of String(text).split('\n')) {
      let line = '';
      for (const word of para.split(/(\s+)/)) {
        const next = line + word;
        if (measureCtx.measureText(next).width <= width || !line.trim()) {
          // A word longer than the line breaks where it must.
          if (measureCtx.measureText(next).width > width && !line.trim()) {
            let part = '';
            for (const ch of next) { if (measureCtx.measureText(part + ch).width > width && part) { out.push(part); part = ''; } part += ch; }
            line = part;
          } else line = next;
        } else { out.push(line.trimEnd()); line = word.trimStart(); }
      }
      out.push(line.trimEnd());
    }
    return out;
  }
  function textEl(text, x, y, w, h, { size = 16, color: col = '#1f1f1f', align = 'middle', valign = 'middle', bold = false, pad = 10 } = {}) {
    const t = document.createElementNS(NS, 'text');
    const lines = wrap(text, Math.max(10, w - pad * 2), size, bold);
    const lh = size * 1.3;
    const total = lines.length * lh;
    const top = valign === 'top' ? y + pad : y + (h - total) / 2;
    const ax = align === 'start' ? x + pad : x + w / 2;
    t.setAttribute('font-size', size);
    t.setAttribute('font-family', FONT);
    if (bold) t.setAttribute('font-weight', '600');
    t.setAttribute('fill', col);
    t.setAttribute('text-anchor', align);
    lines.forEach((l, i) => {
      const s = document.createElementNS(NS, 'tspan');
      s.setAttribute('x', ax);
      s.setAttribute('y', top + i * lh + size * 0.95);
      s.textContent = l || ' ';
      t.appendChild(s);
    });
    return t;
  }
  /** How tall a text item must be for its words. */
  const textHeight = it => Math.max(it.size * 1.3 + 8, wrap(it.text || ' ', it.w - 8, it.size, it.bold).length * it.size * 1.3 + 8);

  // ---------- Drawing ----------
  const el = (tag, attrs = {}) => { const e = document.createElementNS(NS, tag); for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, v); return e; };
  function shapePath(s, x, y, w, h) {
    const r = Math.min(w, h);
    switch (s) {
      case 'round': return `M${x + r * 0.25},${y} H${x + w - r * 0.25} A${r * 0.25},${r * 0.25} 0 0 1 ${x + w},${y + r * 0.25} V${y + h - r * 0.25} A${r * 0.25},${r * 0.25} 0 0 1 ${x + w - r * 0.25},${y + h} H${x + r * 0.25} A${r * 0.25},${r * 0.25} 0 0 1 ${x},${y + h - r * 0.25} V${y + r * 0.25} A${r * 0.25},${r * 0.25} 0 0 1 ${x + r * 0.25},${y} Z`;
      case 'ellipse': return `M${x},${y + h / 2} A${w / 2},${h / 2} 0 1 0 ${x + w},${y + h / 2} A${w / 2},${h / 2} 0 1 0 ${x},${y + h / 2} Z`;
      case 'diamond': return `M${x + w / 2},${y} L${x + w},${y + h / 2} L${x + w / 2},${y + h} L${x},${y + h / 2} Z`;
      case 'triangle': return `M${x + w / 2},${y} L${x + w},${y + h} L${x},${y + h} Z`;
      case 'parallelogram': return `M${x + w * 0.18},${y} H${x + w} L${x + w * 0.82},${y + h} H${x} Z`;
      case 'hexagon': return `M${x + w * 0.2},${y} H${x + w * 0.8} L${x + w},${y + h / 2} L${x + w * 0.8},${y + h} H${x + w * 0.2} L${x},${y + h / 2} Z`;
      case 'cylinder': { const e = Math.min(h * 0.15, 24); return `M${x},${y + e} A${w / 2},${e} 0 0 1 ${x + w},${y + e} V${y + h - e} A${w / 2},${e} 0 0 1 ${x},${y + h - e} Z M${x},${y + e} A${w / 2},${e} 0 0 0 ${x + w},${y + e}`; }
      default: return `M${x},${y} H${x + w} V${y + h} H${x} Z`;
    }
  }
  const SIDES = ['n', 'e', 's', 'w'];
  const sidePoint = (b, s) => ({ n: [b.x + b.w / 2, b.y], s: [b.x + b.w / 2, b.y + b.h], e: [b.x + b.w, b.y + b.h / 2], w: [b.x, b.y + b.h / 2] }[s]);
  const boxOf = i => (i.type === 'connector' ? connectorBox(i) : { x: i.x, y: i.y, w: i.w, h: i.h });
  function facing(a, b) {
    const dx = (b.x + b.w / 2) - (a.x + a.w / 2), dy = (b.y + b.h / 2) - (a.y + a.h / 2);
    return Math.abs(dx) * a.h >= Math.abs(dy) * a.w ? (dx >= 0 ? ['e', 'w'] : ['w', 'e']) : (dy >= 0 ? ['s', 'n'] : ['n', 's']);
  }
  /** A connector's path and its two ends. */
  function connectorGeom(c) {
    const a = c.from && byId(c.from.id), b = c.to && byId(c.to.id);
    let p1 = [c.x1, c.y1], p2 = [c.x2, c.y2], s1 = null, s2 = null;
    if (a && b) { [s1, s2] = facing(a, b); p1 = sidePoint(a, s1); p2 = sidePoint(b, s2); }
    else if (a) { s1 = facing(a, { x: p2[0], y: p2[1], w: 0, h: 0 })[0]; p1 = sidePoint(a, s1); }
    else if (b) { s2 = facing(b, { x: p1[0], y: p1[1], w: 0, h: 0 })[0]; p2 = sidePoint(b, s2); }
    const [x1, y1] = p1, [x2, y2] = p2;
    let d;
    if (c.kind === 'straight') d = `M${x1},${y1} L${x2},${y2}`;
    else if (c.kind === 'curve') {
      const off = Math.max(40, Math.hypot(x2 - x1, y2 - y1) / 3);
      const v = s => ({ n: [0, -1], s: [0, 1], e: [1, 0], w: [-1, 0] }[s] || [Math.sign(x2 - x1), 0]);
      const [ax, ay] = v(s1), [bx, by] = s2 ? v(s2) : [-Math.sign(x2 - x1), 0];
      d = `M${x1},${y1} C${x1 + ax * off},${y1 + ay * off} ${x2 + bx * off},${y2 + by * off} ${x2},${y2}`;
    } else {
      const horiz = s1 ? s1 === 'e' || s1 === 'w' : Math.abs(x2 - x1) >= Math.abs(y2 - y1);
      const vert2 = s2 ? s2 === 'n' || s2 === 's' : !horiz;
      if (horiz && !vert2) { const mx = (x1 + x2) / 2; d = `M${x1},${y1} H${mx} V${y2} H${x2}`; }
      else if (!horiz && vert2) { const my = (y1 + y2) / 2; d = `M${x1},${y1} V${my} H${x2} V${y2}`; }
      else if (horiz) d = `M${x1},${y1} H${x2} V${y2}`;
      else d = `M${x1},${y1} V${y2} H${x2}`;
    }
    return { d, x1, y1, x2, y2 };
  }
  function connectorBox(c) { const g = connectorGeom(c); return { x: Math.min(g.x1, g.x2), y: Math.min(g.y1, g.y2), w: Math.max(1, Math.abs(g.x2 - g.x1)), h: Math.max(1, Math.abs(g.y2 - g.y1)) }; }

  function drawItem(i) {
    const g = el('g', { class: 'item item-' + i.type, 'data-id': i.id });
    if (i.type === 'sticky') {
      g.append(el('rect', { x: i.x + 2, y: i.y + 4, width: i.w, height: i.h, rx: 4, class: 'sticky-shadow' }), el('rect', { x: i.x, y: i.y, width: i.w, height: i.h, rx: 4, fill: i.color }));
      g.appendChild(textEl(i.text, i.x, i.y, i.w, i.h, { size: Math.max(12, Math.min(28, i.w / 9)), pad: 14 }));
    } else if (i.type === 'text') {
      g.append(el('rect', { x: i.x, y: i.y, width: i.w, height: i.h, fill: 'transparent' }));
      g.appendChild(textEl(i.text || '', i.x, i.y, i.w, i.h, { size: i.size, color: i.color, align: 'start', valign: 'top', bold: !!i.bold, pad: 4 }));
    } else if (i.type === 'shape') {
      g.append(el('path', { d: shapePath(i.shape, i.x, i.y, i.w, i.h), fill: i.fill, stroke: i.line, 'stroke-width': 2, 'stroke-linejoin': 'round' }));
      if (i.text) g.appendChild(textEl(i.text, i.x + (i.shape === 'parallelogram' ? i.w * 0.1 : 0), i.y + (i.shape === 'cylinder' ? Math.min(i.h * 0.15, 24) : 0), i.w * (i.shape === 'parallelogram' ? 0.8 : i.shape === 'diamond' || i.shape === 'triangle' ? 0.9 : 1), i.h - (i.shape === 'cylinder' ? Math.min(i.h * 0.15, 24) : 0), { size: i.size, pad: i.shape === 'diamond' ? i.w * 0.18 : 10 }));
    } else if (i.type === 'frame') {
      g.append(el('rect', { x: i.x, y: i.y, width: i.w, height: i.h, rx: 8, fill: i.fill, class: 'frame-box' }));
      const t = textEl(i.title, i.x, i.y - 30, i.w, 28, { size: 16, align: 'start', valign: 'top', bold: true, pad: 2, color: '#5f6368' });
      t.classList.add('frame-title');
      g.appendChild(t);
    } else if (i.type === 'image') {
      g.append(el('image', { x: i.x, y: i.y, width: i.w, height: i.h, href: i.src, preserveAspectRatio: 'none' }));
    } else if (i.type === 'ink') {
      const sx = i.w / i.bw, sy = i.h / i.bh;
      const p = el('path', { d: i.d, transform: `translate(${i.x} ${i.y}) scale(${sx} ${sy}) translate(${-i.bx} ${-i.by})` });
      if (i.f) { p.setAttribute('fill', i.c); if (i.o < 1) p.setAttribute('fill-opacity', i.o); } else { p.setAttribute('fill', 'none'); p.setAttribute('stroke', i.c); p.setAttribute('stroke-width', i.sw); p.setAttribute('stroke-linecap', 'round'); p.setAttribute('stroke-linejoin', 'round'); p.setAttribute('vector-effect', 'non-scaling-stroke'); if (i.o < 1) p.setAttribute('stroke-opacity', i.o); }
      g.append(el('rect', { x: i.x, y: i.y, width: i.w, height: i.h, fill: 'transparent', class: 'ink-hit' }), p);
    } else if (i.type === 'connector') {
      const geo = connectorGeom(i);
      g.append(el('path', { d: geo.d, fill: 'none', stroke: 'transparent', 'stroke-width': 14, class: 'conn-hit' }));
      const p = el('path', { d: geo.d, fill: 'none', stroke: i.color, 'stroke-width': 2, 'stroke-linejoin': 'round' });
      if (i.arrow !== 'none') p.setAttribute('marker-end', 'url(#arrowEnd)');
      if (i.arrow === 'both') p.setAttribute('marker-start', 'url(#arrowEnd)');
      g.appendChild(p);
    }
    return g;
  }
  // Frames at the back, connectors at the front of what they join; otherwise in board order.
  const drawOrder = () => [...board.items.filter(i => i.type === 'frame'), ...board.items.filter(i => i.type !== 'frame' && i.type !== 'connector'), ...board.items.filter(i => i.type === 'connector')];
  let raf = 0;
  function render() {
    cancelAnimationFrame(raf);
    raf = 0;
    itemsG.textContent = '';
    for (const i of drawOrder()) itemsG.appendChild(drawItem(i));
    drawOverlay();
    $('empty').hidden = board.items.length > 0;
    $('stItems').textContent = `${board.items.length} item${board.items.length === 1 ? '' : 's'}`;
    svg.querySelector('.bg').style.fill = board.bg === 'plain' ? 'transparent' : board.bg === 'grid' ? 'url(#bgGrid)' : 'url(#bgDots)';
    $('bgSel').value = board.bg;
  }
  const renderSoon = () => { if (!raf) raf = requestAnimationFrame(render); };
  function changed() { render(); state.markDirty(); }

  // ---------- Selection and its handles ----------
  const sel = new Set();
  function drawOverlay() {
    overlay.textContent = '';
    const list = [...sel].map(byId).filter(Boolean);
    if (!list.length) return;
    const k = view.k;
    list.forEach(i => { const b = boxOf(i); overlay.appendChild(el('rect', { x: b.x - 3 / k, y: b.y - 3 / k, width: b.w + 6 / k, height: b.h + 6 / k, class: 'sel-box', 'stroke-width': 1.5 / k })); });
    if (list.length !== 1 || list[0].type === 'connector') return;
    const i = list[0], b = boxOf(i), hs = 8 / k;
    // Resize handles at the corners.
    [['nw', b.x, b.y], ['ne', b.x + b.w, b.y], ['se', b.x + b.w, b.y + b.h], ['sw', b.x, b.y + b.h]].forEach(([h, x, y]) => overlay.appendChild(el('rect', { x: x - hs / 2, y: y - hs / 2, width: hs, height: hs, class: 'handle', 'data-h': h, 'stroke-width': 1.5 / k })));
    // + handles: a connected copy on that side (diagrams, mind maps).
    if (['sticky', 'shape'].includes(i.type)) SIDES.forEach(s => {
      const [x, y] = sidePoint(b, s), off = 22 / k, [dx, dy] = { n: [0, -off], s: [0, off], e: [off, 0], w: [-off, 0] }[s];
      const g = el('g', { class: 'plus', 'data-side': s, transform: `translate(${x + dx} ${y + dy}) scale(${1 / k})` });
      g.append(el('circle', { r: 9 }), el('path', { d: 'M-4.5,0 H4.5 M0,-4.5 V4.5' }));
      const t = el('title'); t.textContent = 'Add a connected ' + (i.type === 'sticky' ? 'note' : 'shape'); g.appendChild(t);
      overlay.appendChild(g);
    });
  }
  function select(ids, add = false) {
    if (!add) sel.clear();
    ids.forEach(id => sel.add(id));
    drawOverlay();
    paintSwatches();
    live && live.moved();
  }

  // ---------- Creating ----------
  let tool = 'select', inkColor = INK[0], stickyColor = STICKY[0];
  function add(item, { edit = false } = {}) {
    snapshot();
    const it = cleanItem(item, true);
    board.items.push(it);
    changed();
    select([it.id]);
    if (edit) startEditing(it);
    return it;
  }
  const defaults = (type, x, y, extra = {}) => {
    if (type === 'sticky') return { type, x: x - 90, y: y - 90, w: 180, h: 180, color: stickyColor, text: '', ...extra };
    if (type === 'text') return { type, x, y: y - 14, w: 280, h: 34, text: '', size: 22, color: inkColor, ...extra };
    if (type === 'frame') return { type, x: x - 400, y: y - 250, w: 800, h: 500, title: 'Frame', ...extra };
    if (type.startsWith('shape:')) return { type: 'shape', shape: type.slice(6), x: x - 80, y: y - 50, w: 160, h: type.endsWith('diamond') ? 120 : 100, text: '', ...extra };
    return null;
  };
  /** + handle / Tab: a copy of item `src` on side `s`, joined to it by a connector. */
  function addConnected(src, s, opts = {}) {
    const gap = opts.gap || 90;
    const b = boxOf(src);
    let x = b.x, y = b.y;
    if (s === 'e') x = b.x + b.w + gap; if (s === 'w') x = b.x - b.w - gap; if (s === 's') y = b.y + b.h + gap; if (s === 'n') y = b.y - b.h - gap;
    // Don't land on another item: step along until free.
    const clash = (x, y) => board.items.some(o => o.type !== 'connector' && o.type !== 'frame' && o.id !== src.id && x < o.x + o.w && x + b.w > o.x && y < o.y + o.h && y + b.h > o.y);
    for (let n = 0; n < 50 && clash(x, y); n++) { if (s === 'e' || s === 'w') y += b.h + 30; else x += b.w + 30; }
    snapshot();
    const copy = cleanItem({ ...src, id: newId(), x, y, text: '' }, true);
    board.items.push(copy);
    board.items.push(cleanItem({ id: newId(), type: 'connector', from: { id: src.id }, to: { id: copy.id }, kind: opts.kind || $('connKind').value, arrow: opts.arrow || $('connArrow').value }, true));
    changed();
    select([copy.id]);
    startEditing(copy);
    return copy;
  }
  const parentOf = id => { const c = board.items.find(i => i.type === 'connector' && i.to && i.to.id === id && i.from); return c ? byId(c.from.id) : null; };
  const childrenOf = id => board.items.filter(i => i.type === 'connector' && i.from && i.from.id === id && i.to).map(c => byId(c.to.id)).filter(Boolean);

  // ---------- Editing text ----------
  let editing = null;
  function placeEditor() {
    const i = editing && byId(editing.id);
    if (!i) return;
    const r = svg.getBoundingClientRect(), wr = svg.parentElement.getBoundingClientRect();
    const top = i.type === 'frame' ? i.y - 30 : i.y, h = i.type === 'frame' ? 28 : i.h;
    Object.assign(editor.style, { left: (r.left - wr.left + view.x + i.x * view.k) + 'px', top: (r.top - wr.top + view.y + top * view.k) + 'px', width: i.w * view.k + 'px', height: h * view.k + 'px', fontSize: (i.type === 'text' ? i.size : i.type === 'shape' ? i.size : i.type === 'frame' ? 16 : Math.max(12, Math.min(28, i.w / 9))) * view.k + 'px', textAlign: i.type === 'text' || i.type === 'frame' ? 'left' : 'center', background: i.type === 'sticky' ? i.color : i.type === 'shape' ? i.fill : 'var(--surface)' });
  }
  function startEditing(i) {
    if (!['sticky', 'text', 'shape', 'frame'].includes(i.type)) return;
    editing = { id: i.id, before: i.type === 'frame' ? i.title : i.text };
    editor.value = editing.before;
    editor.hidden = false;
    placeEditor();
    // The item's own words hide while they're edited.
    const node = itemsG.querySelector(`[data-id="${i.id}"] text`);
    if (node) node.style.visibility = 'hidden';
    editor.focus();
    editor.select();
  }
  function stopEditing(keep = true) {
    if (!editing) return;
    const i = byId(editing.id), was = editing;
    editing = null;
    editor.hidden = true;
    if (i && keep) {
      const v = str(editor.value, i.type === 'frame' ? 200 : LIMITS.text);
      if (v !== was.before) {
        snapshot();
        if (i.type === 'frame') i.title = v.replace(/\n/g, ' ') || 'Frame'; else i.text = v;
        if (i.type === 'text') { if (!v.trim()) board.items.splice(board.items.indexOf(i), 1); else i.h = textHeight(i); }
        changed();
      } else if (i.type === 'text' && !v.trim()) { board.items.splice(board.items.indexOf(i), 1); changed(); }
    }
    render();
    svg.focus({ preventScroll: true });
  }
  editor.addEventListener('blur', () => stopEditing(true));
  editor.addEventListener('keydown', e => {
    e.stopPropagation();
    if (e.key === 'Escape') { e.preventDefault(); stopEditing(true); }
    // In a mind map, Tab finishes this node and starts a branch; Enter (without Shift) a sibling.
    if (e.key === 'Tab' && editing) { e.preventDefault(); const i = byId(editing.id); stopEditing(true); if (i && byId(i.id)) branch(i); }
    if (e.key === 'Enter' && !e.shiftKey && editing) { const i = byId(editing.id); if (i && i.type !== 'text' && parentOf(i.id)) { e.preventDefault(); stopEditing(true); sibling(i); } }
    if (editing && byId(editing.id) && byId(editing.id).type === 'text') { const i = byId(editing.id); requestAnimationFrame(() => { if (editing && editing.id === i.id) { editor.style.height = Math.max(i.h, wrap(editor.value || ' ', i.w - 8, i.size).length * i.size * 1.3 + 8) * view.k + 'px'; } }); }
  });
  function branch(i) { const kids = childrenOf(i.id); const s = kids.length && kids[0].x < i.x ? 'w' : 'e'; addConnected(i, s, { kind: 'curve', arrow: 'none' }); }
  function sibling(i) { const p = parentOf(i.id); if (!p) return; const s = i.x >= p.x + p.w ? 'e' : i.x + i.w <= p.x ? 'w' : i.y > p.y ? 's' : 'n'; addConnected(p, s, { kind: 'curve', arrow: 'none' }); }

  // ---------- Pointer ----------
  const hitItem = e => { const g = e.target.closest && e.target.closest('.item'); return g ? byId(g.dataset.id) : null; };
  const pointers = new Map();
  let drag = null, spaceDown = false, lastMove = 0, cursor = null;
  svg.addEventListener('pointerdown', e => {
    if (editing) stopEditing(true);
    svg.focus({ preventScroll: true });
    pointers.set(e.pointerId, [e.clientX, e.clientY]);
    // Two fingers: pinch to zoom and move.
    if (pointers.size === 2) {
      const [a, b] = [...pointers.values()];
      drag = { type: 'pinch', d: Math.hypot(a[0] - b[0], a[1] - b[1]), k: view.k, mid: [(a[0] + b[0]) / 2, (a[1] + b[1]) / 2], vx: view.x, vy: view.y };
      return;
    }
    if (e.button === 1 || spaceDown || tool === 'hand' || (e.pointerType === 'touch' && tool === 'select' && !hitItem(e))) {
      drag = { type: 'pan', sx: e.clientX, sy: e.clientY, vx: view.x, vy: view.y };
      svg.setPointerCapture(e.pointerId);
      svg.classList.add('panning');
      return;
    }
    if (e.button !== 0) return;
    const [x, y] = toBoard(e.clientX, e.clientY);
    if (tool === 'pen' || tool === 'highlighter') return; // ink.js draws
    if (tool === 'eraser') {
      // The eraser takes away drawings it touches (whole strokes).
      drag = { type: 'erase', done: false };
      svg.setPointerCapture(e.pointerId);
      eraseAt(e);
      return;
    }
    const plus = e.target.closest && e.target.closest('.plus');
    if (plus) { const i = byId([...sel][0]); if (i) addConnected(i, plus.dataset.side); e.preventDefault(); return; }
    const handle = e.target.closest && e.target.closest('.handle');
    if (handle) {
      const i = byId([...sel][0]);
      drag = { type: 'resize', h: handle.dataset.h, id: i.id, start: { x: i.x, y: i.y, w: i.w, h: i.h }, snapped: false };
      svg.setPointerCapture(e.pointerId);
      return;
    }
    if (tool === 'connector') {
      const from = hitItem(e);
      drag = { type: 'connect', from: from && from.type !== 'connector' ? from.id : null, x1: x, y1: y };
      drag.preview = el('path', { class: 'conn-preview', 'stroke-width': 2 / view.k, fill: 'none' });
      overlay.appendChild(drag.preview);
      svg.setPointerCapture(e.pointerId);
      return;
    }
    if (tool !== 'select') {
      // Create: click for the default size, or drag to size it.
      drag = { type: 'create', tool, x, y };
      drag.preview = el('rect', { class: 'create-preview', x, y, width: 0, height: 0, 'stroke-width': 1.5 / view.k });
      overlay.appendChild(drag.preview);
      svg.setPointerCapture(e.pointerId);
      return;
    }
    const it = hitItem(e);
    if (it) {
      if (e.shiftKey || e.ctrlKey || e.metaKey) { if (sel.has(it.id)) sel.delete(it.id); else sel.add(it.id); drawOverlay(); paintSwatches(); return; }
      if (!sel.has(it.id)) select([it.id]);
      // What moves: the selection, and what sits inside a selected frame.
      const moving = new Set(sel);
      [...sel].map(byId).filter(f => f && f.type === 'frame').forEach(f => board.items.forEach(o => { if (o.type !== 'connector' && o.id !== f.id && o.x >= f.x && o.y >= f.y && o.x + o.w <= f.x + f.w && o.y + o.h <= f.y + f.h) moving.add(o.id); }));
      drag = { type: 'move', sx: x, sy: y, start: new Map([...moving].map(id => byId(id)).filter(Boolean).map(i => [i.id, { x: i.x, y: i.y, x1: i.x1, y1: i.y1, x2: i.x2, y2: i.y2 }])), snapped: false };
      svg.setPointerCapture(e.pointerId);
      return;
    }
    // Empty board: a box selects what it touches.
    if (!e.shiftKey) select([]);
    drag = { type: 'marquee', x, y, rect: el('rect', { class: 'marquee', x, y, width: 0, height: 0, 'stroke-width': 1 / view.k }) };
    overlay.appendChild(drag.rect);
    svg.setPointerCapture(e.pointerId);
  });
  svg.addEventListener('pointermove', e => {
    if (pointers.has(e.pointerId)) pointers.set(e.pointerId, [e.clientX, e.clientY]);
    const [x, y] = toBoard(e.clientX, e.clientY);
    cursor = { x: Math.round(x), y: Math.round(y) };
    if (live && Date.now() - lastMove > 80) { lastMove = Date.now(); live.moved(); }
    if (!drag) return;
    if (drag.type === 'pinch' && pointers.size >= 2) {
      const [a, b] = [...pointers.values()];
      const d = Math.hypot(a[0] - b[0], a[1] - b[1]), mid = [(a[0] + b[0]) / 2, (a[1] + b[1]) / 2];
      view.x = drag.vx + (mid[0] - drag.mid[0]); view.y = drag.vy + (mid[1] - drag.mid[1]);
      applyView();
      zoomAt(drag.k * d / drag.d, mid[0], mid[1]);
      return;
    }
    if (drag.type === 'erase') { eraseAt(e); return; }
    if (drag.type === 'pan') { view.x = drag.vx + e.clientX - drag.sx; view.y = drag.vy + e.clientY - drag.sy; applyView(); return; }
    if (drag.type === 'move') {
      let dx = x - drag.sx, dy = y - drag.sy;
      if (!drag.snapped && Math.hypot(dx, dy) * view.k < 3) return;
      if (!drag.snapped) { snapshot(); drag.snapped = true; }
      if (!e.altKey) { dx = Math.round(dx / 4) * 4; dy = Math.round(dy / 4) * 4; } // small steps line things up
      for (const [id, s] of drag.start) { const i = byId(id); if (!i) continue; i.x = s.x + dx; i.y = s.y + dy; if (i.type === 'connector') { i.x1 = s.x1 + dx; i.y1 = s.y1 + dy; i.x2 = s.x2 + dx; i.y2 = s.y2 + dy; } }
      renderSoon();
      return;
    }
    if (drag.type === 'resize') {
      const i = byId(drag.id), s = drag.start;
      if (!i) return;
      if (!drag.snapped) { snapshot(); drag.snapped = true; }
      let x1 = s.x, y1 = s.y, x2 = s.x + s.w, y2 = s.y + s.h;
      if (drag.h.includes('w')) x1 = Math.min(x, x2 - 20); if (drag.h.includes('e')) x2 = Math.max(x, x1 + 20);
      if (drag.h.includes('n')) y1 = Math.min(y, y2 - 20); if (drag.h.includes('s')) y2 = Math.max(y, y1 + 20);
      // Pictures, drawings and Shift keep their shape.
      if (e.shiftKey || i.type === 'image' || i.type === 'ink') { const r = s.w / s.h, w = x2 - x1, h = y2 - y1; if (w / h > r) { const nw = h * r; if (drag.h.includes('w')) x1 = x2 - nw; else x2 = x1 + nw; } else { const nh = w / r; if (drag.h.includes('n')) y1 = y2 - nh; else y2 = y1 + nh; } }
      Object.assign(i, { x: x1, y: y1, w: x2 - x1, h: y2 - y1 });
      if (i.type === 'text') i.h = Math.max(i.h, textHeight(i));
      renderSoon();
      return;
    }
    if (drag.type === 'marquee') {
      const rx = Math.min(x, drag.x), ry = Math.min(y, drag.y);
      Object.entries({ x: rx, y: ry, width: Math.abs(x - drag.x), height: Math.abs(y - drag.y) }).forEach(([k, v]) => drag.rect.setAttribute(k, v));
      drag.box = { x: rx, y: ry, w: Math.abs(x - drag.x), h: Math.abs(y - drag.y) };
      return;
    }
    if (drag.type === 'create') {
      Object.entries({ x: Math.min(x, drag.x), y: Math.min(y, drag.y), width: Math.abs(x - drag.x), height: Math.abs(y - drag.y) }).forEach(([k, v]) => drag.preview.setAttribute(k, v));
      drag.to = [x, y];
      return;
    }
    if (drag.type === 'connect') {
      const a = drag.from && byId(drag.from);
      const [sx, sy] = a ? sidePoint(a, facing(a, { x, y, w: 0, h: 0 })[0]) : [drag.x1, drag.y1];
      drag.preview.setAttribute('d', `M${sx},${sy} L${x},${y}`);
      drag.to = [x, y];
    }
  });
  const endPointer = e => {
    pointers.delete(e.pointerId);
    if (!drag) return;
    const d = drag;
    drag = null;
    svg.classList.remove('panning');
    if (d.type === 'move' || d.type === 'resize') { if (d.snapped) changed(); return; }
    if (d.type === 'erase') { if (d.done) changed(); return; }
    if (d.type === 'marquee') {
      d.rect.remove();
      if (d.box && d.box.w > 2) select(board.items.filter(i => { const b = boxOf(i); return b.x < d.box.x + d.box.w && b.x + b.w > d.box.x && b.y < d.box.y + d.box.h && b.y + b.h > d.box.y; }).map(i => i.id), e.shiftKey);
      return;
    }
    if (d.type === 'create') {
      d.preview.remove();
      const [cx, cy] = d.to || [d.x, d.y];
      const big = d.to && Math.abs(cx - d.x) > 12 && Math.abs(cy - d.y) > 12;
      const base = defaults(d.tool, d.x, d.y);
      if (!base) return;
      if (big) Object.assign(base, { x: Math.min(cx, d.x), y: Math.min(cy, d.y), w: Math.abs(cx - d.x), h: Math.abs(cy - d.y) });
      const it = add(base, { edit: d.tool !== 'frame' });
      if (d.tool === 'frame') { const items = board.items.filter(o => o !== it); board.items = [it, ...items]; render(); select([it.id]); }
      setTool('select');
      return;
    }
    if (d.type === 'connect') {
      d.preview.remove();
      const target = document.elementsFromPoint(e.clientX, e.clientY).map(n => n.closest && n.closest('.item')).find(Boolean);
      const to = target && byId(target.dataset.id);
      const [x2, y2] = d.to || [d.x1, d.y1];
      if (Math.hypot(x2 - d.x1, y2 - d.y1) * view.k < 8 && !(to && d.from && to.id !== d.from)) return;
      add({ type: 'connector', from: d.from ? { id: d.from } : undefined, to: to && to.type !== 'connector' && to.id !== d.from ? { id: to.id } : undefined, x1: d.x1, y1: d.y1, x2, y2, kind: $('connKind').value, arrow: $('connArrow').value, color: inkColor === INK[0] ? '#5f6368' : inkColor });
    }
  };
  svg.addEventListener('pointerup', endPointer);
  svg.addEventListener('pointercancel', endPointer);
  svg.addEventListener('dblclick', e => {
    const it = hitItem(e);
    if (it) return startEditing(it);
    const [x, y] = toBoard(e.clientX, e.clientY);
    add(defaults('text', x, y), { edit: true });
  });
  svg.addEventListener('wheel', e => {
    e.preventDefault();
    if (e.ctrlKey || e.metaKey) zoomAt(view.k * Math.exp(-e.deltaY * 0.0025), e.clientX, e.clientY);
    else { view.x -= e.shiftKey && !e.deltaX ? e.deltaY : e.deltaX; view.y -= e.shiftKey && !e.deltaX ? 0 : e.deltaY; applyView(); }
  }, { passive: false });

  // Drawing: pen and highlighter strokes become ink items; the eraser removes whole drawings.
  Ink.attach(svg, {
    tool: () => (tool === 'pen' ? { c: inkColor, w: 3, o: 1 } : tool === 'highlighter' ? { c: inkColor === INK[0] ? '#ffeb3b' : inkColor, w: 18, o: 0.4, hl: true } : null),
    at: e => toBoard(e.clientX, e.clientY),
    liveLayer: () => overlay,
    strokes: () => [],
    add(s) {
      const b = Ink.bounds([s]);
      if (!b) return;
      add({ type: 'ink', d: s.d, c: s.c, sw: s.w, o: s.o, f: s.f, x: b.x, y: b.y, w: Math.max(1, b.w), h: Math.max(1, b.h), bx: b.x, by: b.y, bw: Math.max(1, b.w), bh: Math.max(1, b.h) });
      select([]);
    },
    erase() {}
  });

  /** The eraser: every drawing along the pointer's path since the last event (with some slack). */
  function eraseAt(e) {
    const from = drag.last || [e.clientX, e.clientY], to = [e.clientX, e.clientY];
    drag.last = to;
    const steps = Math.max(1, Math.ceil(Math.hypot(to[0] - from[0], to[1] - from[1]) / 4));
    const hit = new Set();
    for (let n = 0; n <= steps; n++) {
      const x = from[0] + (to[0] - from[0]) * n / steps, y = from[1] + (to[1] - from[1]) * n / steps;
      for (const [dx, dy] of [[0, 0], [5, 0], [-5, 0], [0, 5], [0, -5]]) for (const node of document.elementsFromPoint(x + dx, y + dy)) { const g = node.closest && node.closest('.item-ink'); if (g) hit.add(g.dataset.id); }
    }
    const gone = [...hit].map(byId).filter(Boolean);
    if (!gone.length) return;
    if (!drag.done) { snapshot(); drag.done = true; }
    board.items = board.items.filter(i => !hit.has(i.id));
    gone.forEach(i => sel.delete(i.id));
    render();
    state.markDirty();
  }

  // ---------- Keyboard ----------
  const KEYS = { v: 'select', h: 'hand', n: 'sticky', t: 'text', f: 'frame', r: 'shape:rect', o: 'shape:ellipse', d: 'shape:diamond', l: 'connector', p: 'pen', e: 'eraser' };
  let clip = null;
  svg.addEventListener('keydown', e => {
    if (editing) return;
    const mod = e.ctrlKey || e.metaKey, k = e.key.toLowerCase();
    if (e.key === ' ') { spaceDown = true; svg.classList.add('can-pan'); e.preventDefault(); return; }
    if ((e.key === 'Delete' || e.key === 'Backspace') && sel.size) { e.preventDefault(); removeSelected(); return; }
    if (mod && k === 'a') { e.preventDefault(); select(board.items.map(i => i.id)); return; }
    if (mod && k === 'd') { e.preventDefault(); duplicate(); return; }
    if (mod && k === 'c') { copy(); return; }
    if (mod && k === 'x') { copy(); removeSelected(); return; }
    if (mod && (k === 'z' || k === 'y')) { e.preventDefault(); if (k === 'y' || e.shiftKey) redo(); else undo(); return; }
    if (e.key === 'Escape') { select([]); setTool('select'); return; }
    if (e.key === 'Enter' && sel.size === 1) { e.preventDefault(); const i = byId([...sel][0]); if (e.shiftKey || !parentOf(i.id)) startEditing(i); else sibling(i); return; }
    if (e.key === 'Tab' && sel.size === 1) { e.preventDefault(); const i = byId([...sel][0]); if (i && ['sticky', 'shape', 'text'].includes(i.type)) branch(i); return; }
    if (e.key.startsWith('Arrow') && sel.size) {
      e.preventDefault();
      const step = e.shiftKey ? 40 : 4, [dx, dy] = { ArrowLeft: [-step, 0], ArrowRight: [step, 0], ArrowUp: [0, -step], ArrowDown: [0, step] }[e.key];
      snapshot();
      [...sel].map(byId).filter(Boolean).forEach(i => { i.x += dx; i.y += dy; });
      changed();
      return;
    }
    if (e.shiftKey && e.key === '!') { zoomFit(); return; }
    if (e.shiftKey && e.key === ')') { zoomAt(1, ...screenCenter()); return; }
    if (!mod && (e.key === '+' || e.key === '=')) { zoomAt(view.k * 1.25, ...screenCenter()); return; }
    if (!mod && e.key === '-') { zoomAt(view.k / 1.25, ...screenCenter()); return; }
    if (!mod && !e.altKey && KEYS[k]) setTool(KEYS[k]);
  });
  svg.addEventListener('keyup', e => { if (e.key === ' ') { spaceDown = false; svg.classList.remove('can-pan'); } });
  const screenCenter = () => { const r = svg.getBoundingClientRect(); return [r.left + r.width / 2, r.top + r.height / 2]; };
  function removeSelected() {
    if (!sel.size) return;
    snapshot();
    const gone = new Set(sel);
    board.items = board.items.filter(i => !gone.has(i.id));
    // Connectors to what was removed go too.
    board.items = board.items.filter(i => !(i.type === 'connector' && ((i.from && gone.has(i.from.id)) || (i.to && gone.has(i.to.id)))));
    sel.clear();
    changed();
  }
  /** Copies of items (new ids; connectors between copied items join the copies), offset by d. */
  function cloneItems(list, d) {
    const ids = new Map(list.map(i => [i.id, newId()]));
    return list.map(i => {
      const c = JSON.parse(JSON.stringify(i));
      c.id = ids.get(i.id); c.x += d; c.y += d;
      if (c.type === 'connector') { ['x1', 'x2'].forEach(k => { c[k] += d; }); ['y1', 'y2'].forEach(k => { c[k] += d; }); ['from', 'to'].forEach(k => { if (c[k]) { if (ids.has(c[k].id)) c[k].id = ids.get(c[k].id); else delete c[k]; } }); }
      return cleanItem(c, true);
    }).filter(Boolean);
  }
  function duplicate() {
    const list = [...sel].map(byId).filter(Boolean);
    if (!list.length) return;
    snapshot();
    const copies = cloneItems(list, 24);
    board.items.push(...copies);
    changed();
    select(copies.map(c => c.id));
  }
  function copy() {
    const list = [...sel].map(byId).filter(Boolean);
    if (!list.length) return;
    clip = JSON.stringify(list);
    try { navigator.clipboard.writeText('isadek-board:' + clip).catch(() => {}); } catch { /* clipboard not allowed */ }
  }
  document.addEventListener('paste', e => {
    if (editing || (e.target.closest && e.target.closest('input, textarea'))) return;
    const files = [...(e.clipboardData && e.clipboardData.files || [])].filter(f => /^image\//.test(f.type));
    if (files.length) { e.preventDefault(); files.forEach(f => addImageFile(f)); return; }
    const text = e.clipboardData ? e.clipboardData.getData('text/plain') : '';
    let list = null;
    if (text.startsWith('isadek-board:')) { try { list = JSON.parse(text.slice(13)); } catch { list = null; } } else if (!text && clip) list = JSON.parse(clip);
    if (Array.isArray(list)) {
      e.preventDefault();
      const clean = list.map(i => cleanItem(i, true)).filter(Boolean);
      const b = bounds(clean), [cx, cy] = center();
      const copies = cloneItems(clean, 0).map(c => { const dx = cx - (b.x + b.w / 2), dy = cy - (b.y + b.h / 2); c.x += dx; c.y += dy; if (c.type === 'connector') { c.x1 += dx; c.x2 += dx; c.y1 += dy; c.y2 += dy; } return c; });
      snapshot(); board.items.push(...copies); changed(); select(copies.map(c => c.id));
      return;
    }
    if (text.trim()) {
      e.preventDefault();
      const [x, y] = center();
      // Several lines become sticky notes in a grid; one line a text.
      const lines = text.split(/\r?\n/).map(s => s.trim()).filter(Boolean).slice(0, 100);
      if (lines.length > 1) {
        snapshot();
        const made = lines.map((t, n) => cleanItem({ type: 'sticky', x: x - 300 + (n % 4) * 200, y: y - 200 + Math.floor(n / 4) * 200, w: 180, h: 180, color: stickyColor, text: t }, false));
        board.items.push(...made); changed(); select(made.map(m => m.id));
      } else add({ ...defaults('text', x, y), text: lines[0] }).h = textHeight(byId([...sel][0]));
      changed();
    }
  });

  // ---------- Pictures ----------
  function addImageFile(file, at) {
    if (file.size > 10e6) return UI.toast('That picture is too large (over 10 MB)');
    const r = new FileReader();
    r.onload = () => {
      const img = new Image();
      img.onload = () => {
        const [x, y] = at || center();
        const k = Math.min(1, 600 / Math.max(img.naturalWidth, img.naturalHeight));
        const w = Math.max(20, img.naturalWidth * k), h = Math.max(20, img.naturalHeight * k);
        add({ type: 'image', src: r.result, x: x - w / 2, y: y - h / 2, w, h });
      };
      img.onerror = () => UI.toast('That picture can’t be read');
      img.src = r.result;
    };
    r.readAsDataURL(file);
  }
  $('insPicture').addEventListener('click', () => {
    const inp = document.createElement('input');
    inp.type = 'file'; inp.accept = 'image/png,image/jpeg,image/gif,image/webp';
    inp.onchange = () => { if (inp.files[0]) addImageFile(inp.files[0]); };
    inp.click();
  });
  svg.addEventListener('dragover', e => { if ([...e.dataTransfer.items].some(i => i.kind === 'file')) e.preventDefault(); });
  svg.addEventListener('drop', e => {
    const files = [...e.dataTransfer.files].filter(f => /^image\//.test(f.type));
    if (!files.length) return;
    e.preventDefault();
    files.forEach((f, n) => { const [x, y] = toBoard(e.clientX, e.clientY); addImageFile(f, [x + n * 30, y + n * 30]); });
  });

  // ---------- Tools and colours ----------
  const TOOL_NAMES = { select: 'Select', hand: 'Hand', sticky: 'Sticky Note', text: 'Text', frame: 'Frame', connector: 'Connector', pen: 'Pen', highlighter: 'Highlighter', eraser: 'Eraser' };
  function setTool(t) {
    tool = t;
    document.querySelectorAll('[data-tool]').forEach(b => b.classList.toggle('on', b.dataset.tool === t));
    document.body.dataset.boardTool = t.split(':')[0];
    $('stTool').textContent = TOOL_NAMES[t] || (t.startsWith('shape:') ? 'Shape' : t);
    if (t !== 'select') select([]);
    paintSwatches();
  }
  document.querySelectorAll('[data-tool]').forEach(b => b.addEventListener('click', () => { setTool(b.dataset.tool); svg.focus({ preventScroll: true }); }));
  function paintSwatches() {
    const box = $('swatches');
    const list = [...sel].map(byId).filter(Boolean);
    const forSticky = tool === 'sticky' || (list.length && list.every(i => i.type === 'sticky' || i.type === 'shape' || i.type === 'frame'));
    const colours = forSticky ? STICKY : INK;
    const current = forSticky ? (list[0] && (list[0].color || list[0].fill)) || stickyColor : (list[0] && (list[0].c || list[0].color)) || inkColor;
    box.textContent = '';
    colours.forEach(c => {
      const b = document.createElement('button');
      b.className = 'swatch' + (c === current ? ' on' : '');
      b.style.background = c;
      b.title = c;
      b.setAttribute('aria-label', 'Colour ' + c);
      b.addEventListener('click', () => setColour(c, forSticky));
      box.appendChild(b);
    });
  }
  function setColour(c, forSticky) {
    if (forSticky) stickyColor = c; else inkColor = c;
    const list = [...sel].map(byId).filter(Boolean);
    if (list.length) {
      snapshot();
      list.forEach(i => {
        if (i.type === 'sticky') i.color = c;
        else if (i.type === 'shape') i.fill = c;
        else if (i.type === 'frame') i.fill = c;
        else if (i.type === 'text') i.color = c;
        else if (i.type === 'ink') i.c = c;
        else if (i.type === 'connector') i.color = c;
      });
      changed();
    }
    paintSwatches();
  }

  // ---------- Arrange ----------
  $('toFront').addEventListener('click', () => { const s = new Set(sel); if (!s.size) return; snapshot(); board.items = [...board.items.filter(i => !s.has(i.id)), ...board.items.filter(i => s.has(i.id))]; changed(); });
  $('toBack').addEventListener('click', () => { const s = new Set(sel); if (!s.size) return; snapshot(); board.items = [...board.items.filter(i => s.has(i.id)), ...board.items.filter(i => !s.has(i.id))]; changed(); });
  $('dupBtn').addEventListener('click', duplicate);
  $('delBtn').addEventListener('click', removeSelected);
  document.querySelectorAll('[data-align]').forEach(b => b.addEventListener('click', () => {
    const list = [...sel].map(byId).filter(i => i && i.type !== 'connector');
    if (list.length < 2) return UI.toast('Select two or more items to line up');
    snapshot();
    const bb = bounds(list), a = b.dataset.align;
    if (a === 'hspace' || a === 'vspace') {
      const h = a === 'hspace';
      list.sort((p, q) => (h ? p.x - q.x : p.y - q.y));
      const total = list.reduce((s, i) => s + (h ? i.w : i.h), 0), gap = ((h ? bb.w : bb.h) - total) / (list.length - 1);
      let at = h ? bb.x : bb.y;
      list.forEach(i => { if (h) i.x = at; else i.y = at; at += (h ? i.w : i.h) + gap; });
    } else list.forEach(i => {
      if (a === 'left') i.x = bb.x; if (a === 'right') i.x = bb.x + bb.w - i.w; if (a === 'center') i.x = bb.x + bb.w / 2 - i.w / 2;
      if (a === 'top') i.y = bb.y; if (a === 'bottom') i.y = bb.y + bb.h - i.h; if (a === 'middle') i.y = bb.y + bb.h / 2 - i.h / 2;
    });
    changed();
  }));
  /** Tidy Up: a tree from the selected item (along its connectors) laid out to the right. */
  function tidy() {
    const root = byId([...sel][0]);
    if (!root || root.type === 'connector') return UI.toast('Select the first item of a mind map or flowchart');
    snapshot();
    const seen = new Set();
    const GX = 80, GY = 28;
    const height = n => { seen.add(n.id); const kids = childrenOf(n.id).filter(c => !seen.has(c.id)); n._kids = kids; return Math.max(n.h, kids.reduce((s, c) => s + height(c), 0) + GY * Math.max(0, kids.length - 1)); };
    const place = (n, x, top) => {
      const hh = height2.get(n.id);
      n.x = x; n.y = top + hh / 2 - n.h / 2;
      let y = top;
      const mw = Math.max(0, ...n._kids.map(k => k.w));
      for (const c of n._kids) { place(c, x + n.w + GX, y); y += height2.get(c.id) + GY; }
      void mw;
    };
    const height2 = new Map();
    const measure = n => { const kids = n._kids || []; kids.forEach(measure); height2.set(n.id, Math.max(n.h, kids.reduce((s, c) => s + height2.get(c.id), 0) + GY * Math.max(0, kids.length - 1))); };
    height(root);
    measure(root);
    const cy = root.y + root.h / 2;
    place(root, root.x, cy - height2.get(root.id) / 2);
    board.items.forEach(i => { delete i._kids; });
    changed();
  }
  $('tidyBtn').addEventListener('click', tidy);

  // ---------- Templates ----------
  function template(kind) {
    const [cx, cy] = center();
    const made = [];
    const put = o => { const it = cleanItem({ ...o, id: newId() }, true); made.push(it); return it; };
    const link = (a, b, extra = {}) => put({ type: 'connector', from: { id: a.id }, to: { id: b.id }, kind: 'elbow', arrow: 'end', ...extra });
    if (kind === 'flowchart') {
      const s = put({ type: 'shape', shape: 'round', x: cx - 80, y: cy - 300, w: 160, h: 64, text: 'Start', fill: '#e6f4ea', line: '#188038' });
      const p = put({ type: 'shape', shape: 'rect', x: cx - 90, y: cy - 180, w: 180, h: 80, text: 'Do something' });
      const d = put({ type: 'shape', shape: 'diamond', x: cx - 100, y: cy - 50, w: 200, h: 130, text: 'Did it work?', fill: '#fef7e0', line: '#f29900' });
      const y = put({ type: 'shape', shape: 'round', x: cx - 80, y: cy + 150, w: 160, h: 64, text: 'Done', fill: '#e6f4ea', line: '#188038' });
      const n = put({ type: 'shape', shape: 'rect', x: cx + 200, y: cy - 25, w: 180, h: 80, text: 'Fix it' });
      link(s, p); link(p, d); link(d, y); link(d, n); link(n, p);
    } else if (kind === 'mindmap') {
      const root = put({ type: 'shape', shape: 'round', x: cx - 110, y: cy - 40, w: 220, h: 80, text: 'Main idea', size: 20, fill: '#d2e3fc', line: '#1a73e8' });
      ['Idea', 'Idea', 'Idea'].forEach((t, n) => { const c = put({ type: 'sticky', x: cx + 230, y: cy - 260 + n * 200, w: 160, h: 160, color: STICKY[[0, 3, 4][n]], text: t }); link(root, c, { kind: 'curve', arrow: 'none' }); });
    } else if (kind === 'brainstorm') {
      const f = put({ type: 'frame', x: cx - 460, y: cy - 280, w: 920, h: 560, title: 'Brainstorm' });
      put({ type: 'text', x: f.x + 30, y: f.y + 24, w: 600, h: 40, text: 'What problem are we solving?', size: 26, bold: true });
      for (let n = 0; n < 8; n++) put({ type: 'sticky', x: f.x + 30 + (n % 4) * 220, y: f.y + 100 + Math.floor(n / 4) * 220, w: 190, h: 190, color: STICKY[n % 6], text: '' });
    } else if (kind === 'retro' || kind === 'kanban') {
      const cols = kind === 'retro' ? [['Went well', '#e6f4ea', STICKY[3]], ['To improve', '#fce8e6', STICKY[2]], ['Actions', '#e8f0fe', STICKY[4]]] : [['To do', '#f1f3f4', STICKY[0]], ['Doing', '#fef7e0', STICKY[1]], ['Done', '#e6f4ea', STICKY[3]]];
      cols.forEach(([t, fill, c], n) => {
        const f = put({ type: 'frame', x: cx - 600 + n * 410, y: cy - 300, w: 380, h: 600, title: t, fill });
        for (let k = 0; k < 2; k++) put({ type: 'sticky', x: f.x + 20 + k * 175, y: f.y + 20, w: 160, h: 160, color: c, text: '' });
      });
    }
    snapshot();
    // Frames go to the back, so their notes stay on top.
    board.items = [...made.filter(i => i.type === 'frame'), ...board.items, ...made.filter(i => i.type !== 'frame')];
    changed();
    select(made.map(i => i.id));
    const b = bounds(made), r = svg.getBoundingClientRect();
    if (b && (b.w * view.k > r.width || b.h * view.k > r.height)) { const k = Math.min((r.width - 80) / b.w, (r.height - 80) / b.h); zoomAt(k, ...screenCenter()); }
  }
  document.querySelectorAll('[data-template]').forEach(b => b.addEventListener('click', () => template(b.dataset.template)));

  // ---------- View ----------
  $('zoomIn').addEventListener('click', () => zoomAt(view.k * 1.25, ...screenCenter()));
  $('zoomOut').addEventListener('click', () => zoomAt(view.k / 1.25, ...screenCenter()));
  $('zoom100').addEventListener('click', () => zoomAt(1, ...screenCenter()));
  $('zoomFit').addEventListener('click', zoomFit);
  $('bgSel').addEventListener('change', e => { snapshot(); board.bg = e.target.value; changed(); });

  // ---------- Export ----------
  /** The board (or the selection) as a standalone SVG: the same drawing, nothing of the app's. */
  function exportSvg(onlySel) {
    const list = onlySel && sel.size ? [...sel].map(byId).filter(Boolean) : board.items;
    const b = bounds(list) || { x: 0, y: 0, w: 800, h: 600 };
    const frameTop = list.some(i => i.type === 'frame') ? 34 : 0;
    const m = 40, x = b.x - m, y = b.y - m - frameTop, w = b.w + 2 * m, h = b.h + 2 * m + frameTop;
    const out = document.createElementNS(NS, 'svg');
    out.setAttribute('xmlns', NS);
    out.setAttribute('viewBox', `${x} ${y} ${w} ${h}`);
    out.setAttribute('width', Math.round(w)); out.setAttribute('height', Math.round(h));
    out.appendChild(svg.querySelector('defs marker').parentNode.cloneNode(true));
    out.appendChild(el('rect', { x, y, width: w, height: h, fill: '#ffffff' }));
    const keep = new Set(list.map(i => i.id));
    for (const i of drawOrder()) if (keep.has(i.id)) { const g = drawItem(i); g.querySelectorAll('.ink-hit, .conn-hit').forEach(n => n.remove()); g.querySelectorAll('.sticky-shadow').forEach(n => n.setAttribute('fill', 'rgba(0,0,0,0.12)')); g.querySelectorAll('.frame-box').forEach(n => { n.setAttribute('stroke', '#dadce0'); n.setAttribute('stroke-width', '1.5'); }); out.appendChild(g); }
    // The arrow heads take the line's colour.
    out.querySelectorAll('path[marker-end], path[marker-start]').forEach(p => { const id = 'a' + p.getAttribute('stroke').slice(1); if (!out.querySelector('#' + id)) { const mk = out.querySelector('#arrowEnd').cloneNode(true); mk.id = id; mk.querySelector('path').setAttribute('fill', p.getAttribute('stroke')); out.querySelector('defs').appendChild(mk); } if (p.hasAttribute('marker-end')) p.setAttribute('marker-end', `url(#${id})`); if (p.hasAttribute('marker-start')) p.setAttribute('marker-start', `url(#${id})`); });
    out.querySelector('#arrowEnd').remove();
    out.querySelectorAll('pattern').forEach(n => n.remove());
    return { text: new XMLSerializer().serializeToString(out), w, h };
  }
  async function exportPng() {
    if (!board.items.length) return UI.toast('The board is empty');
    const s = exportSvg(true);
    const scale = Math.min(2, 8000 / Math.max(s.w, s.h));
    const img = new Image();
    const url = URL.createObjectURL(new Blob([s.text], { type: 'image/svg+xml' }));
    try {
      await new Promise((res, rej) => { img.onload = res; img.onerror = () => rej(new Error('the drawing could not be made into a picture')); img.src = url; });
      const c = Object.assign(document.createElement('canvas'), { width: Math.round(s.w * scale), height: Math.round(s.h * scale) });
      c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
      const blob = await new Promise(r => c.toBlob(r, 'image/png'));
      await Platform.save({ data: new Uint8Array(await blob.arrayBuffer()), name: Platform.baseName(state.name) + '.png', filters: [{ name: 'PNG picture', extensions: ['png'] }], saveAs: true, noVersion: true });
    } catch (err) { UI.toast('Export failed: ' + err.message, 5000); } finally { URL.revokeObjectURL(url); }
  }
  async function exportSvgFile() {
    if (!board.items.length) return UI.toast('The board is empty');
    const s = exportSvg(true);
    await Platform.save({ data: new TextEncoder().encode(s.text), name: Platform.baseName(state.name) + '.svg', filters: [{ name: 'SVG', extensions: ['svg'] }], saveAs: true, noVersion: true });
  }
  /** Printing and PDF: the whole board, fitted to the page. */
  function preparePrint() {
    let box = $('printBoard');
    if (!box) { box = document.createElement('div'); box.id = 'printBoard'; box.className = 'print-board'; document.body.appendChild(box); }
    box.innerHTML = board.items.length ? exportSvg(false).text : '';
  }
  addEventListener('beforeprint', preparePrint);

  // ---------- Files ----------
  const FILTERS = [{ name: 'Isadek Whiteboard', extensions: ['aboard'] }];
  const serialize = () => new TextEncoder().encode(JSON.stringify({ app: 'isadek-board', v: 1, bg: board.bg, items: board.items }));
  async function save(saveAs) {
    stopEditing(true);
    let name = state.name;
    if (!/\.aboard$/i.test(name)) name = Platform.baseName(name) + '.aboard';
    try {
      const r = await Platform.save({ data: serialize(), name, path: saveAs ? null : state.path, filters: FILTERS, saveAs: saveAs || !state.path });
      if (!r) return false;
      state.set({ name: r.name || name, path: r.path || state.path, dirty: false });
      UI.toast('Saved');
      return true;
    } catch (err) { UI.toast('Save failed: ' + err.message, 5000); return false; }
  }
  function loadFile(f) {
    let next;
    try { next = cleanBoard(JSON.parse(new TextDecoder().decode(f.data))); } catch (err) { UI.toast('Could not open the board: ' + err.message, 5000); return; }
    board = next;
    undoStack.length = redoStack.length = 0;
    sel.clear();
    state.set({ name: f.name, path: f.path || null, dirty: false });
    Platform.rememberRecent(f.name, 'board', f.path);
    render();
    zoomFit();
  }
  Platform.onMenu(async cmd => {
    switch (cmd) {
      case 'new': Platform.newWindow('board'); break;
      case 'open': {
        const f = await Platform.open(FILTERS);
        if (!f) break;
        if (Platform.shouldOpenInPlace(!board.items.length || !state.dirty)) { if (state.dirty && !(await UI.confirm('Open file', 'Discard the changes to this board?'))) break; loadFile(f); } else Platform.openPath(f.path);
        break;
      }
      case 'save': save(false); break;
      case 'save-as': save(true); break;
      case 'save-and-close': if (await save(false)) Platform.closeNow(); break;
      case 'version-history': Platform.versionHistory({ state, app: 'board', load: loadFile }); break;
      case 'open-cloud': Cloud.openDialog({ app: 'board', load: loadFile, isEmpty: async () => !board.items.length || !state.dirty }); break;
      case 'save-cloud': Cloud.saveAsDialog({ state, serialize: () => serialize(), exts: ['aboard'] }); break;
      case 'share': Cloud.shareDialog({ state, serialize: () => serialize(), exts: ['aboard'], save: () => save(false) }); break;
      case 'undo': if (!editing) undo(); break;
      case 'redo': if (!editing) redo(); break;
      case 'export-png': exportPng(); break;
      case 'export-svg': exportSvgFile(); break;
      case 'export-pdf': preparePrint(); Platform.exportPdf(Platform.baseName(state.name) + '.pdf', true); break;
      case 'print': preparePrint(); Platform.print(); break;
      case 'home': Platform.desktop ? Platform.newWindow('home') : (location.href = 'index.html'); break;
    }
  });
  Platform.preview = () => null;

  // ---------- Live editing: each item is one key; others' cursors show on the board ----------
  let peersNow = [];
  /** Other people's cursors (with their names) and what they have selected. */
  function drawPeers(peers) {
    peersNow = peers;
    peersG.textContent = '';
    peers.forEach(p => {
      const pr = p.presence || {};
      if (pr.cursor && isFinite(pr.cursor.x) && isFinite(pr.cursor.y)) {
        const g = el('g', { class: 'peer', transform: `translate(${pr.cursor.x} ${pr.cursor.y}) scale(${1 / view.k})` });
        g.append(el('path', { d: 'M0,0 L0,16 L4.5,12 L8,19 L10.5,18 L7,11 L13,11 Z', fill: p.color, stroke: '#fff', 'stroke-width': 1.2 }));
        const t = el('text', { x: 14, y: 26, fill: p.color, 'font-size': 12, 'font-family': FONT, 'font-weight': 600 }); t.textContent = p.name; g.appendChild(t);
        peersG.appendChild(g);
      }
      (Array.isArray(pr.sel) ? pr.sel : []).forEach(id => { const i = typeof id === 'string' && byId(id); if (!i) return; const b = boxOf(i); peersG.appendChild(el('rect', { class: 'peer peer-sel', x: b.x - 5, y: b.y - 5, width: b.w + 10, height: b.h + 10, stroke: p.color, 'stroke-width': 2 / view.k })); });
    });
  }
  const live = window.Collab ? Collab.attach({
    app: 'board', doc: state,
    busy: () => !!drag || !!editing,
    state() { const out = { board: { bg: board.bg } }; board.items.forEach((i, z) => { out['i:' + i.id] = { ...i, z }; }); return out; },
    apply(changes) {
      const order = new Map(board.items.map((i, z) => [i.id, z]));
      for (const [k, v] of Object.entries(changes)) {
        if (k === 'board' && v && typeof v === 'object') board.bg = oneOf(v.bg, ['dots', 'grid', 'plain'], board.bg);
        else if (k.startsWith('i:')) {
          const id = k.slice(2);
          if (!ID_OK.test(id)) continue;
          const at = board.items.findIndex(i => i.id === id);
          if (v === null) { if (at >= 0) board.items.splice(at, 1); sel.delete(id); continue; }
          const it = cleanItem(v, true);
          if (!it || it.id !== id) continue;
          if (at >= 0) board.items[at] = it; else board.items.push(it);
          order.set(id, num(v.z, board.items.length));
        }
      }
      board.items.sort((a, b) => (order.get(a.id) ?? 1e9) - (order.get(b.id) ?? 1e9));
      for (const id of [...sel]) if (!byId(id)) sel.delete(id);
      render();
    },
    presence: () => ({ cursor, sel: [...sel].slice(0, 50) }),
    showPeers: drawPeers,
    follow(presence) { if (presence && presence.cursor) { const r = svg.getBoundingClientRect(); view.x = r.width / 2 - presence.cursor.x * view.k; view.y = r.height / 2 - presence.cursor.y * view.k; applyView(); } }
  }) : null;

  // ---------- Start ----------
  setTool('select');
  render();
  { const r = svg.getBoundingClientRect(); view.x = r.width / 2; view.y = r.height / 2; applyView(); }
  addEventListener('resize', () => { if (editing) placeEditor(); });
  window.AtlasBoard = {
    get board() { return board; }, get view() { return view; }, get selection() { return [...sel]; }, get live() { return live; },
    loadFile, serialize, exportSvg, select, setTool, zoomFit, template, tidy, add: o => add(o), cleanBoard,
    toScreen: (x, y) => { const r = svg.getBoundingClientRect(); return [r.left + view.x + x * view.k, r.top + view.y + y * view.k]; }
  };
  (async () => {
    const f = await Platform.pending();
    if (f) loadFile(f);
  })();
})();
