'use strict';
// Isadek — reading text in pictures and handwriting, on this device (Tesseract, in a worker).
//
// The engine and its English model come with the app, so nothing is downloaded and no picture
// leaves the computer. Ocr.recognize(image, { line }) → { text, confidence (0–100), lines: [...] }.
// image: a canvas, an <img>, a Blob or a data URL. line: true for a single line of text.
// Ocr.inkLines(strokes) groups ink into lines of writing and Ocr.inkCanvas(strokes) draws them
// cleanly (dark on white, evenly thick) for reading.
//
// library files: node_modules/tesseract.js/dist/tesseract.min.js node_modules/tesseract.js/dist/worker.min.js
// library files: node_modules/tesseract.js-core/tesseract-core-relaxedsimd-lstm.wasm.js node_modules/tesseract.js-core/tesseract-core-simd-lstm.wasm.js node_modules/tesseract.js-core/tesseract-core-lstm.wasm.js
// library folders: node_modules/@tesseract.js-data/eng/4.0.0_best_int/
(function () {
  const LIB = '../../node_modules/';
  const abs = p => new URL(LIB + p, location.href).href;
  let worker = null, loading = null, queue = Promise.resolve();

  function loadScript(src) {
    return new Promise((resolve, reject) => {
      if (window.Tesseract) return resolve();
      const s = document.createElement('script');
      s.src = src;
      s.onload = resolve;
      s.onerror = () => reject(new Error('The text reader could not be loaded'));
      document.head.appendChild(s);
    });
  }
  function engine() {
    if (loading) return loading;
    loading = (async () => {
      await loadScript(LIB + 'tesseract.js/dist/tesseract.min.js');
      // OEM 1: the neural-network reader only (its core is the smaller one that comes with the app).
      worker = await window.Tesseract.createWorker('eng', 1, {
        workerPath: abs('tesseract.js/dist/worker.min.js'),
        corePath: abs('tesseract.js-core/'),
        langPath: abs('@tesseract.js-data/eng/4.0.0_best_int'),
        workerBlobURL: false,
        cacheMethod: 'none',
        errorHandler: () => {}
      });
      return worker;
    })();
    loading.catch(() => { loading = null; worker = null; });
    return loading;
  }

  /** Reads the text in an image. Calls run one at a time. */
  function recognize(image, opts = {}) {
    const run = queue.then(async () => {
      const w = await engine();
      await w.setParameters({ tessedit_pageseg_mode: opts.line ? '7' : '3', preserve_interword_spaces: '1' });
      const r = await w.recognize(image);
      const d = r.data || {};
      const text = String(d.text || '').replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
      return { text, confidence: d.confidence || 0, lines: text ? text.split('\n') : [] };
    });
    queue = run.catch(() => {});
    return run;
  }

  // ---------------- Ink ----------------
  /** Strokes grouped into lines of writing (top to bottom), each line's strokes left to right. */
  function inkLines(strokes) {
    const items = strokes.filter(s => s.o >= 0.9).map(s => ({ s, b: Ink.bounds([s]) })).filter(it => it.b.w + it.b.h > 0);
    if (!items.length) return [];
    // A typical letter's height: dots and dashes are small, so the median of the taller half.
    const hs = items.map(it => it.b.h).sort((a, b) => a - b);
    const letter = Math.max(8, hs[Math.floor(hs.length * 0.75)] || 20);
    items.sort((a, b) => (a.b.y + a.b.h / 2) - (b.b.y + b.b.h / 2));
    const lines = [];
    for (const it of items) {
      const mid = it.b.y + it.b.h / 2;
      // Joins the line whose band its middle falls in (a dot over an i joins the line below it).
      let best = null, gap = Infinity;
      for (const l of lines) {
        const d = mid < l.y1 ? l.y1 - mid : mid > l.y2 ? mid - l.y2 : 0;
        if (d < gap) { gap = d; best = l; }
      }
      if (best && gap <= letter * 0.6) { best.items.push(it); best.y1 = Math.min(best.y1, it.b.y + (it.b.h > letter * 0.4 ? 0 : it.b.h)); best.y2 = Math.max(best.y2, it.b.y + it.b.h); }
      else lines.push({ items: [it], y1: it.b.y, y2: it.b.y + it.b.h });
    }
    return lines.sort((a, b) => a.y1 - b.y1).map(l => l.items.sort((a, b) => a.b.x - b.b.x).map(it => it.s));
  }
  /**
   * Strokes drawn for reading: dark on white with a margin, scaled so the writing is about
   * linePx pixels from its top to its bottom (a line with ascenders and descenders), each stroke
   * as thick as it was drawn (scaled), at least 2.5 pixels.
   */
  function inkCanvas(strokes, linePx = 64) {
    const b = Ink.bounds(strokes);
    const k = Math.max(0.2, Math.min(8, linePx / Math.max(4, b.h))), pad = 24;
    const c = document.createElement('canvas');
    c.width = Math.max(1, Math.min(8000, Math.ceil(b.w * k + pad * 2)));
    c.height = Math.max(1, Math.min(4000, Math.ceil(b.h * k + pad * 2)));
    const g = c.getContext('2d');
    g.fillStyle = '#fff';
    g.fillRect(0, 0, c.width, c.height);
    g.translate(pad, pad);
    g.scale(k, k);
    g.translate(-b.x, -b.y);
    g.strokeStyle = g.fillStyle = '#000';
    g.lineCap = g.lineJoin = 'round';
    for (const s of strokes) {
      const p = new Path2D(s.d);
      if (s.f) g.fill(p); else { g.lineWidth = Math.max(2.5 / k, s.w); g.stroke(p); }
    }
    return c;
  }
  /** Reads handwriting: each line of the strokes on its own. → { text, confidence, lines } */
  async function readInk(strokes) {
    const lines = inkLines(strokes), out = [];
    let conf = 0;
    for (const l of lines) {
      const r = await recognize(inkCanvas(l), { line: true });
      if (r.text) { out.push(r.text.replace(/\n+/g, ' ')); conf += r.confidence; }
    }
    return { text: out.join('\n'), confidence: out.length ? conf / out.length : 0, lines: out };
  }

  window.Ocr = { recognize, inkLines, inkCanvas, readInk, ready: () => !!worker };
})();
