'use strict';
// Isadek Sheets — Analyze Data and Forecast Sheet.
//
// Analyze Data looks at the table around the selection and says what stands out: totals and
// averages, trends over time, the biggest categories, unusual values and columns that move
// together, each with a chart or PivotTable to insert. Forecast Sheet continues a time series
// with FORECAST.ETS (as formulas, so it stays live) and draws it with its confidence band.
(function () {
  const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const mean = a => a.reduce((s, x) => s + x, 0) / a.length;
  const sd = a => { const m = mean(a); return Math.sqrt(a.reduce((s, x) => s + (x - m) ** 2, 0) / Math.max(1, a.length - 1)); };
  const corr = (a, b) => { const ma = mean(a), mb = mean(b); let sab = 0, saa = 0, sbb = 0; a.forEach((x, i) => { sab += (x - ma) * (b[i] - mb); saa += (x - ma) ** 2; sbb += (b[i] - mb) ** 2; }); return saa && sbb ? sab / Math.sqrt(saa * sbb) : 0; };
  const fmt = v => (Math.abs(v) >= 1000 ? Math.round(v).toLocaleString() : (+v.toPrecision(4)).toLocaleString());

  /**
   * What stands out in a table: [{ kind, text, action: { label, run: 'chart'|'pivot'|'select', … }, score }].
   * table: { headers, rows (values), colTypes } — pure, so it can be tested on its own.
   */
  function analyze(t) {
    const out = [];
    const cols = t.headers.map((h, c) => {
      const vals = t.rows.map(r => r[c]);
      const nums = vals.filter(v => typeof v === 'number' && isFinite(v));
      const texts = vals.filter(v => typeof v === 'string' && v.trim());
      return { h, c, vals, nums, texts, numeric: nums.length >= Math.max(3, vals.length * 0.8), date: !!(t.dateCols && t.dateCols[c]), distinct: new Set(texts.map(x => x.trim().toLowerCase())).size };
    });
    const numeric = cols.filter(c => c.numeric && !c.date);
    const cats = cols.filter(c => !c.numeric && c.texts.length >= t.rows.length * 0.8 && c.distinct >= 2 && c.distinct <= Math.min(30, t.rows.length * 0.7));
    const timeCol = cols.find(c => c.date) || null;
    // Totals and averages.
    for (const n of numeric.slice(0, 6)) out.push({ kind: 'summary', score: 1, text: `${n.h}: total ${fmt(n.nums.reduce((s, x) => s + x, 0))}, average ${fmt(mean(n.nums))} (from ${fmt(Math.min(...n.nums))} to ${fmt(Math.max(...n.nums))})` });
    // Trends: over the timeline if there is one, otherwise down the rows.
    for (const n of numeric.slice(0, 6)) {
      const pts = n.vals.map((v, i) => [timeCol ? timeCol.vals[i] : i + 1, v]).filter(([x, y]) => typeof x === 'number' && typeof y === 'number');
      if (pts.length < 5) continue;
      const fit = window.Formula.trendFit(pts.map(p => p[0]), pts.map(p => p[1]));
      if (!fit || fit.r2 < 0.6) continue;
      const first = fit.f(pts[0][0]), last = fit.f(pts[pts.length - 1][0]);
      const change = first ? (last - first) / Math.abs(first) : 0;
      if (Math.abs(change) < 0.05) continue;
      out.push({ kind: 'trend', score: 3 + fit.r2, text: `${n.h} ${change > 0 ? 'rises' : 'falls'} steadily${timeCol ? ` over ${timeCol.h}` : ''}: about ${change > 0 ? '+' : ''}${Math.round(change * 100)}% from first to last (R² ${fit.r2.toFixed(2)})`, action: { label: 'Insert Line Chart', run: 'chart', type: 'line', cols: [timeCol ? timeCol.c : null, n.c].filter(c => c !== null) } });
    }
    // The biggest category, for each category column and the first numbers.
    for (const cat of cats.slice(0, 3)) for (const n of numeric.slice(0, 2)) {
      const sums = new Map();
      t.rows.forEach(r => { const k = typeof r[cat.c] === 'string' ? r[cat.c].trim() : ''; const v = r[n.c]; if (k && typeof v === 'number') sums.set(k, (sums.get(k) || 0) + v); });
      if (sums.size < 2) continue;
      const total = [...sums.values()].reduce((s, x) => s + x, 0);
      const [topK, topV] = [...sums.entries()].sort((a, b) => b[1] - a[1])[0];
      const share = total ? topV / total : 0;
      out.push({ kind: 'category', score: 2 + share, text: `${topK} has the most ${n.h} by ${cat.h}: ${fmt(topV)} (${Math.round(share * 100)}% of ${fmt(total)})`, action: { label: 'Insert PivotTable', run: 'pivot', rows: [cat.h], vals: [{ field: n.h, agg: 'sum' }] } });
    }
    // Unusual values (more than 3 standard deviations, or far outside the middle half).
    for (const n of numeric.slice(0, 6)) {
      if (n.nums.length < 8) continue;
      const s = [...n.nums].sort((a, b) => a - b), q1 = s[Math.floor(s.length * 0.25)], q3 = s[Math.floor(s.length * 0.75)], iqr = q3 - q1;
      const m = mean(n.nums), d = sd(n.nums);
      n.vals.forEach((v, i) => {
        if (typeof v !== 'number' || !d) return;
        if (Math.abs(v - m) / d > 3 || (iqr && (v > q3 + 3 * iqr || v < q1 - 3 * iqr))) out.push({ kind: 'outlier', score: 4, text: `Row ${i + 2}: ${n.h} is ${fmt(v)}, unusually ${v > m ? 'high' : 'low'} (most are ${fmt(q1)}–${fmt(q3)})`, action: { label: 'Go to It', run: 'select', row: i, col: n.c } });
      });
    }
    // Columns that move together.
    for (let i = 0; i < numeric.length && i < 6; i++) for (let j = i + 1; j < numeric.length && j < 6; j++) {
      const a = [], b = [];
      numeric[i].vals.forEach((v, k) => { const w = numeric[j].vals[k]; if (typeof v === 'number' && typeof w === 'number') { a.push(v); b.push(w); } });
      if (a.length < 5) continue;
      const r = corr(a, b);
      if (Math.abs(r) >= 0.7) out.push({ kind: 'correlation', score: 2.5 + Math.abs(r), text: `${numeric[j].h} ${r > 0 ? 'rises' : 'falls'} as ${numeric[i].h} rises (correlation ${r.toFixed(2)})`, action: { label: 'Insert Scatter Chart', run: 'chart', type: 'scatter', cols: [numeric[i].c, numeric[j].c] } });
    }
    // The most frequent value of a category column.
    for (const cat of cats.slice(0, 2)) {
      const counts = new Map();
      cat.texts.forEach(x => counts.set(x.trim(), (counts.get(x.trim()) || 0) + 1));
      const [k, v] = [...counts.entries()].sort((a, b) => b[1] - a[1])[0];
      if (v > 1) out.push({ kind: 'frequency', score: 1.5, text: `${k} is the most common ${cat.h} (${v} of ${cat.texts.length} rows)`, action: { label: 'Insert PivotTable', run: 'pivot', rows: [cat.h], vals: [{ field: cat.h, agg: 'count' }] } });
    }
    return out.sort((a, b) => b.score - a.score).slice(0, 24);
  }

  function SheetsInsights(ctx) {
    const F = window.Formula;
    let pane = null;

    /** The table around the selection: a header row and the rows under it. */
    function region() {
      const g = ctx.dataRegion();
      if (g.r2 <= g.r1) return null;
      const s = ctx.sheet();
      const headers = [], rows = [], dateCols = {};
      for (let c = g.c1; c <= g.c2; c++) { const h = ctx.value(s, g.r1, c); headers.push(h === '' ? `Column ${F.indexToCol(c)}` : String(h)); }
      for (let r = g.r1 + 1; r <= g.r2; r++) { const row = []; for (let c = g.c1; c <= g.c2; c++) row.push(ctx.value(s, r, c)); rows.push(row); }
      for (let c = g.c1; c <= g.c2; c++) { const fmt = ((s.cells[F.addr(g.r1 + 1, c)] || {}).s || {}).fmt || ''; if (/[dy]/i.test(fmt) && !/\[/.test(fmt)) dateCols[c - g.c1] = true; }
      return { g, headers, rows, dateCols };
    }
    function analyzePane(show = !pane || pane.hidden) {
      if (!pane) {
        pane = document.createElement('aside');
        pane.className = 'script-pane insights-pane';
        pane.setAttribute('aria-label', 'Analyze Data');
        document.body.appendChild(pane);
        pane.addEventListener('keydown', e => e.stopPropagation());
        pane.addEventListener('click', e => {
          const b = e.target.closest('button');
          if (!b) return;
          if (b.dataset.a === 'close') return analyzePane(false);
          if (b.dataset.a === 'refresh') return draw();
          const i = +b.dataset.i;
          if (Number.isInteger(i) && current[i]) act(current[i].action);
        });
      }
      pane.hidden = !show;
      document.body.classList.toggle('script-open', show);
      if (show) draw();
    }
    let current = [], table = null;
    function draw() {
      table = region();
      current = table ? analyze(table) : [];
      pane.innerHTML = `<header><b>Analyze Data</b><span class="sp"></span><button data-a="refresh" title="Look again"><span class="ico" data-icon="refresh-cw"></span></button><button class="x" data-a="close" title="Close"><span class="ico" data-icon="x"></span></button></header>
        <div class="ins-list">${!table ? '<p class="hint">Select a cell in a table with a header row.</p>' : !current.length ? '<p class="hint">Nothing stands out in this table yet.</p>' :
        `<p class="hint">${esc(table.headers.length)} columns, ${esc(table.rows.length)} rows (${esc(F.addr(table.g.r1, table.g.c1))}:${esc(F.addr(table.g.r2, table.g.c2))})</p>` +
        current.map((x, i) => `<div class="ins-card ins-${x.kind}"><p>${esc(x.text)}</p>${x.action ? `<button data-i="${i}">${esc(x.action.label)}</button>` : ''}</div>`).join('')}</div>`;
    }
    function act(a) {
      const g = table.g;
      if (a.run === 'select') return ctx.selectRange({ r1: g.r1 + 1 + a.row, c1: g.c1 + a.col, r2: g.r1 + 1 + a.row, c2: g.c1 + a.col });
      if (a.run === 'pivot') return ctx.createPivot(ctx.sheet().name, `${F.addr(g.r1, g.c1)}:${F.addr(g.r2, g.c2)}`, a.rows, a.vals);
      if (a.run === 'chart') {
        // The columns side by side: a chart needs one block, so non-adjacent columns are copied next to the table.
        const cs = a.cols.map(c => c + g.c1);
        if (cs.length === 2 && cs[1] !== cs[0] + 1) {
          const at = ctx.freeColumnRight(g);
          ctx.copyColumns(g, cs, at);
          return ctx.insertChart(a.type, { r1: g.r1, c1: at, r2: g.r2, c2: at + 1 });
        }
        return ctx.insertChart(a.type, { r1: g.r1, c1: Math.min(...cs), r2: g.r2, c2: Math.max(...cs) });
      }
    }

    // ---------------- Forecast Sheet ----------------
    async function forecastSheet() {
      const g0 = ctx.selRange();
      let g = g0.c2 - g0.c1 === 1 && g0.r2 - g0.r1 >= 3 ? g0 : ctx.dataRegion();
      if (g.c2 - g.c1 < 1) return UI.toast('Select a column of dates (or periods) and a column of values beside it');
      g = { ...g, c2: g.c1 + 1 };
      const s = ctx.sheet();
      const hasHeader = typeof ctx.value(s, g.r1, g.c2) === 'string';
      const r1 = g.r1 + (hasHeader ? 1 : 0);
      const ts = [], ys = [];
      for (let r = r1; r <= g.r2; r++) { const t = ctx.value(s, r, g.c1), y = ctx.value(s, r, g.c2); if (typeof t === 'number' && typeof y === 'number') { ts.push(t); ys.push(y); } }
      if (ts.length < 4) return UI.toast('A forecast needs at least 4 dates with values');
      let step = Infinity;
      for (let i = 1; i < ts.length; i++) if (ts[i] > ts[i - 1]) step = Math.min(step, ts[i] - ts[i - 1]);
      if (!isFinite(step)) return UI.toast('The dates must increase');
      const v = await UI.form('Create Forecast Worksheet', [
        { name: 'n', label: 'Forecast end: how many periods ahead', type: 'number', value: Math.max(3, Math.round(ts.length / 3)), min: 1, max: 1000 },
        { name: 'conf', label: 'Confidence interval (%)', type: 'number', value: 95, min: 1, max: 99 },
        { name: 'season', label: 'Seasonality', type: 'select', options: [{ value: '1', label: 'Detect automatically' }, { value: '0', label: 'None' }, ...[4, 7, 12, 24, 52].map(n => ({ value: String(n), label: `${n} periods` }))], value: '1' }
      ], 'Create');
      if (!v) return;
      const n = Math.max(1, Math.min(1000, Math.round(+v.n || 1))), conf = Math.max(0.01, Math.min(0.99, (+v.conf || 95) / 100)), season = +v.season || (v.season === '0' ? 0 : 1);
      const fmtT = ((s.cells[F.addr(r1, g.c1)] || {}).s || {}).fmt || '';
      const name = ctx.uniqueName('Forecast');
      const tName = hasHeader ? String(ctx.value(s, g.r1, g.c1)) : 'Timeline', yName = hasHeader ? String(ctx.value(s, g.r1, g.c2)) : 'Values';
      const head = [tName, yName, `Forecast(${yName})`, 'Lower Confidence Bound', 'Upper Confidence Bound'];
      const N = ts.length, last = N + 1; // the data rows are 2..N+1
      const cells = [];
      head.forEach((h, c) => cells.push([0, c, h, { b: true }]));
      ts.forEach((t, i) => { cells.push([i + 1, 0, String(t), fmtT ? { fmt: fmtT } : null]); cells.push([i + 1, 1, String(ys[i]), null]); });
      // The last actual value starts the forecast line, so the two join.
      cells.push([N, 2, `=B${last}`, null], [N, 3, `=B${last}`, null], [N, 4, `=B${last}`, null]);
      const yr = `$B$2:$B$${last}`, tr = `$A$2:$A$${last}`;
      for (let k = 1; k <= n; k++) {
        const r = N + k, row = r + 1;
        cells.push([r, 0, `=A${row - 1}+${step}`, fmtT ? { fmt: fmtT } : null]);
        cells.push([r, 2, `=FORECAST.ETS(A${row},${yr},${tr},${season},1)`, { fmt: '#,##0.00' }]);
        cells.push([r, 3, `=C${row}-FORECAST.ETS.CONFINT(A${row},${yr},${tr},${conf},${season},1)`, { fmt: '#,##0.00' }]);
        cells.push([r, 4, `=C${row}+FORECAST.ETS.CONFINT(A${row},${yr},${tr},${conf},${season},1)`, { fmt: '#,##0.00' }]);
      }
      ctx.createSheetWith(name, cells, { colWidths: { 0: 110, 1: 100, 2: 150, 3: 170, 4: 170 }, chart: { type: 'line', range: `A1:E${N + n + 1}`, title: `Forecast: ${yName}`, catCol: true } });
      UI.toast('Forecast sheet created');
    }

    return { analyzePane, analyze: () => (region() ? analyze(region()) : []), forecastSheet };
  }
  SheetsInsights.analyze = analyze;
  window.SheetsInsights = SheetsInsights;
})();
