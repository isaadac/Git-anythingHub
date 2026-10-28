// Real-world files (e2e/corpus, from Apache POI's test data): each opens without errors and shows
// what other programs read from it (python-docx, openpyxl, python-pptx), and the copy Isadek Office
// saves opens in those programs with the same content.
const { chromium } = require('playwright');
const { execFileSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };

const CORPUS = path.join(__dirname, '../corpus');
const OUT = path.join(process.env.E2E_WORK, 'corpus-saved');
fs.mkdirSync(OUT, { recursive: true });
const APP = { docx: 'write', xlsx: 'sheets', pptx: 'slides' };
const only = (process.env.CORPUS_ONLY || '').split(',').filter(Boolean);
const files = fs.readdirSync(CORPUS).filter(f => APP[path.extname(f).slice(1)] && (!only.length || only.some(o => f.includes(o)))).sort();
// How much of what other programs read must Isadek Office show and keep.
const TEXT = 0.97, CELLS = 0.98;
// Excel's own results that a formula engine is not expected to match exactly (documented in the
// files themselves as edge cases): compared only for what the file states as values.
const FORMULA_EDGE = new Set(['FormulaEvalTestData_Copy.xlsx']);

function reference(paths) {
  const out = execFileSync('python3', ['-I', path.join(__dirname, '../corpus-extract.py'), ...paths], { encoding: 'utf8', maxBuffer: 256e6 });
  const map = new Map();
  for (const line of out.split('\n').filter(Boolean)) { const r = JSON.parse(line); map.set(path.basename(r.file), r); }
  return map;
}
const norm = w => w.normalize('NFC').toLowerCase().replace(/^[^\p{L}\p{N}]+|[^\p{L}\p{N}]+$/gu, '');
/** Share of the reference's words (counted with repeats) that are also in `got`. */
function recall(want, got) {
  const bag = new Map();
  for (const w of got.map(norm).filter(Boolean)) bag.set(w, (bag.get(w) || 0) + 1);
  const ws = want.map(norm).filter(Boolean);
  if (!ws.length) return { r: 1, missing: [] };
  const missing = [];
  for (const w of ws) { const n = bag.get(w) || 0; if (n) bag.set(w, n - 1); else missing.push(w); }
  return { r: 1 - missing.length / ws.length, missing };
}
const serial = iso => { const m = iso.match(/^(\d{4})-(\d\d)-(\d\d)(?:T(\d\d):(\d\d):(\d\d)(?:\.(\d+))?)?$/) || []; if (!m[1]) { const t = iso.match(/^(\d\d):(\d\d):(\d\d)/); return t ? (+t[1] * 3600 + +t[2] * 60 + +t[3]) / 86400 : NaN; } return (Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0)) - Date.UTC(1899, 11, 30)) / 864e5; };
function same(want, got) {
  // Error values: Isadek Office keeps them as {error}, other programs as their text.
  if (got && typeof got === 'object' && got.error) got = got.error;
  if (got && typeof got === 'object' && got.date) got = serial(got.date);
  if (want && typeof want === 'object' && want.error) want = want.error;
  if (want && typeof want === 'object' && want.date) return typeof got === 'number' ? Math.abs(got - serial(want.date)) < 1e-6 : String(got).startsWith(want.date.slice(0, 10));
  if (typeof want === 'number') { const g = typeof got === 'string' && got.trim() !== '' ? Number(got) : got; return typeof g === 'number' && Math.abs(g - want) <= 1e-9 * Math.max(1, Math.abs(want)); }
  if (typeof want === 'boolean') return got === want || String(got).toUpperCase() === String(want).toUpperCase();
  return String(got == null ? '' : got).trim() === String(want).trim();
}
function cellAgreement(want, got) {
  let n = 0, ok = 0; const bad = [];
  for (const [name, cells] of Object.entries(want.sheets || {})) {
    const g = got[name];
    for (const [ref, v] of Object.entries(cells)) {
      n++;
      if (g && same(v, g[ref])) ok++; else if (bad.length < 6) bad.push(`${name}!${ref}: ${JSON.stringify(v)} ≠ ${JSON.stringify(g ? g[ref] : '(no sheet)')}`);
    }
  }
  return { r: n ? ok / n : 1, n, bad };
}

(async () => {
  const ref = reference(files.map(f => path.join(CORPUS, f)));
  const browser = await chromium.launch();
  const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 } });
  const pages = {};
  const errors = {};
  async function page(app) {
    if (pages[app]) return pages[app];
    const p = await ctx.newPage();
    p.on('pageerror', e => { (errors[p.__file] = errors[p.__file] || []).push(e.message); });
    await p.goto(process.env.E2E_BASE + `/src/renderer/${app}.html`);
    await p.waitForFunction(a => window[{ write: 'AtlasWrite', sheets: 'AtlasSheets', slides: 'AtlasSlides' }[a]], app);
    p.on('dialog', d => d.dismiss());
    return (pages[app] = p);
  }
  const summary = [];
  const saved = [];
  for (const f of files) {
    const ext = path.extname(f).slice(1), app = APP[ext];
    const p = await page(app);
    p.__file = f;
    const bytes = fs.readFileSync(path.join(CORPUS, f));
    const t0 = Date.now();
    let got, out;
    try {
      got = await p.evaluate(async ([app, name, data]) => {
        const words = t => (t || '').split(/\s+/).filter(Boolean);
        const strip = h => { const d = document.createElement('div'); d.innerHTML = h || ''; d.querySelectorAll('br,hr,p,div,li,ul,ol,td,th,tr,table,h1,h2,h3,h4,h5,h6,blockquote,pre,section').forEach(n => { n.before(' '); n.after(' '); if (n.firstChild) n.prepend(' '); }); return d.textContent; };
        if (app === 'write') {
          await AtlasWrite.loadFile({ name, data: new Uint8Array(data) });
          // Footnote and endnote text is in the notes area, as in Word's body text it is not.
          const doc = document.getElementById('doc').cloneNode(true);
          doc.querySelectorAll('.endnotes-gen, .footnotes-gen, sup.note').forEach(n => n.remove());
          return { words: words(strip(doc.innerHTML)) };
        }
        if (app === 'sheets') {
          await AtlasSheets.loadFile({ name, data: new Uint8Array(data) });
          const book = JSON.parse(AtlasSheets.bookJson());
          const sheets = {};
          book.sheets.forEach((s, i) => {
            AtlasSheets.setSheet(i);
            const cells = {};
            const refs = Object.keys(s.cells || {});
            for (const ref of refs) { const v = AtlasSheets.valueOf(ref); if (v !== '' && v != null) cells[ref] = v; }
            sheets[s.name] = cells;
          });
          AtlasSheets.setSheet(0);
          return { sheets };
        }
        await AtlasSlides.loadFile({ name, data: new Uint8Array(data) });
        const texts = [];
        const walk = els => (els || []).forEach(e => { if (e.html) texts.push(strip(e.html)); walk(e.children); });
        AtlasSlides.pres.slides.forEach(s => walk(s.elements));
        return { slides: AtlasSlides.pres.slides.length, words: words(texts.join(' ')) };
      }, [app, f, [...bytes]]);
      out = Buffer.from(await p.evaluate(async ([app, ext]) => Array.from(await window[{ write: 'AtlasWrite', sheets: 'AtlasSheets', slides: 'AtlasSlides' }[app]].serialize(ext)), [app, ext]));
    } catch (e) {
      check(`${f}: opens and saves`, false, String(e.message || e).split('\n')[0]);
      summary.push({ f, open: 'error' });
      continue;
    }
    const secs = (Date.now() - t0) / 1000;
    fs.writeFileSync(path.join(OUT, f), out);
    saved.push(f);
    const errs = errors[f] || [];
    check(`${f}: opens and saves without errors (${secs.toFixed(1)} s)`, !errs.length && secs < 30, errs.slice(0, 3).join(' | '));
    const r = ref.get(f);
    const row = { f, secs };
    if (r.error) { console.log(`INFO ${f}: the reference reader cannot read the original (${r.error.slice(0, 80)}); content not compared`); summary.push(row); continue; }
    if (app === 'sheets') {
      const c = cellAgreement(r, got.sheets);
      row.open = c.r;
      const limit = FORMULA_EDGE.has(f) ? 0.9 : CELLS;
      check(`${f}: cell values as other programs read them (${(c.r * 100).toFixed(1)}% of ${c.n})`, c.r >= limit, c.bad.join('; '));
      row.got = got;
    } else {
      const t = recall(r.words, got.words);
      row.open = t.r;
      check(`${f}: text as other programs read it (${(t.r * 100).toFixed(1)}% of ${r.words.length} words)`, t.r >= TEXT, t.missing.slice(0, 12).join(' '));
      if (app === 'slides') check(`${f}: all ${r.slides} slides`, got.slides === r.slides, got.slides);
    }
    summary.push(row);
  }

  // The saved copies, read by the same independent programs.
  const back = reference(saved.map(f => path.join(OUT, f)));
  for (const row of summary) {
    if (!saved.includes(row.f)) continue;
    const r = ref.get(row.f), b = back.get(row.f);
    check(`${row.f}: the saved copy opens in other programs`, b && !b.error, b && b.error);
    if (!b || b.error || r.error) continue;
    if (b.sheets) {
      // What Isadek Office showed must be what it saved (its own results for formulas).
      const want = { sheets: Object.fromEntries(Object.entries(row.got.sheets).map(([n, cells]) => [n, cells])) };
      const c = cellAgreement(want, Object.fromEntries(Object.entries(b.sheets).map(([n, cells]) => [n, Object.fromEntries(Object.entries(cells).map(([k, v]) => [k, v && v.date ? serial(v.date) : v]))])));
      const orig = cellAgreement(r, b.sheets);
      row.saved = orig.r;
      check(`${row.f}: the saved copy keeps the values (${(c.r * 100).toFixed(1)}% of what was shown, ${(orig.r * 100).toFixed(1)}% of the original)`, c.r >= CELLS && orig.r >= (FORMULA_EDGE.has(row.f) ? 0.9 : CELLS), c.bad.concat(orig.bad).slice(0, 6).join('; '));
    } else {
      const t = recall(r.words, b.words);
      row.saved = t.r;
      check(`${row.f}: the saved copy keeps the text (${(t.r * 100).toFixed(1)}%)`, t.r >= TEXT, t.missing.slice(0, 12).join(' '));
      if (b.slides !== undefined) check(`${row.f}: the saved copy has all ${r.slides} slides`, b.slides === r.slides, b.slides);
    }
  }
  const pct = x => (typeof x === 'number' ? (x * 100).toFixed(1) + '%' : x || '—');
  const table = ['| File | Open | Saved | s |', '|---|---|---|---|', ...summary.map(r => `| ${r.f} | ${pct(r.open)} | ${pct(r.saved)} | ${r.secs ? r.secs.toFixed(1) : ''} |`)].join('\n');
  fs.writeFileSync(path.join(process.env.E2E_WORK, 'shots', 'corpus.md'), table + '\n');
  if (process.env.CORPUS_REPORT) fs.writeFileSync(process.env.CORPUS_REPORT, table + '\n');
  console.log(`INFO ${files.length} files`);
  await browser.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
