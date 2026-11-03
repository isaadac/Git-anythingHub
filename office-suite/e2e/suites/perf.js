// Performance on big documents, and the correctness of what makes it fast: a 300,000-row workbook
// (load, scroll, recalculate) and a 250-page document (typing, pagination). Limits are generous so
// slower machines pass; they catch the kind of slowdown that makes big files unusable.
const { chromium } = require('playwright');
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const BASE = process.env.E2E_BASE + '/src/renderer/';
const N = 300000;

(async () => {
  const browser = await chromium.launch();
  const errors = [];
  const p = await browser.newPage({ viewport: { width: 1400, height: 900 } });
  p.on('pageerror', e => errors.push(e.message));

  // ---- Sheets
  await p.goto(BASE + 'sheets.html');
  await p.waitForFunction(() => window.AtlasSheets);
  const s1 = await p.evaluate(async N => {
    const lines = ['Id,Value,Group'];
    for (let i = 1; i <= N; i++) lines.push(i + ',' + (i % 97) + ',G' + (i % 13));
    const data = new TextEncoder().encode(lines.join('\n'));
    let t = performance.now();
    await AtlasSheets.loadFile({ name: 'big.csv', path: null, data });
    const load = performance.now() - t;
    AtlasSheets.setCell('E1', `=SUM(B2:B${N + 1})`);
    const first = AtlasSheets.valueOf('E1');
    t = performance.now();
    AtlasSheets.setCell('B2', '1000');
    const edited = AtlasSheets.valueOf('E1');
    const edit = performance.now() - t;
    const wrap = document.getElementById('gridWrap'), frames = [];
    for (let i = 0; i < 90; i++) {
      const s = performance.now();
      wrap.scrollTop += 120;
      wrap.dispatchEvent(new Event('scroll'));
      await new Promise(r => requestAnimationFrame(() => r()));
      frames.push(performance.now() - s);
    }
    frames.sort((a, b) => a - b);
    let want = 0; for (let i = 1; i <= N; i++) want += i % 97;
    return { load, first, want, edited, edit, p50: frames[45], p90: frames[81] };
  }, N);
  check(`a ${N.toLocaleString()}-row CSV opens quickly (${Math.round(s1.load)} ms)`, s1.load < 15000);
  check('a SUM over the whole column is right', s1.first === s1.want, `${s1.first} vs ${s1.want}`);
  check(`changing one value updates it quickly (${Math.round(s1.edit)} ms)`, s1.edited === s1.want - 1 + 1000 && s1.edit < 1500, s1.edited);
  check(`scrolling stays smooth (median frame ${Math.round(s1.p50)} ms, 90th ${Math.round(s1.p90)} ms)`, s1.p50 < 60 && s1.p90 < 150);

  // The caches behind that speed stay right through every kind of change.
  const s2 = await p.evaluate(async N => {
    const S = AtlasSheets, out = {};
    const sum = () => S.valueOf('E1');
    const base = sum();
    // Row n holds (n - 1) % 97 in column B: B3 is 2, B4 is 3.
    S.setCell('B3', '');                      // emptied
    out.emptied = sum() === base - 2;
    S.setCell('B3', '=B4*2');                 // a formula inside the summed column
    out.formula = sum() === base - 2 + 6;
    S.setCell('B4', '10');                    // …that depends on another cell
    out.chain = sum() === base - 2 + 20 - 3 + 10;
    S.setCell('B3', '2'); S.setCell('B4', '3');
    out.back = sum() === base;
    return out;
  }, N);
  check('…when a value is emptied', s2.emptied);
  check('…when a formula goes into the summed column', s2.formula);
  check('…when that formula\'s input changes', s2.chain);
  check('…and back', s2.back);
  await p.evaluate(() => { window.__atlasDirty = false; });

  await p.goto(BASE + 'sheets.html');
  await p.waitForFunction(() => window.AtlasSheets);
  const s3 = await p.evaluate(async () => {
    const S = AtlasSheets, out = {};
    for (let i = 1; i <= 200; i++) { S.setCell('A' + i, String(201 - i)); }
    S.setCell('C1', '=SUM(A1:A200)');
    S.setCell('C2', '=MAX(A1:A200)');
    S.setCell('C3', '=SUMPRODUCT((A1:A200>100)*1)');
    out.start = [S.valueOf('C1'), S.valueOf('C2'), S.valueOf('C3')].join();
    return out;
  });
  // Sort column A through the ribbon (sorting moves cells without setCell).
  await p.fill('#nameBox', 'A1:A200');
  await p.press('#nameBox', 'Enter');
  await p.click('.ribbon-tabs button[data-tab="home"]');
  await p.click('#sortAsc');
  const s4 = await p.evaluate(() => { const S = AtlasSheets; S.setCell('D1', '=SUM(A1:A100)'); return { a1: S.valueOf('A1'), half: S.valueOf('D1'), sums: [S.valueOf('C1'), S.valueOf('C2'), S.valueOf('C3')].join() }; });
  check('sorting keeps every total right', s4.sums === s3.start && s4.a1 === 1 && s4.half === 5050, JSON.stringify({ s3, s4 }));
  const s5 = await p.evaluate(() => {
    const S = AtlasSheets;
    // A formula that spills into the summed column.
    S.setCell('E1', '=SEQUENCE(100)');
    S.setCell('F1', '=SUM(E1:E100)');
    const spill = S.valueOf('F1');
    S.setCell('E1', '=SEQUENCE(100,1,2)');
    const spill2 = S.valueOf('F1');
    return { spill, spill2 };
  });
  check('a spilled column sums right, and after it changes', s5.spill === 5050 && s5.spill2 === 5150, JSON.stringify(s5));
  await p.evaluate(() => { window.__atlasDirty = false; });

  // ---- Write
  await p.goto(BASE + 'write.html');
  await p.waitForFunction(() => window.AtlasWrite);
  const w1 = await p.evaluate(async () => {
    const para = i => `<p>Paragraph ${i}. ` + 'The quick brown fox jumps over the lazy dog while the committee reviews the budget for next year. '.repeat(3) + '</p>';
    let html = '';
    for (let i = 0; i < 3400; i++) html += (i % 40 === 0 ? `<h2>Section ${i / 40 + 1}</h2>` : '') + para(i);
    const t = performance.now();
    await AtlasWrite.loadFile({ name: 'long.html', path: null, data: new TextEncoder().encode('<!doctype html><html><body>' + html + '</body></html>') });
    await new Promise(r => setTimeout(r, 500));
    const margins = () => [...document.querySelectorAll('#doc [data-pg]')].map(e => e.style.marginTop).join();
    return { load: performance.now() - t, pages: document.querySelectorAll('.page-bg').length, margins: margins().length };
  });
  check(`a ${w1.pages}-page document opens quickly (${Math.round(w1.load)} ms)`, w1.pages >= 240 && w1.load < 8000);
  await p.click('#doc p:nth-of-type(3)');
  const w2 = await p.evaluate(async () => {
    const long = [];
    const po = new PerformanceObserver(l => l.getEntries().forEach(e => long.push(Math.round(e.duration))));
    po.observe({ entryTypes: ['longtask'] });
    for (let i = 0; i < 8; i++) { document.execCommand('insertText', false, 'more words here and there '); await new Promise(r => setTimeout(r, 40)); }
    await new Promise(r => setTimeout(r, 800));
    po.disconnect();
    return { long, pages: document.querySelectorAll('.page-bg').length };
  });
  check(`typing near the start of it never stalls (${w2.long.join(', ') || 'no'} long tasks)`, w2.long.every(d => d < 250), w2.long.join());
  // Laying out again without a change moves nothing; a full fresh layout gives the same pages.
  const w3 = await p.evaluate(async () => {
    const snap = () => [...document.querySelectorAll('#doc > *')].map(e => (e.dataset.pg ? parseFloat(e.style.marginTop) : null));
    const same = (x, y) => x.length === y.length && x.every((v, i) => (v === null) === (y[i] === null) && (v === null || Math.abs(v - y[i]) <= 0.05));
    window.dispatchEvent(new Event('resize'));
    await new Promise(r => setTimeout(r, 400));
    const a = snap(), pagesA = document.querySelectorAll('.page-bg').length;
    // Throw away every page break and lay out from scratch.
    document.querySelectorAll('#doc [data-pg]').forEach(el => { el.style.marginTop = el.dataset.pgMt || ''; delete el.dataset.pg; delete el.dataset.pgMt; delete el.dataset.pgG; });
    window.dispatchEvent(new Event('resize'));
    await new Promise(r => setTimeout(r, 400));
    return { same: same(a, snap()), pagesA, pagesB: document.querySelectorAll('.page-bg').length };
  });
  check('pagination from the current breaks equals pagination from scratch', w3.same && w3.pagesA === w3.pagesB, JSON.stringify({ pagesA: w3.pagesA, pagesB: w3.pagesB, same: w3.same }));
  const words = await p.evaluate(async () => { await new Promise(r => setTimeout(r, 500)); return document.getElementById('stWords').textContent; });
  check('the word count catches up after typing', /^\d[\d,]* words$/.test(words) && parseInt(words.replace(/,/g, ''), 10) > 100000, words);
  await p.evaluate(() => { window.__atlasDirty = false; });

  check('no page errors', !errors.length, errors.slice(0, 3).join(' | '));
  await browser.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
