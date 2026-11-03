// Regression tests for the issues fixed in the review pass.
const { chromium } = require('playwright');
const BASE = process.env.E2E_BASE + '/src/renderer/';
(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1400, height: 880 } });
  const errors = [];
  page.on('pageerror', e => errors.push('pageerror: ' + e.message));
  page.on('console', m => { if (m.type() === 'error') errors.push('console: ' + m.text()); });
  page.on('dialog', d => { errors.push('dialog fired: ' + d.message()); d.dismiss(); });
  const check = (name, cond, extra = '') => console.log((cond ? 'PASS ' : 'FAIL ') + name + (extra !== '' ? ' :: ' + extra : ''));
  const go = async url => { try { await page.evaluate(() => { window.__atlasDirty = false; }); } catch {} await page.goto(url); };
  const timed = async fn => { const t = Date.now(); const v = await fn(); return [v, Date.now() - t]; };

  // ---------- Sheets: big data ----------
  await go(BASE + 'sheets.html');
  const [bigCsv, tCsv] = await timed(() => page.evaluate(async () => {
    let text = 'id,value,total\n';
    for (let i = 1; i <= 100000; i++) text += `${i},${i % 7},${i === 1 ? '=B2' : `=C${i}+B${i + 1}`}\n`;
    await AtlasSheets.loadFile({ name: 'big.csv', path: null, data: new TextEncoder().encode(text) });
    return { stats: AtlasSheets.domStats(), last: AtlasSheets.valueOf('C100001') };
  }));
  check('sheets: 100k-row CSV opens quickly', tCsv < 8000, tCsv + ' ms');
  check('sheets: only visible rows are in the DOM', bigCsv.stats.rows < 200, JSON.stringify(bigCsv.stats));
  const expected = Array.from({ length: 100000 }, (_, i) => (i + 1) % 7).reduce((a, b) => a + b, 0);
  check('sheets: 100k-long running total computes', bigCsv.last === expected, bigCsv.last + ' vs ' + expected);
  // Jump far down with the Name Box and check the right row is rendered
  await page.fill('#nameBox', 'B90000');
  await page.press('#nameBox', 'Enter');
  await page.waitForTimeout(150);
  const jumped = await page.evaluate(() => {
    const tr = document.querySelector('#grid tr[data-r="89999"]');
    const td = tr && tr.querySelector('td[data-c="1"]');
    const box = document.getElementById('selBox').getBoundingClientRect();
    const r = td && td.getBoundingClientRect();
    return { text: td && td.textContent, dy: r ? Math.abs(r.top - box.top) : -1, dx: r ? Math.abs(r.left - box.left) : -1 };
  });
  check('sheets: Name Box jump renders the target row', jumped.text === String(89999 % 7), JSON.stringify(jumped));
  check('sheets: selection box lines up with the cell', jumped.dy <= 2 && jumped.dx <= 2, JSON.stringify(jumped));
  // Ctrl+A then bold must not hang
  const [, tBold] = await timed(async () => { await page.focus('#gridWrap'); await page.keyboard.press('Control+a'); await page.keyboard.press('Control+b'); });
  check('sheets: select-all + bold on 300k cells is fast', tBold < 8000, tBold + ' ms');
  // Jump to the last cell Excel allows
  const [, tFar] = await timed(async () => { await page.fill('#nameBox', 'XFD1048576'); await page.press('#nameBox', 'Enter'); });
  const far = await page.evaluate(() => ({ stats: AtlasSheets.domStats(), name: document.getElementById('nameBox').value }));
  check('sheets: jump to XFD1048576 stays light', tFar < 5000 && far.stats.cells < 20000 && far.name === 'XFD1048576', tFar + ' ms ' + JSON.stringify(far));

  // ---------- Sheets: hostile/odd files ----------
  await go(BASE + 'sheets.html');
  const hostileBytes = await (async () => {
    // Built in Node: SheetJS can't write a sheet whose declared range is the whole grid.
    const XLSX = require('xlsx');
    const JSZip = require('jszip');
    const wb = XLSX.utils.book_new();
    const ws = { A1: { t: 's', v: 'hello' }, B2: { t: 'n', v: 5 }, '!ref': 'A1:B2' };
    XLSX.utils.book_append_sheet(wb, ws, 'Data');
    const styles = JSON.parse('{"__proto__":{"b":true}}');
    styles.A1 = { color: 'red;background:url(x)', b: true, font: '"><img src=x>' };
    styles.ZZZZ1 = { b: true };
    const colWidths = JSON.parse('{"__proto__":5,"2":120}');
    const meta = JSON.stringify({ active: 0, sheets: [{ name: 'Data', styles, charts: [{ type: 'evil', range: 'A1:B2' }, { type: 'bar', range: 'A1:B2', x: 'NaN' }], colWidths }] });
    wb.Custprops = { AtlasMeta: meta };
    const zip = await JSZip.loadAsync(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
    const sheetXml = (await zip.file('xl/worksheets/sheet1.xml').async('string')).replace(/<dimension ref="[^"]*"\/>/, '<dimension ref="A1:XFD1048576"/>');
    zip.file('xl/worksheets/sheet1.xml', sheetXml);
    return Array.from(await zip.generateAsync({ type: 'uint8array' }));
  })();
  const hostile = await page.evaluate(async arr => {
    const bytes = new Uint8Array(arr);
    const t = Date.now();
    await AtlasSheets.loadFile({ name: 'hostile.xlsx', path: null, data: bytes });
    const s = AtlasSheets.book.sheets[0];
    return { ms: Date.now() - t, ref: XLSX.read(bytes, { type: 'array' }).Sheets.Data['!ref'], polluted: ({}).b === true || Object.getPrototypeOf(s.cells) !== Object.prototype || Object.getPrototypeOf(s.colWidths) !== Object.prototype, a1: s.cells.A1, keys: Object.keys(s.cells), charts: s.charts, widths: s.colWidths, rows: s.rows };
  }, hostileBytes);
  check('sheets: test file really declares the whole grid', hostile.ref === 'A1:XFD1048576', hostile.ref);
  check('sheets: huge declared range opens fast', hostile.ms < 3000, hostile.ms + ' ms');
  check('sheets: no prototype pollution from file metadata', !hostile.polluted);
  check('sheets: unsafe style values dropped', hostile.a1 && hostile.a1.s && hostile.a1.s.b === true && !hostile.a1.s.color && !hostile.a1.s.font, JSON.stringify(hostile.a1));
  check('sheets: invalid cell keys ignored', hostile.keys.join() === 'A1,B2', hostile.keys.join());
  check('sheets: invalid charts dropped', hostile.charts.length === 1 && hostile.charts[0].x === 100, JSON.stringify(hostile.charts));
  check('sheets: sheet size from real cells, not declared range', hostile.rows === 100, String(hostile.rows));
  const csvName = await page.evaluate(async () => {
    await AtlasSheets.loadFile({ name: 'Q1 [draft]: sales*.csv', path: null, data: new TextEncoder().encode('a,b\n1,2\n') });
    const name = AtlasSheets.book.sheets[0].name;
    let ok = true;
    try { await AtlasSheets.serialize('xlsx'); } catch { ok = false; }
    return { name, ok };
  });
  check('sheets: CSV with odd file name still saves as xlsx', csvName.ok && !/[\[\]:*]/.test(csvName.name), JSON.stringify(csvName));

  // ---------- Sheets: wrap, row height, freeze, scientific ----------
  await go(BASE + 'sheets.html');
  await page.evaluate(() => AtlasSheets.setCell('A1', 'This is a long piece of text that should wrap onto several lines in the cell'));
  await page.click('#grid tr[data-r="0"] td[data-c="0"]');
  await page.click('[data-style="wrap"]');
  const wrapInfo = await page.evaluate(() => {
    const tr = document.querySelector('#grid tr[data-r="0"]');
    const box = document.getElementById('selBox').getBoundingClientRect();
    const td = tr.querySelector('td[data-c="0"]').getBoundingClientRect();
    return { h: tr.getBoundingClientRect().height, boxH: box.height, tdH: td.height };
  });
  check('sheets: wrapped text makes the row taller', wrapInfo.h > 30, JSON.stringify(wrapInfo));
  check('sheets: selection box matches the taller row', Math.abs(wrapInfo.boxH - wrapInfo.tdH) <= 3, JSON.stringify(wrapInfo));
  await page.click('[data-tab="view"]');
  await page.click('#freezeRow');
  await page.evaluate(() => { AtlasSheets.setCell('B3', '12345.678'); });
  await page.click('[data-tab="home"]');
  await page.click('#grid tr[data-r="2"] td[data-c="1"]');
  await page.selectOption('#numFmt', '0.00E+00');
  const sci = await page.textContent('#grid tr[data-r="2"] td[data-c="1"]');
  check('sheets: scientific format shows 1.23E+04', sci === '1.23E+04', sci);
  const frozenRt = await page.evaluate(async () => {
    const bytes = await AtlasSheets.serialize('xlsx');
    await AtlasSheets.loadFile({ name: 'f.xlsx', path: null, data: bytes });
    const s = AtlasSheets.book.sheets[0];
    return { freeze: s.freezeRow, wrap: s.cells.A1.s.wrap, height: s.rowHeights[0], frozenDom: !!document.querySelector('#grid thead tr.frozen') };
  });
  check('sheets: freeze, wrap and row height survive save/reopen', frozenRt.freeze && frozenRt.wrap && frozenRt.height > 30 && frozenRt.frozenDom, JSON.stringify(frozenRt));
  // Typing still works after the rewrite
  await page.click('#grid tr[data-r="5"] td[data-c="3"]');
  await page.keyboard.type('=1+2');
  await page.keyboard.press('Enter');
  const typed = await page.evaluate(() => ({ v: AtlasSheets.valueOf('D6'), name: document.getElementById('nameBox').value }));
  check('sheets: typing a formula and Enter moves down', typed.v === 3 && typed.name === 'D7', JSON.stringify(typed));

  // ---------- Write ----------
  await go(BASE + 'write.html');
  const evilHtml = `<!doctype html><html><head><meta name="atlas-setup" content='{"size":"__proto__","color":"red;}</style><script>alert(1)</script><style>","margin":"9in"}'></head>
    <body><header class="atlas-header">Old header</header><main class="atlas-doc"><p id="doc">x</p><div class="backstage open" style="position:fixed;inset:0">overlay</div>
    <a href="java&#10;script:alert(1)">bad link</a><a href="https://example.com">good link</a><svg><a xlink:href="javascript:alert(1)"><text>t</text></a></svg>
    <img src="https://tracker.example/pixel.png"><p>Heading words here</p></main></body></html>`;
  const w = await page.evaluate(async html => {
    await AtlasWrite.loadFile({ name: 'evil.html', path: null, data: new TextEncoder().encode(html) });
    const out = await AtlasWrite.serialize('html');
    const doc = document.getElementById('doc');
    return {
      exported: out, inner: doc.innerHTML, ids: document.querySelectorAll('#doc').length,
      page: getComputedStyle(document.getElementById('page')).width, header: document.getElementById('pageHeader').textContent
    };
  }, evilHtml);
  check('write: hostile page setup ignored (no script in export)', !/<script/i.test(w.exported) && !/alert/.test(w.exported), (w.exported.match(/atlas-setup" content="[^"]*/) || [''])[0]);
  check('write: page size stays valid', w.page === '816px', w.page);
  check('write: id collisions with the app removed', w.ids === 1);
  check('write: overlay class and position removed', !/backstage|position/.test(w.inner), w.inner.slice(0, 200));
  check('write: obfuscated javascript: link removed', !/javascript/i.test(w.inner) && /href="https:\/\/example.com"/.test(w.inner));
  check('write: svg and remote images removed', !/<svg|tracker/.test(w.inner));
  // Loading a docx clears the previous header
  const docxBytes = await page.evaluate(async () => {
    document.getElementById('doc').innerHTML = '<p>plain body</p>';
    document.getElementById('pageHeader').innerHTML = '';
    return Array.from(await AtlasWrite.serialize('docx'));
  });
  await page.evaluate(async () => { const h = document.getElementById('pageHeader'); h.hidden = false; h.textContent = 'Stale header'; });
  const afterDocx = await page.evaluate(async bytes => { await AtlasWrite.loadFile({ name: 'x.docx', path: null, data: new Uint8Array(bytes) }); return { h: document.getElementById('pageHeader').textContent, hidden: document.getElementById('pageHeader').hidden }; }, docxBytes);
  check('write: header from the previous document is cleared', afterDocx.h === '' && afterDocx.hidden, JSON.stringify(afterDocx));
  // Find only matches inside the document (the ribbon also says "Heading")
  await page.evaluate(() => { document.getElementById('doc').innerHTML = '<p>Intro</p><p>A Heading in <b>bo</b>ld text</p>'; });
  await page.keyboard.press('Control+h').catch(() => {});
  await page.evaluate(() => Platform.emit('replace'));
  await page.fill('#findText', 'heading');
  await page.press('#findText', 'Enter');
  const found = await page.evaluate(() => { const h = CSS.highlights.get('find-match'); const r = h && [...h][0]; return { text: r ? r.toString() : '', inDoc: !!r && document.getElementById('doc').contains(r.startContainer), docText: document.getElementById('doc').textContent, focus: document.activeElement.id }; });
  check('write: find highlights the match in the document and keeps focus in the find box', found.text === 'Heading' && found.inDoc && found.focus === 'findText' && found.docText.includes('Heading'), JSON.stringify(found));
  await page.fill('#findText', 'bold');
  await page.press('#findText', 'Enter');
  const across = await page.evaluate(() => { const h = CSS.highlights.get('find-match'); return h ? [...h][0].toString() : ''; });
  await page.fill('#replaceText', 'BOLD');
  await page.evaluate(() => document.getElementById('replaceOne').click());
  const replaced = await page.evaluate(() => document.getElementById('doc').textContent);
  check('write: replace works while typing in the find bar', replaced.includes('BOLD') && !replaced.includes('bold'), replaced);
  check('write: find matches across formatting', across === 'bold', across);
  // Save As without an extension (web) keeps a real format
  const named = await page.evaluate(async () => {
    UI.prompt = async () => 'report';
    return Platform.chooseSavePath('Document1.docx', [{ name: 'Word', extensions: ['docx'] }, { name: 'Text', extensions: ['txt'] }]);
  });
  check('write: Save As adds the extension', named && named.name === 'report.docx', JSON.stringify(named));

  // ---------- Slides ----------
  await go(BASE + 'slides.html');
  const evilSlides = JSON.stringify({ theme: '__proto__', slides: [{ bg: 'red;}', bgImage: 'https://tracker/x.png', elements: [
    { id: 'a"]x', type: 'shape', shape: 'rect', x: 10, y: 10, w: 100, h: 100, fill: '"/><img src=x onerror=alert(1)><g x="', html: '<img src=x onerror=alert(1)>hi<script>alert(2)</script>', size: 20 },
    { type: 'image', src: 'javascript:alert(1)', x: 0, y: 0, w: 10, h: 10 },
    { type: 'text', html: '<b>ok</b>', x: 'NaN', size: 1e9, font: "x'</style>" }
  ] }] });
  const sl = await page.evaluate(async json => {
    await AtlasSlides.loadFile({ name: 'evil.aslides', path: null, data: new TextEncoder().encode(json) });
    const p = AtlasSlides.pres;
    return { theme: p.theme, els: p.slides[0].elements, stage: document.getElementById('stage').innerHTML, bg: p.slides[0].bg, bgImage: p.slides[0].bgImage };
  }, evilSlides);
  check('slides: unknown theme falls back', sl.theme === 'office');
  check('slides: bad image element dropped', sl.els.length === 2, String(sl.els.length));
  check('slides: no injected markup in the stage', !/onerror|<script|<img/i.test(sl.stage), sl.stage.slice(0, 300));
  check('slides: unsafe colours and backgrounds dropped', sl.els[0].fill === null && sl.bg === null && sl.bgImage === null);
  check('slides: numbers and fonts validated', sl.els[1].x === 100 && sl.els[1].size === 400 && !sl.els[1].font, JSON.stringify(sl.els[1]));
  check('slides: ids regenerated', !sl.els.some(e => /"/.test(e.id)));

  check('no unexpected page errors or dialogs', errors.length === 0, errors.join(' | '));
  await browser.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.error('CRASH', e); process.exit(1); });
