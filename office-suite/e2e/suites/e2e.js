const { chromium } = require('playwright');
const SHOTS = process.env.E2E_WORK;
const BASE = process.env.E2E_BASE + '/src/renderer/';
(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1400, height: 880 } });
  const errors = [];
  page.on('pageerror', e => errors.push('pageerror: ' + e.message));
  page.on('console', m => { if (m.type() === 'error') errors.push('console: ' + m.text()); });
  const check = (name, cond, extra = '') => console.log((cond ? 'PASS ' : 'FAIL ') + name + (extra ? ' :: ' + extra : ''));

  // Home
  await page.goto(BASE + 'index.html');
  await page.screenshot({ path: SHOTS + '/home.png' });

  // ---------- Write ----------
  await page.goto(BASE + 'write.html');
  await page.click('#doc');
  await page.keyboard.type('Quarterly Report');
  await page.click('.style-btn[data-block="H1"]:not([data-cls])');
  await page.keyboard.press('End');
  await page.keyboard.press('Enter');
  await page.keyboard.type('This is ');
  await page.keyboard.press('Control+b');
  await page.keyboard.type('bold');
  await page.keyboard.press('Control+b');
  await page.keyboard.type(' text in Isadek Write.');
  await page.click('[data-tab="insert"]');
  page.once('dialog', d => d.dismiss());
  await page.click('#insTable');
  await page.click('.modal .ok');
  await page.click('[data-tab="home"]');
  const docHtml = await page.$eval('#doc', d => d.innerHTML);
  check('write: heading created', /<h1/.test(docHtml), docHtml.slice(0, 120));
  check('write: bold applied', /font-weight: bold|<b>|font-weight: 700/.test(docHtml));
  check('write: table inserted', /<table/.test(docHtml));
  const words = await page.textContent('#stWords');
  check('write: word count', /\d+ words/.test(words), words);
  const docx = await page.evaluate(async () => Array.from(await AtlasWrite.serialize('docx')));
  check('write: docx produced', docx.length > 2000 && docx[0] === 0x50 && docx[1] === 0x4b, docx.length + ' bytes');
  await page.screenshot({ path: SHOTS + '/write.png' });
  // round trip docx
  await page.evaluate(async bytes => { await AtlasWrite.loadFile({ name: 'rt.docx', path: null, data: new Uint8Array(bytes) }); }, docx);
  const rt = await page.$eval('#doc', d => d.innerHTML);
  check('write: docx round-trip keeps heading', /<h1[^>]*>Quarterly Report/.test(rt), rt.slice(0, 200));
  check('write: docx round-trip keeps bold', /<strong>bold<\/strong>/.test(rt));
  check('write: docx round-trip keeps table', /<table/.test(rt));
  const md = await page.evaluate(() => AtlasWrite.markdownToHtml('# Hi\n\n- a **b**\n- c\n\n[x](https://e.com)'));
  check('write: markdown import', md === '<h1>Hi</h1><ul><li>a <b>b</b></li><li>c</li></ul><p><a href="https://e.com">x</a></p>', md);
  const san = await page.evaluate(() => AtlasWrite.sanitize('<p onclick="x()">hi<script>bad()</script><a href="javascript:alert(1)">l</a><img src=x onerror=alert(1)></p>'));
  check('write: sanitizer strips scripts/handlers', !/script|onclick|onerror|javascript/.test(san), san);

  // ---------- Sheets ----------
  await page.goto(BASE + 'sheets.html');
  await page.click('#grid tbody tr:nth-child(1) td:nth-child(2)');
  for (const v of ['Item', 'Apples', 'Pears', 'Plums']) { await page.keyboard.type(v); await page.keyboard.press('Enter'); }
  await page.click('#grid tbody tr:nth-child(1) td:nth-child(3)');
  for (const v of ['Qty', '10', '20', '30', '=SUM(B2:B4)', '=AVERAGE(B2:B4)*2']) { await page.keyboard.type(v); await page.keyboard.press('Enter'); }
  const b5 = await page.evaluate(() => AtlasSheets.valueOf('B5'));
  const b6 = await page.evaluate(() => AtlasSheets.valueOf('B6'));
  check('sheets: SUM typed into grid', b5 === 60, String(b5));
  check('sheets: AVERAGE*2', b6 === 40, String(b6));
  const shown = await page.textContent('#grid tbody tr:nth-child(5) td:nth-child(3)');
  check('sheets: displayed value', shown === '60', shown);
  // formatting: bold header
  await page.click('#grid tbody tr:nth-child(1) td:nth-child(2)');
  await page.keyboard.down('Shift'); await page.keyboard.press('ArrowRight'); await page.keyboard.up('Shift');
  await page.keyboard.press('Control+b');
  const bold = await page.evaluate(() => AtlasSheets.book.sheets[0].cells.A1.s);
  check('sheets: bold style', bold && bold.b === true, JSON.stringify(bold));
  const agg = await page.textContent('#stAgg');
  // currency format on B2:B4
  await page.click('#grid tbody tr:nth-child(2) td:nth-child(3)');
  await page.keyboard.down('Shift'); await page.keyboard.press('ArrowDown'); await page.keyboard.press('ArrowDown'); await page.keyboard.up('Shift');
  await page.click('[data-fmt="$#,##0.00"]');
  const fmtShown = await page.textContent('#grid tbody tr:nth-child(2) td:nth-child(3)');
  check('sheets: currency format', fmtShown === '$10.00', fmtShown);
  const aggText = await page.textContent('#stAgg');
  check('sheets: status bar sum', /Sum: 60/.test(aggText), aggText);
  // copy-paste formula shifts
  await page.evaluate(() => { AtlasSheets.setCell('C2', '=B2*2'); });
  await page.click('#grid tbody tr:nth-child(2) td:nth-child(4)');
  await page.keyboard.press('Control+c');
  await page.click('#grid tbody tr:nth-child(3) td:nth-child(4)');
  await page.keyboard.press('Control+v');
  const c3 = await page.evaluate(() => AtlasSheets.book.sheets[0].cells.C3.v);
  check('sheets: paste shifts formula', c3 === '=B3*2', c3);
  // chart
  await page.click('#grid tbody tr:nth-child(1) td:nth-child(2)');
  await page.keyboard.down('Shift'); for (let i = 0; i < 3; i++) await page.keyboard.press('ArrowDown'); await page.keyboard.press('ArrowRight'); await page.keyboard.up('Shift');
  await page.click('[data-tab="insert"]');
  await page.click('[data-chart="bar"]');
  check('sheets: chart inserted', await page.$('.chart canvas') !== null);
  await page.click('[data-tab="home"]');
  await page.screenshot({ path: SHOTS + '/sheets.png' });
  const xlsx = await page.evaluate(async () => Array.from(await AtlasSheets.serialize('xlsx')));
  check('sheets: xlsx produced', xlsx.length > 1000 && xlsx[0] === 0x50, xlsx.length + ' bytes');
  await page.evaluate(async bytes => { await AtlasSheets.loadFile({ name: 'rt.xlsx', path: null, data: new Uint8Array(bytes) }); }, xlsx);
  const rtCells = await page.evaluate(() => AtlasSheets.book.sheets[0].cells);
  check('sheets: xlsx round-trip formula', rtCells.B5 && rtCells.B5.v === '=SUM(B2:B4)', JSON.stringify(rtCells.B5));
  check('sheets: xlsx round-trip number format', rtCells.B2 && rtCells.B2.s && rtCells.B2.s.fmt === '$#,##0.00', JSON.stringify(rtCells.B2));
  check('sheets: xlsx round-trip bold (custom props)', rtCells.A1 && rtCells.A1.s && rtCells.A1.s.b === true, JSON.stringify(rtCells.A1));
  check('sheets: xlsx round-trip chart', (await page.evaluate(() => AtlasSheets.book.sheets[0].charts.length)) === 1);
  const csv = await page.evaluate(async () => await AtlasSheets.serialize('csv'));
  check('sheets: csv export', csv.startsWith('Item,Qty,\r\nApples,10,20'), JSON.stringify(csv.slice(0, 40)));
  const ods = await page.evaluate(async () => (await AtlasSheets.serialize('ods')).length);
  check('sheets: ods export', ods > 1000, String(ods));

  // ---------- Slides ----------
  await page.goto(BASE + 'slides.html');
  await page.click('#stage .el.text >> nth=0');
  await page.keyboard.type('Isadek Slides Demo');
  await page.keyboard.press('Escape');
  await page.click('#stage .el.text >> nth=1');
  await page.keyboard.type('Built for every platform');
  await page.keyboard.press('Escape');
  await page.click('#newSlide');
  await page.click('#stage .el.text >> nth=0');
  await page.keyboard.type('Agenda');
  await page.keyboard.press('Escape');
  await page.click('#stage .el.text >> nth=1');
  await page.keyboard.type('First point');
  await page.keyboard.press('Enter');
  await page.keyboard.type('Second point');
  await page.keyboard.press('Escape');
  await page.click('[data-tab="insert"]');
  await page.click('[data-shape="ellipse"]');
  await page.click('[data-tab="design"]');
  await page.click('.theme-swatch >> nth=2');
  const pres = await page.evaluate(() => AtlasSlides.pres);
  check('slides: two slides', pres.slides.length === 2);
  check('slides: title text', /Isadek Slides Demo/.test(pres.slides[0].elements[0].html), pres.slides[0].elements[0].html);
  check('slides: bullets', /<li>First point<\/li><li>Second point<\/li>/.test(pres.slides[1].elements[1].html), pres.slides[1].elements[1].html);
  check('slides: shape added', pres.slides[1].elements.some(e => e.type === 'shape'));
  await page.screenshot({ path: SHOTS + '/slides.png' });
  const pptx = await page.evaluate(async () => Array.from(await AtlasSlides.serialize('pptx')));
  check('slides: pptx produced', pptx.length > 5000 && pptx[0] === 0x50, pptx.length + ' bytes');
  await page.evaluate(async bytes => { await AtlasSlides.loadFile({ name: 'rt.pptx', path: null, data: new Uint8Array(bytes) }); }, pptx);
  const rtp = await page.evaluate(() => AtlasSlides.pres);
  check('slides: pptx round-trip slide count', rtp.slides.length === 2, String(rtp.slides.length));
  const allText = JSON.stringify(rtp.slides.map(s => s.elements.map(e => e.html)));
  check('slides: pptx round-trip text', /Isadek Slides Demo/.test(allText) && /First point/.test(allText) && /Agenda/.test(allText), allText.slice(0, 300));
  check('slides: pptx round-trip shape', rtp.slides[1].elements.some(e => e.type === 'shape' && e.shape === 'ellipse'));
  await page.screenshot({ path: SHOTS + '/slides-reimported.png' });
  await page.evaluate(() => AtlasSlides.present(0));
  await page.waitForTimeout(400);
  await page.keyboard.press('ArrowRight');
  await page.waitForTimeout(700);
  await page.screenshot({ path: SHOTS + '/slideshow.png' });
  await page.keyboard.press('Escape');

  check('no page errors', errors.length === 0, errors.join(' | '));
  await browser.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.error('CRASH', e); process.exit(1); });
