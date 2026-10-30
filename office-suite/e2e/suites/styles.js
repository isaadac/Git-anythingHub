// Chart styles, slide themes, and workbook/presentation templates.
const { chromium } = require('playwright');
const JSZip = require('jszip');
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
const BASE = process.env.E2E_BASE + '/src/renderer/';

(async () => {
  const browser = await chromium.launch();
  const errors = [];
  const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 } });
  const p = await ctx.newPage();
  p.on('pageerror', e => errors.push(e.message));

  // ---- Chart styles
  await p.goto(BASE + 'sheets.html');
  await p.waitForFunction(() => window.AtlasSheets);
  await p.evaluate(() => { [['Month', 'A', 'B'], ['Jan', 5, 3], ['Feb', 7, 4], ['Mar', 6, 8]].forEach((r, i) => r.forEach((v, c) => AtlasSheets.setCell('ABC'[c] + (i + 1), v))); });
  await p.fill('#nameBox', 'A1:C4');
  await p.press('#nameBox', 'Enter');
  await p.click('.ribbon-tabs button[data-tab="insert"]');
  await p.click('[data-chart="bar"]');
  const ch = () => p.evaluate(() => AtlasSheets.book.sheets[0].charts[0]);
  check('a new chart uses the Modern style', (await ch()).style === 'modern');
  // Whether the chart has pixels of exactly this colour (the first series' bars).
  const has = rgb => p.evaluate(rgb => {
    const cv = document.querySelector('.chart canvas'), d = cv.getContext('2d').getImageData(0, 0, cv.width, cv.height).data;
    for (let i = 0; i < d.length; i += 4) if (d[i] === rgb[0] && d[i + 1] === rgb[1] && d[i + 2] === rgb[2]) return true;
    return false;
  }, rgb);
  check('…drawn in its palette', await has([59, 111, 216]));
  await p.click('.chart .bar button[data-a="style"]'); // clicks land on the icon inside
  check('the style button opens the styles menu', await p.isVisible('.chart-styles'));
  const items = await p.$$eval('.chart-styles [role=menuitemradio]', bs => bs.map(b => b.textContent + ':' + b.getAttribute('aria-checked')));
  check('…with seven palettes, the current one checked', items.length === 7 && items[0] === 'Modern:true', items.join());
  await p.click('.chart-styles [data-style="sunset"]');
  check('choosing one restyles the chart', (await ch()).style === 'sunset' && await has([209, 73, 91]) && !(await has([59, 111, 216])));
  await p.click('.chart-styles .labels-toggle');
  check('data labels switch on', (await ch()).labels === true);
  await p.keyboard.press('Escape');
  check('Escape closes the menu', !(await p.isVisible('.chart-styles')));
  await p.click('.chart .bar button[data-a="type"]');
  check('the other chart buttons work through their icons too', (await ch()).type === 'line');
  // Saved and opened again: the style stays, and Excel gets the palette.
  const bytes = await p.evaluate(async () => Array.from(new Uint8Array(await AtlasSheets.serialize('xlsx'))));
  const zip = await JSZip.loadAsync(Buffer.from(bytes));
  const chartXml = await zip.file(Object.keys(zip.files).find(f => /charts\/chart\d+\.xml$/.test(f))).async('string');
  check('Excel gets the chart in the style\'s colours', /D1495B/.test(chartXml) && !/4472C4/.test(chartXml));
  await p.evaluate(async b => { await AtlasSheets.loadFile({ name: 'Styled.xlsx', path: null, data: new Uint8Array(b) }); }, bytes);
  const back = await ch();
  check('style and labels come back when the file is opened', back && back.style === 'sunset' && back.labels === true, JSON.stringify(back));
  await p.evaluate(() => { window.__atlasDirty = false; });

  // ---- Slide themes
  await p.goto(BASE + 'slides.html');
  await p.waitForFunction(() => window.AtlasSlides);
  await p.click('.ribbon-tabs button[data-tab="design"]');
  const sw = await p.$$eval('#themes .theme-swatch', bs => bs.map(b => b.dataset.theme + ':' + b.getAttribute('aria-pressed')));
  check('ten themes, the current one marked', sw.length === 10 && sw[0] === 'office:true', sw.join());
  await p.click('.theme-swatch[data-theme="paper"]');
  const look = await p.evaluate(() => { const t = document.querySelector('#stage [data-id] .content'); const st = getComputedStyle(document.getElementById('stage').querySelector('.slide') || document.getElementById('stage')); return [AtlasSlides.pres.theme, getComputedStyle(t).fontFamily, document.querySelector('.theme-swatch[data-theme="paper"]').getAttribute('aria-pressed')]; });
  check('a theme sets the deck\'s colours and fonts', look[0] === 'paper' && /Cambria/.test(look[1]) && look[2] === 'true', look.join(' | '));
  check('themes use fonts that ship with the app', await p.evaluate(() => [...document.querySelectorAll('#themes .theme-swatch .t')].every(t => /Calibri|Cambria|Arial/.test(t.style.fontFamily))));
  const fit = await p.evaluate(() => { const r = document.querySelector('.ribbon'); return [r.className, document.querySelector('.ribbon-panel.active').scrollWidth <= document.querySelector('.ribbon-panel.active').clientWidth + 1]; });
  check('the theme gallery fits without squeezing the ribbon', !/compact/.test(fit[0]) && fit[1], fit.join());
  await p.evaluate(() => { window.__atlasDirty = false; });

  // ---- Templates
  await p.goto(BASE + 'index.html');
  const all = await p.$$eval('#templates .tpl', ts => ts.map(t => (t.dataset.app || 'write') + ':' + t.querySelector(':scope > b').textContent));
  check('Home shows templates for every app', all.some(t => t.startsWith('write:')) && all.some(t => t.startsWith('sheets:')) && all.some(t => t.startsWith('slides:')), all.join());
  await p.click('#tplFilter [data-f="sheets"]');
  const wb = await p.$$eval('#templates .tpl', ts => ts.map(t => [t.dataset.app, t.querySelector('.thumbnail').dataset.kind, t.querySelector('.thumbnail').textContent.includes('$1,200.00') || t.textContent.includes('INVOICE')]));
  check('the Workbooks filter shows workbook miniatures', wb.length === 4 && wb.every(w => w[0] === 'sheets' && w[1] === 'sheets') && wb.some(w => w[2]), JSON.stringify(wb));
  await p.focus('#tplFilter [data-f="sheets"]');
  await p.keyboard.press('ArrowRight');
  check('arrow keys move between filters', await p.evaluate(() => document.querySelector('#tplFilter [aria-checked="true"]').dataset.f === 'slides' && document.activeElement.dataset.f === 'slides'));
  await p.click('#tplFilter [data-f="sheets"]');
  await p.click('.tpl:has-text("Monthly budget")');
  await p.waitForURL('**/sheets.html');
  await p.waitForFunction(() => window.AtlasSheets && AtlasSheets.book.sheets[0].name === 'Budget');
  const budget = await p.evaluate(() => ({ name: document.querySelector('.titlebar .docname input').value, total: AtlasSheets.valueOf('B10'), diff: AtlasSheets.valueOf('D10'), shown: document.querySelector('#grid tr[data-r="9"] td[data-c="1"]').textContent, charts: AtlasSheets.book.sheets[0].charts.length, bold: getComputedStyle(document.querySelector('#grid tr[data-r="2"] td[data-c="0"]')).fontWeight }));
  check('a workbook template opens as a new .xlsx with working formulas', budget.name === 'Monthly budget.xlsx' && budget.total === 2540 && budget.diff === -86, JSON.stringify(budget));
  check('…formatted, with its chart', budget.shown === '$2,540.00' && budget.charts === 1 && +budget.bold >= 700, JSON.stringify(budget));
  await p.evaluate(() => { window.__atlasDirty = false; });
  await p.goto(BASE + 'index.html');
  await p.click('#tplFilter [data-f="slides"]');
  await p.click('.tpl:has-text("Project update")');
  await p.waitForURL('**/slides.html');
  await p.waitForFunction(() => window.AtlasSlides && AtlasSlides.pres.slides.length > 1);
  const deck = await p.evaluate(() => ({ name: document.querySelector('.titlebar .docname input').value, n: AtlasSlides.pres.slides.length, theme: AtlasSlides.pres.theme, title: document.querySelector('#stage [data-id] .content').textContent }));
  check('a presentation template opens as a new .pptx in its theme', deck.name === 'Project update.pptx' && deck.n === 4 && deck.theme === 'paper' && deck.title === 'Project update', JSON.stringify(deck));
  // A marker file only means a template when it is exactly that.
  check('ordinary files are never mistaken for templates', await p.evaluate(() => !SuiteTemplates.fromBytes(new TextEncoder().encode('{"isadekTemplate":"slides","id":"nope"}'), 'slides') && !SuiteTemplates.fromBytes(new TextEncoder().encode('{"isadekTemplate":"sheets","id":"budget"}'), 'slides') && !!SuiteTemplates.fromBytes(new TextEncoder().encode('{"isadekTemplate":"slides","id":"pitch"}'), 'slides')));
  await p.evaluate(() => { window.__atlasDirty = false; });

  check('no page errors', !errors.length, errors.slice(0, 3).join(' | '));
  await browser.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
