// The start screen: greeting, template miniatures, thumbnails of recent files (taken when a file
// is opened or saved), pinning, and searching the recent list.
const { chromium } = require('playwright');
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
const BASE = process.env.E2E_BASE + '/src/renderer/';

(async () => {
  const browser = await chromium.launch();
  const errors = [];
  const ctx = await browser.newContext({ viewport: { width: 1280, height: 900 } });
  await ctx.addInitScript(() => { if (!localStorage.getItem('atlas-user-name')) localStorage.setItem('atlas-user-name', 'Ada Lovelace'); });
  const p = await ctx.newPage();
  p.on('pageerror', e => errors.push(e.message));
  p.on('dialog', d => { errors.push('dialog: ' + d.message()); d.dismiss(); });

  await p.goto(BASE + 'index.html');
  const hello = await p.evaluate(() => [document.getElementById('greeting').textContent, document.getElementById('today').textContent]);
  check('greets the person by first name, with today\'s date', /^Good (morning|afternoon|evening), Ada$/.test(hello[0]) && hello[1].length > 5, hello.join(' / '));
  const tpl = await p.$$eval('#templates .tpl', ts => ts.map(t => [t.querySelector(':scope > b').textContent, !!t.querySelector('.thumbnail .paper-write h1, .thumbnail .paper-write p, .thumbnail .paper-sheets td, .thumbnail .paper-slides .box')]));
  check('templates show a miniature of their first page, sheet or slide', tpl.length === 5 && tpl.every(t => t[1] && t[0]), JSON.stringify(tpl));
  const scale = await p.$eval('#templates .tpl .paper', e => e.style.transform);
  check('…scaled to fit the card', /^scale\(0\.\d+\)$/.test(scale), scale);

  // ---- Opening files takes their thumbnails (text only, nothing from the file runs)
  await p.goto(BASE + 'write.html');
  await p.waitForFunction(() => window.AtlasWrite);
  const html = '<h1>Trip plan</h1><p style="text-align:center">Day one &lt;b&gt;not bold&lt;/b&gt;</p><ul><li>Pack</li><li>Go</li></ul><p><img src="x" onerror="window.__pwned=1">tail</p>';
  await p.evaluate(h => AtlasWrite.loadFile({ name: 'Trip.html', path: null, data: new TextEncoder().encode('<!doctype html><html><body>' + h + '</body></html>') }), html);
  await p.waitForFunction(() => Platform.thumbs.get('Trip.html'), null, { timeout: 5000 });
  const wt = await p.evaluate(() => Platform.thumbs.get('Trip.html'));
  check('Write: opening a file stores its first blocks as text', wt.kind === 'write' && wt.blocks[0].t === 'h1' && wt.blocks[0].s === 'Trip plan' && wt.blocks[1].a === 'center' && wt.blocks.filter(b => b.t === 'li').length === 2, JSON.stringify(wt.blocks));
  await p.evaluate(() => { window.__atlasDirty = false; });

  await p.goto(BASE + 'sheets.html');
  await p.waitForFunction(() => window.AtlasSheets);
  await p.evaluate(() => AtlasSheets.loadFile({ name: 'Budget.csv', path: null, data: new TextEncoder().encode('Item,Cost\nRent,1200\nFood,300\n') }));
  await p.waitForFunction(() => Platform.thumbs.get('Budget.csv'), null, { timeout: 5000 });
  const st = await p.evaluate(() => Platform.thumbs.get('Budget.csv'));
  check('Sheets: the top-left cells as shown', st.kind === 'sheets' && st.rows[0][0] === 'Item' && st.rows[1][1] === '1200' && st.rows[2][0] === 'Food', JSON.stringify(st.rows.slice(0, 3)));
  await p.evaluate(() => { window.__atlasDirty = false; });

  await p.goto(BASE + 'slides.html');
  await p.waitForFunction(() => window.AtlasSlides);
  await p.evaluate(() => { const s = AtlasSlides.pres.slides[0]; s.elements[0].html = 'Hello deck'; Platform.thumbs.capture('Deck.pptx'); });
  const sl = await p.evaluate(() => Platform.thumbs.get('Deck.pptx'));
  check('Slides: the first slide\'s boxes, colours and words', sl.kind === 'slides' && sl.els.some(e => e.s === 'Hello deck' && e.size > 30 && e.w > 300) && /^rgb/.test(sl.bg), JSON.stringify(sl).slice(0, 300));
  await p.evaluate(() => {
    localStorage.setItem('atlas-recent', JSON.stringify([
      { name: 'Deck.pptx', app: 'slides', opened: Date.now() - 1000 },
      ...JSON.parse(localStorage.getItem('atlas-recent') || '[]').filter(r => r.name !== 'Deck.pptx')
    ]));
  });

  // ---- The start screen shows them
  await p.goto(BASE + 'index.html');
  await p.waitForSelector('#recent .file-card');
  const cards = await p.$$eval('#recent .file-card', cs => cs.map(c => [c.querySelector('.n').textContent, c.querySelector('.thumbnail').dataset.kind, c.querySelector('.thumbnail').textContent.replace(/\s+/g, ' ').trim().slice(0, 60)]));
  const names = cards.map(c => c[0]);
  check('recent files are cards, newest first', names.slice(0, 3).join() === 'Deck.pptx,Budget.csv,Trip.html', names.join());
  const trip = cards.find(c => c[0] === 'Trip.html');
  check('…each with a miniature of its contents', trip && trip[1] === 'write' && /Trip plan/.test(trip[2]) && cards.find(c => c[0] === 'Budget.csv')[2].includes('Rent'), JSON.stringify(cards));
  const safe = await p.evaluate(() => { const t = [...document.querySelectorAll('#recent .file-card')].find(c => c.querySelector('.n').textContent === 'Trip.html').querySelector('.paper'); return [!!t.querySelector('img, b'), /<b>not bold<\/b>/.test(t.textContent), !!window.__pwned]; });
  check('file contents are drawn as text, never as markup', !safe[0] && safe[1] && !safe[2], safe.join());
  const slideBox = await p.evaluate(() => { const c = [...document.querySelectorAll('#recent .file-card')].find(c => c.querySelector('.n').textContent === 'Deck.pptx'); const b = c.querySelector('.thumbnail').getBoundingClientRect(), s = c.querySelector('.paper').getBoundingClientRect(); return [Math.round(b.width - s.width), Math.round((b.height - s.height) / 2 - (s.top - b.top))]; });
  check('a slide fills the card\'s width, centred', Math.abs(slideBox[0]) <= 1 && Math.abs(slideBox[1]) <= 1, slideBox.join());

  // ---- Pin, search
  check('nothing pinned at first', await p.evaluate(() => document.getElementById('pinnedBox').hidden));
  await p.hover('#recent .file-card:has-text("Budget.csv")');
  await p.click('#recent .file-card:has-text("Budget.csv") .pin');
  const after = await p.evaluate(() => [document.getElementById('pinnedBox').hidden, [...document.querySelectorAll('#pinned .n')].map(n => n.textContent).join(), [...document.querySelectorAll('#recent .n')].map(n => n.textContent).includes('Budget.csv'), document.querySelector('#pinned .pin').getAttribute('aria-pressed')]);
  check('pinning moves a file to Pinned', !after[0] && after[1] === 'Budget.csv' && !after[2] && after[3] === 'true', JSON.stringify(after));
  await p.reload();
  await p.waitForSelector('#pinned .file-card');
  check('…and stays pinned', (await p.$$eval('#pinned .n', ns => ns.map(n => n.textContent).join())) === 'Budget.csv');
  await p.fill('#findRecent', 'trip');
  check('searching filters the recent list', (await p.$$eval('#recent .n', ns => ns.map(n => n.textContent).join())) === 'Trip.html');
  await p.fill('#findRecent', 'zzz');
  check('…and says when nothing matches', (await p.textContent('#recent .empty')) === 'No recent files match.');
  await p.fill('#findRecent', '');
  await p.click('#pinned .pin');
  check('unpinning puts it back', await p.evaluate(() => document.getElementById('pinnedBox').hidden && [...document.querySelectorAll('#recent .n')].some(n => n.textContent === 'Budget.csv')));
  await p.focus('#recent .file-card');
  await p.keyboard.press('Enter');
  await sleep(200);
  check('a card opens with the keyboard (on the web: a hint to use Open)', await p.evaluate(() => /Open a file/.test(document.querySelector('.toast') && document.querySelector('.toast').textContent)));

  // ---- The template gallery in Write uses the same miniatures
  await p.goto(BASE + 'write.html');
  await p.click('.ribbon-tabs .file-tab');
  await p.click('.backstage nav button[data-cmd="new-template"]');
  check('the Write template gallery shows page miniatures', await p.$$eval('.template-gallery .tpl', ts => ts.length === 10 && ts.every(t => t.querySelector('.thumbnail .paper'))));

  check('no page errors', !errors.length, errors.slice(0, 3).join(' | '));
  await browser.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
