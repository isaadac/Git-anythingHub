// More interface languages: Italian, and Arabic written right to left (the interface mirrored,
// documents in their own direction); tooltips translated; switching back to English.
const { chromium } = require('playwright');
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const BASE = process.env.E2E_BASE + '/src/renderer/';
(async () => {
  const b = await chromium.launch();
  const errors = [];
  const open = async (lang, app) => {
    const ctx = await b.newContext({ viewport: { width: 1300, height: 800 } });
    await ctx.addInitScript(l => localStorage.setItem('atlas-lang', l), lang);
    const p = await ctx.newPage();
    p.on('pageerror', e => errors.push(lang + ' ' + app + ': ' + e.message));
    await p.goto(BASE + app + '.html');
    await p.waitForFunction(() => window.I18n && document.querySelector('.ribbon-tabs button[data-tab="insert"]'));
    await p.waitForTimeout(300);
    return p;
  };
  const it = await open('it', 'write');
  check('Italian: the ribbon is in Italian', (await it.textContent('.ribbon-tabs button[data-tab="insert"]')) === 'Inserisci');
  check('…and so are the tooltips', (await it.getAttribute('[title^="Grassetto"]', 'title')) === 'Grassetto (Ctrl+B)');
  check('…left to right', (await it.evaluate(() => document.documentElement.dir)) === 'ltr');

  const ar = await open('ar', 'write');
  check('Arabic: the interface is right to left', (await ar.evaluate(() => [document.documentElement.dir, document.documentElement.lang].join())) === 'rtl,ar');
  check('…in Arabic', (await ar.textContent('.ribbon-tabs button[data-tab="insert"]')) === 'إدراج' && (await ar.getAttribute('[title^="غامق"]', 'title')) === 'غامق (Ctrl+B)');
  const tabs = await ar.$$eval('.ribbon-tabs button[data-tab]', bs => bs.map(x => Math.round(x.getBoundingClientRect().left)));
  check('…the tabs run from the right', tabs[0] > tabs[tabs.length - 1], tabs.join());
  await ar.click('#doc');
  await ar.keyboard.type('مرحبا بالعالم');
  await ar.keyboard.press('Enter');
  await ar.keyboard.type('Hello world');
  const pos = await ar.$$eval('#doc p', ps => ps.map(x => { const r = document.createRange(); r.selectNodeContents(x); const t = r.getBoundingClientRect(), p = x.getBoundingClientRect(); return [Math.round(t.left - p.left), Math.round(p.right - t.right)]; }));
  check('…each paragraph follows its own words: Arabic from the right, English from the left', pos[0][1] < 5 && pos[0][0] > 50 && pos[1][0] < 5 && pos[1][1] > 50, JSON.stringify(pos));
  await ar.click('.ai-btn');
  const pane = await ar.$eval('.ai-pane', x => [Math.round(x.getBoundingClientRect().left), innerWidth]);
  check('…panes open on the left', pane[0] < 40, pane.join());
  const sh = await open('ar', 'sheets');
  const cols = await sh.$$eval('.grid thead th[data-c]', t => t.slice(0, 3).map(x => [x.textContent.trim(), Math.round(x.getBoundingClientRect().left)]));
  check('…a spreadsheet keeps A, B, C from the left', cols[0][0].startsWith('A') && cols[0][1] < cols[1][1] && cols[1][1] < cols[2][1], JSON.stringify(cols));
  await ar.evaluate(() => I18n.set('en'));
  await ar.waitForTimeout(200);
  check('switching back to English puts everything back', (await ar.evaluate(() => document.documentElement.dir)) === 'ltr' && (await ar.textContent('.ribbon-tabs button[data-tab="insert"]')) === 'Insert');
  check('no page errors', !errors.length, errors.join(' | '));
  await b.close();
})().catch(e => { console.error('CRASH', e.stack); process.exit(1); });
