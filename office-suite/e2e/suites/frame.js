// The window frame: ribbon display modes, the mini toolbar on a selection, a ribbon that compacts
// to fit narrower windows, and side panes as cards.
const { chromium } = require('playwright');
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
const BASE = process.env.E2E_BASE + '/src/renderer/';

(async () => {
  const browser = await chromium.launch();
  const errors = [];
  const ctx = await browser.newContext({ viewport: { width: 1440, height: 900 } });
  const p = await ctx.newPage();
  p.on('pageerror', e => errors.push(e.message));
  await p.goto(BASE + 'write.html');
  await p.waitForFunction(() => window.Platform && Platform.ribbonMode);

  // ---- Ribbon display options
  const ribbonH = () => p.evaluate(() => { const r = document.querySelector('.ribbon'); return getComputedStyle(r).display === 'none' ? 0 : r.getBoundingClientRect().height; });
  const full = await ribbonH();
  check('starts with the full ribbon', (await p.evaluate(() => document.body.classList.contains('ribbon-full'))) && full > 90, full);
  await p.click('.ribbon-mode');
  const items = await p.$$eval('.ribbon-menu [role=menuitemradio]', bs => bs.map(b => b.textContent.trim() + ':' + b.getAttribute('aria-checked')));
  check('the display menu offers full, single line and tabs only', items.join() === 'Full ribbon:true,Single line:false,Tabs only:false', items.join());
  await p.keyboard.press('ArrowDown');
  await p.keyboard.press('Enter');
  const line = await ribbonH();
  check('single line: one row of commands', await p.evaluate(() => document.body.classList.contains('ribbon-line')) && line > 0 && line < 60, line);
  const chromeH = await p.evaluate(() => [parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--chrome-h')), Math.round(document.querySelector('.ribbon').getBoundingClientRect().bottom)]);
  await sleep(100);
  const chromeH2 = await p.evaluate(() => [parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--chrome-h')), Math.round(document.querySelector('.ribbon').getBoundingClientRect().bottom)]);
  check('--chrome-h follows the ribbon\'s bottom edge', Math.abs(chromeH2[0] - chromeH2[1]) <= 1, JSON.stringify([chromeH, chromeH2]));
  await p.reload();
  await p.waitForFunction(() => window.Platform && Platform.ribbonMode);
  check('the choice is remembered', (await p.evaluate(() => Platform.ribbonMode.get())) === 'line');

  await p.evaluate(() => Platform.ribbonMode.set('tabs'));
  check('tabs only hides the commands', (await ribbonH()) === 0);
  await p.click('.ribbon-tabs button[data-tab]:not(.active)');
  check('a tab shows them over the document', await p.evaluate(() => document.body.classList.contains('ribbon-peek') && getComputedStyle(document.querySelector('.ribbon')).position === 'absolute'));
  await p.keyboard.press('Escape');
  check('…until Escape', (await ribbonH()) === 0);
  await p.keyboard.press('Control+F1');
  check('Ctrl+F1 brings the full ribbon back', (await p.evaluate(() => Platform.ribbonMode.get())) === 'full' && (await ribbonH()) > 90);
  await p.dblclick('.ribbon-tabs button[data-tab].active');
  check('double-clicking a tab collapses it', (await p.evaluate(() => Platform.ribbonMode.get())) === 'tabs');
  await p.evaluate(() => Platform.ribbonMode.set('full'));
  await p.click('.ribbon-tabs button[data-tab]');

  // ---- Mini toolbar
  await p.click('#doc');
  await p.keyboard.type('Make this word bold please');
  const word = await p.evaluate(() => { const t = document.querySelector('#doc p').firstChild; const i = t.data.indexOf('word'); const r = document.createRange(); r.setStart(t, i); r.setEnd(t, i + 4); return r.getBoundingClientRect().toJSON(); });
  await p.mouse.dblclick(word.x + word.width / 2, word.y + word.height / 2);
  await sleep(250);
  const bar = await p.evaluate(() => { const b = document.querySelector('.minibar'); const r = b.getBoundingClientRect(); return { hidden: b.hidden, top: r.top, bottom: r.bottom, n: b.querySelectorAll('.btn').length, sel: String(getSelection()) }; });
  check('selecting words shows the mini toolbar above them', !bar.hidden && bar.bottom <= word.y + 1 && bar.n >= 10, JSON.stringify(bar));
  await p.click('.minibar .btn[aria-label="Bold"]');
  check('its Bold makes the selection bold', await p.evaluate(() => document.queryCommandState('bold') && /(<(b|strong)>|font-weight: bold;">)word</.test(document.querySelector('#doc p').innerHTML)), await p.evaluate(() => document.querySelector('#doc p').innerHTML));
  check('…and shows as pressed', await p.evaluate(() => document.querySelector('.minibar .btn[aria-label="Bold"]').classList.contains('on')));
  await p.keyboard.press('ArrowRight');
  await sleep(100);
  check('typing or moving on hides it', await p.evaluate(() => document.querySelector('.minibar').hidden));

  // ---- Side panes are cards
  const card = await p.evaluate(() => { const c = document.querySelector('.comments-pane'); c.hidden = false; const s = getComputedStyle(c); const out = [s.borderTopLeftRadius, s.boxShadow !== 'none']; c.hidden = true; return out; });
  check('side panes are rounded cards', card[0] === '10px' && card[1], card.join());

  // ---- A narrower window: the ribbon compacts instead of cutting groups off
  for (const app of ['write', 'sheets', 'slides', 'notes']) {
    const q = await ctx.newPage();
    q.on('pageerror', e => errors.push(app + ': ' + e.message));
    await q.setViewportSize({ width: 1180, height: 800 });
    await q.goto(BASE + app + '.html');
    await q.waitForFunction(() => window.Platform && Platform.ribbonMode);
    await q.evaluate(() => Platform.ribbonMode.set('full'));
    await sleep(500);
    const fit = await q.evaluate(() => { const r = document.querySelector('.ribbon'); return [r.scrollWidth, r.clientWidth, r.className]; });
    check(`${app}: every ribbon group fits at 1180px`, fit[0] <= fit[1] + 1, fit.join(' '));
    await q.setViewportSize({ width: 1900, height: 800 });
    await sleep(300);
    const wide = await q.evaluate(() => document.querySelector('.ribbon').className);
    check(`${app}: and has its words back in a wide window`, !/compact/.test(wide), wide);
    await q.close();
  }

  check('no page errors', !errors.length, errors.slice(0, 3).join(' | '));
  await browser.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
