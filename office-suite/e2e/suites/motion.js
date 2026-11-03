// Motion: dialogs, panes and ribbon panels animate in; the Sheets cursor glides on short moves;
// all of it stops with File → Office Theme → Reduce motion, or when the system asks for less.
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
  await p.goto(BASE + 'sheets.html');
  await p.waitForFunction(() => window.Platform && window.AtlasSheets && window.Theme);
  const anim = sel => p.evaluate(sel => { const e = document.querySelector(sel); if (!e) return 'missing'; const s = getComputedStyle(e); return s.animationName + ' ' + s.animationDuration; }, sel);

  // ---- Things that appear show where they come from
  await p.click('.ribbon-tabs button[data-tab]:nth-of-type(4)');
  const panel = await anim('.ribbon-panel.active');
  check('a ribbon panel slides in from the side of the new tab', /^panel-in-right 0\.16s/.test(panel), panel);
  await p.click('.ribbon-tabs button[data-tab]:nth-of-type(2)');
  check('…and from the other side going back', /^panel-in 0\.16s/.test(await anim('.ribbon-panel.active')));
  check('the tab marker grows under the new tab', /marker-grow/.test(await p.evaluate(() => getComputedStyle(document.querySelector('.ribbon-tabs button.active'), '::after').animationName)));
  await p.evaluate(() => { UI.choose('Title', 'Message', ['OK']); });
  const dlg = [await anim('.modal-backdrop'), await anim('.modal-backdrop > .modal')];
  check('a dialog fades its backdrop and rises in', /^fade-in/.test(dlg[0]) && /^rise-in 0\.24s/.test(dlg[1]), dlg.join(' | '));
  await p.keyboard.press('Escape');

  // ---- The cell cursor glides between neighbours but jumps far
  await p.click('#grid tr[data-r="2"] td[data-c="2"]');
  await p.keyboard.press('ArrowRight');
  const glide = await p.evaluate(() => [document.getElementById('selBox').classList.contains('glide'), getComputedStyle(document.getElementById('selBox')).transitionDuration]);
  check('arrow keys glide the cell cursor', glide[0] && /0\.09s/.test(glide[1]), glide.join());
  await sleep(200);
  await p.fill('#nameBox', 'H400');
  await p.press('#nameBox', 'Enter');
  check('a far jump does not', !(await p.evaluate(() => document.getElementById('selBox').classList.contains('glide'))));
  await p.fill('#nameBox', 'C3');
  await p.press('#nameBox', 'Enter');
  await sleep(200);
  await p.click('#grid tr[data-r="2"] td[data-c="2"]');
  await p.keyboard.press('Shift+ArrowRight');
  check('nor does growing the selection', !(await p.evaluate(() => document.getElementById('selBox').classList.contains('glide'))));

  // ---- Reduce motion: the setting
  await p.click('.ribbon-tabs .file-tab');
  await p.click('.backstage nav button[data-cmd="theme"]');
  const box = '.modal .check-row input';
  check('Office Theme has a Reduce motion switch, off by default', await p.$eval(box, b => !b.checked));
  await p.check(box);
  check('turning it on marks the page', (await p.evaluate(() => [document.documentElement.dataset.motion, Theme.calm].join())) === 'reduce,true');
  await p.keyboard.press('Escape');
  await p.click('.ribbon-tabs button[data-tab]:nth-of-type(3)');
  check('…and stops animations', /0s$/.test(await anim('.ribbon-panel.active')), await anim('.ribbon-panel.active'));
  await p.click('#grid tr[data-r="5"] td[data-c="2"]');
  await p.keyboard.press('ArrowDown');
  check('…and the cursor gliding', !(await p.evaluate(() => document.getElementById('selBox').classList.contains('glide'))));
  const other = await ctx.newPage();
  await other.goto(BASE + 'write.html');
  await other.waitForFunction(() => window.Theme);
  check('other windows follow (remembered)', (await other.evaluate(() => document.documentElement.dataset.motion)) === 'reduce');
  await other.evaluate(() => Theme.setMotion(false));
  await sleep(200);
  check('turning it off reaches the open windows', (await p.evaluate(() => document.documentElement.dataset.motion || '')) === '');
  await other.close();

  // ---- …and the system's setting
  const calm = await browser.newContext({ reducedMotion: 'reduce' });
  const q = await calm.newPage();
  await q.goto(BASE + 'write.html');
  await q.waitForFunction(() => window.Theme && window.UI);
  await q.evaluate(() => { UI.choose('Title', 'Message', ['OK']); });
  check('a system that asks for less motion gets none', await q.evaluate(() => Theme.calm) && /0s$/.test(await q.evaluate(() => getComputedStyle(document.querySelector('.modal')).animationDuration)));
  await calm.close();

  check('no page errors', !errors.length, errors.slice(0, 3).join(' | '));
  await browser.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
