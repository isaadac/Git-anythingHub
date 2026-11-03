// Live sessions: links that can edit, comment or only view (the server enforces them), the
// session's discussion with replies and @mentions, unread counts and notifications.
const { chromium } = require('playwright');
const path = require('path');
const { createCollabServer } = require(path.join(process.cwd(), 'src/main/collab-server.js'));
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const BASE = process.env.E2E_BASE + '/src/renderer/';

(async () => {
  const srv = createCollabServer({ limits: { idleMs: 3000 } });
  await new Promise(r => srv.listen(0, '127.0.0.1', r));
  const WS = `ws://127.0.0.1:${srv.address().port}/`;
  const browser = await chromium.launch();
  const errors = [];
  async function open(app, name) {
    const ctx = await browser.newContext();
    await ctx.addInitScript(n => { localStorage.setItem('atlas-user-name', n); }, name);
    await ctx.grantPermissions(['clipboard-read', 'clipboard-write', 'notifications']);
    const p = await ctx.newPage();
    p.on('pageerror', e => errors.push(name + ': ' + e.message));
    await p.goto(BASE + app + '.html');
    await p.waitForFunction(() => window.Collab && document.querySelector('.live-btn'));
    // Notifications are recorded (the page counts as in the background when asked).
    await p.evaluate(() => {
      window.__notes = [];
      window.Notification = class { constructor(t, o) { window.__notes.push({ t, body: o.body }); } static get permission() { return 'granted'; } static requestPermission() { return Promise.resolve('granted'); } };
    });
    return p;
  }
  const until = (p, fn, arg, ms = 8000) => p.waitForFunction(fn, arg, { timeout: ms }).then(() => true, () => false);
  async function host(p) {
    await p.click('.live-btn');
    await p.click('.modal .actions button:has-text("Start on a server")');
    await p.fill('.modal input', WS);
    await p.click('.modal .actions .primary');
    await p.waitForSelector('.modal p:has-text("room=")', { timeout: 10000 });
    const text = await p.textContent('.modal p');
    const buttons = await p.$$eval('.modal .actions button', bs => bs.map(b => b.textContent));
    await p.click('.modal .actions button:has-text("Close")');
    const link = role => (text.match(new RegExp(`Can ${role}: (ws://\\S+)`)) || [])[1];
    return { edit: link('edit'), comment: link('comment'), view: link('view'), text, buttons };
  }
  async function join(p, link) {
    await p.click('.live-btn');
    await p.click('.modal .actions button:has-text("Join")');
    await p.fill('.modal input', link);
    await p.click('.modal .actions .primary');
    await p.waitForFunction(() => document.querySelector('.live-btn').textContent.startsWith('Live ·'), null, { timeout: 10000 });
  }

  // ---- Write: three links
  const ed = await open('write', 'Edith'), co = await open('write', 'Colin'), vi = await open('write', 'Vera');
  await ed.click('#doc');
  await ed.keyboard.type('Draft plan');
  const links = await host(ed);
  check('the host gets three links: edit, comment and view', !!links.edit && !!links.comment && !!links.view && new Set([links.edit, links.comment, links.view]).size === 3, links.text);
  check('…with a button to copy each', ['Copy “Can edit” Link', 'Copy “Can comment” Link', 'Copy “Can view” Link', 'Discussion'].every(b => links.buttons.includes(b)), links.buttons.join(' | '));
  await join(co, links.comment);
  await join(vi, links.view);
  check('people joining with any link get the document', await until(co, () => /Draft plan/.test(document.getElementById('doc').textContent)) && await until(vi, () => /Draft plan/.test(document.getElementById('doc').textContent)));
  check('a comment or view link can’t edit the page', (await co.getAttribute('#doc', 'contenteditable')) === 'false' && (await vi.getAttribute('#doc', 'contenteditable')) === 'false' && (await ed.getAttribute('#doc', 'contenteditable')) === 'true');
  check('…the ribbon is locked for them (not for the editor)', await co.evaluate(() => document.body.classList.contains('live-readonly') && getComputedStyle(document.querySelector('.ribbon-panel.active .controls')).pointerEvents === 'none') && !(await ed.evaluate(() => document.body.classList.contains('live-readonly'))));
  // Even if a page changes anyway, nothing is sent; and the server refuses edits from those links.
  await vi.evaluate(() => { const d = document.getElementById('doc'); d.innerHTML = '<p>Vandalised</p>'; d.dispatchEvent(new Event('input', { bubbles: true })); });
  await ed.waitForTimeout(800);
  check('a viewer’s changed page isn’t sent to others', /Draft plan/.test(await ed.textContent('#doc')) && !/Vandalised/.test(await co.textContent('#doc')));
  await ed.click('#doc');
  await ed.keyboard.press('End');
  await ed.keyboard.type(' — Final plan');
  check('the editor’s changes reach everyone', await until(co, () => /Final plan/.test(document.getElementById('doc').textContent)) && await until(vi, () => /Final plan/.test(document.getElementById('doc').textContent)));
  const coLinks = await co.evaluate(() => { const b = document.querySelector('.live-btn'); b.click(); return new Promise(r => setTimeout(() => { const t = document.querySelector('.modal p').textContent; document.querySelector('.modal .actions button:last-child').click(); r(t); }, 300)); });
  check('a commenter sees only their own link, and that they can only comment', /Can comment: ws:/.test(coLinks) && !/Can edit/.test(coLinks) && /can comment on this document, not change it/.test(coLinks), coLinks);

  // ---- Discussion
  const openTalk = p => p.evaluate(() => { document.querySelector('.live-btn').click(); return new Promise(r => setTimeout(() => { [...document.querySelectorAll('.modal .actions button')].find(b => b.textContent === 'Discussion').click(); r(); }, 200)); });
  await openTalk(co);
  check('the discussion opens with a box to comment', await co.isVisible('.talk-pane textarea'));
  // Edith is in the background: a mention notifies her.
  await ed.evaluate(() => { Object.defineProperty(document, 'hidden', { configurable: true, get: () => true }); document.hasFocus = () => false; });
  await co.fill('.talk-pane textarea', 'Is section 2 ready, @Edith?');
  await co.press('.talk-pane textarea', 'Enter');
  check('a comment reaches everyone', await until(ed, () => window.Collab && document.querySelector('.live-badge') && !document.querySelector('.live-badge').hidden) && await until(vi, () => document.querySelector('.live-badge') && document.querySelector('.live-badge').dataset.n === '1'));
  check('…mentioning someone sends them a notification', await until(ed, () => window.__notes.some(n => n.t === 'Colin mentioned you' && /section 2/.test(n.body))));
  check('…others aren’t notified', (await vi.evaluate(() => window.__notes.length)) === 0);
  check('…the unread count shows on the Live button (and to screen readers)', (await ed.evaluate(() => document.querySelector('.live-badge').dataset.n)) === '1' && /1 unread comment/.test(await ed.getAttribute('.live-btn', 'aria-label')));
  await ed.evaluate(() => { delete document.hidden; document.hasFocus = () => true; });
  await openTalk(ed);
  check('opening the discussion clears the count and shows the mention', (await ed.isHidden('.live-badge')) && (await ed.locator('.talk-pane .talk.mentions-me .at').textContent()) === '@Edith');
  await ed.click('.talk-pane .talk .reply-btn');
  await co.evaluate(() => { Object.defineProperty(document, 'hidden', { configurable: true, get: () => true }); document.hasFocus = () => false; });
  await ed.fill('.talk-pane textarea', 'Yes, done');
  await ed.press('.talk-pane textarea', 'Enter');
  check('a reply is shown under the comment it answers', await until(co, () => { const r = document.querySelector('.talk-pane .talk.reply'); return r && /Colin: Is section 2/.test(r.textContent) && /Yes, done/.test(r.textContent); }));
  check('…and notifies the person it replies to', await until(co, () => window.__notes.some(n => n.t === 'Edith replied to you')));
  await openTalk(vi);
  check('a viewer reads the discussion but has no box to comment', (await vi.locator('.talk-pane .talk').count()) === 2 && await vi.isHidden('.talk-pane .compose'));
  check('comments are text, never HTML', await (async () => { await co.evaluate(() => { delete document.hidden; }); await co.fill('.talk-pane textarea', '<img src=x onerror="window.__xss=1">'); await co.press('.talk-pane textarea', 'Enter'); return until(ed, () => [...document.querySelectorAll('.talk-pane .talk .text')].some(t => t.textContent.includes('<img src=x')) && !document.querySelector('.talk-pane img') && !window.__xss); })());
  // Someone joining later sees the discussion so far.
  const late = await open('write', 'Lou');
  await join(late, links.view);
  await openTalk(late);
  check('someone joining later sees the discussion so far', (await late.locator('.talk-pane .talk').count()) === 3);
  // Leaving gives the page back.
  await co.evaluate(() => { document.querySelector('.live-btn').click(); return new Promise(r => setTimeout(() => { [...document.querySelectorAll('.modal .actions button')].find(b => b.textContent === 'Leave Session').click(); r(); }, 200)); });
  check('leaving a comment-only session makes the page editable again', await until(co, () => document.getElementById('doc').getAttribute('contenteditable') === 'true' && !document.body.classList.contains('live-readonly')));

  // ---- Sheets: a view link can't change cells
  const se = await open('sheets', 'Sam'), sv = await open('sheets', 'Val');
  await se.evaluate(() => AtlasSheets.setCell('A1', 'total'));
  const sl = await host(se);
  await join(sv, sl.view);
  check('a viewer gets the workbook', await until(sv, () => AtlasSheets.valueOf('A1') === 'total'));
  await sv.click('.grid tbody tr[data-r="1"] td[data-c="0"]');
  await sv.keyboard.type('99');
  await sv.keyboard.press('Enter');
  check('typing in a cell does nothing for a viewer', (await sv.evaluate(() => AtlasSheets.valueOf('A2'))) === '' || (await sv.evaluate(() => AtlasSheets.valueOf('A2'))) === null || (await sv.evaluate(() => AtlasSheets.valueOf('A2'))) === undefined);
  await se.evaluate(() => AtlasSheets.setCell('A2', 5));
  check('…while the editor’s changes still arrive', await until(sv, () => AtlasSheets.valueOf('A2') === 5));
  await se.waitForTimeout(300);
  check('…and the editor’s cell is untouched', (await se.evaluate(() => AtlasSheets.valueOf('A2'))) === 5);

  check('no page errors', !errors.length, errors.join(' | '));
  await browser.close();
  srv.close();
})().catch(e => { console.error('CRASH', e.stack); process.exit(1); });
