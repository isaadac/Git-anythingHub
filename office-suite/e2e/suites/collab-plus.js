// Live editing, beyond the basics: two people typing in the same paragraph at once both keep their
// words; the Activity list says who changed what; @mentions in comments reach the person
// mentioned; following someone who presents in Slides; and who made each saved version.
const { chromium } = require('playwright');
const path = require('path');
const ROOT = process.cwd();
const { createCollabServer } = require(path.join(ROOT, 'src/main/collab-server.js'));
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
const BASE = process.env.E2E_BASE + '/src/renderer/';

(async () => {
  const srv = createCollabServer();
  await new Promise(res => srv.listen(0, '127.0.0.1', res));
  const WS = `ws://127.0.0.1:${srv.address().port}/`;
  const browser = await chromium.launch();
  const errors = [];
  async function win(app, name) {
    const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 } });
    await ctx.addInitScript(n => { localStorage.setItem('atlas-user-name', n); localStorage.setItem('atlas-author', n); }, name);
    const p = await ctx.newPage();
    p.on('pageerror', e => errors.push(name + ': ' + e.message));
    await p.goto(BASE + app + '.html');
    await p.waitForFunction(() => window.Collab && document.querySelector('.live-btn'));
    return p;
  }
  async function host(p) {
    await p.click('.live-btn');
    await p.click('.modal .actions button:has-text("Start on a server")');
    await p.fill('.modal input', WS);
    await p.click('.modal .actions .primary');
    await p.waitForSelector('.modal p:has-text("room=")', { timeout: 10000 });
    const link = (await p.textContent('.modal p')).match(/ws:\/\/\S+/)[0];
    await p.click('.modal .actions button:has-text("Close")');
    return link;
  }
  async function join(p, link) {
    await p.click('.live-btn');
    await p.click('.modal .actions button:has-text("Join")');
    await p.fill('.modal input', link);
    await p.click('.modal .actions .primary');
    await p.waitForFunction(() => document.querySelector('.live-btn').textContent.startsWith('Live ·'), null, { timeout: 15000 });
  }
  const settle = async (pages, f) => { let all = []; for (let i = 0; i < 40; i++) { await sleep(200); all = await Promise.all(pages.map(f)); if (all.every(x => x === all[0])) break; } return all; };

  // ---- Two people typing in one paragraph at once
  const a = await win('write', 'Ada Lovelace'), b = await win('write', 'Bea Smith');
  await a.click('#doc');
  await a.keyboard.type('The meeting is on Monday.');
  await a.keyboard.press('Enter');
  await a.keyboard.type('Second paragraph.');
  const link = await host(a);
  await join(b, link);
  await sleep(500);
  const caretAt = (p, where) => p.evaluate(where => {
    const t = document.querySelector('#doc p').firstChild, r = document.createRange();
    r.setStart(t, where === 'end' ? t.length : 0); r.collapse(true);
    const s = getSelection(); s.removeAllRanges(); s.addRange(r); document.getElementById('doc').focus();
  }, where);
  await caretAt(a, 'end'); await caretAt(b, 'start');
  // Both type at the same time, in the same paragraph.
  await Promise.all([a.keyboard.type(' Bring the slides.', { delay: 25 }), b.keyboard.type('Reminder: ', { delay: 25 })]);
  const first = p => p.evaluate(() => document.querySelector('#doc p').textContent.replace(/ /g, ' '));
  const both = await settle([a, b], first);
  check('two people typing in one paragraph at once both keep their words', both[0] === both[1] && /^Reminder: The meeting is on Monday\. Bring the slides\.$/.test(both[0]), JSON.stringify(both));
  check('…and nothing else in the document changes', (await a.evaluate(() => document.querySelectorAll('#doc p')[1].textContent)) === 'Second paragraph.');

  // ---- Activity: who changed what
  await a.click('.live-btn');
  await a.click('.modal .actions button:has-text("Activity")');
  await a.waitForSelector('.activity-pane .act');
  const acts = await a.$$eval('.activity-pane .act', rows => rows.map(r => r.textContent));
  check('the Activity list says who changed what', acts.some(t => /Bea Smith/.test(t) && /Reminder: The meeting/.test(t)) && acts.some(t => /Ada Lovelace \(you\)/.test(t)), acts.join(' | '));
  await a.evaluate(() => window.scrollTo(0, 0));
  await a.click('.activity-pane .act:has-text("Bea Smith")');
  check('clicking an entry shows where', await a.evaluate(() => document.querySelector('#doc p').classList.contains('flash')));
  await a.click('.activity-pane .x');

  // ---- @mentions
  await a.evaluate(() => { const p = document.querySelectorAll('#doc p')[1], r = document.createRange(); r.selectNodeContents(p); const s = getSelection(); s.removeAllRanges(); s.addRange(r); });
  await a.click('.ribbon-tabs button[data-tab="review"]');
  await a.click('#addComment');
  await a.waitForSelector('.modal textarea');
  await a.type('.modal textarea', 'Please check this @Be');
  check('typing @ suggests the people in the session', (await a.$$eval('.mention-list .opt', os => os.map(o => o.textContent))).join() === 'Bea Smith');
  await a.keyboard.press('Enter');
  check('…and Enter picks one', (await a.inputValue('.modal textarea')) === 'Please check this @Bea Smith ');
  await a.click('.modal .actions .primary');
  await a.waitForSelector('.comments-pane .card .mention');
  check('the mention stands out in the comment', (await a.textContent('.comments-pane .card .mention')) === '@Bea Smith');
  await b.waitForFunction(() => /Ada Lovelace mentioned you/.test((document.querySelector('.toast') || {}).textContent || ''), null, { timeout: 8000 }).catch(() => {});
  check('the person mentioned is told, and sees the comment marked', /Ada Lovelace mentioned you: “Please check this @Bea Smith”/.test(await b.textContent('.toast')) && await b.evaluate(() => !document.querySelector('.comments-pane').hidden && !!document.querySelector('.comments-pane .card.mentions-me')));

  // ---- Who made the saved version
  const by = await a.evaluate(() => Platform.contributors());
  check('a version saved now is by Ada, with Bea', by.join() === 'Ada Lovelace,Bea Smith', by.join());
  check('…and the next one only by whoever changed things since', (await a.evaluate(() => Platform.contributors())).join() === 'Ada Lovelace');
  for (const p of [a, b]) await p.evaluate(() => { window.__atlasDirty = false; });

  // ---- Slides: follow the presenter
  const s1 = await win('slides', 'Ada Lovelace'), s2 = await win('slides', 'Bea Smith');
  await s1.evaluate(() => { AtlasSlides.addSlide('content'); AtlasSlides.addSlide('content'); AtlasSlides.goTo(0); });
  const slink = await host(s1);
  await join(s2, slink);
  await sleep(500);
  await s1.evaluate(() => AtlasSlides.present(1));
  await sleep(300);
  await s2.click('.live-people .live-avatar');
  await s2.waitForFunction(() => document.querySelector('.follow-banner') && !document.querySelector('.follow-banner').hidden, null, { timeout: 5000 }).catch(() => {});
  const f1 = await s2.evaluate(() => ({ cur: document.getElementById('stSlide').textContent, banner: (document.querySelector('.follow-banner .who') || {}).textContent, ring: document.querySelector('.live-avatar').classList.contains('following') }));
  check('following someone who presents shows their slide, and offers to watch', f1.cur === 'Slide 2 of 3' && f1.banner === 'Ada Lovelace is presenting' && f1.ring, JSON.stringify(f1));
  await s2.click('.follow-banner .watch');
  await sleep(300);
  check('Watch starts the show on the same slide', (await s2.evaluate(() => AtlasSlides.showIndex)) === 1);
  await s1.keyboard.press('ArrowRight');
  await s2.waitForFunction(() => AtlasSlides.showIndex === 2, null, { timeout: 4000 }).catch(() => {});
  check('…and moves on when the presenter does', (await s2.evaluate(() => AtlasSlides.showIndex)) === 2, String(await s2.evaluate(() => AtlasSlides.showIndex)));
  await s2.keyboard.press('Escape');
  await s2.click('.live-people .live-avatar');
  check('clicking again stops following', !(await s2.evaluate(() => document.querySelector('.live-avatar').classList.contains('following'))));

  for (const p of [s1, s2]) await p.evaluate(() => { window.__atlasDirty = false; });
  check('no page errors', !errors.length, errors.slice(0, 3).join(' | '));
  await browser.close();
  await new Promise(res => { srv.close(res); for (const rm of srv.rooms.values()) for (const c of rm.clients) c.socket.destroy(); });
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
