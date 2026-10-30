// The phone remote in the desktop app, collected by "This computer": the phone opens the link on
// the local network, and its taps move the show in the app.
const { _electron: electron, chromium } = require('playwright');
const path = require('path');
const ROOT = process.cwd();
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
(async () => {
  const ud = path.join(process.env.E2E_WORK, 'ud-remote-' + Date.now());
  const app = await electron.launch({ executablePath: path.join(ROOT, 'node_modules/electron/dist/electron'), args: ['--no-sandbox', `--user-data-dir=${ud}`, '.'], env: { ...process.env, NODE_OPTIONS: '' } });
  const home = await app.firstWindow(); await sleep(1200);
  await home.click('.app-card[data-app="slides"]');
  let s; for (let i = 0; i < 40 && !s; i++) { await sleep(250); s = app.windows().find(w => w.url().includes('slides.html')); }
  await s.waitForFunction(() => window.AtlasSlides);
  const errors = [];
  s.on('pageerror', e => errors.push(e.message));
  await s.evaluate(() => { AtlasSlides.addSlide('title'); AtlasSlides.pres.slides[1].notes = 'Remember the demo'; AtlasSlides.goTo(0); });
  await s.click('[data-tab="show"]');
  await s.click('#remoteBtn');
  await s.waitForSelector('.modal button:has-text("This computer")');
  await s.click('.modal button:has-text("This computer")');
  await s.waitForSelector('.share-qr');
  const link = await s.inputValue('.share-qr input');
  check('the remote link is on this computer’s network', /^http:\/\/[\d.]+:\d+\/r\/[\w-]{22}#[\w-]{40,}$/.test(link), link);
  await s.click('.share-qr .primary');
  const b = await chromium.launch();
  const phone = await b.newPage({ viewport: { width: 390, height: 844 } });
  await phone.goto(link.replace(/^http:\/\/[\d.]+:/, 'http://127.0.0.1:'));
  await phone.waitForFunction(() => /hasn’t started/.test(document.getElementById('count').textContent), null, { timeout: 15000 });
  await phone.click('#nextBtn');
  await s.waitForFunction(() => AtlasSlides.showIndex === 0, null, { timeout: 15000 });
  await phone.click('#nextBtn');
  await s.waitForFunction(() => AtlasSlides.showIndex === 1, null, { timeout: 15000 });
  await phone.waitForFunction(() => document.getElementById('notes').textContent === 'Remember the demo', null, { timeout: 15000 });
  check('taps on the phone move the show, and the phone shows the notes', true);
  await b.close();
  check('no page errors', errors.length === 0, errors.join(' | '));
  await require('../close-app')(app);
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
