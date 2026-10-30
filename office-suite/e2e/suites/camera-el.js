// Desktop: the camera is for Cameo in Slides only; other windows can't open it (the microphone works in all).
const { _electron: electron } = require('playwright');
const path = require('path');
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
(async () => {
  const app = await electron.launch({ executablePath: path.join(process.cwd(), 'node_modules/electron/dist/electron'), args: ['--no-sandbox', '--use-fake-device-for-media-stream', `--user-data-dir=${process.env.E2E_WORK}/ud-cam`, '.'], env: { ...process.env, NODE_OPTIONS: '' } });
  const home = await app.firstWindow();
  await sleep(1000);
  const open = async page => {
    await home.evaluate(p => window.atlas.newWindow(p), page);
    let w; for (let i = 0; i < 40 && !w; i++) { await sleep(250); w = app.windows().find(x => x.url().includes(page + '.html')); }
    await w.waitForLoadState();
    return w;
  };
  const tryMedia = (w, c) => w.evaluate(async c => { try { const s = await navigator.mediaDevices.getUserMedia(c); s.getTracks().forEach(t => t.stop()); return 'ok'; } catch (e) { return e.name; } }, c);
  const slides = await open('slides'), write = await open('write');
  check('desktop: Slides can use the camera (Cameo)', (await tryMedia(slides, { video: true })) === 'ok');
  const w = await tryMedia(write, { video: true });
  check('desktop: other apps cannot', w === 'NotAllowedError', w);
  check('desktop: the microphone works in other apps', (await tryMedia(write, { audio: true })) === 'ok');
  for (const x of app.windows()) await x.evaluate(() => { window.__atlasDirty = false; });
  await require('../close-app')(app);
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.error('CRASH', e.message); process.exit(1); });
