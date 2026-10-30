// Desktop: on-device dictation works in the app (where the browser's speech service isn't
// available): Whisper loads from the app's own files, the model from its download, kept.
const { _electron: electron } = require('playwright');
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const speechModel = require('../speech-model');
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
(async () => {
  const model = speechModel();
  const wav = path.join(process.env.E2E_WORK, 'say.wav');
  try { execFileSync('espeak-ng', ['-v', 'en-us', '-s', '140', '-w', wav, 'Good morning everyone.']); } catch { /* no espeak */ }
  if (!model || !fs.existsSync(wav)) { console.log('PASS desktop on-device speech skipped (no model copy or no espeak here)'); return; }
  const b = fs.readFileSync(wav), rate = b.readUInt32LE(24), pcm = new Int16Array(b.buffer, b.byteOffset + 44, (b.length - 44) >> 1), audio = new Array(16000).fill(0);
  for (let t = 0; t < pcm.length / rate * 16000; t++) { const x = t * rate / 16000, i = Math.floor(x), f = x - i; audio.push(((pcm[i] || 0) * (1 - f) + (pcm[i + 1] || 0) * f) / 32768); }
  audio.push(...new Array(16000).fill(0));
  const app = await electron.launch({ executablePath: path.join(process.cwd(), 'node_modules/electron/dist/electron'), args: ['--no-sandbox', `--user-data-dir=${process.env.E2E_WORK}/ud-speech`, '.'], env: { ...process.env, NODE_OPTIONS: '' } });
  await app.context().route('https://huggingface.co/**', r => {
    const m = new URL(r.request().url()).pathname.match(/^\/Xenova\/whisper-tiny\/resolve\/main\/(.+)$/), f = m && path.join(model, m[1]);
    if (f && fs.existsSync(f)) r.fulfill({ body: fs.readFileSync(f), headers: { 'content-type': 'application/octet-stream', 'access-control-allow-origin': '*' } }); else r.fulfill({ status: 404, body: '' });
  });
  const home = await app.firstWindow();
  await sleep(800);
  await home.evaluate(() => window.atlas.newWindow('write'));
  let w; for (let i = 0; i < 40 && !w; i++) { await sleep(250); w = app.windows().find(x => x.url().includes('write.html')); }
  await w.waitForFunction(() => window.AtlasWrite && window.Speech);
  await w.evaluate(a => { localStorage.setItem('atlas-speech', 'device'); Speech.testAudio = new Float32Array(a); }, audio);
  await w.click('#doc');
  await w.evaluate(() => document.getElementById('dictate').click());
  await w.waitForFunction(() => /good morning/i.test(document.getElementById('doc').textContent), null, { timeout: 120000 }).catch(() => {});
  check('desktop: dictation works on this device', /good morning/i.test(await w.textContent('#doc')), await w.textContent('#doc'));
  check('desktop: the model is kept for next time', await w.evaluate(() => Speech.hasModel()));
  for (const x of app.windows()) await x.evaluate(() => { window.__atlasDirty = false; });
  await require('../close-app')(app);
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.error('CRASH', e.message); process.exit(1); });
