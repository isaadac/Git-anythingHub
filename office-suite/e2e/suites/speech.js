// On-device speech: dictation in Write with Whisper running in the page, the model downloaded
// once (with consent) and kept; nothing said leaves the computer. Spoken audio comes from espeak
// when it is installed (else a recording is skipped) and is fed in place of the microphone.
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const speechModel = require('../speech-model');
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));

/** "Good morning everyone. Please send the report." as 16 kHz samples (espeak), or null. */
function speech() {
  const wav = path.join(process.env.E2E_WORK, 'say.wav');
  try { execFileSync('espeak-ng', ['-v', 'en-us', '-s', '140', '-w', wav, 'Good morning everyone.']); } catch { return null; }
  const b = fs.readFileSync(wav), rate = b.readUInt32LE(24);
  const pcm = new Int16Array(b.buffer, b.byteOffset + 44, (b.length - 44) >> 1), out = [];
  for (let t = 0; t < pcm.length / rate * 16000; t++) { const x = t * rate / 16000, i = Math.floor(x), f = x - i; out.push(((pcm[i] || 0) * (1 - f) + (pcm[i + 1] || 0) * f) / 32768); }
  // A second of quiet before and after, as when someone starts and stops talking.
  return [...new Array(16000).fill(0), ...out, ...new Array(16000).fill(0)];
}

(async () => {
  const model = speechModel(), audio = speech();
  if (!model || !audio) { console.log('PASS on-device speech skipped (no model copy or no espeak here)'); return; }
  const browser = await chromium.launch();
  const ctx = await browser.newContext();
  const asked = [];
  await ctx.route('https://huggingface.co/**', r => {
    const u = new URL(r.request().url());
    asked.push(u.pathname);
    const m = u.pathname.match(/^\/Xenova\/whisper-tiny\/resolve\/main\/(.+)$/), f = m && path.join(model, m[1]);
    if (f && fs.existsSync(f)) r.fulfill({ body: fs.readFileSync(f), headers: { 'content-type': 'application/octet-stream', 'access-control-allow-origin': '*' } });
    else r.fulfill({ status: 404, body: '' });
  });
  const others = [];
  ctx.on('request', r => { if (!/^http:\/\/127\.0\.0\.1|^https:\/\/huggingface\.co\/|^data:|^blob:/.test(r.url())) others.push(r.url()); });
  const p = await ctx.newPage();
  const errors = [];
  p.on('pageerror', e => errors.push(e.message));
  await p.goto(process.env.E2E_BASE + '/src/renderer/write.html');
  await p.waitForFunction(() => window.AtlasWrite && window.Speech);
  // No browser recognizer here (as in the desktop app): the first use asks.
  await p.evaluate(a => { delete window.SpeechRecognition; delete window.webkitSpeechRecognition; Speech.testAudio = new Float32Array(a); }, audio);
  await p.click('#doc');
  await p.evaluate(() => document.getElementById('dictate').click());
  await p.waitForSelector('.modal');
  const ask = await p.textContent('.modal');
  check('the first use explains on-device speech and asks', /On this device/.test(ask) && /Hugging Face once/.test(ask), ask.slice(0, 200));
  await p.click('.modal .actions button:has-text("On this device")');
  await p.waitForFunction(() => /good morning/i.test(document.getElementById('doc').textContent), null, { timeout: 120000 }).catch(() => {});
  const text = await p.textContent('#doc');
  check('dictation types what was said, recognized on this device', /good morning/i.test(text), text);
  check('the model came from Hugging Face, and nothing else was contacted', asked.some(a => /encoder_model_quantized\.onnx$/.test(a)) && !others.length, others.slice(0, 3).join(' '));
  check('the choice and the model are kept', await p.evaluate(async () => localStorage.getItem('atlas-speech') === 'device' && await Speech.hasModel()));
  // Offline next time: the model is on the device.
  await p.evaluate(() => document.getElementById('dictate').click()); // stop
  await p.reload();
  await p.waitForFunction(() => window.Speech);
  asked.length = 0;
  // The internet out of reach (Hugging Face refuses), the app's own files still there.
  await ctx.unroute('https://huggingface.co/**');
  await ctx.route('https://huggingface.co/**', r => { asked.push(r.request().url()); r.abort('internetdisconnected'); });
  await p.evaluate(a => { delete window.SpeechRecognition; delete window.webkitSpeechRecognition; Speech.testAudio = new Float32Array(a); document.getElementById('doc').innerHTML = '<p><br></p>'; }, audio);
  await p.click('#doc');
  await p.evaluate(() => document.getElementById('dictate').click());
  await p.waitForFunction(() => /good morning/i.test(document.getElementById('doc').textContent), null, { timeout: 120000 }).catch(() => {});
  check('without the internet, it still works from the kept model', /good morning/i.test(await p.textContent('#doc')) && !asked.length, asked.join(' '));
  await p.evaluate(() => { window.__atlasDirty = false; });
  check('no page errors', !errors.length, errors.slice(0, 3).join(' | '));
  await browser.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
