// Export to Video in the desktop app: the window records its own frames (no screen capture and
// no prompt), and a clip's sound on a slide is mixed into the video.
const { _electron: electron } = require('playwright');
const path = require('path'), fs = require('fs');
const { execFileSync } = require('child_process');
const ROOT = process.cwd();
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
// A second of a 440 Hz tone, as a WAV data URL.
function tone() {
  const rate = 16000, n = rate, buf = Buffer.alloc(44 + n * 2);
  buf.write('RIFF', 0); buf.writeUInt32LE(36 + n * 2, 4); buf.write('WAVEfmt ', 8); buf.writeUInt32LE(16, 16); buf.writeUInt16LE(1, 20); buf.writeUInt16LE(1, 22);
  buf.writeUInt32LE(rate, 24); buf.writeUInt32LE(rate * 2, 28); buf.writeUInt16LE(2, 32); buf.writeUInt16LE(16, 34); buf.write('data', 36); buf.writeUInt32LE(n * 2, 40);
  for (let i = 0; i < n; i++) buf.writeInt16LE(Math.round(Math.sin(2 * Math.PI * 440 * i / rate) * 12000), 44 + i * 2);
  return 'data:audio/wav;base64,' + buf.toString('base64');
}
(async () => {
  const ud = path.join(process.env.E2E_WORK, 'ud-video-' + Date.now());
  const app = await electron.launch({ executablePath: path.join(ROOT, 'node_modules/electron/dist/electron'), args: ['--no-sandbox', `--user-data-dir=${ud}`, '--autoplay-policy=no-user-gesture-required', '.'], env: { ...process.env, NODE_OPTIONS: '' } });
  const home = await app.firstWindow(); await sleep(1200);
  await home.click('.app-card[data-app="slides"]');
  let s; for (let i = 0; i < 40 && !s; i++) { await sleep(250); s = app.windows().find(w => w.url().includes('slides.html')); }
  await s.waitForFunction(() => window.AtlasSlides);
  const errors = [];
  s.on('pageerror', e => errors.push(e.message));
  await s.evaluate(src => {
    AtlasSlides.addSlide('blank');
    const sl = AtlasSlides.pres.slides;
    sl[0].bg = '#20a040'; sl[0].elements = [{ id: 'clip1', type: 'media', kind: 'audio', src, x: 20, y: 20, w: 60, h: 60, autoplay: true, hideIcon: true }];
    sl[1].bg = '#e0c020'; sl[1].elements = [];
    AtlasSlides.goTo(0);
    Platform.save = async o => { window.__saved = { ext: o.filters[0].extensions[0], bytes: Array.from(o.data) }; return { name: o.name }; };
  }, tone());
  await s.evaluate(() => Platform.emit('export-video'));
  await s.waitForSelector('.modal');
  check('the desktop app records without asking to share the screen', true);
  await s.fill('.modal input[type=number]', '2');
  await s.selectOption('.modal select', '720');
  await s.click('.modal .ok');
  await s.waitForFunction(() => window.__saved, null, { timeout: 40000 });
  const ext = await s.evaluate(() => __saved.ext);
  const file = path.join(process.env.E2E_WORK, 'desktop.' + ext);
  fs.writeFileSync(file, Buffer.from(await s.evaluate(() => __saved.bytes)));
  const info = JSON.parse(execFileSync('ffprobe', ['-v', 'error', '-show_streams', '-of', 'json', file]).toString());
  const v = info.streams.find(x => x.codec_type === 'video'), a = info.streams.find(x => x.codec_type === 'audio');
  check('an MP4 video of the slides, 1280 × 720', ext === 'mp4' && v && v.width === 1280 && v.height === 720, `${ext} ${v && v.codec_name} ${v && v.width}x${v && v.height}`);
  let max = -99;
  if (a) { const out = require('child_process').spawnSync('ffmpeg', ['-i', file, '-af', 'volumedetect', '-vn', '-f', 'null', '-']).stderr.toString(); max = +((out.match(/max_volume: (-?[\d.]+) dB/) || [])[1] ?? -99); }
  check('with the clip’s sound', !!a && max > -30, a ? `${a.codec_name}, loudest ${max} dB` : 'no audio stream');
  const pixel = args => [...execFileSync('ffmpeg', ['-v', 'error', ...args, '-vf', 'scale=1:1', '-frames:v', '1', '-f', 'rawvideo', '-pix_fmt', 'rgb24', '-'])];
  const first = pixel(['-ss', '0.5', '-i', file]), last = pixel(['-sseof', '-0.3', '-i', file]);
  check('the first slide (green) and the last (yellow) are in it', first[1] > 120 && first[0] < 90 && last[0] > 170 && last[1] > 150 && last[2] < 90, first.join(',') + ' / ' + last.join(','));
  check('no page errors', errors.length === 0, errors.join(' | '));
  await require('../close-app')(app);
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
