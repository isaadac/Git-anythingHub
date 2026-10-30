// Export to Video in a browser: the show plays by itself and is recorded from this tab (the test
// browser shares it without asking); the file is checked with ffprobe and its frames with ffmpeg.
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const W = process.env.E2E_WORK;
(async () => {
  const b = await chromium.launch({ headless: !process.env.DISPLAY, args: ['--auto-accept-this-tab-capture', '--use-fake-ui-for-media-stream', '--autoplay-policy=no-user-gesture-required'] });
  const p = await b.newPage({ viewport: { width: 1280, height: 720 } });
  const errors = [];
  p.on('pageerror', e => errors.push(e.message));
  await p.goto(process.env.E2E_BASE + '/src/renderer/slides.html');
  await p.waitForFunction(() => window.AtlasSlides);
  // Two slides: red, then blue (with a fade).
  await p.evaluate(() => {
    AtlasSlides.addSlide('blank');
    const s = AtlasSlides.pres.slides;
    s[0].bg = '#d02020'; s[0].elements = [];
    s[1].bg = '#2040d0'; s[1].elements = []; s[1].transition = 'fade';
    AtlasSlides.goTo(0);
    window.__saved = null;
    Platform.save = async o => { window.__saved = { name: o.name, bytes: Array.from(o.data), ext: o.filters[0].extensions[0] }; return { name: o.name }; };
  });
  await p.evaluate(() => Platform.emit('export-video'));
  await p.waitForSelector('.modal');
  await p.fill('.modal input[type=number]', '2');
  await p.selectOption('.modal select', '720');
  await p.click('.modal .ok');
  await p.waitForFunction(() => window.__saved, null, { timeout: 30000 });
  const saved = await p.evaluate(() => ({ name: __saved.name, ext: __saved.ext, size: __saved.bytes.length }));
  check('the video is saved as MP4', saved.ext === 'mp4' && /^Presentation\d*\.mp4$/.test(saved.name) && saved.size > 5000, JSON.stringify(saved));
  const file = path.join(W, 'export.mp4');
  fs.writeFileSync(file, Buffer.from(await p.evaluate(() => __saved.bytes)));
  check('the show is over after the last slide', await p.evaluate(() => AtlasSlides.showIndex === null && !document.body.classList.contains('exporting-video')));
  const info = JSON.parse(execFileSync('ffprobe', ['-v', 'error', '-show_streams', '-show_format', '-of', 'json', file]).toString());
  const v = info.streams.find(s => s.codec_type === 'video');
  check('H.264 video, 1280 × 720', v && v.codec_name === 'h264' && v.width === 1280 && v.height === 720, v && `${v.codec_name} ${v.width}x${v.height}`);
  // The length from the frames' times (a recording may not say its duration up front).
  const times = execFileSync('ffprobe', ['-v', 'error', '-select_streams', 'v:0', '-show_entries', 'frame=pts_time', '-of', 'csv=p=0', file]).toString().trim().split('\n').map(Number).filter(isFinite);
  const length = Math.max(...times) - Math.min(...times);
  check('about four seconds long (two slides × 2 s)', length > 3 && length < 6, length.toFixed(2));
  // The first frame is the red slide, the last the blue one.
  const pixel = args => [...execFileSync('ffmpeg', ['-v', 'error', ...args, '-vf', 'scale=1:1', '-frames:v', '1', '-f', 'rawvideo', '-pix_fmt', 'rgb24', '-'])];
  const first = pixel(['-i', file]), last = pixel(['-sseof', '-0.2', '-i', file]);
  check('the first slide is in the video', first[0] > 150 && first[2] < 90, first.join(','));
  check('and the last', last[2] > 150 && last[0] < 90, last.join(','));
  // Esc during the export discards it.
  await p.evaluate(() => { window.__saved = null; Platform.emit('export-video'); });
  await p.waitForSelector('.modal');
  await p.click('.modal .ok');
  await p.waitForFunction(() => AtlasSlides.showIndex === 0);
  await p.waitForTimeout(1200);
  await p.keyboard.press('Escape');
  await p.waitForFunction(() => /cancelled/.test((document.querySelector('.toast') || {}).textContent || ''), null, { timeout: 8000 });
  check('Esc cancels the export, nothing saved', await p.evaluate(() => window.__saved === null));
  check('no page errors', errors.length === 0, errors.join(' | '));
  await b.close();
})().catch(e => { console.log('FAIL crashed :: ' + e.message); process.exit(1); });
