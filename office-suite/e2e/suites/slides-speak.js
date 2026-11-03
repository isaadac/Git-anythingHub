// Slides, presenting live: Cameo (the camera in a slide), live captions, and Rehearse with Coach.
// Chromium's fake camera stands in for a webcam; a scripted recognizer stands in for speech.
const { chromium } = require('playwright');
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));

(async () => {
  const browser = await chromium.launch({ args: ['--use-fake-device-for-media-stream', '--use-fake-ui-for-media-stream'] });
  const ctx = await browser.newContext({ viewport: { width: 1280, height: 800 }, permissions: ['camera', 'microphone'] });
  // A speech recognizer the test speaks through: window.__say(text) is heard as a final phrase.
  await ctx.addInitScript(() => {
    class FakeRecognition {
      constructor() { FakeRecognition.live = this; window.__recognizers = (window.__recognizers || 0) + 1; }
      start() { this.running = true; }
      stop() { this.running = false; if (this.onend) this.onend(); }
    }
    window.SpeechRecognition = FakeRecognition;
    window.__say = (text, final = true) => {
      const r = FakeRecognition.live;
      if (!r || !r.running) return false;
      const res = [{ transcript: text }]; res.isFinal = final;
      r.onresult({ resultIndex: 0, results: [res] });
      return true;
    };
    window.__listening = () => !!(FakeRecognition.live && FakeRecognition.live.running);
  });
  const p = await ctx.newPage();
  const errors = [];
  p.on('pageerror', e => errors.push(e.message));
  await p.goto(process.env.E2E_BASE + '/src/renderer/slides.html');
  await p.waitForFunction(() => window.AtlasSlides && window.SlidesCoach);
  await p.evaluate(() => { localStorage.removeItem('atlas-captions'); AtlasSlides.addSlide('content'); AtlasSlides.addSlide('content'); AtlasSlides.goTo(0); });

  // ---- Cameo
  await p.click('.ribbon-tabs button[data-tab="insert"]');
  await p.click('#insCameo');
  const el = await p.evaluate(() => AtlasSlides.pres.slides[0].elements.find(e => e.type === 'cameo'));
  check('Insert → Cameo adds a round camera placeholder', el && el.shape === 'circle' && el.w === el.h && await p.isVisible('#stage .el.cameo.cameo-circle .cameo-ph'));
  await p.dblclick('#stage .el.cameo');
  await p.waitForSelector('.modal');
  await p.selectOption('.modal select >> nth=0', 'rounded');
  await p.click('.modal .actions .primary');
  check('…double-click changes its shape', await p.evaluate(() => AtlasSlides.pres.slides[0].elements.find(e => e.type === 'cameo').shape === 'rounded') && await p.isVisible('#stage .el.cameo.cameo-rounded'));
  check('…and the stage shows no camera while editing', !(await p.$('#stage .el.cameo video')));
  await p.evaluate(() => AtlasSlides.present(0));
  await p.waitForFunction(() => { const v = document.querySelector('#presentSlide .el.cameo video'); return v && v.srcObject && v.readyState >= 2 && v.videoWidth > 0; }, null, { timeout: 10000 }).catch(() => {});
  const cam = await p.evaluate(() => { const v = document.querySelector('#presentSlide .el.cameo video'); return v && { live: v.srcObject && v.srcObject.getVideoTracks()[0].readyState, w: v.videoWidth, mirrored: getComputedStyle(v).transform !== 'none' }; });
  check('presenting puts the live camera in the cameo, mirrored', cam && cam.live === 'live' && cam.w > 0 && cam.mirrored, JSON.stringify(cam));
  // Saved and opened again.
  const bytes = await p.evaluate(async () => { const b = await AtlasSlides.serialize('pptx'); return Array.from(b); }).catch(() => null);
  await p.keyboard.press('Escape');
  await sleep(200);
  check('ending the show turns the camera off', await p.evaluate(() => !document.querySelector('#presentSlide video') || [...document.querySelectorAll('#presentSlide video')].every(v => !v.srcObject || v.srcObject.getTracks().every(t => t.readyState === 'ended'))));
  if (bytes) {
    await p.evaluate(async b => AtlasSlides.loadFile({ name: 'cameo.pptx', data: new Uint8Array(b) }), bytes);
    check('a cameo survives saving as .pptx', await p.evaluate(() => AtlasSlides.pres.slides[0].elements.some(e => e.type === 'cameo' && e.shape === 'rounded')));
  }

  // ---- Live captions
  await p.evaluate(() => { AtlasSlides.goTo(0); AtlasSlides.present(0); });
  check('captions are off unless asked for', !(await p.evaluate(() => window.__listening())));
  await p.keyboard.press('c');
  await sleep(100);
  await p.evaluate(() => window.__say('Good morning everyone'));
  check('C turns captions on: what is said appears on the slide', (await p.textContent('#captions')) === 'Good morning everyone' && await p.isVisible('#captions'));
  await p.evaluate(() => { window.__say('thanks for', false); });
  check('…words still being heard show as they come', (await p.textContent('#captions')) === 'Good morning everyone thanks for');
  await p.evaluate(() => { window.__say('thanks for coming'); window.__say('let us begin'); });
  check('…and only the last two lines are kept', (await p.textContent('#captions')) === 'thanks for coming let us begin');
  await p.keyboard.press('Escape');
  check('ending the show stops listening', !(await p.evaluate(() => window.__listening())));
  await p.evaluate(() => AtlasSlides.present(0));
  check('captions stay on for the next show', await p.evaluate(() => window.__listening()));
  await p.keyboard.press('c');
  check('…until turned off', !(await p.evaluate(() => window.__listening())));
  await p.keyboard.press('Escape');

  // ---- Rehearse with Coach
  await p.evaluate(() => {
    const s = AtlasSlides.pres.slides;
    s[1].elements.push({ id: 'coachtext', type: 'text', x: 60, y: 160, w: 800, h: 200, size: 24, html: '<p>Revenue grew forty percent in the third quarter thanks to new customers in Europe</p>' });
  });
  await p.click('.ribbon-tabs button[data-tab="show"]');
  await p.click('#rehearse');
  await sleep(100);
  check('Rehearse starts the show and listens', await p.evaluate(() => !document.getElementById('presenter').hidden && window.__listening()) && /Coach is listening/.test(await p.textContent('#coachPill')));
  await p.evaluate(() => window.__say('um hello everyone uh thanks for coming'));
  check('the coach flags a filler word as it is said', /Try a pause instead of “um”/.test(await p.textContent('#coachPill')), await p.textContent('#coachPill'));
  await p.keyboard.press('ArrowRight');
  await p.evaluate(() => window.__say('revenue grew forty percent in the third quarter thanks to new customers in Europe'));
  await p.keyboard.press('ArrowRight');
  await p.evaluate(() => window.__say('basically any questions'));
  await p.keyboard.press('Escape');
  await p.waitForSelector('.coach-report');
  const rep = await p.evaluate(() => ({
    stats: [...document.querySelectorAll('.coach-report .stat')].map(s => [...s.children].map(c => c.textContent).join(' ')),
    tips: [...document.querySelectorAll('.coach-report .coach-tips li')].map(l => l.textContent),
    bars: [...document.querySelectorAll('.coach-report .coach-bar')].map(b => [b.querySelector('.n').textContent, b.classList.contains('reading')])
  }));
  check('ending the run shows the report', rep.stats.length === 4 && /^3 Filler words$/.test(rep.stats[2]) && /^1 Slides read aloud$/.test(rep.stats[3]), JSON.stringify(rep.stats));
  check('…with tips on fillers and reading the slide', rep.tips.some(t => /“um” 1×/.test(t)) && rep.tips.some(t => /read slide 2 word for word/.test(t)), rep.tips.join(' | '));
  check('…and time on each slide, the one read aloud marked', rep.bars.length === 3 && rep.bars[1][1] && !rep.bars[0][1], JSON.stringify(rep.bars));
  check('the coach stops listening when the run ends', !(await p.evaluate(() => window.__listening())));
  await p.click('.coach-report .ok');

  await p.evaluate(() => { window.__atlasDirty = false; });
  check('no page errors', !errors.length, errors.slice(0, 3).join(' | '));
  await browser.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
