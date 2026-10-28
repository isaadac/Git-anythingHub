const { chromium } = require('playwright');
const JSZip = require('jszip');
const fs = require('fs');
const { execFileSync } = require('child_process');
const SP = process.env.E2E_WORK;
(async () => {
  const b = await chromium.launch(); const p = await b.newPage({ viewport: { width: 1400, height: 900 } });
  const errors = []; p.on('pageerror', e => errors.push(e.message)); p.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  const check = (n, c, x = '') => console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : ''));
  await p.goto(process.env.E2E_BASE + '/src/renderer/slides.html'); await p.waitForTimeout(300);
  const offer = async (file) => p.evaluate(([name, bytes]) => { Platform.open = async () => ({ name, path: null, data: new Uint8Array(bytes) }); }, [file, Array.from(fs.readFileSync(SP + '/files/' + file))]);
  await p.click('[data-tab="insert"]');
  await offer('clip.webm'); await p.click('#insVideo'); await p.waitForTimeout(1500);
  const v = await p.evaluate(() => { const e = AtlasSlides.pres.slides[0].elements.find(x => x.type === 'media'); return e && { kind: e.kind, poster: !!e.poster && e.poster.slice(0, 15), w: e.w, h: e.h, src: e.src.slice(0, 20) }; });
  check('video inserted with a poster frame and its aspect ratio', v && v.kind === 'video' && v.poster === 'data:image/jpeg' && Math.abs(v.w / v.h - 16 / 9) < 0.02 && v.src === 'data:video/webm;base', JSON.stringify(v));
  await p.check('#mediaAuto');
  await offer('tone.wav'); await p.click('#insAudio'); await p.waitForTimeout(300);
  await p.check('#mediaHide'); await p.check('#mediaLoop');
  const a = await p.evaluate(() => AtlasSlides.pres.slides[0].elements.filter(x => x.type === 'media').map(e => `${e.kind}:${!!e.autoplay}:${!!e.loop}:${!!e.hideIcon}`).join());
  check('playback options', a === 'video:true:false:false,audio:false:true:true', a);
  // Show: the video starts by itself; clicking the audio icon plays it without advancing
  await p.evaluate(() => AtlasSlides.present(0)); await p.waitForTimeout(1200);
  const playing = await p.evaluate(() => { const v = document.querySelector('#presentSlide video'); return v && !v.paused && v.currentTime > 0; });
  check('video set to start automatically plays in the show', playing);
  const icon = await p.evaluate(() => getComputedStyle(document.querySelector('#presentSlide .media-audio-icon')).visibility);
  await p.evaluate(() => { const n = [...document.querySelectorAll('#presentSlide .el.media')].pop(); n.dispatchEvent(new MouseEvent('click', { bubbles: true })); });
  await p.waitForTimeout(400);
  const aud = await p.evaluate(() => ({ playing: !document.querySelector('#presentSlide audio').paused, stillShowing: !document.getElementById('presenter').hidden }));
  check('audio icon hidden in the show; clicking the clip plays it and doesn\'t advance', icon === 'hidden' && aud.playing && aud.stillShowing, JSON.stringify({ icon, aud }));
  await p.keyboard.press('Escape');
  // PPTX
  const bytes = await p.evaluate(async () => Array.from(await AtlasSlides.serialize('pptx')));
  fs.writeFileSync(SP + '/files/media.pptx', Buffer.from(bytes));
  const zip = await JSZip.loadAsync(Buffer.from(bytes));
  const sx = await zip.file('ppt/slides/slide1.xml').async('string');
  const parts = Object.keys(zip.files).filter(n => /^ppt\/media\//.test(n));
  check('pptx: video and audio embedded as media', /<a:videoFile r:link=/.test(sx) && parts.some(n => /\.webm$/.test(n)) && parts.some(n => /\.wav$/.test(n)), parts.join());
  const py = execFileSync('python3', ['-I', '-c', 'import sys\nfrom pptx import Presentation\np=Presentation(sys.argv[1])\nprint([sh.shape_type for sh in p.slides[0].shapes])', SP + '/files/media.pptx']).toString().trim();
  check('python-pptx opens it (video as media; it has no audio support) and audio uses <a:audioFile>', /MEDIA: 16/.test(py) && /<a:audioFile r:link=/.test(sx) && (sx.match(/<a:videoFile/g) || []).length === 1, py);
  const back = await p.evaluate(async b => { await AtlasSlides.loadFile({ name: 'm.pptx', path: null, data: new Uint8Array(b) }); return AtlasSlides.pres.slides[0].elements.filter(x => x.type === 'media').map(e => `${e.kind}:${!!e.poster}:${!!e.autoplay}:${!!e.loop}:${!!e.hideIcon}:${e.src.slice(5, 15)}`).join(); }, bytes);
  check('pptx reopens with the clips, posters and playback options', back === 'video:true:true:false:false:video/webm,audio:true:false:true:true:audio/wav;', back);
  // A PowerPoint-style video (python-pptx movie)
  execFileSync('python3', ['-I', '-c', `
import sys
from pptx import Presentation
from pptx.util import Inches
p = Presentation(); s = p.slides.add_slide(p.slide_layouts[6])
s.shapes.add_movie(sys.argv[1], Inches(1), Inches(1), Inches(4), Inches(2.25), mime_type='video/mp4')
p.save(sys.argv[2])`, SP + '/files/clip.mp4', SP + '/files/movie.pptx']);
  const mv = await p.evaluate(async b => { await AtlasSlides.loadFile({ name: 'mv.pptx', path: null, data: new Uint8Array(b) }); return AtlasSlides.pres.slides[0].elements.map(e => e.type + ':' + (e.src || '').slice(5, 14)).join(); }, Array.from(fs.readFileSync(SP + '/files/movie.pptx')));
  check('a PowerPoint movie opens as a video', mv === 'media:video/mp4', mv);
  check('no page errors', errors.length === 0, errors.join(' | '));
  await b.close();
})();
