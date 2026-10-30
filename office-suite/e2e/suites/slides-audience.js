// Slides and the audience: a live poll whose results grow on the slide as people vote, audience
// Q&A, Slide Zoom (go to a slide and come back), and a phone remote that moves the show — all
// through the Isadek Office web server, as in the web version.
const { chromium } = require('playwright');
const path = require('path');
const { createWebServer } = require(path.join(process.cwd(), 'src/server/web-server.js'));
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
(async () => {
  const srv = createWebServer({ live: false, apiKey: '' });
  await new Promise(r => srv.listen(0, '127.0.0.1', r));
  const base = `http://127.0.0.1:${srv.address().port}`;
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 1280, height: 800 } });
  const errors = [];
  p.on('pageerror', e => errors.push(e.message));
  p.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  await p.goto(base + '/src/renderer/slides.html');
  await p.waitForFunction(() => window.AtlasSlides);
  const vote = async (id, a) => (await fetch(`${base}/api/forms/${id}/responses`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ answers: { a } }) })).status;

  // A poll: question and answers, published to the server the app is on.
  await p.click('[data-tab="insert"]');
  await p.click('#insPoll');
  await p.waitForSelector('.modal');
  await p.fill('.modal input', 'Best day for lunch?');
  await p.fill('.modal textarea', 'Mon\nTue\nWed');
  await p.click('.modal .ok');
  await p.waitForFunction(() => AtlasSlides.pres.slides[0].elements.some(e => e.type === 'poll' && e.form));
  const poll = await p.evaluate(() => AtlasSlides.pres.slides[0].elements.find(e => e.type === 'poll'));
  check('Insert → Poll publishes a poll', poll.question === 'Best day for lunch?' && poll.options.join() === 'Mon,Tue,Wed' && /^[\w-]{22}$/.test(poll.form.id), JSON.stringify(poll));
  const qr = await p.waitForFunction(() => { const d = document.querySelector('#stage .poll-qr svg path'); return d && d.getAttribute('d').length > 500; }).then(() => true, () => false);
  check('the poll shows a QR code and the link to answer at', qr && /\/f\//.test(await p.textContent('#stage .poll-link')));
  const pub = await (await fetch(`${base}/api/forms/${poll.form.id}`)).json();
  check('people answer it as a one-question form', pub.form.questions[0].type === 'choice' && pub.form.questions[0].options.join() === 'Mon,Tue,Wed');

  // More slides, and a zoom on slide 1 to slide 3.
  await p.evaluate(() => { AtlasSlides.addSlide('title'); AtlasSlides.addSlide('title'); AtlasSlides.goTo(0); });
  await p.click('#insZoom');
  await p.waitForSelector('.modal select');
  const third = await p.evaluate(() => AtlasSlides.pres.slides[2].id);
  await p.selectOption('.modal select', third);
  await p.click('.modal .ok');
  const zoom = await p.evaluate(() => AtlasSlides.pres.slides[0].elements.find(e => e.type === 'zoom'));
  check('Insert → Zoom links to a slide', zoom && zoom.target === third, JSON.stringify(zoom));
  check('the zoom shows a preview of that slide', (await p.locator('#stage .el.zoom .zoom-inner .el').count()) > 0 && (await p.textContent('#stage .el.zoom .zoom-num')) === '3');
  // Move it clear of the poll, so it can be clicked in the show.
  await p.evaluate(() => { const z = AtlasSlides.pres.slides[0].elements.find(e => e.type === 'zoom'); Object.assign(z, { x: 700, y: 400, w: 220, h: 124 }); AtlasSlides.goTo(0); });

  // Audience Q&A.
  await p.click('[data-tab="show"]');
  await p.click('#qaBtn');
  await p.waitForSelector('.share-qr');
  const qaLink = await p.inputValue('.share-qr input');
  check('Audience Q&A gives a link and a QR code', /\/f\/[\w-]{22}$/.test(qaLink) && (await p.locator('.share-qr svg.qr').count()) === 1, qaLink);
  await p.click('.share-qr .primary');

  // The show: votes arrive and the bars grow.
  await p.evaluate(() => AtlasSlides.present(0));
  await sleep(300);
  for (const a of ['Tue', 'Tue', 'Mon']) await vote(poll.form.id, a);
  check('a wrong answer is refused by the server', (await vote(poll.form.id, 'Sun')) === 422);
  const bars = await p.waitForFunction(() => { const n = document.querySelector('#presentSlide .poll-row[data-option="Tue"] .poll-n'); return n && n.textContent === '67%' && document.querySelector('#presentSlide .poll-total').textContent.includes('3 votes'); }, null, { timeout: 8000 }).then(() => true, () => false);
  check('results appear on the slide as people vote', bars, await p.evaluate(() => document.querySelector('#presentSlide .poll-opts').textContent));
  const qid = qaLink.split('/f/')[1];
  await fetch(`${base}/api/forms/${qid}/responses`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ answers: { q: 'Will the slides be shared?', name: 'Ada' } }) });
  check('the Q&A link shows in the show', await p.waitForSelector('.qa-pill svg.qr', { timeout: 5000 }).then(() => true, () => false));
  await p.waitForFunction(() => /1 question/.test(document.querySelector('.qa-pill .qa-count').textContent), null, { timeout: 8000 });
  await p.keyboard.press('q');
  const q = await p.textContent('.qa-panel .qa-item');
  check('Q shows the audience’s questions', /Will the slides be shared\?/.test(q) && /Ada/.test(q), q);
  await p.click('.qa-panel .qa-done');
  check('a question can be marked answered', (await p.locator('.qa-panel .qa-item').count()) === 0);
  await p.keyboard.press('q');

  // Zoom: click the preview, see slide 3, then the show comes back to slide 1.
  await p.click('#presentSlide .el.zoom');
  check('clicking a zoom goes to its slide', (await p.evaluate(() => AtlasSlides.showIndex)) === 2);
  await p.keyboard.press('ArrowRight');
  check('after the zoomed slide the show comes back', (await p.evaluate(() => AtlasSlides.showIndex)) === 0);
  await p.keyboard.press('Escape');

  // The phone remote.
  await p.click('#remoteBtn');
  await p.waitForSelector('.share-qr');
  const remoteLink = await p.inputValue('.share-qr input');
  check('Phone Remote gives a link with its secret after #', /\/r\/[\w-]{22}#[\w-]{40,}$/.test(remoteLink), remoteLink);
  await p.click('.share-qr .primary');
  const phone = await b.newPage({ viewport: { width: 390, height: 844 }, hasTouch: true, isMobile: true });
  phone.on('pageerror', e => errors.push('remote: ' + e.message));
  await phone.goto(remoteLink);
  await phone.waitForFunction(() => /hasn’t started/.test(document.getElementById('count').textContent));
  check('the phone knows the show hasn’t started', true);
  await phone.tap('#nextBtn');
  await p.waitForFunction(() => AtlasSlides.showIndex === 0, null, { timeout: 8000 });
  check('Next on the phone starts the show', true);
  await phone.waitForFunction(() => document.getElementById('count').textContent === 'Slide 1 of 3');
  await phone.tap('#nextBtn');
  await p.waitForFunction(() => AtlasSlides.showIndex === 1, null, { timeout: 8000 });
  await phone.waitForFunction(() => document.getElementById('count').textContent === 'Slide 2 of 3');
  check('the phone moves the slides and follows them', true);
  await phone.tap('#black');
  await p.waitForFunction(() => document.getElementById('presenter').classList.contains('black'), null, { timeout: 8000 });
  check('Black from the phone', true);
  const forged = await fetch(remoteLink.replace(/\/r\/([\w-]+)#.*/, '/api/remote/$1/command'), { method: 'POST', headers: { Authorization: 'Bearer wrong-token-0000000000000', 'Content-Type': 'application/json' }, body: JSON.stringify({ cmd: 'next' }) });
  check('a guessed token can’t move the show', forged.status === 403);
  await p.keyboard.press('Escape');
  await p.click('#remoteBtn');
  await p.waitForSelector('.modal');
  await p.click('.modal button:has-text("Stop Remote")');
  await phone.waitForFunction(() => /ended/.test(document.getElementById('count').textContent), null, { timeout: 25000 });
  check('stopping the remote ends it on the phone', true);

  // Saved as PowerPoint and back: the poll, its form and the zoom come back.
  const bytes = await p.evaluate(async () => Array.from(await AtlasSlides.serialize('pptx')));
  await p.evaluate(async b => AtlasSlides.loadFile({ name: 'talk.pptx', path: null, data: new Uint8Array(b) }), bytes);
  const back = await p.evaluate(() => { const s = AtlasSlides.pres.slides; const pl = s[0].elements.find(e => e.type === 'poll'); const z = s[0].elements.find(e => e.type === 'zoom'); return { q: pl && pl.question, id: pl && pl.form && pl.form.id, zoom: z && s.findIndex(x => x.id === z.target), qa: !!AtlasSlides.pres.qa }; });
  check('PowerPoint files keep polls, zooms and Q&A', back.q === 'Best day for lunch?' && back.id === poll.form.id && back.zoom === 2 && back.qa, JSON.stringify(back));

  check('no page errors', errors.length === 0, errors.join(' | '));
  await b.close();
  srv.close();
})().catch(e => { console.log('FAIL crashed :: ' + e.message); process.exit(1); });
