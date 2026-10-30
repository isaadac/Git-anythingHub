// The document-aware Assistant: rewriting a section as tracked changes, building decks from
// outlines (in Slides, and from a Write document), a whole notebook as context, questions that take
// in the person's other open files, and dictation. Replies come from a stand-in for the API.
const { chromium } = require('playwright');
const http = require('http');
const path = require('path');
const { createWebServer } = require(path.join(process.cwd(), 'src/server/web-server.js'));
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
const KEY = 'sk-ant-test-key-docs-0001';

(async () => {
  const reqs = [];
  let reply = 'Hello';
  const api = http.createServer(async (req, res) => {
    let b = ''; for await (const c of req) b += c;
    const body = JSON.parse(b || '{}');
    reqs.push(body);
    res.writeHead(200, { 'content-type': 'text/event-stream' });
    const ev = (t, d) => res.write(`event: ${t}\ndata: ${JSON.stringify({ type: t, ...d })}\n\n`);
    ev('message_start', { message: { id: 'msg_1', type: 'message', role: 'assistant', model: body.model, content: [], stop_reason: null, usage: { input_tokens: 10, output_tokens: 0 } } });
    ev('content_block_start', { index: 0, content_block: { type: 'text', text: '' } });
    for (const part of reply.match(/[\s\S]{1,40}/g)) ev('content_block_delta', { index: 0, delta: { type: 'text_delta', text: part } });
    ev('content_block_stop', { index: 0 });
    ev('message_delta', { delta: { stop_reason: 'end_turn' }, usage: { output_tokens: 5 } });
    ev('message_stop', {});
    res.end();
  });
  await new Promise(r => api.listen(0, '127.0.0.1', r));
  const web = createWebServer({ apiKey: KEY, codes: ['letmein'], perMinute: 1000, baseURL: `http://127.0.0.1:${api.address().port}`, env: {} });
  await new Promise(r => web.listen(0, '127.0.0.1', r));
  const BASE = `http://127.0.0.1:${web.address().port}/src/renderer/`;
  const lastUser = () => { const m = reqs[reqs.length - 1].messages; return m[m.length - 1].content; };

  const browser = await chromium.launch();
  const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 } });
  await ctx.addInitScript(() => {
    localStorage.setItem('atlas-assistant-code', 'letmein');
    localStorage.setItem('atlas-author', 'Ada');
    // A stand-in speech recognizer: "hears" two phrases, then stops when asked.
    class FakeSR {
      start() { window.__srStarted = (window.__srStarted || 0) + 1; this.t = setTimeout(() => { this.onresult && this.onresult({ resultIndex: 0, results: [Object.assign([{ transcript: 'hello from my voice' }], { isFinal: true })] }); }, 50); }
      stop() { clearTimeout(this.t); setTimeout(() => this.onend && this.onend(), 10); }
    }
    if (!window.__noSpeech) window.SpeechRecognition = FakeSR;
  });
  const errors = [];
  const watch = pg => pg.on('pageerror', e => errors.push(e.message));
  const ask = async (pg, chipOrText, wait = true) => {
    const n = await pg.$$eval('.ai-msg.assistant', ms => ms.length);
    if (chipOrText.startsWith('chip:')) await pg.click(`.ai-actions .chip:has-text("${chipOrText.slice(5)}")`);
    else { await pg.fill('.ai-input textarea', chipOrText); await pg.click('.ai-input .send'); }
    if (wait) await pg.waitForFunction(n => { const ms = document.querySelectorAll('.ai-msg.assistant'); return ms.length > n && ms[ms.length - 1].querySelector('.ai-use'); }, n, { timeout: 10000 });
  };
  const use = (pg, label) => pg.click(`.ai-msg.assistant:last-of-type .ai-use button:has-text("${label}")`);

  // ---- Write: rewrite a section as tracked changes
  const w = await ctx.newPage(); watch(w);
  await w.goto(BASE + 'write.html');
  await w.waitForFunction(() => window.AtlasWrite && window.Assistant);
  await w.evaluate(() => { document.getElementById('doc').innerHTML = '<h1>Plan</h1><p>Intro text stays.</p><h2>Budget</h2><p>We will spend a lot of money on the new office.</p><p>This paragraph is fine.</p><h2>Timeline</h2><p>Starts in May.</p>'; });
  await w.click('#doc h2:first-of-type + p');
  await w.click('.ai-btn');
  reply = '## Budget\n\nWe will spend about $40,000 on the new office.\n\nThis paragraph is fine.\n\nA reserve of 10% covers surprises.';
  await ask(w, 'chip:Rewrite this section');
  const sent = lastUser();
  check('"Rewrite this section" sends just that section', /<selection>[\s\S]*## Budget[\s\S]*a lot of money[\s\S]*This paragraph is fine[\s\S]*<\/selection>/.test(sent) && !/<selection>[\s\S]*(Starts in May|Intro text)[\s\S]*<\/selection>/.test(sent), sent.slice(sent.indexOf('<selection>'), sent.indexOf('</selection>') + 12));
  await use(w, 'Suggest as tracked changes');
  const tr = await w.evaluate(() => {
    const ps = [...document.querySelectorAll('#doc > *')];
    return {
      html: document.getElementById('doc').innerHTML,
      heading: ps[2].outerHTML,
      p1del: [...ps[3].querySelectorAll('del.tc')].map(d => d.textContent).join('|'),
      p1ins: [...ps[3].querySelectorAll('ins.tc')].map(d => d.textContent).join('|'),
      untouched: ps[4].textContent === 'This paragraph is fine.' && !ps[4].querySelector('.tc'),
      added: ps[5].tagName === 'P' && ps[5].querySelector('ins.tc') && ps[5].textContent === 'A reserve of 10% covers surprises.',
      rest: ps[6].textContent === 'Timeline' && ps[7].textContent === 'Starts in May.' && ps[1].textContent === 'Intro text stays.',
      author: (document.querySelector('#doc ins.tc') || {}).dataset?.author
    };
  });
  check('changed words are tracked, word by word', tr.p1del === 'a lot of money' && tr.p1ins === 'about $40,000', tr.p1del + ' → ' + tr.p1ins);
  check('unchanged paragraphs and the heading are left alone', tr.untouched && !/class="tc"/.test(tr.heading), tr.heading);
  check('a new paragraph is a tracked insertion', !!tr.added);
  check('the rest of the document is untouched', tr.rest);
  check('the changes carry the author', tr.author === 'Ada', tr.author);
  await w.evaluate(() => document.getElementById('doc').focus());
  await w.keyboard.press('Control+z');
  check('one undo takes the whole suggestion back', await w.evaluate(() => !document.querySelector('#doc .tc') && /a lot of money/.test(document.getElementById('doc').textContent)));
  await w.keyboard.press('Control+y');
  await w.click('.ribbon-tabs button[data-tab="review"]');
  await w.click('#acceptAll').catch(async () => { await w.evaluate(() => [...document.querySelectorAll('button')].find(b => /Accept all/i.test(b.textContent)).click()); });
  const accepted = await w.evaluate(() => [...document.querySelectorAll('#doc > *')].map(e => e.textContent).join(' / '));
  check('accepting them gives exactly the rewrite', accepted === 'Plan / Intro text stays. / Budget / We will spend about $40,000 on the new office. / This paragraph is fine. / A reserve of 10% covers surprises. / Timeline / Starts in May.', accepted);

  // A selection inside one paragraph: word changes there only.
  await w.evaluate(() => {
    const p = document.querySelector('#doc p:last-of-type'); p.textContent = 'The launch starts in may if the weather allows.';
    const r = document.createRange(); r.setStart(p.firstChild, 4); r.setEnd(p.firstChild, 24); const s = getSelection(); s.removeAllRanges(); s.addRange(r);
    document.getElementById('doc').dispatchEvent(new Event('keyup'));
    document.dispatchEvent(new Event('selectionchange'));
  });
  await sleep(100);
  reply = 'launch begins in May';
  await ask(w, 'chip:Improve writing');
  await use(w, 'Suggest as tracked changes');
  const inl = await w.evaluate(() => { const p = document.querySelector('#doc p:last-of-type'); return { d: [...p.querySelectorAll('del.tc')].map(x => x.textContent).join('|'), i: [...p.querySelectorAll('ins.tc')].map(x => x.textContent).join('|'), all: p.textContent.replace(/\u00a0/g, ' ') }; });
  check('a rewrite of selected words is tracked inside the paragraph', inl.d === 'starts|may' && inl.i === 'begins|May' && /^The launch startsbegins in mayMay if the weather allows\.$/.test(inl.all), JSON.stringify(inl));

  // ---- Write → Slides: a document becomes a presentation
  reply = '# Plan\nOur office move\n\n# Money\n\n## Budget\n- About $40,000\n- 10% reserve\nNotes: Mention the reserve.\n\n## Timeline\n- Starts in May\n- Done by July';
  await ask(w, 'chip:Make a presentation');
  const [deck] = await Promise.all([ctx.waitForEvent('page'), use(w, 'Open as slides')]);
  watch(deck);
  await deck.waitForFunction(() => window.AtlasSlides && AtlasSlides.pres.slides.length > 1, null, { timeout: 10000 });
  const d = await deck.evaluate(() => {
    const txt = s => s.elements.map(e => (e.html || '').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim()).filter(Boolean).join(' | ');
    return { name: document.querySelector('.titlebar .docname input').value, kinds: AtlasSlides.pres.slides.map(s => s.layout).join(','), texts: AtlasSlides.pres.slides.map(txt), notes: AtlasSlides.pres.slides[2].notes };
  });
  check('"Open as slides" opens a new presentation built from the outline', d.kinds === 'title,section,content,content' && /\.pptx$/.test(d.name), JSON.stringify(d));
  check('…with titles, points and speaker notes', d.texts[0] === 'Plan | Our office move' && d.texts[2] === 'Budget | About $40,000 10% reserve' && d.notes === 'Mention the reserve.', JSON.stringify(d.texts));

  // ---- Slides: a deck from an outline replaces a blank presentation
  await deck.evaluate(() => { window.__atlasDirty = false; });
  const s = await ctx.newPage(); watch(s);
  await s.goto(BASE + 'slides.html');
  await s.waitForFunction(() => window.AtlasSlides && window.Assistant);
  await s.click('.ai-btn');
  await s.click('.ai-actions .chip:has-text("Deck from outline")');
  check('"Deck from outline" waits for the outline', /Outline:\s*$/.test(await s.inputValue('.ai-input textarea')));
  reply = '# Quarterly review\nTeam A\n\n## Results\n- Revenue up\n- Costs flat\n\n## Next\n- Hire two people';
  await s.fill('.ai-input textarea', (await s.inputValue('.ai-input textarea')) + 'results, next steps');
  await s.click('.ai-input .send');
  await s.waitForSelector('.ai-msg.assistant .ai-use');
  await use(s, 'Build deck');
  const built = await s.evaluate(() => ({ n: AtlasSlides.pres.slides.length, kinds: AtlasSlides.pres.slides.map(x => x.layout).join(',') }));
  check('"Build deck" turns a blank presentation into the deck', built.n === 3 && built.kinds === 'title,content,content', JSON.stringify(built));
  await s.evaluate(() => { window.__atlasDirty = false; });

  // ---- Other open files
  const sh = await ctx.newPage(); watch(sh);
  await sh.goto(BASE + 'sheets.html');
  await sh.waitForFunction(() => window.AtlasSheets && window.Assistant);
  await sh.evaluate(() => { AtlasSheets.setCell('A1', 'Office rent'); AtlasSheets.setCell('B1', '3200'); });
  await w.bringToFront();
  await w.check('.ai-others input');
  reply = 'Rent is 3200 a month.';
  await ask(w, 'What does the rent cost, given my spreadsheet?');
  const withOthers = lastUser();
  check('a question can take in the other open files', /<open_file app="sheets"[^>]*>[\s\S]*Office rent[\s\S]*3200[\s\S]*<\/open_file>/.test(withOthers) && /<open_file app="slides"/.test(withOthers) && /<document>/.test(withOthers), withOthers.slice(0, 300));
  check('…and the reply says how many', /with \d other open files/.test(await w.evaluate(() => [...document.querySelectorAll('.ai-msg.assistant .body')].pop().dataset.with || '')));
  await w.uncheck('.ai-others input');
  await ask(w, 'Just this one');
  check('only when asked for', !/<open_file/.test(lastUser()));
  await sh.evaluate(() => { window.__atlasDirty = false; });
  await s.close(); await sh.close(); await deck.close();

  // ---- Notes: the whole notebook
  const n = await ctx.newPage(); watch(n);
  await n.goto(BASE + 'notes.html');
  await n.waitForFunction(() => window.Assistant && document.querySelector('.nb-list li'));
  await n.evaluate(async () => {
    document.getElementById('pageTitle').value = 'Groceries';
    document.getElementById('body').innerHTML = '<p>Milk and eggs</p>';
    document.getElementById('body').dispatchEvent(new Event('input', { bubbles: true }));
  });
  await n.click('#secAdd');
  if (await n.$('.modal input')) { await n.fill('.modal input', 'Travel'); await n.click('.modal .actions .primary'); }
  await n.waitForFunction(() => document.querySelectorAll('#sections li').length === 2);
  await n.evaluate(() => { document.getElementById('pageTitle').value = 'Trip'; document.getElementById('body').innerHTML = '<p>Book the train</p>'; document.getElementById('body').dispatchEvent(new Event('input', { bubbles: true })); });
  await n.click('.ai-btn');
  reply = '# Summary\n- Groceries\n- Trip';
  await ask(n, 'chip:Summarize notebook');
  const nbDoc = lastUser();
  check('"Summarize notebook" sends every section and page', /# Notebook:/.test(nbDoc) && /Milk and eggs/.test(nbDoc) && /Book the train/.test(nbDoc) && (nbDoc.match(/## Section:/g) || []).length >= 1, nbDoc.slice(0, 400));

  // ---- Dictation
  await w.bringToFront();
  await w.click('.ribbon-tabs button[data-tab="home"]');
  await w.evaluate(() => { const p = document.querySelector('#doc p'); const r = document.createRange(); r.selectNodeContents(p); r.collapse(false); const s = getSelection(); s.removeAllRanges(); s.addRange(r); });
  await w.click('#dictate');
  await w.waitForFunction(() => /hello from my voice/.test(document.querySelector('#doc p').textContent), null, { timeout: 3000 }).catch(() => {});
  const dict = await w.evaluate(() => ({ text: document.querySelector('#doc p').textContent, on: document.getElementById('dictate').getAttribute('aria-pressed') }));
  check('Dictate types what is said at the cursor, spaced after a word', /Intro text stays\. hello from my voice $/.test(dict.text.replace(/\u00a0/g, ' ')) && dict.on === 'true', JSON.stringify(dict));
  await w.click('#dictate');
  await sleep(100);
  check('clicking again stops listening', (await w.getAttribute('#dictate', 'aria-pressed')) === 'false');
  await w.click('.ai-input .mic');
  await w.waitForFunction(() => /hello from my voice/.test(document.querySelector('.ai-input textarea').value), null, { timeout: 3000 }).catch(() => {});
  check('the Assistant\'s microphone dictates the question', /hello from my voice/.test(await w.inputValue('.ai-input textarea')));
  await w.click('.ai-input .mic');
  // Without a browser speech recognizer, it offers the on-device one.
  const ns = await ctx.newPage(); watch(ns);
  await ns.addInitScript(() => { window.__noSpeech = true; delete window.webkitSpeechRecognition; });
  await ns.goto(BASE + 'write.html');
  await ns.waitForFunction(() => window.Assistant);
  await ns.evaluate(() => { delete window.SpeechRecognition; delete window.webkitSpeechRecognition; });
  await ns.click('#dictate');
  await ns.waitForSelector('.modal', { timeout: 5000 }).catch(() => {});
  check('…and without a browser recognizer, offers speech recognition on this device', /On this device/.test(await ns.textContent('.modal').catch(() => '')));
  await ns.click('.modal .actions button:has-text("Cancel")').catch(() => {});

  for (const pg of [w, n, ns]) await pg.evaluate(() => { window.__atlasDirty = false; }).catch(() => {});
  check('no page errors', !errors.length, errors.slice(0, 3).join(' | '));
  await browser.close();
  web.close(); api.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
