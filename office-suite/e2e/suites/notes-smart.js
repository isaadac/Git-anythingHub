// Isadek Notes: Ink to Text (on this device, and with the Assistant reading a picture of the
// writing), Picture to Text and finding words in pictures, Math (answers, steps, graphs) and the
// web clipper (an article with its pictures, or a bookmark), through the web server with a
// stand-in for the Anthropic API and a stand-in web site.
const { chromium } = require('playwright');
const http = require('http');
const path = require('path');
const { createWebServer } = require(path.join(process.cwd(), 'src/server/web-server.js'));
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };

(async () => {
  // The Assistant's stand-in: reads a picture as the two lines written in the test.
  const asked = [];
  const api = http.createServer(async (req, res) => {
    let b = ''; for await (const c of req) b += c;
    const body = JSON.parse(b || '{}');
    const last = body.messages[body.messages.length - 1].content;
    asked.push(last);
    const image = Array.isArray(last) && last.find(x => x.type === 'image');
    const reply = image ? 'Buy milk and eggs\nCall Sam at four' : 'Step 1: subtract 3. Step 2: divide by 2.';
    res.writeHead(200, { 'content-type': 'text/event-stream' });
    const ev = (t, d) => res.write(`event: ${t}\ndata: ${JSON.stringify({ type: t, ...d })}\n\n`);
    ev('message_start', { message: { id: 'm', type: 'message', role: 'assistant', model: body.model, content: [], stop_reason: null, usage: { input_tokens: 1, output_tokens: 0 } } });
    ev('content_block_start', { index: 0, content_block: { type: 'text', text: '' } });
    ev('content_block_delta', { index: 0, delta: { type: 'text_delta', text: reply } });
    ev('content_block_stop', { index: 0 });
    ev('message_delta', { delta: { stop_reason: 'end_turn' }, usage: { output_tokens: 1 } });
    ev('message_stop', {});
    res.end();
  });
  await new Promise(r => api.listen(0, '127.0.0.1', r));
  const web = createWebServer({ apiKey: 'sk-ant-test', open: true, baseURL: `http://127.0.0.1:${api.address().port}`, env: {}, live: false });
  await new Promise(r => web.listen(0, '127.0.0.1', r));
  const BASE = `http://127.0.0.1:${web.address().port}`;

  const b = await chromium.launch();
  const ctx = await b.newContext({ viewport: { width: 1400, height: 900 } });
  // The stand-in web site (it allows pages elsewhere to read it, as a site must for the web version).
  const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAFklEQVR4nGP8z8DAwMDAxMDAwMDAAAANHQEDK+mmyQAAAABJRU5ErkJggg==', 'base64');
  const ARTICLE = `<!doctype html><html><head><meta charset="utf-8"><title>Tide pools | Coast News</title>
    <meta property="og:title" content="Life in the tide pools"><meta property="og:site_name" content="Coast News"><meta name="description" content="What lives between the tides.">
    <script>window.evil = 1</script><style>p{color:red}</style></head><body>
    <nav class="menu"><a href="/">Home</a> <a href="/sports">Sports headlines</a></nav>
    <div class="share-buttons">Share this on social media</div>
    <article><h1>Life in the tide pools</h1>
      <p>Twice a day the sea pulls back and leaves small worlds behind in the rocks, full of anemones, crabs and sea stars.</p>
      <figure><img src="/pics/crab.png" alt="A crab"><figcaption>A shore crab</figcaption></figure>
      <p>Visitors should step carefully and <a href="/guide">read the guide</a> before exploring, because many animals hide under weed.</p>
      <p onclick="alert(1)">Never take animals home: they need the cold water of their pool to survive the day.</p>
      <img src="javascript:alert(1)"><img src="https://tracker.example/pixel.gif">
    </article>
    <aside>Related stories you might like</aside><footer>Copyright Coast News</footer></body></html>`;
  await ctx.route('https://coast.example/**', r => {
    const u = new URL(r.request().url());
    const h = { 'access-control-allow-origin': '*' };
    if (u.pathname === '/pics/crab.png') return r.fulfill({ body: PNG, headers: { ...h, 'content-type': 'image/png' } });
    if (u.pathname === '/tides') return r.fulfill({ body: ARTICLE, headers: { ...h, 'content-type': 'text/html; charset=utf-8' } });
    return r.fulfill({ status: 404, body: '', headers: h });
  });
  await ctx.route('https://tracker.example/**', r => r.fulfill({ status: 404, body: '' }));
  const p = await ctx.newPage();
  const errors = [];
  p.on('pageerror', e => errors.push(e.message));
  await p.goto(BASE + '/src/renderer/notes.html');
  await p.waitForFunction(() => window.AtlasNotes && AtlasNotes.page && window.NotesSmart);
  // Dialogs answer from a list; anything else answers as typed (the default values).
  const answer = list => p.evaluate(list => {
    window.__a = list.slice();
    window.__seen = [];
    UI.form = async (title, fields) => { const want = window.__a.shift(); window.__seen.push({ title, fields: fields.map(f => ({ name: f.name, value: f.value, options: f.options })) }); if (want === null) return null; const out = {}; fields.forEach(f => { out[f.name] = f.value !== undefined ? f.value : (f.options ? (f.options[0].value ?? f.options[0]) : ''); }); return Object.assign(out, want || {}); };
  }, list);
  const seen = () => p.evaluate(() => window.__seen);
  const tab = name => p.click(`[data-tab="${name}"]`);

  // ---- Ink to Text: handwriting-like ink (letters with a wobble, drawn as pen strokes).
  const writeInk = () => p.evaluate(() => {
    const lines = ['Buy milk and eggs', 'Call Sam at four'];
    const c = document.createElement('canvas'); c.width = 700; c.height = 220;
    const g = c.getContext('2d');
    g.font = '40px sans-serif';
    let seed = 7; const rnd = () => (seed = (seed * 16807) % 2147483647) / 2147483647;
    lines.forEach((t, li) => { let x = 10; for (const ch of t) { g.save(); g.translate(x, 70 + li * 90 + (rnd() - 0.5) * 4); g.rotate((rnd() - 0.5) * 0.06); g.fillText(ch, 0, 0); g.restore(); x += g.measureText(ch).width + 1 + rnd() * 2; } });
    const d = g.getImageData(0, 0, c.width, c.height).data, strokes = [];
    const top = document.getElementById('body').getBoundingClientRect().bottom - document.getElementById('page').getBoundingClientRect().top + 20;
    for (let x = 0; x < c.width; x++) { let y0 = -1; for (let y = 0; y <= c.height; y++) { const on = y < c.height && d[(y * c.width + x) * 4 + 3] > 128; if (on && y0 < 0) y0 = y; if (!on && y0 >= 0) { strokes.push({ c: '#1f1f1f', w: 1.6, o: 1, d: `M${x + 60} ${y0 + top} L${x + 60} ${y + top}` }); y0 = -1; } } }
    // A highlighter mark is left alone.
    strokes.push({ c: '#fff176', w: 14, o: 0.45, d: `M60 ${top} L400 ${top}` });
    AtlasNotes.page.ink.push(...strokes);
    return strokes.length;
  });
  await p.evaluate(() => { document.getElementById('body').innerHTML = '<p>Shopping</p>'; });
  await writeInk();
  await answer([{}]);
  await tab('draw');
  await p.click('#inkToText');
  await p.waitForFunction(() => /Buy milk/.test(document.getElementById('body').textContent), null, { timeout: 60000 }).catch(() => {});
  const paras = await p.$$eval('#body p', ps => ps.map(x => x.textContent));
  check('Ink to Text reads both lines of handwriting on this device', paras.includes('Buy milk and eggs') && paras.includes('Call Sam at four'), paras.join(' | '));
  check('…below the text that was above the writing', paras[0] === 'Shopping');
  const left = await p.evaluate(() => AtlasNotes.page.ink.map(s => s.o));
  check('…the writing is replaced, the highlighter kept', left.length === 1 && left[0] < 1, String(left));
  check('…the text can be corrected before it goes in (and the Assistant is offered)', (await seen())[0].fields.some(f => f.name === 'text') && (await seen())[0].fields.some(f => f.name === 'reader'));

  // With the Assistant reading it instead (and keeping the drawing).
  await p.evaluate(() => { document.getElementById('body').innerHTML = '<p><br></p>'; AtlasNotes.page.ink.length = 0; });
  await writeInk();
  await answer([{ reader: 'assistant', ink: 'keep' }, {}]);
  await p.click('#inkToText');
  await p.waitForFunction(() => /Call Sam/.test(document.getElementById('body').textContent), null, { timeout: 60000 }).catch(() => {});
  const sent = asked.find(m => Array.isArray(m));
  check('the Assistant is sent a PNG of the writing with the instruction', !!sent && sent[0].type === 'image' && sent[0].source.media_type === 'image/png' && /Transcribe the handwriting/.test(sent[1].text));
  check('…its reading goes in, and the drawing stays when asked', /Buy milk and eggs/.test(await p.textContent('#body')) && (await p.evaluate(() => AtlasNotes.page.ink.length)) > 100);
  await p.evaluate(() => { AtlasNotes.page.ink.length = 0; });

  // ---- Picture to Text, and words in pictures found by search.
  const textPicture = (text, w = 640) => p.evaluate(([text, w]) => { const c = document.createElement('canvas'); c.width = w; c.height = 90; const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, w, 90); g.fillStyle = '#222'; g.font = '30px serif'; g.fillText(text, 14, 56); return c.toDataURL('image/png'); }, [text, w]);
  const pic1 = await textPicture('Quarterly revenue grew 12%');
  await p.evaluate(src => { document.getElementById('body').innerHTML = `<p>Board slide:</p><p><img id="pic1" src="${src}" alt=""></p><p>After</p>`; }, pic1);
  await answer([{}]);
  await tab('insert');
  await p.click('#pic1');
  await p.click('#picToText');
  await p.waitForFunction(() => /Quarterly revenue/.test(document.getElementById('body').textContent), null, { timeout: 60000 }).catch(() => {});
  const after = await p.$$eval('#body p', ps => ps.map(x => x.textContent.trim()));
  check('Picture to Text reads the picture and puts the words below it', after[2] === 'Quarterly revenue grew 12%' && after[3] === 'After', after.join(' | '));
  check('…and the picture keeps its words for search', /Quarterly revenue grew 12/.test(await p.getAttribute('#pic1', 'data-ocr')));
  // A picture added later is read in the background.
  const pic2 = await textPicture('Meeting with the harbour master');
  await p.evaluate(src => { const b = document.getElementById('body'); b.focus(); const r = document.createRange(); r.selectNodeContents(b); r.collapse(false); getSelection().removeAllRanges(); getSelection().addRange(r); document.execCommand('insertHTML', false, `<p><img id="pic2" src="${src}" alt=""></p>`); }, pic2);
  await p.waitForFunction(() => /harbour/i.test(document.getElementById('pic2')?.getAttribute('data-ocr') || ''), null, { timeout: 60000 }).catch(() => {});
  check('a picture added to the page is read in the background', /harbour master/i.test(await p.getAttribute('#pic2', 'data-ocr') || ''));
  await p.fill('#search', 'harbour master');
  await p.waitForTimeout(400);
  check('…so searching finds the page by words in a picture', (await p.locator('#results li:not(.none)').count()) === 1);
  await p.fill('#search', '');
  await p.dispatchEvent('#search', 'input');
  const saved = await p.evaluate(() => AtlasNotes.notebook.sections[0].pages[0].html);
  check('…and the words are kept in the notebook (sanitized attribute)', /data-ocr="[^"]*harbour/i.test(saved));

  // ---- Math
  await p.evaluate(() => { const b = document.getElementById('body'); b.innerHTML = '<p id="eq">2x + 3 = 11</p>'; b.focus(); const r = document.createRange(); r.selectNodeContents(document.getElementById('eq')); getSelection().removeAllRanges(); getSelection().addRange(r); });
  await tab('draw');
  await p.click('#mathBtn');
  check('Math opens with the selected equation solved', (await p.isVisible('#mathPane')) && (await p.textContent('.math-answer')) === 'x = 4');
  const steps = await p.$$eval('#mathOut .math-steps li', l => l.map(x => x.textContent));
  check('…with the steps', steps.some(s => /Subtract 3/.test(s)) && steps.some(s => /Divide both sides by 2/.test(s)), steps.join(' / '));
  await p.evaluate(() => { const b = document.getElementById('body'); const r = document.createRange(); r.selectNodeContents(b); r.collapse(false); getSelection().removeAllRanges(); getSelection().addRange(r); });
  await p.click('.math-actions button:has-text("Insert Answer")');
  check('Insert Answer puts the answer in the page', /2x \+ 3 = 11 → x = 4/.test(await p.textContent('#body')));
  await p.fill('#mathInput', 'x² − 5x + 6 = 0');
  check('typing another equation solves it (quadratic)', (await p.textContent('.math-answer')) === 'x = 2 or x = 3');
  await p.fill('#mathInput', '12 × (3 + 4) =');
  check('…and works out a value', (await p.textContent('.math-answer')) === '84');
  await p.fill('#mathInput', '2 # 3');
  check('…and says what it can’t read', /“#”/.test(await p.textContent('.math-error')));
  await p.fill('#mathInput', 'y = x^2 - 4');
  check('a function is graphed, with where it crosses zero', (await p.locator('.math-graph').count()) === 1 && /−2, x = 2/.test(await p.textContent('.math-roots')));
  await p.click('.math-actions button:has-text("Insert Graph")');
  await p.waitForFunction(() => [...document.querySelectorAll('#body img')].some(i => i.src.startsWith('data:image/png')));
  const g = await p.$eval('#body img[alt^="Graph"]', i => ({ w: i.naturalWidth, src: i.src.slice(0, 22) }));
  check('Insert Graph puts the graph in the page as a picture', g.w === 560 && g.src === 'data:image/png;base64,', JSON.stringify(g));
  await p.click('.math-actions button:has-text("Explain with Assistant")');
  await p.waitForFunction(() => /Step 1/.test(document.querySelector('.ai-log')?.textContent || ''), null, { timeout: 15000 }).catch(() => {});
  check('Explain with Assistant asks the Assistant to explain it', /explain step by step.*y = x\^2 - 4/i.test(JSON.stringify(asked[asked.length - 1])) && await p.isVisible('.ai-pane'));
  await p.click('.ai-pane .x');
  await p.click('#mathClose');

  // ---- Web Clip
  const pagesBefore = await p.evaluate(() => AtlasNotes.notebook.sections[0].pages.length);
  await answer([{ url: 'https://coast.example/tides', mode: 'article' }]);
  await tab('insert');
  await p.click('#webClip');
  await p.waitForFunction(() => document.getElementById('pageTitle').value === 'Life in the tide pools', null, { timeout: 15000 }).catch(() => {});
  const clip = await p.evaluate(() => ({ title: document.getElementById('pageTitle').value, html: document.getElementById('body').innerHTML, text: document.getElementById('body').textContent, pages: AtlasNotes.notebook.sections[0].pages.length }));
  check('Web Clip makes a new page titled like the article', clip.title === 'Life in the tide pools' && clip.pages === pagesBefore + 1, clip.title);
  check('…with the article text', /Twice a day the sea pulls back/.test(clip.text) && /Never take animals home/.test(clip.text));
  check('…without menus, sharing buttons, related stories or footers', !/Sports headlines|Share this|Related stories|Copyright/.test(clip.text));
  check('…a link back to where it came from', /class="clip-src"[^>]*>Clipped from <a href="https:\/\/coast\.example\/tides">Coast News<\/a>/.test(clip.html), clip.html.slice(0, 160));
  check('…links made absolute', /href="https:\/\/coast\.example\/guide"/.test(clip.html));
  check('…its picture kept in the notebook, and nothing that loads from the web or runs', /<img src="data:image\/png;base64,/.test(clip.html) && !/tracker\.example|javascript:|onclick|<script|evil/.test(clip.html) && (await p.locator('#body img').count()) === 1);
  await answer([{ url: 'https://coast.example/tides', mode: 'bookmark' }]);
  await p.click('#webClip');
  await p.waitForFunction(n => AtlasNotes.notebook.sections[0].pages.length === n, pagesBefore + 2, { timeout: 15000 }).catch(() => {});
  const bm = await p.evaluate(() => document.getElementById('body').innerHTML);
  check('a bookmark clip has the title, summary and link only', /<h2><a href="https:\/\/coast\.example\/tides">Life in the tide pools<\/a><\/h2><p>What lives between the tides\.<\/p>/.test(bm) && !/Twice a day/.test(bm), bm.slice(0, 300));
  await answer([{ url: 'https://coast.example/missing', mode: 'article' }]);
  await p.click('#webClip');
  await p.waitForTimeout(800);
  check('a page that isn’t there makes no new page', (await p.evaluate(() => AtlasNotes.notebook.sections[0].pages.length)) === pagesBefore + 2);
  await answer([{ url: 'file:///etc/passwd', mode: 'article' }]);
  await p.click('#webClip');
  check('only web addresses can be clipped', (await p.evaluate(() => AtlasNotes.notebook.sections[0].pages.length)) === pagesBefore + 2);

  check('no page errors', !errors.length, errors.join(' | '));
  await b.close();
  web.close(); api.close();
})().catch(e => { console.error('CRASH', e.stack); process.exit(1); });
