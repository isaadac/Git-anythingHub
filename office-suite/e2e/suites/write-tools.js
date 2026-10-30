// Isadek Write: form fields (and Word content controls), restricting a document to filling in,
// watermark, line numbers, drop caps, the outline, envelopes and labels, readability, and
// translation (through the web server, with a stand-in for the Anthropic API).
const { chromium } = require('playwright');
const http = require('http');
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const JSZip = require(path.join(process.cwd(), 'node_modules/jszip'));
const { createWebServer } = require(path.join(process.cwd(), 'src/server/web-server.js'));
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const W = process.env.E2E_WORK;
(async () => {
  // The Assistant's stand-in: translations into Spanish of the words it knows.
  const ES = { Hello: 'Hola', world: 'mundo', 'Good morning': 'Buenos días', Results: 'Resultados' };
  const tr = t => Object.entries(ES).reduce((s, [a, b]) => s.split(a).join(b), t);
  const api = http.createServer(async (req, res) => {
    let b = ''; for await (const c of req) b += c;
    const body = JSON.parse(b || '{}');
    const text = body.messages[body.messages.length - 1].content;
    const reply = /JSON array of HTML/.test(body.system) ? JSON.stringify(JSON.parse(text).map(tr)) : tr(text);
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
  const p = await b.newPage({ viewport: { width: 1400, height: 900 } });
  const errors = [];
  p.on('pageerror', e => errors.push(e.message));
  await p.goto(BASE + '/src/renderer/write.html');
  await p.waitForFunction(() => window.AtlasWrite && AtlasWrite.forms);
  const answer = list => p.evaluate(list => { window.__a = list.slice(); UI.form = async () => window.__a.shift(); }, list);
  const caretAtEnd = sel => p.evaluate(sel => { const el = document.querySelector(sel); const r = document.createRange(); r.selectNodeContents(el); r.collapse(false); const s = getSelection(); s.removeAllRanges(); s.addRange(r); el.closest('#doc').focus(); }, sel);
  const docXml = async (bytes, name = 'word/document.xml') => (await JSZip.loadAsync(Buffer.from(bytes))).file(name).async('string');
  const wellFormed = xml => { try { execFileSync('python3', ['-I', '-c', 'import sys, xml.dom.minidom as m; m.parseString(sys.stdin.buffer.read())'], { input: xml }); return true; } catch { return false; } };

  // ---- Form fields
  await p.evaluate(() => { document.getElementById('doc').innerHTML = '<p id="f1">Name: </p><p id="f2">Agree: </p><p id="f3">Team: </p><p id="f4">Start: </p>'; });
  const insert = async (sel, fields) => { await caretAtEnd(sel); await answer([fields]); await p.click('[data-tab="insert"]'); await p.click('#insFormField'); };
  await insert('#f1', { kind: 'text', title: 'Full name', ph: 'Your name', opts: '' });
  await insert('#f2', { kind: 'check', title: 'I agree', ph: '', opts: '' });
  await insert('#f3', { kind: 'drop', title: 'Team', ph: 'Pick a team', opts: 'Red\nBlue\nGreen' });
  await insert('#f4', { kind: 'date', title: 'Start date', ph: '', opts: '' });
  check('four kinds of form field go in', (await p.locator('#doc span.ff').count()) === 4 && (await p.textContent('#f1 span.ff')) === 'Your name');
  await p.click('#f2 span.ff');
  check('a check box ticks', (await p.textContent('#f2 span.ff')) === '☒');
  await p.click('#f1 span.ff');
  await p.fill('.ff-pop input', 'Ada Lovelace');
  await p.keyboard.press('Enter');
  await p.click('#f3 span.ff');
  await p.selectOption('.ff-pop select', 'Blue');
  await p.click('#f4 span.ff');
  await p.fill('.ff-pop input', '2026-11-02');
  await p.keyboard.press('Enter');
  const vals = await p.evaluate(() => [...document.querySelectorAll('#doc span.ff')].map(f => f.dataset.v || ''));
  check('fields are filled in', vals.join('|') === 'Ada Lovelace|1|Blue|2026-11-02', vals.join('|'));
  // Restrict to filling in.
  await p.click('#formsOnly');
  check('Restrict to Filling In leaves only the fields editable', await p.evaluate(() => document.getElementById('doc').getAttribute('contenteditable') === 'false' && document.body.classList.contains('forms-only')));
  let bytes = await p.evaluate(async () => Array.from(await AtlasWrite.serialize('docx')));
  let xml = await docXml(bytes);
  check('in Word they are content controls: text, check box, drop-down, date', /<w:sdt>[\s\S]*Ada Lovelace/.test(xml) && /<w14:checked w14:val="1"\/>/.test(xml) && /<w:listItem w:displayText="Blue"/.test(xml) && /w:fullDate="2026-11-02/.test(xml) && wellFormed(xml));
  const settings = await docXml(bytes, 'word/settings.xml');
  check('and the file is protected for filling in forms', /<w:documentProtection w:edit="forms" w:enforcement="1"\/>/.test(settings));
  fs.writeFileSync(path.join(W, 'form.docx'), Buffer.from(bytes));
  const py = execFileSync('python3', ['-I', '-c', 'import sys, docx; d = docx.Document(sys.argv[1]); print(len(d.paragraphs))', path.join(W, 'form.docx')]).toString().trim();
  check('python-docx opens the file', +py >= 4, py);
  await p.evaluate(async b => AtlasWrite.loadFile({ name: 'form.docx', path: null, data: new Uint8Array(b) }), bytes);
  const back = await p.evaluate(() => ({ v: [...document.querySelectorAll('#doc span.ff')].map(f => `${f.dataset.ff}:${f.dataset.v || ''}`), title: document.querySelector('#doc span.ff').title, opts: document.querySelectorAll('#doc span.ff')[2].dataset.opts, only: AtlasWrite.setup.formsOnly }));
  check('reopened, the fields and their values come back', back.v.join('|') === 'text:Ada Lovelace|check:1|drop:Blue|date:2026-11-02' && back.title === 'Full name' && back.opts === '["Red","Blue","Green"]' && back.only, JSON.stringify(back));
  await p.click('#formsOnly');

  // ---- Watermark, line numbers and a drop cap
  await p.evaluate(() => { document.getElementById('doc').innerHTML = '<h1>Report</h1><p id="d1">Once upon a time there was a document with a big first letter.</p><p>Second paragraph.</p><p>Third.</p>'; AtlasWrite.loadFile; });
  await answer([{ preset: 'DRAFT', text: '', color: '#c0c0c0', layout: 'd' }]);
  await p.click('[data-tab="layout"]');
  await p.click('#wmBtn');
  await p.waitForSelector('.page-bg .wm');
  check('the watermark shows on the page', (await p.textContent('.page-bg .wm')) === 'DRAFT');
  await p.selectOption('#lineNumbers', 'page');
  await p.waitForTimeout(500);
  const lns = await p.evaluate(() => [...document.querySelectorAll('.page-bg .ln')].map(n => n.textContent));
  check('line numbers count the lines', lns.length >= 4 && lns[0] === '1' && lns.every((n, i) => +n === i + 1), lns.join(','));
  await caretAtEnd('#d1');
  await p.click('[data-tab="insert"]');
  await p.click('#dropCapBtn');
  check('Drop Cap marks the paragraph', await p.evaluate(() => document.getElementById('d1').classList.contains('dropcap')));
  bytes = await p.evaluate(async () => Array.from(await AtlasWrite.serialize('docx')));
  xml = await docXml(bytes);
  const zip = await JSZip.loadAsync(Buffer.from(bytes));
  const headers = await Promise.all(Object.keys(zip.files).filter(n => /^word\/header\d*\.xml$/.test(n)).map(n => zip.file(n).async('string')));
  check('in Word: a watermark shape in the header', headers.some(h => /<v:textpath\b[^>]*string="DRAFT"/.test(h) && wellFormed(h)));
  check('in Word: line numbers restart each page', /<w:lnNumType\b[^>]*w:restart="newPage"/.test(xml));
  check('in Word: a dropped capital letter', /<w:framePr w:dropCap="drop" w:lines="3"[^>]*\/>[\s\S]*?<w:t>O<\/w:t>/.test(xml) && wellFormed(xml));
  await p.evaluate(async b => AtlasWrite.loadFile({ name: 'report.docx', path: null, data: new Uint8Array(b) }), bytes);
  const st = await p.evaluate(() => ({ wm: AtlasWrite.setup.watermark, ln: AtlasWrite.setup.lineNumbers, dc: [...document.querySelectorAll('#doc p.dropcap')].map(x => x.textContent) }));
  check('reopened: watermark, line numbers and drop cap come back', st.wm && st.wm.text === 'DRAFT' && st.wm.diagonal && st.ln === 'page' && st.dc[0] === 'Once upon a time there was a document with a big first letter.', JSON.stringify(st));

  // ---- Outline
  await p.evaluate(() => { document.getElementById('doc').innerHTML = '<h1 id="ha">Alpha</h1><p>a text</p><h1 id="hb">Beta</h1><p>b text</p><h2 id="hc">Gamma</h2><p>c text</p>'; });
  await caretAtEnd('#hb');
  await p.keyboard.press('Alt+Shift+ArrowUp');
  let order = await p.evaluate(() => [...document.getElementById('doc').children].map(n => n.textContent).join('|'));
  check('Alt+Shift+Up moves a heading with its section', order === 'Beta|b text|Gamma|c text|Alpha|a text', order);
  await caretAtEnd('#hb');
  await p.keyboard.press('Alt+Shift+ArrowRight');
  check('Alt+Shift+Right demotes it and the headings under it', await p.evaluate(() => [...document.querySelectorAll('#doc h2, #doc h3')].map(h => h.tagName + h.textContent).join(',') === 'H2Beta,H3Gamma'));
  await p.click('[data-tab="view"]');
  await p.click('#viewOutline');
  await p.selectOption('#outlineLevel', '2');
  const visible = await p.evaluate(() => [...document.getElementById('doc').children].filter(n => getComputedStyle(n).display !== 'none').map(n => n.textContent).join('|'));
  check('Outline view shows the chosen levels', visible === 'Beta|Alpha', visible);
  await p.click('#viewOutline');
  const saved = await p.evaluate(async () => { const h = await AtlasWrite.serialize('html'); return typeof h === 'string' ? h : new TextDecoder().decode(h); });
  check('outline marks are never saved', !/ol-hide|data-ol/.test(saved));

  // ---- Envelopes and labels
  await p.evaluate(() => { window.__opened = []; Platform.openData = async (app, name, bytes) => { window.__opened.push({ app, name, html: new TextDecoder().decode(bytes) }); return true; }; });
  await answer([{ to: 'Grace Hopper\n1 Navy Way\nArlington', from: 'Ada Lovelace\n12 St James Sq', size: 'env10' }]);
  await p.click('[data-tab="mailings"]');
  await p.click('#envBtn');
  await p.waitForFunction(() => window.__opened.length === 1);
  const env = await p.evaluate(() => window.__opened[0]);
  check('Envelopes opens an envelope-sized document with both addresses', env.app === 'write' && /Grace Hopper<br>1 Navy Way/.test(env.html) && /Ada Lovelace/.test(env.html) && /&quot;size&quot;:&quot;env10&quot;/.test(env.html));
  await answer([{ kind: 'avery5160', text: 'Ada Lovelace\nLondon' }]);
  await p.click('#labelsBtn');
  await p.waitForFunction(() => window.__opened.length === 2);
  const lab = await p.evaluate(() => { const d = new DOMParser().parseFromString(window.__opened[1].html, 'text/html'); return { cells: [...d.querySelectorAll('td')].filter(t => /Ada Lovelace/.test(t.textContent)).length, setup: JSON.parse(d.querySelector('meta[name=atlas-setup]').content) }; });
  check('Labels makes a sheet of 30 Avery 5160 labels', lab.cells === 30 && lab.setup.margin === '0in', JSON.stringify(lab));
  // The labels document opens with its page set up.
  await p.evaluate(async h => AtlasWrite.loadFile({ name: 'Labels.html', path: null, data: new TextEncoder().encode(h) }), (await p.evaluate(() => window.__opened[1].html)));
  check('which opens with no margins', await p.evaluate(() => AtlasWrite.setup.margin === '0in' && document.querySelectorAll('#doc td').length >= 30));

  // ---- Readability
  await p.evaluate(() => { document.getElementById('doc').innerHTML = '<p>The cat sat. The dog ran fast.</p><p>This sentence is written to be considerably longer than the others so that it crosses the line of twenty five words and gets highlighted as a long sentence by the tool.</p>'; });
  const rd = await p.evaluate(() => AtlasWrite.tools.stats());
  check('readability: words, sentences, long sentences, a reading-ease score', rd.words === 38 && rd.sentences === 3 && rd.long === 1 && rd.ease > 0 && rd.grade > 0, JSON.stringify(rd));
  await p.evaluate(() => AtlasWrite.tools.markHard(true));
  check('long sentences are highlighted without changing the document', await p.evaluate(() => CSS.highlights.get('rd-long').size === 1 && !document.getElementById('doc').innerHTML.includes('mark')));
  await p.evaluate(() => AtlasWrite.tools.markHard(false));

  // ---- Translation
  await p.evaluate(() => { document.getElementById('doc').innerHTML = '<h1>Results</h1><p id="t1">Hello <b>world</b></p><p id="t2">Good morning</p>'; });
  await p.evaluate(() => { const t = document.querySelector('#t2').firstChild; const r = document.createRange(); r.setStart(t, 0); r.setEnd(t, t.length); const s = getSelection(); s.removeAllRanges(); s.addRange(r); });
  await answer([{ lang: 'Spanish', what: 'sel' }]);
  await p.click('[data-tab="review"]');
  await p.click('#translateBtn');
  await p.waitForFunction(() => document.getElementById('t2').textContent === 'Buenos días', null, { timeout: 8000 });
  check('Translate replaces the selected text', true);
  await answer([{ lang: 'Spanish', what: 'doc' }]);
  await p.evaluate(() => getSelection().removeAllRanges());
  await p.click('#translateBtn');
  await p.waitForFunction(() => window.__opened.length === 3, null, { timeout: 8000 });
  const trd = await p.evaluate(() => window.__opened[2]);
  check('Translate the document opens a translated copy with its formatting', /<h1>Resultados<\/h1>/.test(trd.html) && /Hola <b>mundo<\/b>/.test(trd.html) && /\(Spanish\)/.test(trd.name), trd.html.slice(0, 300));

  check('no page errors', errors.length === 0, errors.join(' | '));
  await b.close();
  web.close(); api.close();
})().catch(e => { console.log('FAIL crashed :: ' + e.message); process.exit(1); });
