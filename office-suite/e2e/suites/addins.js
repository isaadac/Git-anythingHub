// Add-ins (installed from a file, buttons on the Insert tab, a pane, reading and changing the
// document within their permissions, sandboxed: no network or page access, stopped when they run
// away) and the Assistant reading files added to a question.
const { chromium } = require('playwright');
const fs = require('fs');
const http = require('http');
const path = require('path');
const { createWebServer } = require(path.join(process.cwd(), 'src/server/web-server.js'));
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sample = f => fs.readFileSync(path.join(process.cwd(), 'addins', f), 'utf8');

(async () => {
  // The Assistant's stand-in: says which attached files it was given.
  const asked = [];
  const api = http.createServer(async (req, res) => {
    let b = ''; for await (const c of req) b += c;
    const body = JSON.parse(b || '{}');
    const last = body.messages[body.messages.length - 1].content;
    asked.push({ system: body.system, last });
    const names = [...String(last).matchAll(/<attached_file name="([^"]+)" kind="([^"]+)">/g)].map(m => `${m[1]} (${m[2]})`);
    const reply = 'Files: ' + names.join(', ');
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
  const BASE = `http://127.0.0.1:${web.address().port}/src/renderer/`;

  const b = await chromium.launch();
  const ctx = await b.newContext({ viewport: { width: 1400, height: 900 } });
  const outside = [];
  ctx.on('request', r => { if (!r.url().startsWith(`http://127.0.0.1:${web.address().port}/`) && !/^(data|blob):/.test(r.url())) outside.push(r.url()); });
  const errors = [];
  const page = async app => {
    const p = await ctx.newPage();
    p.on('pageerror', e => errors.push(app + ': ' + e.message));
    await p.goto(BASE + app + '.html');
    await p.waitForFunction(() => window.Addins && document.getElementById('addinsBtn'));
    await p.evaluate(() => { window.__chooses = []; window.__forms = []; UI.choose = async () => window.__chooses.shift() ?? -1; UI.form = async (t, fields) => { const want = window.__forms.shift(); if (want === null) return null; const out = {}; fields.forEach(f => { out[f.name] = f.value !== undefined ? f.value : (f.options ? (f.options[0].value ?? f.options[0]) : ''); }); return Object.assign(out, want || {}); }; });
    return p;
  };
  const install = (p, text) => p.evaluate(t => { window.__chooses = [0]; return (window.AtlasWrite || window.AtlasSheets || window.AtlasSlides).addins.install(t).then(a => !!a); }, text);

  // ---- Write: the Word Goal sample (read only)
  const w = await page('write');
  await w.click('#doc');
  await w.keyboard.type('One two three four five six seven eight nine ten.');
  check('the Add-ins button is on the Insert tab', await w.evaluate(() => !!document.querySelector('.ribbon-panel[data-panel="insert"] #addinsBtn')));
  check('a sample add-in installs (after showing what it asks for)', await install(w, sample('word-goal.isadek-addin')));
  check('…its button appears on the Insert tab', (await w.locator('.addin-cmd[data-cmd="show"]').count()) === 1);
  await w.click('[data-tab="insert"]');
  await w.click('.addin-cmd[data-cmd="show"]');
  await w.waitForSelector('.addin-pane:not([hidden]) progress', { timeout: 10000 });
  check('running it shows its pane, reading the document', /10 of 500 words/.test(await w.textContent('.addin-pane .body')) && /Add-in: Word Goal/.test(await w.textContent('.addin-pane .by')));
  await w.fill('.addin-pane input', '20');
  await w.click('.addin-pane button:has-text("Set Goal")');
  await w.waitForFunction(() => /10 of 20 words/.test(document.querySelector('.addin-pane .body').textContent));
  check('…its pane’s inputs and buttons work', /10 to go/.test(await w.textContent('.addin-pane .body')));

  // An add-in that asks to read only can't change the document.
  const sneaky = JSON.stringify({ id: 'test.sneaky', name: 'Sneaky', version: '1', apps: ['write'], permissions: ['read'], commands: [{ id: 'go', label: 'Go' }],
    script: `isadek.onCommand('go', async () => {
      const r = {};
      try { await isadek.document.insertText('HACKED'); r.insert = 'allowed'; } catch (e) { r.insert = e.message; }
      r.fetch = typeof fetch; r.xhr = typeof XMLHttpRequest; r.ws = typeof WebSocket; r.idb = typeof indexedDB; r.doc = typeof document; r.win = typeof window; r.imp = typeof importScripts;
      await isadek.ui.showPane('Report', [{ type: 'text', text: JSON.stringify(r) }, { type: 'text', text: '<img src=x onerror="window.__xss=1">' }]);
    });` });
  check('a second add-in installs', await install(w, sneaky));
  await w.click('.addin-cmd[data-cmd="go"]');
  await w.waitForFunction(() => /"insert"/.test(document.querySelector('.addin-pane .body').textContent), null, { timeout: 10000 });
  const rep = JSON.parse(await w.locator('.addin-pane .body p').first().textContent());
  check('an add-in without the “write” permission can’t change the document', /may not change the document/.test(rep.insert) && !/HACKED/.test(await w.textContent('#doc')), rep.insert);
  check('…and has no network, storage or page access', rep.fetch === 'undefined' && rep.xhr === 'undefined' && rep.ws === 'undefined' && rep.idb === 'undefined' && rep.doc === 'undefined' && rep.win === 'undefined' && rep.imp === 'undefined', JSON.stringify(rep));
  check('…and what it shows is text, never HTML', !(await w.evaluate(() => !!document.querySelector('.addin-pane img') || window.__xss)));
  // A runaway add-in is stopped; the page keeps working.
  const loop = JSON.stringify({ id: 'test.loop', name: 'Loop', apps: ['write'], permissions: [], commands: [{ id: 'spin', label: 'Spin' }], script: `isadek.onCommand('spin', () => { for (;;) {} });` });
  await install(w, loop);
  await w.evaluate(() => { const i = AtlasWrite.addins; window.__runP = i.run('test.loop', 'command', 'spin'); });
  await w.waitForTimeout(300);
  check('a running add-in doesn’t freeze the page', (await w.evaluate(() => 1 + 1)) === 2 && (await w.evaluate(() => AtlasWrite.addins.running.includes('test.loop'))));
  await w.evaluate(() => AtlasWrite.addins.stop('test.loop'));
  check('…and can be stopped', !(await w.evaluate(() => AtlasWrite.addins.running.includes('test.loop'))));
  // Broken add-in files are refused with a reason.
  const bad = await w.evaluate(() => ['not json', '{"id":"x","name":"X"}', JSON.stringify({ id: 'a.b.c', name: 'N', apps: ['write'], commands: [{ id: 'c', label: 'C' }], script: '' }), JSON.stringify({ id: 'a.b.c', name: 'N', apps: ['mail'], commands: [{ id: 'c', label: 'C' }], script: 'x' })].map(t => { try { Addins.check(t); return 'ok'; } catch (e) { return e.message; } }));
  check('broken add-in files are refused, saying why', bad[0].includes('JSON') && /valid id|no script/.test(bad[1]) && /no script/.test(bad[2]) && /which apps/.test(bad[3]), bad.join(' | '));
  // Turning one off removes its button.
  await w.evaluate(() => { window.__chooses = [1]; window.__forms = [{ id: 'test.sneaky' }]; });
  await w.click('#addinsBtn');
  await w.waitForFunction(() => !document.querySelector('.addin-cmd[data-cmd="go"]'));
  check('turning an add-in off removes its button', (await w.locator('.addin-cmd[data-cmd="show"]').count()) === 1);

  // ---- Sheets: the Selection Statistics sample (read and write)
  const s = await page('sheets');
  await s.evaluate(() => { AtlasSheets.setCell('A1', 4); AtlasSheets.setCell('A2', 8); AtlasSheets.setCell('A3', 15); AtlasSheets.setCell('A4', 'note'); AtlasSheets.select(0, 0); AtlasSheets.select(3, 0, true); });
  check('installed add-ins show in every app they work in (only there)', (await s.locator('.addin-cmd').count()) === 0);
  await install(s, sample('selection-stats.isadek-addin'));
  await s.click('[data-tab="insert"]');
  await s.click('.addin-cmd[data-cmd="stats"]');
  await s.waitForSelector('.addin-pane:not([hidden]) li', { timeout: 10000 });
  const items = await s.$$eval('.addin-pane li', l => l.map(x => x.textContent));
  check('Selection Statistics reads the selected cells', items.includes('Count: 3') && items.includes('Sum: 27') && items.includes('Average: 9') && items.includes('Median: 8'), items.join(' | '));
  await s.click('.addin-pane button:has-text("Put in Sheet")');
  await s.waitForFunction(() => AtlasSheets.valueOf('D2') === 27, null, { timeout: 10000 }).catch(() => {});
  check('…and writes them into the sheet (it has the write permission)', (await s.evaluate(() => [AtlasSheets.valueOf('C1'), AtlasSheets.valueOf('D1'), AtlasSheets.valueOf('C2'), AtlasSheets.valueOf('D2')])).join() === 'Count,3,Sum,27');
  await s.click('.grid tbody tr[data-r="6"] td[data-c="0"]');
  await s.keyboard.press('Control+z');
  check('…as one step that Undo takes back', (await s.evaluate(() => AtlasSheets.valueOf('D2'))) !== 27);

  // ---- The Assistant reads added files
  const docx = Buffer.from(await w.evaluate(async () => {
    const z = new JSZip();
    z.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
    z.file('word/document.xml', '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Budget 2027</w:t></w:r></w:p><w:p><w:r><w:t>Travel: 4000</w:t></w:r></w:p></w:body></w:document>');
    return Array.from(await z.generateAsync({ type: 'uint8array' }));
  }));
  const extracted = await w.evaluate(async d => [await FileText.extract('budget.docx', new Uint8Array(d)), await FileText.extract('notes.md', new TextEncoder().encode('# Plan\nShip it')), await FileText.extract('x.exe', new Uint8Array(3)).catch(e => e.message)], Array.from(docx));
  check('files are read as text: a Word document keeps its headings', extracted[0].text === '# Budget 2027\nTravel: 4000' && extracted[0].kind === 'Word document', JSON.stringify(extracted[0]));
  check('…Markdown as it is; other kinds are refused', extracted[1].text === '# Plan\nShip it' && /can’t be read/.test(extracted[2]));
  await w.click('.ai-btn');
  await w.evaluate(async d => {
    // The pane's Add files button opens a file: answer it with the document.
    Platform.open = async () => ({ name: 'budget.docx', data: new Uint8Array(d) });
  }, Array.from(docx));
  await w.click('.ai-pane .ai-attach');
  await w.waitForSelector('.ai-pane .ai-file');
  check('Add files puts a chip in the Assistant pane', (await w.textContent('.ai-pane .ai-file .nm')) === 'budget.docx');
  await w.fill('.ai-pane textarea', 'What is the travel budget?');
  await w.click('.ai-pane .send');
  await w.waitForFunction(() => /Files:/.test(document.querySelector('.ai-log').textContent), null, { timeout: 15000 });
  check('the question is sent with the file’s text', /Files: budget\.docx \(Word document\)/.test(await w.textContent('.ai-log')) && /<attached_file name="budget\.docx" kind="Word document">\n# Budget 2027\nTravel: 4000\n<\/attached_file>/.test(asked[asked.length - 1].last) && /<attached_file> tags/.test(asked[asked.length - 1].system));
  await w.click('.ai-pane .ai-file .rm');
  check('…and the chip can be removed', (await w.locator('.ai-pane .ai-file').count()) === 0);

  check('nothing was fetched from outside', !outside.length, outside.join(' '));
  check('no page errors', !errors.length, errors.join(' | '));
  await b.close();
  web.close(); api.close();
})().catch(e => { console.error('CRASH', e.stack); process.exit(1); });
