// The Assistant in a browser, through the Isadek Office web server (which holds the API key),
// and the web server itself: only the app's own files, and live editing on the same port.
const { chromium } = require('playwright');
const http = require('http');
const path = require('path');
const { createWebServer } = require(path.join(process.cwd(), 'src/server/web-server.js'));
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
const KEY = 'sk-ant-test-key-1234567890';
(async () => {
  // A stand-in for the Anthropic API.
  const reqs = [];
  let reply = 'Hello', pace = 15;
  const api = http.createServer(async (req, res) => {
    let b = ''; for await (const c of req) b += c;
    const body = JSON.parse(b || '{}');
    reqs.push({ headers: req.headers, body });
    if ((req.headers['x-api-key'] || String(req.headers.authorization).replace(/^Bearer /, '')) !== KEY) { res.writeHead(401, { 'content-type': 'application/json' }); return res.end(JSON.stringify({ type: 'error', error: { type: 'authentication_error', message: 'invalid x-api-key' } })); }
    res.writeHead(200, { 'content-type': 'text/event-stream' });
    const ev = (t, d) => res.write(`event: ${t}\ndata: ${JSON.stringify({ type: t, ...d })}\n\n`);
    ev('message_start', { message: { id: 'msg_1', type: 'message', role: 'assistant', model: body.model, content: [], stop_reason: null, usage: { input_tokens: 10, output_tokens: 0 } } });
    ev('content_block_start', { index: 0, content_block: { type: 'text', text: '' } });
    for (const part of reply.match(/[\s\S]{1,7}/g)) { if (res.destroyed) return; ev('content_block_delta', { index: 0, delta: { type: 'text_delta', text: part } }); await sleep(pace); }
    ev('content_block_stop', { index: 0 });
    ev('message_delta', { delta: { stop_reason: 'end_turn' }, usage: { output_tokens: 5 } });
    ev('message_stop', {});
    res.end();
  });
  await new Promise(r => api.listen(0, '127.0.0.1', r));
  const web = createWebServer({ apiKey: KEY, codes: ['letmein'], perMinute: 6, baseURL: `http://127.0.0.1:${api.address().port}`, env: {} });
  await new Promise(r => web.listen(0, '127.0.0.1', r));
  const BASE = `http://127.0.0.1:${web.address().port}`;

  // The server's files: the app, and nothing else.
  const get = async p => (await fetch(BASE + p)).status;
  check('web server: start page redirects to Home', (await fetch(BASE + '/', { redirect: 'manual' })).headers.get('location') === '/src/renderer/index.html');
  check('web server: app pages served', await get('/src/renderer/write.html') === 200);
  check('web server: libraries the pages use are served', await get('/node_modules/xlsx/dist/xlsx.full.min.js') === 200 || await get('/node_modules/jszip/dist/jszip.min.js') === 200);
  const hidden = await Promise.all(['/src/main/assistant.js', '/package.json', '/node_modules/@anthropic-ai/sdk/package.json', '/src/renderer/../../package.json', '/%2e%2e/package.json', '/src/server/web-server.js'].map(get));
  check('web server: nothing else (server code, package files, other libraries)', hidden.every(s => s === 404), hidden.join(','));

  const browser = await chromium.launch();
  const p = await browser.newPage();
  const errors = [];
  p.on('pageerror', e => errors.push(e.message));
  await p.goto(BASE + '/src/renderer/write.html');
  await p.waitForSelector('.ai-btn');
  await p.click('.ai-btn');
  await p.waitForFunction(() => document.querySelector('.ai-note').textContent);
  check('without an access code the pane says what is needed', /access code/i.test(await p.textContent('.ai-note')), await p.textContent('.ai-note'));
  const denied = await p.evaluate(async () => (await fetch('/api/assistant', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ messages: [{ role: 'user', content: 'hi' }] }) })).status);
  check('the server refuses questions without a valid code', denied === 401 && reqs.length === 0);
  // Settings: enter the code.
  await p.click('.ai-pane .gear');
  await p.waitForSelector('.modal input[type=password]');
  await p.fill('.modal input[type=password]', 'letmein');
  await p.click('.modal .actions .primary');
  await p.waitForFunction(() => /Opus/.test(document.querySelector('.ai-note').textContent), null, { timeout: 5000 }).catch(() => {});
  check('with the code the pane is ready (model shown)', /Opus/.test(await p.textContent('.ai-note')), await p.textContent('.ai-note'));
  // Ask; the reply streams in.
  await p.click('#doc'); await p.keyboard.type('A short document about the sea.');
  reply = 'The sea covers most of the planet.';
  await p.fill('.ai-input textarea', 'Summarise this');
  await p.click('.ai-input .send');
  await p.waitForSelector('.ai-msg.assistant .ai-use', { timeout: 10000 });
  check('the reply arrives in the pane', (await p.textContent('.ai-msg.assistant .body')) === reply);
  const q = reqs[reqs.length - 1];
  check('the server asks with its key, the chosen model and adaptive thinking', q.body.model === 'claude-opus-5-5' && q.body.thinking.type === 'adaptive' && q.body.stream === true);
  check('the document goes with the question', /<document>[\s\S]*about the sea/.test(q.body.messages[0].content));
  check('the page never sees the key', await p.evaluate(async k => !document.documentElement.innerHTML.includes(k) && !JSON.stringify(await (await fetch('/api/assistant/status', { headers: { Authorization: 'Bearer letmein' } })).json()).includes(k), KEY));
  // Stop a long reply.
  reply = 'word '.repeat(400); pace = 40;
  await p.fill('.ai-input textarea', 'Write a lot');
  await p.click('.ai-input .send');
  await p.waitForFunction(() => document.querySelectorAll('.ai-msg.assistant')[1]?.querySelector('.body').textContent.length > 10, null, { timeout: 8000 });
  await p.click('.ai-input .send'); // Stop
  await p.waitForFunction(() => document.querySelector('.ai-input .send').textContent === 'Send', null, { timeout: 5000 });
  const stopped = await p.evaluate(() => document.querySelectorAll('.ai-msg.assistant')[1].querySelector('.body').textContent);
  check('Stop ends the reply part way', stopped.length > 0 && stopped.length < 2000 && !(await p.evaluate(() => !!document.querySelectorAll('.ai-msg.assistant')[1].querySelector('.ai-use'))), stopped.length);
  pace = 1;
  // Questions per minute are limited.
  const codes = await p.evaluate(async () => { const out = []; for (let i = 0; i < 6; i++) out.push((await fetch('/api/assistant', { method: 'POST', headers: { 'Content-Type': 'application/json', Authorization: 'Bearer letmein' }, body: JSON.stringify({ messages: [{ role: 'user', content: 'hi' }] }) })).status); return out; });
  check('too many questions in a minute are turned away', codes.includes(429), codes.join(','));
  // Live editing on the same port.
  const live = await p.evaluate(() => new Promise(res => {
    const ws = new WebSocket(location.origin.replace(/^http/, 'ws') + '/');
    ws.onopen = () => ws.send(JSON.stringify({ t: 'hello', room: 'webserverroom0001', create: true, app: 'write', name: 'T', color: '#123456', init: {} }));
    ws.onmessage = e => { res(JSON.parse(e.data).t); ws.close(); };
    ws.onerror = () => res('error');
    setTimeout(() => res('timeout'), 4000);
  }));
  check('live editing works on the same server', live === 'welcome', live);
  check('no page errors', !errors.length, errors.join(' | '));
  await browser.close();
  web.close(); api.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
