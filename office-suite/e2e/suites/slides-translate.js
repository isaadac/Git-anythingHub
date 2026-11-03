// Translated live captions: what the presenter says is shown at the foot of the slides in another
// language, each phrase translated by the Assistant (through the web server; a stand-in plays
// the Anthropic API), in the order it was said.
const { chromium } = require('playwright');
const http = require('http');
const path = require('path');
const { createWebServer } = require(path.join(process.cwd(), 'src/server/web-server.js'));
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const KEY = 'sk-ant-test-key-1234567890';
const SPANISH = { 'Good morning everyone.': 'Buenos días a todos.', 'Thank you for coming.': 'Gracias por venir.' };
(async () => {
  const asked = [];
  const api = http.createServer(async (req, res) => {
    let b = ''; for await (const c of req) b += c;
    const body = JSON.parse(b || '{}');
    const text = body.messages[body.messages.length - 1].content;
    asked.push({ system: body.system, text });
    // The first phrase takes longer: the captions must still come out in order.
    await new Promise(r => setTimeout(r, text.startsWith('Good') ? 600 : 50));
    const reply = SPANISH[typeof text === 'string' ? text : text[0].text] || '?';
    res.writeHead(200, { 'content-type': 'text/event-stream' });
    const ev = (t, d) => res.write(`event: ${t}\ndata: ${JSON.stringify({ type: t, ...d })}\n\n`);
    ev('message_start', { message: { id: 'msg_1', type: 'message', role: 'assistant', model: body.model, content: [], stop_reason: null, usage: { input_tokens: 10, output_tokens: 0 } } });
    ev('content_block_start', { index: 0, content_block: { type: 'text', text: '' } });
    ev('content_block_delta', { index: 0, delta: { type: 'text_delta', text: reply } });
    ev('content_block_stop', { index: 0 });
    ev('message_delta', { delta: { stop_reason: 'end_turn' }, usage: { output_tokens: 5 } });
    ev('message_stop', {});
    res.end();
  });
  await new Promise(r => api.listen(0, '127.0.0.1', r));
  const web = createWebServer({ apiKey: KEY, open: true, baseURL: `http://127.0.0.1:${api.address().port}`, env: {}, live: false });
  await new Promise(r => web.listen(0, '127.0.0.1', r));
  const BASE = `http://127.0.0.1:${web.address().port}`;

  const browser = await chromium.launch();
  const ctx = await browser.newContext({ viewport: { width: 1280, height: 800 } });
  // A recognizer the test speaks through (the browser's engine, chosen in advance).
  await ctx.addInitScript(() => {
    localStorage.setItem('atlas-speech', 'browser');
    class FakeRecognition {
      constructor() { FakeRecognition.live = this; }
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
  });
  const p = await ctx.newPage();
  const errors = [];
  p.on('pageerror', e => errors.push(e.message));
  await p.goto(BASE + '/src/renderer/slides.html');
  await p.waitForFunction(() => window.AtlasSlides);
  await p.click('[data-tab="show"]');
  await p.selectOption('#captionsLang', 'Spanish');
  await p.click('#captionsToggle');
  await p.evaluate(() => AtlasSlides.present(0));
  await p.waitForFunction(() => window.__say('Good', false));
  check('while speaking, the words so far show faintly', (await p.textContent('#captions .cap-pending')) === 'Good');
  await p.evaluate(() => { window.__say('Good morning everyone.'); window.__say('Thank you for coming.'); });
  await p.waitForFunction(() => /Gracias por venir/.test(document.getElementById('captions').textContent), null, { timeout: 8000 });
  const cap = await p.textContent('#captions');
  check('captions show in Spanish, in the order said', cap === 'Buenos días a todos. Gracias por venir.', cap);
  check('each phrase was sent for translation into Spanish', asked.length === 2 && asked.every(a => /into Spanish/.test(a.system)) && asked[0].text === 'Good morning everyone.', JSON.stringify(asked.map(a => a.text)));
  // Without the Assistant the captions stay in the language spoken, with one note why.
  await p.keyboard.press('Escape');
  await p.evaluate(() => { localStorage.setItem('atlas-assistant-server', 'http://127.0.0.1:9'); });
  await p.evaluate(() => AtlasSlides.present(0));
  await p.waitForFunction(() => window.__say('Hello again.'));
  await p.waitForFunction(() => document.getElementById('captions').textContent === 'Hello again.', null, { timeout: 8000 });
  const toast = await p.evaluate(() => document.querySelector('.toast') && document.querySelector('.toast').textContent);
  check('without the Assistant, captions stay as spoken and say why', /can’t be translated/.test(toast || ''), toast);
  check('no page errors', errors.length === 0, errors.join(' | '));
  await browser.close();
  web.close(); api.close();
})().catch(e => { console.log('FAIL crashed :: ' + e.message); process.exit(1); });
