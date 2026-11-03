// Live editing under load: many people at once on the server (scripts/collab-load.js), its limits,
// a person whose connection stalls, and several app windows editing one document together.
const { chromium } = require('playwright');
const http = require('http');
const net = require('net');
const crypto = require('crypto');
const path = require('path');
const ROOT = process.cwd();
const { createCollabServer } = require(path.join(ROOT, 'src/main/collab-server.js'));
const { run } = require(path.join(ROOT, 'scripts/collab-load.js'));
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
const BASE = process.env.E2E_BASE + '/src/renderer/';

/** A connection that joins a room and then never reads (a laptop that went to sleep). */
function stalled(port, room) {
  return new Promise(resolve => {
    const s = net.connect(port, '127.0.0.1', () => {
      s.once('data', () => {
        // Joined once the server has answered the handshake; then it stops reading.
        s.write(Buffer.concat([head, mask, body]));
        setTimeout(() => { s.pause(); resolve(s); }, 100);
      });
      s.write(`GET / HTTP/1.1\r\nHost: x\r\nUpgrade: websocket\r\nConnection: Upgrade\r\nSec-WebSocket-Key: ${crypto.randomBytes(16).toString('base64')}\r\nSec-WebSocket-Version: 13\r\n\r\n`);
      const payload = Buffer.from(JSON.stringify({ t: 'hello', room, name: 'Sleepy' }));
      const mask = crypto.randomBytes(4);
      const head = payload.length < 126 ? Buffer.from([0x81, 0x80 | payload.length]) : Buffer.from([0x81, 0x80 | 126, payload.length >> 8, payload.length & 255]);
      const body = Buffer.from(payload.map((b, i) => b ^ mask[i & 3]));

    });
  });
}

(async () => {
  // ---- The server with 200 people in 20 sessions, everyone typing
  const r = await run({ rooms: 20, clients: 10, rate: 5, seconds: 8, keys: 100, size: 300, url: null });
  check(`200 people in 20 sessions: every copy ends identical to the server's (${r.delivered} changes delivered, ${r.perSecond}/s)`, !r.problems.length && r.delivered >= r.sent * 9, r.problems.slice(0, 3).join(' | '));
  check(`changes reach the others fast (p50 ${r.latencyMs.p50} ms, p95 ${r.latencyMs.p95} ms, p99 ${r.latencyMs.p99} ms)`, r.latencyMs.p95 < 500, JSON.stringify(r.latencyMs));
  const big = await run({ rooms: 2, clients: 50, rate: 2, seconds: 6, keys: 300, size: 100, url: null });
  check(`a full session (50 people) stays in step (${big.delivered} changes delivered)`, !big.problems.length && big.latencyMs.p95 < 1000, JSON.stringify(big.latencyMs) + ' ' + big.problems.slice(0, 2).join(' | '));

  // ---- Limits, and a person whose connection stalls
  const srv = createCollabServer({ server: http.createServer(), limits: { clientsPerRoom: 5, backlog: 2e6 } });
  await new Promise(res => srv.listen(0, '127.0.0.1', res));
  const port = srv.address().port, url = `ws://127.0.0.1:${port}/`;
  const room = 'limits-' + crypto.randomBytes(8).toString('hex');
  const open = (create, name) => new Promise((resolve) => {
    const ws = new WebSocket(url);
    ws.onopen = () => ws.send(JSON.stringify({ t: 'hello', room, create, name }));
    ws.onmessage = e => { const m = JSON.parse(e.data); if (m.t === 'welcome' || m.t === 'error') resolve({ ws, m }); };
  });
  const people = [await open(true, 'P0')];
  for (let i = 1; i < 4; i++) people.push(await open(false, 'P' + i));
  const sleepy = await stalled(port, room);
  for (let i = 0; i < 50 && srv.rooms.get(room).clients.size < 5; i++) await sleep(20);
  const extra = await open(false, 'P5');
  check('a full session refuses one more person, politely', extra.m.t === 'error' && /full/.test(extra.m.error), JSON.stringify(extra.m));
  extra.ws.close();
  const before = srv.rooms.get(room).clients.size;
  // Others keep typing a lot while Sleepy reads nothing.
  for (let i = 0; i < 3000; i++) {
    people[i % 4].ws.send(JSON.stringify({ t: 'op', n: i, ops: [{ k: 'p' + (i % 50), v: { text: 'y'.repeat(10000) } }] }));
    if (i % 20 === 19) await sleep(10); // the others keep reading as they go
  }
  // (30 MB in all: more than the operating system's socket buffers hold.)
  for (let i = 0; i < 50 && [...srv.rooms.get(room).clients].some(c => c.name === 'Sleepy'); i++) await sleep(100);
  const names = [...srv.rooms.get(room).clients].map(c => c.name);
  check('a stalled connection is let go instead of piling up memory', before === 5 && !names.includes('Sleepy') && names.length === 4, `${before} → ${names.join(',')}`);
  check('…and the others carry on', await new Promise(res => { people[1].ws.onmessage = e => { if (JSON.parse(e.data).t === 'op') res(true); }; people[0].ws.send(JSON.stringify({ t: 'op', n: 999, ops: [{ k: 'last', v: 'ok' }] })); setTimeout(() => res(false), 3000); }));
  sleepy.destroy();
  people.forEach(p => p.ws.close());
  await new Promise(res => { srv.close(res); for (const rm of srv.rooms.values()) for (const c of rm.clients) c.socket.destroy(); });

  // ---- Six Sheets windows and four Write windows editing one document each, all at once
  const live = createCollabServer();
  await new Promise(res => live.listen(0, '127.0.0.1', res));
  const WS = `ws://127.0.0.1:${live.address().port}/`;
  const browser = await chromium.launch();
  const errors = [];
  async function win(app, name) {
    const ctx = await browser.newContext();
    await ctx.addInitScript(n => { localStorage.setItem('atlas-user-name', n); }, name);
    const p = await ctx.newPage();
    p.on('pageerror', e => errors.push(name + ': ' + e.message));
    await p.goto(BASE + app + '.html');
    await p.waitForFunction(() => window.Collab && document.querySelector('.live-btn'));
    return p;
  }
  async function host(p) {
    await p.click('.live-btn');
    await p.click('.modal .actions button:has-text("Start on a server")');
    await p.fill('.modal input', WS);
    await p.click('.modal .actions .primary');
    await p.waitForSelector('.modal p:has-text("room=")', { timeout: 10000 });
    const text = await p.textContent('.modal p');
    await p.click('.modal .actions button:has-text("Close")');
    return text.match(/ws:\/\/\S+/)[0];
  }
  async function join(p, link) {
    await p.click('.live-btn');
    await p.click('.modal .actions button:has-text("Join")');
    await p.fill('.modal input', link);
    await p.click('.modal .actions .primary');
    await p.waitForFunction(() => document.querySelector('.live-btn').textContent.startsWith('Live ·'), null, { timeout: 15000 });
  }

  const sheets = [];
  for (let i = 0; i < 6; i++) sheets.push(await win('sheets', 'Sheet person ' + i));
  const link = await host(sheets[0]);
  for (const p of sheets.slice(1)) await join(p, link);
  await sheets[0].waitForFunction(() => document.querySelector('.live-btn').textContent === 'Live · 6', null, { timeout: 10000 }).catch(() => {});
  // Each person fills their own column and also fights over a shared row.
  await Promise.all(sheets.map((p, i) => p.evaluate(async i => {
    const col = 'ABCDEF'[i];
    for (let r = 1; r <= 40; r++) {
      AtlasSheets.setCell(col + r, i * 1000 + r);
      if (r % 5 === 0) AtlasSheets.setCell('H' + (r / 5), 'by ' + i);
      if (r % 10 === 0) await new Promise(res => setTimeout(res, 20));
    }
    AtlasSheets.setCell('G1', '=SUM(A1:F40)');
  }, i)));
  const values = p => p.evaluate(() => { const out = {}; for (const col of 'ABCDEFGH') for (let r = 1; r <= 40; r++) { const v = AtlasSheets.valueOf(col + r); if (v !== '' && v != null) out[col + r] = v; } return JSON.stringify(out); });
  let same = false, all = [];
  for (let i = 0; i < 40 && !same; i++) { await sleep(250); all = await Promise.all(sheets.map(values)); same = all.every(v => v === all[0]); }
  const sum = JSON.parse(all[0]).G1;
  const want = [0, 1, 2, 3, 4, 5].reduce((s, i) => s + 40 * i * 1000 + 820, 0);
  check('six Sheets windows typing at once end with the same workbook', same, all.map(a => a.length).join(','));
  check('…with everyone\'s cells in it (and formulas agreeing)', sum === want && Object.keys(JSON.parse(all[0])).length === 6 * 40 + 1 + 8, `${sum} vs ${want}`);

  const writers = [];
  for (let i = 0; i < 4; i++) writers.push(await win('write', 'Writer ' + i));
  await writers[0].click('#doc');
  for (let i = 0; i < 4; i++) { await writers[0].keyboard.type('Paragraph of writer ' + i); await writers[0].keyboard.press('Enter'); }
  const wlink = await host(writers[0]);
  for (const p of writers.slice(1)) await join(p, wlink);
  await sleep(800);
  // Each writer types at the end of their own paragraph, all at the same time.
  await Promise.all(writers.map(async (p, i) => {
    await p.evaluate(i => {
      const para = document.querySelectorAll('#doc > p')[i];
      const r = document.createRange(); r.selectNodeContents(para); r.collapse(false);
      const s = getSelection(); s.removeAllRanges(); s.addRange(r);
      document.getElementById('doc').focus();
    }, i);
    await p.keyboard.type(' — and more from ' + i, { delay: 15 });
  }));
  const texts = p => p.$$eval('#doc > p', ps => ps.map(x => x.textContent.replace(/\s+/g, ' ').trim()).filter(Boolean).join('\n'));
  let wsame = false, wall = [];
  for (let i = 0; i < 40 && !wsame; i++) { await sleep(250); wall = await Promise.all(writers.map(texts)); wsame = wall.every(t => t === wall[0]); }
  check('four Write windows typing at once end with the same document', wsame, wall.map(t => t.length).join(','));
  check('…each paragraph with its writer\'s words', [0, 1, 2, 3].every(i => wall[0].includes(`Paragraph of writer ${i} — and more from ${i}`)), wall[0]);

  check('no page errors', !errors.length, errors.slice(0, 3).join(' | '));
  await browser.close();
  await new Promise(res => { live.close(res); for (const rm of live.rooms.values()) for (const c of rm.clients) c.socket.destroy(); });
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
