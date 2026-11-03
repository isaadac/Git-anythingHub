'use strict';
// The servers keep running whatever they are sent: malformed requests, hostile live-editing
// messages, floods and idle sockets.
const test = require('node:test');
const assert = require('node:assert');
const http = require('http');
const net = require('net');
const { createCollabServer } = require('../src/main/collab-server');
const { createWebServer } = require('../src/server/web-server');

const listen = srv => new Promise(r => srv.listen(0, '127.0.0.1', () => r(srv.address().port)));
const get = (port, p, headers = {}) => new Promise((resolve, reject) => {
  // A raw request, so malformed paths reach the server as typed.
  const s = net.connect(port, '127.0.0.1', () => s.write(`GET ${p} HTTP/1.1\r\nHost: x\r\nConnection: close\r\n${Object.entries(headers).map(([k, v]) => `${k}: ${v}\r\n`).join('')}\r\n`));
  let out = '';
  s.on('data', d => { out += d; });
  s.on('end', () => resolve({ status: +(out.match(/^HTTP\/1\.1 (\d+)/) || [])[1], head: out.split('\r\n\r\n')[0], body: out.split('\r\n\r\n').slice(1).join('\r\n\r\n') }));
  s.on('error', reject);
});

test('web server: bad addresses are refused without taking it down', async t => {
  const srv = createWebServer({ live: false, apiKey: '' });
  const port = await listen(srv);
  t.after(() => srv.close());
  assert.strictEqual((await get(port, '/src/renderer/%E0%A4%A')).status, 400);
  assert.strictEqual((await get(port, '/src/renderer/index.html%00.js')).status, 400);
  assert.strictEqual((await get(port, '/src/renderer/../../package.json')).status, 404);
  assert.strictEqual((await get(port, '/node_modules/electron/package.json')).status, 404);
  const ok = await get(port, '/src/renderer/index.html');
  assert.strictEqual(ok.status, 200, 'still serving');
  for (const h of ['x-frame-options: DENY', "content-security-policy: frame-ancestors 'none'", 'x-content-type-options: nosniff', 'referrer-policy: no-referrer']) assert.match(ok.head.toLowerCase(), new RegExp(h.toLowerCase().replace(/[()']/g, '.')), h);
});

test('web server: access codes can\'t be guessed', async t => {
  const fake = { beta: { messages: { stream() { throw new Error('not used'); } } } };
  const srv = createWebServer({ live: false, client: fake, codes: ['correct-horse'], perMinute: 100 });
  const port = await listen(srv);
  t.after(() => srv.close());
  const status = code => new Promise((resolve, reject) => http.get({ port, host: '127.0.0.1', path: '/api/assistant/status', headers: { Authorization: 'Bearer ' + code } }, r => {
    let b = ''; r.on('data', d => { b += d; }); r.on('end', () => resolve({ status: r.statusCode, body: JSON.parse(b) }));
  }).on('error', reject));
  assert.strictEqual((await status('correct-horse')).body.codeOk, true);
  for (let i = 0; i < 20; i++) assert.strictEqual((await status('guess' + i)).body.codeOk, false);
  const r = await status('correct-horse');
  assert.strictEqual(r.status, 429, 'locked out after many wrong codes, even with the right one');
});

function ws(url) {
  const w = new WebSocket(url);
  const got = [];
  w.onmessage = e => got.push(JSON.parse(e.data));
  return { w, got, open: new Promise(r => { w.onopen = r; }), closed: new Promise(r => { w.onclose = e => r(e.code); }), send: m => w.send(typeof m === 'string' ? m : JSON.stringify(m)) };
}
const sleep = ms => new Promise(r => setTimeout(r, ms));

test('live server: hostile messages close the sender, not the server', async t => {
  const srv = createCollabServer({ limits: { helloMs: 300, rate: 2e6, connections: 5 } });
  const port = await listen(srv);
  t.after(() => { srv.close(); for (const rm of srv.rooms.values()) for (const c of rm.clients) c.socket.destroy(); });
  const url = `ws://127.0.0.1:${port}/`, room = 'h'.repeat(20);
  // A value nested so deeply that turning it back into text overflows the stack.
  const a = ws(url); await a.open;
  a.send({ t: 'hello', room, create: true, app: 'write', name: 'A' });
  await sleep(100);
  a.send('{"t":"op","ops":[{"k":"p:1","v":' + '['.repeat(200000) + ']'.repeat(200000) + '}]}');
  const code = await a.closed;
  assert.ok(code === 1011 || code === 1007, 'sender let go: ' + code);
  // The server is still there.
  const b = ws(url); await b.open;
  b.send({ t: 'hello', room, app: 'write', name: 'B' });
  await sleep(150);
  assert.ok(b.got.some(m => m.t === 'welcome'), 'others can still join');
  // Big presence is dropped.
  b.send({ t: 'presence', data: { x: 'y'.repeat(20000) } });
  await sleep(100);
  assert.strictEqual([...srv.rooms.get(room).clients][0].presence, null);
  // A flood is cut off.
  const big = JSON.stringify({ t: 'op', ops: [{ k: 'p:2', v: 'z'.repeat(900000) }] });
  for (let i = 0; i < 4; i++) b.send(big);
  assert.strictEqual(await b.closed, 1008, 'flooding sender closed');
});

test('live server: idle sockets and too many connections', async t => {
  const srv = createCollabServer({ limits: { helloMs: 300, connections: 2 } });
  const port = await listen(srv);
  const socks = [];
  t.after(() => { socks.forEach(s => s.destroy()); srv.close(); });
  // A raw WebSocket handshake: the status line, and whether the server hung up within `ms`.
  const upgrade = ms => new Promise(resolve => {
    const s = net.connect(port, '127.0.0.1', () => s.write('GET / HTTP/1.1\r\nHost: x\r\nUpgrade: websocket\r\nConnection: Upgrade\r\nSec-WebSocket-Key: dGhlIHNhbXBsZSBub25jZQ==\r\nSec-WebSocket-Version: 13\r\n\r\n'));
    socks.push(s);
    let out = '', ended = false;
    s.on('data', d => { out += d.toString('latin1'); });
    s.on('end', () => { ended = true; });
    s.on('error', () => { ended = true; });
    setTimeout(() => resolve({ status: out.split('\r\n')[0], ended }), ms);
  });
  const idle = await upgrade(700);
  assert.match(idle.status, /101/);
  assert.ok(idle.ended, 'a socket that never joins a session is let go');
  const [a, b, c] = await Promise.all([upgrade(150), upgrade(150), upgrade(150)]);
  const statuses = [a, b, c].map(x => x.status).sort();
  assert.deepStrictEqual(statuses.map(x => /101/.test(x) ? 101 : /503/.test(x) ? 503 : 0), [101, 101, 503], statuses.join(' | '));
});
