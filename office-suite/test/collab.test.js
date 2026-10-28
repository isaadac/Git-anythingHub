'use strict';
// Live-editing relay server: sessions, ordering, presence and limits (uses Node's WebSocket client).
const test = require('node:test');
const assert = require('node:assert');
const { createCollabServer } = require('../src/main/collab-server');

function client(url) {
  const ws = new WebSocket(url);
  const inbox = [], waiters = [];
  ws.onmessage = e => {
    const m = JSON.parse(e.data);
    const i = waiters.findIndex(w => w.pred(m));
    if (i >= 0) waiters.splice(i, 1)[0].resolve(m); else inbox.push(m);
  };
  const next = (pred = () => true) => {
    const i = inbox.findIndex(pred);
    if (i >= 0) return Promise.resolve(inbox.splice(i, 1)[0]);
    return new Promise((resolve, reject) => { waiters.push({ pred, resolve }); setTimeout(() => reject(new Error('timeout')), 3000); });
  };
  const opened = new Promise(r => { ws.onopen = r; });
  return { ws, next, opened, send: m => ws.send(JSON.stringify(m)), closed: new Promise(r => { ws.onclose = r; }) };
}

test('collab server', async t => {
  const srv = createCollabServer({ limits: { idleMs: 200, message: 200000 } });
  await new Promise(r => srv.listen(0, '127.0.0.1', r));
  const url = `ws://127.0.0.1:${srv.address().port}/`;
  const room = 'r'.repeat(20);
  t.after(() => srv.close());

  await t.test('create, join, ordered ops and presence', async () => {
    const a = client(url), b = client(url);
    await Promise.all([a.opened, b.opened]);
    a.send({ t: 'hello', room, create: true, app: 'sheets', name: 'A', color: '#112233', init: { 'c:s:A1': { v: 1 } } });
    const wa = await a.next(m => m.t === 'welcome');
    assert.deepStrictEqual(wa.state, { 'c:s:A1': { v: 1 } });
    b.send({ t: 'hello', room, app: 'sheets', name: 'B' });
    const wb = await b.next(m => m.t === 'welcome');
    assert.deepStrictEqual(wb.state, { 'c:s:A1': { v: 1 } });
    assert.strictEqual(wb.peers.length, 2);
    assert.strictEqual((await a.next(m => m.t === 'peers')).peers.length, 2);
    b.send({ t: 'op', n: 1, ops: [{ k: 'c:s:A1', v: { v: 2 } }, { k: 'c:s:B1', v: { v: 'x' } }] });
    a.send({ t: 'op', n: 1, ops: [{ k: 'c:s:A1', v: { v: 3 } }] });
    const seen = [await a.next(m => m.t === 'op'), await a.next(m => m.t === 'op')];
    const seenB = [await b.next(m => m.t === 'op'), await b.next(m => m.t === 'op')];
    // Both get the same order, and the echo of their own change.
    assert.deepStrictEqual(seen.map(m => m.seq), seenB.map(m => m.seq));
    assert.deepStrictEqual(seen.map(m => m.from), seenB.map(m => m.from));
    b.send({ t: 'presence', data: { r1: 4 } });
    assert.deepStrictEqual((await a.next(m => m.t === 'presence')).data, { r1: 4 });
    // A third person sees the last write
    const c = client(url);
    await c.opened;
    c.send({ t: 'hello', room, name: 'C' });
    const wc = await c.next(m => m.t === 'welcome');
    assert.deepStrictEqual(wc.state['c:s:A1'], seen[1].ops[0].v);
    assert.deepStrictEqual(wc.state['c:s:B1'], { v: 'x' });
    // Deleting a key
    c.send({ t: 'op', n: 1, ops: [{ k: 'c:s:B1', v: null }] });
    await c.next(m => m.t === 'op');
    assert.ok(!srv.rooms.get(room).state.has('c:s:B1'));
    [a, b, c].forEach(x => x.ws.close());
  });

  await t.test('refusals', async () => {
    const x = client(url);
    await x.opened;
    x.send({ t: 'hello', room: 'short' });
    assert.match((await x.next()).error, /Invalid/);
    x.send({ t: 'hello', room: 'z'.repeat(20) });
    assert.match((await x.next()).error, /ended|wrong/);
    x.send({ t: 'hello', room: 'w'.repeat(20), create: true, app: 'write' });
    await x.next(m => m.t === 'welcome');
    const y = client(url);
    await y.opened;
    y.send({ t: 'hello', room: 'w'.repeat(20), app: 'sheets' });
    assert.match((await y.next()).error, /Atlas Write/);
    // Oversized message closes the connection
    const z = client(url);
    await z.opened;
    z.ws.send('x'.repeat(300000));
    await z.closed;
    x.ws.close(); y.ws.close();
  });

  await t.test('empty sessions end after a while', async () => {
    const r2 = 'q'.repeat(20);
    const x = client(url);
    await x.opened;
    x.send({ t: 'hello', room: r2, create: true });
    await x.next(m => m.t === 'welcome');
    x.ws.close();
    await x.closed;
    await new Promise(r => setTimeout(r, 400));
    assert.ok(!srv.rooms.has(r2));
  });
});

test('fractional positions sort between their neighbours', () => {
  global.window = {};
  require('../src/renderer/js/collab.js');
  const { between } = global.window.Collab;
  let keys = [between('', '')];
  for (let i = 0; i < 300; i++) {
    const at = (i * 7919) % (keys.length + 1);
    const k = between(keys[at - 1] || '', keys[at] || '');
    assert.ok((at === 0 || keys[at - 1] < k) && (at === keys.length || k < keys[at]), `${keys[at - 1]} < ${k} < ${keys[at]}`);
    keys.splice(at, 0, k);
  }
  // Always appending or prepending stays short.
  let a = '';
  for (let i = 0; i < 1000; i++) a = between(a, '');
  assert.ok(a.length < 50, a.length);
  let b = 'i';
  for (let i = 0; i < 200; i++) b = between('', b);
  assert.ok(b > '' && b.length < 210);
});
