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
    assert.match((await y.next()).error, /Isadek Write/);
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

  await t.test('links that can only comment or view, and the discussion', async () => {
    const r3 = 'e'.repeat(24), ct = 'c'.repeat(24), vt = 'v'.repeat(24);
    const ed = client(url), co = client(url), vi = client(url);
    await Promise.all([ed.opened, co.opened, vi.opened]);
    ed.send({ t: 'hello', room: r3, create: true, app: 'write', name: 'Ed', share: { comment: ct, view: vt, edit: 'x'.repeat(24) }, init: { 'b:1': { h: 'Hi' } } });
    const we = await ed.next(m => m.t === 'welcome');
    assert.strictEqual(we.role, 'edit');
    assert.deepStrictEqual(we.share, { comment: ct, view: vt });
    co.send({ t: 'hello', room: ct, app: 'write', name: 'Cy' });
    const wc = await co.next(m => m.t === 'welcome');
    assert.strictEqual(wc.role, 'comment');
    assert.strictEqual(wc.share, undefined);
    assert.deepStrictEqual(wc.state, { 'b:1': { h: 'Hi' } });
    vi.send({ t: 'hello', room: vt, app: 'write', name: 'Vi' });
    const wv = await vi.next(m => m.t === 'welcome');
    assert.strictEqual(wv.role, 'view');
    assert.deepStrictEqual(wv.peers.map(p => p.role).sort(), ['comment', 'edit', 'view']);
    // Only the edit link changes the document.
    co.send({ t: 'op', n: 1, ops: [{ k: 'b:1', v: { h: 'Hacked' } }] });
    assert.match((await co.next(m => m.t === 'error')).error, /comment on this document but not change it/);
    vi.send({ t: 'op', n: 1, ops: [{ k: 'b:2', v: { h: 'x' } }] });
    assert.match((await vi.next(m => m.t === 'error')).error, /view this document but not change it/);
    ed.send({ t: 'op', n: 1, ops: [{ k: 'b:1', v: { h: 'Hello' } }] });
    await vi.next(m => m.t === 'op');
    assert.deepStrictEqual([...srv.rooms.get(r3).state.keys()], ['b:1']);
    assert.deepStrictEqual(srv.rooms.get(r3).state.get('b:1').v, { h: 'Hello' });
    // An unknown "edit" token was not made a way in; a token can't be taken as a new session.
    assert.ok(!srv.tokens.has('x'.repeat(24)));
    const z = client(url); await z.opened;
    z.send({ t: 'hello', room: ct, create: true });
    assert.match((await z.next(m => m.t === 'error')).error, /already exists/);
    // Comments: from editors and commenters, with replies and mentions, kept for people who join later.
    co.send({ t: 'say', text: '  @Ed is this final?  ', mentions: [wc.peers.find(p => p.name === 'Ed').id, 5] });
    const said = await ed.next(m => m.t === 'say');
    assert.strictEqual(said.item.text, '@Ed is this final?');
    assert.strictEqual(said.item.from.name, 'Cy');
    assert.strictEqual(said.item.mentions.length, 1);
    ed.send({ t: 'say', text: 'Yes', reply: said.item.id });
    assert.strictEqual((await co.next(m => m.t === 'say' && m.item.text === 'Yes')).item.reply, said.item.id);
    vi.send({ t: 'say', text: 'me too' });
    assert.match((await vi.next(m => m.t === 'error')).error, /not comment/);
    const late = client(url); await late.opened;
    late.send({ t: 'hello', room: vt, name: 'Late' });
    assert.deepStrictEqual((await late.next(m => m.t === 'welcome')).talk.map(x => x.text), ['@Ed is this final?', 'Yes']);
    // Flooding is slowed down.
    for (let i = 0; i < 40; i++) co.send({ t: 'say', text: 'spam ' + i });
    assert.match((await co.next(m => m.t === 'error' && /too quickly/.test(m.error))).error, /too quickly/);
    [ed, co, vi, z, late].forEach(c => c.ws.close());
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
