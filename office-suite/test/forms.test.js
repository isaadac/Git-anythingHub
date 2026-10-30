'use strict';
// Isadek Forms: the shared form model, and collecting answers through the web server and the
// desktop app's local-network server.
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const C = require('../src/renderer/js/forms-core');
const { createFormsStore, createFormsServer } = require('../src/main/forms-server');
const { createWebServer } = require('../src/server/web-server');

const listen = srv => new Promise(r => srv.listen(0, '127.0.0.1', () => r(srv.address().port)));
const FORM = {
  title: 'Team lunch', description: 'Pick a day',
  questions: [
    { id: 'name', type: 'text', title: 'Name', required: true },
    { id: 'mail', type: 'email', title: 'Email' },
    { id: 'day', type: 'choice', title: 'Day', options: ['Mon', 'Tue', 'Mon', ''], required: true },
    { id: 'food', type: 'checkboxes', title: 'Food', options: ['Pizza', 'Salad', 'Soup'] },
    { id: 'n', type: 'number', title: 'Guests', min: 0, max: 5, whole: true },
    { id: 'when', type: 'date', title: 'Date' },
    { id: 'stars', type: 'rating', title: 'Rating', max: 99 }
  ]
};

test('forms are cleaned to known fields and limits', () => {
  const f = C.clean({ ...FORM, evil: '<script>', color: 'red', questions: [...FORM.questions, { type: 'html', title: 'x' }, { id: '../x', type: 'text', title: 'a'.repeat(1000), onclick: 1 }, { id: 'name', type: 'text' }] });
  assert.strictEqual(f.evil, undefined);
  assert.strictEqual(f.color, undefined);
  assert.deepStrictEqual(f.questions.find(q => q.id === 'day').options, ['Mon', 'Tue']);
  assert.strictEqual(f.questions.find(q => q.id === 'stars').max, 10);
  assert.strictEqual(f.questions.length, 9);
  assert.strictEqual(new Set(f.questions.map(q => q.id)).size, 9, 'ids stay unique');
  assert.ok(f.questions.every(q => C.ID.test(q.id) && q.title.length <= 300 && !('onclick' in q)));
  assert.strictEqual(C.clean(null), null);
  assert.strictEqual(C.clean({}).title, 'Untitled form');
});

test('answers are checked against the form', () => {
  const f = C.clean(FORM);
  const good = C.check(f, { name: 'Ada\nL', mail: ' ada@example.com ', day: 'Tue', food: ['Soup', 'Pizza', 'Soup'], n: '3', when: '2026-10-29', stars: 4, extra: 'ignored' });
  assert.ok(good.ok, JSON.stringify(good.errors));
  assert.deepStrictEqual(good.values, { name: 'Ada L', mail: 'ada@example.com', day: 'Tue', food: ['Pizza', 'Soup'], n: 3, when: '2026-10-29', stars: 4 });
  const bad = C.check(f, { name: '  ', mail: 'nope', day: 'Sun', food: ['Cake'], n: 2.5, when: '2026-02-30', stars: 11 });
  assert.deepStrictEqual(Object.keys(bad.errors).sort(), ['day', 'food', 'mail', 'n', 'name', 'stars', 'when']);
  assert.ok(!C.check(f, { name: 'x', day: 'Mon', n: 9 }).ok, 'over the maximum');
  assert.ok(!C.check(f, { name: { toString: () => 'x' }, day: 'Mon' }).ok, 'objects are not text');
  assert.ok(C.check(f, { name: 'x', day: 'Mon' }).ok, 'optional questions may be left empty');
});

test('answers become cells that can never be formulas', () => {
  const cells = C.toCells({ a: '=HYPERLINK("http://x","y")', b: '+1 555', c: '@SUM(A1)', d: 'plain', e: 42, f: ['x', 'y'], g: '12', h: 'TRUE', i: '#N/A' }, ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'missing']);
  assert.deepStrictEqual(cells, ["'=HYPERLINK(\"http://x\",\"y\")", "'+1 555", "'@SUM(A1)", 'plain', 42, 'x, y', "'12", "'TRUE", "'#N/A", '']);
});

async function api(port, method, p, body, key) {
  const res = await fetch(`http://127.0.0.1:${port}${p}`, { method, headers: { ...(body ? { 'Content-Type': 'application/json' } : {}), ...(key ? { Authorization: 'Bearer ' + key } : {}) }, body: body ? JSON.stringify(body) : undefined, redirect: 'manual' });
  const text = await res.text();
  let data = null; try { data = JSON.parse(text); } catch { /* not JSON */ }
  return { status: res.status, data, headers: res.headers };
}

test('the web server publishes forms and collects answers', async () => {
  const srv = createWebServer({ live: false, apiKey: '', formsLimits: { perMinute: 5 } });
  const port = await listen(srv);
  try {
    const made = await api(port, 'POST', '/api/forms', { form: FORM });
    assert.strictEqual(made.status, 201);
    const { id, key } = made.data;
    assert.match(id, /^[A-Za-z0-9_-]{22}$/);
    assert.ok(key.length >= 40);
    // The public view has the form, not the key.
    const pub = await api(port, 'GET', '/api/forms/' + id);
    assert.strictEqual(pub.data.form.title, 'Team lunch');
    assert.ok(!JSON.stringify(pub.data).includes(key) && !('keyHash' in pub.data));
    // /f/<id> leads to the answer page.
    const link = await api(port, 'GET', '/f/' + id);
    assert.strictEqual(link.status, 302);
    assert.strictEqual(link.headers.get('location'), '/src/renderer/form.html?id=' + id);
    assert.strictEqual((await api(port, 'GET', '/f/../../etc/passwd')).status, 404);
    // Answers: checked, numbered.
    assert.strictEqual((await api(port, 'POST', `/api/forms/${id}/responses`, { answers: { name: 'Ada', day: 'Mon' } })).status, 201);
    const wrong = await api(port, 'POST', `/api/forms/${id}/responses`, { answers: { day: 'Sun' } });
    assert.strictEqual(wrong.status, 422);
    assert.deepStrictEqual(Object.keys(wrong.data.errors).sort(), ['day', 'name']);
    assert.strictEqual((await api(port, 'POST', `/api/forms/${id}/responses`, { answers: { name: 'Bo', day: 'Tue', n: 2 } })).status, 201);
    // Only the key reads them.
    assert.strictEqual((await api(port, 'GET', `/api/forms/${id}/responses`)).status, 403);
    assert.strictEqual((await api(port, 'GET', `/api/forms/${id}/responses`, null, 'x'.repeat(43))).status, 403);
    const got = await api(port, 'GET', `/api/forms/${id}/responses?after=1`, null, key);
    assert.strictEqual(got.data.count, 2);
    assert.deepStrictEqual(got.data.responses.map(r => [r.n, r.values.name, r.values.n]), [[2, 'Bo', 2]]);
    assert.strictEqual(got.headers.get('access-control-allow-origin'), '*');
    // Too many answers from one address in a minute.
    const codes = [];
    for (let i = 0; i < 4; i++) codes.push((await api(port, 'POST', `/api/forms/${id}/responses`, { answers: { name: 'x', day: 'Mon' } })).status);
    assert.deepStrictEqual(codes, [201, 201, 429, 429]);
    // Closing and changing the form.
    assert.strictEqual((await api(port, 'PUT', '/api/forms/' + id, { open: false }, key)).data.open, false);
    assert.strictEqual((await api(port, 'GET', '/api/forms/' + id)).data.open, false);
    const changed = await api(port, 'PUT', '/api/forms/' + id, { form: { ...FORM, title: 'Team dinner' } }, key);
    assert.strictEqual(changed.status, 200);
    assert.strictEqual((await api(port, 'PUT', '/api/forms/' + id, { form: { title: 'empty', questions: [] } }, key)).status, 400);
    assert.strictEqual((await api(port, 'PUT', '/api/forms/' + id, { open: true }, 'bad')).status, 403);
    assert.strictEqual((await api(port, 'DELETE', '/api/forms/' + id, null, key)).status, 200);
    assert.strictEqual((await api(port, 'GET', '/api/forms/' + id)).status, 404);
    // Junk is refused, not crashed on.
    assert.strictEqual((await api(port, 'POST', '/api/forms', { form: { questions: [] } })).status, 400);
    const junk = await fetch(`http://127.0.0.1:${port}/api/forms`, { method: 'POST', body: '{"form":' });
    assert.strictEqual(junk.status, 400);
    const big = await fetch(`http://127.0.0.1:${port}/api/forms`, { method: 'POST', body: JSON.stringify({ form: { ...FORM, description: 'x'.repeat(600e3) } }) }).catch(() => ({ status: 413 }));
    assert.strictEqual(big.status, 413);
  } finally { srv.close(); }
});

test('forms kept in a folder survive a restart; ids can’t reach other files', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'forms-'));
  try {
    const a = createFormsStore({ dir });
    const { id, key } = a.create(FORM);
    a.respond(id, { name: 'Ada', day: 'Mon', food: ['Salad'] });
    a.close();
    const b = createFormsStore({ dir });
    const got = b.responses(id, key);
    assert.strictEqual(got.count, 1);
    assert.deepStrictEqual(got.responses[0].values.food, ['Salad']);
    assert.strictEqual(b.respond(id, { name: 'Bo', day: 'Tue' }).n, 2);
    const meta = fs.readFileSync(path.join(dir, id + '.json'), 'utf8');
    assert.ok(!meta.includes(key), 'only the key’s hash is kept');
    assert.throws(() => b.publicForm('../' + id), /doesn’t exist/);
    assert.throws(() => b.responses(id, key.slice(1)), /key/);
    b.close();
  } finally { fs.rmSync(dir, { recursive: true, force: true }); }
});

test('the local-network server serves only the answer page and the API', async () => {
  const srv = createFormsServer();
  const port = await listen(srv);
  try {
    const { id } = srv.store.create(FORM);
    const page = await fetch(`http://127.0.0.1:${port}/src/renderer/form.html`);
    assert.strictEqual(page.status, 200);
    assert.match(page.headers.get('content-security-policy') || (await page.text()), /default-src 'none'/);
    for (const p of ['/src/renderer/sheets.html', '/package.json', '/src/main/main.js', '/src/renderer/js/../../main/main.js']) assert.strictEqual((await fetch(`http://127.0.0.1:${port}${p}`)).status, 404, p);
    assert.strictEqual((await api(port, 'POST', `/api/forms/${id}/responses`, { answers: { name: 'Ada', day: 'Mon' } })).status, 201);
  } finally { srv.close(); }
});
