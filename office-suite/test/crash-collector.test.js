'use strict';
// The crash-report collector: script errors as JSON, Crashpad minidumps (gzip multipart), limits.
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const zlib = require('zlib');
const { Readable } = require('stream');
const { createCrashCollector } = require('../src/server/crash-collector');

function call(c, { body, type, gzip, ip = '1.2.3.4', method = 'POST' }) {
  const data = gzip ? zlib.gzipSync(body) : Buffer.from(body);
  const req = Readable.from([data]);
  Object.assign(req, { method, headers: { 'content-type': type, ...(gzip ? { 'content-encoding': 'gzip' } : {}) }, socket: { remoteAddress: ip } });
  return new Promise(resolve => {
    const res = { writeHead(code) { this.code = code; }, end(text) { resolve({ code: this.code, text }); } };
    c.handle(req, res);
  });
}
const tmp = () => fs.mkdtempSync(path.join(os.tmpdir(), 'crash-'));

test('a script error is kept as a line of JSON, with only simple fields', async () => {
  const dir = tmp(), c = createCrashCollector({ dir });
  const r = await call(c, { type: 'application/json', body: JSON.stringify({ message: 'boom', version: '1.0.0', nested: { x: 1 }, 'bad key!': 1 }) });
  assert.strictEqual(r.code, 200);
  const file = fs.readdirSync(dir).find(f => f.endsWith('.jsonl'));
  const e = JSON.parse(fs.readFileSync(path.join(dir, file), 'utf8'));
  assert.strictEqual(e.message, 'boom');
  assert.strictEqual(e.nested, undefined);
  assert.strictEqual(e['bad key!'], undefined);
  c.close();
});

test('a Crashpad upload (gzip multipart) keeps the minidump and its fields', async () => {
  const dir = tmp(), c = createCrashCollector({ dir });
  const b = '---MultipartBoundary-abc---';
  const body = Buffer.concat([
    Buffer.from(`--${b}\r\nContent-Disposition: form-data; name="_version"\r\n\r\n1.2.0\r\n`),
    Buffer.from(`--${b}\r\nContent-Disposition: form-data; name="upload_file_minidump"; filename="x.dmp"\r\nContent-Type: application/octet-stream\r\n\r\n`),
    Buffer.from('MDMP\x00\x01binary'), Buffer.from(`\r\n--${b}--\r\n`)
  ]);
  const r = await call(c, { type: `multipart/form-data; boundary=${b}`, body, gzip: true });
  assert.strictEqual(r.code, 200);
  const dmp = fs.readdirSync(dir).find(f => f.endsWith('.dmp'));
  assert.strictEqual(fs.readFileSync(path.join(dir, dmp)).toString('latin1'), 'MDMP\x00\x01binary');
  assert.strictEqual(JSON.parse(fs.readFileSync(path.join(dir, dmp.replace('.dmp', '.json')), 'utf8'))._version, '1.2.0');
  c.close();
});

test('anything else is refused, and each address is limited', async () => {
  const dir = tmp(), c = createCrashCollector({ dir, perHour: 2 });
  assert.strictEqual((await call(c, { type: 'text/plain', body: 'hi' })).code, 415);
  assert.strictEqual((await call(c, { type: 'multipart/form-data; boundary=x', body: '--x\r\nContent-Disposition: form-data; name="f"; filename="a"\r\n\r\nnot a dump\r\n--x--' })).code, 400);
  assert.strictEqual((await call(c, { type: 'application/json', body: '{}' })).code, 429);
  assert.strictEqual((await call(c, { type: 'application/json', body: '{}', ip: '5.6.7.8' })).code, 200);
  assert.strictEqual((await call(c, { type: 'application/json', body: '[1]', ip: '5.6.7.8' })).code, 400);
  c.close();
});

test('reports older than 90 days are deleted', () => {
  const dir = tmp();
  const old = path.join(dir, 'old.dmp');
  fs.writeFileSync(old, 'x');
  const t = (Date.now() - 91 * 864e5) / 1000;
  fs.utimesSync(old, t, t);
  fs.writeFileSync(path.join(dir, 'new.dmp'), 'x');
  const c = createCrashCollector({ dir });
  assert.deepStrictEqual(fs.readdirSync(dir), ['new.dmp']);
  c.close();
});
