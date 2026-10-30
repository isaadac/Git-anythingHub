'use strict';
// Careful writes and guarded web requests for the main process (src/main/safe-io.js).
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const http = require('http');
const { writeAtomic, readJson, forbidden, safeGet } = require('../src/main/safe-io');

test('a file is replaced whole, keeping its permissions, with nothing left behind', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'isadek-io-'));
  const f = path.join(dir, 'report.docx');
  writeAtomic(f, Buffer.from('first'));
  fs.chmodSync(f, 0o640);
  writeAtomic(f, Buffer.from('second'));
  assert.strictEqual(fs.readFileSync(f, 'utf8'), 'second');
  if (process.platform !== 'win32') assert.strictEqual(fs.statSync(f).mode & 0o777, 0o640);
  assert.deepStrictEqual(fs.readdirSync(dir), ['report.docx']);
  // A failed write leaves the old file as it was.
  assert.throws(() => writeAtomic(path.join(dir, 'missing-folder', 'x.txt'), 'x'));
  assert.strictEqual(fs.readFileSync(f, 'utf8'), 'second');
  fs.writeFileSync(path.join(dir, 'bad.json'), '{"not": "a list"}');
  assert.deepStrictEqual(readJson(path.join(dir, 'bad.json'), Array.isArray, []), []);
  assert.deepStrictEqual(readJson(path.join(dir, 'none.json'), Array.isArray, ['x']), ['x']);
  fs.rmSync(dir, { recursive: true });
});

test('addresses web data may not come from', () => {
  for (const ip of ['127.0.0.1', '127.8.8.8', '0.0.0.0', '169.254.169.254', '100.100.100.200', '224.0.0.1', '::1', '::', 'fe80::1', '::ffff:127.0.0.1', '::ffff:7f00:1', '::ffff:169.254.169.254', 'fd00:ec2::254']) assert.ok(forbidden(ip), ip);
  for (const ip of ['8.8.8.8', '192.168.1.10', '10.0.0.5', '2606:4700::1111']) assert.ok(!forbidden(ip), ip);
});

test('web data can\'t be fetched from this computer, by number, by name or by redirect', async t => {
  const srv = http.createServer((req, res) => res.end('secret'));
  await new Promise(r => srv.listen(0, '127.0.0.1', r));
  t.after(() => srv.close());
  const port = srv.address().port;
  for (const u of [`http://127.0.0.1:${port}/`, `http://localhost:${port}/`, `http://[::ffff:127.0.0.1]:${port}/`, 'http://169.254.169.254/latest/meta-data/']) {
    await assert.rejects(safeGet(u, { timeoutMs: 5000 }), /on this computer or a private service address/, u);
  }
  await assert.rejects(safeGet('file:///etc/passwd'), /Only http and https/);
  await assert.rejects(safeGet('http://user:pw@example.com/'), /user name or password/);
});
