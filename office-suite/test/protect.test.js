'use strict';
const test = require('node:test');
const assert = require('node:assert');
const P = require('../src/renderer/js/sheets-protect.js');

test("Excel's legacy password hash", () => {
  assert.strictEqual(P.legacyHash('test'), 'CBEB');
  assert.strictEqual(P.legacyHash('password'), '83AF');
  assert.strictEqual(P.legacyHash(''), '');
});

test("Excel 2013+ salted SHA-512 hash", async () => {
  const p = { algo: { name: 'SHA-512', salt: '3c6dJaOD1NyRDCd4cX2O9g==', spin: 1000, hash: 'Cmv9ME7h1wdNs2y3a1GpRVMJWS1RtSS3maW2ycZKDQbh/bEOK9v+o0o2yGh2U4rRJR/4Y8nbJTXHvYekVam4ig==' } };
  assert.strictEqual(await P.check(p, 'secret'), true);
  assert.strictEqual(await P.check(p, 'Secret'), false);
});

test('sheetProtection XML both ways', () => {
  const p = P.clean({ hash: 'CBEB', allow: { formatCells: true, sort: true } });
  const x = P.toXml(p);
  assert.strictEqual(x, '<sheetProtection password="CBEB" sheet="1" objects="1" scenarios="1" formatCells="0" sort="0"/>');
  assert.deepStrictEqual(P.fromXml(x), p);
  assert.strictEqual(P.fromXml('<sheetProtection sheet="0"/>'), null);
  const m = P.fromXml('<sheetProtection algorithmName="SHA-512" hashValue="abc=" saltValue="def=" spinCount="100000" sheet="1" objects="1" scenarios="1" insertRows="0"/>');
  assert.deepStrictEqual(m, { hash: '', allow: { insertRows: true }, algo: { name: 'SHA-512', hash: 'abc=', salt: 'def=', spin: 100000 } });
});
