'use strict';
// Password-protected Office documents (src/renderer/js/office-crypto.js): ECMA-376 Agile encryption.
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const C = require('../src/renderer/js/office-crypto.js');
const FX = path.join(__dirname, 'fixtures');

test('SHA-512 (for key stretching) matches Node\'s', () => {
  for (const n of [0, 1, 68, 111, 112, 127, 128, 129, 300]) {
    const m = crypto.randomBytes(n);
    assert.strictEqual(Buffer.from(C.sha512(new Uint8Array(m))).toString('hex'), crypto.createHash('sha512').update(m).digest('hex'), String(n));
  }
});

test('a protected file opens with its password and only with it', async () => {
  const plain = new Uint8Array(crypto.randomBytes(10000));
  const enc = await C.encrypt(plain, 'Pässwörd 1!', { spinCount: 1000 });
  assert.ok(C.isEncrypted(enc));
  assert.ok(!C.isEncrypted(plain));
  assert.deepStrictEqual(await C.decrypt(enc, 'Pässwörd 1!'), plain);
  await assert.rejects(C.decrypt(enc, 'pässwörd 1!'), e => e.name === 'BadPassword');
  // Each save is different (fresh salts and key), even with the same password.
  assert.notDeepStrictEqual(await C.encrypt(plain, 'x', { spinCount: 10 }), await C.encrypt(plain, 'x', { spinCount: 10 }));
});

test('a changed file is noticed', async () => {
  const plain = new Uint8Array(crypto.randomBytes(5000));
  const enc = await C.encrypt(plain, 'pw', { spinCount: 10 });
  // Flip one byte in the encrypted package (found by its ciphertext, deep in the file).
  const CFB = require('cfb');
  const cfb = CFB.read(enc, { type: 'array' });
  const pkg = CFB.find(cfb, 'EncryptedPackage');
  pkg.content[100] ^= 1;
  const tampered = new Uint8Array(CFB.write(cfb, { type: 'array' }));
  await assert.rejects(C.decrypt(tampered, 'pw'), /changed or damaged/);
});

test('files protected by other software open (msoffcrypto-tool, as Office does it)', async () => {
  const theirs = new Uint8Array(fs.readFileSync(path.join(FX, 'protected-msoffcrypto.xlsx')));
  const expected = new Uint8Array(fs.readFileSync(path.join(FX, 'protected-msoffcrypto.plain.xlsx')));
  assert.ok(C.isEncrypted(theirs));
  assert.deepStrictEqual(await C.decrypt(theirs, 'Secret-42'), expected);
  await assert.rejects(C.decrypt(theirs, 'Secret-43'), e => e.name === 'BadPassword');
});
