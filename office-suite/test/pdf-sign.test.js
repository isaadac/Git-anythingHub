'use strict';
const test = require('node:test');
const assert = require('node:assert');
const PDFLib = require('pdf-lib');
const S = require('../src/renderer/js/pdf-sign.js');

async function basePdf(objectStreams) {
  const d = await PDFLib.PDFDocument.create();
  d.addPage([400, 300]).drawText('Contract', { x: 50, y: 200, size: 24 });
  return d.save({ useObjectStreams: objectStreams });
}

test('a Digital ID: a self-issued certificate for signing', async () => {
  const id = await S.createIdentity({ name: 'Ada Lovelace', email: 'ada@example.com', org: 'Engines' });
  const c = S.readCert(id.certs[0]);
  assert.deepStrictEqual(c.subject, { o: 'Engines', cn: 'Ada Lovelace', email: 'ada@example.com' });
  assert.deepStrictEqual(c.issuer, c.subject);
  assert.ok(c.notAfter - c.notBefore > 4 * 365 * 864e5);
  assert.strictEqual(id.key.extractable, false);
  await assert.rejects(S.createIdentity({ name: '  ' }), /Enter your name/);
});

for (const os of [false, true]) {
  test(`signing and checking (${os ? 'object streams' : 'plain cross-references'})`, async () => {
    const base = await basePdf(os);
    const a = await S.createIdentity({ name: 'Ada' }), z = await S.createIdentity({ name: 'Zoë' });
    const s1 = await S.sign(base, a, { reason: 'Approved', location: 'Köln', visible: { page: 'last' } }, PDFLib);
    assert.ok(Buffer.from(s1.subarray(0, base.length)).equals(Buffer.from(base)), 'the original bytes are kept');
    let v = await S.verify(s1);
    assert.strictEqual(v.length, 1);
    assert.ok(v[0].ok && v[0].intact && v[0].selfSigned && v[0].issuerVerified, JSON.stringify(v[0].problems));
    assert.strictEqual(v[0].reason, 'Approved');
    assert.strictEqual(v[0].location, 'Köln');
    // pdf-lib still opens it and finds the field.
    const d = await PDFLib.PDFDocument.load(s1);
    assert.deepStrictEqual(d.getForm().getFields().map(f => f.getName()), ['Signature1']);
    const s2 = await S.sign(s1, z, {}, PDFLib);
    v = await S.verify(s2);
    assert.deepStrictEqual(v.map(x => [x.ok, x.intact, x.signer.cn]), [[true, false, 'Ada'], [true, true, 'Zoë']]);
    assert.deepStrictEqual((await PDFLib.PDFDocument.load(s2)).getForm().getFields().map(f => f.getName()), ['Signature1', 'Signature2']);
    // Any changed byte outside the signature breaks it.
    for (const at of [10, base.length - 5, s2.length - 3]) {
      const bad = Buffer.from(s2); bad[at] ^= 0x20;
      const r = await S.verify(bad);
      assert.ok(r.some(x => !x.ok), 'byte ' + at);
    }
  });
}

test('hostile signature data is reported, never crashes', async () => {
  const base = Buffer.from(await basePdf(false));
  let seed = 1;
  const rnd = () => (seed = (seed * 16807) % 2147483647) / 2147483647;
  for (let i = 0; i < 200; i++) {
    const bytes = Buffer.from(Array.from({ length: 1 + Math.floor(rnd() * 300) }, () => Math.floor(rnd() * 256)));
    if (i % 3 === 0) bytes[0] = 0x30;
    const hex = bytes.toString('hex');
    const pre = Buffer.concat([base, Buffer.from('\n9 0 obj << /Type /Sig /ByteRange [0 ')]);
    const gapStart = pre.length + 30;
    const tail = `${String(gapStart).padStart(10, ' ')} ${String(gapStart + hex.length + 2).padStart(10, ' ')} 0000000010] /Contents <${hex}> >> endobj\n`;
    const pdf = Buffer.concat([pre, Buffer.from(tail)]);
    const r = await S.verify(pdf);
    assert.ok(r.every(x => !x.ok));
  }
  // Deep nesting and huge lengths.
  assert.throws(() => S.parse(new Uint8Array(new Array(200).fill(0x30).concat([0]))));
  assert.throws(() => S.parse(new Uint8Array([0x30, 0x84, 0xff, 0xff, 0xff, 0xff])));
});

test('importing a key with its certificate', async () => {
  const a = await S.createIdentity({ name: 'Imp' }, { extractable: true });
  const pk = Buffer.from(await crypto.subtle.exportKey('pkcs8', a.key)).toString('base64');
  const key = `-----BEGIN PRIVATE KEY-----\n${pk}\n-----END PRIVATE KEY-----`;
  const id = await S.importIdentity(key, S.certPem(a.certs[0]));
  assert.strictEqual(S.readCert(id.certs[0]).subject.cn, 'Imp');
  const other = await S.createIdentity({ name: 'Other' });
  await assert.rejects(S.importIdentity(key, S.certPem(other.certs[0])), /doesn’t belong/);
  await assert.rejects(S.importIdentity('nothing', S.certPem(other.certs[0])), /No private key/);
  await assert.rejects(S.importIdentity(key, 'nothing'), /No certificate/);
});
