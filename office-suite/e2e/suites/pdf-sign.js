// Isadek PDF: digital signatures. A Digital ID is created (or a certificate imported), a PDF is
// signed (an incremental update: a second signature keeps the first valid), the signatures are
// checked when the PDF opens, and changed or damaged documents are reported. Poppler's pdfsig
// (when installed) checks the signatures independently.
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const W = process.env.E2E_WORK;
const pdfsig = file => { try { return execFileSync('pdfsig', [file], { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] }); } catch (e) { return e.code === 'ENOENT' ? null : String(e.stdout || '') + String(e.stderr || ''); } };

(async () => {
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 1300, height: 900 } });
  const errors = [];
  p.on('pageerror', e => errors.push(e.message));
  await p.goto(process.env.E2E_BASE + '/src/renderer/pdf.html');
  await p.waitForFunction(() => window.__pdfReady && window.AtlasPdf && window.PdfSign);
  // Dialogs answer from lists; saving keeps the bytes.
  await p.evaluate(() => {
    window.__forms = []; window.__chooses = []; window.__saved = [];
    UI.form = async (t, fields) => { const want = window.__forms.shift(); if (want === null) return null; const out = {}; fields.forEach(f => { out[f.name] = f.value !== undefined ? f.value : (f.options ? (f.options[0].value ?? f.options[0]) : ''); }); return Object.assign(out, want || {}); };
    UI.choose = async () => window.__chooses.shift() ?? -1;
    UI.confirm = async () => true;
    Platform.save = async o => { window.__saved.push({ name: o.name, data: Array.from(o.data) }); return { name: o.name }; };
  });
  const lastSaved = async () => Buffer.from(await p.evaluate(() => window.__saved[window.__saved.length - 1].data));
  const open = (bytes, name) => p.evaluate(([d, name]) => AtlasPdf.loadFile({ data: new Uint8Array(d), name }), [Array.from(bytes), name]);

  // A two-page PDF, made with object streams (as pdf-lib and many apps write them).
  const base = Buffer.from(await p.evaluate(async () => {
    const d = await PDFLib.PDFDocument.create();
    for (const t of ['Agreement', 'Page two']) d.addPage([400, 300]).drawText(t, { x: 50, y: 200, size: 24 });
    return Array.from(await d.save());
  }));
  await open(base, 'agreement.pdf');
  check('an unsigned PDF shows no signature bar', await p.isHidden('#sigBar'));

  // ---- Digital ID and the first signature
  await p.evaluate(() => { window.__chooses = [0]; window.__forms = [{ name: 'Ada Lovelace', email: 'ada@example.com', org: 'Analytical Engines' }, { reason: 'I approve', location: 'London', look: 'last' }]; });
  await p.click('[data-tab="sign"]');
  await p.click('#signDigital');
  await p.waitForFunction(() => window.__saved.length === 1 && AtlasPdf.signatures.length === 1, null, { timeout: 30000 });
  const s1 = await lastSaved();
  fs.writeFileSync(path.join(W, 's1.pdf'), s1);
  check('signing saves a copy named -signed and opens it', (await p.evaluate(() => window.__saved[0].name)) === 'agreement-signed.pdf');
  check('…the original bytes are kept and a signature added after them', s1.subarray(0, base.length).equals(base) && s1.length > base.length);
  check('…the signature bar says who signed and that nothing changed', (await p.getAttribute('#sigBar', 'class')).includes('good') && /Signed by Ada Lovelace\. The document hasn’t changed/.test(await p.textContent('#sigText')), await p.textContent('#sigText'));
  check('…your own Digital ID is trusted', !/isn’t confirmed/.test(await p.textContent('#sigText')));
  const sig = await p.evaluate(() => { const s = AtlasPdf.signatures[0]; return { ok: s.ok, reason: s.reason, location: s.location, self: s.selfSigned, cn: s.signer.cn, email: s.signer.email, hash: s.hash }; });
  check('…its reason, place, signer and hash are read back', sig.ok && sig.reason === 'I approve' && sig.location === 'London' && sig.self && sig.cn === 'Ada Lovelace' && sig.email === 'ada@example.com' && sig.hash === 'SHA-256', JSON.stringify(sig));
  check('…the visible stamp is on the last page', /Digitally signed by Ada Lovelace/.test(s1.toString('latin1')) && (await p.evaluate(async () => { const pg = await AtlasPdf.sources[0].doc.getPage(2); return (await pg.getAnnotations()).some(a => a.fieldType === 'Sig'); })));
  const ps1 = pdfsig(path.join(W, 's1.pdf'));
  if (ps1 === null) console.log('PASS pdfsig not installed: independent check skipped');
  else check('Poppler’s pdfsig finds the signature valid and covering the whole document', /Signature Validation: Signature is Valid/.test(ps1) && /Total document signed/.test(ps1) && /Ada Lovelace/.test(ps1), ps1);

  // ---- A second signer (an imported certificate) keeps the first signature valid
  // Importing: a PKCS#8 key and its certificate (as exported from another tool).
  const imp = await p.evaluate(async () => {
    const pem = (label, der) => `-----BEGIN ${label}-----\n${btoa(String.fromCharCode(...der)).match(/.{1,64}/g).join('\n')}\n-----END ${label}-----\n`;
    const a = await PdfSign.createIdentity({ name: 'Imported Person' }, { extractable: true });
    const other = await PdfSign.createIdentity({ name: 'Someone Else' });
    const key = pem('PRIVATE KEY', new Uint8Array(await crypto.subtle.exportKey('pkcs8', a.key)));
    const ok = await PdfSign.importIdentity(key, PdfSign.certPem(a.certs[0])).then(id => PdfSign.readCert(id.certs[0]).subject.cn, e => 'ERR ' + e.message);
    const mismatch = await PdfSign.importIdentity(key, PdfSign.certPem(other.certs[0])).then(() => 'accepted', e => e.message);
    const enc = await PdfSign.importIdentity('-----BEGIN ENCRYPTED PRIVATE KEY-----\nAAAA\n-----END ENCRYPTED PRIVATE KEY-----', PdfSign.certPem(a.certs[0])).then(() => 'accepted', e => e.message);
    return { ok, mismatch, enc };
  });
  check('a PKCS#8 key and its certificate import as a Digital ID', imp.ok === 'Imported Person', imp.ok);
  check('…a key must match its certificate', /doesn’t belong/.test(imp.mismatch), imp.mismatch);
  check('…and an encrypted key is explained', /openssl pkcs8/.test(imp.enc), imp.enc);
  // A second Digital ID (made here, then stored as the device's ID) signs the signed copy.
  await p.evaluate(() => { window.__chooses = [1, 0]; window.__forms = [{ name: 'Zoë Bürger' }]; });
  await p.click('#digitalId');
  await p.waitForFunction(() => window.__chooses.length === 0);
  await p.waitForTimeout(500);
  await p.evaluate(() => { window.__forms = [{ reason: 'Gegengezeichnet', look: 'none' }]; });
  await p.click('#signDigital');
  await p.waitForFunction(() => window.__saved.length === 2 && AtlasPdf.signatures.length === 2, null, { timeout: 30000 });
  const s2 = await lastSaved();
  fs.writeFileSync(path.join(W, 's2.pdf'), s2);
  const two = await p.evaluate(() => AtlasPdf.signatures.map(s => ({ ok: s.ok, later: s.laterChanges, cn: s.signer.cn, trusted: s.trusted })));
  check('a second signature keeps the first one valid (it notes the later change)', two[0].ok && two[0].later && two[1].ok && !two[1].later && two[1].cn === 'Zoë Bürger', JSON.stringify(two));
  check('…the second signature’s name keeps its accents', two[1].cn === 'Zoë Bürger');
  check('…the replaced Digital ID’s certificate is trusted too', two.every(s => s.trusted));
  const ps2 = pdfsig(path.join(W, 's2.pdf'));
  if (ps2 !== null) check('pdfsig finds both valid, the first as not covering the whole file', (ps2.match(/Signature is Valid/g) || []).length === 2 && /Signature Field Name: Signature2/.test(ps2) && /Not total document signed/.test(ps2), ps2);

  // ---- Changes and damage
  const changed = Buffer.from(s2); changed[300] ^= 1;
  await open(changed, 'changed.pdf');
  check('a document changed after signing is reported as not valid', (await p.getAttribute('#sigBar', 'class')).includes('bad') && /not valid/.test(await p.textContent('#sigText')), await p.textContent('#sigText'));
  await p.click('#sigDetails');
  check('…the details say what is wrong', /contents don’t match what was signed/.test(await p.textContent('#sigPane .list')));
  await p.click('#sigPaneClose');
  // Text appended after the signature (an edit by another program).
  const appended = Buffer.concat([s1, Buffer.from('\n% edited later\n')]);
  await open(appended, 'appended.pdf');
  check('changes added after the last signature are noted', (await p.getAttribute('#sigBar', 'class')).includes('warn') && /changed after the last signature/.test(await p.textContent('#sigText')));
  // A signature whose bytes were damaged.
  const s1t = s1.toString('latin1'), at = s1t.indexOf('/Contents <') + 200;
  const damaged = Buffer.from(s1); damaged[at] = damaged[at] === 0x41 ? 0x42 : 0x41;
  await open(damaged, 'damaged.pdf');
  check('a damaged signature is not valid', (await p.getAttribute('#sigBar', 'class')).includes('bad'));
  // Someone else's self-issued signature: valid, identity not confirmed, until trusted.
  const other = Buffer.from(await p.evaluate(async b => { const id = await PdfSign.createIdentity({ name: 'Mallory' }); return Array.from(await PdfSign.sign(new Uint8Array(b), id, {})); }, Array.from(base)));
  await open(other, 'other.pdf');
  check('a stranger’s valid signature says their identity isn’t confirmed', /Signed by Mallory\..*isn’t confirmed/.test(await p.textContent('#sigText')));
  await p.click('#sigDetails');
  await p.click('#sigPane button:has-text("Trust This Signer")');
  await p.waitForFunction(() => /Stop Trusting/.test(document.querySelector('#sigPane .list').textContent));
  check('…Trust This Signer remembers them', !/isn’t confirmed/.test(await p.textContent('#sigText')));
  // A PDF that only looks signed: a byte range pointing at nothing.
  const fake = Buffer.concat([base, Buffer.from('\n1 0 obj << /Type /Sig /ByteRange [0 10 20 999999] /Contents <00> >> endobj\n')]);
  await open(fake, 'fake.pdf');
  check('a broken signature dictionary is reported, not trusted', (await p.getAttribute('#sigBar', 'class')).includes('bad'));

  check('no page errors', !errors.length, errors.join(' | '));
  await b.close();
})().catch(e => { console.error('CRASH', e.stack); process.exit(1); });
