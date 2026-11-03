'use strict';
// Isadek PDF — digital signatures (PAdES-style: a CMS signature inside the PDF, adbe.pkcs7.detached).
//
//   PdfSign.createIdentity({ name, email, org })  → { key, certs: [DER] }   a self-issued Digital ID
//   PdfSign.importIdentity(pemKey, pemCerts)       → { key, certs }          a key and its certificates
//   PdfSign.sign(pdfBytes, identity, { reason, location, contact, visible }) → signed bytes
//   PdfSign.verify(pdfBytes) → [{ ok, intact, covers, signer, issuer, selfSigned, time, reason, … }]
//
// Signing appends an incremental update (the original bytes stay as they are, so earlier
// signatures stay valid): a signature field and its value, with a gap for the signature that is
// filled in after the rest of the file is hashed. The cryptography is WebCrypto's (RSA PKCS#1 v1.5
// or ECDSA, SHA-256); this file only reads and writes the DER structures around it.
(function (root) {
  const subtle = () => (root.crypto || globalThis.crypto).subtle;
  const te = new TextEncoder();
  const latin1 = b => { let s = ''; for (let i = 0; i < b.length; i += 8192) s += String.fromCharCode.apply(null, b.subarray(i, i + 8192)); return s; };
  const concat = parts => { const n = parts.reduce((t, p) => t + p.length, 0), out = new Uint8Array(n); let o = 0; for (const p of parts) { out.set(p, o); o += p.length; } return out; };
  const hex = b => [...b].map(x => x.toString(16).padStart(2, '0')).join('');
  const unhex = s => { const out = new Uint8Array(s.length >> 1); for (let i = 0; i < out.length; i++) out[i] = parseInt(s.substr(i * 2, 2), 16); return out; };
  const eq = (a, b) => a.length === b.length && a.every((x, i) => x === b[i]);

  // ---------------- DER ----------------
  function len(n) {
    if (n < 128) return [n];
    const b = [];
    while (n) { b.unshift(n & 255); n = Math.floor(n / 256); }
    return [0x80 | b.length, ...b];
  }
  const tlv = (tag, ...content) => { const c = concat(content.map(x => (x instanceof Uint8Array ? x : new Uint8Array(x)))); return concat([new Uint8Array([tag, ...len(c.length)]), c]); };
  const SEQ = (...c) => tlv(0x30, ...c), SET = (...c) => tlv(0x31, ...c);
  /** A SET OF, its elements in DER order (sorted by their encodings). */
  const SETOF = (...c) => SET(...c.slice().sort((a, b) => { for (let i = 0; i < Math.min(a.length, b.length); i++) if (a[i] !== b[i]) return a[i] - b[i]; return a.length - b.length; }));
  const CTX = (n, ...c) => tlv(0xa0 | n, ...c);
  const NULL = new Uint8Array([5, 0]);
  const OCTETS = b => tlv(4, b);
  const UTF8 = s => tlv(0x0c, te.encode(s));
  const PRINTABLE = s => tlv(0x13, te.encode(s));
  const IA5 = s => tlv(0x16, te.encode(s));
  const BOOL = v => tlv(1, [v ? 0xff : 0]);
  function INT(bytes) {
    let b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    let i = 0;
    while (i < b.length - 1 && b[i] === 0 && !(b[i + 1] & 0x80)) i++;
    b = b.subarray(i);
    return tlv(2, b[0] & 0x80 ? concat([new Uint8Array([0]), b]) : b);
  }
  const BITS = b => tlv(3, concat([new Uint8Array([0]), b]));
  function OID(s) {
    const p = s.split('.').map(Number), out = [40 * p[0] + p[1]];
    for (const n of p.slice(2)) { const b = [n & 127]; let v = Math.floor(n / 128); while (v) { b.unshift(0x80 | (v & 127)); v = Math.floor(v / 128); } out.push(...b); }
    return tlv(6, out);
  }
  function TIME(d) {
    const z = n => String(n).padStart(2, '0');
    const y = d.getUTCFullYear(), s = `${z(d.getUTCMonth() + 1)}${z(d.getUTCDate())}${z(d.getUTCHours())}${z(d.getUTCMinutes())}${z(d.getUTCSeconds())}Z`;
    return y >= 1950 && y < 2050 ? tlv(0x17, te.encode(z(y % 100) + s)) : tlv(0x18, te.encode(y + s));
  }

  /** A DER element: { tag, start, hs (header size), end, kids? }; content is bytes[start + hs, end). */
  function parse(b, at = 0, end = b.length, depth = 0) {
    if (depth > 40 || at + 2 > end) throw new Error('Bad signature data');
    const tag = b[at];
    let n = b[at + 1], hs = 2;
    if (n & 0x80) {
      const k = n & 127;
      if (!k || k > 4 || at + 2 + k > end) throw new Error('Bad signature data');
      n = 0;
      for (let i = 0; i < k; i++) n = n * 256 + b[at + 2 + i];
      hs += k;
    }
    const e = { tag, start: at, hs, end: at + hs + n };
    if (e.end > end) throw new Error('Bad signature data');
    if (tag & 0x20) {
      e.kids = [];
      for (let p = at + hs; p < e.end;) { const k = parse(b, p, e.end, depth + 1); e.kids.push(k); p = k.end; }
    }
    return e;
  }
  const body = (b, e) => b.subarray(e.start + e.hs, e.end);
  const whole = (b, e) => b.subarray(e.start, e.end);
  function oidOf(b, e) {
    const c = body(b, e), out = [Math.floor(c[0] / 40), c[0] % 40];
    let v = 0;
    for (let i = 1; i < c.length; i++) { v = v * 128 + (c[i] & 127); if (!(c[i] & 0x80)) { out.push(v); v = 0; } }
    return out.join('.');
  }
  function timeOf(b, e) {
    const s = new TextDecoder().decode(body(b, e));
    const m = /^(\d{2}|\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})?/.exec(s);
    if (!m) return null;
    let y = +m[1];
    if (m[1].length === 2) y += y < 50 ? 2000 : 1900;
    return new Date(Date.UTC(y, +m[2] - 1, +m[3], +m[4], +m[5], +(m[6] || 0)));
  }

  const OIDS = {
    rsa: '1.2.840.113549.1.1.1', sha256rsa: '1.2.840.113549.1.1.11', sha384rsa: '1.2.840.113549.1.1.12', sha512rsa: '1.2.840.113549.1.1.13', sha1rsa: '1.2.840.113549.1.1.5', pss: '1.2.840.113549.1.1.10',
    ec: '1.2.840.10045.2.1', p256: '1.2.840.10045.3.1.7', p384: '1.3.132.0.34', ecdsa256: '1.2.840.10045.4.3.2', ecdsa384: '1.2.840.10045.4.3.3', ecdsa512: '1.2.840.10045.4.3.4',
    sha256: '2.16.840.1.101.3.4.2.1', sha384: '2.16.840.1.101.3.4.2.2', sha512: '2.16.840.1.101.3.4.2.3', sha1: '1.3.14.3.2.26',
    data: '1.2.840.113549.1.7.1', signedData: '1.2.840.113549.1.7.2', contentType: '1.2.840.113549.1.9.3', messageDigest: '1.2.840.113549.1.9.4', signingTime: '1.2.840.113549.1.9.5',
    signingCertV2: '1.2.840.113549.1.9.16.2.47',
    cn: '2.5.4.3', o: '2.5.4.10', ou: '2.5.4.11', c: '2.5.4.6', email: '1.2.840.113549.1.9.1',
    keyUsage: '2.5.29.15', basic: '2.5.29.19', ski: '2.5.29.14', eku: '2.5.29.37', emailProtection: '1.3.6.1.5.5.7.3.4', docSigning: '1.3.6.1.4.1.311.10.3.12'
  };
  const HASH = { [OIDS.sha256]: 'SHA-256', [OIDS.sha384]: 'SHA-384', [OIDS.sha512]: 'SHA-512', [OIDS.sha1]: 'SHA-1' };
  const digest = async (alg, data) => new Uint8Array(await subtle().digest(alg, data));

  // ---------------- Certificates ----------------
  function name({ name: cn, email, org }) {
    const rdn = (oid, v) => SET(SEQ(OID(oid), v));
    return SEQ(...[org && rdn(OIDS.o, UTF8(org)), rdn(OIDS.cn, UTF8(cn || 'Isadek user')), email && rdn(OIDS.email, IA5(email))].filter(Boolean));
  }
  function readName(b, e) {
    const out = {};
    for (const set of e.kids || []) for (const atv of set.kids || []) {
      const oid = oidOf(b, atv.kids[0]), v = new TextDecoder().decode(body(b, atv.kids[1]));
      const k = { [OIDS.cn]: 'cn', [OIDS.o]: 'o', [OIDS.ou]: 'ou', [OIDS.c]: 'c', [OIDS.email]: 'email' }[oid];
      if (k && !out[k]) out[k] = v;
    }
    return out;
  }
  /** What a certificate says: subject, issuer, validity, its public key and the bytes it signs. */
  function readCert(der) {
    const b = der instanceof Uint8Array ? der : new Uint8Array(der);
    const c = parse(b), tbs = c.kids[0], f = tbs.kids;
    const o = f[0].tag === 0xa0 ? 1 : 0;
    const serial = body(b, f[o]), issuer = f[o + 2], validity = f[o + 3], subject = f[o + 4], spki = f[o + 5];
    const keyAlg = spki.kids[0].kids;
    return {
      der: b, serial, tbs: whole(b, tbs), sigAlg: oidOf(b, c.kids[1].kids[0]), signature: body(b, c.kids[2]).subarray(1),
      issuerDer: whole(b, issuer), subjectDer: whole(b, subject), issuer: readName(b, issuer), subject: readName(b, subject),
      notBefore: timeOf(b, validity.kids[0]), notAfter: timeOf(b, validity.kids[1]),
      spki: whole(b, spki), keyType: oidOf(b, keyAlg[0]), curve: keyAlg[1] && keyAlg[1].tag === 6 ? oidOf(b, keyAlg[1]) : null
    };
  }
  const fingerprint = async der => hex(await digest('SHA-256', der)).toUpperCase().match(/../g).join(':');

  /** WebCrypto parameters for a key from a certificate (or a key type) and a hash. */
  function algOf(keyType, curve, hash = 'SHA-256') {
    if (keyType === OIDS.rsa) return { name: 'RSASSA-PKCS1-v1_5', hash };
    if (keyType === OIDS.ec) return { name: 'ECDSA', namedCurve: curve === OIDS.p384 ? 'P-384' : 'P-256', hash };
    throw new Error('This kind of key isn’t supported');
  }
  const sigAlgOid = (keyType, hash = 'SHA-256') => (keyType === OIDS.ec ? { 'SHA-256': OIDS.ecdsa256, 'SHA-384': OIDS.ecdsa384, 'SHA-512': OIDS.ecdsa512 }[hash] : { 'SHA-256': OIDS.sha256rsa, 'SHA-384': OIDS.sha384rsa, 'SHA-512': OIDS.sha512rsa }[hash]);
  const algId = (oid, keyType) => (keyType === OIDS.ec ? SEQ(OID(oid)) : SEQ(OID(oid), NULL));

  /** ECDSA signatures: WebCrypto's r‖s and DER's SEQUENCE { r, s }. */
  function ecToDer(raw) { const h = raw.length / 2; return SEQ(INT(raw.subarray(0, h)), INT(raw.subarray(h))); }
  function ecFromDer(der, size) {
    const e = parse(der), out = new Uint8Array(size * 2);
    [0, 1].forEach(i => { let v = body(der, e.kids[i]); while (v.length > size && v[0] === 0) v = v.subarray(1); out.set(v, size * (i + 1) - v.length); });
    return out;
  }
  async function signWith(key, keyType, data, hash = 'SHA-256') {
    const alg = key.algorithm.name === 'ECDSA' ? { name: 'ECDSA', hash } : { name: 'RSASSA-PKCS1-v1_5' };
    const s = new Uint8Array(await subtle().sign(alg, key, data));
    return key.algorithm.name === 'ECDSA' ? ecToDer(s) : s;
  }
  async function verifyWith(cert, sig, data, hash) {
    const alg = algOf(cert.keyType, cert.curve, hash);
    const key = await subtle().importKey('spki', cert.spki, alg, false, ['verify']);
    const s = alg.name === 'ECDSA' ? ecFromDer(sig, alg.namedCurve === 'P-384' ? 48 : 32) : sig;
    return subtle().verify(alg.name === 'ECDSA' ? { name: 'ECDSA', hash } : alg, key, s, data);
  }
  const hashOfSigAlg = oid => ({ [OIDS.sha256rsa]: 'SHA-256', [OIDS.sha384rsa]: 'SHA-384', [OIDS.sha512rsa]: 'SHA-512', [OIDS.sha1rsa]: 'SHA-1', [OIDS.ecdsa256]: 'SHA-256', [OIDS.ecdsa384]: 'SHA-384', [OIDS.ecdsa512]: 'SHA-512' }[oid]);

  /** A new Digital ID: an RSA key that can't be copied out, and a certificate the signer issues themselves. */
  async function createIdentity(who, opts = {}) {
    if (!who || !String(who.name || '').trim()) throw new Error('Enter your name');
    const pair = await subtle().generateKey({ name: 'RSASSA-PKCS1-v1_5', modulusLength: 2048, publicExponent: new Uint8Array([1, 0, 1]), hash: 'SHA-256' }, !!opts.extractable, ['sign', 'verify']);
    const spki = new Uint8Array(await subtle().exportKey('spki', pair.publicKey));
    const serial = (root.crypto || globalThis.crypto).getRandomValues(new Uint8Array(16)); serial[0] &= 0x7f;
    const now = new Date(Date.now() - 60000), until = new Date(now.getTime() + (opts.years || 5) * 365.25 * 864e5);
    const subj = name({ name: String(who.name).trim().slice(0, 64), email: who.email && String(who.email).trim().slice(0, 128), org: who.org && String(who.org).trim().slice(0, 64) });
    const ski = (await digest('SHA-1', spki)).subarray(0, 20);
    const ext = (oid, critical, value) => SEQ(OID(oid), ...(critical ? [BOOL(true)] : []), OCTETS(value));
    const tbs = SEQ(
      CTX(0, INT([2])), INT(serial), algId(OIDS.sha256rsa), subj, SEQ(TIME(now), TIME(until)), subj, spki,
      CTX(3, SEQ(
        ext(OIDS.basic, true, SEQ()),
        ext(OIDS.keyUsage, true, tlv(3, [6, 0xc0])), // digitalSignature, nonRepudiation
        ext(OIDS.eku, false, SEQ(OID(OIDS.emailProtection), OID(OIDS.docSigning))),
        ext(OIDS.ski, false, OCTETS(ski))
      ))
    );
    const sig = await signWith(pair.privateKey, OIDS.rsa, tbs);
    return { key: pair.privateKey, publicKey: pair.publicKey, certs: [SEQ(tbs, algId(OIDS.sha256rsa), BITS(sig))] };
  }

  const pemBlocks = (text, label) => [...String(text).matchAll(new RegExp(`-----BEGIN ${label}-----([\\s\\S]*?)-----END ${label}-----`, 'g'))].map(m => { const s = atob(m[1].replace(/\s+/g, '')); const b = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) b[i] = s.charCodeAt(i); return b; });
  /** A Digital ID from a PKCS#8 private key and its certificate (then any issuers), both PEM. */
  async function importIdentity(pemKey, pemCerts) {
    const keyDer = pemBlocks(pemKey, 'PRIVATE KEY')[0];
    if (!keyDer) throw new Error(/ENCRYPTED PRIVATE KEY|RSA PRIVATE KEY|EC PRIVATE KEY/.test(pemKey) ? 'The key must be an unencrypted PKCS#8 key (“BEGIN PRIVATE KEY”). Convert it with: openssl pkcs8 -topk8 -nocrypt' : 'No private key was found (“BEGIN PRIVATE KEY”)');
    const certs = pemBlocks(pemCerts, 'CERTIFICATE');
    if (!certs.length) throw new Error('No certificate was found (“BEGIN CERTIFICATE”)');
    const c = readCert(certs[0]);
    const key = await subtle().importKey('pkcs8', keyDer, algOf(c.keyType, c.curve), false, ['sign']);
    // The key must belong to the certificate.
    const probe = te.encode('isadek'), sig = await signWith(key, c.keyType, probe);
    if (!(await verifyWith(c, sig, probe, 'SHA-256'))) throw new Error('That key doesn’t belong to that certificate');
    return { key, certs };
  }
  const certPem = der => `-----BEGIN CERTIFICATE-----\n${btoa(latin1(der)).match(/.{1,64}/g).join('\n')}\n-----END CERTIFICATE-----\n`;

  // ---------------- CMS ----------------
  async function cms(identity, data, when) {
    const cert = readCert(identity.certs[0]);
    const md = await digest('SHA-256', data);
    const certHash = await digest('SHA-256', cert.der);
    const attrs = [
      SEQ(OID(OIDS.contentType), SET(OID(OIDS.data))),
      SEQ(OID(OIDS.signingTime), SET(TIME(when))),
      SEQ(OID(OIDS.messageDigest), SET(OCTETS(md))),
      // ESS signing-certificate-v2: which certificate signed (SHA-256 is its default hash).
      SEQ(OID(OIDS.signingCertV2), SET(SEQ(SEQ(SEQ(OCTETS(certHash))))))
    ];
    const signedAttrs = SETOF(...attrs);
    const signature = await signWith(identity.key, cert.keyType, signedAttrs);
    const signedAttrsCtx = concat([new Uint8Array([0xa0]), signedAttrs.subarray(1)]);
    const signer = SEQ(INT([1]), SEQ(cert.issuerDer, INT(cert.serial)), SEQ(OID(OIDS.sha256), NULL), signedAttrsCtx,
      cert.keyType === OIDS.ec ? SEQ(OID(OIDS.ecdsa256)) : SEQ(OID(OIDS.rsa), NULL), OCTETS(signature));
    const sd = SEQ(INT([1]), SET(SEQ(OID(OIDS.sha256), NULL)), SEQ(OID(OIDS.data)), CTX(0, ...identity.certs), SET(signer));
    return SEQ(OID(OIDS.signedData), CTX(0, sd));
  }

  /** Reads and checks a CMS signature over data. */
  async function checkCms(der, data) {
    const out = { ok: false, problems: [] };
    const ci = parse(der);
    if (oidOf(der, ci.kids[0]) !== OIDS.signedData) throw new Error('Not a CMS signature');
    const sd = ci.kids[1].kids[0].kids;
    let i = 3;
    const certs = [];
    if (sd[i] && sd[i].tag === 0xa0) { for (const c of sd[i].kids) if (c.tag === 0x30) { try { certs.push(readCert(whole(der, c).slice())); } catch { /* skip */ } } i++; }
    if (sd[i] && sd[i].tag === 0xa1) i++; // CRLs
    const si = sd[i].kids[0].kids;
    const sid = si[1];
    let cert = null;
    if (sid.tag === 0x30) { const iss = whole(der, sid.kids[0]), ser = body(der, sid.kids[1]); cert = certs.find(c => eq(c.issuerDer, iss) && eq(INT(c.serial).subarray(2), INT(ser).subarray(2))); }
    else cert = certs[0];
    if (!cert) throw new Error('The signer’s certificate isn’t in the signature');
    const hash = HASH[oidOf(der, si[2].kids[0])];
    if (!hash) throw new Error('Unsupported digest');
    let k = 3, attrs = null;
    if (si[k].tag === 0xa0) { attrs = si[k]; k++; }
    const sigAlg = oidOf(der, si[k].kids[0]);
    if (sigAlg === OIDS.pss) throw new Error('RSA-PSS signatures aren’t supported yet');
    const signature = body(der, si[k + 1]);
    const md = await digest(hash, data);
    let signedBytes = data;
    if (attrs) {
      let mdAttr = null;
      for (const a of attrs.kids) {
        const t = oidOf(der, a.kids[0]), v = a.kids[1].kids[0];
        if (t === OIDS.messageDigest) mdAttr = body(der, v);
        if (t === OIDS.signingTime) out.time = timeOf(der, v);
      }
      out.digestMatches = !!mdAttr && eq(mdAttr, md);
      if (!out.digestMatches) out.problems.push('The document’s contents don’t match what was signed');
      // The attributes are signed as a SET (DER tag 0x31), not as the [0] they are stored in.
      signedBytes = whole(der, attrs).slice(); signedBytes[0] = 0x31;
    } else out.digestMatches = true;
    out.signatureValid = await verifyWith(cert, signature, signedBytes, hash).catch(() => false);
    if (!out.signatureValid) out.problems.push('The signature doesn’t match the signer’s key');
    out.ok = out.digestMatches && out.signatureValid;
    out.hash = hash;
    out.cert = cert;
    out.signer = cert.subject; out.issuer = cert.issuer;
    out.selfSigned = eq(cert.issuerDer, cert.subjectDer);
    // Who vouches for the signer: an issuer in the signature that really signed the certificate.
    const issuerCert = certs.find(c => c !== cert && eq(c.subjectDer, cert.issuerDer));
    out.issuerVerified = issuerCert ? await verifyWith(issuerCert, cert.signature, cert.tbs, hashOfSigAlg(cert.sigAlg) || 'SHA-256').catch(() => false) : null;
    if (out.selfSigned) out.issuerVerified = await verifyWith(cert, cert.signature, cert.tbs, hashOfSigAlg(cert.sigAlg) || 'SHA-256').catch(() => false);
    const when = out.time || new Date();
    out.certValidAtSigning = (!cert.notBefore || when >= cert.notBefore) && (!cert.notAfter || when <= cert.notAfter);
    if (!out.certValidAtSigning) out.problems.push('The certificate wasn’t valid when it was used');
    out.fingerprint = await fingerprint(cert.der);
    return out;
  }

  // ---------------- PDF ----------------
  const pdfDate = d => { const z = n => String(n).padStart(2, '0'); return `D:${d.getUTCFullYear()}${z(d.getUTCMonth() + 1)}${z(d.getUTCDate())}${z(d.getUTCHours())}${z(d.getUTCMinutes())}${z(d.getUTCSeconds())}Z`; };
  /** A PDF text string: plain ASCII as a literal, else UTF-16 in hex. */
  function pdfText(s) {
    s = String(s || '');
    if (/^[\x20-\x7e]*$/.test(s)) return '(' + s.replace(/[\\()]/g, c => '\\' + c) + ')';
    let h = 'FEFF';
    for (const ch of s) { const c = ch.codePointAt(0); if (c > 0xffff) { const v = c - 0x10000; h += (0xd800 + (v >> 10)).toString(16).padStart(4, '0') + (0xdc00 + (v & 1023)).toString(16).padStart(4, '0'); } else h += c.toString(16).padStart(4, '0'); }
    return '<' + h.toUpperCase() + '>';
  }
  const winAnsi = s => String(s).replace(/[^\x20-\x7e]/g, '?').replace(/[\\()]/g, c => '\\' + c);
  const SIG_BYTES = 24576; // room for the signature (hex doubles it)

  /**
   * Signs a PDF: appends an incremental update with a signature field (invisible, or a stamp at
   * the bottom of a page) and the signature over the whole file. PDFLib: pdf-lib.
   */
  async function sign(bytes, identity, opts = {}, PDFLib = root.PDFLib) {
    const src = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    const { PDFDocument, PDFName, PDFArray, PDFDict, PDFRef } = PDFLib;
    const doc = await PDFDocument.load(src, { updateMetadata: false });
    const ctx = doc.context;
    if (ctx.trailerInfo.Encrypt) throw new Error('Password-protected PDFs can’t be signed here');
    const text = latin1(src.subarray(Math.max(0, src.length - 2048)));
    const sx = /startxref\s+(\d+)\s+%%EOF\s*$/.exec(text) || [...text.matchAll(/startxref\s+(\d+)/g)].pop();
    if (!sx) throw new Error('This PDF’s cross-reference table couldn’t be found');
    const prev = +sx[1];
    const cert = readCert(identity.certs[0]);
    const when = opts.date || new Date();
    const objs = new Map(); // number -> body text (written in the update)
    const pages = doc.getPages();
    const pi = opts.visible ? Math.max(0, Math.min(pages.length - 1, opts.visible.page === 'first' ? 0 : pages.length - 1)) : 0;
    const page = pages[pi];
    const pageRef = page.ref;

    // The field's name: Signature1, 2… not already used.
    // New objects get numbers above every object in the file, including the object streams and
    // cross-reference streams pdf-lib leaves out once it has read them.
    const raw = latin1(src);
    let top = ctx.largestObjectNumber;
    for (const m of raw.matchAll(/\/Size\s+(\d+)/g)) top = Math.max(top, +m[1] - 1);
    for (const m of raw.matchAll(/(?:^|[\r\n\s])(\d{1,9})\s+\d{1,5}\s+obj\b/g)) top = Math.max(top, +m[1]);
    ctx.largestObjectNumber = top;
    const hadForm = !!doc.catalog.get(PDFName.of('AcroForm'));
    const form = doc.getForm();
    const used = new Set(form.getFields().map(f => f.getName()));
    let n = 1; while (used.has('Signature' + n)) n++;
    const fieldName = 'Signature' + n;
    // (After getForm, which may add the form's own object.)
    let next = ctx.largestObjectNumber + 1;
    const sigNum = next++, widgetNum = next++, apNum = opts.visible ? next++ : 0;

    let rect = [0, 0, 0, 0];
    if (opts.visible) {
      const { width } = page.getSize();
      const w = 220, h = 54, x = Math.max(10, width - w - 36), y = 30;
      rect = [x, y, x + w, y + h];
      const who = cert.subject.cn || 'Signer';
      const lines = [`Digitally signed by ${who}`, `Date: ${when.toISOString().replace('T', ' ').slice(0, 19)} UTC`, opts.reason ? `Reason: ${opts.reason}` : '', opts.location ? `Location: ${opts.location}` : ''].filter(Boolean);
      const content = `q 0.93 0.95 1 rg 0 0 ${w} ${h} re f 0.2 0.35 0.7 RG 1 w 0.5 0.5 ${w - 1} ${h - 1} re S Q BT /Helv 8 Tf 0.1 0.1 0.2 rg ${lines.map((l, i) => `1 0 0 1 6 ${h - 13 - i * 10} Tm (${winAnsi(l).slice(0, 60)}) Tj`).join(' ')} ET`;
      objs.set(apNum, `<< /Type /XObject /Subtype /Form /BBox [0 0 ${w} ${h}] /Resources << /Font << /Helv << /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >> >> >> /Length ${content.length} >>\nstream\n${content}\nendstream`);
    }
    const placeholderRange = '/ByteRange [0 ********** ********** **********]';
    objs.set(sigNum, `<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached ${placeholderRange} /Contents <${'0'.repeat(SIG_BYTES * 2)}> /M ${pdfText(pdfDate(when))} /Name ${pdfText(cert.subject.cn || '')}${opts.reason ? ' /Reason ' + pdfText(opts.reason) : ''}${opts.location ? ' /Location ' + pdfText(opts.location) : ''}${opts.contact ? ' /ContactInfo ' + pdfText(opts.contact) : ''} /Prop_Build << /App << /Name /Isadek#20PDF >> >> >>`);
    objs.set(widgetNum, `<< /Type /Annot /Subtype /Widget /FT /Sig /T ${pdfText(fieldName)} /V ${sigNum} 0 R /F ${opts.visible ? 4 : 132} /Rect [${rect.join(' ')}] /P ${pageRef.objectNumber} ${pageRef.generationNumber} R${apNum ? ` /AP << /N ${apNum} 0 R >>` : ''} >>`);

    // The page lists the new widget among its annotations.
    const widgetRef = PDFRef.of(widgetNum);
    const annots = page.node.get(PDFName.of('Annots'));
    if (annots instanceof PDFRef) {
      const arr = ctx.lookup(annots, PDFArray);
      arr.push(widgetRef);
      objs.set(annots.objectNumber, arr.toString());
    } else {
      if (annots instanceof PDFArray) annots.push(widgetRef); else page.node.set(PDFName.of('Annots'), ctx.obj([widgetRef]));
      objs.set(pageRef.objectNumber, page.node.toString());
    }
    // The form lists the new field, and says it has signatures (SigFlags 3: signatures, append only).
    const catalog = doc.catalog, catRef = ctx.trailerInfo.Root;
    const af = catalog.get(PDFName.of('AcroForm'));
    let afDict = af instanceof PDFRef ? ctx.lookup(af, PDFDict) : af instanceof PDFDict ? af : null;
    if (!afDict) { afDict = ctx.obj({}); catalog.set(PDFName.of('AcroForm'), afDict); }
    let fields = afDict.get(PDFName.of('Fields'));
    if (fields instanceof PDFRef) { const arr = ctx.lookup(fields, PDFArray); arr.push(widgetRef); objs.set(fields.objectNumber, arr.toString()); }
    else { if (!(fields instanceof PDFArray)) { fields = ctx.obj([]); afDict.set(PDFName.of('Fields'), fields); } fields.push(widgetRef); }
    afDict.set(PDFName.of('SigFlags'), ctx.obj(3));
    if (af instanceof PDFRef) objs.set(af.objectNumber, afDict.toString());
    if (!(af instanceof PDFRef) || !hadForm) objs.set(catRef.objectNumber, catalog.toString());

    // The update: the objects, a cross-reference section for them, and a trailer pointing back.
    const header = latin1(src.subarray(src.length - 1)) === '\n' ? '' : '\n';
    const parts = [header];
    let offset = src.length + header.length;
    const xref = [];
    for (const num of [...objs.keys()].sort((a, b) => a - b)) {
      const gen = num === pageRef.objectNumber ? pageRef.generationNumber : 0;
      const s = `${num} ${gen} obj\n${objs.get(num)}\nendobj\n`;
      xref.push([num, offset, gen]);
      parts.push(s);
      offset += s.length;
    }
    const xrefAt = offset;
    const t = ctx.trailerInfo;
    const ref = r => (r ? `${r.objectNumber} ${r.generationNumber} R` : '');
    const id = t.ID ? ` /ID ${t.ID.toString()}` : '';
    const trailer = `/Size ${next + 1} /Root ${ref(catRef)}${t.Info ? ' /Info ' + ref(t.Info) : ''}${id} /Prev ${prev}`;
    let xs;
    if (/^\s*\d+\s+\d+\s+obj/.test(latin1(src.subarray(prev, prev + 40)))) {
      // The file keeps its cross-references in a stream: so does the update (readers expect the same kind).
      const xnum = next;
      const rows = [...xref, [xnum, xrefAt, 0]].sort((a, b) => a[0] - b[0]);
      const data = rows.map(([, off, gen]) => '01' + off.toString(16).padStart(8, '0') + gen.toString(16).padStart(4, '0')).join('');
      xs = `${xnum} 0 obj\n<< /Type /XRef ${trailer} /W [1 4 2] /Index [${rows.map(r => r[0] + ' 1').join(' ')}] /Length ${data.length / 2} >>\nstream\n`;
      const bin = unhex(data);
      parts.push(xs);
      parts.push(latin1(bin));
      xs = `\nendstream\nendobj\nstartxref\n${xrefAt}\n%%EOF\n`;
    } else {
      xs = 'xref\n';
      for (const [num, off, gen] of xref) xs += `${num} 1\n${String(off).padStart(10, '0')} ${String(gen).padStart(5, '0')} n \n`;
      xs += `trailer\n<< ${trailer.replace(`/Size ${next + 1}`, `/Size ${next}`)} >>\nstartxref\n${xrefAt}\n%%EOF\n`;
    }
    parts.push(xs);
    const upd = parts.join('');
    // Offsets above count characters: each must be one byte.
    if (/[^\x00-\xff]/.test(upd)) throw new Error('This PDF has page data that can’t be signed here');
    const update = new Uint8Array(upd.length);
    for (let i = 0; i < upd.length; i++) update[i] = upd.charCodeAt(i);
    const out = concat([src, update]);

    // The gap: the Contents hex string. The byte range is everything else.
    const s = latin1(out.subarray(src.length));
    const ci = s.indexOf('/Contents <') + src.length + '/Contents '.length;
    const gapEnd = ci + SIG_BYTES * 2 + 2;
    const range = [0, ci, gapEnd, out.length - gapEnd];
    const ri = s.indexOf(placeholderRange) + src.length;
    const rangeText = `/ByteRange [${range.join(' ')}]`.padEnd(placeholderRange.length, ' ');
    if (rangeText.length > placeholderRange.length) throw new Error('This PDF is too large to sign');
    out.set(te.encode(rangeText), ri);
    const signed = concat([out.subarray(0, range[1]), out.subarray(range[2])]);
    const sig = await cms(identity, signed, when);
    if (sig.length > SIG_BYTES) throw new Error('The signature is too large');
    out.set(te.encode(hex(sig).toUpperCase()), ci + 1);
    return out;
  }

  /** Every signature in a PDF, checked. Later ones first in the file come first. */
  async function verify(bytes) {
    const b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    const s = latin1(b);
    const out = [];
    for (const m of s.matchAll(/\/ByteRange\s*\[\s*(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s*\]/g)) {
      const r = m.slice(1, 5).map(Number);
      const item = { range: r, ok: false, problems: [] };
      out.push(item);
      try {
        if (r[0] !== 0 || r[1] <= 0 || r[2] <= r[1] || r[2] + r[3] > b.length) throw new Error('The signature’s byte range is wrong');
        const gap = s.slice(r[1], r[2]);
        if (!/^<[0-9a-fA-F\s]*>$/.test(gap)) throw new Error('The signature isn’t where its byte range says');
        let h = gap.slice(1, -1).replace(/\s+/g, '');
        if (h.length % 2) h += '0';
        let der = unhex(h);
        der = der.subarray(0, parse(der).end); // the zeros after it are padding
        const data = concat([b.subarray(0, r[1]), b.subarray(r[2], r[2] + r[3])]);
        Object.assign(item, await checkCms(der, data));
        item.covers = r[2] + r[3];
        // Changes after this signature: the file goes on past what it covers.
        const rest = latin1(b.subarray(item.covers)).replace(/\s+$/, '');
        item.intact = !rest.length;
        item.laterChanges = !!rest.length;
        // The signature dictionary's own fields, from around the byte range.
        const dictStart = s.lastIndexOf('<<', m.index), around = s.slice(dictStart, r[2] + 2000);
        const field = k => { const mm = new RegExp('/' + k + '\\s*(\\((?:\\\\.|[^\\\\)])*\\)|<[0-9A-Fa-f]*>)').exec(around); return mm ? decodeText(mm[1]) : ''; };
        item.reason = field('Reason'); item.location = field('Location'); item.name = field('Name');
        if (!item.time) { const d = /^D:(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})/.exec(field('M')); if (d) item.time = new Date(Date.UTC(+d[1], d[2] - 1, +d[3], +d[4], +d[5], +d[6])); }
      } catch (err) { item.problems.push(err.message); item.ok = false; }
    }
    return out;
  }
  function decodeText(t) {
    if (t[0] === '<') { const b = unhex(t.slice(1, -1)); if (b[0] === 0xfe && b[1] === 0xff) { let s = ''; for (let i = 2; i + 1 < b.length; i += 2) s += String.fromCharCode(b[i] << 8 | b[i + 1]); return s; } return latin1(b); }
    return t.slice(1, -1).replace(/\\([nrtbf()\\]|[0-7]{1,3})/g, (m, c) => ({ n: '\n', r: '\r', t: '\t', b: '\b', f: '\f', '(': '(', ')': ')', '\\': '\\' }[c] ?? String.fromCharCode(parseInt(c, 8))));
  }

  const api = { createIdentity, importIdentity, sign, verify, readCert, fingerprint, certPem, parse, _cms: cms, _checkCms: checkCms };
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.PdfSign = api;
})(typeof window !== 'undefined' ? window : globalThis);
