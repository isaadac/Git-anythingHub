'use strict';
// Password-protected Office documents (.docx, .xlsx, .pptx): reading and writing the encryption
// Word, Excel and PowerPoint use since Office 2010 — ECMA-376 "Agile" encryption, AES-256 with
// SHA-512 and 100,000 rounds of key stretching, in a Compound File container [MS-OFFCRYPTO].
//
// OfficeCrypto.isEncrypted(bytes)      → is this a password-protected Office file?
// OfficeCrypto.decrypt(bytes, password) → the plain .docx/.xlsx/.pptx bytes (throws BadPassword)
// OfficeCrypto.encrypt(bytes, password) → the protected file's bytes
//
// Uses the platform's Web Crypto (browsers, Electron, Node) and SheetJS's CFB reader/writer.
(function (root) {
  const CFB = root.CFB || (typeof require === 'function' ? require('cfb') : null);
  const subtle = (root.crypto || (typeof require === 'function' ? require('crypto').webcrypto : null)).subtle;
  const rand = n => (root.crypto || require('crypto').webcrypto).getRandomValues(new Uint8Array(n));

  const BLOCK = { verifierInput: [0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79], verifierValue: [0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e],
    key: [0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6], hmacKey: [0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6], hmacValue: [0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33] };
  const HASHES = { SHA512: 'SHA-512', SHA384: 'SHA-384', SHA256: 'SHA-256', SHA1: 'SHA-1' };
  const SEGMENT = 4096;

  class BadPassword extends Error { constructor() { super('The password is not correct'); this.name = 'BadPassword'; } }

  const concat = (...parts) => { const n = parts.reduce((a, p) => a + p.length, 0), out = new Uint8Array(n); let at = 0; for (const p of parts) { out.set(p, at); at += p.length; } return out; };
  const u32 = n => { const b = new Uint8Array(4); new DataView(b.buffer).setUint32(0, n, true); return b; };
  const b64 = bytes => { let s = ''; for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000)); return btoa(s); };
  const unb64 = s => { const t = atob(s); const b = new Uint8Array(t.length); for (let i = 0; i < t.length; i++) b[i] = t.charCodeAt(i); return b; };
  const utf16 = s => { const b = new Uint8Array(s.length * 2); for (let i = 0; i < s.length; i++) { const c = s.charCodeAt(i); b[2 * i] = c & 255; b[2 * i + 1] = c >> 8; } return b; };
  const sameBytes = (a, b) => a.length === b.length && a.every((x, i) => x === b[i]);
  /** Bytes made exactly n long: cut, or filled with 0x36 as the format says. */
  const fit = (b, n) => { if (b.length >= n) return b.slice(0, n); const out = new Uint8Array(n).fill(0x36); out.set(b); return out; };
  const padTo = (b, n) => (b.length % n ? concat(b, new Uint8Array(n - (b.length % n))) : b);
  const hash = async (alg, ...parts) => new Uint8Array(await subtle.digest(alg, concat(...parts)));

  // AES-CBC without padding, which Web Crypto doesn't offer directly: encrypting drops the padding
  // block it adds; decrypting first appends the padding block that would have been there.
  const aesKey = (raw, use) => subtle.importKey('raw', raw, { name: 'AES-CBC' }, false, [use]);
  async function aesEncrypt(rawKey, iv, data) {
    const k = await aesKey(rawKey, 'encrypt');
    return new Uint8Array(await subtle.encrypt({ name: 'AES-CBC', iv }, k, data)).slice(0, data.length);
  }
  async function aesDecrypt(rawKey, iv, data) {
    if (!data.length) return new Uint8Array(0);
    const ek = await aesKey(rawKey, 'encrypt'), dk = await aesKey(rawKey, 'decrypt');
    const last = data.slice(data.length - 16);
    const pad = new Uint8Array(await subtle.encrypt({ name: 'AES-CBC', iv: last }, ek, new Uint8Array(0)));
    return new Uint8Array(await subtle.decrypt({ name: 'AES-CBC', iv }, dk, concat(data, pad)));
  }

  // SHA-512 in plain JavaScript, for the 100,000 rounds of key stretching: one round at a time
  // through Web Crypto's asynchronous digest takes seconds; this takes a fraction of one.
  const K = (() => {
    const hex = '428a2f98d728ae22 7137449123ef65cd b5c0fbcfec4d3b2f e9b5dba58189dbbc 3956c25bf348b538 59f111f1b605d019 923f82a4af194f9b ab1c5ed5da6d8118 d807aa98a3030242 12835b0145706fbe 243185be4ee4b28c 550c7dc3d5ffb4e2 72be5d74f27b896f 80deb1fe3b1696b1 9bdc06a725c71235 c19bf174cf692694 e49b69c19ef14ad2 efbe4786384f25e3 0fc19dc68b8cd5b5 240ca1cc77ac9c65 2de92c6f592b0275 4a7484aa6ea6e483 5cb0a9dcbd41fbd4 76f988da831153b5 983e5152ee66dfab a831c66d2db43210 b00327c898fb213f bf597fc7beef0ee4 c6e00bf33da88fc2 d5a79147930aa725 06ca6351e003826f 142929670a0e6e70 27b70a8546d22ffc 2e1b21385c26c926 4d2c6dfc5ac42aed 53380d139d95b3df 650a73548baf63de 766a0abb3c77b2a8 81c2c92e47edaee6 92722c851482353b a2bfe8a14cf10364 a81a664bbc423001 c24b8b70d0f89791 c76c51a30654be30 d192e819d6ef5218 d69906245565a910 f40e35855771202a 106aa07032bbd1b8 19a4c116b8d2d0c8 1e376c085141ab53 2748774cdf8eeb99 34b0bcb5e19b48a8 391c0cb3c5c95a63 4ed8aa4ae3418acb 5b9cca4f7763e373 682e6ff3d6b2b8a3 748f82ee5defb2fc 78a5636f43172f60 84c87814a1f0ab72 8cc702081a6439ec 90befffa23631e28 a4506cebde82bde9 bef9a3f7b2c67915 c67178f2e372532b ca273eceea26619c d186b8c721c0c207 eada7dd6cde0eb1e f57d4f7fee6ed178 06f067aa72176fba 0a637dc5a2c898a6 113f9804bef90dae 1b710b35131c471b 28db77f523047d84 32caab7b40c72493 3c9ebe0a15c9bebc 431d67c49c100d4c 4cc5d4becb3e42b6 597f299cfc657e2a 5fcb6fab3ad6faec 6c44198c4a475817'.split(' ');
    const k = new Int32Array(160);
    hex.forEach((h, i) => { k[2 * i] = parseInt(h.slice(0, 8), 16) | 0; k[2 * i + 1] = parseInt(h.slice(8), 16) | 0; });
    return k;
  })();
  const H0 = [0x6a09e667, 0xf3bcc908, 0xbb67ae85, 0x84caa73b, 0x3c6ef372, 0xfe94f82b, 0xa54ff53a, 0x5f1d36f1, 0x510e527f, 0xade682d1, 0x9b05688c, 0x2b3e6c1f, 0x1f83d9ab, 0xfb41bd6b, 0x5be0cd19, 0x137e2179];
  const W = new Int32Array(160);
  function sha512(msg) {
    const len = msg.length, blocks = Math.ceil((len + 17) / 128), buf = new Uint8Array(blocks * 128);
    buf.set(msg); buf[len] = 0x80;
    const bits = len * 8, end = blocks * 128;
    buf[end - 4] = (bits >>> 24) & 255; buf[end - 3] = (bits >>> 16) & 255; buf[end - 2] = (bits >>> 8) & 255; buf[end - 1] = bits & 255;
    buf[end - 5] = Math.floor(bits / 2 ** 32) & 255;
    const h = H0.map(x => x | 0);
    for (let b = 0; b < blocks; b++) {
      const o = b * 128;
      for (let i = 0; i < 32; i++) W[i] = (buf[o + 4 * i] << 24) | (buf[o + 4 * i + 1] << 16) | (buf[o + 4 * i + 2] << 8) | buf[o + 4 * i + 3];
      for (let i = 16; i < 80; i++) {
        // σ0 of W[i-15], σ1 of W[i-2] (64-bit words as high, low 32-bit halves)
        let xh = W[2 * (i - 15)], xl = W[2 * (i - 15) + 1];
        const s0h = ((xh >>> 1) | (xl << 31)) ^ ((xh >>> 8) | (xl << 24)) ^ (xh >>> 7);
        const s0l = ((xl >>> 1) | (xh << 31)) ^ ((xl >>> 8) | (xh << 24)) ^ ((xl >>> 7) | (xh << 25));
        xh = W[2 * (i - 2)]; xl = W[2 * (i - 2) + 1];
        const s1h = ((xh >>> 19) | (xl << 13)) ^ ((xl >>> 29) | (xh << 3)) ^ (xh >>> 6);
        const s1l = ((xl >>> 19) | (xh << 13)) ^ ((xh >>> 29) | (xl << 3)) ^ ((xl >>> 6) | (xh << 26));
        let lo = (s0l >>> 0) + (s1l >>> 0) + (W[2 * (i - 7) + 1] >>> 0) + (W[2 * (i - 16) + 1] >>> 0);
        const hi = s0h + s1h + W[2 * (i - 7)] + W[2 * (i - 16)] + Math.floor(lo / 4294967296);
        W[2 * i] = hi | 0; W[2 * i + 1] = lo | 0;
      }
      let [ah, al, bh, bl, ch, cl, dh, dl, eh, el, fh, fl, gh, gl, hh, hl] = h;
      for (let i = 0; i < 80; i++) {
        const S1h = ((eh >>> 14) | (el << 18)) ^ ((eh >>> 18) | (el << 14)) ^ ((el >>> 9) | (eh << 23));
        const S1l = ((el >>> 14) | (eh << 18)) ^ ((el >>> 18) | (eh << 14)) ^ ((eh >>> 9) | (el << 23));
        const chh = (eh & fh) ^ (~eh & gh), chl = (el & fl) ^ (~el & gl);
        let t1l = (hl >>> 0) + (S1l >>> 0) + (chl >>> 0) + (K[2 * i + 1] >>> 0) + (W[2 * i + 1] >>> 0);
        const t1h = (hh + S1h + chh + K[2 * i] + W[2 * i] + Math.floor(t1l / 4294967296)) | 0;
        t1l = t1l | 0;
        const S0h = ((ah >>> 28) | (al << 4)) ^ ((al >>> 2) | (ah << 30)) ^ ((al >>> 7) | (ah << 25));
        const S0l = ((al >>> 28) | (ah << 4)) ^ ((ah >>> 2) | (al << 30)) ^ ((ah >>> 7) | (al << 25));
        const mjh = (ah & bh) ^ (ah & ch) ^ (bh & ch), mjl = (al & bl) ^ (al & cl) ^ (bl & cl);
        const t2l = (S0l >>> 0) + (mjl >>> 0), t2h = (S0h + mjh + Math.floor(t2l / 4294967296)) | 0;
        hh = gh; hl = gl; gh = fh; gl = fl; fh = eh; fl = el;
        const el2 = (dl >>> 0) + (t1l >>> 0);
        eh = (dh + t1h + Math.floor(el2 / 4294967296)) | 0; el = el2 | 0;
        dh = ch; dl = cl; ch = bh; cl = bl; bh = ah; bl = al;
        const al2 = (t1l >>> 0) + (t2l >>> 0);
        ah = (t1h + t2h + Math.floor(al2 / 4294967296)) | 0; al = al2 | 0;
      }
      const add = [ah, al, bh, bl, ch, cl, dh, dl, eh, el, fh, fl, gh, gl, hh, hl];
      for (let i = 0; i < 16; i += 2) {
        const lo = (h[i + 1] >>> 0) + (add[i + 1] >>> 0);
        h[i] = (h[i] + add[i] + Math.floor(lo / 4294967296)) | 0; h[i + 1] = lo | 0;
      }
    }
    const out = new Uint8Array(64);
    for (let i = 0; i < 16; i++) { out[4 * i] = h[i] >>> 24; out[4 * i + 1] = (h[i] >>> 16) & 255; out[4 * i + 2] = (h[i] >>> 8) & 255; out[4 * i + 3] = h[i] & 255; }
    return out;
  }

  /** The password's key for one purpose (blockKey), after spinCount rounds of hashing. */
  async function passwordHash(password, salt, alg, spinCount) {
    let h = await hash(alg, salt, utf16(password));
    if (alg === 'SHA-512') {
      const m = new Uint8Array(68);
      for (let i = 0; i < spinCount; i++) { m[0] = i & 255; m[1] = (i >>> 8) & 255; m[2] = (i >>> 16) & 255; m[3] = i >>> 24; m.set(h, 4); h = sha512(m); }
      return h;
    }
    for (let i = 0; i < spinCount; i++) h = await hash(alg, u32(i), h);
    return h;
  }
  const derive = async (alg, h, blockKey, keyBytes) => fit(await hash(alg, h, new Uint8Array(blockKey)), keyBytes);

  // ---------- Reading ----------
  function streams(bytes) {
    let cfb;
    try { cfb = CFB.read(bytes, { type: 'array' }); } catch { return null; }
    const get = name => { const e = CFB.find(cfb, name); return e && e.content ? new Uint8Array(e.content) : null; };
    const info = get('EncryptionInfo'), pkg = get('EncryptedPackage');
    return info && pkg ? { info, pkg } : null;
  }
  const isCompound = b => b && b.length > 8 && b[0] === 0xd0 && b[1] === 0xcf && b[2] === 0x11 && b[3] === 0xe0;
  /** Is this an Office document protected with a password? */
  function isEncrypted(bytes) {
    const b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    return isCompound(b) && !!streams(b);
  }

  function parseInfo(info) {
    const major = info[0] | (info[1] << 8), minor = info[2] | (info[3] << 8);
    if (major !== 4 || minor !== 4) throw new Error(major === 3 || major === 4 ? 'This file uses an older kind of Office encryption (Office 2007). Open it in Office and save it again, then try again.' : 'This file is protected in a way Isadek Office can\'t open');
    const xml = new TextDecoder().decode(info.subarray(8));
    const doc = new DOMParserLike(xml);
    const kd = doc.attrs('keyData'), ke = doc.attrs('encryptedKey'), di = doc.attrs('dataIntegrity');
    if (!kd || !ke) throw new Error('The file\'s encryption information is damaged');
    const num = v => parseInt(v, 10);
    return {
      keyData: { salt: unb64(kd.saltValue), alg: HASHES[kd.hashAlgorithm], blockSize: num(kd.blockSize), keyBits: num(kd.keyBits), hashSize: num(kd.hashSize), cipher: kd.cipherAlgorithm, chaining: kd.cipherChaining },
      key: { salt: unb64(ke.saltValue), alg: HASHES[ke.hashAlgorithm], spinCount: num(ke.spinCount), keyBits: num(ke.keyBits), hashSize: num(ke.hashSize),
        verifierInput: unb64(ke.encryptedVerifierHashInput), verifierValue: unb64(ke.encryptedVerifierHashValue), keyValue: unb64(ke.encryptedKeyValue), cipher: ke.cipherAlgorithm },
      integrity: di ? { key: unb64(di.encryptedHmacKey), value: unb64(di.encryptedHmacValue) } : null
    };
  }
  // A tiny attribute reader for the EncryptionInfo XML (works the same in pages, workers and Node).
  function DOMParserLike(xml) {
    this.attrs = local => {
      const m = new RegExp(`<(?:\\w+:)?${local}\\b([^>]*)>`).exec(xml);
      if (!m) return null;
      const out = {};
      for (const a of m[1].matchAll(/([\w:]+)="([^"]*)"/g)) out[a[1].replace(/^\w+:/, '')] = a[2];
      return out;
    };
  }

  /** The plain document from a password-protected one. Throws BadPassword for a wrong password. */
  async function decrypt(bytes, password) {
    const s = streams(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes));
    if (!s) throw new Error('This file is not a password-protected Office document');
    const info = parseInfo(s.info);
    const k = info.key, kd = info.keyData;
    if (k.cipher !== 'AES' || kd.cipher !== 'AES' || !k.alg || !kd.alg || kd.chaining !== 'ChainingModeCBC') throw new Error('This file is protected in a way Isadek Office can\'t open');
    const keyBytes = k.keyBits / 8;
    const h = await passwordHash(String(password), k.salt, k.alg, k.spinCount);
    // Is the password right? Its verifier decrypts to a value whose hash is stored alongside.
    const input = (await aesDecrypt(await derive(k.alg, h, BLOCK.verifierInput, keyBytes), k.salt, k.verifierInput)).slice(0, k.salt.length);
    const value = await aesDecrypt(await derive(k.alg, h, BLOCK.verifierValue, keyBytes), k.salt, k.verifierValue);
    if (!sameBytes((await hash(k.alg, input)).slice(0, k.hashSize), value.slice(0, k.hashSize))) throw new BadPassword();
    const secret = (await aesDecrypt(await derive(k.alg, h, BLOCK.key, keyBytes), k.salt, k.keyValue)).slice(0, kd.keyBits / 8);
    // The package: its size, then 4096-byte segments, each with its own IV.
    const size = Number(new DataView(s.pkg.buffer, s.pkg.byteOffset, 8).getBigUint64(0, true));
    const body = s.pkg.subarray(8);
    const parts = [];
    for (let i = 0, at = 0; at < body.length; i++, at += SEGMENT) {
      const iv = (await hash(kd.alg, kd.salt, u32(i))).slice(0, kd.blockSize);
      parts.push(await aesDecrypt(secret, iv, body.subarray(at, Math.min(at + SEGMENT, body.length))));
    }
    const out = concat(...parts).slice(0, size);
    // Untouched since it was saved? (The HMAC Office writes over the encrypted package.)
    if (info.integrity) {
      const hmacKeyIv = (await hash(kd.alg, kd.salt, new Uint8Array(BLOCK.hmacKey))).slice(0, kd.blockSize);
      const hmacValueIv = (await hash(kd.alg, kd.salt, new Uint8Array(BLOCK.hmacValue))).slice(0, kd.blockSize);
      const hmacKey = (await aesDecrypt(secret, hmacKeyIv, info.integrity.key)).slice(0, kd.hashSize);
      const expected = (await aesDecrypt(secret, hmacValueIv, info.integrity.value)).slice(0, kd.hashSize);
      const mac = new Uint8Array(await subtle.sign('HMAC', await subtle.importKey('raw', hmacKey, { name: 'HMAC', hash: kd.alg }, false, ['sign']), s.pkg));
      if (!sameBytes(mac, expected)) throw new Error('The file has been changed or damaged since it was saved');
    }
    return out;
  }

  // ---------- Writing ----------
  const lpUnicode = s => { const b = utf16(s); return padTo(concat(u32(b.length), b), 4); };
  /** The \x06DataSpaces storage Office expects next to an encrypted package [MS-OFFCRYPTO 2.1]. */
  function dataSpaces() {
    const v1 = u32(1); // a version 1.0: major 1, minor 0, as two 16-bit numbers
    const version = concat(lpUnicode('Microsoft.Container.DataSpaces'), v1, v1, v1); // reader, updater, writer
    const entry = concat(u32(1), u32(0), lpUnicode('EncryptedPackage'), lpUnicode('StrongEncryptionDataSpace'));
    const map = concat(u32(8), u32(1), u32(entry.length + 4), entry);
    const space = concat(u32(8), u32(1), lpUnicode('StrongEncryptionTransform'));
    const id = lpUnicode('{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}');
    // Transform header (its length counts the bytes before the name), then the encryption
    // transform: no name, block size 0, cipher mode 0, reserved 4.
    const primary = concat(u32(8 + id.length), u32(1), id, lpUnicode('Microsoft.Container.EncryptionTransform'), v1, v1, v1, u32(0), u32(0), u32(0), u32(4));
    return { version, map, space, primary };
  }

  /** A password-protected copy of a .docx/.xlsx/.pptx. */
  async function encrypt(bytes, password, opts = {}) {
    const plain = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    const alg = 'SHA-512', spinCount = opts.spinCount || 100000, keyBytes = 32, hashSize = 64, blockSize = 16;
    const kdSalt = rand(16), pwSalt = rand(16), secret = rand(keyBytes), verifier = rand(16);
    const h = await passwordHash(String(password), pwSalt, alg, spinCount);
    const encVerifierInput = await aesEncrypt(await derive(alg, h, BLOCK.verifierInput, keyBytes), pwSalt, verifier);
    const encVerifierValue = await aesEncrypt(await derive(alg, h, BLOCK.verifierValue, keyBytes), pwSalt, await hash(alg, verifier));
    const encKey = await aesEncrypt(await derive(alg, h, BLOCK.key, keyBytes), pwSalt, secret);
    // The package, in 4096-byte segments.
    const parts = [];
    for (let i = 0, at = 0; at < plain.length; i++, at += SEGMENT) {
      const iv = (await hash(alg, kdSalt, u32(i))).slice(0, blockSize);
      parts.push(await aesEncrypt(secret, iv, padTo(plain.subarray(at, Math.min(at + SEGMENT, plain.length)), blockSize)));
    }
    const sizeField = new Uint8Array(8);
    new DataView(sizeField.buffer).setBigUint64(0, BigInt(plain.length), true);
    const pkg = concat(sizeField, ...parts);
    // Integrity: an HMAC of the encrypted package, so tampering is noticed when it's opened.
    const hmacKey = rand(hashSize);
    const mac = new Uint8Array(await subtle.sign('HMAC', await subtle.importKey('raw', hmacKey, { name: 'HMAC', hash: alg }, false, ['sign']), pkg));
    const encHmacKey = await aesEncrypt(secret, (await hash(alg, kdSalt, new Uint8Array(BLOCK.hmacKey))).slice(0, blockSize), hmacKey);
    const encHmacValue = await aesEncrypt(secret, (await hash(alg, kdSalt, new Uint8Array(BLOCK.hmacValue))).slice(0, blockSize), mac);
    const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      '<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password" xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">' +
      `<keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="${b64(kdSalt)}"/>` +
      `<dataIntegrity encryptedHmacKey="${b64(encHmacKey)}" encryptedHmacValue="${b64(encHmacValue)}"/>` +
      '<keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">' +
      `<p:encryptedKey spinCount="${spinCount}" saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="${b64(pwSalt)}" encryptedVerifierHashInput="${b64(encVerifierInput)}" encryptedVerifierHashValue="${b64(encVerifierValue)}" encryptedKeyValue="${b64(encKey)}"/>` +
      '</keyEncryptor></keyEncryptors></encryption>';
    const info = concat(new Uint8Array([4, 0, 4, 0, 0x40, 0, 0, 0]), new TextEncoder().encode(xml));
    const ds = dataSpaces();
    const cfb = CFB.utils.cfb_new();
    CFB.utils.cfb_add(cfb, 'EncryptionInfo', info);
    CFB.utils.cfb_add(cfb, 'EncryptedPackage', pkg);
    CFB.utils.cfb_add(cfb, '\u0006DataSpaces/Version', ds.version);
    CFB.utils.cfb_add(cfb, '\u0006DataSpaces/DataSpaceMap', ds.map);
    CFB.utils.cfb_add(cfb, '\u0006DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace', ds.space);
    CFB.utils.cfb_add(cfb, '\u0006DataSpaces/TransformInfo/StrongEncryptionTransform/\u0006Primary', ds.primary);
    // CFB.utils.cfb_new adds a placeholder stream; it isn't part of an encrypted document.
    CFB.utils.cfb_del(cfb, '\u0001Sh33tJ5');
    return new Uint8Array(CFB.write(cfb, { type: 'array' }));
  }

  const api = { isEncrypted, decrypt, encrypt, BadPassword, sha512 };
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.OfficeCrypto = api;
})(typeof window !== 'undefined' ? window : globalThis);
