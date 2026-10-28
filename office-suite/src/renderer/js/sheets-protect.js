'use strict';
// Isadek Sheets — sheet protection, compatible with Excel's <sheetProtection>.
// sheet.protect = { hash, algo?: { name, hash, salt, spin }, allow: { formatCells, insertRows, … } }
// hash is Excel's legacy 16-bit password hash (as Excel writes in the password attribute);
// files from newer Excel use a salted SHA-512 (algo), which is checked with WebCrypto.
(function (root) {
  const ALLOW = ['formatCells', 'formatColumns', 'formatRows', 'insertColumns', 'insertRows', 'deleteColumns', 'deleteRows', 'sort', 'autoFilter', 'pivotTables', 'objects'];
  /** Excel's legacy password hash (four hex digits); '' for no password. */
  function legacyHash(pw) {
    if (!pw) return '';
    let h = 0;
    for (let i = pw.length - 1; i >= 0; i--) { h = ((h >> 14) & 1) | ((h << 1) & 0x7fff); h ^= pw.charCodeAt(i) & 0xff; }
    h = ((h >> 14) & 1) | ((h << 1) & 0x7fff);
    h ^= pw.length; h ^= 0xce4b;
    return h.toString(16).toUpperCase().padStart(4, '0');
  }
  const b64 = bytes => btoa(String.fromCharCode(...bytes));
  const unb64 = s => Uint8Array.from(atob(s), c => c.charCodeAt(0));
  /** Excel 2013+'s hash: SHA-512 of salt + UTF-16LE password, then spinCount rounds with the round number appended. */
  async function modernHash(pw, saltB64, spin) {
    const subtle = (root.crypto || {}).subtle;
    if (!subtle) return null;
    const pwBytes = new Uint8Array(pw.length * 2);
    for (let i = 0; i < pw.length; i++) { pwBytes[i * 2] = pw.charCodeAt(i) & 0xff; pwBytes[i * 2 + 1] = pw.charCodeAt(i) >> 8; }
    const salt = unb64(saltB64);
    let h = new Uint8Array(await subtle.digest('SHA-512', new Uint8Array([...salt, ...pwBytes])));
    const buf = new Uint8Array(68);
    for (let i = 0; i < spin; i++) {
      buf.set(h, 0); buf[64] = i & 0xff; buf[65] = (i >> 8) & 0xff; buf[66] = (i >> 16) & 0xff; buf[67] = (i >>> 24) & 0xff;
      h = new Uint8Array(await subtle.digest('SHA-512', buf));
    }
    return b64(h);
  }
  async function check(protect, pw) {
    if (protect.algo) {
      if (!/SHA-512/i.test(protect.algo.name)) return false;
      return (await modernHash(pw || '', protect.algo.salt, protect.algo.spin)) === protect.algo.hash;
    }
    return (protect.hash || '') === legacyHash(pw || '');
  }
  function clean(p) {
    if (!p || typeof p !== 'object') return null;
    const out = { hash: typeof p.hash === 'string' && /^[0-9A-F]{0,4}$/i.test(p.hash) ? p.hash.toUpperCase() : '', allow: {} };
    const a = p.algo;
    if (a && typeof a === 'object' && typeof a.hash === 'string' && typeof a.salt === 'string' && /^[A-Za-z0-9+/=]{1,200}$/.test(a.hash + a.salt) && Number.isInteger(a.spin) && a.spin >= 0 && a.spin <= 10000000)
      out.algo = { name: String(a.name || 'SHA-512').slice(0, 20), hash: a.hash, salt: a.salt, spin: a.spin };
    ALLOW.forEach(k => { if (p.allow && p.allow[k] === true) out.allow[k] = true; });
    if (p.allow && p.allow.selectLocked === false) out.allow.selectLocked = false;
    return out;
  }
  /** <sheetProtection> for a protected sheet. Excel's attributes say what is protected, so allowed = "0". */
  function toXml(p) {
    if (!p) return '';
    const attrs = [];
    if (p.algo) attrs.push(`algorithmName="${p.algo.name}" hashValue="${p.algo.hash}" saltValue="${p.algo.salt}" spinCount="${p.algo.spin}"`);
    else if (p.hash) attrs.push(`password="${p.hash}"`);
    attrs.push('sheet="1" objects="1" scenarios="1"');
    ALLOW.forEach(k => { if (k !== 'objects' && p.allow[k]) attrs.push(`${k}="0"`); });
    if (p.allow.selectLocked === false) attrs.push('selectLockedCells="1"');
    return `<sheetProtection ${attrs.join(' ')}/>`;
  }
  function fromXml(xml) {
    const m = /<sheetProtection\b([^>]*)\/?>/.exec(xml || '');
    if (!m) return null;
    const attr = n => { const x = new RegExp(`\\b${n}="([^"]*)"`).exec(m[1]); return x ? x[1] : null; };
    if (attr('sheet') !== '1' && attr('sheet') !== 'true') return null;
    const p = { hash: attr('password') || '', allow: {} };
    if (attr('hashValue')) p.algo = { name: attr('algorithmName') || 'SHA-512', hash: attr('hashValue'), salt: attr('saltValue') || '', spin: parseInt(attr('spinCount') || '0', 10) };
    ALLOW.forEach(k => { if (attr(k) === '0' || attr(k) === 'false') p.allow[k] = true; });
    if (attr('selectLockedCells') === '1') p.allow.selectLocked = false;
    return clean(p);
  }
  const api = { legacyHash, modernHash, check, clean, toXml, fromXml, ALLOW };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.SheetsProtect = api;
})(typeof window !== 'undefined' ? window : globalThis);
