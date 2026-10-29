'use strict';
// Careful file writes and web requests for the main process.
//   writeAtomic: a file is replaced whole or not at all (a crash, a full disk or a lost drive
//     mid-save never leaves half a document).
//   safeGet: a GET for "data from the web" that can't be pointed at this computer or the cloud
//     metadata service (also after redirects or DNS tricks), with a size and time limit.
//   readCapped: a fetch Response's body, refused as soon as it passes a size limit.
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const dns = require('dns');
const net = require('net');
const http = require('http');
const https = require('https');

/** Writes `data` to `target` through a temporary file in the same folder, then renames it. */
function writeAtomic(target, data, opts = {}) {
  const dir = path.dirname(target);
  const tmp = path.join(dir, `.${path.basename(target)}.${crypto.randomBytes(6).toString('hex')}.tmp`);
  let mode = opts.mode;
  try { if (mode === undefined) mode = fs.statSync(target).mode & 0o777; } catch { /* a new file */ }
  let fd = null;
  try {
    fd = fs.openSync(tmp, 'wx', mode === undefined ? 0o666 : mode);
    fs.writeFileSync(fd, data);
    fs.fsyncSync(fd);
    fs.closeSync(fd); fd = null;
    fs.renameSync(tmp, target);
  } catch (err) {
    if (fd !== null) { try { fs.closeSync(fd); } catch { /* closing anyway */ } }
    try { fs.unlinkSync(tmp); } catch { /* never made */ }
    // Folders where new files can't be made (but this one can be written): write in place.
    if (err.code === 'EACCES' || err.code === 'EPERM' || err.code === 'EROFS' || err.code === 'EXDEV') { fs.writeFileSync(target, data, mode === undefined ? undefined : { mode }); return; }
    throw err;
  }
}

/** JSON from a file, or `fallback` when it is missing, unreadable or not what `valid` accepts. */
function readJson(file, valid, fallback) {
  try { const v = JSON.parse(fs.readFileSync(file, 'utf8')); return valid(v) ? v : fallback; } catch { return fallback; }
}

// ---------- Addresses web data may not come from ----------
/** True for this computer, link-local (including the cloud metadata service), unspecified or multicast addresses. */
function forbidden(ip) {
  if (net.isIPv4(ip)) {
    const [a, b] = ip.split('.').map(Number);
    // 100.100.100.200: Alibaba Cloud's metadata service (others use 169.254.169.254).
    return a === 127 || a === 0 || (a === 169 && b === 254) || a >= 224 || ip === '100.100.100.200';
  }
  if (net.isIPv6(ip)) {
    const v = ip.toLowerCase();
    // An IPv4 address inside IPv6 (::ffff:127.0.0.1, which URLs write as ::ffff:7f00:1).
    const mapped = v.match(/^::ffff:(\d+\.\d+\.\d+\.\d+)$/);
    if (mapped) return forbidden(mapped[1]);
    const hex = v.match(/^::ffff:([0-9a-f]{1,4}):([0-9a-f]{1,4})$/);
    if (hex) { const hi = parseInt(hex[1], 16), lo = parseInt(hex[2], 16); return forbidden(`${hi >> 8}.${hi & 255}.${lo >> 8}.${lo & 255}`); }
    return v === '::' || v === '::1' || /^fe[89ab]/.test(v) || /^ff/.test(v) || v.startsWith('fd00:ec2::254');
  }
  return true;
}
/** dns.lookup that refuses forbidden addresses (checked at connect time, so DNS rebinding can't slip past). */
function guardedLookup(host, opts, cb) {
  dns.lookup(host, { ...opts, all: true }, (err, list) => {
    if (err) return cb(err);
    const ok = list.filter(a => !forbidden(a.address));
    if (!ok.length) return cb(Object.assign(new Error(`${host} is on this computer or a private service address, which web data can't come from`), { code: 'EFORBIDDEN' }));
    if (opts && opts.all) return cb(null, ok);
    cb(null, ok[0].address, ok[0].family);
  });
}

/**
 * GET url (http or https) following up to 5 redirects, each checked again. Resolves to
 * { bytes, contentType } or rejects with an Error whose message can be shown.
 */
function safeGet(url, { maxBytes = 50e6, timeoutMs = 30000, headers = {} } = {}) {
  return new Promise((resolve, reject) => {
    const deadline = setTimeout(() => fail(new Error('The site took too long to answer')), timeoutMs);
    let req = null, done = false;
    const fail = err => { if (done) return; done = true; clearTimeout(deadline); if (req) req.destroy(); reject(err); };
    const go = (u, hops) => {
      let target;
      try { target = new URL(u); } catch { return fail(new Error('That is not a web address')); }
      if (!/^https?:$/.test(target.protocol)) return fail(new Error('Only http and https addresses can be used'));
      if (target.username || target.password) return fail(new Error('Addresses with a user name or password can\'t be used'));
      // An address written as a number is checked here (no lookup happens for it).
      const host = target.hostname.replace(/^\[|\]$/g, '');
      if (net.isIP(host) && forbidden(host)) return fail(new Error(`${host} is on this computer or a private service address, which web data can't come from`));
      const lib = target.protocol === 'https:' ? https : http;
      req = lib.get(target, { lookup: guardedLookup, headers: { 'Accept-Encoding': 'identity', ...headers } }, res => {
        if ([301, 302, 303, 307, 308].includes(res.statusCode) && res.headers.location) {
          res.resume();
          if (hops >= 5) return fail(new Error('The site redirected too many times'));
          return go(new URL(res.headers.location, target).href, hops + 1);
        }
        if (res.statusCode < 200 || res.statusCode >= 300) { res.resume(); return fail(new Error(`The site answered ${res.statusCode}`)); }
        if (+res.headers['content-length'] > maxBytes) { res.resume(); return fail(new Error(`That file is too large (over ${Math.round(maxBytes / 1e6)} MB)`)); }
        const parts = []; let n = 0;
        res.on('data', d => { n += d.length; if (n > maxBytes) fail(new Error(`That file is too large (over ${Math.round(maxBytes / 1e6)} MB)`)); else parts.push(d); });
        res.on('end', () => { if (done) return; done = true; clearTimeout(deadline); resolve({ bytes: new Uint8Array(Buffer.concat(parts)), contentType: String(res.headers['content-type'] || '') }); });
        res.on('error', fail);
      });
      req.on('error', err => fail(err.code === 'EFORBIDDEN' ? err : new Error('Could not reach the site: ' + (err.message || err))));
    };
    go(url, 0);
  });
}

/** A fetch Response's body as bytes, refused (with an error) once it is larger than maxBytes. */
async function readCapped(res, maxBytes) {
  if (+res.headers.get('content-length') > maxBytes) { try { await res.body.cancel(); } catch { /* gone */ } throw new Error('response too large'); }
  if (!res.body) return new Uint8Array(0);
  const reader = res.body.getReader();
  const parts = []; let n = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    n += value.length;
    if (n > maxBytes) { try { await reader.cancel(); } catch { /* gone */ } throw new Error('response too large'); }
    parts.push(value);
  }
  const out = new Uint8Array(n);
  let at = 0;
  for (const p of parts) { out.set(p, at); at += p.length; }
  return out;
}

module.exports = { writeAtomic, readJson, forbidden, safeGet, readCapped };
