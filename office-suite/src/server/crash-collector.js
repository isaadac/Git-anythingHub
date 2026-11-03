'use strict';
// Receives crash reports from the desktop app (when people turn them on) and keeps them in a folder:
//   POST application/json            a script error: appended to errors-YYYY-MM-DD.jsonl
//   POST multipart/form-data         a native crash from Crashpad: the minidump as <id>.dmp and
//                                    its fields as <id>.json
// Reports older than 90 days are deleted. Each address may send PER_HOUR reports an hour.
// Enabled in web-server.js with CRASH_DIR=/path/to/folder; the app is built with
// ISADEK_CRASH_URL=https://your-server/api/crash.
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const zlib = require('zlib');

const MAX = 20e6, KEEP_DAYS = 90, PER_HOUR = 60;
const FIELD = /^[\w.-]{1,64}$/;

function multipart(body, type) {
  const b = (String(type).match(/boundary=(?:"([^"]+)"|([^;]+))/i) || []).slice(1).find(Boolean);
  if (!b) return null;
  const sep = Buffer.from('--' + b), fields = {}, files = {};
  let at = body.indexOf(sep);
  while (at >= 0) {
    const start = at + sep.length;
    if (body.slice(start, start + 2).toString() === '--') break;
    const next = body.indexOf(sep, start);
    if (next < 0) break;
    const part = body.slice(start + 2, next - 2); // without the CRLF after the boundary and before the next
    const head = part.indexOf('\r\n\r\n');
    if (head > 0) {
      const h = part.slice(0, head).toString('utf8'), data = part.slice(head + 4);
      const name = (h.match(/name="([^"]*)"/i) || [])[1];
      if (name && FIELD.test(name)) {
        if (/filename="/i.test(h)) files[name] = data;
        else fields[name] = data.toString('utf8').slice(0, 2000);
      }
    }
    at = next;
  }
  return { fields, files };
}

function createCrashCollector({ dir, perHour = PER_HOUR, now = () => Date.now() }) {
  fs.mkdirSync(dir, { recursive: true });
  const hits = new Map();
  const prune = () => {
    const cut = now() - KEEP_DAYS * 864e5;
    for (const f of fs.readdirSync(dir)) {
      try { const p = path.join(dir, f); if (fs.statSync(p).mtimeMs < cut) fs.unlinkSync(p); } catch { /* gone */ }
    }
    hits.clear();
  };
  prune();
  const timer = setInterval(prune, 36e5);
  timer.unref();

  function read(req) {
    return new Promise((resolve, reject) => {
      const parts = []; let n = 0;
      req.on('data', d => { n += d.length; if (n > MAX) { reject(Object.assign(new Error('too big'), { code: 413 })); req.destroy(); } else parts.push(d); });
      req.on('end', () => resolve(Buffer.concat(parts)));
      req.on('error', reject);
    });
  }
  const reply = (res, code, text) => { res.writeHead(code, { 'Content-Type': 'text/plain; charset=utf-8', 'Cache-Control': 'no-store' }); res.end(text); };

  async function handle(req, res) {
    if (req.method !== 'POST') return reply(res, 405, 'POST only');
    const who = req.socket.remoteAddress || '?';
    const n = (hits.get(who) || 0) + 1;
    hits.set(who, n);
    if (n > perHour) return reply(res, 429, 'Too many reports');
    let body;
    try {
      body = await read(req);
      // Crashpad compresses its uploads.
      if (/gzip/i.test(req.headers['content-encoding'] || '')) body = zlib.gunzipSync(body, { maxOutputLength: MAX * 4 });
    } catch (e) { return reply(res, e.code || 400, 'Bad report'); }
    const type = String(req.headers['content-type'] || '');
    const id = new Date(now()).toISOString().replace(/[:.]/g, '-') + '-' + crypto.randomBytes(4).toString('hex');
    if (/^application\/json/i.test(type)) {
      let r;
      try { r = JSON.parse(body.toString('utf8')); } catch { return reply(res, 400, 'Bad report'); }
      if (!r || typeof r !== 'object' || Array.isArray(r)) return reply(res, 400, 'Bad report');
      const keep = {};
      for (const [k, v] of Object.entries(r).slice(0, 30)) if (FIELD.test(k) && (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean')) keep[k] = typeof v === 'string' ? v.slice(0, 8000) : v;
      fs.appendFileSync(path.join(dir, `errors-${id.slice(0, 10)}.jsonl`), JSON.stringify({ id, received: new Date(now()).toISOString(), ...keep }) + '\n');
      return reply(res, 200, id);
    }
    if (/^multipart\/form-data/i.test(type)) {
      const m = multipart(body, type);
      const dump = m && (m.files.upload_file_minidump || Object.values(m.files)[0]);
      if (!dump || dump.slice(0, 4).toString('latin1') !== 'MDMP') return reply(res, 400, 'Bad report');
      fs.writeFileSync(path.join(dir, id + '.dmp'), dump);
      fs.writeFileSync(path.join(dir, id + '.json'), JSON.stringify({ id, received: new Date(now()).toISOString(), ...m.fields }, null, 1));
      return reply(res, 200, id); // Crashpad shows this as the report's id
    }
    return reply(res, 415, 'JSON or multipart only');
  }
  return { handle, prune, close: () => clearInterval(timer) };
}

module.exports = { createCrashCollector, multipart };
