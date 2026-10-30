'use strict';
// Isadek Forms — collecting answers. A form is published with its definition and gets an id
// (the link people answer at, unguessable) and an owner key (needed to change it, read the
// answers or delete it; only its hash is kept). Answers are checked against the form as they
// arrive and kept in order, numbered, so Sheets can fetch the ones it hasn't got yet.
//
// createFormsStore({ dir }) keeps forms in a folder (one .json and one .ndjson per form), or in
// memory without one. store.handle(req, res, url) answers /api/forms/… and /f/<id>; the web
// server uses it, and the desktop app runs createFormsServer() to collect answers on the local
// network: just the answer page and the API.
const http = require('http');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const FormsCore = require('../renderer/js/forms-core');

const LIMITS = { forms: 10000, responses: 50000, storedBytes: 100e6, formBytes: 512e3, answerBytes: 256e3, perMinute: 30, createPerHour: 60 };
const ID_RE = /^[A-Za-z0-9_-]{22}$/;
const token = n => crypto.randomBytes(n).toString('base64url');
const hash = k => crypto.createHash('sha256').update(String(k)).digest('hex');
const same = (a, b) => { const x = Buffer.from(a), y = Buffer.from(b); return x.length === y.length && crypto.timingSafeEqual(x, y); };

function createFormsStore(opts = {}) {
  const limits = { ...LIMITS, ...(opts.limits || {}) };
  const dir = opts.dir ? path.resolve(opts.dir) : null;
  if (dir) fs.mkdirSync(dir, { recursive: true, mode: 0o700 });
  const forms = new Map(); // id -> { form, keyHash, open, created, count, responses (in memory only) }

  // ---------- Keeping ----------
  const file = (id, ext) => path.join(dir, id + ext);
  function saveMeta(id, f) {
    if (!dir) return;
    const tmp = file(id, '.json.tmp');
    fs.writeFileSync(tmp, JSON.stringify({ form: f.form, keyHash: f.keyHash, open: f.open, created: f.created }), { mode: 0o600 });
    fs.renameSync(tmp, file(id, '.json'));
  }
  function load(id) {
    if (forms.has(id)) return forms.get(id);
    if (!dir || !ID_RE.test(id)) return null;
    let meta;
    try { meta = JSON.parse(fs.readFileSync(file(id, '.json'), 'utf8')); } catch { return null; }
    const form = FormsCore.clean(meta.form);
    if (!form || typeof meta.keyHash !== 'string') return null;
    let count = 0, bytes = 0;
    try { const text = fs.readFileSync(file(id, '.ndjson'), 'utf8'); count = text ? text.split('\n').filter(Boolean).length : 0; bytes = Buffer.byteLength(text); } catch { /* none yet */ }
    const f = { form, keyHash: meta.keyHash, open: meta.open !== false, created: meta.created || 0, count, bytes };
    forms.set(id, f);
    return f;
  }
  const total = () => (dir ? fs.readdirSync(dir).filter(n => n.endsWith('.json')).length : forms.size);

  // ---------- The operations (also called directly by the desktop app) ----------
  function create(def) {
    const form = FormsCore.clean(def);
    if (!form || !form.questions.length) throw Object.assign(new Error('A form needs at least one question'), { status: 400 });
    if (total() >= limits.forms) throw Object.assign(new Error('This server holds as many forms as it can'), { status: 507 });
    const id = token(16), key = token(32);
    const f = { form, keyHash: hash(key), open: true, created: Date.now(), count: 0, bytes: 0, responses: dir ? null : [] };
    forms.set(id, f);
    saveMeta(id, f);
    return { id, key };
  }
  function owned(id, key) {
    const f = ID_RE.test(String(id)) ? load(id) : null;
    if (!f) throw Object.assign(new Error('There is no such form'), { status: 404 });
    if (typeof key !== 'string' || !same(hash(key), f.keyHash)) throw Object.assign(new Error('That key doesn’t open this form'), { status: 403 });
    return f;
  }
  function update(id, key, { form: def, open } = {}) {
    const f = owned(id, key);
    if (def !== undefined) {
      const form = FormsCore.clean(def);
      if (!form || !form.questions.length) throw Object.assign(new Error('A form needs at least one question'), { status: 400 });
      f.form = form;
    }
    if (typeof open === 'boolean') f.open = open;
    saveMeta(id, f);
    return { open: f.open, count: f.count };
  }
  function remove(id, key) {
    owned(id, key);
    forms.delete(id);
    if (dir) for (const ext of ['.json', '.ndjson']) try { fs.unlinkSync(file(id, ext)); } catch { /* gone */ }
    return { deleted: true };
  }
  /** The public part: what people answering see. */
  function publicForm(id) {
    const f = ID_RE.test(String(id)) ? load(id) : null;
    if (!f) throw Object.assign(new Error('This form doesn’t exist, or was deleted'), { status: 404 });
    return { form: f.form, open: f.open };
  }
  function respond(id, answers) {
    const f = ID_RE.test(String(id)) ? load(id) : null;
    if (!f) throw Object.assign(new Error('This form doesn’t exist, or was deleted'), { status: 404 });
    if (!f.open) throw Object.assign(new Error('This form is no longer accepting responses'), { status: 403 });
    if (f.count >= limits.responses || f.bytes >= limits.storedBytes) throw Object.assign(new Error('This form has as many responses as it can take'), { status: 507 });
    const r = FormsCore.check(f.form, answers);
    if (!r.ok) throw Object.assign(new Error('Some answers need fixing'), { status: 422, errors: r.errors });
    const entry = { n: f.count + 1, at: new Date().toISOString(), values: r.values };
    const line = JSON.stringify(entry) + '\n';
    if (dir) fs.appendFileSync(file(id, '.ndjson'), line, { mode: 0o600 });
    else f.responses.push(entry);
    f.count++;
    f.bytes += Buffer.byteLength(line);
    return { n: entry.n };
  }
  /** Responses after number `after` (at most `limit`), for the owner. */
  function responses(id, key, after = 0, limit = 1000) {
    const f = owned(id, key);
    after = Math.max(0, Math.floor(+after) || 0);
    limit = Math.min(5000, Math.max(1, Math.floor(+limit) || 1000));
    let list;
    if (dir) {
      let text = '';
      try { text = fs.readFileSync(file(id, '.ndjson'), 'utf8'); } catch { /* none yet */ }
      list = text.split('\n').filter(Boolean).slice(after, after + limit).map(l => { try { return JSON.parse(l); } catch { return null; } }).filter(Boolean);
    } else list = f.responses.slice(after, after + limit);
    return { form: f.form, open: f.open, count: f.count, responses: list };
  }

  // ---------- HTTP ----------
  // Answers per address per minute, and forms created per address per hour.
  const hits = new Map();
  const timer = setInterval(() => { const cut = Date.now() - 3600e3; for (const [k, t] of hits) { const keep = t.filter(x => x > cut); if (keep.length) hits.set(k, keep); else hits.delete(k); } }, 60000);
  timer.unref();
  function allowed(kind, who, max, ms) {
    const k = kind + ' ' + who, cut = Date.now() - ms;
    const t = (hits.get(k) || []).filter(x => x > cut);
    if (t.length >= max) return false;
    t.push(Date.now()); hits.set(k, t.slice(-max));
    return true;
  }
  function readBody(req, max) {
    return new Promise((resolve, reject) => {
      const parts = []; let n = 0;
      req.on('data', d => { n += d.length; if (n > max) { reject(Object.assign(new Error('That is too much to send'), { status: 413 })); req.destroy(); } else parts.push(d); });
      req.on('end', () => { try { resolve(JSON.parse(Buffer.concat(parts).toString('utf8') || '{}')); } catch { reject(Object.assign(new Error('Bad request'), { status: 400 })); } });
      req.on('error', reject);
    });
  }
  const json = (res, code, body) => { res.writeHead(code, { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' }); res.end(JSON.stringify(body)); };
  const keyOf = req => (String(req.headers.authorization || '').match(/^Bearer ([A-Za-z0-9_-]{1,100})$/) || [])[1] || '';

  /** Handles /api/forms… and /f/<id>; returns false for anything else. */
  function handle(req, res, url) {
    if (url.pathname.startsWith('/f/')) {
      const id = url.pathname.slice(3);
      if (!ID_RE.test(id)) { res.writeHead(404); res.end('Not found'); return true; }
      res.writeHead(302, { Location: '/src/renderer/form.html?id=' + id, 'Cache-Control': 'no-store' });
      res.end();
      return true;
    }
    if (!url.pathname.startsWith('/api/forms')) return false;
    // The owner's calls carry their key, never cookies, so any page (the desktop app's, a phone's)
    // may make them.
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Headers', 'Authorization, Content-Type');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE');
    res.setHeader('Access-Control-Max-Age', '600');
    if (req.method === 'OPTIONS') { res.writeHead(204); res.end(); return true; }
    const who = req.socket.remoteAddress || '?';
    const parts = url.pathname.split('/').filter(Boolean); // api, forms, id?, responses?
    const id = parts[2], sub = parts[3];
    const fail = err => { if (!res.headersSent) json(res, err.status || 500, { error: err.status ? err.message : 'Something went wrong', ...(err.errors ? { errors: err.errors } : {}) }); else res.end(); };
    (async () => {
      if (parts.length === 2 && req.method === 'GET') return json(res, 200, { available: true });
      if (parts.length === 2 && req.method === 'POST') {
        if (!allowed('create', who, limits.createPerHour, 3600e3)) return json(res, 429, { error: 'Too many new forms from here in the last hour' });
        const body = await readBody(req, limits.formBytes);
        return json(res, 201, create(body.form));
      }
      if (parts.length === 3 && req.method === 'GET') return json(res, 200, publicForm(id));
      if (parts.length === 3 && req.method === 'PUT') return json(res, 200, update(id, keyOf(req), await readBody(req, limits.formBytes)));
      if (parts.length === 3 && req.method === 'DELETE') return json(res, 200, remove(id, keyOf(req)));
      if (parts.length === 4 && sub === 'responses' && req.method === 'POST') {
        if (!allowed('answer', who, limits.perMinute, 60000)) return json(res, 429, { error: 'Too many responses from here in the last minute. Try again shortly.' });
        const body = await readBody(req, limits.answerBytes);
        return json(res, 201, respond(id, body.answers));
      }
      if (parts.length === 4 && sub === 'responses' && req.method === 'GET') return json(res, 200, responses(id, keyOf(req), url.searchParams.get('after'), url.searchParams.get('limit')));
      json(res, 404, { error: 'Not found' });
    })().catch(fail);
    return true;
  }

  return { create, update, remove, publicForm, respond, responses, handle, close: () => clearInterval(timer) };
}

// The files the answer page needs, and nothing else of the app.
const PAGE_FILES = ['src/renderer/form.html', 'src/renderer/css/form.css', 'src/renderer/js/form.js', 'src/renderer/js/forms-core.js', 'src/renderer/js/forms-view.js'];
const TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8' };
const ROOT = path.resolve(__dirname, '../..');

/** A server for the local network: the answer page and the forms API (the desktop app's "This computer"). */
function createFormsServer(opts = {}) {
  const store = opts.store || createFormsStore(opts);
  const server = http.createServer((req, res) => {
    res.setHeader('X-Content-Type-Options', 'nosniff');
    res.setHeader('X-Frame-Options', 'DENY');
    res.setHeader('Referrer-Policy', 'no-referrer');
    let url;
    try { url = new URL(req.url, 'http://x'); } catch { res.writeHead(400); return res.end(); }
    try { if (store.handle(req, res, url)) return; } catch { if (!res.headersSent) { res.writeHead(500); res.end(); } return; }
    const rel = url.pathname.replace(/^\/+/, '');
    if ((req.method === 'GET' || req.method === 'HEAD') && PAGE_FILES.includes(rel)) {
      fs.readFile(path.join(ROOT, rel), (err, data) => {
        if (err) { res.writeHead(404); return res.end('Not found'); }
        res.writeHead(200, { 'Content-Type': TYPES[path.extname(rel)], 'Content-Length': data.length, 'Cache-Control': 'no-cache' });
        res.end(req.method === 'HEAD' ? undefined : data);
      });
      return;
    }
    res.writeHead(404); res.end('Not found');
  });
  server.headersTimeout = 20000;
  server.requestTimeout = 60000;
  server.keepAliveTimeout = 10000;
  server.store = store;
  server.on('close', () => store.close());
  return server;
}

module.exports = { createFormsStore, createFormsServer, LIMITS };
