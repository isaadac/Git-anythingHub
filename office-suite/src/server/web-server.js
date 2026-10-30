'use strict';
// Isadek Office web server: the web app, the Assistant (Claude) and live editing on one port.
// The Anthropic API key stays here on the server; pages send their questions to /api/assistant
// and get the reply streamed back, so the key never reaches a browser.
//
//   ANTHROPIC_API_KEY=sk-ant-… ASSISTANT_CODES=code1,code2 node src/server/web-server.js [--port 8080] [--host 0.0.0.0]
//
// Settings (environment):
//   ANTHROPIC_API_KEY      the key the Assistant uses (without it the Assistant is off)
//   ASSISTANT_CODES        access codes people type in the Assistant settings (comma separated);
//                          required unless ASSISTANT_OPEN=1, so strangers can't spend your credit
//   ASSISTANT_OPEN=1       let anyone who can reach the server use the Assistant
//   ASSISTANT_ORIGINS      other sites whose pages may call the Assistant (comma separated origins)
//   ASSISTANT_PER_MINUTE   questions per minute per person or address (default 10)
//   ASSISTANT_MODEL        the default model
//   ISADEK_GOOGLE_CLIENT_ID, ISADEK_MICROSOFT_CLIENT_ID, ISADEK_DROPBOX_APP_KEY, …
//                          your cloud app registrations, offered to every page (scripts/cloud-apps.js)
//   FORMS_DIR              keep Isadek Forms and their responses in this folder (without it they are
//                          kept in memory and lost when the server stops)
//   CRASH_DIR              keep crash reports the desktop app sends to /api/crash in this folder
//                          (build the app with ISADEK_CRASH_URL=https://this-server/api/crash)
const http = require('http');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { Anthropic, MODELS, validModel, cleanMessages, ask } = require('../main/assistant-core');
const { createCollabServer } = require('../main/collab-server');
const { fromEnv: cloudAppsFromEnv } = require('../../scripts/cloud-apps');
const { createCrashCollector } = require('./crash-collector');
const { createFormsStore } = require('../main/forms-server');

const ROOT = path.resolve(__dirname, '../..');
const RENDERER = path.join(ROOT, 'src/renderer');
const TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.mjs': 'text/javascript; charset=utf-8', '.zip': 'application/zip', '.css': 'text/css; charset=utf-8', '.json': 'application/json', '.png': 'image/png', '.svg': 'image/svg+xml', '.ico': 'image/x-icon', '.wasm': 'application/wasm', '.woff2': 'font/woff2', '.map': 'application/json', '.bcmap': 'application/octet-stream', '.pfb': 'application/octet-stream', '.ttf': 'font/ttf', '.onnx': 'application/octet-stream' };
const MAX_BODY = 4e6;
// Sent with every response: no framing by other sites (clickjacking), no MIME sniffing, no
// referrer to other sites, and the files only for this site's pages. (No Cross-Origin-Opener-
// Policy: cloud sign-in pop-ups report back to the page that opened them.)
const SECURITY_HEADERS = {
  'X-Content-Type-Options': 'nosniff',
  'X-Frame-Options': 'DENY',
  'Content-Security-Policy': "frame-ancestors 'none'",
  'Referrer-Policy': 'no-referrer',
  'Cross-Origin-Resource-Policy': 'same-origin',
  'Permissions-Policy': 'camera=(self), microphone=(self), geolocation=(), payment=(), usb=()'
};

/** Library files the pages load (../../node_modules/…): only those are served from node_modules. */
function libraryFiles() {
  const wanted = new Set();
  for (const f of fs.readdirSync(RENDERER, { recursive: true })) {
    if (!/\.(html|js)$/.test(f)) continue;
    const text = fs.readFileSync(path.join(RENDERER, f), 'utf8');
    for (const m of text.matchAll(/node_modules\/((?:@[\w.-]+\/)?[\w.-]+\/[\w./-]+\.(?:json|m?js|css|wasm|zip)\b)/g)) wanted.add('node_modules/' + m[1]);
    // Whole folders a library loads by name at run time ("// library folders: node_modules/x/dir/ …").
    for (const m of text.matchAll(/library folders:([^\n]*)/g)) for (const d of m[1].match(/node_modules\/[\w@./-]+\//g) || []) if (!d.includes('..')) wanted.add(d);
  }
  return wanted;
}
const isLibrary = (libs, rel) => libs.has(rel) || [...libs].some(d => d.endsWith('/') && rel.startsWith(d));

function createWebServer(opts = {}) {
  const env = opts.env || process.env;
  const list = v => String(v || '').split(',').map(s => s.trim()).filter(Boolean);
  const cfg = {
    key: opts.apiKey !== undefined ? opts.apiKey : env.ANTHROPIC_API_KEY || '',
    codes: opts.codes || list(env.ASSISTANT_CODES),
    open: opts.open !== undefined ? opts.open : env.ASSISTANT_OPEN === '1',
    origins: opts.origins || list(env.ASSISTANT_ORIGINS),
    perMinute: opts.perMinute || +env.ASSISTANT_PER_MINUTE || 10,
    model: validModel(opts.model || env.ASSISTANT_MODEL)
  };
  const client = opts.client || (cfg.key ? new Anthropic({ apiKey: cfg.key, ...(opts.baseURL ? { baseURL: opts.baseURL } : {}) }) : null);
  const libs = libraryFiles();
  const enabled = () => !!client && (cfg.open || cfg.codes.length > 0);
  const crashDir = opts.crashDir !== undefined ? opts.crashDir : env.CRASH_DIR;
  const crashes = crashDir ? createCrashCollector({ dir: path.resolve(crashDir) }) : null;
  // Forms people answer at /f/<id>, with their responses kept for the workbook that made them.
  const formsDir = opts.formsDir !== undefined ? opts.formsDir : env.FORMS_DIR;
  const forms = opts.forms === false ? null : createFormsStore({ dir: formsDir || null, limits: opts.formsLimits });

  // Questions per minute, per access code (or address when the Assistant is open).
  const hits = new Map();
  setInterval(() => { const cut = Date.now() - 60000; for (const [k, t] of hits) { const keep = t.filter(x => x > cut); if (keep.length) hits.set(k, keep); else hits.delete(k); } }, 60000).unref();
  function allowed(who) {
    const cut = Date.now() - 60000;
    const t = (hits.get(who) || []).filter(x => x > cut);
    if (t.length >= cfg.perMinute) return false;
    t.push(Date.now()); hits.set(who, t);
    return true;
  }
  // Wrong access codes, per address: guessing is limited to a few tries a minute.
  const misses = new Map();
  setInterval(() => { const cut = Date.now() - 600000; for (const [k, t] of misses) { const keep = t.filter(x => x > cut); if (keep.length) misses.set(k, keep); else misses.delete(k); } }, 60000).unref();
  const addr = req => req.socket.remoteAddress || '?';
  const lockedOut = req => (misses.get(addr(req)) || []).filter(x => x > Date.now() - 600000).length >= 20;
  const missed = req => { const t = misses.get(addr(req)) || []; t.push(Date.now()); misses.set(addr(req), t.slice(-50)); };
  const same = (a, b) => { const x = Buffer.from(a), y = Buffer.from(b); return x.length === y.length && crypto.timingSafeEqual(x, y); };
  const codeOf = req => (String(req.headers.authorization || '').match(/^Bearer (.{1,200})$/) || [])[1] || '';
  const validCode = code => cfg.codes.some(c => same(c, code));

  function cors(req, res) {
    const o = req.headers.origin;
    if (o && cfg.origins.includes(o)) {
      res.setHeader('Access-Control-Allow-Origin', o);
      res.setHeader('Vary', 'Origin');
      res.setHeader('Access-Control-Allow-Headers', 'Authorization, Content-Type');
      res.setHeader('Access-Control-Allow-Methods', 'GET, POST');
      res.setHeader('Access-Control-Max-Age', '600');
    }
  }
  const json = (res, code, body) => { res.writeHead(code, { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' }); res.end(JSON.stringify(body)); };

  function readBody(req) {
    return new Promise((resolve, reject) => {
      const parts = []; let n = 0;
      req.on('data', d => { n += d.length; if (n > MAX_BODY) { reject(new Error('too big')); req.destroy(); } else parts.push(d); });
      req.on('end', () => resolve(Buffer.concat(parts).toString('utf8')));
      req.on('error', reject);
    });
  }

  async function assistant(req, res) {
    if (!enabled()) return json(res, 503, { error: 'The Assistant is not set up on this server.' });
    const code = codeOf(req);
    if (!cfg.open && lockedOut(req)) return json(res, 429, { error: 'Too many wrong access codes. Try again in a few minutes.' });
    if (!cfg.open && !validCode(code)) { missed(req); return json(res, 401, { error: 'Enter a valid access code in the Assistant settings.', needsCode: true }); }
    if (!allowed(code || req.socket.remoteAddress)) return json(res, 429, { error: 'Too many questions in the last minute. Try again shortly.' });
    let body;
    try { body = JSON.parse(await readBody(req)); } catch { return json(res, 400, { error: 'Bad request' }); }
    const messages = cleanMessages(body && body.messages);
    if (!messages) return json(res, 400, { error: 'Nothing to ask' });
    // The reply streams as lines of JSON: {"t":"delta","text"} … then {"t":"done",…}.
    res.writeHead(200, { 'Content-Type': 'application/x-ndjson; charset=utf-8', 'Cache-Control': 'no-store', 'X-Accel-Buffering': 'no' });
    const abort = new AbortController();
    res.on('close', () => { if (!res.writableEnded) abort.abort(); });
    const out = await ask(client, { model: validModel(body.model || cfg.model), system: body.system, messages, effort: body.effort },
      text => res.write(JSON.stringify({ t: 'delta', text }) + '\n'), abort.signal);
    if (!res.destroyed) res.end(JSON.stringify({ t: 'done', ...out }) + '\n');
  }

  // Cloud app registrations from the environment replace the file built into the pages.
  const cloudApps = cloudAppsFromEnv(env);
  function serveFile(req, res, url) {
    let rel;
    // A malformed address (bad %-escape) or a NUL in it is simply not found.
    try { rel = decodeURIComponent(url.pathname).replace(/^\/+/, ''); } catch { res.writeHead(400); return res.end('Bad request'); }
    if (rel.includes('\0')) { res.writeHead(400); return res.end('Bad request'); }
    if (rel === 'src/renderer/cloud-apps.json' && Object.keys(cloudApps).length) return json(res, 200, cloudApps);
    if (!rel) { res.writeHead(302, { Location: '/src/renderer/index.html' }); return res.end(); }
    rel = path.posix.normalize(rel);
    const ok = (rel.startsWith('src/renderer/') && !rel.includes('..')) || (!rel.includes('..') && isLibrary(libs, rel));
    const file = path.join(ROOT, rel);
    if (!ok || !file.startsWith(ROOT + path.sep)) { res.writeHead(404); return res.end('Not found'); }
    fs.stat(file, (err, st) => {
      if (err || !st.isFile()) { res.writeHead(404); return res.end('Not found'); }
      res.writeHead(200, { 'Content-Type': TYPES[path.extname(file)] || 'application/octet-stream', 'Content-Length': st.size, 'X-Content-Type-Options': 'nosniff', 'Cache-Control': 'no-cache' });
      if (req.method === 'HEAD') return res.end();
      const stream = fs.createReadStream(file);
      stream.on('error', () => res.destroy());
      stream.pipe(res);
    });
  }

  const server = http.createServer((req, res) => {
    for (const [k, v] of Object.entries(SECURITY_HEADERS)) res.setHeader(k, v);
    try { route(req, res); } catch (err) {
      // Never let one request take the server down.
      if (!res.headersSent) { res.writeHead(500); res.end('Server error'); } else res.destroy();
      console.error('request failed:', err && err.message);
    }
  });
  function route(req, res) {
    let url;
    try { url = new URL(req.url, 'http://x'); } catch { res.writeHead(400); return res.end(); }
    if (forms && forms.handle(req, res, url)) return;
    if (url.pathname.startsWith('/api/')) {
      cors(req, res);
      if (req.method === 'OPTIONS') { res.writeHead(204); return res.end(); }
      if (url.pathname === '/api/assistant/status' && req.method === 'GET') {
        const code = codeOf(req);
        let codeOk = cfg.open;
        if (!cfg.open && code) {
          if (lockedOut(req)) return json(res, 429, { error: 'Too many wrong access codes. Try again in a few minutes.' });
          codeOk = validCode(code);
          if (!codeOk) missed(req);
        }
        return json(res, 200, { available: enabled(), needsCode: !cfg.open, codeOk, models: MODELS, model: cfg.model });
      }
      if (url.pathname === '/api/crash' && crashes) return crashes.handle(req, res).catch(() => { if (!res.headersSent) json(res, 500, { error: 'Could not keep the report' }); });
      if (url.pathname === '/api/assistant' && req.method === 'POST') return assistant(req, res).catch(err => { if (!res.headersSent) json(res, 500, { error: String(err.message || err) }); else res.end(); });
      return json(res, 404, { error: 'Not found' });
    }
    if (req.method !== 'GET' && req.method !== 'HEAD') { res.writeHead(405); return res.end(); }
    serveFile(req, res, url);
  }
  // Slow or idle clients are let go instead of holding sockets open.
  server.headersTimeout = 20000;
  server.requestTimeout = 120000;
  server.keepAliveTimeout = 10000;
  if (opts.live !== false) createCollabServer({ server });
  if (forms) server.on('close', () => forms.close());
  return server;
}

module.exports = { createWebServer };

if (require.main === module) {
  const arg = (n, d) => { const i = process.argv.indexOf(n); return i > 0 ? process.argv[i + 1] : d; };
  const port = +arg('--port', process.env.PORT || 8080), host = arg('--host', '0.0.0.0');
  const srv = createWebServer();
  srv.listen(port, host, () => {
    console.log(`Isadek Office on http://${host}:${port}/  (live editing: ws://${host}:${port}/)`);
    if (!process.env.ANTHROPIC_API_KEY) console.log('Assistant: off (set ANTHROPIC_API_KEY to turn it on)');
    else if (process.env.ASSISTANT_OPEN !== '1' && !process.env.ASSISTANT_CODES) console.log('Assistant: off until you set ASSISTANT_CODES (or ASSISTANT_OPEN=1)');
    if (!process.env.FORMS_DIR) console.log('Forms: responses are kept in memory (set FORMS_DIR to keep them on disk)');
  });
}
