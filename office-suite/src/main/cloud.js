'use strict';
// Cloud accounts for the desktop app: the OAuth redirect (a one-shot loopback server, as
// RFC 8252 recommends for installed apps), HTTP requests to the storage services (from here,
// so the pages' CSP and CORS don't apply), and the tokens, encrypted with the OS keychain.
const { app, ipcMain, safeStorage, shell } = require('electron');
const http = require('http');
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const settings = require('./settings');

// Hosts the storage services use; WebDAV accounts and endpoint overrides add their own.
const BASE_HOSTS = ['accounts.google.com', 'oauth2.googleapis.com', 'www.googleapis.com', 'login.microsoftonline.com', 'graph.microsoft.com',
  'www.dropbox.com', 'api.dropboxapi.com', 'content.dropboxapi.com'];
const MAX_BYTES = 300e6;

function cleanConfig(c) {
  const str = (v, n = 300) => (typeof v === 'string' ? v.trim().slice(0, n) : '');
  const url = v => { try { const u = new URL(str(v, 2000)); return /^https?:$/.test(u.protocol) ? u.toString() : ''; } catch { return ''; } };
  const out = {
    google: { clientId: str(c?.google?.clientId), clientSecret: str(c?.google?.clientSecret) },
    microsoft: { clientId: str(c?.microsoft?.clientId), tenant: /^[\w.-]{1,100}$/.test(str(c?.microsoft?.tenant)) ? str(c.microsoft.tenant) : 'common' },
    dropbox: { clientId: str(c?.dropbox?.clientId) },
    // Endpoint overrides (self-hosted or test servers): { provider: { key: url } }.
    endpoints: {}
  };
  if (c && c.endpoints && typeof c.endpoints === 'object') {
    for (const [p, e] of Object.entries(c.endpoints)) {
      if (!['google', 'microsoft', 'dropbox'].includes(p) || !e || typeof e !== 'object') continue;
      out.endpoints[p] = {};
      for (const [k, v] of Object.entries(e)) if (/^\w{1,40}$/.test(k) && url(v)) out.endpoints[p][k] = url(v);
    }
  }
  return out;
}
function allowedHosts() {
  const c = cleanConfig(settings.get('cloud') || {});
  const hosts = new Set(BASE_HOSTS);
  const add = u => { try { hosts.add(new URL(u).host); } catch { /* skip */ } };
  Object.values(c.endpoints).forEach(e => Object.values(e).forEach(add));
  for (const a of readTokens().accounts || []) if (a && a.provider === 'webdav' && typeof a.url === 'string') add(a.url);
  return hosts;
}

// ---------- OAuth redirect ----------
const waits = new Map();
ipcMain.handle('cloud:oauth-begin', async () => {
  const server = http.createServer();
  await new Promise((res, rej) => { server.once('error', rej); server.listen(0, '127.0.0.1', res); });
  const id = crypto.randomBytes(12).toString('hex');
  const redirectUri = `http://127.0.0.1:${server.address().port}/callback`;
  let settle;
  const result = new Promise(res => { settle = res; });
  const timer = setTimeout(() => { settle({ error: 'timeout' }); server.close(); }, 5 * 60 * 1000);
  // One sign-in only: after the first callback nothing more is answered, even on a kept-alive connection.
  let done = false;
  server.on('request', (req, resp) => {
    if (done) { req.socket.destroy(); return; }
    const u = new URL(req.url, redirectUri);
    if (u.pathname !== '/callback') { resp.writeHead(404); resp.end(); return; }
    done = true;
    resp.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8', 'Content-Security-Policy': "default-src 'none'; style-src 'unsafe-inline'", Connection: 'close' });
    resp.end('<!doctype html><title>Isadek Office</title><body style="font:16px system-ui;margin:3em;text-align:center"><h2>Signed in</h2><p>You can close this tab and return to Isadek Office.</p>');
    clearTimeout(timer);
    settle({ code: u.searchParams.get('code'), state: u.searchParams.get('state'), error: u.searchParams.get('error') });
    resp.on('finish', () => { server.close(); if (server.closeAllConnections) server.closeAllConnections(); });
  });
  waits.set(id, { result, server, timer });
  return { id, redirectUri };
});
ipcMain.handle('cloud:oauth-wait', async (_e, { id, authUrl }) => {
  const w = waits.get(id);
  if (!w) return { error: 'unknown sign-in' };
  waits.delete(id);
  let u;
  try { u = new URL(authUrl); } catch { return { error: 'bad sign-in address' }; }
  if (!(u.protocol === 'https:' || (u.protocol === 'http:' && allowedHosts().has(u.host)))) return { error: 'bad sign-in address' };
  await shell.openExternal(u.toString());
  return w.result;
});

// ---------- HTTP for the services ----------
ipcMain.handle('cloud:fetch', async (_e, req) => {
  let u;
  try { u = new URL(String(req && req.url)); } catch { return { status: 0, error: 'bad address' }; }
  if (!/^https?:$/.test(u.protocol) || !allowedHosts().has(u.host)) return { status: 0, error: `Isadek Office doesn't connect to ${u.host}` };
  const method = /^(GET|POST|PUT|PATCH|DELETE|PROPFIND|MKCOL)$/.test(req.method || 'GET') ? req.method || 'GET' : 'GET';
  const headers = {};
  for (const [k, v] of Object.entries(req.headers || {})) if (/^[\w-]{1,60}$/.test(k) && typeof v === 'string' && v.length < 10000) headers[k] = v;
  const body = req.body instanceof Uint8Array ? req.body : typeof req.body === 'string' ? req.body : undefined;
  try {
    const r = await fetch(u, { method, headers, body, redirect: 'follow' });
    const buf = new Uint8Array(await r.arrayBuffer());
    if (buf.length > MAX_BYTES) return { status: 0, error: 'response too large' };
    return { status: r.status, headers: Object.fromEntries(r.headers.entries()), body: buf };
  } catch (err) {
    return { status: 0, error: String(err.message || err) };
  }
});

// ---------- Settings and tokens ----------
// The publisher's own app registrations, built into this copy (see scripts/cloud-apps.js).
function builtIn() {
  try { return cleanConfig(JSON.parse(fs.readFileSync(path.join(__dirname, '../renderer/cloud-apps.json'), 'utf8'))); } catch { return cleanConfig({}); }
}
ipcMain.handle('cloud:config-get', () => cleanConfig(settings.get('cloud') || {}));
ipcMain.handle('cloud:builtin-get', () => builtIn());
ipcMain.handle('cloud:config-set', (_e, c) => { settings.set('cloud', cleanConfig(c)); return cleanConfig(c); });
const tokenFile = () => path.join(app.getPath('userData'), 'cloud-accounts.bin');
function readTokens() {
  try {
    const raw = fs.readFileSync(tokenFile());
    const text = safeStorage.isEncryptionAvailable() && raw[0] !== 0x7b ? safeStorage.decryptString(raw) : raw.toString('utf8');
    const t = JSON.parse(text);
    return t && typeof t === 'object' ? t : {};
  } catch { return {}; }
}
ipcMain.handle('cloud:accounts-get', () => readTokens());
ipcMain.handle('cloud:accounts-set', (_e, t) => {
  const text = JSON.stringify(t && typeof t === 'object' ? t : {});
  if (text.length > 1e6) return false;
  fs.mkdirSync(path.dirname(tokenFile()), { recursive: true });
  // Encrypted with the OS keychain where there is one (macOS Keychain, Windows DPAPI, libsecret).
  fs.writeFileSync(tokenFile(), safeStorage.isEncryptionAvailable() ? safeStorage.encryptString(text) : Buffer.from(text, 'utf8'), { mode: 0o600 });
  return true;
});

module.exports = { cleanConfig, isCloudPath: p => typeof p === 'string' && /^cloud:\/\/[\w-]{1,40}\//.test(p) && p.length < 4000 };
