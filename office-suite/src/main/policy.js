'use strict';
// Policies an organisation sets for Isadek Office on its computers (see ENTERPRISE.md).
//
// Read once at start, from (later ones win):
//   a JSON file — Windows %ProgramData%\Isadek Office\policies.json, macOS
//     /Library/Application Support/Isadek Office/policies.json, Linux /etc/isadek-office/policies.json;
//   Windows Group Policy — HKLM\SOFTWARE\Policies\Isadek\Office (from the ADMX template in policies/);
//   macOS configuration profiles — /Library/Managed Preferences/io.github.isaadac.isadekoffice.plist.
// These are places only an administrator can write. Unknown names and values of the wrong kind are
// ignored. What a policy turns off is refused here, in the main process (so a page can't get round
// it), and the pages hide it too.
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const { app, ipcMain } = require('electron');

const SCHEMA = {
  DisableAssistant: 'bool',       // the Assistant (Claude)
  DisableCloudStorage: 'bool',    // Google Drive, OneDrive, Dropbox, WebDAV
  DisableLiveEditing: 'bool',     // live sessions (hosting and joining)
  LiveServers: 'list',            // only these live-editing servers (ws:// or wss:// prefixes); "local" allows This computer
  DisableForms: 'bool',           // publishing forms, polls, Q&A and the phone remote
  DisableAddins: 'bool',          // installing and running add-ins
  DisableWebData: 'bool',         // From Web, Web Clip and other downloads a page asks for
  DisableAutoUpdate: 'bool',      // automatic updates (deploy new versions yourself)
  DisableCrashReports: 'bool',    // sending crash reports
  DisableSpeechDownload: 'bool',  // downloading the on-device speech model
  DefaultLanguage: 'lang'         // the interface language until a person picks one
};
const LANGS = ['en', 'es', 'fr', 'de', 'pt', 'it', 'ar'];

function clean(raw) {
  const out = {};
  if (!raw || typeof raw !== 'object') return out;
  for (const [k, kind] of Object.entries(SCHEMA)) {
    const v = raw[k];
    if (v === undefined || v === null) continue;
    if (kind === 'bool') {
      if (typeof v === 'boolean') out[k] = v;
      else if (v === 1 || v === 0 || v === '1' || v === '0') out[k] = String(v) === '1';
    } else if (kind === 'list') {
      const list = (Array.isArray(v) ? v : String(v).split(/[\n;,]+/)).map(x => String(x).trim()).filter(x => x === 'local' || /^wss?:\/\/[^\s]+$/i.test(x)).slice(0, 50);
      out[k] = list;
    } else if (kind === 'lang' && LANGS.includes(String(v))) out[k] = String(v);
  }
  return out;
}

function jsonFile() {
  if (!app.isPackaged && process.env.ISADEK_POLICY_FILE) return process.env.ISADEK_POLICY_FILE; // tests only
  if (process.platform === 'win32') return path.join(process.env.ProgramData || 'C:\\ProgramData', 'Isadek Office', 'policies.json');
  if (process.platform === 'darwin') return '/Library/Application Support/Isadek Office/policies.json';
  return '/etc/isadek-office/policies.json';
}
function readJson(file) {
  try {
    const st = fs.statSync(file);
    if (!st.isFile() || st.size > 100000) return {};
    return JSON.parse(fs.readFileSync(file, 'utf8'));
  } catch { return {}; }
}
/** Windows Group Policy: values under HKLM (then HKCU, which HKLM overrides). */
function readRegistry() {
  if (process.platform !== 'win32') return {};
  const out = {};
  for (const hive of ['HKCU', 'HKLM']) {
    let text = '';
    try { text = execFileSync('reg', ['query', `${hive}\\SOFTWARE\\Policies\\Isadek\\Office`], { encoding: 'utf8', timeout: 5000, windowsHide: true }); } catch { continue; }
    for (const line of text.split(/\r?\n/)) {
      const m = /^\s+(\w+)\s+(REG_DWORD|REG_SZ|REG_MULTI_SZ)\s+(.*)$/.exec(line);
      if (!m) continue;
      out[m[1]] = m[2] === 'REG_DWORD' ? String(parseInt(m[3], 16)) : m[2] === 'REG_MULTI_SZ' ? m[3].split('\\0') : m[3];
    }
  }
  return out;
}
/** macOS configuration profiles (managed preferences). */
function readManagedPrefs() {
  if (process.platform !== 'darwin') return {};
  const f = '/Library/Managed Preferences/io.github.isaadac.isadekoffice.plist';
  if (!fs.existsSync(f)) return {};
  try { return JSON.parse(execFileSync('plutil', ['-convert', 'json', '-o', '-', f], { encoding: 'utf8', timeout: 5000 })); } catch { return {}; }
}

const policy = Object.freeze({ ...clean(readJson(jsonFile())), ...clean(readRegistry()), ...clean(readManagedPrefs()) });
const OFF = 'This was turned off by your organisation';
/** Main-process functions each policy turns off. */
const BLOCKS = {
  DisableAssistant: ['ai:ask', 'ai:set-key'],
  DisableCloudStorage: ['cloud:oauth-begin', 'cloud:oauth-wait', 'cloud:fetch'],
  DisableLiveEditing: ['collab:host', 'collab:join'],
  DisableForms: ['forms:host', 'forms:call'],
  DisableWebData: ['net:fetch-data']
};
const blocked = new Map();
for (const [k, chans] of Object.entries(BLOCKS)) if (policy[k]) chans.forEach(c => blocked.set(c, k));
if (policy.LiveServers && !policy.LiveServers.includes('local')) blocked.set('collab:host', 'LiveServers');

// Wraps handlers registered after this (main.js loads this right after the IPC guard).
const handle = ipcMain.handle.bind(ipcMain);
ipcMain.handle = (channel, fn) => handle(channel, (e, ...args) => {
  // (Refused the way these functions report problems: { error } for the page to show.)
  if (blocked.has(channel)) return { error: OFF };
  return fn(e, ...args);
});
ipcMain.on('policy:get', e => { e.returnValue = policy; });

/** Private and loopback addresses: where "This computer" live sessions run. */
function localHost(host) {
  const h = host.replace(/^\[|\]$/g, '').toLowerCase();
  if (h === 'localhost' || h === '::1' || /^127\./.test(h) || /^10\./.test(h) || /^192\.168\./.test(h) || /^169\.254\./.test(h)) return true;
  const m = /^172\.(\d+)\./.exec(h);
  return (m && +m[1] >= 16 && +m[1] <= 31) || /^f[cd][0-9a-f]{2}:/.test(h) || /^fe80:/.test(h);
}
/**
 * Whether a page may make this network request under the policies. Live sessions (WebSockets),
 * forms and the speech model download go straight from the pages, so they are checked here, on
 * the way out, as well as in the pages.
 */
function allowed(url) {
  let u;
  try { u = new URL(url); } catch { return true; }
  if (u.protocol === 'ws:' || u.protocol === 'wss:') {
    if (policy.DisableLiveEditing) return false;
    const list = policy.LiveServers;
    if (!list) return true;
    if (list.includes('local') && localHost(u.hostname)) return true;
    return list.some(p => p !== 'local' && url.toLowerCase().startsWith(p.toLowerCase()));
  }
  if (u.protocol === 'http:' || u.protocol === 'https:') {
    if (policy.DisableForms && /^\/api\/(forms|remote)(\/|$)/.test(u.pathname)) return false;
    if (policy.DisableSpeechDownload && /(^|\.)huggingface\.co$/i.test(u.hostname)) return false;
  }
  return true;
}
/**
 * Applies to a session's requests (main.js, when the app is ready): no scripts from the network
 * (the app's are all its own files), and the organisation's network policies.
 */
function applyNetwork(ses) {
  ses.webRequest.onBeforeRequest({ urls: ['ws://*/*', 'wss://*/*', 'http://*/*', 'https://*/*'] }, (d, cb) => cb({ cancel: d.resourceType === 'script' || !allowed(d.url) }));
}

module.exports = { policy, clean, OFF, blocked: c => blocked.has(c), allowed, applyNetwork };
