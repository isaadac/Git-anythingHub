'use strict';
// Crash reports. Off until the person turns them on (Help → Send Crash Reports), and never with
// any document content.
//   - Native crashes (any process) are caught by Electron's crash reporter as minidumps.
//   - Script errors in the main process and the windows, and windows that stop working, are
//     written to userData/crash-reports/errors.jsonl (the latest 200), with home folders and user
//     names taken out of file paths.
// Everything stays on the computer unless the copy was built with a report address
// (ISADEK_CRASH_URL, see scripts/release-config.js) and the person turned reports on; then
// minidumps go to that address (Crashpad's multipart upload) and errors are posted there as JSON.
const { app, crashReporter, ipcMain, net, shell } = require('electron');
const fs = require('fs');
const os = require('os');
const path = require('path');
const settings = require('./settings');

const KEEP = 200, PER_SESSION = 50;
function releaseConfig() {
  try { return JSON.parse(fs.readFileSync(path.join(__dirname, 'release.json'), 'utf8')); } catch { return {}; }
}
const reportUrl = () => {
  const u = String(process.env.ISADEK_CRASH_URL || releaseConfig().crashUrl || '').trim();
  return /^https?:\/\//i.test(u) ? u : '';
};
const enabled = () => !!settings.get('crashReports');
const dir = () => path.join(app.getPath('userData'), 'crash-reports');
const logFile = () => path.join(dir(), 'errors.jsonl');

// File paths say who the person is (C:\Users\Ana\…): keep only what follows the home folder.
const home = os.homedir();
const user = (() => { try { return os.userInfo().username; } catch { return ''; } })();
const esc = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
function scrub(s, max = 4000) {
  let t = String(s == null ? '' : s).slice(0, max * 2);
  if (home) t = t.split(home).join('~').split(home.replace(/\\/g, '/')).join('~');
  if (user && user.length > 2) t = t.replace(new RegExp(`([\\\\/])${esc(user)}([\\\\/])`, 'gi'), '$1~$2');
  return t.slice(0, max);
}

let sent = 0;
function record(entry) {
  if (sent >= PER_SESSION) return null;
  sent++;
  const e = {
    time: new Date().toISOString(), version: app.getVersion(), platform: process.platform, arch: process.arch, electron: process.versions.electron,
    kind: String(entry.kind || 'error').slice(0, 40), where: String(entry.where || '').slice(0, 40),
    message: scrub(entry.message, 1000), stack: scrub(entry.stack, 6000), source: scrub(entry.source, 500),
    ...(entry.reason ? { reason: String(entry.reason).slice(0, 40), exitCode: entry.exitCode | 0 } : {})
  };
  try {
    fs.mkdirSync(dir(), { recursive: true });
    let lines = [];
    try { lines = fs.readFileSync(logFile(), 'utf8').split('\n').filter(Boolean); } catch { /* first one */ }
    lines.push(JSON.stringify(e));
    fs.writeFileSync(logFile(), lines.slice(-KEEP).join('\n') + '\n');
  } catch { /* nowhere to write: drop it */ }
  upload(e);
  return e;
}
function upload(e) {
  const url = reportUrl();
  if (!url || !enabled() || !app.isReady()) return;
  try {
    const req = net.request({ method: 'POST', url, useSessionCookies: false });
    req.setHeader('Content-Type', 'application/json');
    req.on('error', () => {});
    req.on('response', r => { r.on('data', () => {}); r.on('error', () => {}); });
    req.end(JSON.stringify({ type: 'error', product: 'Isadek Office', ...e }));
  } catch { /* offline: it is still in the local log */ }
}

function start() {
  const url = reportUrl();
  crashReporter.start({
    productName: 'Isadek Office', companyName: 'Isadek', submitURL: url || 'https://127.0.0.1/none',
    uploadToServer: !!url && enabled(), compress: true, ignoreSystemCrashHandler: false,
    globalExtra: { _version: app.getVersion() }
  });
  process.on('uncaughtException', err => { record({ kind: 'exception', where: 'main', message: err && err.message, stack: err && err.stack }); console.error(err); });
  process.on('unhandledRejection', err => { record({ kind: 'rejection', where: 'main', message: err && err.message || String(err), stack: err && err.stack }); });
  app.on('render-process-gone', (_e, wc, d) => {
    if (d.reason === 'clean-exit') return;
    let where = '';
    try { where = (new URL(wc.getURL()).pathname.match(/(\w+)\.html$/) || [])[1] || ''; } catch { /* no page */ }
    record({ kind: 'window-gone', where, reason: d.reason, exitCode: d.exitCode, message: `The ${where || 'window'} process ended (${d.reason})` });
  });
  app.on('child-process-gone', (_e, d) => { if (d.reason !== 'clean-exit' && d.reason !== 'killed') record({ kind: 'process-gone', where: d.type, reason: d.reason, exitCode: d.exitCode, message: `${d.type} process ended (${d.reason})` }); });
}

/** The windows' script errors (sent from platform.js). */
ipcMain.handle('crash:report', (e, r) => {
  if (!r || typeof r !== 'object') return false;
  let where = '';
  try { where = (new URL(e.sender.getURL()).pathname.match(/(\w+)\.html$/) || [])[1] || ''; } catch { /* no page */ }
  return !!record({ kind: r.kind === 'rejection' ? 'rejection' : 'error', where, message: r.message, stack: r.stack, source: r.source });
});
ipcMain.handle('crash:status', () => ({ enabled: enabled(), url: !!reportUrl() }));

function setEnabled(on) {
  settings.set('crashReports', !!on);
  try { crashReporter.setUploadToServer(!!on && !!reportUrl()); } catch { /* not started */ }
}
function openFolder() {
  fs.mkdirSync(dir(), { recursive: true });
  return shell.openPath(dir());
}
function list() {
  try { return fs.readFileSync(logFile(), 'utf8').split('\n').filter(Boolean).map(l => JSON.parse(l)); } catch { return []; }
}

module.exports = { start, record, setEnabled, enabled, reportUrl, openFolder, list, scrub };
