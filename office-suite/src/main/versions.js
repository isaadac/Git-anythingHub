'use strict';
// Version history for the desktop app: each save keeps a copy of the document in the app's data
// folder (at most one copy per 10 minutes, the latest 30 per document), so an earlier version can
// be opened or restored. A document is identified by where it is saved (a file or cloud path).
const { app, ipcMain } = require('electron');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const KEEP = 30, WINDOW_MS = 10 * 60 * 1000, MAX_BYTES = 150e6;
// All versions of all documents together: past this, the oldest versions go first.
const TOTAL_BYTES = 2e9;
const root = () => path.join(app.getPath('userData'), 'versions');
const validKey = k => typeof k === 'string' && k.length > 0 && k.length <= 4000;
const dirOf = key => path.join(root(), crypto.createHash('sha1').update(key).digest('hex'));
function readIndex(dir) {
  try { const x = JSON.parse(fs.readFileSync(path.join(dir, 'index.json'), 'utf8')); return Array.isArray(x.versions) ? x : { versions: [] }; } catch { return { versions: [] }; }
}
function writeIndex(dir, idx) {
  const tmp = path.join(dir, 'index.json.tmp');
  fs.writeFileSync(tmp, JSON.stringify(idx));
  fs.renameSync(tmp, path.join(dir, 'index.json'));
}

ipcMain.handle('versions:add', (_e, key, name, data, by) => {
  if (!validKey(key) || !(data instanceof Uint8Array) || data.length > MAX_BYTES) return false;
  const dir = dirOf(key);
  fs.mkdirSync(dir, { recursive: true });
  const idx = readIndex(dir);
  const hash = crypto.createHash('sha1').update(data).digest('hex');
  const last = idx.versions[idx.versions.length - 1];
  if (last && last.hash === hash) return true; // nothing changed since the last version
  const now = Date.now();
  // Saves close together (AutoSave) update one version instead of piling up.
  if (last && now - last.created < WINDOW_MS) { try { fs.unlinkSync(path.join(dir, last.id + '.bin')); } catch { /* gone */ } idx.versions.pop(); }
  const id = String(now);
  fs.writeFileSync(path.join(dir, id + '.bin'), data);
  idx.versions.push({ id, name: String(name || '').slice(0, 255), by: (Array.isArray(by) ? by : []).filter(x => typeof x === 'string').map(x => x.replace(/[^\w .@'-]/g, '').slice(0, 60)).filter(Boolean).slice(0, 10), size: data.length, hash, created: last && now - last.created < WINDOW_MS ? last.created : now, saved: now });
  while (idx.versions.length > KEEP) { const old = idx.versions.shift(); try { fs.unlinkSync(path.join(dir, old.id + '.bin')); } catch { /* gone */ } }
  idx.key = key;
  writeIndex(dir, idx);
  pruneSoon();
  return true;
});

/** Keeps the whole version store under TOTAL_BYTES, removing the oldest versions of any document. */
let pruneTimer = null;
function pruneSoon() { if (!pruneTimer) pruneTimer = setTimeout(() => { pruneTimer = null; try { prune(); } catch { /* next time */ } }, 5000); }
function prune(limit = TOTAL_BYTES) {
  let dirs = [];
  try { dirs = fs.readdirSync(root()).map(d => path.join(root(), d)); } catch { return; }
  const all = [];
  for (const d of dirs) for (const v of readIndex(d).versions) all.push({ dir: d, v });
  let total = all.reduce((n, x) => n + (x.v.size || 0), 0);
  if (total <= limit) return;
  all.sort((a, b) => a.v.saved - b.v.saved);
  const drop = new Map();
  for (const x of all) {
    if (total <= limit) break;
    total -= x.v.size || 0;
    if (!drop.has(x.dir)) drop.set(x.dir, new Set());
    drop.get(x.dir).add(x.v.id);
  }
  for (const [d, ids] of drop) {
    const idx = readIndex(d);
    ids.forEach(id => { try { fs.unlinkSync(path.join(d, id + '.bin')); } catch { /* gone */ } });
    idx.versions = idx.versions.filter(v => !ids.has(v.id));
    if (idx.versions.length) writeIndex(d, idx); else fs.rmSync(d, { recursive: true, force: true });
  }
}
ipcMain.handle('versions:list', (_e, key) => {
  if (!validKey(key)) return [];
  return readIndex(dirOf(key)).versions.map(v => ({ id: v.id, name: v.name, by: v.by || [], size: v.size, saved: v.saved })).reverse();
});
ipcMain.handle('versions:read', (_e, key, id) => {
  if (!validKey(key) || !/^\d{1,16}$/.test(String(id))) return null;
  try { return new Uint8Array(fs.readFileSync(path.join(dirOf(key), id + '.bin'))); } catch { return null; }
});

module.exports = { KEEP, prune };
