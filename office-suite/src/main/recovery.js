'use strict';
// Crash recovery: windows with unsaved changes keep a draft in userData/recovery.
// Drafts are removed when the document is saved or closed on purpose, so any draft
// still there at startup belongs to a session that crashed or lost power.
const { app, ipcMain, BrowserWindow, dialog } = require('electron');
const fs = require('fs');
const path = require('path');

const ID = /^[a-z0-9-]{8,64}$/;
const APPS = new Set(['write', 'sheets', 'slides', 'notes']);
const dir = () => path.join(app.getPath('userData'), 'recovery');
const fileFor = id => path.join(dir(), id + '.json');
const draftOf = new Map();       // webContents.id -> draft id
const pendingRecovery = new Map(); // webContents.id -> draft to restore on load

function remove(id) { if (ID.test(id || '')) { try { fs.unlinkSync(fileFor(id)); } catch { /* already gone */ } } }
function read(id) {
  if (!ID.test(id || '')) return null;
  try { return JSON.parse(fs.readFileSync(fileFor(id), 'utf8')); } catch { return null; }
}
function list() {
  let names = [];
  try { names = fs.readdirSync(dir()).filter(n => n.endsWith('.json')); } catch { return []; }
  const open = new Set(draftOf.values()); // drafts already open in a window aren't offered again
  return names.map(n => {
    const id = n.slice(0, -5);
    if (open.has(id)) return null;
    const d = read(id);
    return d && APPS.has(d.app) ? { id, app: d.app, name: String(d.name || 'Untitled'), savedAt: d.savedAt || 0, path: d.path || null } : null;
  }).filter(Boolean).sort((a, b) => b.savedAt - a.savedAt);
}

let hooks = null;
/**
 * h: { createWindow(app), approve(path), isApproved(path) }
 */
function init(h) {
  hooks = h;
  ipcMain.handle('recovery:save', (e, d) => {
    if (!d || !ID.test(d.draftId || '') || !APPS.has(d.app) || typeof d.data !== 'string') return false;
    if (d.data.length > 200e6) return false;
    fs.mkdirSync(dir(), { recursive: true });
    // Only remember a file path the user already chose, so a draft can't be used to gain write access.
    const filePath = d.path && hooks.isApproved(d.path) ? d.path : null;
    const draft = { app: d.app, name: String(d.name || 'Untitled').slice(0, 255), path: filePath, data: d.data, savedAt: Date.now() };
    const tmp = fileFor(d.draftId) + '.tmp';
    fs.writeFileSync(tmp, JSON.stringify(draft));
    fs.renameSync(tmp, fileFor(d.draftId)); // atomic: a crash mid-write never leaves half a draft
    draftOf.set(e.sender.id, d.draftId);
    return true;
  });
  ipcMain.handle('recovery:clear', e => { remove(draftOf.get(e.sender.id)); return true; });
  ipcMain.handle('recovery:list', () => list());
  ipcMain.handle('recovery:discard', (_e, id) => { remove(id); return true; });
  ipcMain.handle('recovery:open', (_e, id) => {
    const d = read(id);
    if (!d || !APPS.has(d.app)) return false;
    if (d.path) hooks.approve(d.path);
    const win = hooks.createWindow(d.app);
    pendingRecovery.set(win.webContents.id, { draftId: id, name: d.name, path: d.path, data: d.data });
    return true;
  });
}

/** Called by the window code when a document is closed on purpose ("Don't Save" or nothing to save). */
function closedCleanly(wc) {
  remove(draftOf.get(wc.id));
  draftOf.delete(wc.id);
}

/** A renderer that crashed: offer to reload it with its last draft. */
function watchWindow(win) {
  const wcId = win.webContents.id;
  win.webContents.on('render-process-gone', async (_e, details) => {
    if (details.reason === 'clean-exit') return;
    const id = draftOf.get(wcId);
    const { response } = await dialog.showMessageBox(win, {
      type: 'error', buttons: ['Reopen', 'Close'], defaultId: 0,
      message: 'This window stopped working',
      detail: id ? 'Reopen it with your recovered changes?' : 'Reopen it?'
    });
    if (win.isDestroyed()) return;
    if (response === 1) { win.destroy(); return; }
    const d = id && read(id);
    if (d) {
      if (d.path) hooks.approve(d.path);
      pendingRecovery.set(wcId, { draftId: id, name: d.name, path: d.path, data: d.data });
    }
    win.webContents.reload();
  });
  win.on('closed', () => { draftOf.delete(wcId); pendingRecovery.delete(wcId); });
}

function takePending(wcId) {
  const d = pendingRecovery.get(wcId);
  pendingRecovery.delete(wcId);
  if (d) draftOf.set(wcId, d.draftId);
  return d || null;
}

module.exports = { init, list, closedCleanly, watchWindow, takePending };
