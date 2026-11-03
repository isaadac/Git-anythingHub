'use strict';
// Isadek Notes storage: one JSON file per notebook in the app's data folder. Every change is
// saved as you type, so there is no Save command. Deleted notebooks go to a trash folder.
const { app, ipcMain } = require('electron');
const fs = require('fs');
const path = require('path');

const MAX = 200e6;
const dir = () => path.join(app.getPath('userData'), 'notebooks');
const validId = id => typeof id === 'string' && /^[\w-]{1,40}$/.test(id);
const file = id => path.join(dir(), id + '.json');

ipcMain.handle('notes:list', () => {
  try {
    return fs.readdirSync(dir()).filter(n => n.endsWith('.json')).map(n => {
      const id = n.slice(0, -5);
      try {
        // Only the head of the file is needed for the list.
        const fd = fs.openSync(file(id), 'r');
        const buf = Buffer.alloc(4096);
        const len = fs.readSync(fd, buf, 0, 4096, 0);
        fs.closeSync(fd);
        const head = buf.toString('utf8', 0, len);
        const name = (/"name"\s*:\s*"((?:[^"\\]|\\.)*)"/.exec(head) || [])[1];
        const color = (/"color"\s*:\s*"(#[0-9a-fA-F]{6})"/.exec(head) || [])[1];
        return { id, name: name ? JSON.parse(`"${name}"`) : id, color: color || '#7719aa', modified: fs.statSync(file(id)).mtimeMs };
      } catch { return null; }
    }).filter(Boolean).sort((a, b) => b.modified - a.modified);
  } catch { return []; }
});
ipcMain.handle('notes:read', (_e, id) => {
  if (!validId(id)) return null;
  try { return fs.readFileSync(file(id), 'utf8'); } catch { return null; }
});
ipcMain.handle('notes:write', (_e, id, json) => {
  if (!validId(id) || typeof json !== 'string' || json.length > MAX) return false;
  fs.mkdirSync(dir(), { recursive: true });
  // Write a temporary file and rename it, so a crash can't leave half a notebook.
  const tmp = file(id) + '.tmp';
  fs.writeFileSync(tmp, json);
  fs.renameSync(tmp, file(id));
  return true;
});
ipcMain.handle('notes:delete', (_e, id) => {
  if (!validId(id)) return false;
  const trash = path.join(dir(), '.trash');
  fs.mkdirSync(trash, { recursive: true });
  try { fs.renameSync(file(id), path.join(trash, `${id}-${Date.now()}.json`)); return true; } catch { return false; }
});
