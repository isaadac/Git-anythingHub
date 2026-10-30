'use strict';
// Every message a page sends to the main process is checked to come from one of the app's own
// pages (src/renderer, loaded from disk), in a window's main frame. So nothing else that might
// ever load in a window — a page from the web, a frame, an injected document — can use the
// app's file, network or account functions. Loaded first, before anything registers a handler.
const { ipcMain } = require('electron');
const path = require('path');
const { fileURLToPath } = require('url');

const RENDERER = path.join(__dirname, '../renderer') + path.sep;

function trusted(e) {
  const f = e && e.senderFrame;
  if (!f || !e.sender || f !== e.sender.mainFrame) return false;
  try {
    const u = new URL(f.url);
    if (u.protocol !== 'file:') return false;
    return path.normalize(fileURLToPath(u)).startsWith(RENDERER);
  } catch { return false; }
}

const handle = ipcMain.handle.bind(ipcMain);
ipcMain.handle = (channel, fn) => handle(channel, (e, ...args) => {
  if (!trusted(e)) throw new Error('Not allowed');
  return fn(e, ...args);
});
const on = ipcMain.on.bind(ipcMain);
ipcMain.on = (channel, fn) => on(channel, (e, ...args) => {
  if (!trusted(e)) { e.returnValue = null; return; }
  return fn(e, ...args);
});

module.exports = { trusted };
