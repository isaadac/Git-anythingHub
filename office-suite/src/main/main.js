'use strict';
// Atlas Office — Electron main process.
// Owns windows, native menus, file dialogs, file associations and the recent-files list.
const { app, BrowserWindow, Menu, dialog, ipcMain, session, shell } = require('electron');
const path = require('path');
const fs = require('fs');
const settings = require('./settings');
const updater = require('./updater');
const recovery = require('./recovery');

const APPS = {
  home: { page: 'index.html', title: 'Atlas Office', width: 1100, height: 720 },
  write: { page: 'write.html', title: 'Atlas Write', width: 1280, height: 860 },
  sheets: { page: 'sheets.html', title: 'Atlas Sheets', width: 1360, height: 860 },
  slides: { page: 'slides.html', title: 'Atlas Slides', width: 1360, height: 860 }
};

const EXT_TO_APP = {
  '.docx': 'write', '.html': 'write', '.htm': 'write', '.txt': 'write', '.md': 'write',
  '.xlsx': 'sheets', '.xls': 'sheets', '.csv': 'sheets', '.ods': 'sheets', '.tsv': 'sheets',
  '.pptx': 'slides', '.aslides': 'slides'
};

const pendingFiles = new Map(); // webContents.id -> file path to open on load
// Files the user chose (dialogs, command line, Finder, recent list). The renderer may only
// read or overwrite these, so a compromised page can't touch arbitrary files.
const approvedPaths = new Set();
const approve = p => { const abs = path.resolve(p); approvedPaths.add(abs); return abs; };
const isApproved = p => typeof p === 'string' && approvedPaths.has(path.resolve(p));
const SAFE_EXTERNAL = /^(https?:|mailto:)/i;
const recentPath = () => path.join(app.getPath('userData'), 'recent.json');

function readRecent() {
  try { return JSON.parse(fs.readFileSync(recentPath(), 'utf8')); } catch { return []; }
}

function addRecent(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  const entry = { path: filePath, name: path.basename(filePath), app: EXT_TO_APP[ext] || 'write', opened: Date.now() };
  const list = [entry, ...readRecent().filter(r => r.path !== filePath)].slice(0, 20);
  try { fs.writeFileSync(recentPath(), JSON.stringify(list, null, 2)); } catch { /* non-fatal */ }
  app.addRecentDocument(filePath);
}

function createWindow(appName = 'home', filePath = null) {
  const cfg = APPS[appName] || APPS.home;
  const win = new BrowserWindow({
    width: cfg.width,
    height: cfg.height,
    minWidth: 760,
    minHeight: 500,
    title: cfg.title,
    backgroundColor: '#f3f2f1',
    icon: path.join(__dirname, '../../build/icon.png'),
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      spellcheck: true
    }
  });
  if (filePath) pendingFiles.set(win.webContents.id, filePath);
  recovery.watchWindow(win);
  win.loadFile(path.join(__dirname, '../renderer', cfg.page));

  // Open external links in the user's browser, never inside the app.
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (SAFE_EXTERNAL.test(url)) shell.openExternal(url);
    return { action: 'deny' };
  });
  // The app's pages never navigate; links in documents open in the browser instead.
  win.webContents.on('will-navigate', (e, url) => {
    e.preventDefault();
    if (SAFE_EXTERNAL.test(url)) shell.openExternal(url);
  });
  win.webContents.on('will-attach-webview', e => e.preventDefault());

  // Ask the renderer whether there are unsaved changes before closing.
  let allowClose = false;
  let asking = false;
  win.on('close', e => {
    if (allowClose) return;
    e.preventDefault();
    if (asking) return;
    asking = true;
    // A hung or crashed page must not keep the window open forever.
    const timeout = new Promise(resolve => setTimeout(() => resolve(false), 1500));
    Promise.race([win.webContents.executeJavaScript('window.__atlasDirty === true').catch(() => false), timeout]).then(dirty => {
      asking = false;
      if (win.isDestroyed()) return;
      if (!dirty) { recovery.closedCleanly(win.webContents); allowClose = true; win.close(); return; }
      const choice = dialog.showMessageBoxSync(win, {
        type: 'warning',
        buttons: ['Save', "Don't Save", 'Cancel'],
        defaultId: 0,
        cancelId: 2,
        message: 'Do you want to save your changes?',
        detail: 'Your changes will be lost if you don\'t save them.'
      });
      if (choice === 1) { recovery.closedCleanly(win.webContents); allowClose = true; win.close(); }
      else if (choice === 0) win.webContents.send('menu', 'save-and-close');
    });
  });
  return win;
}

function openPath(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  const appName = EXT_TO_APP[ext];
  if (!appName) {
    dialog.showErrorBox('Unsupported file', `Atlas Office can't open "${path.basename(filePath)}".`);
    return;
  }
  createWindow(appName, approve(filePath));
}

// ---------- IPC ----------

ipcMain.handle('file:open-dialog', async (e, filters) => {
  const win = BrowserWindow.fromWebContents(e.sender);
  const res = await dialog.showOpenDialog(win, { properties: ['openFile'], filters });
  if (res.canceled || !res.filePaths.length) return null;
  const filePath = approve(res.filePaths[0]);
  addRecent(filePath);
  return { path: filePath, name: path.basename(filePath), data: new Uint8Array(fs.readFileSync(filePath)) };
});

ipcMain.handle('file:save', async (e, { data, filePath, defaultName, filters, saveAs }) => {
  const win = BrowserWindow.fromWebContents(e.sender);
  let target = filePath;
  if (target && !isApproved(target)) throw new Error('Choose where to save with Save As.');
  if (!target || saveAs) {
    const res = await dialog.showSaveDialog(win, { defaultPath: target || defaultName, filters });
    if (res.canceled || !res.filePath) return null;
    target = approve(res.filePath);
  }
  const buf = typeof data === 'string' ? Buffer.from(data, 'utf8') : Buffer.from(data);
  fs.writeFileSync(target, buf);
  addRecent(target);
  return { path: target, name: path.basename(target) };
});

ipcMain.handle('file:choose-save-path', async (e, { defaultName, filters }) => {
  const win = BrowserWindow.fromWebContents(e.sender);
  const res = await dialog.showSaveDialog(win, { defaultPath: defaultName, filters });
  if (res.canceled || !res.filePath) return null;
  // Linux save dialogs don't add the extension for you.
  let target = res.filePath;
  const exts = (filters || []).flatMap(f => f.extensions);
  if (exts.length && !exts.includes(path.extname(target).slice(1).toLowerCase())) target += '.' + exts[0];
  approve(target);
  return { path: target, name: path.basename(target) };
});

ipcMain.handle('file:pending', e => {
  const draft = recovery.takePending(e.sender.id);
  if (draft) return { recovered: draft };
  const p = pendingFiles.get(e.sender.id);
  pendingFiles.delete(e.sender.id);
  if (!p) return null;
  try {
    addRecent(p);
    return { path: p, name: path.basename(p), data: new Uint8Array(fs.readFileSync(p)) };
  } catch (err) {
    return { error: String(err.message || err) };
  }
});

ipcMain.handle('file:export-pdf', async (e, { defaultName, landscape }) => {
  const win = BrowserWindow.fromWebContents(e.sender);
  const res = await dialog.showSaveDialog(win, { defaultPath: defaultName, filters: [{ name: 'PDF', extensions: ['pdf'] }] });
  if (res.canceled || !res.filePath) return null;
  const target = /\.pdf$/i.test(res.filePath) ? res.filePath : res.filePath + '.pdf';
  const pdf = await e.sender.printToPDF({ printBackground: true, landscape: !!landscape, pageSize: 'Letter', preferCSSPageSize: true });
  fs.writeFileSync(target, pdf);
  shell.openPath(target);
  return target;
});

ipcMain.handle('app:print', e => new Promise(resolve => e.sender.print({ printBackground: true }, ok => resolve(ok))));
ipcMain.handle('app:recent', () => readRecent().filter(r => fs.existsSync(r.path)).map(r => ({ ...r, path: approve(r.path) })));
ipcMain.handle('app:clear-recent', () => { try { fs.unlinkSync(recentPath()); } catch { /* none */ } app.clearRecentDocuments(); return true; });
ipcMain.handle('app:new-window', (_e, appName) => { createWindow(APPS[appName] ? appName : 'home'); return true; });
ipcMain.handle('app:open-path', (_e, p) => {
  if (!isApproved(p)) return false;
  openPath(p);
  return true;
});
ipcMain.handle('app:set-title', (e, title) => { BrowserWindow.fromWebContents(e.sender)?.setTitle(title); return true; });
ipcMain.handle('app:close-now', e => {
  const win = BrowserWindow.fromWebContents(e.sender);
  if (win) { recovery.closedCleanly(e.sender); win.removeAllListeners('close'); win.close(); }
  return true;
});
ipcMain.handle('app:fullscreen', (e, on) => { BrowserWindow.fromWebContents(e.sender)?.setFullScreen(!!on); return true; });

recovery.init({ createWindow, approve, isApproved });

// ---------- Menu ----------
function send(cmd) {
  const win = BrowserWindow.getFocusedWindow();
  if (win) win.webContents.send('menu', cmd);
}

function buildMenu() {
  const isMac = process.platform === 'darwin';
  const template = [
    ...(isMac ? [{ role: 'appMenu' }] : []),
    {
      label: 'File',
      submenu: [
        { label: 'New Document', accelerator: 'CmdOrCtrl+Shift+W', click: () => createWindow('write') },
        { label: 'New Spreadsheet', accelerator: 'CmdOrCtrl+Shift+E', click: () => createWindow('sheets') },
        { label: 'New Presentation', accelerator: 'CmdOrCtrl+Shift+P', click: () => createWindow('slides') },
        { type: 'separator' },
        { label: 'New', accelerator: 'CmdOrCtrl+N', click: () => send('new') },
        { label: 'Open…', accelerator: 'CmdOrCtrl+O', click: () => send('open') },
        { label: 'Save', accelerator: 'CmdOrCtrl+S', click: () => send('save') },
        { label: 'Save As…', accelerator: 'CmdOrCtrl+Shift+S', click: () => send('save-as') },
        { type: 'separator' },
        { label: 'Export as PDF…', click: () => send('export-pdf') },
        { label: 'Print…', accelerator: 'CmdOrCtrl+P', click: () => send('print') },
        { type: 'separator' },
        { label: 'Home', click: () => createWindow('home') },
        isMac ? { role: 'close' } : { role: 'quit' }
      ]
    },
    {
      label: 'Edit',
      submenu: [
        { label: 'Undo', accelerator: 'CmdOrCtrl+Z', click: () => send('undo') },
        { label: 'Redo', accelerator: isMac ? 'Cmd+Shift+Z' : 'Ctrl+Y', click: () => send('redo') },
        { type: 'separator' },
        { role: 'cut' }, { role: 'copy' }, { role: 'paste' }, { role: 'selectAll' },
        { type: 'separator' },
        { label: 'Find and Replace…', accelerator: 'CmdOrCtrl+F', click: () => send('find') }
      ]
    },
    {
      label: 'View',
      submenu: [
        { role: 'zoomIn' }, { role: 'zoomOut' }, { role: 'resetZoom' },
        { type: 'separator' },
        { role: 'togglefullscreen' },
        { role: 'toggleDevTools' }
      ]
    },
    { role: 'windowMenu' },
    {
      role: 'help',
      submenu: [{
        label: 'Check for Updates…',
        click: () => updater.check(true)
      }, {
        label: 'Update Automatically',
        type: 'checkbox',
        checked: !!settings.get('autoUpdate'),
        click: item => updater.setEnabled(item.checked)
      }, { type: 'separator' }, {
        label: 'About Atlas Office',
        click: () => dialog.showMessageBox({
          type: 'info',
          title: 'About Atlas Office',
          message: `Atlas Office ${app.getVersion()}`,
          detail: 'Write, Sheets and Slides — a free cross-platform office suite.\nOpens and saves .docx, .xlsx, .csv and .pptx files.'
        })
      }]
    }
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

// ---------- Lifecycle ----------
const fileArgs = (argv, cwd = process.cwd()) => argv.slice(app.isPackaged ? 1 : 2)
  .filter(a => !a.startsWith('-'))
  .map(a => path.resolve(cwd, a))
  .filter(a => EXT_TO_APP[path.extname(a).toLowerCase()] && fs.existsSync(a));
let macOpenQueue = [];

app.on('open-file', (e, filePath) => { // macOS: double-click in Finder
  e.preventDefault();
  if (app.isReady()) openPath(filePath); else macOpenQueue.push(filePath);
});

if (!app.requestSingleInstanceLock()) {
  app.quit();
} else {
  app.on('second-instance', (_e, argv, workingDirectory) => {
    const files = fileArgs(argv, workingDirectory);
    if (files.length) files.forEach(f => openPath(f));
    else createWindow('home');
  });

  app.whenReady().then(() => {
    // Only the permissions the apps use: clipboard (Paste button) and full screen (slide show).
    const allowed = new Set(['clipboard-read', 'clipboard-sanitized-write', 'fullscreen']);
    session.defaultSession.setPermissionRequestHandler((_wc, permission, cb) => cb(allowed.has(permission)));
    session.defaultSession.setPermissionCheckHandler((_wc, permission) => allowed.has(permission));
    buildMenu();
    const files = [...fileArgs(process.argv), ...macOpenQueue];
    macOpenQueue = [];
    if (files.length) files.forEach(f => openPath(f));
    // Drafts left behind mean the last session ended unexpectedly: show them on Home.
    if (!files.length || recovery.list().length) createWindow('home');
    updater.start();
    app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow('home'); });
  });

  app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });
}
