'use strict';
// Atlas Office — Electron main process.
// Owns windows, native menus, file dialogs, file associations and the recent-files list.
const { app, BrowserWindow, Menu, dialog, ipcMain, shell } = require('electron');
const path = require('path');
const fs = require('fs');

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
  win.loadFile(path.join(__dirname, '../renderer', cfg.page));

  // Open external links in the user's browser, never inside the app.
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (/^https?:/i.test(url)) shell.openExternal(url);
    return { action: 'deny' };
  });
  win.webContents.on('will-navigate', (e, url) => {
    if (!url.startsWith('file://')) { e.preventDefault(); shell.openExternal(url); }
  });

  // Ask the renderer whether there are unsaved changes before closing.
  let allowClose = false;
  win.on('close', e => {
    if (allowClose) return;
    e.preventDefault();
    win.webContents.executeJavaScript('window.__atlasDirty === true').catch(() => false).then(dirty => {
      if (!dirty) { allowClose = true; win.close(); return; }
      const choice = dialog.showMessageBoxSync(win, {
        type: 'warning',
        buttons: ['Save', "Don't Save", 'Cancel'],
        defaultId: 0,
        cancelId: 2,
        message: 'Do you want to save your changes?',
        detail: 'Your changes will be lost if you don\'t save them.'
      });
      if (choice === 1) { allowClose = true; win.close(); }
      else if (choice === 0) win.webContents.send('menu', 'save-and-close');
    });
  });
  return win;
}

function openPath(filePath, fromWin) {
  const ext = path.extname(filePath).toLowerCase();
  const appName = EXT_TO_APP[ext];
  if (!appName) {
    dialog.showErrorBox('Unsupported file', `Atlas Office can't open "${path.basename(filePath)}".`);
    return;
  }
  createWindow(appName, filePath);
  if (fromWin && fromWin.__isEmptyHome) fromWin.close();
}

// ---------- IPC ----------
ipcMain.handle('file:read', (_e, filePath) => {
  const data = fs.readFileSync(filePath);
  addRecent(filePath);
  return { path: filePath, name: path.basename(filePath), data: new Uint8Array(data) };
});

ipcMain.handle('file:open-dialog', async (e, filters) => {
  const win = BrowserWindow.fromWebContents(e.sender);
  const res = await dialog.showOpenDialog(win, { properties: ['openFile'], filters });
  if (res.canceled || !res.filePaths.length) return null;
  const filePath = res.filePaths[0];
  addRecent(filePath);
  return { path: filePath, name: path.basename(filePath), data: new Uint8Array(fs.readFileSync(filePath)) };
});

ipcMain.handle('file:save', async (e, { data, filePath, defaultName, filters, saveAs }) => {
  const win = BrowserWindow.fromWebContents(e.sender);
  let target = filePath;
  if (!target || saveAs) {
    const res = await dialog.showSaveDialog(win, { defaultPath: target || defaultName, filters });
    if (res.canceled || !res.filePath) return null;
    target = res.filePath;
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
  return { path: res.filePath, name: path.basename(res.filePath) };
});

ipcMain.handle('file:pending', e => {
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
  const pdf = await e.sender.printToPDF({ printBackground: true, landscape: !!landscape, pageSize: 'Letter', preferCSSPageSize: true });
  fs.writeFileSync(res.filePath, pdf);
  shell.openPath(res.filePath);
  return res.filePath;
});

ipcMain.handle('app:print', e => new Promise(resolve => e.sender.print({ printBackground: true }, ok => resolve(ok))));
ipcMain.handle('app:recent', () => readRecent().filter(r => fs.existsSync(r.path)));
ipcMain.handle('app:clear-recent', () => { try { fs.unlinkSync(recentPath()); } catch { /* none */ } app.clearRecentDocuments(); return true; });
ipcMain.handle('app:new-window', (_e, appName) => { createWindow(appName); return true; });
ipcMain.handle('app:open-path', (e, p) => { openPath(p, BrowserWindow.fromWebContents(e.sender)); return true; });
ipcMain.handle('app:set-title', (e, title) => { BrowserWindow.fromWebContents(e.sender)?.setTitle(title); return true; });
ipcMain.handle('app:close-now', e => {
  const win = BrowserWindow.fromWebContents(e.sender);
  if (win) { win.removeAllListeners('close'); win.close(); }
  return true;
});
ipcMain.handle('app:fullscreen', (e, on) => { BrowserWindow.fromWebContents(e.sender)?.setFullScreen(!!on); return true; });

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
const fileArgs = argv => argv.slice(app.isPackaged ? 1 : 2).filter(a => !a.startsWith('-') && fs.existsSync(a) && EXT_TO_APP[path.extname(a).toLowerCase()]);
let macOpenQueue = [];

app.on('open-file', (e, filePath) => { // macOS: double-click in Finder
  e.preventDefault();
  if (app.isReady()) openPath(filePath); else macOpenQueue.push(filePath);
});

if (!app.requestSingleInstanceLock()) {
  app.quit();
} else {
  app.on('second-instance', (_e, argv) => {
    const files = fileArgs(argv);
    if (files.length) files.forEach(f => openPath(f));
    else createWindow('home');
  });

  app.whenReady().then(() => {
    buildMenu();
    const files = [...fileArgs(process.argv), ...macOpenQueue];
    macOpenQueue = [];
    if (files.length) files.forEach(f => openPath(f));
    else createWindow('home');
    app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow('home'); });
  });

  app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });
}
