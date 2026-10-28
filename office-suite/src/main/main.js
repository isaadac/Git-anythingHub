'use strict';
// Isadek Office — Electron main process.
// Owns windows, native menus, file dialogs, file associations and the recent-files list.
const { app, BrowserWindow, Menu, dialog, ipcMain, nativeTheme, screen, session, shell } = require('electron');
const { LANGS, menuText } = require('./menu-i18n');
const path = require('path');
const fs = require('fs');
const settings = require('./settings');
const crash = require('./crash');
crash.start();
const updater = require('./updater');
const recovery = require('./recovery');
const cloud = require('./cloud');
require('./assistant');
require('./notes-store');
require('./versions');
const os = require('os');
const { createCollabServer } = require('./collab-server');

const APPS = {
  home: { page: 'index.html', title: 'Isadek Office', width: 1100, height: 720 },
  write: { page: 'write.html', title: 'Isadek Write', width: 1280, height: 860 },
  sheets: { page: 'sheets.html', title: 'Isadek Sheets', width: 1360, height: 860 },
  slides: { page: 'slides.html', title: 'Isadek Slides', width: 1360, height: 860 },
  notes: { page: 'notes.html', title: 'Isadek Notes', width: 1280, height: 840 },
  help: { page: 'help.html', title: 'Isadek Office Help', width: 1000, height: 760 }
};

const EXT_TO_APP = {
  '.docx': 'write', '.html': 'write', '.htm': 'write', '.txt': 'write', '.md': 'write',
  '.xlsx': 'sheets', '.xls': 'sheets', '.csv': 'sheets', '.ods': 'sheets', '.tsv': 'sheets',
  '.pptx': 'slides', '.aslides': 'slides', '.anote': 'notes'
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
    backgroundColor: nativeTheme.shouldUseDarkColors ? '#292929' : '#f3f2f1',
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
  // Right-click in text: spelling suggestions, Add to Dictionary, and the usual editing commands.
  win.webContents.on('context-menu', (_e, p) => {
    if (!p.isEditable) return;
    const wc = win.webContents;
    const items = [];
    if (p.misspelledWord) {
      p.dictionarySuggestions.slice(0, 6).forEach(w => items.push({ label: w, click: () => wc.replaceMisspelling(w) }));
      if (!p.dictionarySuggestions.length) items.push({ label: 'No suggestions', enabled: false });
      items.push({ type: 'separator' }, { label: 'Add to Dictionary', click: () => wc.session.addWordToSpellCheckerDictionary(p.misspelledWord) }, { type: 'separator' });
    }
    items.push({ role: 'cut', enabled: p.editFlags.canCut }, { role: 'copy', enabled: p.editFlags.canCopy }, { role: 'paste', enabled: p.editFlags.canPaste }, { type: 'separator' }, { role: 'selectAll' });
    Menu.buildFromTemplate(items).popup({ window: win });
  });

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
    dialog.showErrorBox('Unsupported file', `Isadek Office can't open "${path.basename(filePath)}".`);
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
  // New, unsaved content (a template, a mail-merge result): no file behind it.
  if (typeof p === 'object') return { path: p.cloudPath || null, name: p.name, data: p.data };
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
  const pdf = await e.sender.printToPDF({ printBackground: true, landscape: !!landscape, pageSize: 'Letter', preferCSSPageSize: true, generateTaggedPDF: true, generateDocumentOutline: true });
  fs.writeFileSync(target, pdf);
  shell.openPath(target);
  return target;
});

ipcMain.handle('app:print', e => new Promise(resolve => e.sender.print({ printBackground: true }, ok => resolve(ok))));
ipcMain.handle('app:recent', () => readRecent().filter(r => fs.existsSync(r.path)).map(r => ({ ...r, path: approve(r.path) })));
ipcMain.handle('app:clear-recent', () => { try { fs.unlinkSync(recentPath()); } catch { /* none */ } app.clearRecentDocuments(); return true; });
ipcMain.handle('app:new-window', (_e, appName) => { createWindow(APPS[appName] ? appName : 'home'); return true; });
ipcMain.handle('app:open-data', (_e, { app: appName, name, data, cloudPath }) => {
  if (!APPS[appName] || appName === 'home' || typeof name !== 'string' || !(data instanceof Uint8Array) || data.length > 300e6) return false;
  const win = createWindow(appName);
  // A file from a cloud account keeps its cloud address, so Save writes back to it.
  pendingFiles.set(win.webContents.id, { name: path.basename(name).slice(0, 255), data, cloudPath: cloud.isCloudPath(cloudPath) ? cloudPath : null });
  return true;
});
// ---------- Live editing: a session server on this computer, for people on the same network ----------
let collabHost = null;
const pendingCollab = new Map(); // webContents.id -> invite link to join on load
function listen(srv, port) {
  return new Promise((resolve, reject) => {
    const onErr = err => { srv.removeListener('listening', onOk); reject(err); };
    const onOk = () => { srv.removeListener('error', onErr); resolve(); };
    srv.once('error', onErr);
    srv.once('listening', onOk);
    srv.listen(port, '0.0.0.0');
  });
}
ipcMain.handle('collab:host', async () => {
  try {
    if (!collabHost) {
      const srv = createCollabServer();
      try { await listen(srv, 8787); } catch { await listen(srv, 0); }
      collabHost = srv;
    }
    const port = collabHost.address().port;
    const ips = Object.values(os.networkInterfaces()).flat().filter(i => i && i.family === 'IPv4' && !i.internal).map(i => i.address);
    return { urls: [...ips, '127.0.0.1'].map(ip => `ws://${ip}:${port}/`) };
  } catch (err) {
    return { error: String(err.message || err) };
  }
});
ipcMain.handle('collab:join', (_e, appName, link) => {
  if (!APPS[appName] || appName === 'home' || typeof link !== 'string' || link.length > 2000) return false;
  const win = createWindow(appName);
  pendingCollab.set(win.webContents.id, link);
  return true;
});
ipcMain.handle('collab:pending', e => { const l = pendingCollab.get(e.sender.id) || null; pendingCollab.delete(e.sender.id); return l; });
// ---------- Theme ----------
// The pages keep the choice in their shared storage; this side makes native menus, dialogs and
// window backgrounds match it.
function applyTheme(mode) {
  nativeTheme.themeSource = ['light', 'dark'].includes(mode) ? mode : 'system';
  const bg = nativeTheme.shouldUseDarkColors ? '#292929' : '#f3f2f1';
  BrowserWindow.getAllWindows().forEach(w => { if (!w.isDestroyed()) w.setBackgroundColor(bg); });
}
app.whenReady().then(() => applyTheme(settings.get('theme')));
ipcMain.handle('app:set-theme', (_e, mode) => {
  if (!['system', 'light', 'dark'].includes(mode)) return false;
  settings.set('theme', mode);
  applyTheme(mode);
  return true;
});

// ---------- Sheets: Get Data from the Web ----------
// A plain GET (no cookies or custom headers), size- and time-limited; the page only parses the result.
ipcMain.handle('net:fetch-data', async (_e, url) => {
  let u;
  try { u = new URL(String(url)); } catch { return { error: 'Enter a web address such as https://example.com/data.csv' }; }
  if (!/^https?:$/.test(u.protocol)) return { error: 'Only http and https addresses can be used' };
  const ctl = new AbortController();
  const timer = setTimeout(() => ctl.abort(), 30000);
  try {
    const r = await fetch(u, { redirect: 'follow', signal: ctl.signal, headers: { 'User-Agent': 'IsadekOffice/' + app.getVersion() } });
    if (!r.ok) return { error: `The site answered ${r.status}` };
    if (+r.headers.get('content-length') > 50e6) return { error: 'That file is too large (over 50 MB)' };
    const buf = new Uint8Array(await r.arrayBuffer());
    if (buf.length > 50e6) return { error: 'That file is too large (over 50 MB)' };
    return { bytes: buf, contentType: r.headers.get('content-type') || '' };
  } catch (err) {
    return { error: ctl.signal.aborted ? 'The site took too long to answer' : 'Could not reach the site: ' + (err.message || err) };
  } finally { clearTimeout(timer); }
});

// ---------- Proofing languages ----------
function applySpellLanguages() {
  const want = settings.get('spellLanguages');
  if (process.platform === 'darwin' || !Array.isArray(want) || !want.length) return; // macOS picks the language itself
  const avail = new Set(session.defaultSession.availableSpellCheckerLanguages);
  const langs = want.filter(l => avail.has(l)).slice(0, 4);
  if (langs.length) session.defaultSession.setSpellCheckerLanguages(langs);
}
app.whenReady().then(applySpellLanguages);
ipcMain.handle('spell:languages', () => ({
  available: process.platform === 'darwin' ? [] : session.defaultSession.availableSpellCheckerLanguages,
  current: process.platform === 'darwin' ? [] : session.defaultSession.getSpellCheckerLanguages()
}));
ipcMain.handle('spell:set-languages', (_e, langs) => {
  if (!Array.isArray(langs)) return false;
  settings.set('spellLanguages', langs.filter(l => typeof l === 'string' && /^[a-zA-Z-]{2,12}$/.test(l)).slice(0, 4));
  applySpellLanguages();
  return true;
});
ipcMain.handle('spell:add-word', (_e, w) => typeof w === 'string' && w.length < 100 && session.defaultSession.addWordToSpellCheckerDictionary(w));
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
// ---------- Presenter View: an audience window driven by the presenter's window ----------
const showPairs = new Map(); // webContents id -> the other window of the pair
function showWindowFor(presenter) {
  const displays = screen.getAllDisplays();
  const here = screen.getDisplayMatching(presenter.getBounds());
  const other = displays.find(d => d.id !== here.id);
  const audience = new BrowserWindow({
    ...(other ? { x: other.bounds.x, y: other.bounds.y, width: other.bounds.width, height: other.bounds.height, fullscreen: true } : { width: 960, height: 580 }),
    title: 'Slide Show', backgroundColor: '#000000', autoHideMenuBar: true,
    icon: path.join(__dirname, '../../build/icon.png'),
    webPreferences: { preload: path.join(__dirname, 'preload.js'), contextIsolation: true, nodeIntegration: false, sandbox: true }
  });
  audience.setMenuBarVisibility(false);
  audience.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
  audience.webContents.on('will-navigate', e => e.preventDefault());
  const pid = presenter.webContents.id, aid = audience.webContents.id;
  showPairs.set(pid, audience); showPairs.set(aid, presenter);
  audience.on('closed', () => {
    showPairs.delete(pid); showPairs.delete(aid);
    if (!presenter.isDestroyed()) presenter.webContents.send('show:message', { cmd: 'closed' });
  });
  presenter.once('closed', () => { if (!audience.isDestroyed()) audience.close(); });
  audience.loadFile(path.join(__dirname, '../renderer/slides.html'), { query: { audience: '1' } });
  return audience;
}
ipcMain.handle('show:open', e => {
  const win = BrowserWindow.fromWebContents(e.sender);
  if (!win) return false;
  const old = showPairs.get(e.sender.id);
  if (old && !old.isDestroyed()) old.close();
  showWindowFor(win);
  return true;
});
ipcMain.handle('show:send', (e, msg) => {
  if (!msg || typeof msg !== 'object' || typeof msg.cmd !== 'string') return false;
  const to = showPairs.get(e.sender.id);
  if (!to || to.isDestroyed()) return false;
  to.webContents.send('show:message', msg);
  return true;
});
ipcMain.handle('show:close', e => { const to = showPairs.get(e.sender.id); if (to && !to.isDestroyed()) to.close(); return true; });
ipcMain.handle('app:fullscreen', (e, on) => { BrowserWindow.fromWebContents(e.sender)?.setFullScreen(!!on); return true; });

recovery.init({ createWindow, approve, isApproved });

// Interface language: the pages translate themselves; the native menus are rebuilt here.
ipcMain.handle('app:set-language', (_e, lang) => {
  if (!LANGS.includes(lang)) return false;
  settings.set('language', lang);
  buildMenu();
  return true;
});

// ---------- Menu ----------
/** The Help window, at one of its sections. */
function openHelp(section) {
  const open = BrowserWindow.getAllWindows().find(w => !w.isDestroyed() && w.webContents.getURL().includes('help.html'));
  const go = w => w.webContents.executeJavaScript(`location.hash = ${JSON.stringify('#' + section)}`).catch(() => {});
  if (open) { open.focus(); go(open); return; }
  const w = createWindow('help');
  w.webContents.once('did-finish-load', () => go(w));
}
ipcMain.handle('app:help', (_e, section) => { openHelp(/^[\w-]{1,40}$/.test(String(section)) ? String(section) : 'start'); return true; });
ipcMain.handle('app:crash-reports', (_e, on) => { if (typeof on === 'boolean') { crash.setEnabled(on); buildMenu(); } return crash.enabled(); });

function send(cmd) {
  const win = BrowserWindow.getFocusedWindow();
  if (win) win.webContents.send('menu', cmd);
}

function buildMenu() {
  const isMac = process.platform === 'darwin';
  const L = menuText(settings.get('language'));
  const template = [
    ...(isMac ? [{ role: 'appMenu' }] : []),
    {
      label: L('File'),
      submenu: [
        { label: L('New Document'), accelerator: 'CmdOrCtrl+Shift+W', click: () => createWindow('write') },
        { label: L('New Spreadsheet'), accelerator: 'CmdOrCtrl+Shift+E', click: () => createWindow('sheets') },
        { label: L('New Presentation'), accelerator: 'CmdOrCtrl+Shift+P', click: () => createWindow('slides') },
        { label: L('Isadek Notes'), accelerator: 'CmdOrCtrl+Shift+O', click: () => createWindow('notes') },
        { type: 'separator' },
        { label: L('New'), accelerator: 'CmdOrCtrl+N', click: () => send('new') },
        { label: L('Open…'), accelerator: 'CmdOrCtrl+O', click: () => send('open') },
        { label: L('Save'), accelerator: 'CmdOrCtrl+S', click: () => send('save') },
        { label: L('Save As…'), accelerator: 'CmdOrCtrl+Shift+S', click: () => send('save-as') },
        { type: 'separator' },
        { label: L('Export as PDF…'), click: () => send('export-pdf') },
        { label: L('Print…'), accelerator: 'CmdOrCtrl+P', click: () => send('print') },
        { type: 'separator' },
        { label: L('Home'), click: () => createWindow('home') },
        isMac ? { role: 'close', label: L('Close') } : { role: 'quit', label: L('Quit') }
      ]
    },
    {
      label: L('Edit'),
      submenu: [
        { label: L('Undo'), accelerator: 'CmdOrCtrl+Z', click: () => send('undo') },
        { label: L('Redo'), accelerator: isMac ? 'Cmd+Shift+Z' : 'Ctrl+Y', click: () => send('redo') },
        { type: 'separator' },
        { role: 'cut', label: L('Cut') }, { role: 'copy', label: L('Copy') }, { role: 'paste', label: L('Paste') }, { role: 'selectAll', label: L('Select All') },
        { type: 'separator' },
        { label: L('Find and Replace…'), accelerator: 'CmdOrCtrl+F', click: () => send('find') }
      ]
    },
    {
      label: L('View'),
      submenu: [
        { role: 'zoomIn', label: L('Zoom In') }, { role: 'zoomOut', label: L('Zoom Out') }, { role: 'resetZoom', label: L('Actual Size') },
        { type: 'separator' },
        { role: 'togglefullscreen', label: L('Toggle Full Screen') },
        { role: 'toggleDevTools', label: L('Toggle Developer Tools') }
      ]
    },
    { role: 'windowMenu', label: L('Window') },
    {
      role: 'help',
      label: L('Help'),
      submenu: [{
        label: L('Help and Support'),
        accelerator: 'F1',
        click: () => createWindow('help')
      }, {
        label: L('Privacy'),
        click: () => openHelp('privacy')
      }, { type: 'separator' }, {
        label: L('Send Crash Reports'),
        type: 'checkbox',
        checked: crash.enabled(),
        click: item => crash.setEnabled(item.checked)
      }, {
        label: L('Show Crash Reports'),
        click: () => crash.openFolder()
      }, { type: 'separator' }, {
        label: L('Check for Updates…'),
        click: () => updater.check(true)
      }, {
        label: L('Update Automatically'),
        type: 'checkbox',
        checked: !!settings.get('autoUpdate'),
        click: item => updater.setEnabled(item.checked)
      }, { type: 'separator' }, {
        label: L('About Isadek Office'),
        click: () => dialog.showMessageBox({
          type: 'info',
          title: 'About Isadek Office',
          message: `Isadek Office ${app.getVersion()}`,
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
    // Only the permissions the apps use: clipboard (Paste button), full screen (slide show) and the
    // microphone (recording narration in Slides and audio notes) — never the camera or screen.
    const allowed = new Set(['clipboard-read', 'clipboard-sanitized-write', 'fullscreen']);
    const audioOnly = d => !!d && Array.isArray(d.mediaTypes) && d.mediaTypes.length > 0 && d.mediaTypes.every(t => t === 'audio');
    session.defaultSession.setPermissionRequestHandler((_wc, permission, cb, details) => cb(allowed.has(permission) || (permission === 'media' && audioOnly(details))));
    session.defaultSession.setPermissionCheckHandler((_wc, permission, _origin, details) => allowed.has(permission) || (permission === 'media' && (!details || !details.mediaType || details.mediaType === 'audio')));
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
