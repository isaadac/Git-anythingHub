'use strict';
// Automatic updates from the project's GitHub Releases (electron-updater).
// The release workflow uploads latest*.yml next to the installers; the app reads
// .../releases/latest/download/latest*.yml to find newer versions.
const { app, dialog, BrowserWindow } = require('electron');
const settings = require('./settings');

const CHECK_EVERY = 6 * 60 * 60 * 1000; // 6 hours
let autoUpdater = null;
let manual = false;
let timer = null;
let downloaded = false;

function updater() {
  if (autoUpdater) return autoUpdater;
  ({ autoUpdater } = require('electron-updater'));
  autoUpdater.autoDownload = true;
  autoUpdater.autoInstallOnAppQuit = true;
  autoUpdater.logger = process.env.ATLAS_UPDATE_DEBUG ? console : null;
  // Self-hosted update server (and testing): ATLAS_UPDATE_URL=https://example.com/atlas
  if (process.env.ATLAS_UPDATE_URL) autoUpdater.setFeedURL({ provider: 'generic', url: process.env.ATLAS_UPDATE_URL });

  autoUpdater.on('update-available', info => {
    if (manual) dialog.showMessageBox({ type: 'info', message: `Atlas Office ${info.version} is available`, detail: 'It is downloading in the background. You will be asked to restart when it is ready.' });
  });
  autoUpdater.on('update-not-available', () => {
    if (manual) dialog.showMessageBox({ type: 'info', message: 'Atlas Office is up to date', detail: `You have version ${app.getVersion()}.` });
    manual = false;
  });
  autoUpdater.on('error', err => {
    if (manual) dialog.showMessageBox({ type: 'warning', message: 'Could not check for updates', detail: String(err && err.message || err).split('\n')[0] });
    manual = false;
  });
  autoUpdater.on('update-downloaded', async info => {
    manual = false;
    if (downloaded) return;
    downloaded = true;
    const win = BrowserWindow.getFocusedWindow() || BrowserWindow.getAllWindows()[0];
    const { response } = await dialog.showMessageBox(win, {
      type: 'info',
      buttons: ['Restart Now', 'Later'],
      defaultId: 0,
      cancelId: 1,
      message: `Atlas Office ${info.version} is ready to install`,
      detail: 'Restart to finish updating. If you choose Later, the update installs the next time you quit.'
    });
    // Windows ask to save unsaved work as usual when they close.
    if (response === 0) setImmediate(() => autoUpdater.quitAndInstall());
  });
  return autoUpdater;
}

/** Updates only apply to installed builds, not to `npm start` or the portable .exe. */
function supported() {
  if (!app.isPackaged) return false;
  if (process.platform === 'win32' && process.env.PORTABLE_EXECUTABLE_DIR) return false;
  if (process.platform === 'linux' && !process.env.APPIMAGE && !isDeb()) return false;
  return true;
}
function isDeb() {
  // electron-builder marks .deb installs with resources/package-type.
  try { return require('fs').readFileSync(require('path').join(process.resourcesPath, 'package-type'), 'utf8').trim() === 'deb'; } catch { return false; }
}

function check(isManual) {
  if (!supported()) {
    if (isManual) dialog.showMessageBox({ type: 'info', message: 'Updates are not available for this copy', detail: app.isPackaged ? 'Download the installer from the Releases page to get automatic updates.' : 'Automatic updates only work in installed builds.' });
    return;
  }
  manual = !!isManual;
  updater().checkForUpdates().catch(() => { /* reported through the 'error' event */ });
}

function start() {
  clearInterval(timer);
  if (!settings.get('autoUpdate') || !supported()) return;
  setTimeout(() => check(false), +process.env.ATLAS_UPDATE_DELAY || 15000);
  timer = setInterval(() => check(false), CHECK_EVERY);
}

function setEnabled(on) {
  settings.set('autoUpdate', !!on);
  if (on) start(); else clearInterval(timer);
}

module.exports = { start, check, setEnabled, supported };
