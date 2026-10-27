'use strict';
// Exposes a small, safe API to the renderer. The renderer never gets Node access.
const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('atlas', {
  isDesktop: true,
  platform: process.platform,
  openDialog: filters => ipcRenderer.invoke('file:open-dialog', filters),
  save: opts => ipcRenderer.invoke('file:save', opts),
  chooseSavePath: opts => ipcRenderer.invoke('file:choose-save-path', opts),
  pendingFile: () => ipcRenderer.invoke('file:pending'),
  exportPdf: opts => ipcRenderer.invoke('file:export-pdf', opts),
  print: () => ipcRenderer.invoke('app:print'),
  recent: () => ipcRenderer.invoke('app:recent'),
  clearRecent: () => ipcRenderer.invoke('app:clear-recent'),
  newWindow: appName => ipcRenderer.invoke('app:new-window', appName),
  openData: (appName, name, data) => ipcRenderer.invoke('app:open-data', { app: appName, name, data }),
  openPath: p => ipcRenderer.invoke('app:open-path', p),
  setTitle: t => ipcRenderer.invoke('app:set-title', t),
  closeNow: () => ipcRenderer.invoke('app:close-now'),
  fullscreen: on => ipcRenderer.invoke('app:fullscreen', on),
  recoverySave: d => ipcRenderer.invoke('recovery:save', d),
  recoveryClear: () => ipcRenderer.invoke('recovery:clear'),
  recoveryList: () => ipcRenderer.invoke('recovery:list'),
  recoveryOpen: id => ipcRenderer.invoke('recovery:open', id),
  recoveryDiscard: id => ipcRenderer.invoke('recovery:discard', id),
  onMenu: cb => ipcRenderer.on('menu', (_e, cmd) => cb(cmd))
});
