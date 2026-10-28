'use strict';
// Isadek Assistant: Claude, with the user's own Anthropic API key. Requests are made here in the
// main process with the official SDK, so the key never reaches a page; replies stream to the
// window that asked.
const { app, ipcMain, safeStorage, BrowserWindow } = require('electron');
const fs = require('fs');
const path = require('path');
const settings = require('./settings');
const { Anthropic, MODELS, validModel, cleanMessages, ask } = require('./assistant-core');

const keyFile = () => path.join(app.getPath('userData'), 'assistant-key.bin');
function readKey() {
  try {
    const raw = fs.readFileSync(keyFile());
    return safeStorage.isEncryptionAvailable() && raw[0] !== 0x73 /* 's' of sk- */ ? safeStorage.decryptString(raw) : raw.toString('utf8');
  } catch { return process.env.ANTHROPIC_API_KEY || ''; }
}
const model = () => validModel(settings.get('assistantModel'));
let client = null, clientKey = null;
function getClient() {
  const key = readKey();
  if (!key) return null;
  if (!client || clientKey !== key) { client = new Anthropic({ apiKey: key }); clientKey = key; }
  return client;
}

ipcMain.handle('ai:status', () => ({ hasKey: !!readKey(), model: model(), models: MODELS, fromEnv: !fs.existsSync(keyFile()) && !!process.env.ANTHROPIC_API_KEY }));
ipcMain.handle('ai:set-key', (_e, key) => {
  key = String(key || '').trim();
  if (key && !/^sk-ant-[\w-]{10,300}$/.test(key)) return { error: 'That does not look like an Anthropic API key (sk-ant-…)' };
  fs.mkdirSync(path.dirname(keyFile()), { recursive: true });
  if (!key) { try { fs.unlinkSync(keyFile()); } catch { /* none */ } return { ok: true }; }
  fs.writeFileSync(keyFile(), safeStorage.isEncryptionAvailable() ? safeStorage.encryptString(key) : Buffer.from(key, 'utf8'), { mode: 0o600 });
  return { ok: true };
});
ipcMain.handle('ai:set-model', (_e, m) => { if (MODELS.some(x => x.id === m)) settings.set('assistantModel', m); return model(); });

const running = new Map(); // request id -> AbortController
ipcMain.handle('ai:cancel', (_e, id) => { const c = running.get(id); if (c) c.abort(); return true; });
ipcMain.handle('ai:ask', async (e, req) => {
  const c = getClient();
  if (!c) return { error: 'Add your Anthropic API key in the Assistant settings first.' };
  const id = String(req && req.id || '');
  const messages = cleanMessages(req && req.messages);
  if (!messages) return { error: 'Nothing to ask' };
  const win = BrowserWindow.fromWebContents(e.sender);
  const send = (ch, d) => { if (win && !win.isDestroyed()) e.sender.send(ch, d); };
  const abort = new AbortController();
  running.set(id, abort);
  try {
    return await ask(c, { model: model(), system: req.system, messages, effort: req.effort }, delta => send('ai:delta', { id, text: delta }), abort.signal);
  } finally {
    running.delete(id);
  }
});

module.exports = { MODELS };
