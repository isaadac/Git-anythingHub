'use strict';
// Atlas Assistant: Claude, with the user's own Anthropic API key. Requests are made here in the
// main process with the official SDK, so the key never reaches a page; replies stream to the
// window that asked.
const { app, ipcMain, safeStorage, BrowserWindow } = require('electron');
const fs = require('fs');
const path = require('path');
const Anthropic = require('@anthropic-ai/sdk');
const settings = require('./settings');

const MODELS = [
  { id: 'claude-opus-5-5', label: 'Claude Opus 5.5 (best)' },
  { id: 'claude-sonnet-5-5', label: 'Claude Sonnet 5.5 (faster)' },
  { id: 'claude-haiku-5-5', label: 'Claude Haiku 5.5 (fastest, cheapest)' }
];
const keyFile = () => path.join(app.getPath('userData'), 'assistant-key.bin');
function readKey() {
  try {
    const raw = fs.readFileSync(keyFile());
    return safeStorage.isEncryptionAvailable() && raw[0] !== 0x73 /* 's' of sk- */ ? safeStorage.decryptString(raw) : raw.toString('utf8');
  } catch { return process.env.ANTHROPIC_API_KEY || ''; }
}
function model() {
  const m = settings.get('assistantModel');
  return MODELS.some(x => x.id === m) ? m : MODELS[0].id;
}
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

const running = new Map(); // request id -> stream
ipcMain.handle('ai:cancel', (_e, id) => { const s = running.get(id); if (s) s.abort(); return true; });
ipcMain.handle('ai:ask', async (e, req) => {
  const c = getClient();
  if (!c) return { error: 'Add your Anthropic API key in the Assistant settings first.' };
  const id = String(req && req.id || '');
  const messages = (Array.isArray(req.messages) ? req.messages : [])
    .filter(m => m && (m.role === 'user' || m.role === 'assistant') && typeof m.content === 'string' && m.content)
    .slice(-40);
  if (!messages.length || messages[0].role !== 'user') return { error: 'Nothing to ask' };
  const win = BrowserWindow.fromWebContents(e.sender);
  const send = (ch, d) => { if (win && !win.isDestroyed()) e.sender.send(ch, d); };
  try {
    const stream = c.beta.messages.stream({
      model: model(),
      max_tokens: 32000,
      system: String(req.system || '').slice(0, 400000),
      thinking: { type: 'adaptive' },
      output_config: { effort: req.effort === 'high' ? 'high' : 'medium' },
      // If a safety classifier declines, the API retries on Anthropic's recommended model.
      betas: ['server-side-fallback-2026-07-01'],
      fallbacks: 'default',
      messages
    });
    running.set(id, stream);
    stream.on('text', delta => send('ai:delta', { id, text: delta }));
    const msg = await stream.finalMessage();
    if (msg.stop_reason === 'refusal') return { error: 'Claude declined to help with this request.', refusal: true };
    const text = msg.content.filter(b => b.type === 'text').map(b => b.text).join('');
    return { text, stopReason: msg.stop_reason, model: msg.model, truncated: msg.stop_reason === 'max_tokens' };
  } catch (err) {
    if (err instanceof Anthropic.APIUserAbortError) return { cancelled: true };
    if (err instanceof Anthropic.AuthenticationError) return { error: 'The API key was not accepted. Check it in the Assistant settings.' };
    if (err instanceof Anthropic.PermissionDeniedError) return { error: 'This API key cannot use that model.' };
    if (err instanceof Anthropic.RateLimitError) return { error: 'Too many requests right now. Try again in a moment.' };
    if (err instanceof Anthropic.BadRequestError) return { error: 'Request rejected: ' + err.message };
    if (err instanceof Anthropic.APIConnectionError) return { error: 'Could not reach the Anthropic API. Check your internet connection.' };
    if (err instanceof Anthropic.APIError) return { error: `The API returned an error (${err.status}): ${err.message}` };
    return { error: String(err.message || err) };
  } finally {
    running.delete(id);
  }
});

module.exports = { MODELS };
