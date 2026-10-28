'use strict';
// What the desktop app (assistant.js) and the web server (src/server/web-server.js) share about
// asking Claude: the models offered, which messages are accepted, the request, and error text.
const Anthropic = require('@anthropic-ai/sdk');

const MODELS = [
  { id: 'claude-opus-5-5', label: 'Claude Opus 5.5 (best)' },
  { id: 'claude-sonnet-5-5', label: 'Claude Sonnet 5.5 (faster)' },
  { id: 'claude-haiku-5-5', label: 'Claude Haiku 5.5 (fastest, cheapest)' }
];
const validModel = m => (MODELS.some(x => x.id === m) ? m : MODELS[0].id);

/** The conversation from a page, or null if there is nothing to ask. */
function cleanMessages(list) {
  const messages = (Array.isArray(list) ? list : [])
    .filter(m => m && (m.role === 'user' || m.role === 'assistant') && typeof m.content === 'string' && m.content)
    .slice(-40);
  return messages.length && messages[0].role === 'user' ? messages : null;
}

/**
 * Streams a reply. onText(delta) gets the text as it arrives; resolves to
 * { text, stopReason, model, truncated } or { error, refusal? } or { cancelled: true }.
 * signal: an AbortSignal that cancels the request.
 */
async function ask(client, { model, system, messages, effort }, onText, signal) {
  try {
    const stream = client.beta.messages.stream({
      model: validModel(model),
      max_tokens: 32000,
      system: String(system || '').slice(0, 400000),
      thinking: { type: 'adaptive' },
      output_config: { effort: effort === 'high' ? 'high' : 'medium' },
      // If a safety classifier declines, the API retries on Anthropic's recommended model.
      betas: ['server-side-fallback-2026-07-01'],
      fallbacks: 'default',
      messages
    }, signal ? { signal } : undefined);
    stream.on('text', onText);
    const msg = await stream.finalMessage();
    if (msg.stop_reason === 'refusal') return { error: 'Claude declined to help with this request.', refusal: true };
    const text = msg.content.filter(b => b.type === 'text').map(b => b.text).join('');
    return { text, stopReason: msg.stop_reason, model: msg.model, truncated: msg.stop_reason === 'max_tokens' };
  } catch (err) {
    if (err instanceof Anthropic.APIUserAbortError || (signal && signal.aborted)) return { cancelled: true };
    return { error: errorText(err) };
  }
}

function errorText(err) {
  if (err instanceof Anthropic.AuthenticationError) return 'The API key was not accepted. Check it in the Assistant settings.';
  if (err instanceof Anthropic.PermissionDeniedError) return 'This API key cannot use that model.';
  if (err instanceof Anthropic.RateLimitError) return 'Too many requests right now. Try again in a moment.';
  if (err instanceof Anthropic.BadRequestError) return 'Request rejected: ' + err.message;
  if (err instanceof Anthropic.APIConnectionError) return 'Could not reach the Anthropic API. Check your internet connection.';
  if (err instanceof Anthropic.APIError) return `The API returned an error (${err.status}): ${err.message}`;
  return String((err && err.message) || err);
}

module.exports = { Anthropic, MODELS, validModel, cleanMessages, ask };
