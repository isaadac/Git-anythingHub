'use strict';
// Interface languages: every dictionary covers the same words, and those cover the interface.
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const R = path.join(__dirname, '../src/renderer');
const ctx = { window: {} };
for (const l of ['es', 'fr', 'de', 'pt']) vm.runInNewContext(fs.readFileSync(path.join(R, 'i18n', l + '.js'), 'utf8'), ctx);
const D = ctx.window.ISADEK_I18N;
const decode = s => s.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&nbsp;/g, ' ');

// The words on the ribbon: tab names, group names and button labels (a big button's two lines as one).
function ribbonWords(html) {
  const out = new Set();
  const ribbon = html.slice(html.indexOf('<div class="ribbon-tabs"'), html.indexOf('</div>\n</div>', html.indexOf('<div class="ribbon"')) + 1);
  for (const m of ribbon.matchAll(/<button[^>]*>([\s\S]*?)<\/button>|<div class="label">([^<]*)<\/div>/g)) {
    const text = decode((m[1] || m[2] || '').replace(/<span class="ico"[^>]*>[\s\S]*?<\/span>/g, '').replace(/<br>/g, ' ').replace(/<[^>]+>/g, '')).replace(/\s+/g, ' ').trim();
    // Letters and symbols (B, I, ab, Aa) stay as they are.
    if (/\p{L}{3}/u.test(text)) out.add(text);
  }
  return out;
}

test('every dictionary has the same words, each translated', () => {
  const keys = Object.keys(D.es).sort();
  assert.ok(keys.length > 400);
  for (const l of ['fr', 'de', 'pt']) {
    assert.deepStrictEqual(Object.keys(D[l]).sort(), keys, l);
    for (const k of keys) assert.ok(typeof D[l][k] === 'string' && D[l][k].trim(), `${l}: ${k}`);
  }
});

test('the dictionaries cover the ribbons of every app', () => {
  const missing = [];
  for (const app of ['write', 'sheets', 'slides', 'notes', 'pdf']) {
    for (const w of ribbonWords(fs.readFileSync(path.join(R, app + '.html'), 'utf8'))) if (!D.es[w]) missing.push(`${app}: ${w}`);
  }
  assert.deepStrictEqual(missing, []);
});

test('every page loads the language script', () => {
  // form.html is the public page people answer a form on: the form's own words, in the
  // language its author wrote them, and none of the app's.
  for (const f of fs.readdirSync(R).filter(f => f.endsWith('.html') && f !== 'form.html')) assert.match(fs.readFileSync(path.join(R, f), 'utf8'), /<script src="js\/i18n\.js"><\/script>/, f);
});

test('the native menus are translated in every language', () => {
  const { LABELS } = require('../src/main/menu-i18n');
  const keys = Object.keys(LABELS.es).sort();
  for (const l of ['fr', 'de', 'pt']) assert.deepStrictEqual(Object.keys(LABELS[l]).sort(), keys, l);
});
