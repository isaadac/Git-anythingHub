'use strict';
// Every icon the interface asks for is in js/icons.js (rebuild with node scripts/build-icons.js).
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { usedNames } = require('../scripts/build-icons');

test('icons.js has every icon the pages and scripts use', () => {
  const ctx = { window: {}, document: { readyState: 'complete', querySelectorAll: () => [], documentElement: {} }, MutationObserver: class { observe() {} } };
  vm.runInNewContext(fs.readFileSync(path.join(__dirname, '../src/renderer/js/icons.js'), 'utf8'), ctx);
  const have = new Set(ctx.window.Icons.names());
  const missing = usedNames().filter(n => !have.has(n));
  assert.deepStrictEqual(missing, [], 'run node scripts/build-icons.js');
});

test('no ribbon button is left with an emoji for an icon', () => {
  for (const app of ['write', 'sheets', 'slides', 'notes']) {
    const html = fs.readFileSync(path.join(__dirname, `../src/renderer/${app}.html`), 'utf8');
    const ribbon = html.slice(html.indexOf('<div class="ribbon"'), html.indexOf('<div class="backstage'));
    const emoji = [...ribbon.matchAll(/<(?:button|span)[^>]*>([^<]*[\u{1F300}-\u{1FAFF}][^<]*)</gu)].map(m => m[1].trim());
    assert.deepStrictEqual(emoji, [], app);
  }
});
