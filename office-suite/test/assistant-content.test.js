'use strict';
const test = require('node:test');
const assert = require('node:assert');
const { cleanMessages } = require('../src/main/assistant-core');

const img = (data = 'iVBORw0KGgo=', media_type = 'image/png') => ({ type: 'image', source: { type: 'base64', media_type, data } });
test('text messages pass as before', () => {
  assert.deepStrictEqual(cleanMessages([{ role: 'user', content: 'hi', extra: 1 }]), [{ role: 'user', content: 'hi' }]);
  assert.strictEqual(cleanMessages([{ role: 'assistant', content: 'x' }]), null);
  assert.strictEqual(cleanMessages([]), null);
});
test('a question may carry pictures', () => {
  const m = cleanMessages([{ role: 'user', content: [img(), { type: 'text', text: 'read this', cache_control: { type: 'ephemeral' } }] }]);
  assert.deepStrictEqual(m, [{ role: 'user', content: [img(), { type: 'text', text: 'read this' }] }]);
});
test('only small base64 PNG, JPEG, GIF or WebP pictures, with some text, from the user', () => {
  const ask = content => cleanMessages([{ role: 'user', content }]);
  assert.strictEqual(ask([img(), { type: 'text', text: 'x' }].concat([img(), img(), img(), img()])), null, 'at most four pictures');
  assert.strictEqual(ask([img()]), null, 'some text');
  assert.strictEqual(ask([img('AAAA', 'image/svg+xml'), { type: 'text', text: 'x' }]), null);
  assert.strictEqual(ask([{ type: 'image', source: { type: 'url', url: 'http://127.0.0.1/' } }, { type: 'text', text: 'x' }]), null);
  assert.strictEqual(ask([img('not base64!'), { type: 'text', text: 'x' }]), null);
  assert.strictEqual(ask([img('A'.repeat(3e6 + 4)), { type: 'text', text: 'x' }]), null);
  assert.strictEqual(ask([{ type: 'document', source: {} }, { type: 'text', text: 'x' }]), null);
  assert.strictEqual(ask([{ type: 'tool_result', content: 'x' }]), null);
  // The Assistant's own turns stay text.
  assert.strictEqual(cleanMessages([{ role: 'user', content: 'a' }, { role: 'assistant', content: [img(), { type: 'text', text: 'x' }] }]).length, 1);
});
