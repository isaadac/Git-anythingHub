'use strict';
// Three-way merge used by live editing (src/renderer/js/merge3.js).
const test = require('node:test');
const assert = require('node:assert');
const { merge3, mergeText } = require('../src/renderer/js/merge3.js');

test('two people typing in different places of one paragraph both keep their words', () => {
  assert.strictEqual(mergeText('The cat sat.', 'The black cat sat.', 'The cat sat down.'), 'The black cat sat down.');
  assert.strictEqual(mergeText('<p>One two three</p>', '<p>Zero one two three</p>', '<p>One two three four</p>'), '<p>Zero one two three four</p>');
});

test('formatting and words merge independently', () => {
  assert.strictEqual(mergeText('Hello world', '<b>Hello</b> world', 'Hello brave world'), '<b>Hello</b> brave world');
});

test('typing at the very same spot: the later change', () => {
  assert.strictEqual(mergeText('A B', 'A x B', 'A y B'), 'A x B');
  assert.strictEqual(mergeText('A B', 'A x B', 'A y B', true), 'A y B');
});

test('changing the very same words: ours', () => {
  assert.strictEqual(mergeText('It costs 50 dollars', 'It costs 60 dollars', 'It costs 70 dollars'), 'It costs 60 dollars');
});

test('a deletion and an edit elsewhere both happen', () => {
  assert.strictEqual(mergeText('one two three four', 'one three four', 'one two three four five'), 'one three four five');
});

test('one side unchanged: the other side wins whole', () => {
  assert.strictEqual(mergeText('a', 'a', 'b'), 'b');
  assert.strictEqual(mergeText('a', 'b', 'a'), 'b');
});

test('objects merge field by field (a paragraph moved by one, edited by the other)', () => {
  const base = { p: 'a1', h: '<p>Text here</p>' };
  assert.deepStrictEqual(merge3(base, { p: 'a5', h: '<p>Text here</p>' }, { p: 'a1', h: '<p>Text here now</p>' }, ['h']), { p: 'a5', h: '<p>Text here now</p>' });
  assert.deepStrictEqual(merge3(base, { p: 'a1', h: '<p>Mine here</p>' }, { p: 'a1', h: '<p>Text here now</p>' }, ['h']), { p: 'a1', h: '<p>Mine here now</p>' });
});

test('only the fields named as text merge as text; other strings: ours', () => {
  assert.deepStrictEqual(merge3({ v: 'x' }, { v: 'by 0' }, { v: 'by 1' }, ['h']), { v: 'by 0' });
  assert.strictEqual(merge3('a b', 'a b c', 'z a b'), 'a b c');
});

test('a spreadsheet cell changed by both keeps ours, its style change by the other is kept', () => {
  assert.deepStrictEqual(merge3({ v: '50' }, { v: '100' }, { v: '50', s: { b: true } }), { v: '100', s: { b: true } });
  assert.deepStrictEqual(merge3({ v: '50' }, { v: '100' }, { v: '200' }), { v: '100' });
});

test('merging is symmetric when the edits do not overlap', () => {
  const base = 'alpha beta gamma delta', a = 'ALPHA beta gamma delta', b = 'alpha beta gamma DELTA';
  assert.strictEqual(mergeText(base, a, b), mergeText(base, b, a));
});

test('tags are never split', () => {
  const out = mergeText('<p>a <i>b</i> c</p>', '<p>a <i>b</i> c d</p>', '<p>z a <i>b</i> c</p>');
  assert.strictEqual(out, '<p>z a <i>b</i> c d</p>');
});

test('both sides work out the same merge (who was first is all that matters)', () => {
  const base = '<p>The meeting is on Monday.</p>', a = '<p>Reminder: The meeting is on Monday.</p>', b = '<p>The meeting is on Monday. Bring the slides.</p>';
  assert.strictEqual(mergeText(base, a, b, true), mergeText(base, b, a, false));
  assert.strictEqual(mergeText(base, a, b, true), '<p>Reminder: The meeting is on Monday. Bring the slides.</p>');
  assert.strictEqual(mergeText('A B', 'A x B', 'A y B', true), mergeText('A B', 'A y B', 'A x B', false));
  assert.strictEqual(mergeText('cost 5', 'cost 6', 'cost 7', true), 'cost 7');
  assert.deepStrictEqual(merge3({ v: '1' }, { v: '2' }, { v: '3' }, [], '', true), { v: '3' });
});

test('the same words typed on both sides are not doubled', () => {
  assert.strictEqual(mergeText('A B', 'A xy B', 'A xy B', true), 'A xy B');
  assert.ok(!/xyxy|xyzxy|xyxyz/.test(mergeText('A B', 'A xyz B', 'A xy B', true)));
});
