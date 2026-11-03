'use strict';
// Clean & Transform steps (src/renderer/js/sheets-clean.js).
const test = require('node:test');
const assert = require('node:assert');
const { run, describe, cleanStep } = require('../src/renderer/js/sheets-clean.js');

const messy = [
  ['Name', 'City', 'Sales', 'Date'],
  ['  ada   lovelace ', 'london', '1,200', '2024-01-05'],
  ['BOB SMITH', ' Paris', '300', '2024-02-10'],
  ['ada lovelace', 'London ', '1200', '2024-01-05'],
  ['', '', '', ''],
  ['carol​ jones', 'berlin', 'n/a', '']
];

test('trim, case and type clean up text and numbers', () => {
  const out = run(messy, [{ op: 'trim', cols: ['Name', 'City'] }, { op: 'clean', cols: ['Name'] }, { op: 'case', cols: ['Name', 'City'], to: 'proper' }, { op: 'type', cols: ['Sales'], to: 'number' }]);
  assert.deepStrictEqual(out[1], ['Ada Lovelace', 'London', 1200, '2024-01-05']);
  assert.deepStrictEqual(out[2], ['Bob Smith', 'Paris', 300, '2024-02-10']);
  assert.strictEqual(out[5][2], 'n/a');
});

test('blank rows and duplicates go', () => {
  const out = run(messy, [{ op: 'trim', cols: ['Name', 'City'] }, { op: 'case', cols: ['Name', 'City'], to: 'proper' }, { op: 'type', cols: ['Sales'], to: 'number' }, { op: 'removeBlankRows' }, { op: 'dedupe' }]);
  assert.strictEqual(out.length, 4);
  assert.deepStrictEqual(out.map(r => r[0]), ['Name', 'Ada Lovelace', 'Bob Smith', 'Carol​ Jones']);
});

test('split and merge columns', () => {
  const split = run([['Full'], ['Ada Lovelace'], ['Bob Smith']], [{ op: 'split', col: 'Full', by: ' ' }]);
  assert.deepStrictEqual(split, [['Full.1', 'Full.2'], ['Ada', 'Lovelace'], ['Bob', 'Smith']]);
  const merged = run(split, [{ op: 'merge', cols: ['Full.2', 'Full.1'], sep: ', ', name: 'Name' }]);
  assert.deepStrictEqual(merged, [['Name'], ['Lovelace, Ada'], ['Smith, Bob']]);
});

test('fill down, filter, sort, top', () => {
  const t = [['Region', 'Rep', 'Amount'], ['North', 'A', 5], ['', 'B', 9], ['South', 'C', 2], ['', 'D', 7]];
  const out = run(t, [{ op: 'fillDown', cols: ['Region'] }, { op: 'filter', col: 'Amount', cond: 'greater than', value: '4' }, { op: 'sort', col: 'Amount', desc: true }, { op: 'top', n: 2 }]);
  assert.deepStrictEqual(out, [['Region', 'Rep', 'Amount'], ['North', 'B', 9], ['South', 'D', 7]]);
});

test('group by with a sum, and unpivot', () => {
  const t = [['Region', 'Q1', 'Q2'], ['North', 10, 20], ['South', 5, ''], ['North', 1, 2]];
  assert.deepStrictEqual(run(t, [{ op: 'group', by: ['Region'], col: 'Q1', agg: 'sum' }]), [['Region', 'Sum of Q1'], ['North', 11], ['South', 5]]);
  assert.deepStrictEqual(run(t.slice(0, 3), [{ op: 'unpivot', cols: ['Q1', 'Q2'] }]), [['Region', 'Attribute', 'Value'], ['North', 'Q1', 10], ['North', 'Q2', 20], ['South', 'Q1', 5]]);
});

test('replace, rename, remove columns', () => {
  const out = run([['A', 'B', 'C'], ['x-y', 1, 2]], [{ op: 'replace', cols: ['A'], find: '-', with: '+' }, { op: 'rename', col: 'B', to: 'Bee' }, { op: 'remove', cols: ['C'] }]);
  assert.deepStrictEqual(out, [['A', 'Bee'], ['x+y', 1]]);
});

test('a step that no longer fits is skipped; steps describe themselves; files are cleaned', () => {
  assert.deepStrictEqual(run([['A'], [1]], [{ op: 'split', col: 'Gone', by: ',' }, { op: 'nope' }]), [['A'], [1]]);
  assert.strictEqual(describe({ op: 'filter', col: 'Amount', cond: 'greater than', value: '4' }), 'Kept rows where Amount greater than “4”');
  assert.strictEqual(cleanStep({ op: 'evil', code: 'x' }), null);
  assert.deepStrictEqual(cleanStep({ op: 'trim', cols: ['A', { x: 1 }], extra: 1 }), { op: 'trim', cols: ['A'] });
});
