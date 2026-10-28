'use strict';
const test = require('node:test');
const assert = require('node:assert');
const { extend, series } = require('../src/renderer/js/sheets-fill.js');
const src = (...xs) => xs.map(raw => ({ raw }));

test('numbers: one is copied, two or more continue the trend', () => {
  assert.deepStrictEqual(extend(src('5'), 3), ['5', '5', '5']);
  assert.deepStrictEqual(extend(src('5'), 3, { increment: true }), ['6', '7', '8']);
  assert.deepStrictEqual(extend(src('1', '3'), 3), ['5', '7', '9']);
  assert.deepStrictEqual(extend(src('0.1', '0.2'), 2), ['0.3', '0.4']);
  assert.deepStrictEqual(extend(src('1', '2', '4'), 1), ['5.33333333333333']);
  assert.deepStrictEqual(extend(src('10', '20'), 2, { reverse: true }), ['0', '-10']);
  assert.deepStrictEqual(extend(src('1', '3'), 2, { copy: true }), ['1', '3']);
});

test('a single date counts up by a day', () => {
  assert.deepStrictEqual(extend([{ raw: '45292', isDate: true }], 2), ['45293', '45294']);
});

test('text with a number counts on, keeping zero padding', () => {
  assert.deepStrictEqual(extend(src('Item 1'), 3), ['Item 2', 'Item 3', 'Item 4']);
  assert.deepStrictEqual(extend(src('Week 07', 'Week 09'), 2), ['Week 11', 'Week 13']);
  assert.deepStrictEqual(extend(src('Q1'), 2, { reverse: true }), ['Q0', 'Q-1']);
});

test('day and month names cycle with their style', () => {
  assert.deepStrictEqual(extend(src('Mon'), 3), ['Tue', 'Wed', 'Thu']);
  assert.deepStrictEqual(extend(src('Friday'), 3), ['Saturday', 'Sunday', 'Monday']);
  assert.deepStrictEqual(extend(src('JAN', 'MAR'), 3), ['MAY', 'JUL', 'SEP']);
  assert.deepStrictEqual(extend(src('december'), 2), ['january', 'february']);
  assert.deepStrictEqual(extend(src('March'), 2, { reverse: true }), ['February', 'January']);
});

test('other text and formulas repeat', () => {
  assert.deepStrictEqual(extend(src('a', 'b'), 3), ['a', 'b', 'a']);
  assert.deepStrictEqual(extend(src('=A1', 'x'), 3), [{ formula: '=A1', from: 0 }, 'x', { formula: '=A1', from: 0 }]);
});

test('Series dialog: linear, growth and stop value', () => {
  assert.deepStrictEqual(series(1, 5, { step: 2 }), [1, 3, 5, 7, 9]);
  assert.deepStrictEqual(series(2, 10, { type: 'growth', step: 3, stop: 100 }), [2, 6, 18, 54]);
});
