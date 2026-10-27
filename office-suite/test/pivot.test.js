'use strict';
const test = require('node:test');
const assert = require('node:assert');
const P = require('../src/renderer/js/sheets-pivot.js');

const table = {
  headers: ['Region', 'Rep', 'Year', 'Sales', 'Units'],
  rows: [
    ['East', 'Ann', 2023, 100, 1],
    ['East', 'Bob', 2023, 50, 2],
    ['West', 'Cy', 2024, 70, 3],
    ['East', 'Ann', 2024, 30, 4],
    ['West', 'Cy', 2023, 'n/a', 5],
    ['', '', '', '', ''],
    ['North', '', 2024, 10, 6]
  ]
};
const grid = r => r.cells.map(row => row.map(c => c.v));

test('rows and a sum', () => {
  const r = P.compute(table, { rows: ['Region'], vals: [{ field: 'Sales', agg: 'sum' }] });
  assert.deepStrictEqual(grid(r), [['Region', 'Sum of Sales'], ['East', 180], ['North', 10], ['West', 70], ['Grand Total', 260]]);
  assert.strictEqual(r.cells[0][0].k, 'hdr');
  assert.strictEqual(r.cells[4][1].k, 'total');
});

test('count counts non-empty values, average and max skip text', () => {
  const r = P.compute(table, { rows: ['Region'], vals: [{ field: 'Sales', agg: 'count' }, { field: 'Sales', agg: 'average' }, { field: 'Sales', agg: 'max' }] });
  assert.deepStrictEqual(grid(r)[3], ['West', 2, 70, 70]);
  assert.deepStrictEqual(grid(r)[0], ['Region', 'Count of Sales', 'Average of Sales', 'Max of Sales']);
});

test('columns with a grand total column', () => {
  const r = P.compute(table, { rows: ['Region'], cols: ['Year'], vals: [{ field: 'Units', agg: 'sum' }] });
  assert.deepStrictEqual(grid(r), [
    ['Sum of Units', 'Column Labels', '', ''],
    ['Region', '2023', '2024', 'Grand Total'],
    ['East', 3, 4, 7],
    ['North', '', 6, 6],
    ['West', 5, 3, 8],
    ['Grand Total', 8, 13, 21]
  ]);
});

test('two row fields: outer labels once, subtotals, (blank) last', () => {
  const r = P.compute(table, { rows: ['Region', 'Rep'], vals: [{ field: 'Units', agg: 'sum' }] });
  assert.deepStrictEqual(grid(r), [
    ['Region', 'Rep', 'Sum of Units'],
    ['East', 'Ann', 5], ['', 'Bob', 2], ['East Total', '', 7],
    ['North', '(blank)', 6], ['North Total', '', 6],
    ['West', 'Cy', 8], ['West Total', '', 8],
    ['Grand Total', '', 21]
  ]);
});

test('filters keep only matching records', () => {
  const r = P.compute(table, { rows: ['Rep'], vals: [{ field: 'Sales', agg: 'sum' }], filters: [{ field: 'Year', value: 2023 }] });
  assert.deepStrictEqual(grid(r), [['Rep', 'Sum of Sales'], ['Ann', 100], ['Bob', 50], ['Cy', 0], ['Grand Total', 150]]);
});

test('several values across columns', () => {
  const r = P.compute(table, { cols: ['Year'], vals: [{ field: 'Sales', agg: 'sum' }, { field: 'Units', agg: 'min' }] });
  assert.deepStrictEqual(grid(r)[0], ['', '2023 · Sum of Sales', '2023 · Min of Units', '2024 · Sum of Sales', '2024 · Min of Units', 'Total Sum of Sales', 'Total Min of Units']);
  assert.deepStrictEqual(grid(r).slice(1), [['Grand Total', 150, 1, 110, 3, 260, 1]]);
});

test('missing fields are reported, not thrown', () => {
  const r = P.compute(table, { rows: ['Nope'], vals: [] });
  assert.ok(r.error);
});

test('clean drops bad definitions and fields', () => {
  assert.strictEqual(P.clean({ src: { sheet: 'S', ref: 'A1;B2' }, r: 0, c: 0 }), null);
  const c = P.clean({ src: { sheet: 'S', ref: 'A1:E8' }, r: 2, c: 3, rows: ['Region', 5, { x: 1 }], vals: [{ field: 'Sales', agg: 'eval' }, { field: 'Units', agg: 'sum' }], id: '<x>' });
  assert.deepStrictEqual(c.rows, ['Region']);
  assert.deepStrictEqual(c.vals, [{ field: 'Units', agg: 'sum' }]);
  assert.match(c.id, /^pv[a-z0-9]+$/);
});
