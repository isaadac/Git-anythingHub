'use strict';
// Workbook names and table (structured) references in the formula engine; Goal Seek and Solver.
const test = require('node:test');
const assert = require('node:assert');
const F = require('../src/renderer/js/formula.js');
const W = require('../src/renderer/js/sheets-whatif.js');

function engine(cells, opts = {}) {
  return new F.Engine({
    getRaw: (s, r, c) => cells[F.addr(r, c)] ?? '',
    sheetNames: () => ['Sheet1'],
    resolveName: n => (opts.names || {})[n] ?? null,
    tableRef: opts.tableRef
  });
}

test('defined names: ranges, constants and formulas', () => {
  const e = engine({ A1: '1', A2: '2', A3: '3', B1: '=SUM(Data)*Rate', B2: '=Twice', B3: '=Nope+1' }, { names: { DATA: 'Sheet1!$A$1:$A$3', RATE: '0.5', TWICE: '=A3*2' } });
  assert.strictEqual(e.value('Sheet1', 0, 1), 3);
  assert.strictEqual(e.value('Sheet1', 1, 1), 6);
  assert.strictEqual(e.value('Sheet1', 2, 1).error, '#NAME?');
});

test('table references resolve through the workbook', () => {
  // Table1 = A1:B4 (header row 1), columns Item, Qty.
  const cells = { A1: 'Item', B1: 'Qty', A2: 'x', B2: '2', A3: 'y', B3: '3', A4: 'z', B4: '5', C2: '=[@Qty]*10', D1: '=SUM(Table1[Qty])', D2: '=ROWS(Table1[#All])' };
  const tableRef = (spec, sheet, cur) => {
    const cols = ['item', 'qty'];
    let c1 = 0, c2 = 1;
    if (spec.cols) { c1 = cols.indexOf(spec.cols[0].toLowerCase()); c2 = cols.indexOf(spec.cols[1].toLowerCase()); }
    if (spec.thisRow) return { sheet, r1: cur.row, c1, r2: cur.row, c2 };
    if (spec.parts.has('#all')) return { sheet, r1: 0, c1, r2: 3, c2 };
    return { sheet, r1: 1, c1, r2: 3, c2 };
  };
  const e = engine(cells, { tableRef });
  assert.strictEqual(e.value('Sheet1', 0, 3), 10);
  assert.strictEqual(e.value('Sheet1', 1, 2), 20);
  assert.strictEqual(e.value('Sheet1', 1, 3), 4);
});

test('structured references survive copying, inserting rows and Excel conversion', () => {
  assert.strictEqual(F.shiftFormula('=SUM(T[Q1])+A1', 1, 0), '=SUM(T[Q1])+A2');
  assert.strictEqual(F.adjustForInsert('=T[[#Headers],[B2]]+B2', 'row', 0, 1, 'S', 'S'), '=T[[#Headers],[B2]]+B3');
  assert.strictEqual(F.toExcelFormula('[@Qty]*Rate', { names: new Set(['RATE']), table: 'Orders' }), 'Orders[[#This Row],[Qty]]*Rate');
  assert.strictEqual(F.toExcelFormula('LET(x,1,x+Rate)', { names: new Set(['RATE']) }), '_xlfn.LET(_xlpm.x,1,_xlpm.x+Rate)');
  assert.strictEqual(F.fromExcelFormula('Orders[[#This Row],[Unit Price]]'), 'Orders[@[Unit Price]]');
});

test('Goal Seek and Solver', () => {
  assert.ok(Math.abs(W.goalSeek(x => x * x, 1, 9).x - 3) < 1e-6);
  const r = W.solve({ x0: [0, 0], evaluate: ([x, y]) => ({ obj: 3 * x + 5 * y, cons: [{ lhs: x, op: '<=', rhs: 4 }, { lhs: 2 * y, op: '<=', rhs: 12 }, { lhs: 3 * x + 2 * y, op: '<=', rhs: 18 }] }), goal: 'max' });
  assert.deepStrictEqual([r.ok, r.x, r.obj], [true, [2, 6], 36]);
});
