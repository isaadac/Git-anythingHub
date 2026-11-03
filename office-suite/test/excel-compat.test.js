'use strict';
// Formula behaviour found by opening real Excel files (e2e/corpus): what Excel does, Isadek does.
const test = require('node:test');
const assert = require('node:assert');
const F = require('../src/renderer/js/formula.js');

function book(sheets, legacy = new Set()) {
  return new F.Engine({
    getRaw: (s, r, c) => { const v = (sheets[s] || {})[F.addr(r, c)]; return v === undefined ? '' : String(v); },
    sheetNames: () => Object.keys(sheets),
    isLegacy: (s, r, c) => legacy.has(s + '!' + F.addr(r, c))
  });
}
const val = (e, s, ref) => { const p = F.parseRef(ref); const v = e.value(s, p.row, p.col); return v instanceof F.FormulaError ? v.error : v; };

test('3D references, TRUE() and error constants', () => {
  const e = book({ S1: { C1: 5, A1: '=SUM(S1:S3!C1)', A2: '=TRUE()', A3: "='S2'!#REF!", A4: '=IFERROR(#DIV/0!,7)' }, S2: { C1: 7 }, S3: { C1: 'x' } });
  assert.strictEqual(val(e, 'S1', 'A1'), 12);
  assert.strictEqual(val(e, 'S1', 'A2'), true);
  assert.strictEqual(val(e, 'S1', 'A3'), '#REF!');
  assert.strictEqual(val(e, 'S1', 'A4'), 7);
});

test('formulas from an older Excel take the value in their own row (implicit intersection)', () => {
  const cells = { A1: 1, A2: 4, A3: 9, B2: '=SQRT(A1:A3)', C2: '=A1:A3*10', D2: '=SQRT(A1:A3)', B5: '=SQRT(A1:A3)' };
  const e = book({ S: cells }, new Set(['S!B2', 'S!C2', 'S!B5']));
  assert.strictEqual(val(e, 'S', 'B2'), 2);           // row 2 → A2
  assert.strictEqual(val(e, 'S', 'C2'), 40);
  assert.strictEqual(val(e, 'S', 'B5'), '#VALUE!');   // no row 5 in A1:A3
  assert.strictEqual(val(e, 'S', 'D2'), 1);           // a new formula spills: D2 shows the first value
});

test('dates as Excel numbers them, including its 29 February 1900', () => {
  const e = book({ S: { A1: '=DATE(1900,1,1)', A2: '=DATE(1900,2,29)', A3: '=DATE(1900,3,1)', A4: '=DAY(60)', A5: '=MONTH(0)', A6: '=DATE(2024,14,0)', A7: '=DAYS360(DATE(2008,1,16),DATE(2008,5,16))', A8: '=YEAR(-1)' } });
  assert.deepStrictEqual(['A1', 'A2', 'A3', 'A4', 'A5', 'A7', 'A8'].map(r => val(e, 'S', r)), [1, 60, 61, 29, 1, 120, '#NUM!']);
  assert.strictEqual(F.formatValue(val(e, 'S', 'A6'), 'yyyy-mm-dd'), '2025-01-31');
});

test('aggregates skip text in referenced cells; logical functions skip text in ranges', () => {
  const e = book({ S: { A1: 2, A2: 'x', A3: true, B1: '=SUM(A1,A2)', B2: '=AND(A1:A3)', B3: '=OR(A2)', B4: '=PRODUCT(A2)', B5: '=SUM(A1,"3")' } });
  assert.deepStrictEqual(['B1', 'B2', 'B3', 'B4', 'B5'].map(r => val(e, 'S', r)), [2, true, '#VALUE!', 0, 5]);
});

test('functions that were missing or differed', () => {
  const e = book({ S: { A1: 1, A2: 2, B1: 3, B2: 'x', C1: '=SUMX2MY2(A1:A2,B1:B2)', C2: '=SUMXMY2(A1:A2,B1:B2)', C3: '=CHOOSE(2,1/0,"ok")', C4: '=CHAR(128)', C5: '=SUBSTITUTE("aaa","a","b",2.9)', C6: '=DOLLAR(-1234.5)', C7: '=ROUND(1.26,1.9)', C8: '=FLOOR(0,0)', C9: '=ERROR.TYPE(#N/A)' } });
  assert.deepStrictEqual(['C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8', 'C9'].map(r => val(e, 'S', r)), [-8, 4, 'ok', '€', 'aba', '($1,234.50)', 1.3, 0, 7]);
});
