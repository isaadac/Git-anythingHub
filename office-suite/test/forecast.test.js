'use strict';
// Forecasting and trend functions: FORECAST.ETS (+ CONFINT, SEASONALITY), TREND, GROWTH, LINEST,
// LOGEST, and the trend fits charts use. Regression results are checked against numpy.polyfit.
const test = require('node:test');
const assert = require('node:assert');
const F = require('../src/renderer/js/formula.js');

function sheet(cols) {
  const cells = {};
  Object.entries(cols).forEach(([col, list]) => list.forEach((v, i) => { cells[col + (i + 1)] = String(v); }));
  return cells;
}
function calc(formula, cells, ref = 'Z99') {
  const all = { ...cells, [ref]: formula };
  const e = new F.Engine({ getRaw: (s, r, c) => all[F.addr(r, c)] ?? '', sheetNames: () => ['Sheet1'] });
  const p = F.parseRef(ref);
  const v = e.value('Sheet1', p.row, p.col);
  return v instanceof F.FormulaError ? v.error : v;
}
const near = (a, b, tol = 1e-6) => assert.ok(Math.abs(a - b) <= tol * Math.max(1, Math.abs(b)), `${a} ≉ ${b}`);
const X = [1, 2, 3, 4, 5, 6, 7, 8], Y = [2.1, 3.9, 6.2, 7.8, 10.1, 12.2, 13.8, 16.1];

test('LINEST and TREND match a least-squares line (numpy.polyfit)', () => {
  const c = sheet({ A: X, B: Y });
  near(calc('=INDEX(LINEST(B1:B8,A1:A8),1,1)', c), 1.99761905);
  near(calc('=INDEX(LINEST(B1:B8,A1:A8),1,2)', c), 0.03571429, 1e-5);
  near(calc('=INDEX(TREND(B1:B8,A1:A8,10),1,1)', c), 1.99761905 * 10 + 0.03571429, 1e-6);
  near(calc('=INDEX(LINEST(B1:B8,A1:A8,TRUE,TRUE),3,1)', c), calc('=RSQ(B1:B8,A1:A8)', c));
});

test('GROWTH and LOGEST fit an exponential curve', () => {
  const c = sheet({ A: X, B: [3, 5.5, 9.1, 15, 24.5, 40.3, 66.1, 108.9] });
  near(calc('=INDEX(LOGEST(B1:B8,A1:A8),1,1)', c), 1.6591505137411398, 1e-9);
  near(calc('=INDEX(LOGEST(B1:B8,A1:A8),1,2)', c), 1.9322067582939688, 1e-9);
  near(calc('=GROWTH(B1:B8,A1:A8,9)', c), 1.9322067582939688 * 1.6591505137411398 ** 9, 1e-9);
  assert.strictEqual(calc('=GROWTH(A1:A3,B1:B3,4)', sheet({ A: [1, -1, 2], B: [1, 2, 3] })), '#NUM!');
});

test('trend fits for charts: polynomial matches numpy, R² is reported', () => {
  const poly = F.trendFit(X, Y, 'poly', 2);
  near(poly.f(10), 0.00238095 * 100 + 1.97619048 * 10 + 0.07142857, 1e-6);
  assert.ok(poly.r2 > 0.99 && poly.r2 <= 1);
  assert.match(F.trendFit(X, Y).eq, /^y = 1\.998x \+ 0\.03571$/);
  assert.strictEqual(F.trendFit([1], [2]), null);
});

test('FORECAST.ETS continues a trend with a season', () => {
  // 6 years of quarters: a rising trend and the same quarterly pattern each year.
  const t = Array.from({ length: 24 }, (_, i) => i + 1);
  const y = t.map(i => 100 + 2 * i + [12, -6, 0, 6][(i - 1) % 4]);
  const c = sheet({ A: t, B: y });
  assert.strictEqual(calc('=FORECAST.ETS.SEASONALITY(B1:B24,A1:A24)', c), 4);
  for (const h of [25, 26, 27, 28, 30]) near(calc(`=FORECAST.ETS(${h},B1:B24,A1:A24)`, c), 100 + 2 * h + [12, -6, 0, 6][(h - 1) % 4], 0.0005);
  const ci = calc('=FORECAST.ETS.CONFINT(26,B1:B24,A1:A24)', c);
  assert.ok(typeof ci === 'number' && ci >= 0 && ci < 5, String(ci));
  // No season asked for: a straight trend.
  near(calc('=FORECAST.ETS(25,B1:B24,A1:A24,0)', c), calc('=FORECAST.ETS(25,B1:B24,A1:A24,0)', c));
});

test('FORECAST.ETS fills gaps, averages equal dates and refuses bad timelines', () => {
  const c = sheet({ A: [1, 2, 4, 5, 5, 6, 7, 8], B: [10, 12, 16, 17, 19, 20, 22, 24] });
  const v = calc('=FORECAST.ETS(9,B1:B8,A1:A8,0)', c);
  assert.ok(v > 24 && v < 28, String(v));
  assert.strictEqual(calc('=FORECAST.ETS(9,B1:B3,A1:A2)', c), '#N/A');
  assert.strictEqual(calc('=FORECAST.ETS(9,B1:B3,C1:C3)', { ...c, C1: 'a', C2: 'b', C3: 'c' }), '#VALUE!');
  assert.strictEqual(calc('=FORECAST.ETS(0,B1:B8,A1:A8)', c), '#NUM!');
  assert.strictEqual(calc('=FORECAST.ETS(9,B1:B8,A1:A8,-1)', c), '#NUM!');
});

test('new forecast functions are saved with Excel’s _xlfn. prefix', () => {
  assert.match(F.toExcelFormula('=FORECAST.ETS(A1,B1:B9,C1:C9)'), /_xlfn\.FORECAST\.ETS\(/);
  assert.match(F.toExcelFormula('=FORECAST.ETS.CONFINT(A1,B1:B9,C1:C9)'), /_xlfn\.FORECAST\.ETS\.CONFINT\(/);
  assert.match(F.toExcelFormula('=TREND(B1:B9)'), /^=?TREND\(/);
});
