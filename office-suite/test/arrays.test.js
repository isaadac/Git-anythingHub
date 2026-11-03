'use strict';
const test = require('node:test');
const assert = require('node:assert');
const F = require('../src/renderer/js/formula.js');

function book(cells) {
  const data = { S: { ...cells } };
  const engine = new F.Engine({
    getRaw: (s, r, c) => (data[s] || {})[F.addr(r, c)] ?? '',
    sheetNames: () => Object.keys(data),
    formulasInRange: (s, r1, c1, r2, c2) => Object.keys(data[s] || {}).filter(k => String(data[s][k])[0] === '=').map(k => { const p = F.parseRef(k); return [p.row, p.col]; }).filter(([r, c]) => r >= r1 && r <= r2 && c >= c1 && c <= c2)
  });
  return { v: ref => { const p = F.parseRef(ref); return engine.value('S', p.row, p.col); }, engine, data };
}
const near = (a, b, eps = 1e-6) => assert.ok(Math.abs(a - b) < eps, `${a} vs ${b}`);

test('SEQUENCE spills and A1# reads the whole spill', () => {
  const b = book({ A1: '=SEQUENCE(3,2,10,5)', D1: '=SUM(A1#)', E1: '=ROWS(A1#)' });
  assert.strictEqual(b.v('A1'), 10);
  assert.strictEqual(b.v('B1'), 15);
  assert.strictEqual(b.v('B3'), 35);
  assert.strictEqual(b.v('D1'), 135);
  assert.strictEqual(b.v('E1'), 3);
});

test('a blocked spill is #SPILL!', () => {
  const b = book({ A1: '=SEQUENCE(3)', A2: 'x' });
  assert.strictEqual(b.v('A1').error, '#SPILL!');
  assert.strictEqual(b.v('A3'), '');
});

test('reading a spilled cell before its formula works', () => {
  const b = book({ C1: '=A3*10', A1: '=SEQUENCE(3)' });
  assert.strictEqual(b.v('C1'), 30);
});

test('FILTER, SORT, UNIQUE', () => {
  const b = book({ A1: 'b', A2: 'a', A3: 'b', A4: 'c', B1: '3', B2: '1', B3: '2', B4: '5',
    D1: '=FILTER(A1:B4,B1:B4>1)', G1: '=SORT(A1:B4,2,-1)', J1: '=UNIQUE(A1:A4)', K1: '=FILTER(A1:A4,B1:B4>10,"none")' });
  assert.deepStrictEqual([b.v('D1'), b.v('E1'), b.v('D2'), b.v('D3'), b.v('E3')], ['b', 3, 'b', 'c', 5]);
  assert.strictEqual(b.v('D4'), '');
  assert.deepStrictEqual([b.v('G1'), b.v('H1'), b.v('G4'), b.v('H4')], ['c', 5, 'a', 1]);
  assert.deepStrictEqual([b.v('J1'), b.v('J2'), b.v('J3'), b.v('J4')], ['b', 'a', 'c', '']);
  assert.strictEqual(b.v('K1'), 'none');
});

test('array arithmetic and lifted functions spill', () => {
  const b = book({ A1: '1', A2: '2', A3: '3', B1: '=A1:A3*2', C1: '=ROUND(A1:A3/3,1)', D1: '=SUMPRODUCT((A1:A3>1)*A1:A3)', E1: '=IF(A1:A3>1,"big","small")' });
  assert.deepStrictEqual([b.v('B1'), b.v('B2'), b.v('B3')], [2, 4, 6]);
  assert.deepStrictEqual([b.v('C1'), b.v('C3')], [0.3, 1]);
  assert.strictEqual(b.v('D1'), 5);
  assert.deepStrictEqual([b.v('E1'), b.v('E2')], ['small', 'big']);
});

test('LET, INDIRECT, OFFSET, ROW, COLUMN, XLOOKUP returning a row', () => {
  const b = book({ A1: '5', A2: '6', A3: '7', B1: 'x', B2: 'y', B3: 'z',
    C1: '=LET(x,A1*2,y,x+1,x*y)', C2: '=INDIRECT("A"&3)', C3: '=SUM(OFFSET(A1,1,0,2,1))', C4: '=ROW()+COLUMN(B1)',
    D1: '=XLOOKUP("y",B1:B3,A1:B3)', D5: '=XMATCH(6,A1:A3)', D6: '=XLOOKUP("q",B1:B3,A1:A3,"missing")' });
  assert.strictEqual(b.v('C1'), 110);
  assert.strictEqual(b.v('C2'), 7);
  assert.strictEqual(b.v('C3'), 13);
  assert.strictEqual(b.v('C4'), 6);
  assert.deepStrictEqual([b.v('D1'), b.v('E1')], [6, 'y']);
  assert.strictEqual(b.v('D5'), 2);
  assert.strictEqual(b.v('D6'), 'missing');
});

test('statistics', () => {
  const b = book({ A1: '1', A2: '2', A3: '3', A4: '4', A5: '10', B1: '2', B2: '4', B3: '5', B4: '4', B5: '5',
    C1: '=CORREL(A1:A5,B1:B5)', C2: '=SLOPE(B1:B5,A1:A5)', C3: '=PERCENTILE.INC(A1:A5,0.3)', C4: '=QUARTILE.EXC(A1:A5,1)',
    C5: '=NORM.S.DIST(1.5,TRUE)', C6: '=NORM.S.INV(0.975)', C7: '=STDEV.S(A1:A5)', C8: '=MEDIAN(A1:A5)', C9: '=RANK.AVG(4,B1:B5)' });
  near(b.v('C1'), 0.6350852961, 1e-9);
  near(b.v('C2'), 0.22, 1e-12);
  near(b.v('C3'), 2.2);
  near(b.v('C4'), 1.5);
  near(b.v('C5'), 0.9331927987, 1e-9);
  near(b.v('C6'), 1.959963985, 1e-8);
  near(b.v('C7'), 3.5355339, 1e-6);
  assert.strictEqual(b.v('C9'), 3.5);
});

test('finance', () => {
  const b = book({ A1: '-1000', A2: '300', A3: '400', A4: '500',
    B1: '=IRR(A1:A4)', B2: '=RATE(36,-300,10000)', B3: '=NPER(0.01,-100,1000)', B4: '=IPMT(0.1/12,1,36,8000)', B5: '=PPMT(0.1/12,1,36,8000)', B6: '=DDB(2400,300,10,1)', B7: '=EFFECT(0.0525,4)' });
  near(b.v('B1'), 0.0889633947, 1e-8);
  near(b.v('B2'), 0.0042206675, 1e-8);
  near(b.v('B3'), 10.58864445942, 1e-9);
  near(b.v('B4'), -66.6666667, 1e-6);
  near(b.v('B5'), -191.470830884, 1e-8);
  near(b.v('B6'), 480);
  near(b.v('B7'), 0.0535427, 1e-6);
});

test('text, math and dates', () => {
  const b = book({ A1: '=TEXTBEFORE("a-b-c","-",2)', A2: '=TEXTAFTER("a-b-c","-",-1)', A3: '=TEXTSPLIT("a,b;c,d",",",";")',
    C1: '=ROMAN(1999)', C2: '=HEX2DEC("FF")', C3: '=DEC2BIN(-1)', C4: '=GCD(24,36)', C5: '=MROUND(10,3)', C6: '=COMBIN(10,3)',
    D1: '=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,31))', D2: '=DATEDIF(DATE(2020,5,15),DATE(2024,3,1),"Y")', D3: '=ISOWEEKNUM(DATE(2024,12,30))',
    D4: '=WORKDAY(DATE(2024,1,5),1)', D5: '=YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),3)' });
  assert.strictEqual(b.v('A1'), 'a-b');
  assert.strictEqual(b.v('A2'), 'c');
  assert.deepStrictEqual([b.v('A3'), b.v('B3'), b.v('A4'), b.v('B4')], ['a', 'b', 'c', 'd']);
  assert.deepStrictEqual([b.v('C1'), b.v('C2'), b.v('C3'), b.v('C4'), b.v('C5'), b.v('C6')], ['MCMXCIX', 255, '1111111111', 12, 9, 120]);
  assert.strictEqual(b.v('D1'), 23);
  assert.strictEqual(b.v('D2'), 3);
  assert.strictEqual(b.v('D3'), 1);
  assert.strictEqual(b.v('D4'), F.dateToSerial(new Date(2024, 0, 8)));
  near(b.v('D5'), 182 / 365);
});

test('Excel names for newer functions round trip', () => {
  const f = 'LET(x,FILTER(A1:A9,A1:A9>0),SUM(x)+XLOOKUP("a",B1:B9,C1:C9))&"SORT(x)"';
  const x = F.toExcelFormula(f);
  assert.strictEqual(x, '_xlfn.LET(_xlpm.x,_xlfn._xlws.FILTER(A1:A9,A1:A9>0),SUM(_xlpm.x)+_xlfn.XLOOKUP("a",B1:B9,C1:C9))&"SORT(x)"');
  assert.strictEqual(F.fromExcelFormula(x), f);
  assert.strictEqual(F.toExcelFormula("SUM('My Sheet'!A1:B2,Sheet2!$A$1)"), "SUM('My Sheet'!A1:B2,Sheet2!$A$1)");
});
