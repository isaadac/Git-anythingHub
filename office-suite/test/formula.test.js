'use strict';
const test = require('node:test');
const assert = require('node:assert');
const F = require('../src/renderer/js/formula.js');

function book(cells, sheets = { Sheet1: cells }) {
  return new F.Engine({
    getRaw: (s, r, c) => (sheets[s] || {})[F.addr(r, c)] ?? '',
    sheetNames: () => Object.keys(sheets)
  });
}
const val = (cells, ref = 'Z99', sheets) => {
  const e = book(cells, sheets);
  const p = F.parseRef(ref);
  const v = e.value('Sheet1', p.row, p.col);
  return v instanceof F.FormulaError ? v.error : v;
};
const calc = (formula, cells = {}) => val({ ...cells, Z99: formula });

test('arithmetic and precedence', () => {
  assert.strictEqual(calc('=1+2*3'), 7);
  assert.strictEqual(calc('=(1+2)*3'), 9);
  assert.strictEqual(calc('=2^3^2'), 512);
  assert.strictEqual(calc('=-2^2'), 4);
  assert.strictEqual(calc('=10/4'), 2.5);
  assert.strictEqual(calc('=50%'), 0.5);
  assert.strictEqual(calc('=1/0'), '#DIV/0!');
  assert.strictEqual(calc('="a"&"b"&1'), 'ab1');
});

test('references and ranges', () => {
  const c = { A1: '1', A2: '2', A3: '3', B1: 'x', B2: '=A1+A2' };
  assert.strictEqual(calc('=SUM(A1:A3)', c), 6);
  assert.strictEqual(calc('=B2*2', c), 6);
  assert.strictEqual(calc('=AVERAGE(A1:B3)', c), 2.25);
  assert.strictEqual(calc('=COUNT(A1:B3)', c), 4);
  assert.strictEqual(calc('=COUNTA(A1:B3)', c), 5);
  assert.strictEqual(calc('=$A$1+A$2', c), 3);
  assert.strictEqual(calc('=A1+B1', c), '#VALUE!');
});

test('cross-sheet references', () => {
  const sheets = { Sheet1: { Z99: "=Data!A1*2+'My Sheet'!B2" }, Data: { A1: '5' }, 'My Sheet': { B2: '1' } };
  assert.strictEqual(val({}, 'Z99', sheets), 11);
});

test('circular references are detected', () => {
  assert.strictEqual(val({ A1: '=B1', B1: '=A1' }, 'A1'), '#CIRC!');
});

test('logic functions', () => {
  assert.strictEqual(calc('=IF(1>2,"y","n")'), 'n');
  assert.strictEqual(calc('=IF(A1="",1,2)'), 1);
  assert.strictEqual(calc('=IFERROR(1/0,"bad")'), 'bad');
  assert.strictEqual(calc('=AND(TRUE,1)'), true);
  assert.strictEqual(calc('=OR(FALSE,0)'), false);
  assert.strictEqual(calc('=NOT(FALSE)'), true);
  assert.strictEqual(calc('=IFS(1>2,"a",2>1,"b")'), 'b');
  assert.strictEqual(calc('=SWITCH(2,1,"one",2,"two","other")'), 'two');
});

test('text functions', () => {
  assert.strictEqual(calc('=UPPER("abc")'), 'ABC');
  assert.strictEqual(calc('=LEFT("hello",2)&RIGHT("hello",2)&MID("hello",2,3)'), 'heloell');
  assert.strictEqual(calc('=PROPER("hello world")'), 'Hello World');
  assert.strictEqual(calc('=TRIM("  a   b ")'), 'a b');
  assert.strictEqual(calc('=SUBSTITUTE("a-b-c","-","+")'), 'a+b+c');
  assert.strictEqual(calc('=TEXTJOIN(", ",TRUE,"a","","b")'), 'a, b');
  assert.strictEqual(calc('=LEN("four")'), 4);
  assert.strictEqual(calc('=FIND("l","hello")'), 3);
  assert.strictEqual(calc('=TEXT(1234.5,"$#,##0.00")'), '$1,234.50');
});

test('conditional aggregates', () => {
  const c = { A1: 'apple', A2: 'pear', A3: 'apple', B1: '10', B2: '20', B3: '30' };
  assert.strictEqual(calc('=SUMIF(A1:A3,"apple",B1:B3)', c), 40);
  assert.strictEqual(calc('=COUNTIF(B1:B3,">15")', c), 2);
  assert.strictEqual(calc('=AVERAGEIF(A1:A3,"ap*",B1:B3)', c), 20);
  assert.strictEqual(calc('=SUMIFS(B1:B3,A1:A3,"apple",B1:B3,">10")', c), 30);
  assert.strictEqual(calc('=COUNTIFS(A1:A3,"apple",B1:B3,"<20")', c), 1);
});

test('lookups', () => {
  const c = { A1: 'a', A2: 'b', A3: 'c', B1: '1', B2: '2', B3: '3' };
  assert.strictEqual(calc('=VLOOKUP("b",A1:B3,2,FALSE)', c), 2);
  assert.strictEqual(calc('=VLOOKUP("z",A1:B3,2,FALSE)', c), '#N/A');
  assert.strictEqual(calc('=INDEX(B1:B3,MATCH("c",A1:A3,0))', c), 3);
  assert.strictEqual(calc('=XLOOKUP("a",A1:A3,B1:B3)', c), 1);
  assert.strictEqual(calc('=VLOOKUP(2.5,B1:B3,1,TRUE)', c), 2);
});

test('math and statistics', () => {
  assert.strictEqual(calc('=ROUND(2.345,2)'), 2.35);
  assert.strictEqual(calc('=ROUNDUP(2.341,2)'), 2.35);
  assert.strictEqual(calc('=ROUNDDOWN(2.349,2)'), 2.34);
  assert.strictEqual(calc('=MOD(-3,2)'), 1);
  assert.strictEqual(calc('=MEDIAN(1,3,2,4)'), 2.5);
  assert.strictEqual(calc('=MAX(1,5,3)'), 5);
  assert.strictEqual(calc('=SQRT(-1)'), '#NUM!');
  assert.strictEqual(calc('=SUMPRODUCT(A1:A2,B1:B2)', { A1: '2', A2: '3', B1: '4', B2: '5' }), 23);
  assert.ok(Math.abs(calc('=PMT(0.05/12,360,200000)') + 1073.64) < 0.01);
});

test('dates', () => {
  assert.strictEqual(calc('=DATE(2024,1,1)'), 45292);
  assert.strictEqual(calc('=YEAR(45292)'), 2024);
  assert.strictEqual(calc('=MONTH(DATE(2024,3,15))'), 3);
  assert.strictEqual(calc('=EOMONTH(DATE(2024,2,10),0)'), F.dateToSerial(new Date(2024, 1, 29)));
  assert.strictEqual(F.formatValue(45292, 'yyyy-mm-dd'), '2024-01-01');
  assert.strictEqual(F.formatValue(45292.5, 'h:mm AM/PM'), '12:00 PM');
});

test('unknown functions and syntax errors', () => {
  assert.strictEqual(calc('=NOPE(1)'), '#NAME?');
  assert.strictEqual(calc('=1+'), '#ERROR!');
});

test('number formats', () => {
  assert.strictEqual(F.formatValue(1234.567, '#,##0.00'), '1,234.57');
  assert.strictEqual(F.formatValue(0.256, '0.0%'), '25.6%');
  assert.strictEqual(F.formatValue(-5, '$#,##0.00'), '-$5.00');
  assert.strictEqual(F.formatValue(7, '000'), '007');
  assert.strictEqual(F.formatGeneral(0.1 + 0.2), '0.3');
});

test('literal parsing', () => {
  assert.strictEqual(F.parseLiteral('42'), 42);
  assert.strictEqual(F.parseLiteral('1,234.5'), 1234.5);
  assert.strictEqual(F.parseLiteral('$12'), 12);
  assert.strictEqual(F.parseLiteral('15%'), 0.15);
  assert.strictEqual(F.parseLiteral('true'), true);
  assert.strictEqual(F.parseLiteral("'007"), '007');
  assert.strictEqual(F.parseLiteral('hello'), 'hello');
});

test('shifting formulas on copy', () => {
  assert.strictEqual(F.shiftFormula('=A1+$B$1+C$2+$D3', 1, 1), '=B2+$B$1+D$2+$D4');
  assert.strictEqual(F.shiftFormula('=SUM(A1:A3)', 0, 1), '=SUM(B1:B3)');
  assert.strictEqual(F.shiftFormula('="A1"&A1', 1, 0), '="A1"&A2');
  assert.strictEqual(F.shiftFormula('=LOG10(A1)', 1, 0), '=LOG10(A2)');
  assert.strictEqual(F.shiftFormula('=A1', -1, 0), '=#REF!');
});

test('adjusting references for inserted and deleted rows', () => {
  assert.strictEqual(F.adjustForInsert('=SUM(A1:A5)+B7', 'row', 2, 1, 'Sheet1', 'Sheet1'), '=SUM(A1:A6)+B8');
  assert.strictEqual(F.adjustForInsert('=SUM(A1:A5)', 'row', 4, -1, 'Sheet1', 'Sheet1'), '=SUM(A1:A4)');
  assert.strictEqual(F.adjustForInsert('=A3', 'row', 2, -1, 'Sheet1', 'Sheet1'), '=#REF!');
  assert.strictEqual(F.adjustForInsert('=C1', 'col', 0, 2, 'Sheet1', 'Sheet1'), '=E1');
  assert.strictEqual(F.adjustForInsert('=Other!A5', 'row', 0, 1, 'Sheet1', 'Sheet1'), '=Other!A5');
});
