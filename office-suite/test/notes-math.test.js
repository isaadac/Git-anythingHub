'use strict';
const test = require('node:test');
const assert = require('node:assert');
const M = require('../src/renderer/js/notes-math.js');

test('values and fractions', () => {
  assert.strictEqual(M.analyze('12 × (3 + 4) =').text, '84');
  assert.strictEqual(M.analyze('3/4').text, '3/4 ≈ 0.75');
  assert.strictEqual(M.analyze('2^10').value, 1024);
  assert.ok(Math.abs(M.analyze('sin(pi/2) + √16').value - 5) < 1e-12);
  assert.strictEqual(M.analyze('2 + 2 = 4').value, true);
  assert.strictEqual(M.analyze('2 + 2 = 5').value, false);
});
test('linear equations with steps', () => {
  const r = M.analyze('2x + 3 = 11');
  assert.strictEqual(r.method, 'linear');
  assert.deepStrictEqual(r.solutions, [4]);
  assert.strictEqual(r.text, 'x = 4');
  assert.ok(r.steps.some(s => /Subtract 3/.test(s)) && r.steps.some(s => /Divide both sides by 2/.test(s)), r.steps.join('\n'));
  const r2 = M.analyze('5a - 2 = 3a + 7');
  assert.strictEqual(r2.text, 'a = 9/2 ≈ 4.5');
  assert.strictEqual(M.analyze('3(x + 1) = 3x + 3').text, 'True for every x');
  assert.strictEqual(M.analyze('x + 1 = x + 2').text, 'No solution');
});
test('quadratics, real and complex', () => {
  const r = M.analyze('x² - 5x + 6 = 0');
  assert.deepStrictEqual(r.solutions, [2, 3]);
  assert.strictEqual(r.text, 'x = 2 or x = 3');
  assert.deepStrictEqual(M.analyze('(x + 1)(x - 1) = 0').solutions, [-1, 1]);
  assert.deepStrictEqual(M.analyze('x^2 + 2x + 1 = 0').solutions, [-1]);
  const c = M.analyze('x^2 + 1 = 0');
  assert.strictEqual(c.solutions.length, 0);
  assert.strictEqual(c.text, 'x = 0 ± 1i');
});
test('other equations numerically, graphs, errors', () => {
  const r = M.analyze('2^x = 8');
  assert.strictEqual(r.method, 'numeric');
  assert.ok(Math.abs(r.solutions[0] - 3) < 1e-9, String(r.solutions));
  const s = M.analyze('sin(x) = 0.5');
  assert.ok(s.solutions.some(x => Math.abs(x - Math.PI / 6) < 1e-6));
  const g = M.analyze('y = x^2 - 4');
  assert.strictEqual(g.kind, 'graph');
  assert.strictEqual(g.fn(3), 5);
  assert.deepStrictEqual(g.roots().solutions, [-2, 2]);
  assert.match(M.graphSvg(g.fn, g.label), /^<svg[^>]+>.*<path d="M/);
  assert.strictEqual(M.analyze('1/x').kind, 'graph');
  assert.throws(() => M.analyze('2x + 3y = 4'), /2 letters/);
  assert.throws(() => M.analyze('(2 + 3'), /bracket/);
  assert.throws(() => M.analyze('2 # 3'), /“#”/);
  assert.throws(() => M.analyze('1 = 2 = 3'), /one = sign/);
});
