'use strict';
// Write's equation syntax: the tree, MathML and Office Math made from it.
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ctx = { window: {} };
vm.runInNewContext(fs.readFileSync(path.join(__dirname, '../src/renderer/js/write-math.js'), 'utf8'), ctx);
const M = ctx.window.WriteMath;

test('fractions take the factors either side, without their parentheses', () => {
  const t = M.parse('(a+b)/(2c)');
  assert.strictEqual(t.t, 'frac');
  assert.strictEqual(t.n.t, 'row');
  assert.match(M.toMathML('(a+b)/(2c)'), /<mfrac><mrow><mi>a<\/mi><mo>\+<\/mo><mi>b<\/mi><\/mrow><mrow><mn>2<\/mn><mi>c<\/mi><\/mrow><\/mfrac>/);
});

test('the quadratic formula', () => {
  const ml = M.toMathML('x = (-b \\pm sqrt(b^2 - 4ac))/(2a)');
  assert.match(ml, /<mo>±<\/mo>/);
  assert.match(ml, /<msqrt>.*<msup><mi>b<\/mi><mn>2<\/mn><\/msup>.*<\/msqrt>/);
  assert.match(ml, /<mfrac>/);
});

test('scripts, roots, sums and Greek letters', () => {
  assert.match(M.toMathML('a_i^2'), /<msubsup><mi>a<\/mi><mi>i<\/mi><mn>2<\/mn><\/msubsup>/);
  assert.match(M.toMathML('sqrt[3](x)'), /<mroot><mi>x<\/mi><mn>3<\/mn><\/mroot>/);
  assert.match(M.toMathML('\\sum_(i=1)^n i'), /<munderover><mo largeop="true">∑<\/mo>.*<mi>n<\/mi><\/munderover><mi>i<\/mi>/);
  assert.match(M.toMathML('\\alpha + \\beta'), /<mi>α<\/mi><mo>\+<\/mo><mi>β<\/mi>/);
  assert.match(M.toMathML('sin x'), /<mi>sin<\/mi><mo>&#x2061;<\/mo><mi>x<\/mi>/);
});

test('Office Math for Word', () => {
  const o = M.toOMML('x^2/(2a) + sqrt(y)');
  assert.match(o, /^<m:oMath><m:f><m:num><m:sSup><m:e>.*<m:t xml:space="preserve">x<\/m:t>.*<\/m:sSup><\/m:num><m:den>/);
  assert.match(o, /<m:rad><m:radPr><m:degHide m:val="1"\/><\/m:radPr><m:deg\/><m:e>/);
  assert.match(M.toOMML('\\sum_(k=0)^n k'), /<m:nary><m:naryPr><m:chr m:val="∑"\/><m:limLoc m:val="undOvr"\/><\/m:naryPr><m:sub>/);
  assert.match(M.toOMML('[a]'), /<m:d><m:dPr><m:begChr m:val="\["\/><m:endChr m:val="\]"\/><\/m:dPr>/);
  assert.doesNotMatch(M.toOMML('a<b & "x"'), /<b|& /);
});

test('anything typed gives some equation (no exceptions)', () => {
  for (const s of ['', ')', '((', '^', '/', 'a/', '\\unknown', 'sqrt', '\\sum', '{', '"open', 'a__b^^c']) assert.doesNotThrow(() => { M.toMathML(s); M.toOMML(s); }, s);
});
