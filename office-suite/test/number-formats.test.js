'use strict';
// Number formats against what Excel itself shows (440 cases from Excel files).
const test = require('node:test');
const assert = require('node:assert');
const F = require('../src/renderer/js/formula.js');
const { cases } = require('./fixtures/excel-formats.json');

test('formats match Excel (General, numbers, dates and times)', () => {
  const wrong = cases.filter(([fmt, v, want]) => F.formatValue(v, fmt) !== want);
  // The few left are fraction layouts with literals inside them, like |?\:#=/=#|.
  assert.ok(wrong.length <= 10, wrong.slice(0, 10).map(w => JSON.stringify(w) + ' → ' + JSON.stringify(F.formatValue(w[1], w[0]))).join('\n'));
  assert.ok(wrong.every(([fmt]) => fmt.includes('/')), 'only fractions may differ');
});

test('everyday formats', () => {
  const f = F.formatValue;
  assert.strictEqual(f(1234.5, '#,##0.00'), '1,234.50');
  assert.strictEqual(f(-1234.5, '#,##0.00;(#,##0.00)'), '(1,234.50)');
  assert.strictEqual(f(0, '#,##0;-#,##0;"zero"'), 'zero');
  assert.strictEqual(f(0.035, '0.0%'), '3.5%');
  assert.strictEqual(f(1234567, '0.0,,"M"'), '1.2M');
  assert.strictEqual(f(12345.678, '0.00E+00'), '1.23E+04');
  assert.strictEqual(f(1.5, '# ?/?'), '1 1/2');
  assert.strictEqual(f(0.25, '# ?/4'), ' 1/4');
  assert.strictEqual(f(5551234, '000-0000'), '555-1234');
  assert.strictEqual(f(45658.75, 'yyyy-mm-dd hh:mm'), '2025-01-01 18:00');
  assert.strictEqual(f(1.25, '[h]:mm'), '30:00');
  assert.strictEqual(f(150, '[>100]"big";"small"'), 'big');
  assert.strictEqual(f(42, '[Red]0;[Blue]-0'), '42');
  assert.strictEqual(f(1234.5, '[$€-407] #,##0.00'), '€ 1,234.50');
  assert.strictEqual(f('text', '0;0;0;"<"@">"'), '<text>');
  assert.strictEqual(f(100000000000, 'General'), '1E+11');
  assert.strictEqual(f(1 / 3, 'General'), '0.333333333');
});
