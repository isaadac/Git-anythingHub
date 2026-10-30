'use strict';
// Ink geometry (src/renderer/js/ink.js).
const test = require('node:test');
const assert = require('node:assert');
const Ink = require('../src/renderer/js/ink.js');

test('a smoothed line passes through midpoints and ends at the last point', () => {
  assert.strictEqual(Ink.smoothPath([[0, 0], [10, 0], [20, 10]]), 'M0 0 Q10 0 15 5 L20 10');
  assert.match(Ink.smoothPath([[5, 5]]), /^M5 5 L5\.1 5\.1$/); // a dot
});

test('pressure makes the outline wider', () => {
  const width = p => { const b = Ink.bounds([{ d: Ink.outlinePath([[0, 50, p], [50, 50, p], [100, 50, p]], 10), f: 1, w: 10 }]); return b.h; };
  assert.ok(width(1) > width(0.1) * 1.5, `${width(1)} vs ${width(0.1)}`);
  assert.ok(Math.abs(width(1) - 10) < 1.5);
  assert.match(Ink.outlinePath([[0, 0, 0.5], [20, 0, 0.5]], 4), /^M[\d. -]+.* Z$/);
  assert.match(Ink.outlinePath([[3, 3, 0.5]], 4), / Z$/); // a dot is a little circle
});

test('strokes from files are checked', () => {
  assert.deepStrictEqual(Ink.clean({ c: '#ff0000', w: 3, o: 0.5, d: 'M0 0 L1 1', f: 1, x: 'evil' }), { c: '#ff0000', w: 3, o: 0.5, d: 'M0 0 L1 1', f: 1 });
  assert.strictEqual(Ink.clean({ c: '#000', d: 'M0 0 <script>' }), null);
  assert.strictEqual(Ink.clean({ c: 'url(javascript:x)', d: 'M0 0' }).c, '#1f1f1f');
});

test('hit testing, bounds and moving', () => {
  const s = [{ c: '#000', w: 4, o: 1, d: 'M0 0 L100 0' }, { c: '#000', w: 4, o: 1, d: 'M0 50 L100 50' }];
  assert.strictEqual(Ink.hit(s, 50, 3), 0);
  assert.strictEqual(Ink.hit(s, 50, 48), 1);
  assert.strictEqual(Ink.hit(s, 50, 25), -1);
  assert.deepStrictEqual(Ink.bounds(s), { x: -2, y: -2, w: 104, h: 54 });
  assert.strictEqual(Ink.moved(s, 10, -5)[1].d, 'M10 45 L110 45');
});
