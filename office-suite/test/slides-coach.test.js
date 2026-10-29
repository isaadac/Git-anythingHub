'use strict';
// Rehearse with Coach (src/renderer/js/slides-coach.js).
const test = require('node:test');
const assert = require('node:assert');
const C = require('../src/renderer/js/slides-coach.js');

test('filler words, including two-word ones, but not ordinary words', () => {
  const f = C.fillersIn('Um, so basically we, uh, grew. You know, I like it. You know what I mean? Umbrella.');
  assert.deepStrictEqual(Object.fromEntries(f), { um: 1, basically: 1, uh: 1, 'you know': 2, 'i mean': 1 });
});

test('reading a slide aloud is spotted; talking around it is not', () => {
  const slide = 'Revenue grew forty percent in the third quarter thanks to new customers in Europe';
  assert.ok(C.readAloud(slide, 'so revenue grew forty percent in the third quarter thanks to new customers in Europe') > 0.9);
  assert.ok(C.readAloud(slide, 'we had a great summer, mostly because European buyers discovered us') < 0.2);
  assert.strictEqual(C.readAloud('Short title', 'short title'), 0);
});

test('the report: pace, fillers, reading, timing and tips', () => {
  const slides = [{ text: 'Welcome' }, { text: 'Revenue grew forty percent in the third quarter thanks to new customers in Europe' }, { text: 'Questions' }];
  const heard = [
    { slide: 0, text: 'um hello everyone uh thanks for coming um today we look at the quarter' },
    { slide: 1, text: 'revenue grew forty percent in the third quarter thanks to new customers in Europe basically' },
    { slide: 2, text: 'any questions' }
  ];
  const r = C.analyze({ slides, heard, times: [5, 15, 150], seconds: 170 });
  assert.strictEqual(r.words, 31);
  assert.strictEqual(r.wpm, 11);
  assert.strictEqual(r.pace, 'slow');
  assert.deepStrictEqual(r.fillers, { total: 4, top: [['um', 2], ['uh', 1], ['basically', 1]] });
  assert.deepStrictEqual(r.reading, [1]);
  assert.deepStrictEqual(r.longest, { slide: 2, seconds: 150 });
  assert.strictEqual(r.tips.length, 5);
  assert.match(r.tips[0], /about 11 words a minute/);
  assert.match(r.tips[2], /read slide 2 word for word/);
  assert.match(r.tips[3], /Slide 3 took 2:30/);
  assert.match(r.tips[4], /little on slide 3/);
  const good = C.analyze({ slides, heard: [{ slide: 0, text: 'word '.repeat(140) }], times: [60], seconds: 60 });
  assert.strictEqual(good.pace, 'good');
  assert.match(good.tips[0], /Well paced/);
  assert.strictEqual(C.analyze({ slides, heard: [{ slide: 0, text: 'word '.repeat(200) }], times: [60], seconds: 60 }).pace, 'fast');
});
