'use strict';
// Grammar rules used by the Spelling & Grammar pane.
const test = require('node:test');
const assert = require('node:assert');
global.window = global.window || {};
require('../src/renderer/js/proofing.js');
const { grammar } = global.window.Proofing;
const fix = (text, lang) => {
  // Apply the first suggestion of every issue, last first.
  let out = text;
  for (const x of grammar(text, lang).filter(x => x.suggestions.length).reverse()) out = out.slice(0, x.start) + x.suggestions[0] + out.slice(x.end);
  return out;
};
const rules = (text, lang) => grammar(text, lang).map(x => x.rule);

test('catches common slips', () => {
  assert.strictEqual(fix('This is is a test.'), 'This is a test.');
  assert.strictEqual(fix('Two  spaces here.'), 'Two spaces here.');
  assert.strictEqual(fix('Hello , world .'), 'Hello, world.');
  assert.strictEqual(fix('one,two and three;four'), 'one, two and three; four');
  assert.strictEqual(fix('It rained. we stayed in.'), 'It rained. We stayed in.');
  assert.strictEqual(fix('Then i said no.'), 'Then I said no.');
  assert.strictEqual(fix('She ate a apple and an banana.'), 'She ate an apple and a banana.');
  assert.strictEqual(fix('I could of gone, it was better then nothing.'), 'I could have gone, it was better than nothing.');
  assert.strictEqual(fix('We have alot to do.'), 'We have a lot to do.');
});

test('leaves correct text alone', () => {
  for (const ok of [
    'A university is an honest place to spend an hour.',
    'A one-off event for an MBA class, e.g. the review. It went well.',
    'Prices rose 3.5 percent. Mr. Smith agreed, i.e. he said yes.',
    'She had had enough. That that is is that that is not.',
    'Version 2.0 shipped on time; see example.com/docs for details.',
    'C’est vrai ; il a dit : oui !'
  ]) assert.deepStrictEqual(rules(ok, /C’est/.test(ok) ? 'fr-FR' : 'en-US').filter(r => r !== 'repeat' || !/That that/.test(ok)), [], ok);
});

test('English-only rules stay off for other languages', () => {
  assert.deepStrictEqual(rules('Ich habe a apple gesehen.', 'de-DE'), []);
  assert.deepStrictEqual(rules('Das ist ist gut.', 'de-DE'), ['repeat']);
});

test('flags very long sentences', () => {
  const long = Array.from({ length: 50 }, (_, i) => 'word' + i).join(' ') + '.';
  const x = grammar(long, 'en-US').find(i => i.rule === 'long');
  assert.ok(x && x.suggestions.length === 0 && /50 words/.test(x.message));
});

test('double space typed in the editor (space + no-break space)', () => {
  assert.strictEqual(fix('End.\u00a0 Next'), 'End. Next');
  assert.strictEqual(fix('End. \u00a0Next'), 'End. Next');
});
