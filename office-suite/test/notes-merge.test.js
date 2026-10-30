'use strict';
// Notes sync: merging two copies of a notebook.
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ctx = { window: {} };
vm.runInNewContext(fs.readFileSync(path.join(__dirname, '../src/renderer/js/notes-extras.js'), 'utf8'), ctx);
const { merge, cleanFiles } = ctx.window.NotesExtras;
const plain = x => JSON.parse(JSON.stringify(x)); // values from the vm realm, as plain data
const page = (id, title, modified) => ({ id, title, html: `<p>${title}</p>`, ink: [], level: 0, created: 1, modified });
const nb = (sections, extra = {}) => ({ format: 'atlas-notes', version: 1, id: 'n1', name: 'NB', color: '#7719aa', modified: 1, deleted: {}, files: {}, sections, ...extra });

test('the newer copy of each page wins, and pages from both sides are kept', () => {
  const a = nb([{ id: 's1', name: 'S', color: '#000000', modified: 1, pages: [page('p1', 'one (A)', 5), page('p2', 'two', 1)] }]);
  const b = nb([{ id: 's1', name: 'S', color: '#000000', modified: 1, pages: [page('p1', 'one (B)', 3), page('p2', 'two (B)', 9), page('p3', 'three', 2)] }]);
  const m = merge(a, b);
  assert.deepStrictEqual(plain(m.sections[0].pages.map(p => p.title)), ['one (A)', 'two (B)', 'three']);
});

test('deletions carry over unless the page changed afterwards', () => {
  const a = nb([{ id: 's1', name: 'S', color: '#000000', modified: 1, pages: [page('p1', 'keep', 1)] }], { deleted: { p2: 10, p3: 10 } });
  const b = nb([{ id: 's1', name: 'S', color: '#000000', modified: 1, pages: [page('p1', 'keep', 1), page('p2', 'old', 5), page('p3', 'edited later', 20)] }]);
  assert.deepStrictEqual(plain(merge(a, b).sections[0].pages.map(p => p.title)), ['keep', 'edited later']);
});

test('a page moved to another section is where its newer copy is; new sections come along', () => {
  const a = nb([{ id: 's1', name: 'S1', color: '#000000', modified: 1, pages: [page('p1', 'moved', 9)] }, { id: 's2', name: 'S2', color: '#000000', modified: 1, pages: [page('p2', 'x', 1)] }]);
  const b = nb([{ id: 's1', name: 'S1', color: '#000000', modified: 1, pages: [page('p3', 'y', 1)] }, { id: 's2', name: 'S2', color: '#000000', modified: 1, pages: [page('p1', 'moved (old)', 2), page('p2', 'x', 1)] }, { id: 's3', name: 'New', color: '#000000', modified: 4, pages: [page('p4', 'z', 4)] }]);
  const m = merge(a, b);
  assert.deepStrictEqual(plain(m.sections.map(s => s.name + ':' + s.pages.map(p => p.title).join('/'))), ['S1:moved/y', 'S2:x', 'New:z']);
});

test('attachments: only well-formed files are kept', () => {
  const ok = cleanFiles({ fabcde: { name: 'a.pdf', type: 'application/pdf', size: 10, data: 'JVBERi0BAgP6+w==' }, fzzzzz: { name: 'bad', data: '<script>' }, bad: { data: 'AAAA' } });
  assert.deepStrictEqual(plain(Object.keys(ok)), ['fabcde']);
});
