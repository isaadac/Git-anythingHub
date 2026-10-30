'use strict';
// Reference formats for Write's Sources manager (src/renderer/js/write-refs.js).
const test = require('node:test');
const assert = require('node:assert');
const R = require('../src/renderer/js/write-refs.js');

const BIB = `% exported from Zotero
@article{lovelace1843,
  author = {Lovelace, Ada and Babbage, Charles},
  title = {{Notes} on the {Analytical Engine}},
  journal = "Scientific Memoirs",
  year = 1843,
  volume = {3},
  number = {1},
  pages = {666--731},
  doi = {https://doi.org/10.1000/xyz123}
}
@comment{ignored {nested} }
@book{G{\\"o}del1931,
  author = {G{\\"o}del, Kurt},
  title = {{\\"U}ber formal unentscheidbare S{\\"a}tze \\& more},
  publisher = {Monatshefte},
  year = {1931}
}
@online{site, title = {A web page}, url = {https://example.com/a}, date = {2021-05-01}, author = {{World Health Organization}}}`;

test('BibTeX: fields, authors, accents, braces, page ranges and DOIs', () => {
  const [a, b, c] = R.parseBibTeX(BIB);
  assert.deepStrictEqual({ ...a }, { type: 'article', authors: 'Lovelace, Ada; Babbage, Charles', title: 'Notes on the Analytical Engine', year: '1843', publisher: '', journal: 'Scientific Memoirs', volume: '3', issue: '1', pages: '666–731', url: 'https://doi.org/10.1000/xyz123', doi: '10.1000/xyz123', accessed: '' });
  assert.strictEqual(b.authors, 'Gödel, Kurt');
  assert.strictEqual(b.title, 'Über formal unentscheidbare Sätze & more');
  assert.strictEqual(c.type, 'website');
  assert.strictEqual(c.year, '2021');
  assert.strictEqual(c.authors, 'World Health Organization');
  assert.strictEqual(R.parseBibTeX(BIB).length, 3);
});

test('RIS records', () => {
  const ris = 'TY  - JOUR\r\nAU  - Hopper, Grace\r\nAU  - Mauchly, John\r\nTI  - The education of a computer\r\nJO  - Proceedings of the ACM\r\nPY  - 1952/05/02\r\nSP  - 243\r\nEP  - 249\r\nDO  - 10.1145/609784.609818\r\nER  - \r\nTY  - BOOK\r\nA1  - Knuth, Donald\r\nT1  - The Art of Computer Programming\r\nPB  - Addison-Wesley\r\nY1  - 1968\r\nER  -';
  const [a, b] = R.parseRIS(ris);
  assert.deepStrictEqual({ type: a.type, authors: a.authors, year: a.year, pages: a.pages, journal: a.journal, url: a.url }, { type: 'article', authors: 'Hopper, Grace; Mauchly, John', year: '1952', pages: '243–249', journal: 'Proceedings of the ACM', url: 'https://doi.org/10.1145/609784.609818' });
  assert.deepStrictEqual({ type: b.type, authors: b.authors, title: b.title, publisher: b.publisher, year: b.year }, { type: 'book', authors: 'Knuth, Donald', title: 'The Art of Computer Programming', publisher: 'Addison-Wesley', year: '1968' });
});

test('Crossref records and DOIs in text', () => {
  const s = R.fromCrossref({ DOI: '10.1038/171737a0', type: 'journal-article', title: ['Molecular Structure of Nucleic Acids'], author: [{ family: 'Watson', given: 'J. D.' }, { family: 'Crick', given: 'F. H. C.' }], 'container-title': ['Nature'], volume: '171', issue: '4356', page: '737-738', issued: { 'date-parts': [[1953, 4]] }, publisher: 'Springer' });
  assert.deepStrictEqual(s, { type: 'article', authors: 'Watson, J. D.; Crick, F. H. C.', title: 'Molecular Structure of Nucleic Acids', year: '1953', publisher: 'Springer', journal: 'Nature', volume: '171', issue: '4356', pages: '737–738', url: 'https://doi.org/10.1038/171737a0', doi: '10.1038/171737a0' });
  assert.strictEqual(R.doiIn('see https://doi.org/10.1038/171737a0.'), '10.1038/171737a0');
  assert.strictEqual(R.doiIn('no doi here'), '');
});

test('BibTeX out, with unique keys, reads back the same', () => {
  const list = R.parseBibTeX(BIB);
  const out = R.toBibTeX([...list, list[0]]);
  assert.match(out, /@article\{lovelace1843notes,/);
  assert.match(out, /@article\{lovelace1843notesa,/);
  assert.match(out, /pages = \{666--731\}/);
  const back = R.parseBibTeX(out);
  assert.strictEqual(back.length, 4);
  for (const k of ['authors', 'title', 'year', 'journal', 'pages', 'doi']) assert.strictEqual(back[0][k], list[0][k], k);
  assert.strictEqual(back[1].title, list[1].title);
});
