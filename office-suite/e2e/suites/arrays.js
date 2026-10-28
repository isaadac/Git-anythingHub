const { chromium } = require('playwright');
const JSZip = require('jszip');
(async () => {
  const b = await chromium.launch(); const p = await b.newPage({ viewport: { width: 1400, height: 900 } });
  const errors = []; p.on('pageerror', e => errors.push(e.message)); p.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  const check = (n, c, x = '') => console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : ''));
  await p.goto(process.env.E2E_BASE + '/src/renderer/sheets.html');
  await p.evaluate(() => {
    [['A1', 'Name'], ['B1', 'Score'], ['A2', 'Ann'], ['B2', '70'], ['A3', 'Bob'], ['B3', '95'], ['A4', 'Cy'], ['B4', '82'], ['A5', 'Ann'], ['B5', '60']].forEach(([k, v]) => AtlasSheets.setCell(k, v));
    AtlasSheets.setCell('D1', '=SORT(FILTER(A2:B5,B2:B5>65),2,-1)');
    AtlasSheets.setCell('G1', '=UNIQUE(A2:A5)');
    AtlasSheets.setCell('H1', '=SUM(D1#)');
  });
  const td = (r, c) => p.evaluate(([r, c]) => { const t = document.querySelector(`#grid tr[data-r="${r}"] td[data-c="${c}"]`); return t ? t.textContent : null; }, [r, c]);
  const grid = await Promise.all([[0, 3], [0, 4], [1, 3], [2, 4], [3, 3], [0, 6], [2, 6], [3, 6], [0, 7]].map(x => td(...x)));
  check('spilled results show in the grid', grid.join('|') === 'Bob|95|Cy|70||Ann|Cy||247', grid.join('|'));
  await p.evaluate(() => AtlasSheets.select(1, 4));
  const ui = await p.evaluate(() => ({ fx: document.getElementById('fx').value, ghost: document.getElementById('fx').classList.contains('ghost'), box: !document.querySelector('.spill-box').hidden }));
  check('a spilled cell shows its formula greyed and the spill outline', ui.fx === '=SORT(FILTER(A2:B5,B2:B5>65),2,-1)' && ui.ghost && ui.box, JSON.stringify(ui));
  await p.evaluate(() => AtlasSheets.setCell('E2', 'block'));
  check('typing in the spill range gives #SPILL!', (await td(0, 3)) === '#SPILL!');
  await p.evaluate(() => AtlasSheets.setCell('E2', ''));
  check('clearing it spills again', (await td(1, 4)) === '82');
  // xlsx: dynamic array formula, _xlfn names, cached values
  const bytes = await p.evaluate(async () => Array.from(await AtlasSheets.serialize('xlsx')));
  const zip = await JSZip.loadAsync(Buffer.from(bytes)); const sx = await zip.file('xl/worksheets/sheet1.xml').async('string');
  const d1 = (sx.match(/<c r="D1"[\s\S]*?<\/c>/) || [''])[0];
  check('xlsx: dynamic array formula with Excel names', /cm="1"/.test(d1) && /t="array" ref="D1:E3"/.test(d1) && /_xlfn\._xlws\.SORT\(_xlfn\._xlws\.FILTER\(/.test(d1) && !!zip.file('xl/metadata.xml'), d1);
  check('xlsx: spilled values cached', /<c r="E2"[^>]*><v>82<\/v>/.test(sx), (sx.match(/<c r="E2"[\s\S]*?<\/c>/) || [''])[0]);
  await p.evaluate(async b => { await AtlasSheets.loadFile({ name: 'a.xlsx', path: null, data: new Uint8Array(b) }); }, bytes);
  const back = await p.evaluate(() => ({ d1: AtlasSheets.book.sheets[0].cells.D1.v, e2: AtlasSheets.book.sheets[0].cells.E2, g1: AtlasSheets.book.sheets[0].cells.G1.v, v: AtlasSheets.valueOf('E2'), h: AtlasSheets.valueOf('H1') }));
  check('reopened: formulas back without prefixes, spill rebuilt', back.d1 === '=SORT(FILTER(A2:B5,B2:B5>65),2,-1)' && back.g1 === '=UNIQUE(A2:A5)' && !back.e2 && back.v === 82 && back.h === 247, JSON.stringify(back));
  check('no page errors', errors.length === 0, errors.join(' | '));
  await b.close();
})();
