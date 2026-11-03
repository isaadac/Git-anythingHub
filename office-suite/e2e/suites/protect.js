const { chromium } = require('playwright');
const fs = require('fs');
const { execFileSync } = require('child_process');
const SP = process.env.E2E_WORK;
(async () => {
  const b = await chromium.launch(); const p = await b.newPage({ viewport: { width: 1400, height: 900 } });
  const errors = []; p.on('pageerror', e => errors.push(e.message)); p.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  const check = (n, c, x = '') => console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : ''));
  await p.goto(process.env.E2E_BASE + '/src/renderer/sheets.html');
  await p.evaluate(() => { localStorage.setItem('atlas-author', 'Ada'); ['A1', 'A2'].forEach((k, i) => AtlasSheets.setCell(k, ['Price', '10'][i])); });
  const answer = list => p.evaluate(list => { window.__a = list.slice(); UI.form = async () => window.__a.shift(); }, list);
  const sel = (r, c, r2 = r, c2 = c) => p.evaluate(([r, c, r2, c2]) => { AtlasSheets.select(r, c); if (r2 !== r || c2 !== c) AtlasSheets.select(r2, c2, true); }, [r, c, r2, c2]);
  await p.click('[data-tab="review"]');
  // Notes
  await sel(1, 0); await answer([{ t: 'Before tax' }]); await p.click('#noteNew');
  await sel(4, 3); await answer([{ t: 'Note on an empty cell' }]); await p.click('#noteNew');
  const ind = await p.evaluate(() => [...document.querySelectorAll('#grid td.has-note')].length);
  check('notes show a red corner', ind === 2, ind);
  const td = await p.evaluate(() => { const t = document.querySelector('#grid tr[data-r="1"] td[data-c="0"]'); const r = t.getBoundingClientRect(); return { x: r.left + 10, y: r.top + 8 }; });
  await p.mouse.move(td.x, td.y); await p.waitForTimeout(100);
  const hover = await p.evaluate(() => { const n = document.querySelector('.note-pop:not([hidden])'); return n && n.textContent; });
  check('hovering shows the note with its author', hover === 'Ada:Before tax', hover);
  await p.click('#noteShowAll');
  check('Show All Notes', (await p.evaluate(() => document.querySelectorAll('.note-layer .note-pop').length)) === 2);
  await p.click('#noteShowAll');
  // Clearing contents keeps the note
  await sel(1, 0); await p.evaluate(() => document.getElementById('gridWrap').dispatchEvent(new KeyboardEvent('keydown', { key: 'Delete', bubbles: true })));
  check('Delete clears the value but keeps the note', await p.evaluate(() => AtlasSheets.book.sheets[0].cells.A2 && AtlasSheets.book.sheets[0].cells.A2.n.text === 'Before tax' && AtlasSheets.book.sheets[0].cells.A2.v === ''));
  await p.evaluate(() => AtlasSheets.setCell('A2', '10'));
  // Unlock B2:B3 for input, then protect with a password
  await sel(1, 1, 2, 1); await p.click('#lockCell');
  await answer([{ pw: 'test', pw2: 'test', formatCells: '', insertRows: '', insertColumns: '', sort: '' }]); await p.click('#protectSheet');
  const st = await p.evaluate(() => ({ prot: AtlasSheets.book.sheets[0].protect, label: document.querySelector('#protectSheet .lbl').textContent }));
  check('Protect Sheet stores Excel\'s password hash', st.prot && st.prot.hash === 'CBEB' && st.label === 'Unprotect Sheet', JSON.stringify(st));
  // Typing in a locked cell is refused; unlocked cell works
  await sel(0, 0); await p.focus('#gridWrap'); await p.keyboard.type('x');
  const toast = await p.evaluate(() => document.querySelector('.toast.show') && document.querySelector('.toast').textContent);
  const editing = await p.evaluate(() => !document.getElementById('cellEditor').hidden);
  check('locked cells can\'t be edited', /protected sheet/.test(toast || '') && !editing && (await p.evaluate(() => AtlasSheets.valueOf('A1'))) === 'Price', toast);
  await sel(1, 1); await p.keyboard.type('42'); await p.keyboard.press('Enter');
  check('unlocked cells can be edited', (await p.evaluate(() => AtlasSheets.valueOf('B2'))) === 42);
  // Formatting, clearing, inserting rows are blocked
  await sel(0, 0); await p.click('[data-tab="home"]'); await p.click('[data-style="b"]');
  await p.evaluate(() => document.getElementById('gridWrap').dispatchEvent(new KeyboardEvent('keydown', { key: 'Delete', bubbles: true })));
  await p.click('#insRow');
  const after = await p.evaluate(() => ({ b: !!(AtlasSheets.book.sheets[0].cells.A1.s || {}).b, a1: AtlasSheets.valueOf('A1'), a2: AtlasSheets.valueOf('A2') }));
  check('formatting, clearing and inserting rows are blocked', !after.b && after.a1 === 'Price' && after.a2 === 10, JSON.stringify(after));
  // xlsx: protection, unlocked cells and notes for Excel
  const bytes = await p.evaluate(async () => Array.from(await AtlasSheets.serialize('xlsx')));
  fs.writeFileSync(SP + '/files/protect-out.xlsx', Buffer.from(bytes));
  const py = execFileSync('python3', ['-I', '-c', `
import sys
from openpyxl import load_workbook
ws = load_workbook(sys.argv[1]).active
print(ws.protection.sheet, ws.protection.password, ws.protection.formatCells)
print(ws['A1'].protection.locked, ws['B2'].protection.locked, ws['B3'].protection.locked)
print(ws['A2'].comment.text if ws['A2'].comment else None, '|', ws['A2'].comment.author if ws['A2'].comment else None, '|', ws['D5'].comment.text if ws['D5'].comment else None)
`, SP + '/files/protect-out.xlsx']).toString().trim().split('\n');
  check('openpyxl sees the protection, the unlocked cells and the notes', py[0] === 'True CBEB True' && py[1] === 'True False False' && py[2] === 'Before tax | Ada | Note on an empty cell', JSON.stringify(py));
  // Reopen as an Excel file (no Atlas metadata) via openpyxl re-save
  execFileSync('python3', ['-I', '-c', `
import sys
from openpyxl import load_workbook
wb = load_workbook(sys.argv[1]); wb.save(sys.argv[2])`, SP + '/files/protect-out.xlsx', SP + '/files/protect-resaved.xlsx']);
  const re = await p.evaluate(async b => {
    await AtlasSheets.loadFile({ name: 'r.xlsx', path: null, data: new Uint8Array(b) });
    const s = AtlasSheets.book.sheets[0];
    return { prot: s.protect, b2: (s.cells.B2.s || {}).unlocked, b3: (s.cells.B3 && s.cells.B3.s || {}).unlocked, a1: (s.cells.A1.s || {}).unlocked, note: s.cells.A2.n, d5: s.cells.D5 && s.cells.D5.n };
  }, Array.from(fs.readFileSync(SP + '/files/protect-resaved.xlsx')));
  check('an Excel-saved file keeps protection, unlocked cells and notes', re.prot && re.prot.hash === 'CBEB' && re.b2 && re.b3 && !re.a1 && re.note.text === 'Before tax' && re.note.author === 'Ada' && re.d5.text === 'Note on an empty cell', JSON.stringify(re));
  // Unprotect needs the password
  await p.click('[data-tab="review"]');
  await answer([{ pw: 'wrong' }]); await p.click('#protectSheet');
  const still = await p.evaluate(() => !!AtlasSheets.book.sheets[0].protect);
  await answer([{ pw: 'test' }]); await p.click('#protectSheet'); await p.waitForTimeout(50);
  const gone = await p.evaluate(() => !AtlasSheets.book.sheets[0].protect);
  check('Unprotect Sheet checks the password', still && gone);
  check('no page errors', errors.length === 0, errors.join(' | '));
  await b.close();
})();
