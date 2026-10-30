const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch(); const p = await b.newPage({ viewport: { width: 1400, height: 900 } });
  const errors = []; p.on('pageerror', e => errors.push(e.message)); p.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  const check = (n, c, x = '') => console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : ''));
  await p.goto(process.env.E2E_BASE + '/src/renderer/sheets.html');
  const set = o => p.evaluate(o => Object.entries(o).forEach(([k, v]) => AtlasSheets.setCell(k, v)), o);
  const val = refs => p.evaluate(refs => refs.map(r => AtlasSheets.valueOf(r)), refs);
  const raw = refs => p.evaluate(refs => refs.map(r => { const m = /^([A-Z]+)(\d+)$/.exec(r); return (AtlasSheets.book.sheets[0].cells[r] || {}).v ?? ''; }), refs);
  const sel = (r1, c1, r2, c2) => p.evaluate(([r1, c1, r2, c2]) => { AtlasSheets.select(r1, c1); AtlasSheets.select(r2, c2, true); }, [r1, c1, r2, c2]);
  const cellCenter = (r, c) => p.evaluate(([r, c]) => { const t = document.querySelector(`#grid tr[data-r="${r}"] td[data-c="${c}"]`).getBoundingClientRect(); return { x: t.left + t.width / 2, y: t.top + t.height / 2 }; }, [r, c]);
  async function drag(r, c, mods = []) {
    const h = await p.evaluate(() => { const r = document.querySelector('.fill-handle').getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; });
    const to = await cellCenter(r, c);
    await p.mouse.move(h.x, h.y); await p.mouse.down();
    await p.mouse.move((h.x + to.x) / 2, (h.y + to.y) / 2, { steps: 3 }); await p.mouse.move(to.x, to.y, { steps: 3 });
    for (const m of mods) await p.keyboard.down(m);
    await p.mouse.up();
    for (const m of mods) await p.keyboard.up(m);
  }
  await set({ A1: '1', A2: '3', B1: 'Mon', C1: 'Item 1', D1: '=A1*10', E1: 'x', E2: 'y', F1: '5' });
  await sel(0, 0, 1, 4);
  await drag(5, 2);
  const r1 = await val(['A3', 'A6', 'B3', 'B5', 'B6', 'C5', 'D5', 'D6', 'E3', 'E6']);
  check('drag down continues each column (blanks keep their places)', r1.join() === '5,11,Tue,Wed,,Item 3,90,,x,y', r1.join());
  check('formulas shift as they fill', (await raw(['D5']))[0] === '=A5*10');
  check('selection grows over the filled cells', (await p.evaluate(() => document.getElementById('nameBox').value)) === 'A1:E6');
  // A single number is copied; Ctrl makes it a series
  await sel(0, 5, 0, 5); await drag(3, 5);
  await sel(0, 6, 0, 6); await set({ G1: '5' }); await sel(0, 6, 0, 6); await drag(3, 6, ['Control']);
  const r2 = await val(['F4', 'G4']);
  check('one number copies, Ctrl+drag counts', r2.join() === '5,8', r2.join());
  // Fill right and up
  await set({ A10: 'Jan', B10: 'Feb' }); await sel(9, 0, 9, 1); await drag(9, 4);
  const r3 = await val(['C10', 'E10']);
  check('drag right', r3.join() === 'Mar,May', r3.join());
  await set({ H10: '10', H11: '20' }); await sel(9, 7, 10, 7); await drag(7, 7);
  const r4 = await val(['H9', 'H8']);
  check('drag up runs the series backwards', r4.join() === '0,-10', r4.join());
  // Double-click fills down to match the next column
  await set({ J1: '1', K1: 'a', K2: 'b', K3: 'c', K4: 'd' }); await sel(0, 9, 0, 9);
  await p.evaluate(() => document.querySelector('.fill-handle').dispatchEvent(new MouseEvent('dblclick', { bubbles: true })));
  const r5 = await val(['J4', 'J5']);
  check('double-click fills down beside data', r5.join() === '1,', r5.join());
  // Undo restores
  await p.evaluate(() => { const e = new KeyboardEvent('keydown', { key: 'z', ctrlKey: true, bubbles: true }); document.getElementById('gridWrap').dispatchEvent(e); });
  check('undo removes the fill', (await val(['J4']))[0] === '');
  // Series dialog: growth to stop value
  await set({ M1: '2' }); await sel(0, 12, 0, 12);
  await p.evaluate(() => { UI.form = async () => ({ dir: 'cols', type: 'growth', unit: 'day', step: '3', stop: '100' }); });
  await p.click('#fillSeries');
  const r6 = await val(['M2', 'M4', 'M5']);
  check('Series dialog: growth with stop value', r6.join() === '6,54,', r6.join());
  // Dates
  await p.evaluate(() => { AtlasSheets.setCell('N1', '=DATE(2024,1,31)'); });
  await sel(0, 13, 0, 13);
  await p.evaluate(() => { UI.form = async () => ({ dir: 'cols', type: 'date', unit: 'month', step: '1', stop: '' }); });
  await sel(0, 13, 2, 13); await p.click('#fillSeries');
  const r7 = await p.evaluate(() => ['N2', 'N3'].map(k => { const d = new Date(Date.UTC(1899, 11, 30) + AtlasSheets.valueOf(k) * 864e5); return d.toISOString().slice(0, 10); }));
  check('Series dialog: months keep the end of month', r7.join() === '2024-02-29,2024-03-31', r7.join());
  check('no page errors', errors.length === 0, errors.join(' | '));
  await b.close();
})();
