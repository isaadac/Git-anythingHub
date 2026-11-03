const { chromium } = require('playwright');
const JSZip = require('jszip');
const fs = require('fs');
const SP = process.env.E2E_WORK;
(async () => {
  const b = await chromium.launch(); const p = await b.newPage({ viewport: { width: 1400, height: 900 } });
  const errors = []; p.on('pageerror', e => errors.push(e.message)); p.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  const check = (n, c, x = '') => console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : ''));
  await p.goto(process.env.E2E_BASE + '/src/renderer/slides.html'); await p.waitForTimeout(300);
  await p.evaluate(() => {
    const s = AtlasSlides.pres.slides[0];
    s.elements = [
      { id: 'A', type: 'shape', shape: 'rect', fill: '#4472c4', html: '', x: 100, y: 100, w: 200, h: 100, size: 18, align: 'center', valign: 'middle' },
      { id: 'B', type: 'shape', shape: 'ellipse', fill: '#ed7d31', html: '', x: 600, y: 300, w: 160, h: 100, size: 18, align: 'center', valign: 'middle' },
      { id: 'C', type: 'shape', shape: 'triangle', fill: '#70ad47', html: '', x: 400, y: 380, w: 100, h: 80, size: 18, align: 'center', valign: 'middle' }
    ];
    AtlasSlides.goTo(0);
  });
  const click = async (id, opts) => p.click(`#stage [data-id="${id}"]`, { position: { x: 20, y: 20 }, ...opts });
  // Connector A -> B
  await p.click('[data-tab="insert"]');
  await p.click('[data-connector="elbow"]');
  await click('A'); await click('B');
  const conn = await p.evaluate(() => { const c = AtlasSlides.pres.slides[0].elements.find(e => e.type === 'connector'); return c && { kind: c.kind, from: c.from.id, to: c.to.id, path: document.querySelector('#stage .el.connector svg > path:not(.conn-hit)').getAttribute('d') }; });
  check('elbow connector joins two shapes', conn && conn.kind === 'elbow' && conn.from === 'A' && conn.to === 'B' && /^M\d+,\d+ H[\d.]+ V\d+ H\d+$/.test(conn.path), JSON.stringify(conn));
  // Move B: the connector follows
  const bb = await p.evaluate(() => { const r = document.querySelector('#stage [data-id="B"]').getBoundingClientRect(); return { x: r.left + 30, y: r.top + 30 }; });
  await p.mouse.move(bb.x, bb.y); await p.mouse.down(); await p.mouse.move(bb.x, bb.y + 60, { steps: 4 }); await p.mouse.up();
  const after = await p.evaluate(() => ({ by: AtlasSlides.pres.slides[0].elements.find(e => e.id === 'B').y, d: document.querySelector('#stage .el.connector svg > path:not(.conn-hit)').getAttribute('d') }));
  check('moving a shape moves the connector end', after.by > 330 && after.d !== conn.path, JSON.stringify(after));
  // Group A and C (Shift+click), resize scales children, ungroup restores
  await p.keyboard.press('Escape');
  await click('A'); await click('C', { modifiers: ['Shift'] });
  await p.keyboard.press('Control+g');
  const g = await p.evaluate(() => { const g = AtlasSlides.pres.slides[0].elements.find(e => e.type === 'group'); const c = AtlasSlides.pres.slides[0].elements.find(e => e.type === 'connector'); return g && { x: g.x, y: g.y, w: g.w, h: g.h, kids: g.children.map(k => k.id + '@' + k.x + ',' + k.y).join(' '), connFrom: c.from.id === g.id }; });
  check('Shift+click then Ctrl+G groups; connectors now attach to the group', g && g.x === 100 && g.y === 100 && g.w === 400 && g.h === 360 && g.kids === 'A@0,0 C@300,280' && g.connFrom, JSON.stringify(g));
  await p.click('[data-tab="arrange"]'); await p.click('#rotRight'); await p.click('#flipH');
  const rot = await p.evaluate(() => { const g = AtlasSlides.pres.slides[0].elements.find(e => e.type === 'group'); return [g.rot, g.flipH]; });
  check('Rotate right 90° and Flip horizontal', rot[0] === 90 && rot[1] === true, JSON.stringify(rot));
  // Rotation handle with Shift snaps to 15°
  const h = await p.evaluate(() => { const r = document.querySelector('#stage .handle.rot').getBoundingClientRect(); const s = document.querySelector('#stage .el.selected').getBoundingClientRect(); return { hx: r.left + 6, hy: r.top + 6, cx: s.left + s.width / 2, cy: s.top + s.height / 2 }; });
  await p.mouse.move(h.hx, h.hy); await p.mouse.down(); await p.keyboard.down('Shift');
  await p.mouse.move(h.cx + 100, h.cy - 98, { steps: 5 }); await p.mouse.up(); await p.keyboard.up('Shift');
  const r2 = await p.evaluate(() => AtlasSlides.pres.slides[0].elements.find(e => e.type === 'group').rot);
  check('rotation handle snaps to 15° with Shift', r2 === 45, String(r2));
  await p.evaluate(() => { const g = AtlasSlides.pres.slides[0].elements.find(e => e.type === 'group'); delete g.rot; delete g.flipH; });
  // pptx round trip
  const bytes = await p.evaluate(async () => Array.from(await AtlasSlides.serialize('pptx')));
  const zip = await JSZip.loadAsync(Buffer.from(bytes));
  const sx = await zip.file('ppt/slides/slide1.xml').async('string');
  const ids = Object.fromEntries([...sx.matchAll(/<p:cNvPr id="(\d+)" name="atlas-([^"]+)"/g)].map(m => [m[2], m[1]]));
  const cx = (sx.match(/<p:cxnSp>[\s\S]*?<\/p:cxnSp>/) || [''])[0];
  check('pptx: a real elbow connector joined to the shapes', /prst="bentConnector3"/.test(cx) && new RegExp(`<a:stCxn id="${ids.C}"`).test(cx) /* the group child nearest the start */ && new RegExp(`<a:endCxn id="${ids.B}" idx="\\d"`).test(cx) && !/<p:grpSp>/.test(sx), cx.slice(0, 300));
  const back = await p.evaluate(async b => { await AtlasSlides.loadFile({ name: 'r.pptx', path: null, data: new Uint8Array(b) }); const els = AtlasSlides.pres.slides[0].elements; const g = els.find(e => e.type === 'group'); const c = els.find(e => e.type === 'connector'); return { types: els.map(e => e.type).join(), kids: g && g.children.map(k => k.shape + '@' + k.x + ',' + k.y).join(' '), connected: c && !!c.from && !!c.to && c.from.id === g.id && els.find(e => e.id === c.to.id).shape === 'ellipse', kind: c && c.kind }; }, bytes);
  check('reopened: group, connector (still joined) and shapes', back.types.split(',').sort().join() === 'connector,group,shape' && back.kids === 'rect@0,0 triangle@300,280' && back.connected && back.kind === 'elbow', JSON.stringify(back));
  // A PowerPoint connector (python-pptx)
  const pp = await p.evaluate(async b => { await AtlasSlides.loadFile({ name: 'c.pptx', path: null, data: new Uint8Array(b) }); const els = AtlasSlides.pres.slides[0].elements; const c = els.find(e => e.type === 'connector'); return c && { kind: c.kind, from: els.find(e => e.id === (c.from || {}).id)?.shape, to: els.find(e => e.id === (c.to || {}).id)?.shape }; }, Array.from(fs.readFileSync(SP + '/files/conn.pptx')));
  check('PowerPoint connectors open attached to their shapes', pp && pp.kind === 'elbow' && pp.from === 'rect' && pp.to === 'ellipse', JSON.stringify(pp));
  check('no page errors', errors.length === 0, errors.join(' | '));
  await b.close();
})();
