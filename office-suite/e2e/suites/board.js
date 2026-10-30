// Isadek Whiteboard: sticky notes, shapes, + handles, connectors, mind maps (Tab, Enter, Tidy
// Up), frames, drawing and erasing, undo, files, and export as SVG and PNG.
const { chromium } = require('playwright');
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
(async () => {
  const b = await chromium.launch();
  const ctx = await b.newContext({ viewport: { width: 1400, height: 900 }, permissions: ['clipboard-read', 'clipboard-write'] });
  const p = await ctx.newPage();
  const errors = [];
  p.on('pageerror', e => errors.push(e.message));
  p.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  await p.goto(process.env.E2E_BASE + '/src/renderer/board.html');
  await p.waitForFunction(() => window.AtlasBoard);
  check('a new board is empty and says how to start', await p.isVisible('#empty'));
  const B = () => p.evaluate(() => JSON.parse(JSON.stringify(AtlasBoard.board.items)));
  const scr = (x, y) => p.evaluate(([x, y]) => AtlasBoard.toScreen(x, y), [x, y]);

  // A sticky note: click, type, Escape.
  await p.click('[data-tool="sticky"]');
  let [sx, sy] = await scr(0, 0);
  await p.mouse.click(sx, sy);
  await p.waitForSelector('#editor:not([hidden])');
  await p.keyboard.type('Hello board');
  await p.keyboard.press('Escape');
  let items = await B();
  check('a sticky note is added and typed in', items.length === 1 && items[0].type === 'sticky' && items[0].text === 'Hello board', JSON.stringify(items));
  check('its words are drawn on the board', (await p.textContent('#items .item-sticky text')).includes('Hello board'));

  // A rectangle, dragged to size (R).
  await p.focus('#board');
  await p.keyboard.press('r');
  [sx, sy] = await scr(300, -60);
  await p.mouse.move(sx, sy); await p.mouse.down(); await p.mouse.move(sx + 100, sy + 60, { steps: 4 }); await p.mouse.move(sx + 200, sy + 120, { steps: 4 }); await p.mouse.up();
  await p.keyboard.type('Step one');
  await p.keyboard.press('Escape');
  items = await B();
  const rect = items.find(i => i.type === 'shape');
  check('a shape is drawn to the size dragged', rect && rect.shape === 'rect' && Math.round(rect.w) === 200 && Math.round(rect.h) === 120 && rect.text === 'Step one', JSON.stringify(rect));

  // The + handle on its right: a connected copy.
  await p.evaluate(id => AtlasBoard.select([id]), rect.id);
  await p.click('#overlay .plus[data-side="e"]');
  await p.keyboard.type('Step two');
  await p.keyboard.press('Escape');
  items = await B();
  const two = items.find(i => i.type === 'shape' && i.text === 'Step two'), link = items.find(i => i.type === 'connector');
  check('+ adds a connected shape beside it', two && two.x > rect.x + rect.w && link && link.from.id === rect.id && link.to.id === two.id, JSON.stringify(items.map(i => [i.type, i.text, i.x])));

  // The connector tool: from the sticky to the first shape.
  const sticky = items.find(i => i.type === 'sticky');
  await p.click('[data-tool="connector"]');
  const [ax, ay] = await scr(sticky.x + sticky.w / 2, sticky.y + sticky.h / 2), [bx, by] = await scr(rect.x + rect.w / 2, rect.y + rect.h / 2);
  await p.mouse.move(ax, ay); await p.mouse.down(); await p.mouse.move(bx, by, { steps: 8 }); await p.mouse.up();
  items = await B();
  check('the connector tool joins two items', items.filter(i => i.type === 'connector').some(c => c.from && c.from.id === sticky.id && c.to && c.to.id === rect.id));

  // Moving an item: its connectors follow.
  await p.click('[data-tool="select"]');
  const before = await p.getAttribute(`#items [data-id="${link.id}"] path:not(.conn-hit)`, 'd');
  const [mx, my] = await scr(two.x + 20, two.y + 20);
  await p.mouse.move(mx, my); await p.mouse.down(); await p.mouse.move(mx + 60, my + 120, { steps: 6 }); await p.mouse.up();
  const moved = (await B()).find(i => i.id === two.id);
  const after = await p.getAttribute(`#items [data-id="${link.id}"] path:not(.conn-hit)`, 'd');
  check('dragging moves an item and its connector follows', Math.round(moved.y - two.y) === 120 && before !== after, `${moved.y - two.y} ${before} -> ${after}`);

  // Undo and redo.
  await p.focus('#board');
  await p.keyboard.press('Control+z');
  check('undo puts it back', Math.round((await B()).find(i => i.id === two.id).y) === Math.round(two.y));
  await p.keyboard.press('Control+y');
  check('redo moves it again', Math.round((await B()).find(i => i.id === two.id).y) === Math.round(moved.y));

  // A mind map: Tab adds a branch, Enter a sibling, Tidy Up lays it out.
  await p.evaluate(() => { AtlasBoard.board.items.length = 0; AtlasBoard.template('mindmap'); });
  const root = (await B()).find(i => i.type === 'shape');
  await p.evaluate(id => AtlasBoard.select([id]), root.id);
  await p.focus('#board');
  await p.keyboard.press('Tab');
  await p.keyboard.type('Branch A');
  await p.keyboard.press('Enter');
  await p.keyboard.type('Branch B');
  await p.keyboard.press('Escape');
  items = await B();
  const kids = items.filter(i => i.type === 'connector' && i.from && i.from.id === root.id).map(c => items.find(i => i.id === c.to.id));
  check('Tab adds a branch and Enter a sibling', kids.length === 5 && kids.some(k => k.text === 'Branch A') && kids.some(k => k.text === 'Branch B'), kids.map(k => k.text).join(','));
  await p.evaluate(id => { AtlasBoard.select([id]); AtlasBoard.tidy(); }, root.id);
  items = await B();
  const laid = items.filter(i => i.type === 'connector' && i.from && i.from.id === root.id).map(c => items.find(i => i.id === c.to.id));
  const col = [...laid].sort((a, b) => a.y - b.y);
  check('Tidy Up puts the branches in a column beside the root, without overlaps', laid.every(k => k.x > root.x + root.w) && col.every((k, n) => n === 0 || k.y >= col[n - 1].y + col[n - 1].h), col.map(k => `${k.y}+${k.h}`).join(','));

  // Frames carry what is on them.
  await p.evaluate(() => { AtlasBoard.board.items.length = 0; AtlasBoard.template('retro'); });
  items = await B();
  const frame = items.find(i => i.type === 'frame'), inside = items.find(i => i.type === 'sticky' && i.x >= frame.x && i.x + i.w <= frame.x + frame.w);
  check('the Retro template makes three frames with notes', items.filter(i => i.type === 'frame').length === 3 && items.filter(i => i.type === 'sticky').length === 6);
  await p.evaluate(() => AtlasBoard.zoomFit());
  const [fx, fy] = await scr(frame.x + 20, frame.y - 14);
  await p.mouse.move(fx, fy); await p.mouse.down(); await p.mouse.move(fx + 50, fy + 40, { steps: 5 }); await p.mouse.up();
  items = await B();
  const k = await p.evaluate(() => AtlasBoard.view.k);
  const f2 = items.find(i => i.id === frame.id), in2 = items.find(i => i.id === inside.id);
  check('dragging a frame by its title moves its notes too', Math.abs((f2.x - frame.x) - (in2.x - inside.x)) < 1 && Math.abs(f2.x - frame.x - 50 / k) < 6, `${f2.x - frame.x} ${in2.x - inside.x}`);

  // Drawing, then erasing.
  await p.evaluate(() => { AtlasBoard.board.items.length = 0; AtlasBoard.zoomFit(); });
  await p.click('[data-tool="pen"]');
  [sx, sy] = await scr(-100, 100);
  await p.mouse.move(sx, sy); await p.mouse.down();
  for (let n = 1; n <= 20; n++) await p.mouse.move(sx + n * 10, sy + Math.sin(n / 3) * 30);
  await p.mouse.up();
  items = await B();
  check('the pen draws', items.length === 1 && items[0].type === 'ink' && items[0].d.length > 20);
  await p.click('[data-tool="eraser"]');
  await p.mouse.move(sx + 100, sy - 60); await p.mouse.down(); await p.mouse.move(sx + 100, sy + 60, { steps: 12 }); await p.mouse.up();
  check('the eraser removes the drawing', (await B()).length === 0);

  // Files: save and open; hostile content is dropped.
  await p.evaluate(() => { AtlasBoard.template('flowchart'); });
  const bytes = await p.evaluate(() => Array.from(AtlasBoard.serialize()));
  const n0 = (await B()).length;
  await p.evaluate(b => AtlasBoard.loadFile({ name: 'Plan.aboard', data: new Uint8Array(b) }), bytes);
  check('a saved board opens the same', (await B()).length === n0 && (await p.inputValue('.docname input')) === 'Plan.aboard');
  const clean = await p.evaluate(() => AtlasBoard.cleanBoard({ items: [{ type: 'script', x: 0 }, { type: 'image', src: 'javascript:alert(1)' }, { type: 'image', src: 'https://evil.example/x.png' }, { type: 'ink', d: 'M0 0 L10 10"/><script>' }, { type: 'sticky', text: 'ok', color: 'red;background:url(x)' }, { type: 'connector', from: { id: 'nope' } }] }));
  check('a board file can hold only plain items (no scripts, remote pictures or odd colours)', clean.items.length === 2 && clean.items[0].color === '#fff176' && !clean.items[1].from, JSON.stringify(clean.items));

  // Export: SVG with real text, PNG.
  const svgText = await p.evaluate(() => AtlasBoard.exportSvg(false).text);
  check('SVG export draws the words as text, with nothing of the page', /<text[^>]*>.*Did it work\?/s.test(svgText) && !/foreignObject|<script|class="handle"/.test(svgText));
  await p.evaluate(() => { window.__png = null; Platform.save = async o => { window.__png = Array.from(o.data.slice(0, 8)); return { name: o.name }; }; AtlasBoard.select([]); Platform.emit('export-png'); });
  await p.waitForFunction(() => window.__png, null, { timeout: 10000 });
  check('Export Picture makes a PNG', (await p.evaluate(() => window.__png)).join(',') === '137,80,78,71,13,10,26,10');

  // Copy and paste.
  const first = (await B()).find(i => i.type === 'shape');
  await p.evaluate(id => AtlasBoard.select([id]), first.id);
  await p.focus('#board');
  await p.keyboard.press('Control+c');
  await p.keyboard.press('Control+v');
  await p.waitForTimeout(200);
  check('copy and paste make a copy', (await B()).filter(i => i.type === 'shape' && i.text === first.text).length === 2);

  // Zoom with Ctrl + wheel.
  const k0 = await p.evaluate(() => AtlasBoard.view.k);
  await p.mouse.move(700, 500);
  await p.keyboard.down('Control'); await p.mouse.wheel(0, -300); await p.keyboard.up('Control');
  check('Ctrl + wheel zooms in', (await p.evaluate(() => AtlasBoard.view.k)) > k0 * 1.5);
  check('no page errors', errors.length === 0, errors.join(' | '));
  await b.close();
})().catch(e => { console.log('FAIL crashed :: ' + e.message); process.exit(1); });
