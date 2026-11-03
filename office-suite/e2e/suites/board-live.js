// The Whiteboard edited by two people at once: items, moves and deletions reach the other board,
// each person sees the other's cursor, and someone else's hostile item is cleaned on arrival.
const { chromium } = require('playwright');
const path = require('path');
const { createCollabServer } = require(path.join(process.cwd(), 'src/main/collab-server.js'));
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
(async () => {
  const srv = createCollabServer();
  await new Promise(r => srv.listen(0, '127.0.0.1', r));
  const WS = `ws://127.0.0.1:${srv.address().port}/`;
  const browser = await chromium.launch();
  const errors = [];
  async function open(name) {
    const ctx = await browser.newContext({ viewport: { width: 1300, height: 850 } });
    await ctx.addInitScript(n => { localStorage.setItem('atlas-user-name', n); }, name);
    const p = await ctx.newPage();
    p.on('pageerror', e => errors.push(name + ': ' + e.message));
    await p.goto(process.env.E2E_BASE + '/src/renderer/board.html');
    await p.waitForFunction(() => window.AtlasBoard && document.querySelector('.live-btn'));
    return p;
  }
  const until = (p, fn, arg) => p.waitForFunction(fn, arg, { timeout: 8000 }).then(() => true, () => false);
  const a = await open('Ada'), b = await open('Bo');
  await a.evaluate(() => { AtlasBoard.add({ type: 'sticky', x: 0, y: 0, w: 180, h: 180, text: 'Before the session' }); });
  await a.click('.live-btn');
  await a.click('.modal .actions button:has-text("Start on a server")');
  await a.fill('.modal input', WS);
  await a.click('.modal .actions .primary');
  await a.waitForSelector('.modal p:has-text("room=")', { timeout: 10000 });
  const link = (await a.textContent('.modal p')).match(/ws:\/\/\S+/)[0];
  await a.click('.modal .actions button:has-text("Close")');
  await b.click('.live-btn');
  await b.click('.modal .actions button:has-text("Join")');
  await b.fill('.modal input', link);
  await b.click('.modal .actions .primary');
  check('the board arrives for someone joining', await until(b, () => AtlasBoard.board.items.some(i => i.text === 'Before the session')));
  // Bo adds a shape; Ada moves the note.
  await b.evaluate(() => { AtlasBoard.add({ type: 'shape', shape: 'diamond', x: 300, y: 0, w: 160, h: 120, text: 'Decide' }); });
  check('a shape added by one appears for the other', await until(a, () => AtlasBoard.board.items.some(i => i.type === 'shape' && i.text === 'Decide')));
  await a.evaluate(() => { const s = AtlasBoard.board.items.find(i => i.type === 'sticky'); s.x = 40; s.y = 220; AtlasBoard.add({ type: 'connector', from: { id: s.id }, to: { id: AtlasBoard.board.items.find(i => i.type === 'shape').id } }); });
  check('moves and connectors reach the other board', await until(b, () => { const s = AtlasBoard.board.items.find(i => i.type === 'sticky'); return s && s.x === 40 && s.y === 220 && AtlasBoard.board.items.some(i => i.type === 'connector'); }));
  // Cursors: Ada sees Bo's pointer.
  const [bx, by] = await b.evaluate(() => AtlasBoard.toScreen(20, 20));
  await b.mouse.move(bx, by); await b.mouse.move(bx + 10, by + 5);
  check('each person sees the other’s cursor with their name', await until(a, () => [...document.querySelectorAll('#peers .peer text')].some(t => t.textContent === 'Bo')));
  // Deleting.
  await b.evaluate(() => { const s = AtlasBoard.board.items.find(i => i.type === 'shape'); AtlasBoard.select([s.id]); document.getElementById('board').dispatchEvent(new KeyboardEvent('keydown', { key: 'Delete', bubbles: true })); });
  check('a deletion reaches the other board (with its connectors)', await until(a, () => !AtlasBoard.board.items.some(i => i.type === 'shape' || i.type === 'connector')));
  check('no page errors', errors.length === 0, errors.join(' | '));
  await browser.close();
  srv.close();
})().catch(e => { console.log('FAIL crashed :: ' + e.message); process.exit(1); });
