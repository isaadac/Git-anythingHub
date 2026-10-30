const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch(); const p = await b.newPage();
  await p.goto(process.env.E2E_BASE + '/src/renderer/write.html');
  await p.evaluate(() => { localStorage.setItem('atlas-author', 'Ada'); document.getElementById('doc').innerHTML = '<p>Hello world</p>'; });
  await p.click('[data-tab="review"]'); await p.click('#trackChanges');
  await p.evaluate(() => { const t = document.querySelector('#doc p').firstChild; const r = document.createRange(); r.setStart(t, 5); r.collapse(true); getSelection().removeAllRanges(); getSelection().addRange(r); document.getElementById('doc').focus(); });
  await p.keyboard.type(' there');
  const before = await p.$eval('#doc', d => d.innerHTML);
  const steps = [];
  for (let i = 0; i < 4; i++) { await p.keyboard.press('Control+z'); steps.push(await p.$eval('#doc', d => d.textContent.replace(/\u00a0/g, ' '))); }
  console.log(JSON.stringify(steps));
  const after = steps[steps.length - 1];
  console.log((/<ins/.test(before) && after.replace(/ /g, ' ') === 'Hello world' ? 'PASS' : 'FAIL') + ' undo removes the tracked insertion :: ' + JSON.stringify(after));
  await b.close();
})();
