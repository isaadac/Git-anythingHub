// The download website: the right installer for each computer, from the newest release; Help and
// Privacy pages built from the app's Help.
const { chromium } = require('playwright');
const path = require('path');
const { build } = require('../../scripts/build-site');
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };

const out = build(path.join(process.env.E2E_WORK, 'site'));
const v = '1.2.0', dl = n => `https://github.com/isaadac/Git-anythingHub/releases/download/office-v${v}/${n}`;
const names = [`Isadek-Office-Setup-${v}-x64.exe`, `Isadek-Office-Setup-${v}-arm64.exe`, `Isadek-Office-${v}-portable.exe`, `Isadek-Office-${v}-universal.dmg`, `Isadek-Office-${v}-universal.zip`,
  `Isadek-Office-${v}-x86_64.AppImage`, `Isadek-Office-${v}-arm64.AppImage`, `Isadek-Office-${v}-amd64.deb`, `Isadek-Office-${v}-arm64.deb`, `Isadek-Office-${v}-android-release.apk`, 'SHA256SUMS.txt', 'latest.yml'];
const releases = [
  { tag_name: 'other-v9', draft: false, prerelease: false, assets: [] },
  { tag_name: 'office-v1.3.0-beta', draft: false, prerelease: true, assets: [] },
  { tag_name: 'office-v' + v, draft: false, prerelease: false, published_at: '2026-10-20T10:00:00Z', assets: names.map((n, i) => ({ name: n, size: (90 + i) * 1048576, browser_download_url: dl(n) })) }
];
const UA = {
  windows: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0 Safari/537.36',
  mac: 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0 Safari/537.36',
  linux: 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0 Safari/537.36',
  android: 'Mozilla/5.0 (Linux; Android 15; Pixel 9) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0 Mobile Safari/537.36'
};

(async () => {
  const browser = await chromium.launch();
  const errors = [];
  const open = async (ua, api = releases) => {
    const ctx = await browser.newContext({ userAgent: ua });
    const p = await ctx.newPage();
    p.on('pageerror', e => errors.push(e.message));
    p.on('console', m => { if (m.type() === 'error' && !/Failed to load resource/.test(m.text())) errors.push(m.text()); });
    await ctx.route('https://api.github.com/**', r => (api ? r.fulfill({ json: api }) : r.fulfill({ status: 503, body: '' })));
    await p.goto('file://' + path.join(out, 'index.html'));
    await p.waitForFunction(() => document.getElementById('version').textContent !== '' || document.querySelector('.os li'), null, { timeout: 5000 }).catch(() => {});
    return p;
  };
  const main = p => p.$eval('#mainBtn', a => [a.textContent, a.href]);

  for (const [os, want, text] of [['windows', names[0], 'Download for Windows'], ['mac', names[3], 'Download for macOS'], ['linux', names[5], 'Download for Linux'], ['android', names[9], 'Download for Android']]) {
    const p = await open(UA[os]);
    const [t, href] = await main(p);
    check(`${os}: the big button downloads ${want}`, t === text && href === dl(want), `${t} ${href}`);
    if (os === 'windows') {
      const lists = await p.$$eval('.os', o => o.map(x => [x.dataset.os, x.querySelectorAll('li a').length, x.classList.contains('mine')]));
      check('every installer is listed under its system, this one marked', JSON.stringify(lists) === JSON.stringify([['windows', 3, true], ['mac', 2, false], ['linux', 4, false], ['android', 1, false]]), JSON.stringify(lists));
      check('the newest final release is the one shown (not betas or other projects)', (await p.$eval('#version', e => e.textContent)).startsWith('Version 1.2.0'));
    }
    await p.context().close();
  }
  const off = await open(UA.windows, null);
  check('without GitHub the button still leads to the releases page', (await main(off))[1] === 'https://github.com/isaadac/Git-anythingHub/releases');
  await off.context().close();

  const ctx = await browser.newContext();
  const p = await ctx.newPage();
  p.on('pageerror', e => errors.push(e.message));
  await p.goto('file://' + path.join(out, 'privacy.html'));
  const pv = await p.evaluate(() => ({ h: document.querySelector('h2').textContent, crash: /Crash reports/.test(document.body.textContent), setting: !!document.querySelector('.setting') }));
  check('privacy.html is the policy from the app', pv.h === 'Privacy' && pv.crash && !pv.setting, JSON.stringify(pv));
  await p.goto('file://' + path.join(out, 'help.html'));
  check('help.html has every Help topic', (await p.$$eval('main section', s => s.length)) >= 10);
  check('no page errors', !errors.length, errors.join(' | '));
  await browser.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
