'use strict';
// Settings built into a release of the desktop app (src/main/release.json). From the environment:
//
//   ISADEK_CRASH_URL   where crash reports go when people turn them on (an https address that
//                      accepts them, such as the web server's /api/crash with CRASH_DIR set)
//
//   node scripts/release-config.js     writes the file (an empty {} when nothing is set)
const fs = require('fs');
const path = require('path');

function fromEnv(env = process.env) {
  const out = {};
  const crash = String(env.ISADEK_CRASH_URL || '').trim();
  if (/^https:\/\//i.test(crash)) out.crashUrl = crash;
  return out;
}

module.exports = { fromEnv };

if (require.main === module) {
  const cfg = fromEnv();
  fs.writeFileSync(path.resolve(__dirname, '../src/main/release.json'), JSON.stringify(cfg, null, 2) + '\n');
  console.log('Crash reports: ' + (cfg.crashUrl ? 'sent to ' + new URL(cfg.crashUrl).host + ' when turned on' : 'kept on the computer only'));
}
