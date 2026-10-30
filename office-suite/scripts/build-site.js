'use strict';
// Builds the download website into dist-site/ (published to GitHub Pages by the release workflow):
// the download page from site/, and help.html and privacy.html made from the app's own Help
// (src/renderer/help.html), so the policy on the web is the one in the app.
//
//   node scripts/build-site.js [outDir]
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const OUT = path.resolve(process.argv[2] || path.join(ROOT, 'dist-site'));

function sections(html) {
  return [...html.matchAll(/<section id="([\w-]+)"[^>]*>([\s\S]*?)<\/section>/g)].map(m => ({ id: m[1], html: m[2].replace(/<div class="setting"[\s\S]*?<\/div>\n?/g, '') }));
}
const page = (title, body) => `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta http-equiv="Content-Security-Policy" content="default-src 'self'; img-src 'self' data:; style-src 'self'; object-src 'none'; base-uri 'none'; form-action 'none'">
<title>${title}</title>
<link rel="icon" href="icon.svg">
<link rel="stylesheet" href="site.css">
</head>
<body>
<header class="top">
  <div class="brand"><img src="icon.svg" alt="" width="28" height="28"> <a class="home" href="index.html">Isadek Office</a></div>
  <nav><a href="index.html#download">Download</a><a href="help.html">Help</a><a href="privacy.html">Privacy</a></nav>
</header>
<main class="doc">
${body}
</main>
<footer class="muted small">Isadek Office · <a href="privacy.html">Privacy</a> · <a href="help.html">Help</a></footer>
</body>
</html>
`;

function build(out = OUT) {
  fs.rmSync(out, { recursive: true, force: true });
  fs.mkdirSync(out, { recursive: true });
  for (const f of fs.readdirSync(path.join(ROOT, 'site'))) fs.copyFileSync(path.join(ROOT, 'site', f), path.join(out, f));
  fs.copyFileSync(path.join(ROOT, 'build/icon.svg'), path.join(out, 'icon.svg'));
  const all = sections(fs.readFileSync(path.join(ROOT, 'src/renderer/help.html'), 'utf8'));
  const privacy = all.find(s => s.id === 'privacy');
  if (!privacy) throw new Error('help.html has no privacy section');
  fs.writeFileSync(path.join(out, 'privacy.html'), page('Privacy — Isadek Office', privacy.html.trim()));
  const toc = '<p>' + all.map(s => `<a href="#${s.id}">${s.html.match(/<h2>([^<]*)<\/h2>/)[1]}</a>`).join(' · ') + '</p>';
  fs.writeFileSync(path.join(out, 'help.html'), page('Help — Isadek Office', '<h1>Isadek Office Help</h1>\n' + toc + all.map(s => `<section id="${s.id}">${s.html}</section>`).join('\n')));
  // GitHub Pages: serve files as they are.
  fs.writeFileSync(path.join(out, '.nojekyll'), '');
  return out;
}

module.exports = { build, sections };
if (require.main === module) console.log('Site built in ' + path.relative(process.cwd(), build()));
