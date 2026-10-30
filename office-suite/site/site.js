'use strict';
// The download page: finds the newest Isadek Office release on GitHub and links its installers,
// with the one for this computer as the big button.
(function () {
  const REPO = 'isaadac/Git-anythingHub';
  const KINDS = [
    { os: 'windows', re: /Setup-.*-x64\.exe$/, label: 'Installer (64-bit)', main: 'Download for Windows' },
    { os: 'windows', re: /Setup-.*-arm64\.exe$/, label: 'Installer (ARM)' },
    { os: 'windows', re: /portable\.exe$/, label: 'Portable (no install)' },
    { os: 'mac', re: /\.dmg$/, label: 'Disk image (Apple silicon and Intel)', main: 'Download for macOS' },
    { os: 'mac', re: /mac\.zip$|universal.*\.zip$/, label: 'Zip' },
    { os: 'linux', re: /x86_64\.AppImage$/, label: 'AppImage (64-bit)', main: 'Download for Linux' },
    { os: 'linux', re: /arm64\.AppImage$/, label: 'AppImage (ARM)' },
    { os: 'linux', re: /amd64\.deb$/, label: 'Debian/Ubuntu (.deb, 64-bit)' },
    { os: 'linux', re: /arm64\.deb$/, label: 'Debian/Ubuntu (.deb, ARM)' },
    { os: 'android', re: /release\.apk$/, label: 'APK', main: 'Download for Android' }
  ];
  const ua = navigator.userAgent;
  const mine = /Android/i.test(ua) ? 'android' : /Windows/i.test(ua) ? 'windows' : /Mac OS X|Macintosh/i.test(ua) && !/iPhone|iPad/i.test(ua) ? 'mac' : /Linux|X11|CrOS/i.test(ua) ? 'linux' : null;
  const size = n => (n > 1048576 ? (n / 1048576).toFixed(0) + ' MB' : (n / 1024).toFixed(0) + ' KB');

  function show(release) {
    const assets = release.assets || [];
    let main = null;
    for (const k of KINDS) {
      const a = assets.find(x => k.re.test(x.name));
      if (!a) continue;
      const li = document.createElement('li');
      const link = document.createElement('a');
      link.href = a.browser_download_url;
      link.textContent = k.label;
      li.append(link, ' ');
      const s = document.createElement('small');
      s.textContent = size(a.size);
      li.appendChild(s);
      document.querySelector(`.os[data-os="${k.os}"] ul`).appendChild(li);
      if (k.os === mine && k.main && !main) main = { a, k };
    }
    document.querySelectorAll('.os').forEach(o => { if (!o.querySelector('li')) o.querySelector('ul').innerHTML = '<li class="muted">Coming soon</li>'; o.classList.toggle('mine', o.dataset.os === mine); });
    document.getElementById('version').textContent = `Version ${String(release.tag_name).replace(/^office-v/, '')}${release.published_at ? ' · ' + new Date(release.published_at).toLocaleDateString() : ''}`;
    if (main) {
      const b = document.getElementById('mainBtn');
      b.href = main.a.browser_download_url;
      b.textContent = main.k.main;
      document.getElementById('mainNote').textContent = `${main.k.label} · ${size(main.a.size)} · other systems below`;
    } else {
      document.getElementById('mainBtn').href = '#download';
    }
  }

  fetch(`https://api.github.com/repos/${REPO}/releases?per_page=20`, { headers: { Accept: 'application/vnd.github+json' } })
    .then(r => (r.ok ? r.json() : Promise.reject(new Error(r.status))))
    .then(list => {
      const release = (list || []).find(r => /^office-v/.test(r.tag_name) && !r.draft && !r.prerelease);
      if (release) show(release); else document.getElementById('version').textContent = 'The first release is on its way.';
    })
    .catch(() => { document.getElementById('version').textContent = ''; });
})();
