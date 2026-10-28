'use strict';
// The Help window: contents, search, and the crash-report setting (desktop).
(function () {
  const main = document.getElementById('main'), toc = document.getElementById('toc'), q = document.getElementById('q');
  const sections = [...main.querySelectorAll('section')];
  sections.forEach(s => {
    const a = document.createElement('a');
    a.href = '#' + s.id;
    a.textContent = s.querySelector('h2').textContent;
    toc.appendChild(a);
  });
  const mark = () => {
    const id = (location.hash || '#start').slice(1);
    toc.querySelectorAll('a').forEach(a => a.classList.toggle('on', a.hash === '#' + id));
  };
  window.addEventListener('hashchange', mark);
  mark();

  q.addEventListener('input', () => {
    const words = q.value.toLowerCase().split(/\s+/).filter(Boolean);
    let shown = 0;
    sections.forEach(s => {
      const hit = words.every(w => s.textContent.toLowerCase().includes(w));
      s.classList.toggle('hidden', !hit);
      toc.querySelector(`a[href="#${s.id}"]`).classList.toggle('hidden', !hit);
      if (hit) shown++;
    });
    document.getElementById('none').classList.toggle('hidden', shown > 0);
  });
  q.addEventListener('keydown', e => { if (e.key === 'Escape') { q.value = ''; q.dispatchEvent(new Event('input')); } });

  // Desktop: the crash-report setting, here as well as in the Help menu.
  const atlas = window.atlas;
  if (atlas && atlas.crashStatus) {
    const box = document.getElementById('crashSetting'), t = document.getElementById('crashToggle'), note = document.getElementById('crashNote');
    atlas.crashStatus().then(st => {
      box.hidden = false;
      t.checked = st.enabled;
      note.textContent = st.url ? '' : '(this copy of Isadek Office keeps reports on your computer only)';
    }).catch(() => {});
    t.addEventListener('change', () => atlas.setCrashReports(t.checked));
  }
})();
