'use strict';
// Isadek Remote: a phone moves a slide show on another screen. The link's #fragment holds the
// phone's token (never sent to the server except as the Authorization header).
(function () {
  const $ = id => document.getElementById(id);
  const id = new URLSearchParams(location.search).get('id') || '';
  const tok = location.hash.slice(1).replace(/^t=/, '');
  let v = 0, started = 0, stopped = false;
  const status = t => { $('status').textContent = t; };
  const auth = { Authorization: 'Bearer ' + tok };
  if (!/^[A-Za-z0-9_-]{22}$/.test(id) || !/^[A-Za-z0-9_-]{20,100}$/.test(tok)) { $('count').textContent = 'This link doesn’t control a show'; document.body.classList.add('off'); return; }

  async function send(cmd) {
    if (navigator.vibrate) navigator.vibrate(15);
    try {
      const r = await fetch(`../../api/remote/${id}/command`, { method: 'POST', headers: { ...auth, 'Content-Type': 'application/json' }, body: JSON.stringify({ cmd }), credentials: 'omit', cache: 'no-store' });
      if (!r.ok) status((await r.json().catch(() => ({}))).error || 'Not sent');
      else status('');
    } catch { status('Not connected. Check the Wi-Fi.'); }
  }
  [['nextBtn', 'next'], ['prev', 'prev'], ['first', 'first'], ['last', 'last'], ['black', 'black']].forEach(([el, cmd]) => $(el).addEventListener('click', () => send(cmd)));
  // Volume keys can't be read by a page; arrow keys (a keyboard or clicker) work.
  document.addEventListener('keydown', e => { if (e.key === 'ArrowRight' || e.key === ' ') send('next'); else if (e.key === 'ArrowLeft') send('prev'); });
  // Swipes: left for next, right for back.
  let x0 = null;
  $('notes').parentElement.addEventListener('touchstart', e => { x0 = e.touches[0].clientX; }, { passive: true });
  $('notes').parentElement.addEventListener('touchend', e => { if (x0 === null) return; const dx = e.changedTouches[0].clientX - x0; if (Math.abs(dx) > 80) send(dx < 0 ? 'next' : 'prev'); x0 = null; });

  function show(st) {
    v = st.v;
    started = st.started || 0;
    document.body.classList.toggle('off', !st.showing);
    $('count').textContent = st.showing ? `Slide ${st.slide + 1} of ${st.total}` : 'The show hasn’t started';
    $('title').textContent = st.title || (st.showing ? `Slide ${st.slide + 1}` : '');
    $('next').textContent = st.next ? 'Next: ' + st.next : st.showing && st.slide + 1 >= st.total ? 'Last slide' : '';
    $('notes').textContent = st.notes || '';
    $('black').setAttribute('aria-pressed', String(!!st.black));
  }
  (async function follow() {
    while (!stopped) {
      try {
        const r = await fetch(`../../api/remote/${id}/state?since=${v}`, { headers: auth, credentials: 'omit', cache: 'no-store' });
        if (r.status === 404 || r.status === 403) { $('count').textContent = 'The show has ended'; document.body.classList.add('off'); stopped = true; return; }
        if (r.ok) { show(await r.json()); status(''); } else await new Promise(res => setTimeout(res, 2000));
      } catch { status('Not connected. Check the Wi-Fi.'); await new Promise(res => setTimeout(res, 2000)); }
    }
  })();
  setInterval(() => {
    if (!started) { $('clock').textContent = ''; return; }
    const s = Math.max(0, Math.floor((Date.now() - started) / 1000));
    $('clock').textContent = `${Math.floor(s / 60)}:${String(s % 60).padStart(2, '0')}`;
  }, 1000);
  // Keep the phone's screen on while presenting.
  const awake = () => { if (navigator.wakeLock && document.visibilityState === 'visible') navigator.wakeLock.request('screen').catch(() => {}); };
  document.addEventListener('visibilitychange', awake);
  awake();
})();
