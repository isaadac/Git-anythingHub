'use strict';
// The page people answer an Isadek form on (/f/<id> brings them here): loads the form from the
// server that published it, and sends the answers back to the same server.
(function () {
  const box = document.getElementById('box');
  const id = new URLSearchParams(location.search).get('id') || '';
  const message = (title, text) => {
    box.textContent = '';
    const card = document.createElement('div'); card.className = 'fv-body fv-done';
    const h = document.createElement('h1'); h.className = 'fv-form-title'; h.textContent = title;
    const p = document.createElement('p'); p.className = 'fv-desc'; p.textContent = text;
    card.append(h, p); box.appendChild(card);
  };
  async function call(method, url, body) {
    const res = await fetch(url, { method, headers: body ? { 'Content-Type': 'application/json' } : {}, body: body ? JSON.stringify(body) : undefined, credentials: 'omit', cache: 'no-store' });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw Object.assign(new Error(data.error || `The server answered ${res.status}`), { errors: data.errors, status: res.status });
    return data;
  }
  (async () => {
    if (!/^[A-Za-z0-9_-]{22}$/.test(id)) return message('Form not found', 'Check the link you were given.');
    let got;
    try { got = await call('GET', `../../api/forms/${id}`); } catch (err) { return message('Form not found', err.status === 404 ? 'This form doesn’t exist, or was deleted.' : 'The form couldn’t be loaded. Check your connection and try again.'); }
    document.title = got.form.title + ' — Isadek Forms';
    if (!got.open) return message(got.form.title, 'This form is no longer accepting responses.');
    FormsView.render(box, got.form, { onSubmit: answers => call('POST', `../../api/forms/${id}/responses`, { answers }) });
  })();
})();
