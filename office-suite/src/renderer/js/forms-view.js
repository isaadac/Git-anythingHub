'use strict';
// Isadek Forms — draws a form for someone to answer: on the answer page, and in Sheets (preview,
// and filling one in on this device). Everything from the form is set as text, never as markup.
//
//   FormsView.render(box, form, { onSubmit: async answers => {}, submitLabel, preview })
// onSubmit may throw an Error with .errors ({questionId: message}) to show beside the questions.
(function () {
  const C = window.FormsCore;
  const el = (tag, cls, text) => { const e = document.createElement(tag); if (cls) e.className = cls; if (text !== undefined) e.textContent = text; return e; };
  let seq = 0;

  function question(q, n) {
    const id = `fv${++seq}`;
    const card = el('fieldset', 'fv-q');
    card.dataset.q = q.id;
    card.dataset.type = q.type;
    const legend = el('legend', 'fv-title', q.title);
    if (q.required) { const star = el('span', 'fv-req', ' *'); star.setAttribute('aria-hidden', 'true'); legend.appendChild(star); }
    card.appendChild(legend);
    if (q.help) card.appendChild(el('p', 'fv-help', q.help));
    const err = el('p', 'fv-err');
    err.id = id + 'e';
    err.setAttribute('role', 'alert');
    const describe = input => { input.setAttribute('aria-describedby', err.id); if (q.required) input.setAttribute('aria-required', 'true'); input.setAttribute('aria-label', q.title); };
    let get;
    if (q.type === 'paragraph') {
      const t = el('textarea', 'fv-input'); t.rows = 4; t.maxLength = C.LIMITS.paragraph; describe(t); card.appendChild(t);
      get = () => t.value;
    } else if (['text', 'number', 'email', 'date'].includes(q.type)) {
      const i = el('input', 'fv-input');
      i.type = q.type === 'text' ? 'text' : q.type;
      if (q.type === 'number') { i.step = q.whole ? '1' : 'any'; if (q.min !== undefined) i.min = q.min; if (q.max !== undefined) i.max = q.max; i.inputMode = q.whole ? 'numeric' : 'decimal'; }
      if (q.type === 'email') i.autocomplete = 'email';
      if (q.type === 'text') i.maxLength = C.LIMITS.answer;
      describe(i); card.appendChild(i);
      get = () => (q.type === 'number' ? (i.value === '' ? (i.validity.badInput ? 'x' : '') : i.value) : i.value);
    } else if (q.type === 'dropdown') {
      const s = el('select', 'fv-input');
      s.appendChild(el('option', '', 'Choose')).value = '';
      q.options.forEach(o => { s.appendChild(el('option', '', o)).value = o; });
      describe(s); card.appendChild(s);
      get = () => s.value;
    } else if (q.type === 'choice' || q.type === 'checkboxes') {
      const list = el('div', 'fv-options');
      q.options.forEach((o, k) => {
        const lab = el('label', 'fv-option');
        const i = el('input');
        i.type = q.type === 'choice' ? 'radio' : 'checkbox';
        i.name = id; i.value = o; i.id = `${id}o${k}`;
        i.setAttribute('aria-describedby', err.id);
        lab.append(i, el('span', '', o));
        list.appendChild(lab);
      });
      card.appendChild(list);
      get = () => { const on = [...list.querySelectorAll('input:checked')].map(i => i.value); return q.type === 'choice' ? (on[0] || '') : on; };
    } else if (q.type === 'rating') {
      const row = el('div', 'fv-rating');
      row.setAttribute('role', 'radiogroup');
      row.setAttribute('aria-label', q.title);
      for (let k = 1; k <= q.max; k++) {
        const lab = el('label', 'fv-star');
        const i = el('input'); i.type = 'radio'; i.name = id; i.value = String(k);
        i.setAttribute('aria-label', `${k} of ${q.max}`);
        lab.append(i, el('span', '', String(k)));
        row.appendChild(lab);
      }
      row.addEventListener('change', () => { const v = +(row.querySelector('input:checked') || {}).value || 0; row.querySelectorAll('.fv-star').forEach((l, k) => l.classList.toggle('on', k < v)); });
      card.appendChild(row);
      get = () => { const i = row.querySelector('input:checked'); return i ? +i.value : ''; };
    }
    card.appendChild(err);
    return { card, get, err, q, n };
  }

  function render(box, raw, opts = {}) {
    const form = C.clean(raw);
    box.textContent = '';
    box.classList.add('fv');
    if (form.color) box.style.setProperty('--fv-accent', form.color);
    const head = el('header', 'fv-head');
    head.appendChild(el('h1', 'fv-form-title', form.title));
    if (form.description) head.appendChild(el('p', 'fv-desc', form.description));
    if (form.questions.some(q => q.required)) head.appendChild(el('p', 'fv-note', '* Required'));
    const f = el('form', 'fv-body');
    f.noValidate = true;
    f.appendChild(head);
    const items = form.questions.map(question);
    items.forEach(it => f.appendChild(it.card));
    const bar = el('div', 'fv-actions');
    const send = el('button', 'fv-submit', opts.submitLabel || 'Submit'); send.type = 'submit';
    const clear = el('button', 'fv-clear', 'Clear form'); clear.type = 'button';
    const status = el('p', 'fv-status'); status.setAttribute('role', 'status');
    bar.append(send, clear);
    f.append(bar, status);
    box.appendChild(f);

    const show = errors => {
      let first = null;
      items.forEach(it => {
        const m = errors[it.q.id] || '';
        it.err.textContent = m;
        it.card.classList.toggle('bad', !!m);
        it.card.querySelectorAll('input, textarea, select').forEach(i => i.setAttribute('aria-invalid', m ? 'true' : 'false'));
        if (m && !first) first = it.card;
      });
      if (first) { first.scrollIntoView({ block: 'center', behavior: 'smooth' }); const i = first.querySelector('input, textarea, select'); if (i) i.focus({ preventScroll: true }); }
    };
    clear.addEventListener('click', () => { f.reset(); f.querySelectorAll('.fv-star.on').forEach(s => s.classList.remove('on')); show({}); status.textContent = ''; });
    f.addEventListener('submit', async e => {
      e.preventDefault();
      if (send.disabled) return;
      const answers = {};
      items.forEach(it => { answers[it.q.id] = it.get(); });
      const r = C.check(form, answers);
      show(r.errors);
      if (!r.ok) { status.textContent = 'Some answers need fixing.'; return; }
      send.disabled = true;
      status.textContent = 'Sending…';
      try {
        await opts.onSubmit(r.values);
        done();
      } catch (err) {
        if (err && err.errors) show(err.errors);
        status.textContent = (err && err.message) || 'Your response couldn’t be sent. Try again.';
      } finally { send.disabled = false; }
    });
    function done() {
      box.textContent = '';
      const card = el('div', 'fv-body fv-done');
      const head2 = el('header', 'fv-head');
      head2.appendChild(el('h1', 'fv-form-title', form.title));
      head2.appendChild(el('p', 'fv-desc', form.confirm || 'Your response has been recorded.'));
      card.appendChild(head2);
      const again = el('button', 'fv-clear', 'Submit another response'); again.type = 'button';
      again.addEventListener('click', () => render(box, raw, opts));
      card.appendChild(again);
      box.appendChild(card);
      if (opts.onDone) opts.onDone();
    }
    return { form, element: f };
  }

  window.FormsView = { render };
})();
