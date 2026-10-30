'use strict';
// Isadek Forms — the form model shared by Sheets (which builds forms), the page people answer
// them on, and the server that collects the answers. Everything that crosses from one to another
// goes through clean() and check(): plain data with known fields and limits, nothing else.
(function (root) {
  const TYPES = {
    text: 'Short answer', paragraph: 'Paragraph', number: 'Number', email: 'Email', date: 'Date',
    choice: 'Multiple choice', checkboxes: 'Checkboxes', dropdown: 'Dropdown', rating: 'Rating'
  };
  const WITH_OPTIONS = new Set(['choice', 'checkboxes', 'dropdown']);
  const LIMITS = { questions: 200, options: 200, title: 300, help: 2000, description: 5000, option: 300, answer: 10000, paragraph: 50000 };
  const ID = /^[A-Za-z0-9_-]{1,40}$/;
  const COLOR = /^#[0-9a-f]{6}$/i;
  const EMAIL = /^[^\s@<>()[\]\\,;:"]{1,64}@[^\s@<>()[\]\\,;:"]{1,255}\.[A-Za-z]{2,63}$/;
  const str = (v, n) => (typeof v === 'string' ? v.replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '').slice(0, n) : '');
  const finite = v => (typeof v === 'number' && isFinite(v) ? v : typeof v === 'string' && v.trim() !== '' && isFinite(+v) ? +v : null);

  /** A question with known fields only. */
  function cleanQuestion(q, i) {
    if (!q || typeof q !== 'object' || !TYPES[q.type]) return null;
    const out = { id: typeof q.id === 'string' && ID.test(q.id) ? q.id : 'q' + (i + 1), type: q.type, title: str(q.title, LIMITS.title) || 'Question', required: q.required === true };
    const help = str(q.help, LIMITS.help);
    if (help) out.help = help;
    if (WITH_OPTIONS.has(q.type)) {
      const seen = new Set();
      out.options = (Array.isArray(q.options) ? q.options : []).map(o => str(o, LIMITS.option).trim()).filter(o => o && !seen.has(o) && seen.add(o)).slice(0, LIMITS.options);
      if (!out.options.length) out.options = ['Option 1'];
    }
    if (q.type === 'number') {
      const lo = finite(q.min), hi = finite(q.max);
      if (lo !== null) out.min = lo;
      if (hi !== null && (lo === null || hi >= lo)) out.max = hi;
      if (q.whole === true) out.whole = true;
    }
    if (q.type === 'rating') out.max = Math.min(10, Math.max(3, Math.round(finite(q.max) || 5)));
    return out;
  }

  /** A form definition: title, description, questions and how it looks. Ids stay unique. */
  function clean(f) {
    if (!f || typeof f !== 'object') return null;
    const used = new Set();
    const questions = (Array.isArray(f.questions) ? f.questions : []).slice(0, LIMITS.questions).map(cleanQuestion).filter(Boolean).map((q, i) => {
      let id = q.id;
      for (let k = 2; used.has(id); k++) id = q.id.slice(0, 36) + '_' + k;
      used.add(id);
      return { ...q, id };
    });
    const out = { title: str(f.title, LIMITS.title) || 'Untitled form', description: str(f.description, LIMITS.description), questions };
    if (COLOR.test(f.color)) out.color = f.color.toLowerCase();
    const confirm = str(f.confirm, LIMITS.help);
    if (confirm) out.confirm = confirm;
    return out;
  }

  /**
   * Checks answers against the form: { ok, errors: {questionId: message}, values: {questionId: value} }.
   * Values are numbers for Number and Rating, arrays of chosen options for Checkboxes, otherwise text.
   */
  function check(form, answers) {
    const errors = {}, values = {};
    const a = answers && typeof answers === 'object' && !Array.isArray(answers) ? answers : {};
    for (const q of form.questions) {
      const raw = Object.prototype.hasOwnProperty.call(a, q.id) ? a[q.id] : undefined;
      const empty = raw === undefined || raw === null || raw === '' || (Array.isArray(raw) && !raw.length);
      if (empty) { if (q.required) errors[q.id] = 'This question needs an answer'; continue; }
      let v = null, err = '';
      switch (q.type) {
        case 'text': case 'paragraph': {
          if (typeof raw !== 'string') { err = 'Enter text'; break; }
          v = raw.slice(0, q.type === 'paragraph' ? LIMITS.paragraph : LIMITS.answer);
          if (q.type === 'text') v = v.replace(/[\r\n]+/g, ' ');
          if (!v.trim()) v = null;
          break;
        }
        case 'email': {
          const t = typeof raw === 'string' ? raw.trim() : '';
          if (!EMAIL.test(t)) err = 'Enter an email address, like name@example.com'; else v = t;
          break;
        }
        case 'number': {
          const n = finite(raw);
          if (n === null) err = 'Enter a number';
          else if (q.whole && !Number.isInteger(n)) err = 'Enter a whole number';
          else if (q.min !== undefined && n < q.min) err = `Enter ${q.min} or more`;
          else if (q.max !== undefined && n > q.max) err = `Enter ${q.max} or less`;
          else v = n;
          break;
        }
        case 'date': {
          const t = typeof raw === 'string' ? raw.trim() : '';
          const m = t.match(/^(\d{4})-(\d{2})-(\d{2})$/);
          const d = m && new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
          if (!d || d.getUTCMonth() !== +m[2] - 1 || +m[1] < 1900 || +m[1] > 9999) err = 'Enter a date'; else v = t;
          break;
        }
        case 'rating': {
          const n = finite(raw);
          if (n === null || !Number.isInteger(n) || n < 1 || n > q.max) err = `Choose 1 to ${q.max}`; else v = n;
          break;
        }
        case 'choice': case 'dropdown': {
          if (typeof raw !== 'string' || !q.options.includes(raw)) err = 'Choose one of the options'; else v = raw;
          break;
        }
        case 'checkboxes': {
          const list = Array.isArray(raw) ? raw : [raw];
          if (list.some(x => typeof x !== 'string' || !q.options.includes(x))) err = 'Choose from the options';
          else v = q.options.filter(o => list.includes(o)); // in the form's order, once each
          break;
        }
      }
      if (err) errors[q.id] = err;
      else if (v !== null) values[q.id] = v;
      else if (q.required) errors[q.id] = 'This question needs an answer';
    }
    return { ok: !Object.keys(errors).length, errors, values };
  }

  /**
   * The cell values for one response, in the order of `columns` (question ids): text that a
   * spreadsheet would read as a formula or a number is kept as text, so an answer can never
   * become a formula (the leading ' is the spreadsheet's own "this is text" mark).
   */
  function toCells(values, columns) {
    return columns.map(id => {
      const v = values ? values[id] : undefined;
      if (v === undefined || v === null) return '';
      if (typeof v === 'number') return v;
      const t = Array.isArray(v) ? v.join(', ') : String(v);
      return /^\s*[=+\-@'#$.\d]/.test(t) || /^[\t\r]/.test(t) || /^\s*(true|false)\s*$/i.test(t) ? "'" + t : t;
    });
  }

  const api = { TYPES, WITH_OPTIONS, LIMITS, clean, cleanQuestion, check, toCells, ID };
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.FormsCore = api;
})(typeof self !== 'undefined' ? self : this);
