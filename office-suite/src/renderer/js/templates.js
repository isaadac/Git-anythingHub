'use strict';
// Isadek Office — workbook and presentation templates (Write's are in write-templates.js).
// A template is plain data. Opening one hands the app a tiny marker file ({ isadekTemplate, id })
// through the normal open path, and the app builds the workbook or deck from the data here, so it
// opens as a new, unsaved .xlsx or .pptx in a window of its own (desktop) or tab (web).
(function () {
  const H = { b: true, bg: '#1f3864', color: '#ffffff' }; // header row
  const T = { b: true, size: 20, color: '#1f3864' };       // sheet title
  const MONEY = '$#,##0.00';
  const c = (v, s) => ({ v: String(v), s });               // a styled cell
  const money = v => c(v, { fmt: MONEY });
  const total = v => c(v, { b: true, fmt: MONEY, border: { t: 1 } });

  const SHEETS = [
    {
      id: 'budget', name: 'Monthly budget', desc: 'Income, spending and what is left',
      sheet: {
        name: 'Budget', widths: { 0: 190, 1: 110, 2: 110, 3: 110 }, heights: { 0: 42 },
        rows: [
          [c('Monthly budget', T)],
          [],
          [c('Category', H), c('Planned', H), c('Actual', H), c('Difference', H)],
          ['Rent', money(1200), money(1200), money('=B4-C4')],
          ['Groceries', money(450), money(512), money('=B5-C5')],
          ['Transport', money(160), money(138), money('=B6-C6')],
          ['Utilities', money(180), money(171), money('=B7-C7')],
          ['Eating out', money(150), money(205), money('=B8-C8')],
          ['Savings', money(400), money(400), money('=B9-C9')],
          [c('Total', { b: true, border: { t: 1 } }), total('=SUM(B4:B9)'), total('=SUM(C4:C9)'), total('=B10-C10')]
        ],
        shown: { B10: '$2,540.00', C10: '$2,626.00', D10: '-$86.00', D4: '$0.00', D5: '-$62.00', D6: '$22.00', D7: '$9.00', D8: '-$55.00', D9: '$0.00' },
        charts: [{ type: 'doughnut', range: 'A3:C9', x: 560, y: 20, w: 400, h: 300, title: 'Where the money goes', style: 'modern', labels: true }]
      }
    },
    {
      id: 'invoice', name: 'Invoice', desc: 'Items, quantities and totals',
      sheet: {
        name: 'Invoice', widths: { 0: 260, 1: 70, 2: 110, 3: 120 }, heights: { 0: 42 },
        rows: [
          [c('INVOICE', { b: true, size: 24, color: '#13834a' })],
          ['[Your company]', '', 'Invoice #', '0001'],
          ['[Street, City]', '', 'Date', '[Date]'],
          [],
          [c('Bill to', { b: true })],
          ['[Customer name]'],
          [],
          [c('Description', H), c('Qty', H), c('Unit price', H), c('Amount', H)],
          ['Design work', '12', money(85), money('=B9*C9')],
          ['Hosting (12 months)', '1', money(240), money('=B10*C10')],
          ['Domain', '1', money(18), money('=B11*C11')],
          [],
          ['', '', c('Subtotal', { b: true }), money('=SUM(D9:D11)')],
          ['', '', c('Tax (10%)', { b: true }), money('=D13*0.1')],
          ['', '', c('Total', { b: true, size: 13 }), total('=D13+D14')],
          [],
          ['Thank you for your business!']
        ],
        shown: { D9: '$1,020.00', D10: '$240.00', D11: '$18.00', D13: '$1,278.00', D14: '$127.80', D15: '$1,405.80' },
        charts: []
      }
    },
    {
      id: 'tracker', name: 'Project tracker', desc: 'Tasks, owners, status and dates',
      sheet: {
        name: 'Project', widths: { 0: 230, 1: 110, 2: 110, 3: 100, 4: 90 }, heights: { 0: 42 },
        rows: [
          [c('Project tracker', T)],
          [],
          [c('Task', H), c('Owner', H), c('Status', H), c('Due', H), c('Done %', H)],
          ['Kick-off meeting', 'Ana', 'Done', '[Date]', c('1', { fmt: '0%' })],
          ['Research and interviews', 'Ben', 'Done', '[Date]', c('1', { fmt: '0%' })],
          ['First design', 'Chen', 'In progress', '[Date]', c('0.6', { fmt: '0%' })],
          ['Build', 'Dev team', 'In progress', '[Date]', c('0.3', { fmt: '0%' })],
          ['Testing', 'Eve', 'Not started', '[Date]', c('0', { fmt: '0%' })],
          ['Launch', 'Ana', 'Not started', '[Date]', c('0', { fmt: '0%' })],
          [],
          [c('Overall', { b: true }), '', '', '', c('=AVERAGE(E4:E9)', { b: true, fmt: '0%' })]
        ],
        shown: { E4: '100%', E5: '100%', E6: '60%', E7: '30%', E8: '0%', E9: '0%', E11: '48%' },
        charts: []
      }
    },
    {
      id: 'sales', name: 'Sales report', desc: 'Monthly sales by region, with a chart',
      sheet: {
        name: 'Sales', widths: { 0: 110, 1: 90, 2: 90, 3: 90, 4: 100 }, heights: { 0: 42 },
        rows: [
          [c('Sales report', T)],
          [],
          [c('Month', H), c('North', H), c('South', H), c('West', H), c('Total', H)],
          ['Jan', '120', '95', '80', c('=SUM(B4:D4)', { b: true })],
          ['Feb', '135', '102', '91', c('=SUM(B5:D5)', { b: true })],
          ['Mar', '128', '118', '99', c('=SUM(B6:D6)', { b: true })],
          ['Apr', '150', '121', '110', c('=SUM(B7:D7)', { b: true })],
          ['May', '171', '133', '118', c('=SUM(B8:D8)', { b: true })],
          ['Jun', '182', '140', '131', c('=SUM(B9:D9)', { b: true })]
        ],
        shown: { E4: '295', E5: '328', E6: '345', E7: '381', E8: '422', E9: '453' },
        charts: [{ type: 'bar', range: 'A3:D9', x: 560, y: 20, w: 460, h: 300, title: 'Sales by region', style: 'modern' }]
      }
    }
  ];

  // ---- Presentations. Elements use the deck's own format; ph marks title/subtitle/body text.
  const W = 960;
  const title = (text, o = {}) => ({ type: 'text', ph: 'title', html: text, x: 60, y: 36, w: 840, h: 90, size: 40, align: 'left', valign: 'middle', ...o });
  const sub = (text, o = {}) => ({ type: 'text', ph: 'subtitle', html: text, x: 120, y: 318, w: 720, h: 70, size: 24, align: 'center', valign: 'top', ...o });
  const bullets = (items, o = {}) => ({ type: 'text', ph: 'body', html: '<ul>' + items.map(i => `<li>${i}</li>`).join('') + '</ul>', x: 60, y: 140, w: 840, h: 360, size: 24, align: 'left', valign: 'top', ...o });
  const bar = (fill, o = {}) => ({ type: 'shape', shape: 'rect', html: '', fill, x: 60, y: 128, w: 96, h: 6, size: 18, align: 'center', valign: 'middle', ...o });
  const stat = (x, big, small, color) => [
    { type: 'text', html: big, x, y: 190, w: 260, h: 90, size: 54, align: 'left', valign: 'bottom', color },
    { type: 'text', html: small, x, y: 285, w: 260, h: 80, size: 18, align: 'left', valign: 'top' }
  ];
  const cover = (t, s, accent) => ({ layout: 'title', elements: [title(t, { x: 80, y: 150, w: 800, h: 150, size: 54, align: 'center', valign: 'bottom' }), bar(accent, { x: 432, y: 314, w: 96, h: 6 }), sub(s, { y: 334 })] });

  const SLIDES = [
    {
      id: 'pitch', name: 'Pitch deck', desc: 'Problem, solution, market, team, ask',
      theme: 'aurora', colors: { bg: '#1e1b4b', title: '#ffffff', text: '#c7d2fe', accent: '#f472b6' },
      slides: [
        cover('[Company name]', 'The one-line description of what you do', '#f472b6'),
        { layout: 'content', elements: [title('The problem'), bar('#f472b6'), bullets(['Who has the problem, and how often', 'What it costs them today', 'Why existing answers fall short'])] },
        { layout: 'content', elements: [title('Our solution'), bar('#f472b6'), bullets(['What the product does, in one sentence', 'The moment it saves the day', 'Why now'])] },
        { layout: 'titleonly', elements: [title('Market'), bar('#f472b6'), ...stat(60, '$4.2B', 'Total market', '#f472b6'), ...stat(350, '$610M', 'Reachable market', '#f472b6'), ...stat(640, '18%', 'Growth a year', '#f472b6')] },
        { layout: 'content', elements: [title('Team'), bar('#f472b6'), bullets(['[Name] — CEO, [background]', '[Name] — CTO, [background]', '[Name] — Design, [background]'])] },
        { layout: 'section', elements: [title('The ask', { x: 80, y: 170, w: 800, h: 110, size: 48, valign: 'bottom' }), sub('[Amount] to reach [milestone] by [date]', { x: 80, align: 'left', w: 800, y: 290 })] }
      ]
    },
    {
      id: 'update', name: 'Project update', desc: 'Status, wins, risks and next steps',
      theme: 'paper', colors: { bg: '#fbf8f3', title: '#1c1917', text: '#44403c', accent: '#b45309' },
      slides: [
        cover('Project update', '[Project] · [Date]', '#b45309'),
        { layout: 'titleonly', elements: [title('At a glance'), bar('#b45309'), ...stat(60, 'On track', 'Overall status', '#b45309'), ...stat(350, '72%', 'Work complete', '#b45309'), ...stat(640, '3', 'Open risks', '#b45309')] },
        { layout: 'two', elements: [title('Wins and risks'), bar('#b45309'), bullets(['Shipped [feature]', 'Signed [partner]', 'Cut costs by [x]'], { w: 410 }), bullets(['[Risk] — owner, plan', '[Risk] — owner, plan'], { x: 490, w: 410 })] },
        { layout: 'content', elements: [title('Next steps'), bar('#b45309'), bullets(['[Step] — [owner], [date]', '[Step] — [owner], [date]', '[Step] — [owner], [date]'])] }
      ]
    },
    {
      id: 'lesson', name: 'Lesson', desc: 'Objectives, ideas, practice, recap',
      theme: 'mint', colors: { bg: '#ecfdf5', title: '#064e3b', text: '#065f46', accent: '#059669' },
      slides: [
        cover('[Lesson title]', '[Course] · [Teacher]', '#059669'),
        { layout: 'content', elements: [title('Today you will'), bar('#059669'), bullets(['Understand [idea]', 'Practise [skill]', 'Apply it to [problem]'])] },
        { layout: 'two', elements: [title('Key ideas'), bar('#059669'), bullets(['[Idea one]', '[Idea two]'], { w: 410 }), bullets(['[Example]', '[Example]'], { x: 490, w: 410 })] },
        { layout: 'content', elements: [title('Your turn'), bar('#059669'), bullets(['[Exercise one]', '[Exercise two]', 'Share with a partner'])] },
        { layout: 'content', elements: [title('Recap'), bar('#059669'), bullets(['[What we learned]', '[Homework]'])] }
      ]
    },
    {
      id: 'review', name: 'Quarterly review', desc: 'Results, highlights and goals',
      theme: 'dark', colors: { bg: '#111827', title: '#f9fafb', text: '#cbd5e1', accent: '#60a5fa' },
      slides: [
        cover('Q[n] review', '[Team] · [Date]', '#60a5fa'),
        { layout: 'titleonly', elements: [title('Results'), bar('#60a5fa'), ...stat(60, '+14%', 'Revenue on last quarter', '#60a5fa'), ...stat(350, '92', 'Customer score', '#60a5fa'), ...stat(640, '1.2k', 'New customers', '#60a5fa')] },
        { layout: 'content', elements: [title('Highlights'), bar('#60a5fa'), bullets(['[Highlight]', '[Highlight]', '[Highlight]'])] },
        { layout: 'content', elements: [title('Goals for next quarter'), bar('#60a5fa'), bullets(['[Goal] — measured by [metric]', '[Goal] — measured by [metric]'])] }
      ]
    }
  ];

  const byId = (list, id) => list.find(t => t.id === id) || null;
  const plain = html => String(html).replace(/<\/li>/g, '\n').replace(/<[^>]+>/g, '').replace(/&amp;/g, '&').trim();

  /** A miniature, as an app's Platform.preview would describe the result. */
  function preview(app, t) {
    if (app === 'sheets') {
      const rows = [], bold = [];
      t.sheet.rows.slice(0, 14).forEach((row, r) => {
        const cells = [], b = [];
        for (let k = 0; k < 8; k++) {
          const cell = row[k], ref = 'ABCDEFGH'[k] + (r + 1);
          const v = cell && typeof cell === 'object' ? cell.v : cell;
          const shown = t.sheet.shown[ref] || (typeof v === 'string' && v[0] !== '=' ? (cell && cell.s && cell.s.fmt === MONEY ? '$' + Number(v).toLocaleString('en-US', { minimumFractionDigits: 2 }) : cell && cell.s && cell.s.fmt === '0%' ? Math.round(v * 100) + '%' : v) : '');
          cells.push(shown || ''); b.push(!!(cell && cell.s && cell.s.b));
        }
        rows.push(cells); bold.push(b);
      });
      return { kind: 'sheets', rows, bold };
    }
    const first = t.slides[0], C = t.colors;
    return {
      kind: 'slides', w: W, h: 540, bg: C.bg,
      els: first.elements.map(e => ({ x: e.x, y: e.y, w: e.w, h: e.h, s: plain(e.html) || undefined, size: e.size, fill: e.type === 'shape' ? e.fill : undefined, color: e.color || (e.ph === 'title' ? C.title : C.text), a: e.align === 'left' ? undefined : e.align, bold: e.ph === 'title' ? 1 : undefined }))
    };
  }

  /** Opens a template as a new, unsaved file named for saving as .xlsx or .pptx. */
  function open(app, t, sameTab) {
    const bytes = new TextEncoder().encode(JSON.stringify({ isadekTemplate: app, id: t.id }));
    return Platform.openData(app, t.name + (app === 'sheets' ? '.xlsx' : '.pptx'), bytes, sameTab);
  }

  /** The template a marker file names, or null for any other file. */
  function fromBytes(bytes, app) {
    if (!bytes || bytes.length > 200 || bytes[0] !== 0x7b) return null;
    try {
      const m = JSON.parse(new TextDecoder().decode(bytes));
      if (!m || m.isadekTemplate !== app || typeof m.id !== 'string') return null;
      return byId(app === 'sheets' ? SHEETS : SLIDES, m.id);
    } catch { return null; }
  }

  /** Opens an outline (Markdown) as a new presentation built from it. */
  function openOutline(name, markdown) {
    const bytes = new TextEncoder().encode(JSON.stringify({ isadekOutline: String(markdown).slice(0, 200000) }));
    return Platform.openData('slides', String(name).replace(/\.[^.]+$/, '') + '.pptx', bytes);
  }
  /** The outline a handed-over file carries, or null for any other file. */
  function outlineFrom(bytes) {
    if (!bytes || bytes.length > 1e6 || bytes[0] !== 0x7b) return null;
    try {
      const m = JSON.parse(new TextDecoder().decode(bytes));
      return m && typeof m.isadekOutline === 'string' && Object.keys(m).length === 1 ? m.isadekOutline : null;
    } catch { return null; }
  }

  /** A template's card: its miniature, name and what it is for. */
  function card(app, t, onPick) {
    const b = document.createElement('button');
    b.className = 'tpl';
    b.dataset.app = app;
    b.appendChild(UI.thumbnail(preview(app, t), app));
    const name = document.createElement('b'), desc = document.createElement('small');
    name.textContent = t.name; desc.textContent = t.desc;
    b.append(name, desc);
    b.addEventListener('click', () => onPick(t));
    return b;
  }

  window.SuiteTemplates = { sheets: SHEETS, slides: SLIDES, preview, open, fromBytes, card, openOutline, outlineFrom };
})();
