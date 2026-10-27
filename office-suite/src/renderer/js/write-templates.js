'use strict';
// Atlas Write — document templates and the template gallery.
// Each template is the HTML that Write saves (body + optional header/footer + page setup),
// so opening one goes through the normal HTML loader.
(function () {
  const ph = t => `<span style="color:#767676">[${t}]</span>`;
  const T = [
    {
      id: 'blank', name: 'Blank document', icon: '📄', desc: 'Start from scratch',
      body: '<p><br></p>'
    },
    {
      id: 'resume', name: 'Resume', icon: '👤', desc: 'Clean one-page CV',
      setup: { margin: '0.75in' },
      body: `<h1 class="title" style="text-align:center">${ph('Your Name')}</h1>
<p style="text-align:center">${ph('City, Country')} · ${ph('phone')} · ${ph('email@example.com')} · ${ph('linkedin.com/in/you')}</p>
<hr>
<h2>Summary</h2>
<p>${ph('Two or three sentences about your experience, strengths and the role you are looking for.')}</p>
<h2>Experience</h2>
<p><b>${ph('Job Title')}</b>, ${ph('Company')} — <i>${ph('Month Year')} – ${ph('Present')}</i></p>
<ul><li>${ph('Achievement with a measurable result')}</li><li>${ph('Responsibility or project you led')}</li></ul>
<p><b>${ph('Job Title')}</b>, ${ph('Company')} — <i>${ph('Month Year')} – ${ph('Month Year')}</i></p>
<ul><li>${ph('Achievement with a measurable result')}</li></ul>
<h2>Education</h2>
<p><b>${ph('Degree')}</b>, ${ph('School')} — <i>${ph('Year')}</i></p>
<h2>Skills</h2>
<p>${ph('Skill')} · ${ph('Skill')} · ${ph('Skill')} · ${ph('Language')}</p>`
    },
    {
      id: 'cover-letter', name: 'Cover letter', icon: '✉️', desc: 'Job application letter',
      body: `<p>${ph('Your Name')}<br>${ph('Address')}<br>${ph('email@example.com')}</p>
<p>${ph('Date')}</p>
<p>${ph('Hiring Manager')}<br>${ph('Company')}<br>${ph('Address')}</p>
<p>Dear ${ph('Hiring Manager')},</p>
<p>I am writing to apply for the ${ph('position')} role at ${ph('Company')}. ${ph('Why you are interested and how you heard about it.')}</p>
<p>${ph('A paragraph about relevant experience and an achievement that matches the job.')}</p>
<p>${ph('A paragraph about why you would be a good fit for the team and company.')}</p>
<p>Thank you for your time and consideration. I look forward to hearing from you.</p>
<p>Sincerely,</p>
<p>${ph('Your Name')}</p>`
    },
    {
      id: 'business-letter', name: 'Business letter', icon: '🏢', desc: 'Formal letter with letterhead',
      header: `<p style="text-align:right"><b>${ph('Company Name')}</b><br>${ph('Street, City')} · ${ph('phone')}</p>`,
      body: `<p>${ph('Date')}</p>
<p>${ph('Recipient Name')}<br>${ph('Title')}<br>${ph('Company')}<br>${ph('Address')}</p>
<p><b>Re: ${ph('Subject')}</b></p>
<p>Dear ${ph('Recipient Name')},</p>
<p>${ph('Body of the letter.')}</p>
<p>Yours faithfully,</p>
<p><br></p>
<p>${ph('Your Name')}<br>${ph('Title')}</p>`
    },
    {
      id: 'report', name: 'Report', icon: '📊', desc: 'Title page, contents and sections',
      setup: { differentFirst: true },
      footer: '<p style="text-align:center">Page <span class="fld" data-fld="page">1</span> of <span class="fld" data-fld="numpages">1</span></p>',
      body: `<p><br></p><p><br></p><p><br></p>
<h1 class="title" style="text-align:center">${ph('Report Title')}</h1>
<p style="text-align:center">${ph('Subtitle')}</p>
<p style="text-align:center">${ph('Author')} · ${ph('Date')}</p>
<hr class="page-break">
<h1>Executive Summary</h1>
<p>${ph('A short summary of the purpose, findings and recommendations.')}</p>
<h1>Introduction</h1>
<p>${ph('Background and scope of the report.')}</p>
<h1>Findings</h1>
<h2>${ph('Finding one')}</h2>
<p>${ph('Details and evidence.')}</p>
<h2>${ph('Finding two')}</h2>
<p>${ph('Details and evidence.')}</p>
<h1>Recommendations</h1>
<ol><li>${ph('Recommendation')}</li><li>${ph('Recommendation')}</li></ol>
<h1>Conclusion</h1>
<p>${ph('Closing remarks.')}</p>`
    },
    {
      id: 'memo', name: 'Memo', icon: '📝', desc: 'Internal memorandum',
      body: `<h1 class="title">Memo</h1>
<table><tbody>
<tr><td style="width:20%"><b>To:</b></td><td>${ph('Recipients')}</td></tr>
<tr><td><b>From:</b></td><td>${ph('Your Name')}</td></tr>
<tr><td><b>Date:</b></td><td>${ph('Date')}</td></tr>
<tr><td><b>Subject:</b></td><td>${ph('Subject')}</td></tr>
</tbody></table>
<hr>
<p>${ph('Message.')}</p>`
    },
    {
      id: 'minutes', name: 'Meeting minutes', icon: '🗓', desc: 'Agenda, decisions and actions',
      body: `<h1 class="title">${ph('Meeting Name')}</h1>
<p><b>Date:</b> ${ph('Date')} · <b>Time:</b> ${ph('Time')} · <b>Location:</b> ${ph('Location')}</p>
<p><b>Attendees:</b> ${ph('Names')}</p>
<p><b>Apologies:</b> ${ph('Names')}</p>
<h2>Agenda</h2>
<ol><li>${ph('Item')}</li><li>${ph('Item')}</li><li>${ph('Item')}</li></ol>
<h2>Discussion</h2>
<p>${ph('Notes on each agenda item.')}</p>
<h2>Decisions</h2>
<ul><li>${ph('Decision')}</li></ul>
<h2>Action items</h2>
<table><tbody><tr><th>Action</th><th>Owner</th><th>Due</th></tr>
<tr><td>${ph('Action')}</td><td>${ph('Name')}</td><td>${ph('Date')}</td></tr>
<tr><td>${ph('Action')}</td><td>${ph('Name')}</td><td>${ph('Date')}</td></tr></tbody></table>
<p><b>Next meeting:</b> ${ph('Date')}</p>`
    },
    {
      id: 'essay', name: 'Essay (MLA)', icon: '🎓', desc: 'Double-spaced academic paper',
      setup: { citeStyle: 'mla' },
      header: '<p style="text-align:right">' + ph('Last name') + ' <span class="fld" data-fld="page">1</span></p>',
      body: `<p style="line-height:2">${ph('Your Name')}</p>
<p style="line-height:2">${ph('Instructor')}</p>
<p style="line-height:2">${ph('Course')}</p>
<p style="line-height:2">${ph('Day Month Year')}</p>
<p style="text-align:center;line-height:2">${ph('Title')}</p>
<p style="line-height:2;text-indent:0.5in">${ph('Introduction with your thesis statement.')}</p>
<p style="line-height:2;text-indent:0.5in">${ph('Body paragraph.')}</p>
<p style="line-height:2;text-indent:0.5in">${ph('Conclusion.')}</p>`
    },
    {
      id: 'newsletter', name: 'Newsletter', icon: '📰', desc: 'Two-column layout',
      setup: { columns: 2, margin: '0.75in' },
      body: `<h1 class="title">${ph('Newsletter Name')}</h1>
<p><i>${ph('Issue')} · ${ph('Month Year')}</i></p>
<h2>${ph('Lead story headline')}</h2>
<p>${ph('Lead story text.')}</p>
<h2>${ph('Second story')}</h2>
<p>${ph('Story text.')}</p>
<h2>${ph('Upcoming events')}</h2>
<ul><li>${ph('Event, date')}</li><li>${ph('Event, date')}</li></ul>
<h2>${ph('Contact us')}</h2>
<p>${ph('Contact details.')}</p>`
    },
    {
      id: 'invoice', name: 'Invoice', icon: '🧾', desc: 'Simple invoice',
      body: `<h1 class="title" style="text-align:right">INVOICE</h1>
<p><b>${ph('Your Business')}</b><br>${ph('Address')}<br>${ph('email@example.com')}</p>
<p><b>Bill to:</b><br>${ph('Client Name')}<br>${ph('Address')}</p>
<p><b>Invoice #:</b> ${ph('0001')} · <b>Date:</b> ${ph('Date')} · <b>Due:</b> ${ph('Date')}</p>
<table><tbody><tr><th>Description</th><th style="width:12%">Qty</th><th style="width:18%">Unit price</th><th style="width:18%">Amount</th></tr>
<tr><td>${ph('Item')}</td><td>1</td><td>${ph('0.00')}</td><td>${ph('0.00')}</td></tr>
<tr><td>${ph('Item')}</td><td>1</td><td>${ph('0.00')}</td><td>${ph('0.00')}</td></tr>
<tr><td colspan="3" style="text-align:right"><b>Total</b></td><td><b>${ph('0.00')}</b></td></tr></tbody></table>
<p>${ph('Payment terms and bank details.')}</p>
<p>Thank you for your business!</p>`
    }
  ];

  /** The template as an HTML document in the format Write saves. */
  function toHtml(t) {
    const setup = JSON.stringify(Object.assign({ margin: '1in', orientation: 'portrait', size: 'letter', columns: 1, color: '#ffffff' }, t.setup || {}));
    const attr = s => s.replace(/&/g, '&amp;').replace(/"/g, '&quot;');
    return `<!doctype html><html><head><meta charset="utf-8"><title>${t.name}</title><meta name="atlas-setup" content="${attr(setup)}"></head><body>`
      + (t.header ? `<header class="atlas-header">${t.header}</header>` : '')
      + `<main class="atlas-doc">${t.body}</main>`
      + (t.footer ? `<footer class="atlas-footer">${t.footer}</footer>` : '')
      + '</body></html>';
  }

  /** Shows the gallery; resolves to a template or null. */
  function pick(title = 'New document') {
    return new Promise(resolve => {
      const back = document.createElement('div');
      back.className = 'modal-backdrop';
      back.innerHTML = `<div class="modal template-gallery" role="dialog" aria-label="${title}"><h2></h2><div class="tpl-grid"></div><div class="actions"><button class="cancel">Cancel</button></div></div>`;
      back.querySelector('h2').textContent = title;
      const grid = back.querySelector('.tpl-grid');
      T.forEach(t => {
        const b = document.createElement('button');
        b.className = 'tpl';
        b.innerHTML = '<span class="tpl-icon"></span><b></b><small></small>';
        b.querySelector('.tpl-icon').textContent = t.icon;
        b.querySelector('b').textContent = t.name;
        b.querySelector('small').textContent = t.desc;
        b.addEventListener('click', () => { back.remove(); resolve(t); });
        grid.appendChild(b);
      });
      const close = () => { back.remove(); resolve(null); };
      back.querySelector('.cancel').addEventListener('click', close);
      back.addEventListener('mousedown', e => { if (e.target === back) close(); });
      back.addEventListener('keydown', e => { if (e.key === 'Escape') close(); });
      document.body.appendChild(back);
      grid.querySelector('button').focus();
    });
  }

  /** Opens a template as a new, unsaved document (named for saving as .docx). */
  async function open(t, sameTab) {
    const bytes = new TextEncoder().encode(toHtml(t));
    return Platform.openData('write', (t.id === 'blank' ? 'Document' : t.name) + '.docx', bytes, sameTab);
  }

  window.WriteTemplates = { list: T, toHtml, pick, open };
})();
