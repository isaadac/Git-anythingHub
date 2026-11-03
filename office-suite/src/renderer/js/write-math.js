'use strict';
// Isadek Write — equations. An equation is typed in a linear form close to Word's:
//   x = (-b ± sqrt(b^2 - 4ac))/(2a)      a_i^2      \sum_(i=1)^n i      sqrt[3](x)      \alpha + \beta
// It is parsed once into a small tree, which becomes MathML on the page and Office Math (OMML)
// in .docx; equations in opened .docx files are turned back into the linear form.
(function () {
  const SYMBOLS = {
    alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ε', zeta: 'ζ', eta: 'η', theta: 'θ', iota: 'ι', kappa: 'κ', lambda: 'λ', mu: 'μ',
    nu: 'ν', xi: 'ξ', pi: 'π', rho: 'ρ', sigma: 'σ', tau: 'τ', upsilon: 'υ', phi: 'φ', chi: 'χ', psi: 'ψ', omega: 'ω',
    Gamma: 'Γ', Delta: 'Δ', Theta: 'Θ', Lambda: 'Λ', Xi: 'Ξ', Pi: 'Π', Sigma: 'Σ', Phi: 'Φ', Psi: 'Ψ', Omega: 'Ω',
    pm: '±', mp: '∓', times: '×', div: '÷', cdot: '·', le: '≤', leq: '≤', ge: '≥', geq: '≥', ne: '≠', neq: '≠', approx: '≈', equiv: '≡',
    to: '→', rightarrow: '→', leftarrow: '←', Rightarrow: '⇒', infty: '∞', partial: '∂', nabla: '∇', in: '∈', notin: '∉', subset: '⊂',
    cup: '∪', cap: '∩', forall: '∀', exists: '∃', deg: '°', prime: '′', ldots: '…', cdots: '⋯',
    sum: '∑', prod: '∏', int: '∫', iint: '∬', oint: '∮'
  };
  const NARY = new Set(['∑', '∏', '∫', '∬', '∮']);
  const FUNCS = new Set(['sin', 'cos', 'tan', 'cot', 'sec', 'csc', 'log', 'ln', 'exp', 'lim', 'max', 'min', 'det', 'sinh', 'cosh', 'tanh', 'arcsin', 'arccos', 'arctan']);
  const OPS = '+-=<>±∓×÷·≤≥≠≈≡→←⇒∈∉⊂∪∩∀∃,;:!|′°…⋯*';
  const OPEN = { '(': ')', '[': ']', '{': '}' };
  const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

  // ---------------- Parsing ----------------
  function tokenize(src) {
    const out = [];
    let i = 0;
    while (i < src.length) {
      const c = src[i];
      if (/\s/.test(c)) { i++; continue; }
      if (c === '\\') {
        const m = src.slice(i + 1).match(/^[A-Za-z]+/);
        if (m && SYMBOLS[m[0]]) { const v = SYMBOLS[m[0]]; out.push(NARY.has(v) ? { k: 'nary', v } : /[α-ωΑ-Ω]/.test(v) ? { k: 'id', v } : { k: 'op', v }); i += 1 + m[0].length; continue; }
        if (m && m[0] === 'sqrt') { out.push({ k: 'sqrt' }); i += 5; continue; }
        out.push({ k: 'op', v: m ? m[0] : '\\' }); i += 1 + (m ? m[0].length : 0); continue;
      }
      if (/[0-9.]/.test(c)) { const m = src.slice(i).match(/^[0-9]*\.?[0-9]+|^[0-9]+/); out.push({ k: 'num', v: m[0] }); i += m[0].length; continue; }
      if (/[A-Za-z]/.test(c)) {
        const m = src.slice(i).match(/^[A-Za-z]+/)[0];
        if (m === 'sqrt') { out.push({ k: 'sqrt' }); i += 4; continue; }
        // Function names stay whole; other letters are separate variables (ab = a·b).
        const fn = [...FUNCS].filter(f => m.startsWith(f)).sort((a, b) => b.length - a.length)[0];
        if (fn) { out.push({ k: 'fn', v: fn }); i += fn.length; continue; }
        out.push({ k: 'id', v: c }); i++; continue;
      }
      if (NARY.has(c)) { out.push({ k: 'nary', v: c }); i++; continue; }
      if (OPEN[c] || ')]}'.includes(c) || '^_/'.includes(c)) { out.push({ k: c }); i++; continue; }
      if (c === '"') { const j = src.indexOf('"', i + 1); const t = src.slice(i + 1, j < 0 ? src.length : j); out.push({ k: 'text', v: t }); i = j < 0 ? src.length : j + 1; continue; }
      out.push(OPS.includes(c) ? { k: 'op', v: c === '*' ? '·' : c } : { k: 'id', v: c });
      i++;
    }
    return out;
  }

  function parse(src) {
    const toks = tokenize(String(src || ''));
    let p = 0;
    const peek = () => toks[p], next = () => toks[p++];
    const strip = n => (n && n.t === 'paren' && n.o === '(' ? n.e : n);
    function row(close) {
      const c = [];
      while (p < toks.length && !(peek().k === close)) {
        if (')]}'.includes(peek().k)) { next(); continue; } // stray closer
        let f = factor();
        while (peek() && peek().k === '/') { next(); const d = factor(); f = { t: 'frac', n: strip(f), d: strip(d) }; }
        c.push(f);
      }
      if (close && peek() && peek().k === close) next();
      return c.length === 1 ? c[0] : { t: 'row', c };
    }
    function scripts(base) {
      let sb = null, sp = null;
      while (peek() && (peek().k === '^' || peek().k === '_')) {
        const k = next().k;
        const a = strip(atom());
        if (k === '^') sp = a; else sb = a;
      }
      if (!sb && !sp) return base;
      return sb && sp ? { t: 'subsup', b: base, sb, sp } : sp ? { t: 'sup', b: base, s: sp } : { t: 'sub', b: base, s: sb };
    }
    function factor() {
      const t = peek();
      if (t && t.k === 'nary') {
        next();
        let sub = null, sup = null;
        while (peek() && (peek().k === '^' || peek().k === '_')) { const k = next().k; const a = strip(atom()); if (k === '^') sup = a; else sub = a; }
        const e = p < toks.length && !')]}'.includes(peek().k) ? factor() : { t: 'row', c: [] };
        return { t: 'nary', op: t.v, sub, sup, e };
      }
      return scripts(atom());
    }
    function atom() {
      const t = next();
      if (!t) return { t: 'row', c: [] };
      if (t.k === 'num') return { t: 'num', v: t.v };
      if (t.k === 'id') return { t: 'id', v: t.v };
      if (t.k === 'fn') return { t: 'fn', v: t.v };
      if (t.k === 'op') return { t: 'op', v: t.v };
      if (t.k === 'text') return { t: 'text', v: t.v };
      if (t.k === '(' || t.k === '[') return { t: 'paren', o: t.k, c: OPEN[t.k], e: row(OPEN[t.k]) };
      if (t.k === '{') return row('}');
      if (t.k === 'sqrt') {
        let deg = null;
        if (peek() && peek().k === '[') { next(); deg = row(']'); }
        const e = peek() && (peek().k === '(' || peek().k === '{') ? strip(atom()) : atom();
        return { t: 'sqrt', e, deg };
      }
      if (t.k === 'nary') { p--; return factor(); }
      return { t: 'op', v: t.k };
    }
    return row(null);
  }

  // ---------------- MathML (on the page) ----------------
  function mathml(n) {
    switch (n.t) {
      case 'row': return `<mrow>${n.c.map(mathml).join('')}</mrow>`;
      case 'num': return `<mn>${esc(n.v)}</mn>`;
      case 'id': return `<mi>${esc(n.v)}</mi>`;
      case 'fn': return `<mi>${esc(n.v)}</mi><mo>&#x2061;</mo>`;
      case 'op': return `<mo>${esc(n.v)}</mo>`;
      case 'text': return `<mtext>${esc(n.v)}</mtext>`;
      case 'frac': return `<mfrac>${mathml(n.n)}${mathml(n.d)}</mfrac>`;
      case 'sup': return `<msup>${mathml(n.b)}${mathml(n.s)}</msup>`;
      case 'sub': return `<msub>${mathml(n.b)}${mathml(n.s)}</msub>`;
      case 'subsup': return `<msubsup>${mathml(n.b)}${mathml(n.sb)}${mathml(n.sp)}</msubsup>`;
      case 'sqrt': return n.deg ? `<mroot>${mathml(n.e)}${mathml(n.deg)}</mroot>` : `<msqrt>${mathml(n.e)}</msqrt>`;
      case 'paren': return `<mrow><mo>${n.o}</mo>${mathml(n.e)}<mo>${n.c}</mo></mrow>`;
      case 'nary': {
        const op = `<mo largeop="true">${n.op}</mo>`, sb = n.sub ? mathml(n.sub) : '<mrow></mrow>', sp = n.sup ? mathml(n.sup) : '<mrow></mrow>';
        const under = n.op === '∑' || n.op === '∏';
        const head = !n.sub && !n.sup ? op : under ? (n.sub && n.sup ? `<munderover>${op}${sb}${sp}</munderover>` : n.sub ? `<munder>${op}${sb}</munder>` : `<mover>${op}${sp}</mover>`)
          : (n.sub && n.sup ? `<msubsup>${op}${sb}${sp}</msubsup>` : n.sub ? `<msub>${op}${sb}</msub>` : `<msup>${op}${sp}</msup>`);
        return `<mrow>${head}${mathml(n.e)}</mrow>`;
      }
    }
    return '';
  }
  const toMathML = (src, display) => `<math xmlns="http://www.w3.org/1998/Math/MathML"${display ? ' display="block"' : ''}>${mathml(parse(src))}</math>`;

  // ---------------- Office Math (.docx) ----------------
  const FONT = '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>';
  const run = (t, plain) => `<m:r>${plain ? '<m:rPr><m:sty m:val="p"/></m:rPr>' : ''}${FONT}<m:t xml:space="preserve">${esc(t)}</m:t></m:r>`;
  function omml(n) {
    switch (n.t) {
      case 'row': return n.c.map(omml).join('');
      case 'num': return run(n.v, true);
      case 'id': return run(n.v, false);
      case 'fn': return run(n.v, true);
      case 'op': return run(n.v, true);
      case 'text': return run(n.v, true);
      case 'frac': return `<m:f><m:num>${omml(n.n)}</m:num><m:den>${omml(n.d)}</m:den></m:f>`;
      case 'sup': return `<m:sSup><m:e>${omml(n.b)}</m:e><m:sup>${omml(n.s)}</m:sup></m:sSup>`;
      case 'sub': return `<m:sSub><m:e>${omml(n.b)}</m:e><m:sub>${omml(n.s)}</m:sub></m:sSub>`;
      case 'subsup': return `<m:sSubSup><m:e>${omml(n.b)}</m:e><m:sub>${omml(n.sb)}</m:sub><m:sup>${omml(n.sp)}</m:sup></m:sSubSup>`;
      case 'sqrt': return n.deg ? `<m:rad><m:deg>${omml(n.deg)}</m:deg><m:e>${omml(n.e)}</m:e></m:rad>` : `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${omml(n.e)}</m:e></m:rad>`;
      case 'paren': return `<m:d>${n.o === '(' ? '' : `<m:dPr><m:begChr m:val="${n.o}"/><m:endChr m:val="${n.c}"/></m:dPr>`}<m:e>${omml(n.e)}</m:e></m:d>`;
      case 'nary': {
        const pr = `<m:naryPr>${n.op === '∫' ? '' : `<m:chr m:val="${n.op}"/>`}<m:limLoc m:val="${n.op === '∑' || n.op === '∏' ? 'undOvr' : 'subSup'}"/>${n.sub ? '' : '<m:subHide m:val="1"/>'}${n.sup ? '' : '<m:supHide m:val="1"/>'}</m:naryPr>`;
        return `<m:nary>${pr}<m:sub>${n.sub ? omml(n.sub) : ''}</m:sub><m:sup>${n.sup ? omml(n.sup) : ''}</m:sup><m:e>${omml(n.e)}</m:e></m:nary>`;
      }
    }
    return '';
  }
  const toOMML = src => `<m:oMath>${omml(parse(src))}</m:oMath>`;

  /** An <m:oMath> element (from a .docx) in the linear form. */
  function fromOMML(el) {
    const kids = (e, name) => [...e.children].filter(c => c.localName === name);
    const val = (e, name) => { const c = e.getElementsByTagNameNS('*', name)[0]; return c ? (c.getAttribute('m:val') || c.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/math', 'val') || '') : ''; };
    const grp = s => ([...s].length <= 1 ? s : `{${s}}`);
    const part = (e, name) => { const c = kids(e, name)[0]; return c ? walk(c) : ''; };
    function walk(e) {
      let out = '';
      for (const c of e.children) {
        switch (c.localName) {
          case 'r': out += [...c.getElementsByTagNameNS('*', 't')].map(t => t.textContent).join(''); break;
          case 'f': out += `${grp(part(c, 'num'))}/${grp(part(c, 'den'))}`; break;
          case 'sSup': out += `${grp(part(c, 'e'))}^${grp(part(c, 'sup'))}`; break;
          case 'sSub': out += `${grp(part(c, 'e'))}_${grp(part(c, 'sub'))}`; break;
          case 'sSubSup': out += `${grp(part(c, 'e'))}_${grp(part(c, 'sub'))}^${grp(part(c, 'sup'))}`; break;
          case 'rad': {
            const deg = part(c, 'deg');
            out += deg && val(c, 'degHide') !== '1' ? `sqrt[${deg}](${part(c, 'e')})` : `sqrt(${part(c, 'e')})`;
            break;
          }
          case 'd': {
            const pr = kids(c, 'dPr')[0];
            const beg = pr && kids(pr, 'begChr')[0] ? val(pr, 'begChr') : '(';
            const end = pr && kids(pr, 'endChr')[0] ? val(pr, 'endChr') : ')';
            out += beg + kids(c, 'e').map(walk).join(', ') + end;
            break;
          }
          case 'nary': {
            const pr = kids(c, 'naryPr')[0];
            const chr = pr && kids(pr, 'chr')[0] ? val(pr, 'chr') : '∫';
            const sb = part(c, 'sub'), sp = part(c, 'sup');
            out += chr + (sb ? `_${grp(sb)}` : '') + (sp ? `^${grp(sp)}` : '') + grp(part(c, 'e'));
            break;
          }
          case 'func': out += `${part(c, 'fName')} ${grp(part(c, 'e'))}`; break;
          case 'rPr': case 'ctrlPr': case 'oMathParaPr': case 'argPr': break;
          default: out += walk(c);
        }
      }
      return out;
    }
    return walk(el).trim();
  }

  window.WriteMath = { parse, toMathML, toOMML, fromOMML };
})();
