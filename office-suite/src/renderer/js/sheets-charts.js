'use strict';
// Isadek Sheets — charts as Excel DrawingML (xl/charts/chartN.xml), both ways.
// toXml(spec) writes a chart Excel, LibreOffice and Google Sheets open as a real, editable chart;
// fromXml(xml) reads one back into { type, title, series: [{ nameRef, catRef, valRef }] }.
(function (root) {
  const NS = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const esc = s => String(s).replace(/[&<>"]/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]));
  const hex = c => (c || '#4472c4').slice(1).toUpperCase();

  function strRef(ref, values) {
    if (!ref) return '';
    const pts = (values || []).map((v, i) => `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`).join('');
    return `<c:strRef><c:f>${esc(ref)}</c:f><c:strCache><c:ptCount val="${(values || []).length}"/>${pts}</c:strCache></c:strRef>`;
  }
  function numRef(ref, values) {
    const pts = (values || []).map((v, i) => (typeof v === 'number' && isFinite(v) ? `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>` : '')).join('');
    return `<c:numRef><c:f>${esc(ref)}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${(values || []).length}"/>${pts}</c:numCache></c:numRef>`;
  }
  const fill = c => `<c:spPr><a:solidFill><a:srgbClr val="${hex(c)}"/></a:solidFill></c:spPr>`;
  const lineSp = c => `<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="${hex(c)}"/></a:solidFill><a:round/></a:ln></c:spPr>`;
  const tx = s => (s.nameRef ? `<c:tx>${strRef(s.nameRef, [s.name])}</c:tx>` : s.name ? `<c:tx><c:v>${esc(s.name)}</c:v></c:tx>` : '');
  const cat = s => (s.catRef ? `<c:cat>${s.catNumeric ? numRef(s.catRef, s.cats) : strRef(s.catRef, s.cats)}</c:cat>` : '');
  const val = s => `<c:val>${numRef(s.valRef, s.values)}</c:val>`;
  const axes = (horizontal, scatter) => {
    const cAx = scatter
      ? `<c:valAx><c:axId val="500000001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="500000002"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>`
      : `<c:catAx><c:axId val="500000001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${horizontal ? 'l' : 'b'}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="500000002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx>`;
    return cAx + `<c:valAx><c:axId val="500000002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${horizontal ? 'b' : 'l'}"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="500000001"/><c:crosses val="autoZero"/><c:crossBetween val="${scatter ? 'midCat' : 'between'}"/></c:valAx>`;
  };
  const AX = '<c:axId val="500000001"/><c:axId val="500000002"/>';
  const ser = (s, i, body) => `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx(s)}${body}</c:ser>`;
  const barSer = (s, i) => ser(s, i, fill(s.color) + '<c:invertIfNegative val="0"/>' + cat(s) + val(s));
  const lineSer = (s, i) => ser(s, i, lineSp(s.color) + `<c:marker><c:symbol val="circle"/><c:size val="5"/>${fill(s.color)}</c:marker>` + cat(s) + val(s) + '<c:smooth val="0"/>');

  /**
   * spec: { type: bar|hbar|line|area|pie|doughnut|scatter|combo, title, series: [{ name, nameRef, color,
   * catRef, cats, catNumeric, valRef, values }] }
   */
  function toXml(spec) {
    const S = spec.series;
    let plot;
    switch (spec.type) {
      case 'line': plot = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${S.map(lineSer).join('')}<c:marker val="1"/>${AX}</c:lineChart>` + axes(false); break;
      case 'area': plot = `<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>${S.map((s, i) => ser(s, i, fill(s.color) + cat(s) + val(s))).join('')}${AX}</c:areaChart>` + axes(false); break;
      case 'pie': case 'doughnut': {
        const s = S[0];
        const pts = (s.values || []).map((_, i) => `<c:dPt><c:idx val="${i}"/><c:bubble3D val="0"/>${fill(spec.palette[i % spec.palette.length])}</c:dPt>`).join('');
        const one = ser(s, 0, pts + cat(s) + val(s));
        plot = spec.type === 'pie'
          ? `<c:pieChart><c:varyColors val="1"/>${one}<c:firstSliceAng val="0"/></c:pieChart>`
          : `<c:doughnutChart><c:varyColors val="1"/>${one}<c:firstSliceAng val="0"/><c:holeSize val="50"/></c:doughnutChart>`;
        break;
      }
      case 'scatter': {
        const body = S.map((s, i) => ser(s, i, `<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="6"/>${fill(s.color)}</c:marker>` +
          `<c:xVal>${numRef(s.catRef, s.cats)}</c:xVal><c:yVal>${numRef(s.valRef, s.values)}</c:yVal><c:smooth val="0"/>`)).join('');
        plot = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${body}${AX}</c:scatterChart>` + axes(false, true);
        break;
      }
      case 'combo': {
        const bars = S.slice(0, Math.max(1, S.length - 1)), lines = S.length > 1 ? S.slice(-1) : [];
        plot = `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${bars.map(barSer).join('')}<c:gapWidth val="150"/>${AX}</c:barChart>`
          + (lines.length ? `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${lines.map((s, i) => lineSer(s, bars.length + i)).join('')}<c:marker val="1"/>${AX}</c:lineChart>` : '')
          + axes(false);
        break;
      }
      default: {
        const h = spec.type === 'hbar';
        plot = `<c:barChart><c:barDir val="${h ? 'bar' : 'col'}"/><c:grouping val="clustered"/><c:varyColors val="0"/>${S.map(barSer).join('')}<c:gapWidth val="150"/>${AX}</c:barChart>` + axes(h);
      }
    }
    const title = spec.title ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="0"/></a:pPr><a:r><a:rPr lang="en-US" sz="1400" b="0"/><a:t>${esc(spec.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>` : '<c:autoTitleDeleted val="1"/>';
    const legend = S.length > 1 || spec.type === 'pie' || spec.type === 'doughnut' ? '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>' : '';
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace ${NS}><c:roundedCorners val="0"/><c:chart>${title}<c:plotArea><c:layout/>${plot}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`;
  }

  /** The drawing that places charts on a sheet: anchors are { from: {col,colOff,row,rowOff}, to: {…} } in EMU offsets. */
  function drawingXml(anchors) {
    const pos = (tag, p) => `<xdr:${tag}><xdr:col>${p.col}</xdr:col><xdr:colOff>${p.colOff}</xdr:colOff><xdr:row>${p.row}</xdr:row><xdr:rowOff>${p.rowOff}</xdr:rowOff></xdr:${tag}>`;
    const items = anchors.map((a, i) => `<xdr:twoCellAnchor editAs="oneCell">${pos('from', a.from)}${pos('to', a.to)}<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${i + 2}" name="${esc(a.name || 'Chart ' + (i + 1))}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId${i + 1}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${items}</xdr:wsDr>`;
  }

  // ---------- Reading ----------
  const kids = (el, name) => (el ? [...el.children].filter(e => e.localName === name) : []);
  const kid = (el, name) => kids(el, name)[0];
  const deep = (el, name) => (el ? [...el.getElementsByTagName('*')].filter(e => e.localName === name) : []);
  /** A chart part's type, title and series references. */
  function fromXml(xml) {
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    if (doc.querySelector('parsererror')) return null;
    const plot = deep(doc.documentElement, 'plotArea')[0];
    if (!plot) return null;
    const groups = [...plot.children].filter(e => /Chart$/.test(e.localName));
    if (!groups.length) return null;
    const first = groups[0].localName;
    let type = { barChart: 'bar', bar3DChart: 'bar', lineChart: 'line', line3DChart: 'line', areaChart: 'area', area3DChart: 'area', pieChart: 'pie', pie3DChart: 'pie', ofPieChart: 'pie', doughnutChart: 'doughnut', scatterChart: 'scatter', bubbleChart: 'scatter', radarChart: 'line' }[first];
    if (!type) return null;
    if (type === 'bar' && kid(groups[0], 'barDir') && kid(groups[0], 'barDir').getAttribute('val') === 'bar') type = 'hbar';
    if (groups.length > 1 && first === 'barChart' && groups.some(g => g.localName === 'lineChart')) type = 'combo';
    const f = el => { const x = deep(el, 'f')[0]; return x ? x.textContent.trim() : null; };
    const series = [];
    groups.forEach(g => kids(g, 'ser').forEach(s => {
      const txEl = kid(s, 'tx');
      series.push({
        nameRef: txEl ? f(txEl) : null,
        name: txEl ? (deep(txEl, 'v')[0] || {}).textContent || '' : '',
        catRef: f(kid(s, 'cat') || kid(s, 'xVal')),
        valRef: f(kid(s, 'val') || kid(s, 'yVal'))
      });
    }));
    const titleEl = deep(doc.documentElement, 'title')[0];
    const title = titleEl ? deep(titleEl, 't').map(t => t.textContent).join('') : '';
    return { type, title: title.slice(0, 200), series: series.filter(s => s.valRef).slice(0, 50) };
  }
  /** The placement of each chart in a drawing part, with the relationship id of its chart. */
  function anchorsFromDrawing(xml) {
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    if (doc.querySelector('parsererror')) return [];
    const num = (el, n) => Math.max(0, parseInt((kid(el, n) || {}).textContent, 10) || 0);
    return [...doc.documentElement.children].filter(a => /^(two|one)CellAnchor$/.test(a.localName)).map(a => {
      const from = kid(a, 'from'), to = kid(a, 'to'), ext = kid(a, 'ext');
      const chart = deep(a, 'chart')[0];
      if (!from || !chart) return null;
      const rid = chart.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') || chart.getAttribute('r:id');
      return {
        rid,
        from: { col: num(from, 'col'), colOff: num(from, 'colOff'), row: num(from, 'row'), rowOff: num(from, 'rowOff') },
        to: to ? { col: num(to, 'col'), colOff: num(to, 'colOff'), row: num(to, 'row'), rowOff: num(to, 'rowOff') } : null,
        ext: ext ? { cx: +ext.getAttribute('cx') || 0, cy: +ext.getAttribute('cy') || 0 } : null
      };
    }).filter(Boolean);
  }

  const api = { toXml, drawingXml, fromXml, anchorsFromDrawing };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.SheetsCharts = api;
})(typeof window !== 'undefined' ? window : globalThis);
