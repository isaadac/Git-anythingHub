'use strict';
// Adds the parts of a workbook that SheetJS (community edition) can't write — conditional
// formatting, data validation and the like — to the .xlsx it produced, by editing its XML.
(function () {
  // Order of a worksheet's children after <sheetData> (ECMA-376 CT_Worksheet).
  const ORDER = ['sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews',
    'mergeCells', 'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter',
    'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF',
    'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];

  /** Inserts `xml` (elements named `tag`) where the schema puts them in a worksheet. */
  function insertInSheet(sheetXml, tag, xml) {
    if (!xml) return sheetXml;
    const after = ORDER.slice(ORDER.indexOf(tag) + 1);
    let at = -1;
    for (const t of after) {
      const i = sheetXml.search(new RegExp(`<${t}[\\s/>]`));
      if (i !== -1 && (at === -1 || i < at)) at = i;
    }
    if (at === -1) at = sheetXml.lastIndexOf('</worksheet>');
    return sheetXml.slice(0, at) + xml + sheetXml.slice(at);
  }

  const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

  /**
   * patch(bytes, { sheets: [{ conditionalFormatting, dataValidations, … }], dxfs: [xml] })
   * Each sheet entry maps a worksheet element name to the XML to add (sheets in workbook order).
   */
  const colNum = ref => [...ref.replace(/\d+$/, '')].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  /** Adds empty cells (with the given XML) to their rows, keeping rows and cells in order. */
  function insertCells(xml, cells) {
    const byRow = new Map();
    for (const [ref, x] of cells) { const m = /^([A-Z]+)(\d+)$/.exec(ref); if (!m) continue; const r = +m[2]; if (!byRow.has(r)) byRow.set(r, []); byRow.get(r).push({ ref, x }); }
    if (!byRow.size) return xml;
    if (/<sheetData\s*\/>/.test(xml)) xml = xml.replace(/<sheetData\s*\/>/, '<sheetData></sheetData>');
    for (const [r, list] of [...byRow].sort((a, b) => a[0] - b[0])) {
      list.sort((a, b) => colNum(a.ref) - colNum(b.ref));
      const rowRe = new RegExp(`<row r="${r}"([^>]*?)(\\/?)>`);
      const rm = rowRe.exec(xml);
      if (rm && rm[2] === '/') { xml = xml.replace(rowRe, `<row r="${r}"${rm[1]}>${list.map(c => c.x).join('')}</row>`); continue; }
      if (rm) {
        const start = rm.index + rm[0].length, end = xml.indexOf('</row>', start);
        let body = xml.slice(start, end);
        for (const c of list) {
          const after = [...body.matchAll(/<c r="([A-Z]+)\d+"/g)].find(m => colNum(m[1]) > colNum(c.ref));
          body = after ? body.slice(0, after.index) + c.x + body.slice(after.index) : body + c.x;
        }
        xml = xml.slice(0, start) + body + xml.slice(end);
        continue;
      }
      const next = [...xml.matchAll(/<row r="(\d+)"/g)].find(m => +m[1] > r);
      const at = next ? next.index : xml.indexOf('</sheetData>');
      xml = xml.slice(0, at) + `<row r="${r}">${list.map(c => c.x).join('')}</row>` + xml.slice(at);
    }
    return xml;
  }
  /** Frozen rows/columns as Excel panes. */
  function freezePane(xml, rows, cols) {
    if (!rows && !cols) return xml;
    const ref = String.fromCharCode(65 + cols) + (rows + 1);
    const pane = `<pane${cols ? ` xSplit="${cols}"` : ''}${rows ? ` ySplit="${rows}"` : ''} topLeftCell="${ref}" activePane="${rows && cols ? 'bottomRight' : rows ? 'bottomLeft' : 'topRight'}" state="frozen"/>`;
    const view = `<sheetViews><sheetView workbookViewId="0">${pane}</sheetView></sheetViews>`;
    if (/<sheetViews>[\s\S]*?<\/sheetViews>/.test(xml)) return xml.replace(/<sheetViews>[\s\S]*?<\/sheetViews>/, view);
    if (/<sheetViews\s*\/>/.test(xml)) return xml.replace(/<sheetViews\s*\/>/, view);
    const at = xml.search(/<sheetFormatPr[\s/>]|<cols[\s>]|<sheetData[\s/>]/);
    return at < 0 ? xml : xml.slice(0, at) + view + xml.slice(at);
  }

  /**
   * patch(bytes, { sheets: [{ conditionalFormatting, …, cellStyles, freeze, tables: [xml], extLst }], dxfs: [xml], styles: xml })
   * styles replaces styles.xml; each sheet's cellStyles maps a cell to its index in it.
   */
  async function patch(bytes, { sheets = [], dxfs = [], styles: stylesXml = null } = {}) {
    if (!sheets.some(s => s && Object.values(s).some(Boolean)) && !dxfs.length && !stylesXml) return bytes;
    const zip = await JSZip.loadAsync(bytes);
    const rels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
    const wbXml = await zip.file('xl/workbook.xml').async('string');
    const target = {};
    for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
      const id = (m[0].match(/\bId="([^"]+)"/) || [])[1], t = (m[0].match(/\bTarget="([^"]+)"/) || [])[1];
      if (id && t) target[id] = t.startsWith('/') ? t.slice(1) : 'xl/' + t;
    }
    const paths = [...wbXml.matchAll(/<sheet\b[^>]*>/g)].map(m => target[(m[0].match(/\br:id="([^"]+)"/) || [])[1]]);
    let drawingN = 0, chartN = 0, tableN = 0;
    const overrides = [];
    let styles = stylesXml; // styles.xml: given, or loaded when needed
    const xfClones = new Map();
    let xfs = null;
    for (let i = 0; i < sheets.length; i++) {
      const add = sheets[i];
      if (!add || !paths[i] || !zip.file(paths[i])) continue;
      let xml = await zip.file(paths[i]).async('string');
      for (const tag of ORDER) if (typeof add[tag] === 'string' && add[tag]) xml = insertInSheet(xml, tag, add[tag]);
      // Every cell gets its index in the new style table; cells that only carry formatting are added.
      if (add.cellStyles) {
        const want = add.cellStyles;
        const seen = new Set();
        xml = xml.replace(/<c r="([A-Z]+\d+)"([^>]*?)(\/?)>/g, (all, ref, attrs, close) => {
          seen.add(ref);
          const a = attrs.replace(/\ss="\d+"/, '');
          const i = want.get(ref);
          return `<c r="${ref}"${a}${i ? ` s="${i}"` : ''}${close}>`;
        });
        const missing = [];
        for (const [ref, i] of want) if (i && !seen.has(ref)) missing.push([ref, `<c r="${ref}" s="${i}"/>`]);
        xml = insertCells(xml, missing);
        // Whole-row styles SheetJS may have written would point into the old table.
        xml = xml.replace(/(<row\b[^>]*?)\s(s|customFormat)="\d+"/g, '$1');
        xml = xml.replace(/(<row\b[^>]*?)\s(s|customFormat)="\d+"/g, '$1');
      }
      if (add.freeze) xml = freezePane(xml, add.freeze.rows || 0, add.freeze.cols || 0);
      // Unlocked cells (for protected sheets): a copy of each cell's style with <protection locked="0"/>.
      if (add.unlocked && add.unlocked.length) {
        if (!styles) styles = await zip.file('xl/styles.xml').async('string');
        if (!xfs) { const m = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(styles); xfs = m ? m[1].match(/<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g) || [] : []; }
        const unlockedXf = i => {
          if (!xfClones.has(i)) {
            let x = (xfs[i] || xfs[0] || '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>').replace(/<protection\b[^>]*\/>/, '').replace(/\sapplyProtection="[^"]*"/, '');
            x = x.endsWith('/>') ? x.slice(0, -2) + ' applyProtection="1"><protection locked="0"/></xf>'
              : x.replace(/^<xf\b/, '<xf applyProtection="1"').replace(/(<extLst\b|<\/xf>)/, '<protection locked="0"/>$1');
            xfs.push(x);
            xfClones.set(i, xfs.length - 1);
          }
          return xfClones.get(i);
        };
        const want = new Set(add.unlocked.slice(0, 100000));
        xml = xml.replace(/<c r="([A-Z]+\d+)"([^>]*?)(\/?)>/g, (all, ref, attrs, close) => {
          if (!want.has(ref)) return all;
          want.delete(ref);
          const sm = /\ss="(\d+)"/.exec(attrs);
          const nx = unlockedXf(sm ? +sm[1] : 0);
          return `<c r="${ref}"${sm ? attrs.replace(/\ss="\d+"/, ` s="${nx}"`) : attrs + ` s="${nx}"`}${close}>`;
        });
        // Empty unlocked cells get a cell element in their row (rows and cells in order).
        const byRow = new Map();
        for (const ref of want) { const m = /^([A-Z]+)(\d+)$/.exec(ref); const r = +m[2]; if (!byRow.has(r)) byRow.set(r, []); byRow.get(r).push(ref); }
        const colNum = ref => [...ref.replace(/\d+$/, '')].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
        if (/<sheetData\s*\/>/.test(xml)) xml = xml.replace(/<sheetData\s*\/>/, '<sheetData></sheetData>');
        for (const [r, refs] of [...byRow].sort((a, b) => a[0] - b[0])) {
          const cells = refs.sort((a, b) => colNum(a) - colNum(b)).map(ref => ({ ref, x: `<c r="${ref}" s="${unlockedXf(0)}"/>` }));
          const rowRe = new RegExp(`<row r="${r}"([^>]*?)(\\/?)>`);
          const rm = rowRe.exec(xml);
          if (rm && rm[2] === '/') { xml = xml.replace(rowRe, `<row r="${r}"${rm[1]}>${cells.map(c => c.x).join('')}</row>`); continue; }
          if (rm) {
            const start = rm.index + rm[0].length, end = xml.indexOf('</row>', start);
            let body = xml.slice(start, end);
            for (const c of cells) {
              const after = [...body.matchAll(/<c r="([A-Z]+)\d+"/g)].find(m => colNum(m[1]) > colNum(c.ref));
              body = after ? body.slice(0, after.index) + c.x + body.slice(after.index) : body + c.x;
            }
            xml = xml.slice(0, start) + body + xml.slice(end);
            continue;
          }
          const next = [...xml.matchAll(/<row r="(\d+)"/g)].find(m => +m[1] > r);
          const rowXml = `<row r="${r}">${cells.map(c => c.x).join('')}</row>`;
          const at = next ? next.index : xml.indexOf('</sheetData>');
          xml = xml.slice(0, at) + rowXml + xml.slice(at);
        }
      }
      // Charts: a drawing part for the sheet, one chart part per chart, and the relationships between them.
      if (add.charts && add.charts.parts.length) {
        const dPath = `xl/drawings/drawing${++drawingN}.xml`;
        zip.file(dPath, add.charts.drawing);
        overrides.push(`<Override PartName="/${dPath}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`);
        const dRels = add.charts.parts.map((part, j) => {
          const cPath = `xl/charts/chart${++chartN}.xml`;
          zip.file(cPath, part);
          overrides.push(`<Override PartName="/${cPath}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`);
          return `<Relationship Id="rId${j + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${chartN}.xml"/>`;
        });
        zip.file(`xl/drawings/_rels/drawing${drawingN}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${dRels.join('')}</Relationships>`);
        const relsPath = paths[i].replace(/([^/]+)$/, '_rels/$1.rels');
        const rel = `<Relationship Id="rIdAtlasDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${drawingN}.xml"/>`;
        const existing = zip.file(relsPath) ? await zip.file(relsPath).async('string') : null;
        zip.file(relsPath, existing ? existing.replace('</Relationships>', rel + '</Relationships>')
          : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel}</Relationships>`);
        if (!/<worksheet\b[^>]*xmlns:r=/.test(xml)) xml = xml.replace(/<worksheet\b/, '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
        xml = insertInSheet(xml, 'drawing', '<drawing r:id="rIdAtlasDrawing"/>');
      }
      // Tables (Format as Table): one part each, linked from the sheet.
      if (Array.isArray(add.tables) && add.tables.length) {
        const relsPath = paths[i].replace(/([^/]+)$/, '_rels/$1.rels');
        const rels = [], parts = [];
        for (const tx of add.tables) {
          const n = ++tableN, tPath = `xl/tables/table${n}.xml`;
          zip.file(tPath, tx);
          overrides.push(`<Override PartName="/${tPath}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`);
          rels.push(`<Relationship Id="rIdAtlasTable${n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table${n}.xml"/>`);
          parts.push(`<tablePart r:id="rIdAtlasTable${n}"/>`);
        }
        const existing = zip.file(relsPath) ? await zip.file(relsPath).async('string') : null;
        zip.file(relsPath, existing ? existing.replace('</Relationships>', rels.join('') + '</Relationships>')
          : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`);
        if (!/<worksheet\b[^>]*xmlns:r=/.test(xml)) xml = xml.replace(/<worksheet\b/, '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
        xml = insertInSheet(xml, 'tableParts', `<tableParts count="${parts.length}">${parts.join('')}</tableParts>`);
      }
      zip.file(paths[i], xml);
    }
    if (overrides.length) {
      const ct = await zip.file('[Content_Types].xml').async('string');
      zip.file('[Content_Types].xml', ct.replace('</Types>', overrides.join('') + '</Types>'));
    }
    if (xfs && xfClones.size) styles = styles.replace(/<cellXfs\b[^>]*>[\s\S]*?<\/cellXfs>/, `<cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs>`);
    if (styles) zip.file('xl/styles.xml', styles);
    if (dxfs.length) {
      let st = await zip.file('xl/styles.xml').async('string');
      const block = `<dxfs count="${dxfs.length}">${dxfs.join('')}</dxfs>`;
      if (/<dxfs\b[^>]*\/>/.test(st)) st = st.replace(/<dxfs\b[^>]*\/>/, block);
      else if (/<dxfs\b[\s\S]*?<\/dxfs>/.test(st)) st = st.replace(/<dxfs\b[\s\S]*?<\/dxfs>/, block);
      else st = st.replace(/(<tableStyles\b|<colors\b|<extLst\b|<\/styleSheet>)/, block + '$1');
      zip.file('xl/styles.xml', st);
    }
    return new Uint8Array(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
  }

  window.XlsxPatch = { patch, insertInSheet, esc };
})();
