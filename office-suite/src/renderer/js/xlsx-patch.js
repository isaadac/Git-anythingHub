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
  async function patch(bytes, { sheets = [], dxfs = [] } = {}) {
    if (!sheets.some(s => s && Object.values(s).some(Boolean)) && !dxfs.length) return bytes;
    const zip = await JSZip.loadAsync(bytes);
    const rels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
    const wbXml = await zip.file('xl/workbook.xml').async('string');
    const target = {};
    for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
      const id = (m[0].match(/\bId="([^"]+)"/) || [])[1], t = (m[0].match(/\bTarget="([^"]+)"/) || [])[1];
      if (id && t) target[id] = t.startsWith('/') ? t.slice(1) : 'xl/' + t;
    }
    const paths = [...wbXml.matchAll(/<sheet\b[^>]*>/g)].map(m => target[(m[0].match(/\br:id="([^"]+)"/) || [])[1]]);
    let drawingN = 0, chartN = 0;
    const overrides = [];
    for (let i = 0; i < sheets.length; i++) {
      const add = sheets[i];
      if (!add || !paths[i] || !zip.file(paths[i])) continue;
      let xml = await zip.file(paths[i]).async('string');
      for (const tag of ORDER) if (typeof add[tag] === 'string' && add[tag]) xml = insertInSheet(xml, tag, add[tag]);
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
      zip.file(paths[i], xml);
    }
    if (overrides.length) {
      const ct = await zip.file('[Content_Types].xml').async('string');
      zip.file('[Content_Types].xml', ct.replace('</Types>', overrides.join('') + '</Types>'));
    }
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
