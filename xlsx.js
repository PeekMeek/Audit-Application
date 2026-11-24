/* Builds the "Additive Application" workbook (.xlsx) in the browser with JSZip.
   One form tab per application in the template layout, plus a Summary tab.
   Works in the browser (window.AAXlsx) and in Node (module.exports) for testing. */
(function (root) {
  'use strict';
  const F = root.AAFields || (typeof require !== 'undefined' ? require('./fields.js') : null);

  const esc = s => String(s == null ? '' : s)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const colLetter = i => { let s = ''; i++; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
  const ref = (r, c) => colLetter(c - 1) + r;

  /* ---------- styles ---------- */
  const LABEL_FILL = 'FFDDE1F5', HEADER_FILL = 'FF282156';
  const URG_FILLS = Array.from({ length: 10 }, (_, i) => 'FF' + F.urgencyHex(i + 1, 68));
  const S = { title: 1, label: 2, value: 3, header: 4, cell: 5, link: 6, section: 7, urg0: 8, num: 18, backlink: 19 };
  function stylesXml() {
    const fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>',
      `<fill><patternFill patternType="solid"><fgColor rgb="${LABEL_FILL}"/><bgColor indexed="64"/></patternFill></fill>`,
      `<fill><patternFill patternType="solid"><fgColor rgb="${HEADER_FILL}"/><bgColor indexed="64"/></patternFill></fill>`]
      .concat(URG_FILLS.map(c => `<fill><patternFill patternType="solid"><fgColor rgb="${c}"/><bgColor indexed="64"/></patternFill></fill>`));
    const thin = '<left style="thin"><color rgb="FF000000"/></left><right style="thin"><color rgb="FF000000"/></right><top style="thin"><color rgb="FF000000"/></top><bottom style="thin"><color rgb="FF000000"/></bottom><diagonal/>';
    const xf = (font, fill, border, align) => `<xf numFmtId="0" fontId="${font}" fillId="${fill}" borderId="${border}" xfId="0"${font ? ' applyFont="1"' : ''}${fill ? ' applyFill="1"' : ''}${border ? ' applyBorder="1"' : ''}${align ? ' applyAlignment="1"><alignment ' + align + '/></xf>' : '/>'}`;
    const xfs = [
      xf(0, 0, 0),
      xf(2, 0, 1, 'horizontal="center" vertical="center"'),       // 1 title
      xf(0, 2, 1, 'vertical="top" wrapText="1"'),                 // 2 label
      xf(0, 0, 1, 'vertical="top" wrapText="1"'),                 // 3 value
      xf(3, 3, 1, 'horizontal="center" vertical="center" wrapText="1"'), // 4 header
      xf(0, 0, 1, 'vertical="top" wrapText="1"'),                 // 5 cell
      xf(4, 0, 1, 'vertical="top"'),                              // 6 link
      xf(5, 2, 1, 'vertical="center"')                            // 7 section
    ];
    for (let i = 0; i < 10; i++) xfs.push(xf(1, 4 + i, 1, 'horizontal="center" vertical="top"')); // 8..17 urgency
    xfs.push(xf(0, 0, 1, 'horizontal="center" vertical="top"'));  // 18 number
    xfs.push(xf(4, 0, 0, 'vertical="center"'));                   // 19 back link
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="6"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="16"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font><font><u/><sz val="11"/><color rgb="FF0563C1"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="12"/><name val="Calibri"/><family val="2"/></font></fonts>
<fills count="${fills.length}">${fills.join('')}</fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border>${thin}</border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;
  }

  /* ---------- sheet model ---------- */
  class Sheet {
    constructor(name) { this.name = name; this.cells = {}; this.heights = {}; this.cols = []; this.merges = []; this.links = []; this.images = []; this.freezeRow = 0; this.filter = null; this.fitWidth = false; }
    set(r, c, v, s) { (this.cells[r] = this.cells[r] || {})[c] = { v, s }; }
    xml() {
      const rows = Object.keys(this.cells).map(Number).concat(Object.keys(this.heights).map(Number))
        .filter((x, i, a) => a.indexOf(x) === i).sort((a, b) => a - b);
      const cellXml = (r, c, { v, s }) => {
        const at = ` r="${ref(r, c)}"${s ? ` s="${s}"` : ''}`;
        if (v == null || v === '') return `<c${at}/>`;
        if (typeof v === 'number') return `<c${at}><v>${v}</v></c>`;
        if (Array.isArray(v)) return `<c${at} t="inlineStr"><is>${v.map(run => `<r><rPr>${run.b ? '<b/>' : ''}${run.i ? '<i/>' : ''}<sz val="11"/><rFont val="Calibri"/><family val="2"/></rPr><t xml:space="preserve">${esc(run.t)}</t></r>`).join('')}</is></c>`;
        return `<c${at} t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`;
      };
      const data = rows.map(r => {
        const cs = this.cells[r] || {};
        const h = this.heights[r] ? ` ht="${this.heights[r]}" customHeight="1"` : '';
        return `<row r="${r}"${h}>${Object.keys(cs).map(Number).sort((a, b) => a - b).map(c => cellXml(r, c, cs[c])).join('')}</row>`;
      }).join('');
      const pane = this.freezeRow ? `<pane ySplit="${this.freezeRow}" topLeftCell="A${this.freezeRow + 1}" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A${this.freezeRow + 1}" sqref="A${this.freezeRow + 1}"/>` : '';
      return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        (this.fitWidth ? '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' : '') +
        `<sheetViews><sheetView workbookViewId="0"${this.index === 0 ? ' tabSelected="1"' : ''}>${pane}</sheetView></sheetViews>` +
        '<sheetFormatPr defaultRowHeight="15"/>' +
        (this.cols.length ? `<cols>${this.cols.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('')}</cols>` : '') +
        `<sheetData>${data}</sheetData>` +
        (this.filter ? `<autoFilter ref="${this.filter}"/>` : '') +
        (this.merges.length ? `<mergeCells count="${this.merges.length}">${this.merges.map(m => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>` : '') +
        (this.links.length ? `<hyperlinks>${this.links.map(l => `<hyperlink ref="${l.ref}" location="${esc(l.location)}" display="${esc(l.display)}"/>`).join('')}</hyperlinks>` : '') +
        '<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>' +
        `<pageSetup orientation="${this.landscape ? 'landscape' : 'portrait'}"${this.fitWidth ? ' fitToWidth="1" fitToHeight="0"' : ''}/>` +
        (this.images.length ? '<drawing r:id="rId1"/>' : '') +
        '</worksheet>';
    }
  }

  // Rough row height for wrapped text so rows aren't cut off (Excel doesn't auto-fit generated rows).
  function linesFor(text, width) {
    const cpl = Math.max(8, Math.floor(width * 1.1));
    return String(text || '').split('\n').reduce((n, p) => n + Math.max(1, Math.ceil(p.length / cpl)), 0);
  }
  const heightFor = (lines) => Math.min(409, Math.max(20, lines * 15 + 5));

  function sheetNameFor(i, appName, used) {
    let base = `${String(i + 1).padStart(2, '0')} ${(appName || 'Untitled').replace(/[\[\]:*?\/\\]/g, ' ').replace(/\s+/g, ' ').trim()}`;
    base = base.slice(0, 31).trim().replace(/^'+|'+$/g, '');
    let name = base, n = 2;
    while (used.has(name.toLowerCase()) || name.toLowerCase() === 'summary') { const suf = ` (${n++})`; name = base.slice(0, 31 - suf.length) + suf; }
    used.add(name.toLowerCase());
    return name;
  }
  const quoteSheet = n => `'${n.replace(/'/g, "''")}'!A1`;
  function fmtDate(iso) {
    if (!iso) return '';
    const d = new Date(iso), p = n => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
  }
  function fitImage(img, maxW, maxH) {
    const w = img.w || 800, h = img.h || 600, k = Math.min(1, maxW / w, maxH / h);
    return { w: Math.round(w * k), h: Math.round(h * k) };
  }

  /* ---------- form tab (template layout) ---------- */
  const WA = 46, WB = 92, PXB = WB * 7 + 5;
  function buildForm(entry, name, withBackLink) {
    const d = F.derive(entry);
    const sh = new Sheet(name);
    sh.cols = [WA, WB, 14];
    sh.fitWidth = true;
    sh.set(1, 1, 'Additive Application', S.title); sh.set(1, 2, '', S.title);
    sh.merges.push('A1:B1'); sh.heights[1] = 26;
    if (withBackLink) { sh.set(1, 3, '← Summary', S.backlink); sh.links.push({ ref: 'C1', location: "'Summary'!A1", display: '← Summary' }); }
    let r = 2;
    for (const [key, label, hint] of F.TEMPLATE_ROWS) {
      sh.set(r, 1, hint ? [{ t: label, b: true }, { t: hint, i: true }] : [{ t: label, b: true }], S.label);
      if (key === 'image') {
        if (entry.partPhoto) {
          const sz = fitImage(entry.partPhoto, PXB - 16, 340);
          sh.images.push({ photo: entry.partPhoto, row: r - 1, col: 1, w: sz.w, h: sz.h });
          sh.set(r, 2, '', S.value); sh.heights[r] = Math.round((sz.h + 12) * 0.75);
        } else { sh.set(r, 2, '(No part photo)', S.value); sh.heights[r] = 20; }
      } else if (key === 'urgency') {
        sh.set(r, 2, d.urgency || '', d.urgency ? S.urg0 + d.urgency - 1 : S.value);
        sh.heights[r] = 20;
      } else {
        const v = d[key] || '';
        sh.set(r, 2, v, S.value);
        sh.heights[r] = heightFor(Math.max(linesFor(label + hint, WA), linesFor(v, WB)));
      }
      r++;
    }
    // Notebook details, as captured
    r++;
    sh.set(r, 1, 'Notebook Details (as captured)', S.section); sh.set(r, 2, '', S.section); sh.merges.push(`A${r}:B${r}`); sh.heights[r] = 22; r++;
    const meta = [['Captured', fmtDate(entry.createdAt)], ['Last edited', fmtDate(entry.updatedAt)], ['Posted to Teams', fmtDate(entry.postedAt) || 'Not posted']];
    for (const [l, v] of meta) { sh.set(r, 1, [{ t: l, b: true }], S.label); sh.set(r, 2, v, S.value); sh.heights[r] = 20; r++; }
    for (const [title, list] of [['Required', F.REQUIRED], ['Supporting', F.SUPPORTING]]) {
      sh.set(r, 1, [{ t: title, b: true, i: true }], S.label); sh.set(r, 2, '', S.label); sh.merges.push(`A${r}:B${r}`); sh.heights[r] = 20; r++;
      list.forEach((f, i) => {
        let v = (entry.fields || {})[f.key] || '';
        if (f.key === 'photo') v = entry.partPhoto ? 'Part photo attached (see Image above)' : (v || '');
        sh.set(r, 1, `${F.fieldNumber(f.key)}. ${f.label}`, S.label); sh.set(r, 2, v, S.value);
        sh.heights[r] = heightFor(Math.max(linesFor(f.label, WA), linesFor(v, WB))); r++;
      });
    }
    if (entry.otherNotes) { sh.set(r, 1, [{ t: 'Other notes', b: true }], S.label); sh.set(r, 2, entry.otherNotes, S.value); sh.heights[r] = heightFor(linesFor(entry.otherNotes, WB)); r++; }
    (entry.notebookPhotos || []).forEach((p, i) => {
      const sz = fitImage(p, PXB - 16, 520);
      sh.set(r, 1, [{ t: `Notebook page ${i + 1}`, b: true }], S.label); sh.set(r, 2, '', S.value);
      sh.images.push({ photo: p, row: r - 1, col: 1, w: sz.w, h: sz.h });
      sh.heights[r] = Math.round((sz.h + 12) * 0.75); r++;
    });
    return sh;
  }

  /* ---------- summary tab ---------- */
  function buildSummary(entries, names) {
    const sh = new Sheet('Summary');
    const cols = [['#', 5], ['Captured', 17], ['Originator', 18], ['Department/Area/Product Line', 22], ['Workstream Owner', 18], ['Application Name', 30],
      ['Urgency (1-10)', 11], ['Confidence (1-10)', 12], ['EAU', 12], ['Current Issue/Challenge', 50], ['Posted to Teams', 17], ['Form', 10]];
    sh.cols = cols.map(c => c[1]);
    sh.landscape = true; sh.fitWidth = true;
    cols.forEach((c, i) => sh.set(1, i + 1, c[0], S.header));
    sh.heights[1] = 32; sh.freezeRow = 1;
    entries.forEach((e, i) => {
      const d = F.derive(e), r = i + 2;
      const vals = [i + 1, fmtDate(e.createdAt), d.originator, d.department, d.workstreamOwner, d.appName, d.urgency || '', d.confidence || '', F.clean((e.fields || {}).eau), d.issue, fmtDate(e.postedAt) || 'No'];
      vals.forEach((v, c) => {
        let s = S.cell;
        if (c === 0 || c === 7) s = S.num;
        if (c === 6) s = d.urgency ? S.urg0 + d.urgency - 1 : S.num;
        sh.set(r, c + 1, v, s);
      });
      sh.set(r, 12, 'Open', S.link); sh.links.push({ ref: `L${r}`, location: quoteSheet(names[i]), display: 'Open' });
      sh.heights[r] = heightFor(Math.max(linesFor(d.issue, 50), linesFor(d.appName, 30), linesFor(d.department, 22)));
    });
    const last = Math.max(2, entries.length + 1);
    sh.filter = `A1:L${last}`;
    return sh;
  }

  /* ---------- assemble the package ---------- */
  function drawingXml(images, firstImageNo) {
    const EMU = 9525;
    const anchors = images.map((im, i) => `<xdr:oneCellAnchor><xdr:from><xdr:col>${im.col}</xdr:col><xdr:colOff>${6 * EMU}</xdr:colOff><xdr:row>${im.row}</xdr:row><xdr:rowOff>${6 * EMU}</xdr:rowOff></xdr:from><xdr:ext cx="${im.w * EMU}" cy="${im.h * EMU}"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${i + 2}" name="Picture ${firstImageNo + i}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId${i + 1}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${im.w * EMU}" cy="${im.h * EMU}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${anchors}</xdr:wsDr>`;
  }
  const rels = list => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${list.map(([id, type, target]) => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>`).join('')}</Relationships>`;
  function photoBase64(p) {
    const s = p.dataUrl || '';
    const i = s.indexOf(',');
    return i >= 0 ? s.slice(i + 1) : s;
  }

  /**
   * entries: array of saved entries (oldest first).
   * opts.summary (default true): include the Summary tab.
   * returns a Promise of a Blob (browser) or Uint8Array (Node, opts.type='uint8array').
   */
  function buildWorkbook(entries, opts) {
    opts = opts || {};
    const JSZip = opts.JSZip || root.JSZip;
    const withSummary = opts.summary !== false;
    const used = new Set();
    const names = entries.map((e, i) => sheetNameFor(i, F.clean((e.fields || {}).appName), used));
    const sheets = [];
    if (withSummary) sheets.push(buildSummary(entries, names));
    entries.forEach((e, i) => sheets.push(buildForm(e, names[i], withSummary)));
    sheets.forEach((s, i) => { s.index = i; });

    const zip = new JSZip();
    const overrides = [], wbRels = [];
    let imgNo = 1, drawNo = 1;
    sheets.forEach((sh, i) => {
      const n = i + 1;
      zip.file(`xl/worksheets/sheet${n}.xml`, sh.xml());
      overrides.push(`<Override PartName="/xl/worksheets/sheet${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`);
      wbRels.push([`rId${n}`, 'worksheet', `worksheets/sheet${n}.xml`]);
      if (sh.images.length) {
        const dn = drawNo++;
        zip.file(`xl/worksheets/_rels/sheet${n}.xml.rels`, rels([['rId1', 'drawing', `../drawings/drawing${dn}.xml`]]));
        const imgRels = [];
        sh.images.forEach((im, k) => {
          const file = `image${imgNo + k}.jpeg`;
          zip.file(`xl/media/${file}`, photoBase64(im.photo), { base64: true });
          imgRels.push([`rId${k + 1}`, 'image', `../media/${file}`]);
        });
        zip.file(`xl/drawings/drawing${dn}.xml`, drawingXml(sh.images, imgNo));
        zip.file(`xl/drawings/_rels/drawing${dn}.xml.rels`, rels(imgRels));
        overrides.push(`<Override PartName="/xl/drawings/drawing${dn}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`);
        imgNo += sh.images.length;
      }
    });
    wbRels.push([`rId${sheets.length + 1}`, 'styles', 'styles.xml']);
    const summarySheet = withSummary ? sheets[0] : null;
    const defined = summarySheet && summarySheet.filter
      ? `<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">'Summary'!$A$1:$L$${summarySheet.filter.split(':L')[1]}</definedName></definedNames>` : '';
    zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView activeTab="0"/></bookViews><sheets>${sheets.map((s, i) => `<sheet name="${esc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets>${defined}</workbook>`);
    zip.file('xl/_rels/workbook.xml.rels', rels(wbRels));
    zip.file('xl/styles.xml', stylesXml());
    const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    zip.file('docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Additive Applications</dc:title><dc:creator>Additive Applications app</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`);
    zip.file('_rels/.rels', rels([['rId1', 'officeDocument', 'xl/workbook.xml']]).replace('</Relationships>',
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>'));
    zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>${overrides.join('')}</Types>`);
    return zip.generateAsync({
      type: opts.type || 'blob', compression: 'DEFLATE',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
  }

  const api = { buildWorkbook };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.AAXlsx = api;
})(typeof window !== 'undefined' ? window : globalThis);
