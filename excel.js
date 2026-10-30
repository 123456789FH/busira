/* Basirah XLSX export. No network calls or external dependencies.
 * Writes a small, standards-based SpreadsheetML package in a ZIP container.
 * All user text is inlineStr (never a formula); numeric values stay numeric.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.BasirahExcel = factory();
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';
  const MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  const letters = ['أ', 'ب', 'ج', 'د', 'هـ'];
  const encoder = new TextEncoder();

  function escapeXML(value) {
    return String(value ?? '')
      .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/g, '')
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&apos;');
  }
  // OOXML uses _xHHHH_ escapes. Protect literal sequences before encoding XML.
  function textXML(value) {
    const text = String(value ?? '');
    if (text.length > 32767) throw Error('يوجد نص أطول من الحد المسموح في خلية Excel.');
    return escapeXML(text.replace(/_x[0-9a-fA-F]{4}_/g, match => '_x005F_' + match.slice(1)));
  }
  function column(index) {
    let name = '';
    for (let n = index + 1; n; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + (n - 1) % 26) + name;
    return name;
  }
  const number = (value, style = Number.isInteger(value) ? 7 : 4) => ({ value, style });
  function dateCell(value) {
    const d = new Date(value);
    if (!Number.isFinite(d.getTime())) return String(value ?? '');
    // Export the same local calendar date displayed by the application.
    return number(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / 86400000 + 25569, 6);
  }
  function cell(value, reference, style = 0) {
    if (value && typeof value === 'object') { style = value.style; value = value.value; }
    if (typeof value === 'number') {
      if (!Number.isFinite(value)) throw Error('تعذّر تصدير قيمة رقمية غير صالحة.');
      return `<c r="${reference}" s="${style}" t="n"><v>${value}</v></c>`;
    }
    return `<c r="${reference}" s="${style}" t="inlineStr"><is><t xml:space="preserve">${textXML(value)}</t></is></c>`;
  }

  function worksheet(sheet, quiz, count) {
    if (sheet.rows.length > 1048571) throw Error('عدد النتائج يتجاوز سعة ورقة Excel.');
    const lastColumn = column(sheet.headers.length - 1), lastRow = sheet.rows.length + 5;
    const titleEnd = column(Math.min(sheet.headers.length, 8) - 1);
    const metadata = [sheet.name, `الاختبار: ${quiz.title}     الفصل: ${quiz.className || '—'}`, `عدد الأوراق المعتمدة: ${count}     ${sheet.note}`];
    let data = metadata.map((value, i) => `<row r="${i + 1}" ht="${i ? 40 : 34}" customHeight="1">${cell(value, 'A' + (i + 1), i ? 2 : 1)}</row>`).join('');
    data += `<row r="5" ht="42" customHeight="1">${sheet.headers.map((label, c) => cell(label, column(c) + '5', 3)).join('')}</row>`;
    data += sheet.rows.map((values, r) => {
      const height = Math.min(409, Math.max(28, ...values.map((value, c) => typeof value === 'string' ? Math.ceil(value.length / Math.max(8, sheet.widths[c] - 3)) * 19 : 28)));
      return `<row r="${r + 6}" ht="${height}" customHeight="1">${values.map((value, c) => cell(value, column(c) + (r + 6))).join('')}</row>`;
    }).join('');
    return XML + `<worksheet xmlns="${NS}"><dimension ref="A1:${lastColumn}${lastRow}"/>` +
      '<sheetViews><sheetView rightToLeft="1" showGridLines="0" workbookViewId="0"><pane ySplit="5" topLeftCell="A6" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A6" sqref="A6"/></sheetView></sheetViews>' +
      '<sheetFormatPr defaultRowHeight="28"/><cols>' + sheet.widths.map((width, i) => `<col min="${i + 1}" max="${i + 1}" width="${width}" customWidth="1"/>`).join('') + '</cols>' +
      `<sheetData>${data}</sheetData><autoFilter ref="A5:${lastColumn}${lastRow}"/>` +
      `<mergeCells count="3">${[1, 2, 3].map(r => `<mergeCell ref="A${r}:${titleEnd}${r}"/>`).join('')}</mergeCells>` +
      '<pageMargins left="0.3" right="0.3" top="0.4" bottom="0.4" header="0.2" footer="0.2"/></worksheet>';
  }

  const styles = XML + `<styleSheet xmlns="${NS}">
    <numFmts count="3"><numFmt numFmtId="164" formatCode="0.0%"/><numFmt numFmtId="165" formatCode="yyyy-mm-dd"/><numFmt numFmtId="166" formatCode="0.##"/></numFmts>
    <fonts count="4">
      <font><sz val="12"/><color rgb="FF203A43"/><name val="Arial"/></font>
      <font><b/><sz val="20"/><color rgb="FF006C67"/><name val="Arial"/></font>
      <font><sz val="11"/><color rgb="FF506773"/><name val="Arial"/></font>
      <font><b/><sz val="12"/><color rgb="FFFFFFFF"/><name val="Arial"/></font>
    </fonts>
    <fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF006C67"/><bgColor indexed="64"/></patternFill></fill></fills>
    <borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left/><right/><top/><bottom style="thin"><color rgb="FFDCE9E7"/></bottom><diagonal/></border></borders>
    <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
    <cellXfs count="8">
      <xf numFmtId="49" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="right" vertical="center" wrapText="1" readingOrder="2"/></xf>
      <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="right" vertical="center" readingOrder="2"/></xf>
      <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="right" vertical="center" wrapText="1" readingOrder="2"/></xf>
      <xf numFmtId="0" fontId="3" fillId="2" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1" readingOrder="2"/></xf>
      <xf numFmtId="166" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
      <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
      <xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
      <xf numFmtId="1" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    </cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  </styleSheet>`;

  const crcTable = Uint32Array.from({ length: 256 }, (_, n) => {
    for (let k = 0; k < 8; k++) n = n & 1 ? 0xedb88320 ^ (n >>> 1) : n >>> 1;
    return n >>> 0;
  });
  function crc32(bytes) {
    let crc = 0xffffffff;
    for (const b of bytes) crc = crcTable[(crc ^ b) & 255] ^ (crc >>> 8);
    return (crc ^ 0xffffffff) >>> 0;
  }
  function concat(parts) {
    const bytes = new Uint8Array(parts.reduce((sum, p) => sum + p.length, 0));
    let at = 0;
    for (const part of parts) { bytes.set(part, at); at += part.length; }
    return bytes;
  }
  // ZIP method 0 (STORE) works offline, including browsers without CompressionStream.
  function zip(files) {
    const local = [], central = [];
    let offset = 0;
    for (const [path, text] of files) {
      const name = encoder.encode(path), content = encoder.encode(text), crc = crc32(content);
      const header = new Uint8Array(30), h = new DataView(header.buffer);
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true);
      h.setUint16(12, 33, true); // valid DOS date: 1980-01-01
      h.setUint32(14, crc, true); h.setUint32(18, content.length, true); h.setUint32(22, content.length, true); h.setUint16(26, name.length, true);
      local.push(header, name, content);
      const directory = new Uint8Array(46), d = new DataView(directory.buffer);
      d.setUint32(0, 0x02014b50, true); d.setUint16(4, 20, true); d.setUint16(6, 20, true); d.setUint16(8, 0x0800, true); d.setUint16(14, 33, true);
      d.setUint32(16, crc, true); d.setUint32(20, content.length, true); d.setUint32(24, content.length, true); d.setUint16(28, name.length, true); d.setUint32(42, offset, true);
      central.push(directory, name);
      offset += header.length + name.length + content.length;
      if (offset > 0xffffffff) throw Error('حجم التصدير أكبر من الحد المسموح.');
    }
    const directory = concat(central), end = new Uint8Array(22), e = new DataView(end.buffer);
    e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true);
    e.setUint32(12, directory.length, true); e.setUint32(16, offset, true);
    return concat([...local, directory, end]);
  }

  function build(quiz, analysis) {
    if (!analysis.rows.length) throw Error('لا توجد نتائج معتمدة لهذا الاختبار.');
    const sheets = [
      { name: 'درجات الطلبة', note: 'الدرجات والإجابات المحفوظة للاختبار المحدد.',
        headers: ['الاسم', 'الرمز', 'الدرجة', 'الدرجة الكلية', 'النسبة', 'تاريخ الاعتماد', ...quiz.questions.map((_, i) => 'س' + (i + 1))],
        widths: [30, 16, 14, 18, 16, 22, ...quiz.questions.map(() => 8)],
        rows: analysis.rows.map(r => [r.name, String(r.studentId ?? ''), number(r.score), number(r.total), number(r.percent / 100, 5), dateCell(r.created), ...quiz.questions.map((_, i) => r.answers[i] == null ? 'فارغ' : letters[r.answers[i]])]) },
      { name: 'تحليل الأسئلة', note: 'خطأ: إجابة محددة غير صحيحة. فارغ: بلا إجابة معتمدة، ويشمل التظليل المتعدد المعتمد كفارغ.',
        headers: ['السؤال', 'المهارة', 'الإجابة الصحيحة', 'درجة السؤال', 'صحيح', 'خطأ', 'فارغ', 'نسبة الصحة'],
        widths: [12, 38, 20, 18, 12, 12, 12, 18],
        rows: analysis.items.map((r, i) => [number(i + 1, 7), r.skill, letters[quiz.questions[i].answer], number(quiz.questions[i].weight), number(r.correct, 7), number(analysis.rows.length - r.correct - r.blank, 7), number(r.blank, 7), number(r.percent / 100, 5)]) },
      { name: 'تحليل المهارات', note: 'النسبة = الدرجات المحققة ÷ الدرجات الممكنة لجميع الأوراق المعتمدة، مرجّحة بدرجات الأسئلة.',
        headers: ['المهارة', 'عدد الأسئلة', 'الدرجات المحققة', 'الدرجات الممكنة', 'نسبة الإتقان'],
        widths: [38, 18, 24, 24, 20],
        rows: analysis.skills.map(s => [s.name, number(s.questions, 7), number(s.earned), number(s.total), number(s.percent / 100, 5)]) }
    ];
    const files = [
      ['[Content_Types].xml', XML + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' + sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('') + '</Types>'],
      ['_rels/.rels', XML + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
      ['xl/workbook.xml', XML + `<workbook xmlns="${NS}" xmlns:r="${REL}"><bookViews><workbookView activeTab="0"/></bookViews><sheets>` + sheets.map((s, i) => `<sheet name="${escapeXML(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('') + '</sheets></workbook>'],
      ['xl/_rels/workbook.xml.rels', XML + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="${REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('') + `<Relationship Id="rId4" Type="${REL}/styles" Target="styles.xml"/></Relationships>`],
      ['xl/styles.xml', styles],
      ...sheets.map((s, i) => [`xl/worksheets/sheet${i + 1}.xml`, worksheet(s, quiz, analysis.rows.length)])
    ];
    return zip(files);
  }
  return { build, MIME };
});
