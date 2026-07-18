(function (global) {
  'use strict';

  const encoder = new TextEncoder();
  const MIME_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const STYLE = {
    default: 0,
    title: 1,
    subtitle: 2,
    section: 3,
    label: 4,
    value: 5,
    currency: 6,
    currencyNegative: 7,
    date: 8,
    tableHeader: 9,
    tableText: 10,
    totalLabel: 11,
    totalCurrency: 12,
    note: 13
  };

  function xmlEscape(value) {
    return String(value ?? '')
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  function columnName(index) {
    let value = index + 1;
    let result = '';
    while (value > 0) {
      value -= 1;
      result = String.fromCharCode(65 + (value % 26)) + result;
      value = Math.floor(value / 26);
    }
    return result;
  }

  function excelDate(value) {
    const date = value instanceof Date ? value : new Date(value || '');
    if (Number.isNaN(date.getTime())) return null;
    return (Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()) / 86400000) + 25569;
  }

  function cellXml(rowIndex, columnIndex, cell) {
    const reference = `${columnName(columnIndex)}${rowIndex + 1}`;
    const config = cell && typeof cell === 'object' && !Array.isArray(cell) && !(cell instanceof Date)
      ? cell
      : { value: cell };
    const style = Number.isInteger(config.style) ? ` s="${config.style}"` : '';
    const value = config.value;

    if (config.formula) {
      const cached = Number(config.cachedValue ?? value ?? 0) || 0;
      return `<c r="${reference}"${style}><f>${xmlEscape(config.formula)}</f><v>${cached}</v></c>`;
    }
    if (value === null || value === undefined || value === '') {
      return `<c r="${reference}"${style}/>`;
    }
    if (value instanceof Date || config.type === 'date') {
      const serial = excelDate(value);
      return serial === null
        ? `<c r="${reference}"${style}/>`
        : `<c r="${reference}"${style}><v>${serial}</v></c>`;
    }
    if (typeof value === 'number' && Number.isFinite(value)) {
      return `<c r="${reference}"${style}><v>${value}</v></c>`;
    }
    if (typeof value === 'boolean') {
      return `<c r="${reference}"${style} t="b"><v>${value ? 1 : 0}</v></c>`;
    }
    return `<c r="${reference}"${style} t="inlineStr"><is><t xml:space="preserve">${xmlEscape(value)}</t></is></c>`;
  }

  function sheetXml(rows, options = {}) {
    const columnWidths = options.columnWidths || [];
    const maxColumns = Math.max(columnWidths.length, ...rows.map((row) => row.length), 1);
    const maxRows = Math.max(rows.length, 1);
    const cols = columnWidths.length
      ? `<cols>${columnWidths.map((width, index) => `<col min="${index + 1}" max="${index + 1}" width="${width}" customWidth="1"/>`).join('')}</cols>`
      : '';
    const rowXml = rows.map((row, rowIndex) => {
      const cells = row.map((cell, columnIndex) => cellXml(rowIndex, columnIndex, cell)).join('');
      const height = options.rowHeights?.[rowIndex];
      const rowStyle = height ? ` ht="${height}" customHeight="1"` : '';
      return `<row r="${rowIndex + 1}"${rowStyle}>${cells}</row>`;
    }).join('');
    const merges = options.merges?.length
      ? `<mergeCells count="${options.merges.length}">${options.merges.map((range) => `<mergeCell ref="${range}"/>`).join('')}</mergeCells>`
      : '';
    const tabSelected = options.tabSelected ? ' tabSelected="1"' : '';
    const pane = options.freezeRows
      ? `<sheetViews><sheetView workbookViewId="0" showGridLines="${options.showGridLines === false ? 0 : 1}"${tabSelected}><pane ySplit="${options.freezeRows}" topLeftCell="A${options.freezeRows + 1}" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>`
      : `<sheetViews><sheetView workbookViewId="0" showGridLines="${options.showGridLines === false ? 0 : 1}"${tabSelected}/></sheetViews>`;
    const autoFilter = options.autoFilter ? `<autoFilter ref="${options.autoFilter}"/>` : '';
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${pane}<dimension ref="A1:${columnName(maxColumns - 1)}${maxRows}"/>${cols}<sheetData>${rowXml}</sheetData>${autoFilter}${merges}<pageMargins left="0.35" right="0.35" top="0.5" bottom="0.5" header="0.2" footer="0.2"/><pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/></worksheet>`;
  }

  function stylesXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="2"><numFmt numFmtId="164" formatCode="&quot;AED&quot; #,##0.00;[Red]-&quot;AED&quot; #,##0.00;&quot;AED&quot; -"/><numFmt numFmtId="165" formatCode="yyyy-mm-dd"/></numFmts>
  <fonts count="7">
    <font><sz val="11"/><name val="Aptos"/><family val="2"/></font>
    <font><b/><sz val="20"/><color rgb="FFFFFFFF"/><name val="Aptos Display"/></font>
    <font><b/><sz val="11"/><color rgb="FF111115"/><name val="Aptos"/></font>
    <font><b/><sz val="11"/><color rgb="FFC1FF72"/><name val="Aptos"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Aptos"/></font>
    <font><b/><sz val="11"/><color rgb="FF111115"/><name val="Aptos"/></font>
    <font><i/><sz val="10"/><color rgb="FF666666"/><name val="Aptos"/></font>
  </fonts>
  <fills count="7">
    <fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF111115"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFC1FF72"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF1F3EE"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF7F8F4"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFECEC"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="4">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left/><right/><top/><bottom style="thin"><color rgb="FFD9DDD4"/></bottom><diagonal/></border>
    <border><left/><right/><top style="medium"><color rgb="FFC1FF72"/></top><bottom/><diagonal/></border>
    <border><left style="thin"><color rgb="FFD9DDD4"/></left><right style="thin"><color rgb="FFD9DDD4"/></right><top style="thin"><color rgb="FFD9DDD4"/></top><bottom style="thin"><color rgb="FFD9DDD4"/></bottom><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="14">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="2" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="5" borderId="1" xfId="0" applyFill="1" applyBorder="1"/>
    <xf numFmtId="164" fontId="0" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="164" fontId="0" fillId="6" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="165" fontId="0" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="4" fillId="2" borderId="3" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="5" fillId="3" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="164" fontId="5" fillId="3" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="0" fontId="6" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment wrapText="1"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>`;
  }

  function workbookXml(sheetNames, activeSheetIndex = 0) {
    const activeTab = Math.max(0, Math.min(sheetNames.length - 1, Math.round(Number(activeSheetIndex) || 0)));
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="24000" windowHeight="14000" activeTab="${activeTab}"/></bookViews><sheets>${sheetNames.map((name, index) => `<sheet name="${xmlEscape(name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`).join('')}</sheets><calcPr calcId="191029" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>`;
  }

  function workbookRelationshipsXml(sheetCount) {
    const sheets = Array.from({ length: sheetCount }, (_, index) => `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets}<Relationship Id="rId${sheetCount + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
  }

  function contentTypesXml(sheetCount) {
    const sheets = Array.from({ length: sheetCount }, (_, index) => `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${sheets}<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`;
  }

  function crc32(bytes) {
    let crc = 0xFFFFFFFF;
    for (let index = 0; index < bytes.length; index += 1) {
      crc ^= bytes[index];
      for (let bit = 0; bit < 8; bit += 1) {
        crc = (crc >>> 1) ^ (0xEDB88320 & -(crc & 1));
      }
    }
    return (crc ^ 0xFFFFFFFF) >>> 0;
  }

  function uint16(value) {
    const bytes = new Uint8Array(2);
    new DataView(bytes.buffer).setUint16(0, value, true);
    return bytes;
  }

  function uint32(value) {
    const bytes = new Uint8Array(4);
    new DataView(bytes.buffer).setUint32(0, value >>> 0, true);
    return bytes;
  }

  function concat(parts) {
    const total = parts.reduce((sum, part) => sum + part.length, 0);
    const output = new Uint8Array(total);
    let offset = 0;
    parts.forEach((part) => {
      output.set(part, offset);
      offset += part.length;
    });
    return output;
  }

  function zipStore(entries) {
    const localParts = [];
    const centralParts = [];
    let localOffset = 0;
    const now = new Date();
    const time = (now.getHours() << 11) | (now.getMinutes() << 5) | Math.floor(now.getSeconds() / 2);
    const date = ((Math.max(now.getFullYear(), 1980) - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();

    entries.forEach(([name, content]) => {
      const nameBytes = encoder.encode(name);
      const data = typeof content === 'string' ? encoder.encode(content) : content;
      const checksum = crc32(data);
      const localHeader = concat([
        uint32(0x04034B50), uint16(20), uint16(0x0800), uint16(0), uint16(time), uint16(date),
        uint32(checksum), uint32(data.length), uint32(data.length), uint16(nameBytes.length), uint16(0), nameBytes
      ]);
      localParts.push(localHeader, data);

      const centralHeader = concat([
        uint32(0x02014B50), uint16(20), uint16(20), uint16(0x0800), uint16(0), uint16(time), uint16(date),
        uint32(checksum), uint32(data.length), uint32(data.length), uint16(nameBytes.length), uint16(0), uint16(0),
        uint16(0), uint16(0), uint32(0), uint32(localOffset), nameBytes
      ]);
      centralParts.push(centralHeader);
      localOffset += localHeader.length + data.length;
    });

    const centralDirectory = concat(centralParts);
    const end = concat([
      uint32(0x06054B50), uint16(0), uint16(0), uint16(entries.length), uint16(entries.length),
      uint32(centralDirectory.length), uint32(localOffset), uint16(0)
    ]);
    return concat([...localParts, centralDirectory, end]);
  }

  function createMonthlyHistoryWorkbook(report) {
    const activities = Array.isArray(report.activities) ? report.activities : [];
    const receipts = Array.isArray(report.receipts) ? report.receipts : [];
    const contributionTotal = activities.filter((item) => item.type === 'contribution').reduce((sum, item) => sum + (Number(item.amount) || 0), 0);
    const expenseTotal = activities.filter((item) => item.type === 'expense').reduce((sum, item) => sum + (Number(item.amount) || 0), 0);
    const netTotal = contributionTotal - expenseTotal;
    const firstDataRow = 2;
    const lastDataRow = Math.max(firstDataRow, activities.length + 1);
    const generatedAt = report.generatedAt instanceof Date ? report.generatedAt : new Date();
    const preferredSheetIndex = Math.max(0, Math.min(2, Math.round(Number(report.activeSheetIndex) || 0)));

    const summaryRows = [
      [{ value: 'BudgetIQ Monthly History', style: STYLE.title }, null, null, null],
      [{ value: `${report.month || 'Monthly'} shared-fund Excel report`, style: STYLE.subtitle }, null, null, null],
      [],
      [{ value: 'Key totals', style: STYLE.section }, null, null, null],
      [{ value: 'Activities', style: STYLE.label }, { value: activities.length, style: STYLE.value }],
      [{ value: 'Contributions', style: STYLE.label }, { formula: `SUM('Transactions'!H${firstDataRow}:H${lastDataRow})`, cachedValue: contributionTotal, style: STYLE.currency }],
      [{ value: 'Expenses', style: STYLE.label }, { formula: `SUM('Transactions'!I${firstDataRow}:I${lastDataRow})`, cachedValue: expenseTotal, style: STYLE.currencyNegative }],
      [{ value: 'Net balance', style: STYLE.totalLabel }, { formula: `SUM('Transactions'!J${firstDataRow}:J${lastDataRow})`, cachedValue: netTotal, style: STYLE.totalCurrency }],
      [],
      [{ value: 'Report details', style: STYLE.section }, null, null, null],
      [{ value: 'Month', style: STYLE.label }, { value: report.month || 'Monthly', style: STYLE.value }],
      [{ value: 'Scope', style: STYLE.label }, { value: report.scope || 'All fund members', style: STYLE.value }],
      [{ value: 'Generated by', style: STYLE.label }, { value: report.owner || 'BudgetIQ', style: STYLE.value }],
      [{ value: 'Generated on', style: STYLE.label }, { value: generatedAt, type: 'date', style: STYLE.date }],
      [{ value: 'Receipt records', style: STYLE.label }, { value: receipts.length, style: STYLE.value }],
      [],
      [{ value: 'How to use', style: STYLE.section }, null, null, null],
      [{ value: 'Filter and sort the Transactions sheet to review contributions and expenses. Totals on this Summary sheet update automatically in Excel.', style: STYLE.note }, null, null, null]
    ];

    const transactionRows = [[
      'Date', 'Activity', 'Person', 'Type', 'Category / Fund', 'Reference', 'Status', 'Contribution (AED)', 'Expense (AED)', 'Net (AED)'
    ].map((value) => ({ value, style: STYLE.tableHeader }))];
    if (activities.length) {
      activities.forEach((activity) => {
        const isExpense = activity.type === 'expense';
        const amount = Number(activity.amount) || 0;
        transactionRows.push([
          { value: new Date(activity.createdAt || ''), type: 'date', style: STYLE.date },
          { value: activity.title || 'Activity', style: STYLE.tableText },
          { value: activity.person || 'Not recorded', style: STYLE.tableText },
          { value: isExpense ? 'Expense' : 'Contribution', style: STYLE.tableText },
          { value: activity.detail || '', style: STYLE.tableText },
          { value: activity.reference || '', style: STYLE.tableText },
          { value: activity.status || 'Recorded', style: STYLE.tableText },
          { value: isExpense ? 0 : amount, style: STYLE.currency },
          { value: isExpense ? Math.abs(amount) : 0, style: STYLE.currencyNegative },
          { value: isExpense ? -Math.abs(amount) : amount, style: isExpense ? STYLE.currencyNegative : STYLE.currency }
        ]);
      });
    } else {
      transactionRows.push([
        { value: null, style: STYLE.date }, { value: 'No activity recorded for this month', style: STYLE.tableText },
        { value: '', style: STYLE.tableText }, { value: '', style: STYLE.tableText }, { value: '', style: STYLE.tableText },
        { value: '', style: STYLE.tableText }, { value: '', style: STYLE.tableText },
        { value: 0, style: STYLE.currency }, { value: 0, style: STYLE.currency }, { value: 0, style: STYLE.currency }
      ]);
    }

    transactionRows.push([
      { value: 'TOTAL', style: STYLE.totalLabel }, { value: '', style: STYLE.totalLabel }, { value: '', style: STYLE.totalLabel },
      { value: '', style: STYLE.totalLabel }, { value: '', style: STYLE.totalLabel }, { value: '', style: STYLE.totalLabel },
      { value: '', style: STYLE.totalLabel }, { formula: `SUM(H${firstDataRow}:H${lastDataRow})`, cachedValue: contributionTotal, style: STYLE.totalCurrency },
      { formula: `SUM(I${firstDataRow}:I${lastDataRow})`, cachedValue: expenseTotal, style: STYLE.totalCurrency },
      { formula: `SUM(J${firstDataRow}:J${lastDataRow})`, cachedValue: netTotal, style: STYLE.totalCurrency }
    ]);

    const receiptRows = [[
      '#', 'Expense', 'Purchased By', 'Category', 'Date', 'Reference', 'Amount (AED)', 'Receipt File', 'Note'
    ].map((value) => ({ value, style: STYLE.tableHeader }))];
    if (receipts.length) {
      receipts.forEach((receipt, index) => receiptRows.push([
        { value: index + 1, style: STYLE.tableText },
        { value: receipt.name || 'Saved Expense', style: STYLE.tableText },
        { value: receipt.buyer || 'Not recorded', style: STYLE.tableText },
        { value: receipt.category || 'Other', style: STYLE.tableText },
        { value: new Date(receipt.createdAt || ''), type: 'date', style: STYLE.date },
        { value: receipt.reference || '', style: STYLE.tableText },
        { value: Number(receipt.amount) || 0, style: STYLE.currency },
        { value: receipt.receiptName || `Receipt ${index + 1}`, style: STYLE.tableText },
        { value: receipt.note || '', style: STYLE.tableText }
      ]));
    } else {
      receiptRows.push([{ value: 1, style: STYLE.tableText }, { value: 'No receipt pictures recorded for this month', style: STYLE.tableText }]);
    }
    receiptRows.push([{ value: 'Receipt pictures stay securely inside BudgetIQ and the PDF report; this sheet records their searchable details.', style: STYLE.note }]);
    const receiptNoteRow = receiptRows.length;

    let sheets = [
      ['Summary', sheetXml(summaryRows, { columnWidths: [24, 24, 18, 18], rowHeights: { 0: 34, 1: 24, 3: 22, 9: 22, 16: 22, 17: 36 }, merges: ['A1:D1', 'A2:D2', 'A4:D4', 'A10:D10', 'A17:D17', 'A18:D18'], showGridLines: false, tabSelected: preferredSheetIndex === 0 })],
      ['Transactions', sheetXml(transactionRows, { columnWidths: [14, 28, 20, 15, 20, 19, 15, 20, 18, 17], rowHeights: { 0: 30 }, freezeRows: 1, autoFilter: `A1:J${lastDataRow}`, showGridLines: false, tabSelected: preferredSheetIndex === 1 })],
      ['Receipts', sheetXml(receiptRows, { columnWidths: [7, 28, 20, 16, 14, 19, 17, 24, 36], rowHeights: { 0: 30, [receiptNoteRow - 1]: 32 }, merges: [`A${receiptNoteRow}:I${receiptNoteRow}`], freezeRows: 1, autoFilter: `A1:I${Math.max(2, receipts.length + 1)}`, showGridLines: false, tabSelected: preferredSheetIndex === 2 })]
    ];
    if (preferredSheetIndex > 0) {
      sheets = [sheets[preferredSheetIndex], ...sheets.filter((_, index) => index !== preferredSheetIndex)];
    }

    const timestamp = generatedAt.toISOString();
    const entries = [
      ['[Content_Types].xml', contentTypesXml(sheets.length)],
      ['_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`],
      ['docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>BudgetIQ ${xmlEscape(report.month || 'Monthly')} History</dc:title><dc:creator>BudgetIQ</dc:creator><cp:lastModifiedBy>BudgetIQ</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">${timestamp}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${timestamp}</dcterms:modified></cp:coreProperties>`],
      ['docProps/app.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>BudgetIQ</Application><AppVersion>1.0</AppVersion></Properties>`],
      ['xl/workbook.xml', workbookXml(sheets.map(([name]) => name), 0)],
      ['xl/_rels/workbook.xml.rels', workbookRelationshipsXml(sheets.length)],
      ['xl/styles.xml', stylesXml()],
      ...sheets.map(([, xml], index) => [`xl/worksheets/sheet${index + 1}.xml`, xml])
    ];

    return new Blob([zipStore(entries)], { type: MIME_TYPE });
  }

  global.BudgetIQExcel = Object.freeze({ MIME_TYPE, createMonthlyHistoryWorkbook });
})(window);
