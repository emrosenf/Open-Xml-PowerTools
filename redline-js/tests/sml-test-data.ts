import JSZip from 'jszip';

export interface ExcelFileSpec {
  sheets: SheetSpec[];
}

export interface SheetSpec {
  name: string;
  rows: RowSpec[];
}

export interface RowSpec {
  index: number;
  cells: CellSpec[];
}

export interface CellSpec {
  address: string;
  value?: string;
  formula?: string;
}

export async function generateExcelFile(spec: ExcelFileSpec): Promise<Buffer> {
  const zip = new JSZip();

  const sheetOverrides = spec.sheets.map((_, i) =>
    `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
  ).join('\n  ');

  const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  ${sheetOverrides}
</Types>`;
  zip.file('[Content_Types].xml', contentTypesXml);

  const rootRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
  zip.file('_rels/.rels', rootRelsXml);

  const workbookRelsEntries = spec.sheets.map((_, i) =>
    `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`
  ).join('\n  ');

  const workbookRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${workbookRelsEntries}
  <Relationship Id="rId${spec.sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
  zip.file('xl/_rels/workbook.xml.rels', workbookRelsXml);

  const sheetEntries = spec.sheets.map((sheet, i) =>
    `<sheet name="${sheet.name}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`
  ).join('\n    ');

  const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    ${sheetEntries}
  </sheets>
</workbook>`;
  zip.file('xl/workbook.xml', workbookXml);

  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border/>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
</styleSheet>`;
  zip.file('xl/styles.xml', stylesXml);

  for (let i = 0; i < spec.sheets.length; i++) {
    const sheet = spec.sheets[i];
    const rowsXml = sheet.rows.map(row => {
      const cellsXml = row.cells.map(cell => {
        if (cell.formula) {
          return `<c r="${cell.address}"><f>${escapeXml(cell.formula)}</f></c>`;
        }
        return `<c r="${cell.address}" t="str"><v>${escapeXml(cell.value || '')}</v></c>`;
      }).join('');
      return `<row r="${row.index}">${cellsXml}</row>`;
    }).join('\n    ');

    const worksheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    ${rowsXml}
  </sheetData>
</worksheet>`;
    zip.file(`xl/worksheets/sheet${i + 1}.xml`, worksheetXml);
  }

  return await zip.generateAsync({ type: 'nodebuffer' });
}

function escapeXml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}
