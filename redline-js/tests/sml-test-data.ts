// Copyright (c) Microsoft. All rights reserved.
// Licensed under MIT license. See LICENSE file in project root for full license information.

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
  format?: CellFormatSpec;
  comment?: CommentSpec;
  dataValidation?: DataValidationSpec;
}

export interface CellFormatSpec {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  fontSize?: number;
  fontColor?: string;
  fillForegroundColor?: string;
}

export interface CommentSpec {
  author: string;
  text: string;
}

export interface DataValidationSpec {
  type: string;
  operator?: string;
  formula1?: string;
}

/**
 * Generate a minimal .xlsx file from specification
 */
export async function generateExcelFile(spec: ExcelFileSpec): Promise<Buffer> {
  const zip = new JSZip();

  const contentTypesXml = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`;
  zip.file('[Content_Types].xml', contentTypesXml);

  const workbookRelsXml = spec.sheets.map((sheet, i) => {
    const id = `rId${i + 1}`;
    return `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`;
  }).join('\n');

  zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${workbookRelsXml}
</Relationships>`);

  const workbookXml = `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  ${spec.sheets.map((sheet, i) => `<sheet name="${sheet.name}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('\n')}
</workbook>`;
  zip.file('xl/workbook.xml', workbookXml);

  const stylesXml = `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts>
    <numFmt numFmtId="164" formatCode="General"/>
  </numFmts>
  <fonts>
    <font><sz val="11"/><color rgb="00000000"/><name val="Calibri"/></font>
  </fonts>
  <fills>
    <fill><patternFill patternType="none"/></fill>
  </fills>
  <borders>
    <border/>
  </borders>
  <cellXfs>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellXfs>
</styleSheet>`;
  zip.file('xl/styles.xml', stylesXml);

  for (let i = 0; i < spec.sheets.length; i++) {
    const sheet = spec.sheets[i];
    const rowsXml = sheet.rows.map(row => {
      const cellsXml = row.cells.map(cell => {
        const cellContent = cell.value
          ? `<v>${cell.value}</v>`
          : `<f>${cell.formula}</f>`;
        return `<c r="${cell.address}">${cellContent}</c>`;
      }).join('\n');
      return `  <row r="${row.index}">\n${cellsXml}\n  </row>`;
    }).join('\n');

    const worksheetXml = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
${rowsXml}
  </sheetData>
</worksheet>`;
    zip.file(`xl/worksheets/sheet${i + 1}.xml`, worksheetXml);
  }

  for (let i = 0; i < spec.sheets.length; i++) {
    const sheetRelsXml = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`;
    zip.file(`xl/worksheets/_rels/sheet${i + 1}.xml.rels`, sheetRelsXml);
  }

  return await zip.generateAsync({ type: 'nodebuffer' });
}

export { generateExcelFile, type ExcelFileSpec, type SheetSpec, type RowSpec, type CellSpec };
