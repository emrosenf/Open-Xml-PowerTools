// Copyright (c) Microsoft. All rights reserved.
// Licensed under MIT license. See LICENSE file in the project root for full license information.

/**
 * SML Markup Renderer
 *
 * Renders a marked workbook showing differences between two Excel spreadsheets.
 * Adds highlight styles, cell comments, and a summary sheet.
 */

import {
  type OoxmlPackage,
  clonePackage,
  savePackage,
  getPartAsXml,
  getPartAsString,
  setPartFromString,
  setPartFromXml,
  getRelationships,
} from '../core/package';
import {
  parseXml,
  buildXml,
  addXmlDeclaration,
  getTagName,
  getChildren,
  getAttribute,
  setAttribute,
  createNode,
  type XmlNode,
} from '../core/xml';
import {
  SmlChangeType,
  type SmlChange,
  type SmlComparerSettings,
  type SmlComparisonResult,
  type HighlightColors,
} from './types';

/**
 * Track style IDs for highlighted cells
 */
interface HighlightStyles {
  addedFillId: number;
  modifiedValueFillId: number;
  modifiedFormulaFillId: number;
  modifiedFormatFillId: number;
  addedStyleId: number;
  modifiedValueStyleId: number;
  modifiedFormulaStyleId: number;
  modifiedFormatStyleId: number;
}

/**
 * Default highlight colors (ARGB hex without #)
 */
const DEFAULT_COLORS: Required<HighlightColors> = {
  addedCellColor: '90EE90', // Light green
  deletedCellColor: 'FFCCCB', // Light red
  modifiedValueColor: 'FFD700', // Gold
  modifiedFormulaColor: '87CEEB', // Sky blue
  modifiedFormatColor: 'E6E6FA', // Lavender
  insertedRowColor: 'E0FFFF', // Light cyan
  deletedRowColor: 'FFE4E1', // Misty rose
  namedRangeChangeColor: 'DDA0DD', // Light purple
  commentChangeColor: 'FFFACD', // Light yellow
  dataValidationChangeColor: 'FFDAB9', // Light orange
  mergedCellRangeColor: 'D3D3D3', // Light gray
};

/**
 * Render a marked workbook with all differences highlighted.
 *
 * @param pkg The newer workbook package to mark up
 * @param result The comparison result with all changes
 * @param settings Comparison settings including colors
 * @returns Buffer containing the marked workbook
 */
export async function renderMarkedWorkbook(
  pkg: OoxmlPackage,
  result: SmlComparisonResult,
  settings: SmlComparerSettings
): Promise<Buffer> {
  const markedPkg = await clonePackage(pkg);
  const colors = { ...DEFAULT_COLORS, ...settings.highlightColors };
  const highlightStyles = await addHighlightStyles(markedPkg, colors);
  const changesBySheet = groupChangesBySheet(result.changes);

  for (const [sheetName, changes] of changesBySheet) {
    await applySheetHighlights(markedPkg, sheetName, changes, highlightStyles, settings);
  }

  await addDiffSummarySheet(markedPkg, result, settings);
  return savePackage(markedPkg);
}

/**
 * Group changes by sheet name for processing
 */
function groupChangesBySheet(changes: SmlChange[]): Map<string, SmlChange[]> {
  const grouped = new Map<string, SmlChange[]>();

  for (const change of changes) {
    let sheetName = change.sheetName ?? change.cellAddress;

    if (
      change.changeType === SmlChangeType.SheetAdded ||
      change.changeType === SmlChangeType.SheetDeleted ||
      change.changeType === SmlChangeType.SheetRenamed
    ) {
      continue;
    }

    if (sheetName && sheetName.includes('!')) {
      const parts = sheetName.split('!');
      sheetName = parts[0].replace(/^'/, '').replace(/'$/, '');
    }

    if (!sheetName) continue;

    if (!grouped.has(sheetName)) {
      grouped.set(sheetName, []);
    }
    grouped.get(sheetName)!.push(change);
  }

  return grouped;
}

/**
 * Add highlight fill styles to the workbook stylesheet
 */
async function addHighlightStyles(
  pkg: OoxmlPackage,
  colors: Required<HighlightColors>
): Promise<HighlightStyles> {
  const stylesXml = await getPartAsXml(pkg, 'xl/styles.xml');
  if (!stylesXml) {
    throw new Error('Invalid Excel workbook: missing xl/styles.xml');
  }

  let stylesheetNode: XmlNode | null = null;
  for (const node of stylesXml) {
    if (getTagName(node) === 'styleSheet') {
      stylesheetNode = node;
      break;
    }
  }

  if (!stylesheetNode) {
    throw new Error('Invalid styles.xml: missing styleSheet element');
  }

  const children = getChildren(stylesheetNode);
  let fillsNode: XmlNode | null = null;
  let cellXfsNode: XmlNode | null = null;

  for (const child of children) {
    const tagName = getTagName(child);
    if (tagName === 'fills') {
      fillsNode = child;
    } else if (tagName === 'cellXfs') {
      cellXfsNode = child;
    }
  }

  const styles: HighlightStyles = {
    addedFillId: 0,
    modifiedValueFillId: 0,
    modifiedFormulaFillId: 0,
    modifiedFormatFillId: 0,
    addedStyleId: 0,
    modifiedValueStyleId: 0,
    modifiedFormulaStyleId: 0,
    modifiedFormatStyleId: 0,
  };

  if (fillsNode) {
    const fills = getChildren(fillsNode);
    let fillCount = fills.length;

    styles.addedFillId = fillCount++;
    fills.push(createSolidFill(colors.addedCellColor));

    styles.modifiedValueFillId = fillCount++;
    fills.push(createSolidFill(colors.modifiedValueColor));

    styles.modifiedFormulaFillId = fillCount++;
    fills.push(createSolidFill(colors.modifiedFormulaColor));

    styles.modifiedFormatFillId = fillCount++;
    fills.push(createSolidFill(colors.modifiedFormatColor));

    setAttribute(fillsNode, 'count', String(fillCount));
  }

  if (cellXfsNode) {
    const xfs = getChildren(cellXfsNode);
    let xfCount = xfs.length;

    styles.addedStyleId = xfCount++;
    xfs.push(createXfWithFill(styles.addedFillId));

    styles.modifiedValueStyleId = xfCount++;
    xfs.push(createXfWithFill(styles.modifiedValueFillId));

    styles.modifiedFormulaStyleId = xfCount++;
    xfs.push(createXfWithFill(styles.modifiedFormulaFillId));

    styles.modifiedFormatStyleId = xfCount++;
    xfs.push(createXfWithFill(styles.modifiedFormatFillId));

    setAttribute(cellXfsNode, 'count', String(xfCount));
  }

  setPartFromXml(pkg, 'xl/styles.xml', stylesXml);

  return styles;
}

/**
 * Create a solid fill element
 */
function createSolidFill(color: string): XmlNode {
  const argbColor = color.length === 6 ? `FF${color}` : color;

  return {
    fill: [
      {
        patternFill: [
          {
            fgColor: [],
            ':@': { '@_rgb': argbColor },
          },
          {
            bgColor: [],
            ':@': { '@_indexed': '64' },
          },
        ],
        ':@': { '@_patternType': 'solid' },
      },
    ],
  };
}

/**
 * Create a cell format (xf) element with a fill reference
 */
function createXfWithFill(fillId: number): XmlNode {
  return {
    xf: [],
    ':@': {
      '@_numFmtId': '0',
      '@_fontId': '0',
      '@_fillId': String(fillId),
      '@_borderId': '0',
      '@_applyFill': '1',
    },
  };
}

/**
 * Apply highlights to a specific worksheet
 */
async function applySheetHighlights(
  pkg: OoxmlPackage,
  sheetName: string,
  changes: SmlChange[],
  styles: HighlightStyles,
  settings: SmlComparerSettings
): Promise<void> {
  const worksheetPath = await findWorksheetPath(pkg, sheetName);
  if (!worksheetPath) return;

  const wsXml = await getPartAsXml(pkg, worksheetPath);
  if (!wsXml) return;

  let sheetDataNode: XmlNode | null = null;
  for (const node of wsXml) {
    if (getTagName(node) === 'worksheet') {
      for (const child of getChildren(node)) {
        if (getTagName(child) === 'sheetData') {
          sheetDataNode = child;
          break;
        }
      }
    }
  }

  if (!sheetDataNode) return;

  for (const change of changes) {
    if (!change.cellAddress) continue;

    let cellRef = change.cellAddress;
    if (cellRef.includes('!')) {
      cellRef = cellRef.split('!')[1];
    }

    const styleId = getStyleIdForChange(change, styles);
    if (styleId >= 0) {
      applyCellStyle(sheetDataNode, cellRef, styleId);
    }
  }

  setPartFromXml(pkg, worksheetPath, wsXml);
  await addCommentsForChanges(pkg, worksheetPath, changes, settings);
}

/**
 * Get the style ID for a change type
 */
function getStyleIdForChange(change: SmlChange, styles: HighlightStyles): number {
  switch (change.changeType) {
    case SmlChangeType.CellAdded:
      return styles.addedStyleId;
    case SmlChangeType.ValueChanged:
      return styles.modifiedValueStyleId;
    case SmlChangeType.FormulaChanged:
      return styles.modifiedFormulaStyleId;
    case SmlChangeType.FormatChanged:
      return styles.modifiedFormatStyleId;
    default:
      return -1;
  }
}

/**
 * Apply a style to a cell in the sheet data
 */
function applyCellStyle(sheetDataNode: XmlNode, cellRef: string, styleId: number): void {
  const { col, row } = parseCellRef(cellRef);
  const rows = getChildren(sheetDataNode);

  let rowNode = rows.find(
    (r) => getTagName(r) === 'row' && getAttribute(r, 'r') === String(row)
  );

  if (!rowNode) {
    rowNode = createNode('row', { r: String(row) }, []);
    insertRowInOrder(rows, rowNode, row);
  }

  const cells = getChildren(rowNode);
  let cellNode = cells.find(
    (c) => getTagName(c) === 'c' && getAttribute(c, 'r') === cellRef
  );

  if (!cellNode) {
    cellNode = createNode('c', { r: cellRef }, []);
    insertCellInOrder(cells, cellNode, col);
  }

  setAttribute(cellNode, 's', String(styleId));
}

/**
 * Insert a row node in the correct position (ordered by row number)
 */
function insertRowInOrder(rows: XmlNode[], newRow: XmlNode, rowNum: number): void {
  let insertIdx = rows.length;
  for (let i = 0; i < rows.length; i++) {
    const existingRow = parseInt(getAttribute(rows[i], 'r') || '0', 10);
    if (existingRow > rowNum) {
      insertIdx = i;
      break;
    }
  }
  rows.splice(insertIdx, 0, newRow);
}

/**
 * Insert a cell node in the correct position (ordered by column)
 */
function insertCellInOrder(cells: XmlNode[], newCell: XmlNode, colNum: number): void {
  let insertIdx = cells.length;
  for (let i = 0; i < cells.length; i++) {
    const cellRef = getAttribute(cells[i], 'r') || '';
    const { col } = parseCellRef(cellRef);
    if (col > colNum) {
      insertIdx = i;
      break;
    }
  }
  cells.splice(insertIdx, 0, newCell);
}

/**
 * Parse a cell reference (e.g., "A1") into column and row numbers
 */
function parseCellRef(cellRef: string): { col: number; row: number } {
  let col = 0;
  let i = 0;

  while (i < cellRef.length && /[A-Za-z]/.test(cellRef[i])) {
    col = col * 26 + (cellRef[i].toUpperCase().charCodeAt(0) - 64);
    i++;
  }

  const row = parseInt(cellRef.slice(i), 10);
  return { col, row };
}

/**
 * Get column letter from column number (1-indexed)
 */
function getColumnLetter(col: number): string {
  let result = '';
  let n = col;
  while (n > 0) {
    n--;
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26);
  }
  return result;
}

/**
 * Find the worksheet path for a given sheet name
 */
async function findWorksheetPath(
  pkg: OoxmlPackage,
  sheetName: string
): Promise<string | null> {
  const workbookXml = await getPartAsXml(pkg, 'xl/workbook.xml');
  if (!workbookXml) return null;

  for (const node of workbookXml) {
    if (getTagName(node) === 'workbook') {
      for (const child of getChildren(node)) {
        if (getTagName(child) === 'sheets') {
          for (const sheet of getChildren(child)) {
            if (getTagName(sheet) === 'sheet' && getAttribute(sheet, 'name') === sheetName) {
              const rId = getAttribute(sheet, 'r:id');
              if (rId) {
                const rels = await getRelationships(pkg, 'xl/workbook.xml');
                const rel = rels.find((r) => r.id === rId);
                if (rel) {
                  return `xl/${rel.target}`;
                }
              }
            }
          }
        }
      }
    }
  }

  return null;
}

/**
 * Add comments to worksheet for changes
 */
async function addCommentsForChanges(
  pkg: OoxmlPackage,
  worksheetPath: string,
  changes: SmlChange[],
  settings: SmlComparerSettings
): Promise<void> {
  if (changes.length === 0) return;

  const author = settings.authorForChanges || 'redline-js';
  const commentNodes: XmlNode[] = [];
  const authorId = 0;

  for (const change of changes) {
    if (!change.cellAddress) continue;

    let cellRef = change.cellAddress;
    if (cellRef.includes('!')) {
      cellRef = cellRef.split('!')[1];
    }

    const commentText = buildCommentText(change);
    if (!commentText) continue;

    commentNodes.push({
      comment: [
        {
          text: [
            {
              r: [
                {
                  t: [{ '#text': commentText }],
                },
              ],
            },
          ],
        },
      ],
      ':@': {
        '@_ref': cellRef,
        '@_authorId': String(authorId),
      },
    });
  }

  if (commentNodes.length === 0) return;

  const commentsXml: XmlNode = {
    comments: [
      {
        authors: [
          {
            author: [{ '#text': author }],
          },
        ],
      },
      {
        commentList: commentNodes,
      },
    ],
    ':@': {
      '@_xmlns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    },
  };

  const parts = worksheetPath.split('/');
  parts.pop();
  const dir = parts.join('/');
  const commentsPath = `${dir}/comments1.xml`;

  const commentsContent = addXmlDeclaration(buildXml(commentsXml));
  setPartFromString(pkg, commentsPath, commentsContent);
  await addCommentsRelationship(pkg, worksheetPath, commentsPath);
}

/**
 * Build comment text for a change
 */
function buildCommentText(change: SmlChange): string {
  const lines: string[] = [];
  lines.push(`[${SmlChangeType[change.changeType]}]`);

  switch (change.changeType) {
    case SmlChangeType.CellAdded:
      lines.push(`New value: ${change.newValue || ''}`);
      if (change.newFormula) {
        lines.push(`Formula: =${change.newFormula}`);
      }
      break;

    case SmlChangeType.CellDeleted:
      lines.push(`Deleted value: ${change.oldValue || ''}`);
      break;

    case SmlChangeType.ValueChanged:
      lines.push(`Old value: ${change.oldValue || ''}`);
      lines.push(`New value: ${change.newValue || ''}`);
      break;

    case SmlChangeType.FormulaChanged:
      lines.push(`Old formula: =${change.oldFormula || ''}`);
      lines.push(`New formula: =${change.newFormula || ''}`);
      break;

    case SmlChangeType.FormatChanged:
      lines.push('Formatting changed');
      break;

    default:
      return '';
  }

  return lines.join('\n');
}

/**
 * Add relationship for comments to worksheet
 */
async function addCommentsRelationship(
  pkg: OoxmlPackage,
  worksheetPath: string,
  commentsPath: string
): Promise<void> {
  const parts = worksheetPath.split('/');
  const fileName = parts.pop()!;
  const dir = parts.join('/');
  const relsPath = `${dir}/_rels/${fileName}.rels`;

  // Get existing rels or create new
  let relsContent = await getPartAsString(pkg, relsPath);
  let relsXml: XmlNode[];
  let nextId = 1;

  if (relsContent) {
    relsXml = parseXml(relsContent);
    // Find highest existing Id
    for (const node of relsXml) {
      if (getTagName(node) === 'Relationships') {
        for (const child of getChildren(node)) {
          const id = getAttribute(child, 'Id');
          if (id) {
            const num = parseInt(id.replace('rId', ''), 10);
            if (num >= nextId) nextId = num + 1;
          }
        }
      }
    }
  } else {
    relsXml = [
      {
        Relationships: [],
        ':@': {
          '@_xmlns': 'http://schemas.openxmlformats.org/package/2006/relationships',
        },
      },
    ];
  }

  // Add comments relationship
  const commentsFileName = commentsPath.split('/').pop()!;
  const relNode: XmlNode = {
    Relationship: [],
    ':@': {
      '@_Id': `rId${nextId}`,
      '@_Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
      '@_Target': `../${commentsFileName}`,
    },
  };

  // Add to relationships
  for (const node of relsXml) {
    if (getTagName(node) === 'Relationships') {
      getChildren(node).push(relNode);
    }
  }

  // Save rels
  setPartFromXml(pkg, relsPath, relsXml);

  // Update content types
  await addContentType(pkg, commentsPath);
}

/**
 * Add content type for comments part
 */
async function addContentType(pkg: OoxmlPackage, partPath: string): Promise<void> {
  const contentTypesXml = await getPartAsString(pkg, '[Content_Types].xml');
  if (!contentTypesXml) return;

  const nodes = parseXml(contentTypesXml);

  for (const node of nodes) {
    if (getTagName(node) === 'Types') {
      const children = getChildren(node);

      // Check if override already exists
      const exists = children.some(
        (c) => getAttribute(c, 'PartName') === `/${partPath}`
      );

      if (!exists) {
        children.push({
          Override: [],
          ':@': {
            '@_PartName': `/${partPath}`,
            '@_ContentType':
              'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
          },
        });
      }
    }
  }

  setPartFromXml(pkg, '[Content_Types].xml', nodes);
}

/**
 * Add a summary sheet with all changes
 */
async function addDiffSummarySheet(
  pkg: OoxmlPackage,
  result: SmlComparisonResult,
  _settings: SmlComparerSettings
): Promise<void> {
  // Build summary sheet content
  const rows: XmlNode[] = [];
  let rowNum = 1;

  // Header
  rows.push(createDataRow(rowNum++, ['Spreadsheet Comparison Summary']));
  rows.push(createDataRow(rowNum++, ['']));

  // Statistics
  rows.push(createDataRow(rowNum++, ['Total Changes:', String(result.changes.length)]));

  const stats = computeStats(result);
  rows.push(createDataRow(rowNum++, ['Value Changes:', String(stats.valueChanges)]));
  rows.push(createDataRow(rowNum++, ['Formula Changes:', String(stats.formulaChanges)]));
  rows.push(createDataRow(rowNum++, ['Format Changes:', String(stats.formatChanges)]));
  rows.push(createDataRow(rowNum++, ['Cells Added:', String(stats.cellsAdded)]));
  rows.push(createDataRow(rowNum++, ['Cells Deleted:', String(stats.cellsDeleted)]));
  rows.push(createDataRow(rowNum++, ['Sheets Added:', String(stats.sheetsAdded)]));
  rows.push(createDataRow(rowNum++, ['Sheets Deleted:', String(stats.sheetsDeleted)]));
  rows.push(createDataRow(rowNum++, ['']));

  // Change details header
  rows.push(
    createDataRow(rowNum++, [
      'Change Type',
      'Cell',
      'Old Value',
      'New Value',
      'Description',
    ])
  );

  // Change rows
  for (const change of result.changes) {
    rows.push(
      createDataRow(rowNum++, [
        SmlChangeType[change.changeType],
        change.cellAddress || '',
        change.oldValue || change.oldFormula || '',
        change.newValue || change.newFormula || '',
        getChangeDescription(change),
      ])
    );
  }

  // Create worksheet XML
  const worksheetXml: XmlNode = {
    worksheet: [
      {
        sheetData: rows,
      },
    ],
    ':@': {
      '@_xmlns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      '@_xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    },
  };

  // Add worksheet part
  const sheetPath = 'xl/worksheets/_DiffSummary.xml';
  setPartFromXml(pkg, sheetPath, worksheetXml);

  // Add to workbook
  await addSheetToWorkbook(pkg, '_DiffSummary', sheetPath);

  // Add content type
  await addWorksheetContentType(pkg, sheetPath);
}

/**
 * Create a data row for the summary sheet
 */
function createDataRow(rowNum: number, values: string[]): XmlNode {
  const cells: XmlNode[] = values.map((value, i) => ({
    c: [
      {
        is: [
          {
            t: [{ '#text': value }],
          },
        ],
      },
    ],
    ':@': {
      '@_r': `${getColumnLetter(i + 1)}${rowNum}`,
      '@_t': 'inlineStr',
    },
  }));

  return {
    row: cells,
    ':@': { '@_r': String(rowNum) },
  };
}

/**
 * Compute statistics from changes
 */
function computeStats(result: SmlComparisonResult): {
  valueChanges: number;
  formulaChanges: number;
  formatChanges: number;
  cellsAdded: number;
  cellsDeleted: number;
  sheetsAdded: number;
  sheetsDeleted: number;
} {
  const stats = {
    valueChanges: 0,
    formulaChanges: 0,
    formatChanges: 0,
    cellsAdded: 0,
    cellsDeleted: 0,
    sheetsAdded: 0,
    sheetsDeleted: 0,
  };

  for (const change of result.changes) {
    switch (change.changeType) {
      case SmlChangeType.ValueChanged:
        stats.valueChanges++;
        break;
      case SmlChangeType.FormulaChanged:
        stats.formulaChanges++;
        break;
      case SmlChangeType.FormatChanged:
        stats.formatChanges++;
        break;
      case SmlChangeType.CellAdded:
        stats.cellsAdded++;
        break;
      case SmlChangeType.CellDeleted:
        stats.cellsDeleted++;
        break;
      case SmlChangeType.SheetAdded:
        stats.sheetsAdded++;
        break;
      case SmlChangeType.SheetDeleted:
        stats.sheetsDeleted++;
        break;
    }
  }

  return stats;
}

/**
 * Get human-readable description for a change
 */
function getChangeDescription(change: SmlChange): string {
  switch (change.changeType) {
    case SmlChangeType.SheetAdded:
      return `Sheet added: ${change.newSheetName ?? change.sheetName ?? change.cellAddress}`;
    case SmlChangeType.SheetDeleted:
      return `Sheet deleted: ${change.oldSheetName ?? change.sheetName ?? change.cellAddress}`;
    case SmlChangeType.SheetRenamed:
      return `Sheet renamed: ${change.oldSheetName} → ${change.newSheetName ?? change.cellAddress}`;
    case SmlChangeType.CellAdded:
      return `Cell added with value: ${change.newValue}`;
    case SmlChangeType.CellDeleted:
      return `Cell deleted, had value: ${change.oldValue}`;
    case SmlChangeType.ValueChanged:
      return `Value changed: ${change.oldValue} → ${change.newValue}`;
    case SmlChangeType.FormulaChanged:
      return `Formula changed: =${change.oldFormula} → =${change.newFormula}`;
    case SmlChangeType.FormatChanged:
      return 'Cell formatting changed';
    case SmlChangeType.RowInserted:
      return `Row ${change.rowIndex} inserted`;
    case SmlChangeType.RowDeleted:
      return `Row ${change.rowIndex} deleted`;
    case SmlChangeType.CommentAdded:
      return `Comment added: ${change.newComment}`;
    case SmlChangeType.CommentDeleted:
      return `Comment deleted`;
    case SmlChangeType.CommentChanged:
      return `Comment changed`;
    default:
      return SmlChangeType[change.changeType];
  }
}

/**
 * Add a sheet to the workbook
 */
async function addSheetToWorkbook(
  pkg: OoxmlPackage,
  sheetName: string,
  sheetPath: string
): Promise<void> {
  // First add relationship
  const rId = await addWorkbookRelationship(pkg, sheetPath);

  // Then add to workbook.xml
  const workbookXml = await getPartAsXml(pkg, 'xl/workbook.xml');
  if (!workbookXml) return;

  for (const node of workbookXml) {
    if (getTagName(node) === 'workbook') {
      for (const child of getChildren(node)) {
        if (getTagName(child) === 'sheets') {
          const sheets = getChildren(child);

          // Find highest sheetId
          let maxId = 0;
          for (const sheet of sheets) {
            const id = parseInt(getAttribute(sheet, 'sheetId') || '0', 10);
            if (id > maxId) maxId = id;
          }

          // Add new sheet
          sheets.push({
            sheet: [],
            ':@': {
              '@_name': sheetName,
              '@_sheetId': String(maxId + 1),
              '@_r:id': rId,
            },
          });
        }
      }
    }
  }

  setPartFromXml(pkg, 'xl/workbook.xml', workbookXml);
}

/**
 * Add a relationship to the workbook
 */
async function addWorkbookRelationship(
  pkg: OoxmlPackage,
  targetPath: string
): Promise<string> {
  const relsPath = 'xl/_rels/workbook.xml.rels';
  const relsContent = await getPartAsString(pkg, relsPath);

  if (!relsContent) {
    throw new Error('Missing workbook relationships');
  }

  const relsXml = parseXml(relsContent);
  let nextId = 1;

  for (const node of relsXml) {
    if (getTagName(node) === 'Relationships') {
      const children = getChildren(node);

      // Find next available Id
      for (const child of children) {
        const id = getAttribute(child, 'Id');
        if (id) {
          const num = parseInt(id.replace('rId', ''), 10);
          if (num >= nextId) nextId = num + 1;
        }
      }

      // Add new relationship
      const relTarget = targetPath.replace('xl/', '');
      children.push({
        Relationship: [],
        ':@': {
          '@_Id': `rId${nextId}`,
          '@_Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
          '@_Target': relTarget,
        },
      });
    }
  }

  setPartFromXml(pkg, relsPath, relsXml);
  return `rId${nextId}`;
}

/**
 * Add worksheet content type
 */
async function addWorksheetContentType(pkg: OoxmlPackage, sheetPath: string): Promise<void> {
  const contentTypesXml = await getPartAsString(pkg, '[Content_Types].xml');
  if (!contentTypesXml) return;

  const nodes = parseXml(contentTypesXml);

  for (const node of nodes) {
    if (getTagName(node) === 'Types') {
      const children = getChildren(node);
      children.push({
        Override: [],
        ':@': {
          '@_PartName': `/${sheetPath}`,
          '@_ContentType':
            'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
        },
      });
    }
  }

  setPartFromXml(pkg, '[Content_Types].xml', nodes);
}
