// Copyright (c) Microsoft. All rights reserved.
// Licensed under MIT license. See LICENSE file in the project root for full license information.

/**
 * Excel workbook canonicalization
 *
 * This module handles the canonicalization of Excel workbooks for comparison.
 * It extracts shared strings, expands style indices, and creates internal
 * signature structures that are used for efficient comparison.
 */

import {
  openPackage,
  getPartAsXml,
  type OoxmlPackage,
} from '../core/package';
import {
  getTagName,
  getChildren,
  getTextContent,
  findNodes,
  type XmlNode,
} from '../core/xml';
import { hashString } from '../core/hash';
import type {
  WorkbookSignature,
  WorksheetSignature,
  CellSignature,
  CellFormatSignature,
  CommentSignature,
  DataValidationSignature,
  HyperlinkSignature,
  MergedCellRange,
  ConditionalFormatRange,
} from './types';

/**
 * Canonicalize an Excel workbook for comparison.
 *
 * This function performs the canonicalization phases:
 * 1. Load workbook and parse relationships
 * 2. Extract and resolve shared strings
 * 3. Extract and expand cell formatting
 * 4. Parse worksheets and extract cells
 * 5. Extract Phase 3 features (comments, data validation, etc.)
 *
 * @param pkg The OOXML package
 * @param settings Comparison settings
 * @returns Canonicalized workbook signature
 */
export async function canonicalize(
  pkg: OoxmlPackage,
  settings: any
): Promise<WorkbookSignature> {
  // Parse workbook.xml to get worksheet relationships
  const workbookXml = await getPartAsXml(pkg, 'xl/workbook.xml');
  if (!workbookXml) {
    throw new Error('Invalid Excel workbook: missing xl/workbook.xml');
  }

  // Extract worksheet relationships
  const worksheetRels = extractWorksheetRels(workbookXml);

  // Parse shared strings table
  const sharedStrings = await extractSharedStrings(pkg);

  // Parse styles
  const styles = await parseStyles(pkg);

  // Create workbook signature
  const workbookSig: WorkbookSignature = {
    sheets: new Map<string, WorksheetSignature>(),
    definedNames: await extractDefinedNames(pkg),
  };

  // Process each worksheet
  for (const rel of worksheetRels) {
    const sheetSig = await canonicalizeWorksheet(
      pkg,
      rel.id,
      rel.name,
      sharedStrings,
      styles,
      settings
    );
    workbookSig.sheets.set(rel.name, sheetSig);
  }

  return workbookSig;
}

/**
 * Extract worksheet relationships from workbook.xml
 */
function extractWorksheetRels(workbookXml: XmlNode[]): Array<{ id: string; name: string }> {
  const rels: Array<{ id: string; name: string }> = [];

  for (const node of workbookXml) {
    if (getTagName(node) === 'workbook') {
      const children = getChildren(node);

      // Find sheets element
      for (const child of children) {
        if (getTagName(child) === 'sheets') {
          const sheets = getChildren(child);

          for (const sheet of sheets) {
            const attrs = sheet[':@'] as Record<string, string> | undefined;
            if (attrs) {
              rels.push({
                id: attrs['r:id'] || '',
                name: attrs['name'] || '',
              });
            }
          }
        }
      }
    }
  }

  return rels;
}

/**
 * Extract shared strings from xl/sharedStrings.xml
 */
async function extractSharedStrings(
  pkg: OoxmlPackage
): Promise<string[] | null> {
  const sharedStringsXml = await getPartAsXml(pkg, 'xl/sharedStrings.xml');
  if (!sharedStringsXml) return null;

  const strings: string[] = [];

  for (const node of sharedStringsXml) {
    if (getTagName(node) === 'sst') {
      const children = getChildren(node);

      for (const child of children) {
        if (getTagName(child) === 'si') {
          // Extract text from si element
          const text = extractSiText(child);
          strings.push(text);
        }
      }
    }
  }

  return strings;
}

/**
 * Extract text content from an si (shared string item) element
 */
function extractSiText(siNode: XmlNode): string {
  let text = '';

  // si can contain <t> (text), <r> (rich text run), etc.
  const children = getChildren(siNode);

  for (const child of children) {
    const tagName = getTagName(child);

    if (tagName === 't') {
      text += getTextContent(child);
    } else if (tagName === 'r') {
      // Rich text run - extract text from <t> inside
      const tNode = findNodes(child, (n) => getTagName(n) === 't');
      for (const t of tNode) {
        text += getTextContent(t);
      }
    }
  }

  return text;
}

/**
 * Parse styles from xl/styles.xml
 *
 * This extracts the cellXfs (cell formats) and builds a map to resolve
 * style indices to actual formatting properties.
 */
async function parseStyles(
  pkg: OoxmlPackage
): Promise<Map<number, CellFormatSignature>> {
  const stylesXml = await getPartAsXml(pkg, 'xl/styles.xml');
  if (!stylesXml) return new Map();

  const formats = new Map<number, CellFormatSignature>();

  for (const node of stylesXml) {
    if (getTagName(node) === 'styleSheet') {
      const children = getChildren(node);

      // Find cellXfs (cell formats)
      for (const child of children) {
        if (getTagName(child) === 'cellXfs') {
          const xfs = getChildren(child);

          for (let i = 0; i < xfs.length; i++) {
            const xf = xfs[i];
            const format = parseXf(xf);
            formats.set(i, format);
          }
        }
      }
    }
  }

  return formats;
}

/**
 * Parse an xf (cell format) element
 */
function parseXf(xfNode: XmlNode): CellFormatSignature {
  const attrs = xfNode[':@'] as Record<string, string> | undefined;
  const format: CellFormatSignature = {};

  if (attrs) {
    format.numberFormatCode = attrs['numFmtId'];

    if (attrs['fontId']) {
      format.bold = attrs['fontId'] === '1';
    }

    if (attrs['fillId']) {
      // Could parse fill from fills element
      format.fillForegroundColor = attrs['fillId'];
    }
  }

  return format;
}

/**
 * Canonicalize a single worksheet
 */
async function canonicalizeWorksheet(
  pkg: OoxmlPackage,
  relId: string,
  name: string,
  sharedStrings: string[] | null,
  styles: Map<number, CellFormatSignature>,
  settings: any
): Promise<WorksheetSignature> {
  // Get worksheet XML path from relationship
  const worksheetPath = resolveWorksheetPath(pkg, relId);

  if (!worksheetPath) {
    throw new Error(`Cannot find worksheet for relId: ${relId}`);
  }

  const worksheetXml = await getPartAsXml(pkg, worksheetPath);
  if (!worksheetXml) {
    throw new Error(`Cannot parse worksheet: ${worksheetPath}`);
  }

  // Create worksheet signature
  const sheetSig: WorksheetSignature = {
    name,
    relationshipId: relId,
    cells: new Map<string, CellSignature>(),
    populatedRows: new Set<number>(),
    populatedColumns: new Set<number>(),
    rowSignatures: new Map<number, string>(),
    columnSignatures: new Map<number, string>(),
    comments: new Map<string, CommentSignature>(),
    dataValidations: new Map<string, DataValidationSignature>(),
    mergedCellRanges: new Set<string>(),
    hyperlinks: new Map<string, HyperlinkSignature>(),
  };

  // Parse sheet data
  for (const node of worksheetXml) {
    if (getTagName(node) === 'worksheet') {
      const children = getChildren(node);

      for (const child of children) {
        const tagName = getTagName(child);

        if (tagName === 'sheetData') {
          parseSheetData(child, sheetSig, sharedStrings, styles, settings);
        } else if (tagName === 'mergeCells') {
          parseMergedCells(child, sheetSig);
        } else if (tagName === 'hyperlinks') {
          parseHyperlinks(child, sheetSig);
        } else if (tagName === 'comments') {
          // Comments are in separate part, not inline
        } else if (tagName === 'dataValidations') {
          parseDataValidations(child, sheetSig);
        } else if (tagName === 'conditionalFormatting') {
          parseConditionalFormatting(child, sheetSig);
        }
      }
    }
  }

  // Compute row and column signatures
  computeRowSignatures(sheetSig);
  computeColumnSignatures(sheetSig);

  return sheetSig;
}

/**
 * Resolve worksheet path from relationship ID
 */
function resolveWorksheetPath(
  pkg: OoxmlPackage,
  relId: string
): string | null {
  // Relationship ID -> target path mapping is in xl/_rels/workbook.xml.rels
  // For now, assume standard naming: xl/worksheets/sheet1.xml, etc.
  const sheets = pkg.zip.file(/xl\/worksheets\/sheet\d+\.xml/);

  if (sheets && sheets.length > 0) {
    return sheets[0].name;
  }

  return null;
}

/**
 * Parse sheet data and extract cells
 */
function parseSheetData(
  sheetDataNode: XmlNode,
  sheetSig: WorksheetSignature,
  sharedStrings: string[] | null,
  styles: Map<number, CellFormatSignature>,
  settings: any
): void {
  const rows = getChildren(sheetDataNode);

  for (const rowNode of rows) {
    if (getTagName(rowNode) !== 'row') continue;

    const attrs = rowNode[':@'] as Record<string, string> | undefined;
    const rowIndex = attrs ? parseInt(attrs['r'] || '0', 10) : 0;

    // Track populated rows
    sheetSig.populatedRows.add(rowIndex);

    const cells = getChildren(rowNode);

    for (const cellNode of cells) {
      if (getTagName(cellNode) !== 'c') continue;

      const cellSig = parseCell(cellNode, rowIndex, sharedStrings, styles, settings);
      if (cellSig) {
        sheetSig.cells.set(cellSig.address, cellSig);
        sheetSig.populatedColumns.add(cellSig.column);
      }
    }
  }
}

/**
 * Parse a single cell element
 */
function parseCell(
  cellNode: XmlNode,
  rowIndex: number,
  sharedStrings: string[] | null,
  styles: Map<number, CellFormatSignature>,
  settings: any
): CellSignature | null {
  const attrs = cellNode[':@'] as Record<string, string> | undefined;
  if (!attrs) return null;

  // Parse cell address (e.g., "A1")
  const cellAddress = attrs['r'] || '';
  const columnIndex = parseColumnIndex(cellAddress);

  // Extract cell content
  let resolvedValue = '';
  let formula = '';
  let styleIndex = -1;

  const children = getChildren(cellNode);

  for (const child of children) {
    const tagName = getTagName(child);

    if (tagName === 'v') {
      // Value - resolve from shared strings if needed
      const rawValue = getTextContent(child);
      const type = attrs['t']; // s = shared string, str = string, etc.

      if (type === 's' && sharedStrings) {
        const index = parseInt(rawValue, 10);
        if (index >= 0 && index < sharedStrings.length) {
          resolvedValue = sharedStrings[index];
        }
      } else {
        resolvedValue = rawValue;
      }
    } else if (tagName === 'f') {
      // Formula
      formula = getTextContent(child);
    }
  }

  // Get style index
  if (attrs['s']) {
    styleIndex = parseInt(attrs['s'], 10);
  }

  // Get cell format
  const format = styleIndex >= 0 ? styles.get(styleIndex) : {};

  // Compute content hash
  const content = formula || resolvedValue;
  const contentHash = hashString(content);

  return {
    address: cellAddress,
    row: rowIndex,
    column: columnIndex,
    resolvedValue,
    formula,
    contentHash,
    format,
  };
}

/**
 * Parse column index from cell address (e.g., "A1" -> 0, "B1" -> 1)
 */
function parseColumnIndex(cellAddress: string): number {
  const match = cellAddress.match(/^([A-Z]+)/);
  if (!match) return 0;

  const col = match[1];
  let index = 0;

  for (let i = 0; i < col.length; i++) {
    index = index * 26 + (col.charCodeAt(i) - 64);
  }

  return index - 1;
}

/**
 * Parse merged cells
 */
function parseMergedCells(
  mergeCellsNode: XmlNode,
  sheetSig: WorksheetSignature
): void {
  const merges = getChildren(mergeCellsNode);

  for (const mergeNode of merges) {
    if (getTagName(mergeNode) !== 'mergeCell') continue;

    const attrs = mergeNode[':@'] as Record<string, string> | undefined;
    if (attrs && attrs['ref']) {
      sheetSig.mergedCellRanges.add(attrs['ref']);
    }
  }
}

/**
 * Parse hyperlinks
 */
function parseHyperlinks(
  hyperlinksNode: XmlNode,
  sheetSig: WorksheetSignature
): void {
  const links = getChildren(hyperlinksNode);

  for (const linkNode of links) {
    if (getTagName(linkNode) !== 'hyperlink') continue;

    const attrs = linkNode[':@'] as Record<string, string> | undefined;
    if (attrs) {
      const sig: HyperlinkSignature = {
        cellAddress: attrs['ref'] || '',
        target: attrs['target'] || '',
        hash: hashString(attrs['target'] || ''),
      };
      sheetSig.hyperlinks.set(sig.cellAddress, sig);
    }
  }
}

/**
 * Parse data validations
 */
function parseDataValidations(
  dataValidationsNode: XmlNode,
  sheetSig: WorksheetSignature
): void {
  const validations = getChildren(dataValidationsNode);

  for (const validationNode of validations) {
    if (getTagName(validationNode) !== 'dataValidation') continue;

    const attrs = validationNode[':@'] as Record<string, string> | undefined;
    if (!attrs) continue;

    const sig: DataValidationSignature = {
      cellRange: attrs['sqref'] || '',
      type: attrs['type'] || '',
      operator: attrs['operator'],
      formula1: attrs['formula1'],
      formula2: attrs['formula2'],
      allowBlank: attrs['allowBlank'] === '1',
      showDropDown: attrs['showDropDown'] === '1',
      showInputMessage: attrs['showInputMessage'] === '1',
      showErrorMessage: attrs['showErrorMessage'] === '1',
      errorTitle: attrs['errorTitle'],
      error: attrs['error'],
      promptTitle: attrs['promptTitle'],
      prompt: attrs['prompt'],
      hash: '',
    };

    // Compute hash from all properties
    const hashParts = [
      sig.cellRange,
      sig.type,
      sig.operator,
      sig.formula1,
      sig.formula2,
    ];
    sig.hash = hashString(hashParts.join('|'));

    sheetSig.dataValidations.set(sig.cellRange, sig);
  }
}

/**
 * Parse conditional formatting
 */
function parseConditionalFormatting(
  cfNode: XmlNode,
  sheetSig: WorksheetSignature
): void {
  // For now, just track that conditional formatting exists
  // Full implementation would parse cfRule elements
}

/**
 * Compute row signatures for LCS-based alignment
 */
function computeRowSignatures(sheetSig: WorksheetSignature): void {
  const maxRow = Math.max(...sheetSig.populatedRows.values());

  for (let row = 1; row <= maxRow; row++) {
    if (!sheetSig.populatedRows.has(row)) continue;

    // Build signature from all cells in this row
    const cellSignatures: string[] = [];

    for (const [address, cell] of sheetSig.cells) {
      if (cell.row === row) {
        cellSignatures.push(cell.contentHash);
      }
    }

    // Hash the row signature
    const rowSig = hashString(cellSignatures.join('|'));
    sheetSig.rowSignatures.set(row, rowSig);
  }
}

/**
 * Compute column signatures for LCS-based alignment
 */
function computeColumnSignatures(sheetSig: WorksheetSignature): void {
  const maxCol = Math.max(...sheetSig.populatedColumns.values());

  for (let col = 1; col <= maxCol; col++) {
    if (!sheetSig.populatedColumns.has(col)) continue;

    // Build signature from all cells in this column
    const cellSignatures: string[] = [];

    for (const [address, cell] of sheetSig.cells) {
      if (cell.column === col) {
        cellSignatures.push(cell.contentHash);
      }
    }

    // Hash the column signature
    const colSig = hashString(cellSignatures.join('|'));
    sheetSig.columnSignatures.set(col, colSig);
  }
}

/**
 * Extract defined names (named ranges) from workbook
 */
async function extractDefinedNames(
  pkg: OoxmlPackage
): Promise<Map<string, string>> {
  const definedNames = new Map<string, string>();

  // Try to parse from workbook.xml
  const workbookXml = await getPartAsXml(pkg, 'xl/workbook.xml');
  if (!workbookXml) return definedNames;

  for (const node of workbookXml) {
    if (getTagName(node) === 'workbook') {
      const children = getChildren(node);

      for (const child of children) {
        if (getTagName(child) === 'definedNames') {
          const names = getChildren(child);

          for (const nameNode of names) {
            if (getTagName(nameNode) === 'definedName') {
              const attrs = nameNode[':@'] as Record<string, string> | undefined;
              if (attrs && attrs['name']) {
                definedNames.set(attrs['name'], getTextContent(nameNode));
              }
            }
          }
        }
      }
    }
  }

  return definedNames;
}
