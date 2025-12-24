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
  getRelationships,
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

interface FontInfo {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  size?: number;
  name?: string;
  color?: string;
}

interface FillInfo {
  pattern?: string;
  foregroundColor?: string;
  backgroundColor?: string;
}

interface BorderInfo {
  leftStyle?: string;
  leftColor?: string;
  rightStyle?: string;
  rightColor?: string;
  topStyle?: string;
  topColor?: string;
  bottomStyle?: string;
  bottomColor?: string;
}

interface CellXf {
  numFmtId: number;
  fontId: number;
  fillId: number;
  borderId: number;
  horizontalAlignment?: string;
  verticalAlignment?: string;
  wrapText?: boolean;
  indent?: number;
}

interface StyleInfo {
  numberFormats: Map<number, string>;
  fonts: FontInfo[];
  fills: FillInfo[];
  borders: BorderInfo[];
  cellFormats: CellXf[];
}

async function parseStyles(
  pkg: OoxmlPackage
): Promise<Map<number, CellFormatSignature>> {
  const stylesXml = await getPartAsXml(pkg, 'xl/styles.xml');
  if (!stylesXml) return new Map();

  const styleInfo: StyleInfo = {
    numberFormats: new Map(),
    fonts: [],
    fills: [],
    borders: [],
    cellFormats: [],
  };

  for (const node of stylesXml) {
    if (getTagName(node) === 'styleSheet') {
      const children = getChildren(node);

      for (const child of children) {
        const tagName = getTagName(child);

        if (tagName === 'numFmts') {
          parseNumFmts(child, styleInfo);
        } else if (tagName === 'fonts') {
          parseFonts(child, styleInfo);
        } else if (tagName === 'fills') {
          parseFills(child, styleInfo);
        } else if (tagName === 'borders') {
          parseBorders(child, styleInfo);
        } else if (tagName === 'cellXfs') {
          parseCellXfs(child, styleInfo);
        }
      }
    }
  }

  return expandAllStyles(styleInfo);
}

function parseNumFmts(numFmtsNode: XmlNode, styleInfo: StyleInfo): void {
  for (const child of getChildren(numFmtsNode)) {
    if (getTagName(child) !== 'numFmt') continue;
    const attrs = child[':@'] as Record<string, string> | undefined;
    if (attrs) {
      const id = parseInt(attrs['numFmtId'] || '0', 10);
      const code = attrs['formatCode'] || '';
      styleInfo.numberFormats.set(id, code);
    }
  }
}

function parseFonts(fontsNode: XmlNode, styleInfo: StyleInfo): void {
  for (const fontNode of getChildren(fontsNode)) {
    if (getTagName(fontNode) !== 'font') continue;
    const font: FontInfo = {};

    for (const child of getChildren(fontNode)) {
      const tagName = getTagName(child);
      const attrs = child[':@'] as Record<string, string> | undefined;

      if (tagName === 'b') font.bold = true;
      else if (tagName === 'i') font.italic = true;
      else if (tagName === 'u') font.underline = true;
      else if (tagName === 'strike') font.strikethrough = true;
      else if (tagName === 'sz' && attrs) font.size = parseFloat(attrs['val'] || '0');
      else if (tagName === 'name' && attrs) font.name = attrs['val'];
      else if (tagName === 'color' && attrs) font.color = getColorValue(attrs);
    }

    styleInfo.fonts.push(font);
  }
}

function parseFills(fillsNode: XmlNode, styleInfo: StyleInfo): void {
  for (const fillNode of getChildren(fillsNode)) {
    if (getTagName(fillNode) !== 'fill') continue;
    const fill: FillInfo = {};

    for (const child of getChildren(fillNode)) {
      if (getTagName(child) === 'patternFill') {
        const attrs = child[':@'] as Record<string, string> | undefined;
        if (attrs) fill.pattern = attrs['patternType'];

        for (const pfChild of getChildren(child)) {
          const pfAttrs = pfChild[':@'] as Record<string, string> | undefined;
          if (getTagName(pfChild) === 'fgColor' && pfAttrs) {
            fill.foregroundColor = getColorValue(pfAttrs);
          } else if (getTagName(pfChild) === 'bgColor' && pfAttrs) {
            fill.backgroundColor = getColorValue(pfAttrs);
          }
        }
      }
    }

    styleInfo.fills.push(fill);
  }
}

function parseBorders(bordersNode: XmlNode, styleInfo: StyleInfo): void {
  for (const borderNode of getChildren(bordersNode)) {
    if (getTagName(borderNode) !== 'border') continue;
    const border: BorderInfo = {};

    for (const child of getChildren(borderNode)) {
      const tagName = getTagName(child);
      const attrs = child[':@'] as Record<string, string> | undefined;
      const style = attrs?.['style'];

      let color: string | undefined;
      for (const colorChild of getChildren(child)) {
        if (getTagName(colorChild) === 'color') {
          const colorAttrs = colorChild[':@'] as Record<string, string> | undefined;
          if (colorAttrs) color = getColorValue(colorAttrs);
        }
      }

      if (tagName === 'left') {
        border.leftStyle = style;
        border.leftColor = color;
      } else if (tagName === 'right') {
        border.rightStyle = style;
        border.rightColor = color;
      } else if (tagName === 'top') {
        border.topStyle = style;
        border.topColor = color;
      } else if (tagName === 'bottom') {
        border.bottomStyle = style;
        border.bottomColor = color;
      }
    }

    styleInfo.borders.push(border);
  }
}

function parseCellXfs(cellXfsNode: XmlNode, styleInfo: StyleInfo): void {
  for (const xfNode of getChildren(cellXfsNode)) {
    if (getTagName(xfNode) !== 'xf') continue;
    const attrs = xfNode[':@'] as Record<string, string> | undefined;

    const xf: CellXf = {
      numFmtId: parseInt(attrs?.['numFmtId'] || '0', 10),
      fontId: parseInt(attrs?.['fontId'] || '0', 10),
      fillId: parseInt(attrs?.['fillId'] || '0', 10),
      borderId: parseInt(attrs?.['borderId'] || '0', 10),
    };

    for (const child of getChildren(xfNode)) {
      if (getTagName(child) === 'alignment') {
        const alignAttrs = child[':@'] as Record<string, string> | undefined;
        if (alignAttrs) {
          xf.horizontalAlignment = alignAttrs['horizontal'];
          xf.verticalAlignment = alignAttrs['vertical'];
          xf.wrapText = alignAttrs['wrapText'] === '1';
          xf.indent = parseInt(alignAttrs['indent'] || '0', 10);
        }
      }
    }

    styleInfo.cellFormats.push(xf);
  }
}

function getColorValue(attrs: Record<string, string>): string {
  if (attrs['rgb']) return attrs['rgb'];
  if (attrs['indexed']) return `indexed:${attrs['indexed']}`;
  if (attrs['theme']) return `theme:${attrs['theme']}`;
  return '';
}

function getBuiltInNumberFormat(numFmtId: number): string {
  const formats: Record<number, string> = {
    0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
    9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
    14: 'mm-dd-yy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy',
    18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss',
    22: 'm/d/yy h:mm', 37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)',
    39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)',
    45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 48: '##0.0E+0', 49: '@',
  };
  return formats[numFmtId] || 'General';
}

function expandAllStyles(styleInfo: StyleInfo): Map<number, CellFormatSignature> {
  const formats = new Map<number, CellFormatSignature>();

  for (let i = 0; i < styleInfo.cellFormats.length; i++) {
    const xf = styleInfo.cellFormats[i];
    const format: CellFormatSignature = {};

    const customFmt = styleInfo.numberFormats.get(xf.numFmtId);
    format.numberFormatCode = customFmt ?? getBuiltInNumberFormat(xf.numFmtId);

    if (xf.fontId >= 0 && xf.fontId < styleInfo.fonts.length) {
      const font = styleInfo.fonts[xf.fontId];
      format.bold = font.bold;
      format.italic = font.italic;
      format.underline = font.underline;
      format.strikethrough = font.strikethrough;
      format.fontName = font.name;
      format.fontSize = font.size;
      format.fontColor = font.color;
    }

    if (xf.fillId >= 0 && xf.fillId < styleInfo.fills.length) {
      const fill = styleInfo.fills[xf.fillId];
      format.fillPattern = fill.pattern;
      format.fillForegroundColor = fill.foregroundColor;
      format.fillBackgroundColor = fill.backgroundColor;
    }

    if (xf.borderId >= 0 && xf.borderId < styleInfo.borders.length) {
      const border = styleInfo.borders[xf.borderId];
      format.borderLeftStyle = border.leftStyle;
      format.borderLeftColor = border.leftColor;
      format.borderRightStyle = border.rightStyle;
      format.borderTopStyle = border.topStyle;
      format.borderTopColor = border.topColor;
      format.borderBottomStyle = border.bottomStyle;
      format.borderBottomColor = border.bottomColor;
    }

    format.horizontalAlignment = xf.horizontalAlignment;
    format.verticalAlignment = xf.verticalAlignment;
    format.wrapText = xf.wrapText;
    format.indent = xf.indent;

    formats.set(i, format);
  }

  return formats;
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
  const worksheetPath = await resolveWorksheetPath(pkg, relId);

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
 *
 * Parses xl/_rels/workbook.xml.rels to find the target path for a given relationship ID.
 */
async function resolveWorksheetPath(
  pkg: OoxmlPackage,
  relId: string
): Promise<string | null> {
  // Get relationships from xl/_rels/workbook.xml.rels
  const relationships = await getRelationships(pkg, 'xl/workbook.xml');

  // Find the relationship with matching ID
  const rel = relationships.find((r) => r.id === relId);

  if (!rel) {
    return null;
  }

  // Target is relative to xl/ directory
  // e.g., "worksheets/sheet1.xml" -> "xl/worksheets/sheet1.xml"
  const target = rel.target;

  if (target.startsWith('/')) {
    // Absolute path from root
    return target.slice(1);
  } else if (target.startsWith('../')) {
    // Relative path going up - resolve from xl/
    return target.replace('../', '');
  } else {
    // Relative path from xl/
    return `xl/${target}`;
  }
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
