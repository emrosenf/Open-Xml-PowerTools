// Copyright (c) Microsoft. All rights reserved.
// Licensed under MIT license. See LICENSE file in the project root for full license information.

/**
 * SmlComparer - Excel spreadsheet comparison
 *
 * Compares two Excel spreadsheets and produces a structured result document.
 *
 * This is a TypeScript port of C# SmlComparer from Open-Xml-PowerTools.
 */

import {
  SmlChangeType,
  type SmlChange,
  type SmlComparisonResult,
  type SmlComparerSettings,
  type SmlChangeListItem,
  type SmlChangeListOptions,
  type WorksheetSignature,
  type CommentSignature,
  type DataValidationSignature,
  type HyperlinkSignature,
} from './types';

import { openPackage } from '../core/package';
import { canonicalize } from './canonicalize';
import { sheetsMatch } from './sheets';
import { compareRows } from './diff';
import { compareCells } from './cells';
import { renderMarkedWorkbook } from './markup';

/**
 * Compare two Excel spreadsheets and produce a structured result.
 *
 * @param older - The original/older workbook buffer
 * @param newer - The revised/newer workbook buffer
 * @param settings - Comparison settings
 * @returns A result object containing all detected changes
 */
export async function compare(
  older: Buffer | Uint8Array | ArrayBuffer,
  newer: Buffer | Uint8Array | ArrayBuffer,
  settings: SmlComparerSettings = {}
): Promise<SmlComparisonResult> {
  const result: SmlComparisonResult = {
    changes: [],
  };

  // 1. Open packages
  const pkg1 = await openPackage(older);
  const pkg2 = await openPackage(newer);

  // 2. Canonicalize both workbooks
  const sig1 = await canonicalize(pkg1, settings);
  const sig2 = await canonicalize(pkg2, settings);

  // 3. Match sheets between workbooks
  const sheetMatches = sheetsMatch(sig1, sig2, settings);

  // 4. Compare matched sheets
  for (const match of sheetMatches) {
    switch (match.type) {
      case 'added': {
        result.changes.push({
          changeType: SmlChangeType.SheetAdded,
          sheetName: match.newName,
          cellAddress: match.newName,
          newSheetName: match.newName,
        });
        break;
      }

      case 'deleted': {
        result.changes.push({
          changeType: SmlChangeType.SheetDeleted,
          sheetName: match.oldName,
          cellAddress: match.oldName,
          oldSheetName: match.oldName,
        });
        break;
      }

      case 'renamed': {
        result.changes.push({
          changeType: SmlChangeType.SheetRenamed,
          oldSheetName: match.oldName,
          newSheetName: match.newName,
          sheetName: match.newName,
          cellAddress: match.newName,
        });

        if (match.oldName && match.newName) {
          const sheet1 = sig1.sheets.get(match.oldName);
          const sheet2 = sig2.sheets.get(match.newName);

          if (sheet1 && sheet2) {
            compareSheets(sheet1, sheet2, settings, result);
          }
        }
        break;
      }

      case 'matched': {
        if (match.name) {
          const sheet1 = sig1.sheets.get(match.name);
          const sheet2 = sig2.sheets.get(match.name);

          if (sheet1 && sheet2) {
            compareSheets(sheet1, sheet2, settings, result);
          }
        }
        break;
      }
    }
  }

  // 5. Compare named ranges at workbook level
  compareNamedRanges(sig1.definedNames, sig2.definedNames, result);

  return result;
}

/**
 * Compare two Excel spreadsheets and produce a marked workbook with highlights.
 *
 * @param older - The original/older workbook buffer
 * @param newer - The revised/newer workbook buffer
 * @param settings - Comparison settings including highlight colors
 * @returns Buffer containing the marked workbook with differences highlighted
 */
export async function produceMarkedWorkbook(
  older: Buffer | Uint8Array | ArrayBuffer,
  newer: Buffer | Uint8Array | ArrayBuffer,
  settings: SmlComparerSettings = {}
): Promise<Buffer> {
  const pkg2 = await openPackage(newer);
  const result = await compare(older, newer, settings);
  return renderMarkedWorkbook(pkg2, result, settings);
}

function compareSheets(
  sheet1: WorksheetSignature,
  sheet2: WorksheetSignature,
  settings: SmlComparerSettings,
  result: SmlComparisonResult
): void {
  const rowChanges = compareRows(sheet1, sheet2, settings, sheet1.name);
  result.changes.push(...rowChanges);

  const cellChanges = compareCells(sheet1.cells, sheet2.cells, settings, sheet1.name);
  result.changes.push(...cellChanges);

  if (settings.compareComments !== false) {
    compareComments(sheet1.comments, sheet2.comments, sheet1.name, result);
  }

  if (settings.compareDataValidations !== false) {
    compareDataValidations(sheet1.dataValidations, sheet2.dataValidations, sheet1.name, result);
  }

  if (settings.compareMergedCells !== false) {
    compareMergedCells(sheet1.mergedCellRanges, sheet2.mergedCellRanges, sheet1.name, result);
  }

  if (settings.compareHyperlinks !== false) {
    compareHyperlinks(sheet1.hyperlinks, sheet2.hyperlinks, sheet1.name, result);
  }
}

/**
 * Compare named ranges between two workbooks.
 */
function compareNamedRanges(
  names1: Map<string, string>,
  names2: Map<string, string>,
  result: SmlComparisonResult
): void {
  // Find added named ranges
  for (const [name, value2] of names2) {
    if (!names1.has(name)) {
      result.changes.push({
        changeType: SmlChangeType.NamedRangeAdded,
        namedRangeName: name,
        newNamedRangeValue: value2,
      });
    } else {
      const value1 = names1.get(name);

      if (value1 !== value2) {
        result.changes.push({
          changeType: SmlChangeType.NamedRangeChanged,
          namedRangeName: name,
          oldNamedRangeValue: value1,
          newNamedRangeValue: value2,
        });
      }
    }
  }

  // Find deleted named ranges
  for (const [name, value1] of names1) {
    if (!names2.has(name)) {
      result.changes.push({
        changeType: SmlChangeType.NamedRangeDeleted,
        namedRangeName: name,
        oldNamedRangeValue: value1,
      });
    }
  }
}

function compareComments(
  comments1: Map<string, CommentSignature>,
  comments2: Map<string, CommentSignature>,
  _sheetName: string,
  result: SmlComparisonResult
): void {
  const allAddresses = new Set([...comments1.keys(), ...comments2.keys()]);

  for (const addr of allAddresses) {
    const c1 = comments1.get(addr);
    const c2 = comments2.get(addr);

    if (!c1 && c2) {
      result.changes.push({
        changeType: SmlChangeType.CommentAdded,
        sheetName: _sheetName,
        cellAddress: addr,
        newComment: c2.text,
        commentAuthor: c2.author,
      });
    } else if (c1 && !c2) {
      result.changes.push({
        changeType: SmlChangeType.CommentDeleted,
        sheetName: _sheetName,
        cellAddress: addr,
        oldComment: c1.text,
        commentAuthor: c1.author,
      });
    } else if (c1 && c2 && (c1.text !== c2.text || c1.author !== c2.author)) {
      result.changes.push({
        changeType: SmlChangeType.CommentChanged,
        sheetName: _sheetName,
        cellAddress: addr,
        oldComment: c1.text,
        newComment: c2.text,
        commentAuthor: c2.author,
      });
    }
  }
}

function compareDataValidations(
  dvs1: Map<string, DataValidationSignature>,
  dvs2: Map<string, DataValidationSignature>,
  _sheetName: string,
  result: SmlComparisonResult
): void {
  const allKeys = new Set([...dvs1.keys(), ...dvs2.keys()]);

  for (const key of allKeys) {
    const dv1 = dvs1.get(key);
    const dv2 = dvs2.get(key);

    if (!dv1 && dv2) {
      result.changes.push({
        changeType: SmlChangeType.DataValidationAdded,
        sheetName: _sheetName,
        cellAddress: key,
        dataValidationType: dv2.type,
        newDataValidation: formatDataValidation(dv2),
      });
    } else if (dv1 && !dv2) {
      result.changes.push({
        changeType: SmlChangeType.DataValidationDeleted,
        sheetName: _sheetName,
        cellAddress: key,
        dataValidationType: dv1.type,
        oldDataValidation: formatDataValidation(dv1),
      });
    } else if (dv1 && dv2 && dv1.hash !== dv2.hash) {
      result.changes.push({
        changeType: SmlChangeType.DataValidationChanged,
        sheetName: _sheetName,
        cellAddress: key,
        dataValidationType: dv2.type,
        oldDataValidation: formatDataValidation(dv1),
        newDataValidation: formatDataValidation(dv2),
      });
    }
  }
}

function formatDataValidation(dv: DataValidationSignature): string {
  let s = `${dv.type}`;
  if (dv.operator) s += ` ${dv.operator}`;
  if (dv.formula1) s += ` ${dv.formula1}`;
  if (dv.formula2) s += ` ${dv.formula2}`;
  return s;
}

function compareMergedCells(
  merged1: Set<string>,
  merged2: Set<string>,
  _sheetName: string,
  result: SmlComparisonResult
): void {
  for (const range of merged2) {
    if (!merged1.has(range)) {
      result.changes.push({
        changeType: SmlChangeType.MergedCellAdded,
        sheetName: _sheetName,
        mergedCellRange: range,
        cellRange: range,
      });
    }
  }

  for (const range of merged1) {
    if (!merged2.has(range)) {
      result.changes.push({
        changeType: SmlChangeType.MergedCellDeleted,
        sheetName: _sheetName,
        mergedCellRange: range,
        cellRange: range,
      });
    }
  }
}

function compareHyperlinks(
  hls1: Map<string, HyperlinkSignature>,
  hls2: Map<string, HyperlinkSignature>,
  _sheetName: string,
  result: SmlComparisonResult
): void {
  const allAddresses = new Set([...hls1.keys(), ...hls2.keys()]);

  for (const addr of allAddresses) {
    const hl1 = hls1.get(addr);
    const hl2 = hls2.get(addr);

    if (!hl1 && hl2) {
      result.changes.push({
        changeType: SmlChangeType.HyperlinkAdded,
        sheetName: _sheetName,
        cellAddress: addr,
        newHyperlink: hl2.target,
      });
    } else if (hl1 && !hl2) {
      result.changes.push({
        changeType: SmlChangeType.HyperlinkDeleted,
        sheetName: _sheetName,
        cellAddress: addr,
        oldHyperlink: hl1.target,
      });
    } else if (hl1 && hl2 && hl1.hash !== hl2.hash) {
      result.changes.push({
        changeType: SmlChangeType.HyperlinkChanged,
        sheetName: _sheetName,
        cellAddress: addr,
        oldHyperlink: hl1.target,
        newHyperlink: hl2.target,
      });
    }
  }
}

export function buildChangeList(
  result: SmlComparisonResult,
  options: SmlChangeListOptions = {}
): SmlChangeListItem[] {
  const groupAdjacentCells = options.groupAdjacentCells !== false;
  const baseItems = result.changes.map((change, index) =>
    toChangeListItem(change, index)
  );

  if (!groupAdjacentCells) {
    return baseItems;
  }

  return groupAdjacentChangeItems(baseItems);
}

function toChangeListItem(change: SmlChange, index: number): SmlChangeListItem {
  const sheetName = change.sheetName ?? undefined;
  const cellAddress = normalizeCellAddress(change.cellAddress);
  const cellRange = normalizeCellRange(change.cellRange ?? change.mergedCellRange);
  const summary = summarizeChange(change);

  return {
    id: `change-${index + 1}`,
    changeType: change.changeType,
    sheetName,
    cellAddress,
    cellRange,
    rowIndex: change.rowIndex,
    columnIndex: change.columnIndex,
    summary,
    details: {
      oldValue: change.oldValue,
      newValue: change.newValue,
      oldFormula: change.oldFormula,
      newFormula: change.newFormula,
      oldFormat: change.oldFormat,
      newFormat: change.newFormat,
      oldComment: change.oldComment,
      newComment: change.newComment,
      commentAuthor: change.commentAuthor,
      dataValidationType: change.dataValidationType,
      oldDataValidation: change.oldDataValidation,
      newDataValidation: change.newDataValidation,
      mergedCellRange: change.mergedCellRange,
      oldHyperlink: change.oldHyperlink,
      newHyperlink: change.newHyperlink,
      oldSheetName: change.oldSheetName,
      newSheetName: change.newSheetName,
    },
    anchor: buildAnchor(sheetName, cellRange ?? cellAddress),
  };
}

function normalizeCellAddress(address?: string): string | undefined {
  if (!address) return undefined;
  if (address.includes('!')) {
    return address.split('!')[1];
  }
  return address;
}

function normalizeCellRange(range?: string): string | undefined {
  if (!range) return undefined;
  if (range.includes('!')) {
    return range.split('!')[1];
  }
  return range;
}

function buildAnchor(sheetName?: string, cellRef?: string): string | undefined {
  if (!sheetName) return undefined;
  if (!cellRef) return sheetName;
  return `${sheetName}!${cellRef}`;
}

function summarizeChange(change: SmlChange): string {
  switch (change.changeType) {
    case SmlChangeType.SheetAdded:
      return `Worksheet inserted: ${change.newSheetName ?? change.sheetName ?? ''}`.trim();
    case SmlChangeType.SheetDeleted:
      return `Worksheet deleted: ${change.oldSheetName ?? change.sheetName ?? ''}`.trim();
    case SmlChangeType.SheetRenamed:
      return `Worksheet renamed: ${change.oldSheetName ?? ''} → ${change.newSheetName ?? ''}`.trim();
    case SmlChangeType.CellAdded:
      return 'Cell added';
    case SmlChangeType.CellDeleted:
      return 'Cell deleted';
    case SmlChangeType.ValueChanged:
      return 'Value modified';
    case SmlChangeType.FormulaChanged:
      return 'Formula modified';
    case SmlChangeType.FormatChanged:
      return 'Format modified';
    case SmlChangeType.RowInserted:
      return 'Row inserted';
    case SmlChangeType.RowDeleted:
      return 'Row deleted';
    case SmlChangeType.CommentAdded:
      return 'Comment added';
    case SmlChangeType.CommentDeleted:
      return 'Comment deleted';
    case SmlChangeType.CommentChanged:
      return 'Comment changed';
    case SmlChangeType.DataValidationAdded:
      return 'Data validation added';
    case SmlChangeType.DataValidationDeleted:
      return 'Data validation deleted';
    case SmlChangeType.DataValidationChanged:
      return 'Data validation changed';
    case SmlChangeType.MergedCellAdded:
      return 'Merged range added';
    case SmlChangeType.MergedCellDeleted:
      return 'Merged range deleted';
    case SmlChangeType.HyperlinkAdded:
      return 'Hyperlink added';
    case SmlChangeType.HyperlinkDeleted:
      return 'Hyperlink deleted';
    case SmlChangeType.HyperlinkChanged:
      return 'Hyperlink changed';
    default:
      return SmlChangeType[change.changeType];
  }
}

function groupAdjacentChangeItems(items: SmlChangeListItem[]): SmlChangeListItem[] {
  const grouped: SmlChangeListItem[] = [];
  const sorted = [...items].sort(compareByLocation);
  const lastByKey = new Map<string, SmlChangeListItem>();

  for (const item of sorted) {
    if (!isGroupableCellChange(item)) {
      grouped.push({ ...item });
      continue;
    }

    const key = `${item.sheetName}|${item.changeType}|${item.columnIndex}`;
    const last = lastByKey.get(key);

    if (!last || !canGroup(last, item)) {
      const fresh = { ...item };
      grouped.push(fresh);
      lastByKey.set(key, fresh);
      continue;
    }

    const range = mergeRange(last, item);
    last.cellRange = range;
    last.count = (last.count ?? 1) + 1;
    last.anchor = buildAnchor(last.sheetName, range);
  }

  return grouped;
}

function isGroupableCellChange(item: SmlChangeListItem): boolean {
  if (!item.sheetName) return false;
  if (item.rowIndex === undefined || item.columnIndex === undefined) return false;
  return (
    item.changeType === SmlChangeType.CellAdded ||
    item.changeType === SmlChangeType.CellDeleted ||
    item.changeType === SmlChangeType.ValueChanged ||
    item.changeType === SmlChangeType.FormulaChanged ||
    item.changeType === SmlChangeType.FormatChanged
  );
}

function compareByLocation(a: SmlChangeListItem, b: SmlChangeListItem): number {
  if ((a.sheetName ?? '') !== (b.sheetName ?? '')) {
    return (a.sheetName ?? '').localeCompare(b.sheetName ?? '');
  }

  if ((a.rowIndex ?? 0) !== (b.rowIndex ?? 0)) {
    return (a.rowIndex ?? 0) - (b.rowIndex ?? 0);
  }

  return (a.columnIndex ?? 0) - (b.columnIndex ?? 0);
}

function canGroup(base: SmlChangeListItem, next: SmlChangeListItem): boolean {
  if (!base.sheetName || !next.sheetName) return false;
  if (base.sheetName !== next.sheetName) return false;
  if (base.changeType !== next.changeType) return false;
  if (base.rowIndex === undefined || next.rowIndex === undefined) return false;
  if (base.columnIndex === undefined || next.columnIndex === undefined) return false;
  if (base.columnIndex !== next.columnIndex) return false;
  if (next.rowIndex !== base.rowIndex + (base.count ?? 1)) return false;
  return true;
}

function mergeRange(base: SmlChangeListItem, next: SmlChangeListItem): string {
  const baseRow = base.rowIndex ?? 0;
  const nextRow = next.rowIndex ?? baseRow;
  const col = base.columnIndex ?? 0;
  const colLetter = columnIndexToLetters(col + 1);
  const start = `${colLetter}${baseRow}`;
  const end = `${colLetter}${nextRow}`;
  return start === end ? start : `${start}:${end}`;
}

function columnIndexToLetters(index: number): string {
  let result = '';
  let n = index;
  while (n > 0) {
    n--;
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26);
  }
  return result;
}
