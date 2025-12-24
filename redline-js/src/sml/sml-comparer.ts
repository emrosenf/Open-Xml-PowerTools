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
  type SmlComparisonResult,
  type SmlComparerSettings,
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
          cellAddress: match.newName,
        });
        break;
      }

      case 'deleted': {
        result.changes.push({
          changeType: SmlChangeType.SheetDeleted,
          cellAddress: match.oldName,
        });
        break;
      }

      case 'renamed': {
        result.changes.push({
          changeType: SmlChangeType.SheetRenamed,
          oldSheetName: match.oldName,
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

function compareSheets(
  sheet1: WorksheetSignature,
  sheet2: WorksheetSignature,
  settings: SmlComparerSettings,
  result: SmlComparisonResult
): void {
  const rowChanges = compareRows(sheet1, sheet2, settings);
  result.changes.push(...rowChanges);

  const cellChanges = compareCells(sheet1.cells, sheet2.cells, settings);
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
        cellAddress: addr,
        newComment: c2.text,
        commentAuthor: c2.author,
      });
    } else if (c1 && !c2) {
      result.changes.push({
        changeType: SmlChangeType.CommentDeleted,
        cellAddress: addr,
        oldComment: c1.text,
        commentAuthor: c1.author,
      });
    } else if (c1 && c2 && (c1.text !== c2.text || c1.author !== c2.author)) {
      result.changes.push({
        changeType: SmlChangeType.CommentChanged,
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
        cellAddress: key,
        dataValidationType: dv2.type,
        newDataValidation: formatDataValidation(dv2),
      });
    } else if (dv1 && !dv2) {
      result.changes.push({
        changeType: SmlChangeType.DataValidationDeleted,
        cellAddress: key,
        dataValidationType: dv1.type,
        oldDataValidation: formatDataValidation(dv1),
      });
    } else if (dv1 && dv2 && dv1.hash !== dv2.hash) {
      result.changes.push({
        changeType: SmlChangeType.DataValidationChanged,
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
        mergedCellRange: range,
      });
    }
  }

  for (const range of merged1) {
    if (!merged2.has(range)) {
      result.changes.push({
        changeType: SmlChangeType.MergedCellDeleted,
        mergedCellRange: range,
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
        cellAddress: addr,
        newHyperlink: hl2.target,
      });
    } else if (hl1 && !hl2) {
      result.changes.push({
        changeType: SmlChangeType.HyperlinkDeleted,
        cellAddress: addr,
        oldHyperlink: hl1.target,
      });
    } else if (hl1 && hl2 && hl1.hash !== hl2.hash) {
      result.changes.push({
        changeType: SmlChangeType.HyperlinkChanged,
        cellAddress: addr,
        oldHyperlink: hl1.target,
        newHyperlink: hl2.target,
      });
    }
  }
}


