// Copyright (c) Microsoft. All rights reserved.
// Licensed under MIT license. See LICENSE file in the project root for full license information.

/**
 * SmlComparer - Excel spreadsheet comparison
 *
 * Compares two Excel spreadsheets and produces a structured result document.
 *
 * This is a TypeScript port of C# SmlComparer from Open-Xml-PowerTools.
 */

import type {
  SmlChange,
  SmlChangeType,
  SmlComparisonResult,
  SmlComparerSettings,
  WorkbookSignature,
  WorksheetSignature,
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

/**
 * Compare two worksheets and add detected changes to result.
 */
function compareSheets(
  sheet1: WorksheetSignature,
  sheet2: WorksheetSignature,
  settings: SmlComparerSettings,
  result: SmlComparisonResult
): void {
  // Compare rows (with LCS-based alignment if enabled)
  const rowChanges = compareRows(sheet1, sheet2, settings);
  result.changes.push(...rowChanges);

  // Compare cell-level changes (values, formulas, formatting)
  const cellChanges = compareCells(sheet1.cells, sheet2.cells, settings);
  result.changes.push(...cellChanges);
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

export { compare };
