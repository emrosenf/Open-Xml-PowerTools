// Copyright (c) Microsoft. All rights reserved.
// Licensed under MIT license. See LICENSE file in the project root for full license information.

import type {
  WorksheetSignature,
  CellSignature,
  SmlChange,
  SmlChangeType,
} from './types';
import { hashString } from '../core/hash';
import { computeCorrelation, CorrelationStatus } from '../core/lcs';

/**
 * Represents result of matching sheets between two workbooks.
 */
interface SheetMatch {
  type: 'added' | 'deleted' | 'renamed' | 'matched';
  oldName?: string;
  newName?: string;
  name?: string;
}

interface SheetUnit {
  name: string;
  signature: WorksheetSignature;
  hash: string;
}

/**
 * Match sheets between two workbooks using multi-pass strategy.
 *
 * Matching phases:
 * 1. Exact name match
 * 2. Content hash match
 * 3. Fuzzy similarity matching (if enabled)
 *
 * @param sig1 First workbook signature
 * @param sig2 Second workbook signature
 * @param settings Comparison settings
 * @returns Array of sheet matches
 */
export function sheetsMatch(
  sig1: any,
  sig2: any,
  settings: any
): SheetMatch[] {
  const matches: SheetMatch[] = [];
  const matched1 = new Set<string>();
  const matched2 = new Set<string>();

  const sheets1 = Array.from(sig1.sheets.entries()).map(([name, sig]) => ({
    name,
    signature: sig,
  }));
  const sheets2 = Array.from(sig2.sheets.entries()).map(([name, sig]) => ({
    name,
    signature: sig,
  }));

  for (const sheet1 of sheets1) {
    if (matched1.has(sheet1.name)) continue;

    const sheet2 = sheets2.find(
      (s) => s.name === sheet1.name && !matched2.has(s.name)
    );

    if (sheet2) {
      matches.push({
        type: 'matched',
        name: sheet1.name,
      });
      matched1.add(sheet1.name);
      matched2.add(sheet2.name);
    }
  }

  const remaining1 = sheets1.filter((s) => !matched1.has(s.name));
  const remaining2 = sheets2.filter((s) => !matched2.has(s.name));

  for (const sheet1 of remaining1) {
    const hash1 = computeSheetHash(sheet1.signature);

    for (const sheet2 of remaining2) {
      if (matched2.has(sheet2.name)) continue;

      const hash2 = computeSheetHash(sheet2.signature);

      if (hash1 === hash2) {
        matches.push({
          type: 'renamed',
          oldName: sheet1.name,
          newName: sheet2.name,
        });
        matched1.add(sheet1.name);
        matched2.add(sheet2.name);
        break;
      }
    }
  }

  if (settings.enableFuzzyShapeMatching) {
    const fuzzyRemaining1 = sheets1.filter((s) => !matched1.has(s.name));
    const fuzzyRemaining2 = sheets2.filter((s) => !matched2.has(s.name));

    for (const sheet1 of fuzzyRemaining1) {
      let bestMatch: { name: string; similarity: number } | null = null;

      for (const sheet2 of fuzzyRemaining2) {
        if (matched2.has(sheet2.name)) continue;

        const similarity = computeSheetSimilarity(sheet1.signature, sheet2.signature);

        if (
          similarity > settings.sheetSimilarityThreshold &&
          (!bestMatch || similarity > bestMatch.similarity)
        ) {
          bestMatch = { name: sheet2.name, similarity };
        }
      }

      if (bestMatch && bestMatch.similarity > settings.sheetSimilarityThreshold) {
        matches.push({
          type: 'renamed',
          oldName: sheet1.name,
          newName: bestMatch.name,
        });
        matched1.add(sheet1.name);
        matched2.add(bestMatch.name);
      }
    }
  }

  const deletedSheets = sheets1.filter((s) => !matched1.has(s.name));
  const addedSheets = sheets2.filter((s) => !matched2.has(s.name));

  for (const sheet of deletedSheets) {
    matches.push({
      type: 'deleted',
      oldName: sheet.name,
    });
  }

  for (const sheet of addedSheets) {
    matches.push({
      type: 'added',
      newName: sheet.name,
    });
  }

  return matches;
}

/**
 * Compute a hash of a worksheet's content for matching.
 *
 * Combines row signatures to create a unique fingerprint.
 */
function computeSheetHash(sheet: WorksheetSignature): string {
  const rowSigs: string[] = [];

  for (const [row, sig] of sheet.rowSignatures) {
    rowSigs.push(`${row}:${sig}`);
  }

  return hashString(rowSigs.join('|'));
}

/**
 * Compute similarity between two worksheets using LCS on row signatures.
 *
 * Returns a value between 0 and 1.
 */
function computeSheetSimilarity(
  sheet1: WorksheetSignature,
  sheet2: WorksheetSignature
): number {
  const rows1: Array<{ row: number; hash: string }> = [];
  const rows2: Array<{ row: number; hash: string }> = [];

  for (const [row, sig] of sheet1.rowSignatures) {
    rows1.push({ row, hash: sig });
  }

  for (const [row, sig] of sheet2.rowSignatures) {
    rows2.push({ row, hash: sig });
  }

  const correlation = computeCorrelation(rows1, rows2);

  let equalCount = 0;
  let total1 = 0;
  let total2 = 0;

  for (const seq of correlation) {
    if (seq.status === CorrelationStatus.Equal && seq.items1) {
      equalCount += seq.items1.length;
    }
    if (seq.items1) total1 += seq.items1.length;
    if (seq.items2) total2 += seq.items2.length;
  }

  const total = Math.max(total1, total2);
  return total > 0 ? equalCount / total : 0;
}
