// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

import type {
  WorksheetSignature,
  WorkbookSignature,
  SmlComparerSettings,
  SmlComparisonResult,
} from './types';

interface SheetMatch {
  sheet1Name: string;
  sheet2Name: string;
  sheet1: WorksheetSignature;
  sheet2: WorksheetSignature;
}

export async function matchSheets(
  sig1: WorkbookSignature,
  sig2: WorkbookSignature,
  _settings: SmlComparerSettings
): Promise<SheetMatch[]> {
  const matches: SheetMatch[] = [];
  
  for (const [name, sheet1] of sig1.sheets) {
    const sheet2 = sig2.sheets.get(name);
    if (sheet2) {
      matches.push({ sheet1Name: name, sheet2Name: name, sheet1, sheet2 });
    }
  }
  
  return matches;
}

export async function compareSheets(
  _sig1: WorkbookSignature,
  _sig2: WorkbookSignature,
  _settings: SmlComparerSettings
): Promise<SmlComparisonResult> {
  return { changes: [] };
}
