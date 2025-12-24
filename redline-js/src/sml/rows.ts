// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

import {
  WorksheetSignature,
  CellSignature,
} from './types';

import {
  computeLCS,
} from '../core/diff';

/**
 * Matches worksheets between two workbooks.
 */
export async function matchSheets(
  sig1: WorkbookSignature,
  sig2: WorkbookSignature,
  settings: SmlComparerSettings
): Promise<SheetMatch[]> {
  // Implementation placeholder
  return [];
}

/**
 * Computes differences between matched worksheets.
 */
export async function compareSheets(
  sig1: WorkbookSignature,
  sig2: WorkbookSignature,
  settings: SmlComparerSettings
): Promise<SmlComparisonResult> {
  const result = new SmlComparisonResult();
  return result;
}
