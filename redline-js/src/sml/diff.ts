import {
  SmlChangeType,
  type WorksheetSignature,
  type CellSignature,
  type CellFormatSignature,
  type SmlChange,
} from './types';
import { computeCorrelation, CorrelationStatus } from '../core/lcs';

/**
 * Compare rows between two worksheets and compute differences.
 *
 * Uses LCS (Longest Common Subsequence) to align rows and detect
 * insertions, deletions, and modifications.
 *
 * @param sheet1 First worksheet signature
 * @param sheet2 Second worksheet signature
 * @param settings Comparison settings
 * @returns Array of detected row changes
 */
export function compareRows(
  sheet1: WorksheetSignature,
  sheet2: WorksheetSignature,
  settings: any
): SmlChange[] {
  const changes: SmlChange[] = [];

  if (!settings.enableRowAlignment) {
    compareCellsDirect(sheet1, sheet2, changes, settings);
    return changes;
  }

  // Convert rows to units for LCS
  const rows1 = convertRowsToUnits(sheet1);
  const rows2 = convertRowsToUnits(sheet2);

  // Compute LCS correlation
  const correlation = computeCorrelation(rows1, rows2);

  // Process correlation to detect changes
  let row1Index = 0;
  let row2Index = 0;

  for (const seq of correlation) {
    const seqType = seq.status;

    if (seqType === CorrelationStatus.Equal) {
      if (seq.items1 && seq.items2) {
        for (let i = 0; i < seq.items1.length; i++) {
          const unit1 = seq.items1[i];
          const unit2 = seq.items2[i];

          compareAlignedRows(unit1, unit2, sheet1, sheet2, changes, settings);
          row1Index++;
          row2Index++;
        }
      }
    } else if (seqType === CorrelationStatus.Deleted) {
      if (seq.items1) {
        for (const unit of seq.items1) {
          changes.push({
            changeType: SmlChangeType.RowDeleted,
            rowIndex: unit.row,
          });
          row1Index++;
        }
      }
    } else if (seqType === CorrelationStatus.Inserted) {
      if (seq.items2) {
        for (const unit of seq.items2) {
          changes.push({
            changeType: SmlChangeType.RowInserted,
            rowIndex: unit.row,
          });
          row2Index++;
        }
      }
    }
  }

  return changes;
}

/**
 * Convert worksheet row signatures to units for LCS comparison.
 */
function convertRowsToUnits(
  sheet: WorksheetSignature
): Array<{ row: number; hash: string }> {
  const units: Array<{ row: number; hash: string }> = [];

  for (const [row, sig] of sheet.rowSignatures) {
    units.push({ row, hash: sig });
  }

  return units;
}

/**
 * Compare two aligned rows to detect cell-level changes.
 */
function compareAlignedRows(
  unit1: { row: number; hash: string },
  unit2: { row: number; hash: string },
  sheet1: WorksheetSignature,
  sheet2: WorksheetSignature,
  changes: SmlChange[],
  settings: any
): void {
  if (unit1.hash === unit2.hash) {
    return;
  }

  const cells1 = getCellsInRow(sheet1, unit1.row);
  const cells2 = getCellsInRow(sheet2, unit2.row);

  for (const [address, cell2] of cells2) {
    const cell1 = cells1.get(address);

    if (!cell1) {
      changes.push({
        changeType: SmlChangeType.CellAdded,
        cellAddress: address,
        rowIndex: unit2.row,
        columnIndex: cell2.column,
        newValue: cell2.resolvedValue,
      });
    } else {
      if (cell1.contentHash !== cell2.contentHash) {
        if (cell1.formula !== cell2.formula) {
          changes.push({
            changeType: SmlChangeType.FormulaChanged,
            cellAddress: address,
            rowIndex: unit2.row,
            columnIndex: cell2.column,
            oldFormula: cell1.formula,
            newFormula: cell2.formula,
            oldValue: cell1.resolvedValue,
            newValue: cell2.resolvedValue,
          });
        } else {
          changes.push({
            changeType: SmlChangeType.ValueChanged,
            cellAddress: address,
            rowIndex: unit2.row,
            columnIndex: cell2.column,
            oldValue: cell1.resolvedValue,
            newValue: cell2.resolvedValue,
          });
        }
      }

      if (settings.compareFormatting && !formatsEqual(cell1.format, cell2.format)) {
        changes.push({
          changeType: SmlChangeType.FormatChanged,
          cellAddress: address,
          rowIndex: unit2.row,
          columnIndex: cell2.column,
          oldFormat: cell1.format,
          newFormat: cell2.format,
        });
      }
    }
  }

  for (const [address, cell1] of cells1) {
    if (!cells2.has(address) && (cell1.resolvedValue || cell1.formula)) {
      changes.push({
        changeType: SmlChangeType.CellDeleted,
        cellAddress: address,
        rowIndex: unit1.row,
        columnIndex: cell1.column,
        oldValue: cell1.resolvedValue,
      });
    }
  }
}

/**
 * Compare cells without row alignment (direct comparison).
 */
function compareCellsDirect(
  sheet1: WorksheetSignature,
  sheet2: WorksheetSignature,
  changes: SmlChange[],
  settings: any
): void {
  for (const [address, cell2] of sheet2.cells) {
    const cell1 = sheet1.cells.get(address);

    if (!cell1) {
      changes.push({
        changeType: SmlChangeType.CellAdded,
        cellAddress: address,
        rowIndex: cell2.row,
        columnIndex: cell2.column,
        newValue: cell2.resolvedValue,
      });
    } else {
      if (cell1.contentHash !== cell2.contentHash) {
        if (cell1.formula !== cell2.formula) {
          changes.push({
            changeType: SmlChangeType.FormulaChanged,
            cellAddress: address,
            rowIndex: cell2.row,
            columnIndex: cell2.column,
            oldFormula: cell1.formula,
            newFormula: cell2.formula,
            oldValue: cell1.resolvedValue,
            newValue: cell2.resolvedValue,
          });
        } else {
          changes.push({
            changeType: SmlChangeType.ValueChanged,
            cellAddress: address,
            rowIndex: cell2.row,
            columnIndex: cell2.column,
            oldValue: cell1.resolvedValue,
            newValue: cell2.resolvedValue,
          });
        }
      }

      if (settings.compareFormatting && !formatsEqual(cell1.format, cell2.format)) {
        changes.push({
          changeType: SmlChangeType.FormatChanged,
          cellAddress: address,
          rowIndex: cell2.row,
          columnIndex: cell2.column,
          oldFormat: cell1.format,
          newFormat: cell2.format,
        });
      }
    }
  }

  for (const [address, cell1] of sheet1.cells) {
    if (!sheet2.cells.has(address) && (cell1.resolvedValue || cell1.formula)) {
      changes.push({
        changeType: SmlChangeType.CellDeleted,
        cellAddress: address,
        rowIndex: cell1.row,
        columnIndex: cell1.column,
        oldValue: cell1.resolvedValue,
      });
    }
  }
}

function getCellsInRow(sheet: WorksheetSignature, row: number): Map<string, CellSignature> {
  const cells = new Map<string, CellSignature>();

  for (const [address, cell] of sheet.cells) {
    if (cell.row === row) {
      cells.set(address, cell);
    }
  }

  return cells;
}

function formatsEqual(f1: CellFormatSignature, f2: CellFormatSignature): boolean {
  return (
    f1.numberFormatCode === f2.numberFormatCode &&
    f1.bold === f2.bold &&
    f1.italic === f2.italic &&
    f1.underline === f2.underline &&
    f1.strikethrough === f2.strikethrough &&
    f1.fontName === f2.fontName &&
    f1.fontSize === f2.fontSize &&
    f1.fontColor === f2.fontColor &&
    f1.fillPattern === f2.fillPattern &&
    f1.fillForegroundColor === f2.fillForegroundColor &&
    f1.fillBackgroundColor === f2.fillBackgroundColor &&
    f1.borderLeftStyle === f2.borderLeftStyle &&
    f1.borderLeftColor === f2.borderLeftColor &&
    f1.borderRightStyle === f2.borderRightStyle &&
    f1.borderTopStyle === f2.borderTopStyle &&
    f1.borderTopColor === f2.borderTopColor &&
    f1.borderBottomStyle === f2.borderBottomStyle &&
    f1.borderBottomColor === f2.borderBottomColor &&
    f1.horizontalAlignment === f2.horizontalAlignment &&
    f1.verticalAlignment === f2.verticalAlignment &&
    f1.wrapText === f2.wrapText &&
    f1.indent === f2.indent
  );
}
