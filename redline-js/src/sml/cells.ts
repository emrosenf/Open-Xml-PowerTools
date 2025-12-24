import type { CellSignature, SmlChange, SmlChangeType } from './types';
import { computeCorrelation } from '../core/lcs';

/**
 * Compare cells and compute differences.
 *
 * @param cells1 Cells from first sheet
 * @param cells2 Cells from second sheet
 * @param settings Comparison settings
 * @returns Array of detected changes
 */
export function compareCells(
  cells1: Map<string, CellSignature>,
  cells2: Map<string, CellSignature>,
  settings: any
): SmlChange[] {
  const changes: SmlChange[] = [];
  const processedAddresses = new Set<string>();

  // Check for cell deletions
  for (const [address, cell1] of cells1) {
    if (cells2.has(address)) {
      processedAddresses.add(address);
    } else if (cell1.resolvedValue || cell1.formula) {
      changes.push({
        changeType: SmlChangeType.CellDeleted,
        cellAddress: address,
        rowIndex: cell1.row,
        columnIndex: cell1.column,
        oldValue: cell1.resolvedValue,
        oldFormula: cell1.formula,
        oldFormat: cell1.format,
      });
    }
  }

  // Check for cell additions and modifications
  for (const [address, cell2] of cells2) {
    const cell1 = cells1.get(address);

    if (!cell1) {
      if (cell2.resolvedValue || cell2.formula) {
        changes.push({
          changeType: SmlChangeType.CellAdded,
          cellAddress: address,
          rowIndex: cell2.row,
          columnIndex: cell2.column,
          newValue: cell2.resolvedValue,
          newFormula: cell2.formula,
          newFormat: cell2.format,
        });
      }
    } else if (cell1.contentHash !== cell2.contentHash) {
      compareCellValues(cell1, cell2, address, changes, settings);
    }
  }

  return changes;
}

/**
 * Compare two cells at the same address and detect specific changes.
 */
function compareCellValues(
  cell1: CellSignature,
  cell2: CellSignature,
  address: string,
  changes: SmlChange[],
  settings: any
): void {
  const hasValueChange =
    settings.compareValues !== false &&
    cell1.resolvedValue !== cell2.resolvedValue;

  const hasFormulaChange =
    settings.compareFormulas !== false &&
    cell1.formula !== cell2.formula;

  const hasFormatChange =
    settings.compareFormatting !== false &&
    !formatsEqual(cell1.format, cell2.format);

  if (hasValueChange && hasFormulaChange) {
    changes.push({
      changeType: SmlChangeType.ValueChanged,
      cellAddress: address,
      rowIndex: cell1.row,
      columnIndex: cell1.column,
      oldValue: cell1.resolvedValue,
      newValue: cell2.resolvedValue,
    });

    changes.push({
      changeType: SmlChangeType.FormulaChanged,
      cellAddress: address,
      rowIndex: cell1.row,
      columnIndex: cell1.column,
      oldFormula: cell1.formula,
      newFormula: cell2.formula,
    });
  } else if (hasValueChange) {
    changes.push({
      changeType: SmlChangeType.ValueChanged,
      cellAddress: address,
      rowIndex: cell1.row,
      columnIndex: cell1.column,
      oldValue: cell1.resolvedValue,
      newValue: cell2.resolvedValue,
    });
  } else if (hasFormulaChange) {
    changes.push({
      changeType: SmlChangeType.FormulaChanged,
      cellAddress: address,
      rowIndex: cell1.row,
      columnIndex: cell1.column,
      oldFormula: cell1.formula,
      newFormula: cell2.formula,
    });
  }

  if (hasFormatChange) {
    changes.push({
      changeType: SmlChangeType.FormatChanged,
      cellAddress: address,
      rowIndex: cell1.row,
      columnIndex: cell1.column,
      oldFormat: cell1.format,
      newFormat: cell2.format,
    });
  }
}

/**
 * Compare cell format signatures for equality.
 */
function formatsEqual(
  fmt1: any,
  fmt2: any
): boolean {
  if (!fmt1 && !fmt2) return true;
  if (!fmt1 || !fmt2) return false;

  return (
    fmt1.numberFormatCode === fmt2.numberFormatCode &&
    fmt1.bold === fmt2.bold &&
    fmt1.italic === fmt2.italic &&
    fmt1.underline === fmt2.underline &&
    fmt1.strikethrough === fmt2.strikethrough &&
    fmt1.fontName === fmt2.fontName &&
    fmt1.fontSize === fmt2.fontSize &&
    fmt1.fontColor === fmt2.fontColor &&
    fmt1.fillPattern === fmt2.fillPattern &&
    fmt1.fillForegroundColor === fmt2.fillForegroundColor &&
    fmt1.fillBackgroundColor === fmt2.fillBackgroundColor &&
    fmt1.borderLeftStyle === fmt2.borderLeftStyle &&
    fmt1.borderLeftColor === fmt2.borderLeftColor &&
    fmt1.borderRightStyle === fmt2.borderRightStyle &&
    fmt1.borderTopStyle === fmt2.borderTopStyle &&
    fmt1.borderTopColor === fmt2.borderTopColor &&
    fmt1.borderBottomStyle === fmt2.borderBottomStyle &&
    fmt1.borderBottomColor === fmt2.borderBottomColor &&
    fmt1.horizontalAlignment === fmt2.horizontalAlignment &&
    fmt1.verticalAlignment === fmt2.verticalAlignment &&
    fmt1.wrapText === fmt2.wrapText &&
    fmt1.indent === fmt2.indent
  );
}
