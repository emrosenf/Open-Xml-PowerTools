// Copyright (c) Microsoft. All rights reserved.
// Licensed under MIT license. See LICENSE file in project root for full license information.

/**
 * Excel spreadsheet comparison types
 *
 * Defines types and interfaces for comparing Excel spreadsheets.
 */

/**
 * Settings for controlling spreadsheet comparison behavior.
 */
export interface SmlComparerSettings {
  compareValues?: boolean;
  compareFormulas?: boolean;
  compareFormatting?: boolean;
  compareSheetStructure?: boolean;
  compareComments?: boolean;
  compareDataValidations?: boolean;
  compareMergedCells?: boolean;
  compareHyperlinks?: boolean;
  caseInsensitiveValues?: boolean;
  numericTolerance?: number;
  enableRowAlignment?: boolean;
  enableColumnAlignment?: boolean;
  enableSheetRenameDetection?: boolean;
  sheetRenameSimilarityThreshold?: number;
  enableFuzzyShapeMatching?: boolean;
  slideSimilarityThreshold?: number;
  positionTolerance?: number;
  authorForChanges?: string;
  highlightColors?: HighlightColors;
}

/**
 * Highlight colors for change markup.
 */
export interface HighlightColors {
  addedCellColor?: string;
  deletedCellColor?: string;
  modifiedValueColor?: string;
  modifiedFormulaColor?: string;
  modifiedFormatColor?: string;
  insertedRowColor?: string;
  deletedRowColor?: string;
  namedRangeChangeColor?: string;
  commentChangeColor?: string;
  dataValidationChangeColor?: string;
  mergedCellRangeColor?: string;
}

/**
 * Internal canonical representation of a workbook for comparison.
 */
export interface WorkbookSignature {
  sheets: Map<string, WorksheetSignature>;
  definedNames: Map<string, string>;
}

/**
 * Internal canonical representation of a worksheet for comparison.
 */
export interface WorksheetSignature {
  name: string;
  relationshipId: string;
  cells: Map<string, CellSignature>;
  populatedRows: Set<number>;
  populatedColumns: Set<number>;
  rowSignatures: Map<number, string>;
  columnSignatures: Map<number, string>;
  comments: Map<string, CommentSignature>;
  dataValidations: Map<string, DataValidationSignature>;
  mergedCellRanges: Set<string>;
  hyperlinks: Map<string, HyperlinkSignature>;
}

/**
 * Internal canonical representation of a cell for comparison.
 */
export interface CellSignature {
  address: string;
  row: number;
  column: number;
  resolvedValue: string;
  formula: string;
  contentHash: string;
  format: CellFormatSignature;
}

/**
 * Represents expanded formatting of a cell for comparison purposes.
 * Style indices are resolved to actual formatting properties.
 */
export interface CellFormatSignature {
  numberFormatCode?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  fontName?: string;
  fontSize?: number;
  fontColor?: string;
  fillPattern?: string;
  fillForegroundColor?: string;
  fillBackgroundColor?: string;
  borderLeftStyle?: string;
  borderLeftColor?: string;
  borderRightStyle?: string;
  borderTopStyle?: string;
  borderTopColor?: string;
  borderBottomStyle?: string;
  borderBottomColor?: string;
  horizontalAlignment?: string;
  verticalAlignment?: string;
  wrapText?: boolean;
  indent?: number;
}

/**
 * Types of changes detected during spreadsheet comparison.
 */
export enum SmlChangeType {
  SheetAdded,
  SheetDeleted,
  SheetRenamed,
  RowInserted,
  RowDeleted,
  ColumnInserted,
  ColumnDeleted,
  CellAdded,
  CellDeleted,
  ValueChanged,
  FormulaChanged,
  FormatChanged,
  NamedRangeAdded,
  NamedRangeDeleted,
  NamedRangeChanged,
  CommentAdded,
  CommentDeleted,
  CommentChanged,
  DataValidationAdded,
  DataValidationDeleted,
  DataValidationChanged,
  MergedCellAdded,
  MergedCellDeleted,
  ConditionalFormatAdded,
  ConditionalFormatDeleted,
  ConditionalFormatChanged,
  HyperlinkAdded,
  HyperlinkDeleted,
  HyperlinkChanged,
}

/**
 * Represents a single change between two spreadsheets.
 */
export interface SmlChange {
  changeType: SmlChangeType;
  sheetName?: string;
  rowIndex?: number;
  columnIndex?: number;
  oldSheetName?: string;
  newSheetName?: string;
  cellAddress?: string;
  cellRange?: string;
  oldValue?: string;
  newValue?: string;
  oldFormula?: string;
  newFormula?: string;
  oldFormat?: CellFormatSignature;
  newFormat?: CellFormatSignature;
  namedRangeName?: string;
  oldNamedRangeValue?: string;
  newNamedRangeValue?: string;
  oldComment?: string;
  newComment?: string;
  commentAuthor?: string;
  dataValidationType?: string;
  oldDataValidation?: string;
  newDataValidation?: string;
  mergedCellRange?: string;
  oldHyperlink?: string;
  newHyperlink?: string;
}

/**
 * Result of comparing two spreadsheets, containing all detected changes.
 */
export interface SmlComparisonResult {
  changes: SmlChange[];
}

export interface SmlChangeListItem {
  id: string;
  changeType: SmlChangeType;
  sheetName?: string;
  cellAddress?: string;
  cellRange?: string;
  rowIndex?: number;
  columnIndex?: number;
  count?: number;
  summary: string;
  details?: {
    oldValue?: string;
    newValue?: string;
    oldFormula?: string;
    newFormula?: string;
    oldFormat?: CellFormatSignature;
    newFormat?: CellFormatSignature;
    oldComment?: string;
    newComment?: string;
    commentAuthor?: string;
    dataValidationType?: string;
    oldDataValidation?: string;
    newDataValidation?: string;
    mergedCellRange?: string;
    oldHyperlink?: string;
    newHyperlink?: string;
    oldSheetName?: string;
    newSheetName?: string;
  };
  anchor?: string;
}

export interface SmlChangeListOptions {
  groupAdjacentCells?: boolean;
}

/**
 * Represents a cell comment for comparison.
 */
export interface CommentSignature {
  cellAddress: string;
  author: string;
  text: string;
  hash: string;
}

/**
 * Represents a data validation rule for comparison.
 */
export interface DataValidationSignature {
  cellRange: string;
  type: string;
  operator?: string;
  formula1?: string;
  formula2?: string;
  allowBlank?: boolean;
  showDropDown?: boolean;
  showInputMessage?: boolean;
  showErrorMessage?: boolean;
  errorTitle?: string;
  error?: string;
  promptTitle?: string;
  prompt?: string;
  hash: string;
}

/**
 * Represents a hyperlink for comparison.
 */
export interface HyperlinkSignature {
  cellAddress: string;
  target: string;
  hash: string;
}
