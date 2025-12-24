/**
 * Core type definitions for document comparison
 */

// ============================================================================
// Common Types
// ============================================================================

/**
 * Represents a document as a byte buffer
 */
export interface Document {
  fileName: string;
  data: Buffer | Uint8Array;
}

/**
 * Revision types for tracked changes
 */
export enum RevisionType {
  Insertion = 'Insertion',
  Deletion = 'Deletion',
  MoveFrom = 'MoveFrom',
  MoveTo = 'MoveTo',
  ParagraphPropertiesChange = 'ParagraphPropertiesChange',
  RunPropertiesChange = 'RunPropertiesChange',
  SectionPropertiesChange = 'SectionPropertiesChange',
  StyleDefinitionChange = 'StyleDefinitionChange',
  StyleInsertion = 'StyleInsertion',
  NumberingChange = 'NumberingChange',
  CellDeletion = 'CellDeletion',
  CellInsertion = 'CellInsertion',
  CellMerge = 'CellMerge',
  CellPropertiesChange = 'CellPropertiesChange',
  TablePropertiesChange = 'TablePropertiesChange',
  TableGridChange = 'TableGridChange',
}

/**
 * A single revision/change in a document
 */
export interface Revision {
  type: RevisionType;
  author?: string;
  date?: string;
  text?: string;
}

// ============================================================================
// Word Document Types (WML)
// ============================================================================

/**
 * Settings for Word document comparison
 */
export interface WmlComparerSettings {
  /** Author name for tracked changes (default: 'Unknown') */
  author?: string;

  /** Date for tracked changes (default: current date) */
  dateTime?: Date;

  /** Whether to compare paragraph properties */
  compareParagraphProperties?: boolean;

  /** Whether to compare run properties */
  compareRunProperties?: boolean;

  /** Whether to compare section properties */
  compareSectionProperties?: boolean;

  /** Percentage threshold for paragraph matching (0-100) */
  matchThreshold?: number;
}

/**
 * Result of a Word document comparison
 */
export interface WmlComparisonResult {
  /** The compared document with tracked changes */
  document: Document;

  /** List of detected revisions */
  revisions: Revision[];
}

// ============================================================================
// Excel Document Types (SML)
// ============================================================================

/**
 * Change types for Excel comparison
 */
export enum SmlChangeType {
  ValueChanged = 'ValueChanged',
  FormulaChanged = 'FormulaChanged',
  FormatChanged = 'FormatChanged',
  CellAdded = 'CellAdded',
  CellDeleted = 'CellDeleted',
  SheetAdded = 'SheetAdded',
  SheetDeleted = 'SheetDeleted',
  SheetRenamed = 'SheetRenamed',
  RowInserted = 'RowInserted',
  RowDeleted = 'RowDeleted',
  ColumnInserted = 'ColumnInserted',
  ColumnDeleted = 'ColumnDeleted',
  NamedRangeAdded = 'NamedRangeAdded',
  NamedRangeDeleted = 'NamedRangeDeleted',
  NamedRangeChanged = 'NamedRangeChanged',
  MergedCellAdded = 'MergedCellAdded',
  MergedCellDeleted = 'MergedCellDeleted',
  HyperlinkAdded = 'HyperlinkAdded',
  HyperlinkChanged = 'HyperlinkChanged',
  HyperlinkDeleted = 'HyperlinkDeleted',
  DataValidationAdded = 'DataValidationAdded',
  DataValidationDeleted = 'DataValidationDeleted',
  DataValidationChanged = 'DataValidationChanged',
}

/**
 * Settings for Excel comparison
 */
export interface SmlComparerSettings {
  /** Compare cell values */
  compareValues?: boolean;

  /** Compare formulas */
  compareFormulas?: boolean;

  /** Compare cell formatting */
  compareFormatting?: boolean;

  /** Enable row alignment using LCS */
  enableRowAlignment?: boolean;

  /** Enable sheet rename detection */
  enableSheetRenameDetection?: boolean;

  /** Threshold for sheet rename similarity (0-1) */
  sheetRenameSimilarityThreshold?: number;

  /** Case insensitive value comparison */
  caseInsensitiveValues?: boolean;

  /** Numeric tolerance for floating point comparison */
  numericTolerance?: number;

  /** Compare named ranges */
  compareNamedRanges?: boolean;

  /** Compare merged cells */
  compareMergedCells?: boolean;

  /** Compare hyperlinks */
  compareHyperlinks?: boolean;

  /** Compare data validation */
  compareDataValidation?: boolean;

  /** Compare comments */
  compareComments?: boolean;
}

/**
 * A single change in an Excel comparison
 */
export interface SmlChange {
  changeType: SmlChangeType;
  sheetName: string;
  cellAddress?: string;
  rowIndex?: number;
  columnIndex?: number;
  oldValue?: string;
  newValue?: string;
  oldSheetName?: string;
}

/**
 * Result of an Excel comparison
 */
export interface SmlComparisonResult {
  changes: SmlChange[];
  totalChanges: number;
  valueChanges: number;
  formulaChanges: number;
  formatChanges: number;
  cellsAdded: number;
  cellsDeleted: number;
  sheetsAdded: number;
  sheetsDeleted: number;
  sheetsRenamed: number;
  rowsInserted: number;
  rowsDeleted: number;
  columnsInserted: number;
  columnsDeleted: number;
}

// ============================================================================
// PowerPoint Document Types (PML)
// ============================================================================

/**
 * Change types for PowerPoint comparison
 */
export enum PmlChangeType {
  SlideSizeChanged = 'SlideSizeChanged',
  SlideInserted = 'SlideInserted',
  SlideDeleted = 'SlideDeleted',
  SlideMoved = 'SlideMoved',
  ShapeInserted = 'ShapeInserted',
  ShapeDeleted = 'ShapeDeleted',
  ShapeMoved = 'ShapeMoved',
  ShapeResized = 'ShapeResized',
  TextChanged = 'TextChanged',
  ImageReplaced = 'ImageReplaced',
}

/**
 * Settings for PowerPoint comparison
 */
export interface PmlComparerSettings {
  /** Compare slide structure */
  compareSlideStructure?: boolean;

  /** Compare shape structure */
  compareShapeStructure?: boolean;

  /** Compare text content */
  compareTextContent?: boolean;

  /** Compare text formatting */
  compareTextFormatting?: boolean;

  /** Compare shape transforms (position, size) */
  compareShapeTransforms?: boolean;

  /** Compare shape styles */
  compareShapeStyles?: boolean;

  /** Compare image content */
  compareImageContent?: boolean;

  /** Compare charts */
  compareCharts?: boolean;

  /** Compare tables */
  compareTables?: boolean;

  /** Compare notes */
  compareNotes?: boolean;

  /** Compare slide transitions */
  compareTransitions?: boolean;

  /** Enable fuzzy shape matching */
  enableFuzzyShapeMatching?: boolean;

  /** Use LCS for slide alignment */
  useSlideAlignmentLCS?: boolean;

  /** Add summary slide */
  addSummarySlide?: boolean;

  /** Add notes annotations */
  addNotesAnnotations?: boolean;
}

/**
 * A single change in a PowerPoint comparison
 */
export interface PmlChange {
  changeType: PmlChangeType;
  slideNumber?: number;
  shapeName?: string;
  oldValue?: string;
  newValue?: string;
}

/**
 * Result of a PowerPoint comparison
 */
export interface PmlComparisonResult {
  changes: PmlChange[];
  totalChanges: number;
  slidesInserted: number;
  slidesDeleted: number;
  shapesInserted: number;
  shapesDeleted: number;
  shapesMoved: number;
  shapesResized: number;
  textChanges: number;
}
