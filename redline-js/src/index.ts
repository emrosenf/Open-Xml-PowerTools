/**
 * @docredline/core - Document comparison library
 *
 * TypeScript port of Open-Xml-PowerTools comparers for Word, Excel, and PowerPoint.
 *
 * @packageDocumentation
 */

// Core utilities - re-export common types
export {
  type Document,
  RevisionType,
  type Revision,
  type PmlChangeType,
  type PmlComparerSettings,
  type PmlChange,
  type PmlComparisonResult,
} from './types';

// Core module exports
export * from './core';

// Word document handling
export * from './wml/document';
export * from './wml/revision';

// Word document comparison - explicit exports to avoid ambiguity with ./types
export {
  type WmlComparerSettings,
  type WmlComparisonResult,
  compareDocuments,
  countDocumentRevisions,
} from './wml/wml-comparer';

// Excel document comparison - explicit exports to avoid ambiguity with ./types
export {
  compare as compareSpreadsheets,
} from './sml/sml-comparer';

export {
  SmlChangeType,
  type SmlChange,
  type SmlComparerSettings,
  type SmlComparisonResult,
  type WorkbookSignature,
  type WorksheetSignature,
  type CellSignature,
  type CellFormatSignature,
  type HighlightColors,
  type CommentSignature,
  type DataValidationSignature,
  type HyperlinkSignature,
} from './sml/types';

// PowerPoint document comparison (to be implemented)
// export * from './pml/pml-comparer';
