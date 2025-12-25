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
  type PmlChangeListItem,
  type PmlChangeListOptions,
} from './types';

// Core module exports
export * from './core';

// Word document handling
export * from './wml/document';
export * from './wml/revision';

// Word document comparison - explicit exports to avoid ambiguity with ./types
export {
  WmlChangeType,
  type WmlChange,
  type WmlComparerSettings,
  type WmlComparisonResult,
  type WmlChangeListItem,
  type WmlChangeListOptions,
  type WmlWordCount,
  compareDocuments,
  countDocumentRevisions,
  buildChangeList as buildWmlChangeList,
} from './wml/wml-comparer';

// Excel document comparison - explicit exports to avoid ambiguity with ./types
export {
  compare as compareSpreadsheets,
  produceMarkedWorkbook,
  buildChangeList,
} from './sml/sml-comparer';

export {
  SmlChangeType,
  type SmlChange,
  type SmlComparerSettings,
  type SmlComparisonResult,
  type SmlChangeListItem,
  type SmlChangeListOptions,
  type WorkbookSignature,
  type WorksheetSignature,
  type CellSignature,
  type CellFormatSignature,
  type HighlightColors,
  type CommentSignature,
  type DataValidationSignature,
  type HyperlinkSignature,
} from './sml/types';

export {
  comparePresentations,
  produceMarkedPresentation,
  canonicalizePresentationDocument,
  buildChangeList as buildPmlChangeList,
} from './pml/pml-comparer';
