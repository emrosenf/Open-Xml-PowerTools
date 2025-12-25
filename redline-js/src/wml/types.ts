/**
 * WML (Word) document comparison types
 *
 * Defines types and interfaces for comparing Word documents with
 * detailed change tracking for UI display.
 */

/**
 * Types of changes detected during Word document comparison.
 */
export enum WmlChangeType {
  /** Text content inserted */
  TextInserted = 'TextInserted',
  /** Text content deleted */
  TextDeleted = 'TextDeleted',
  /** Text content replaced (delete + insert at same location) */
  TextReplaced = 'TextReplaced',
  /** Entire paragraph inserted */
  ParagraphInserted = 'ParagraphInserted',
  /** Entire paragraph deleted */
  ParagraphDeleted = 'ParagraphDeleted',
  /** Text formatting changed (bold, italic, etc.) */
  FormatChanged = 'FormatChanged',
  /** Table row inserted */
  TableRowInserted = 'TableRowInserted',
  /** Table row deleted */
  TableRowDeleted = 'TableRowDeleted',
  /** Table cell content changed */
  TableCellChanged = 'TableCellChanged',
  /** Content moved from one location */
  MovedFrom = 'MovedFrom',
  /** Content moved to new location */
  MovedTo = 'MovedTo',
  /** Footnote/endnote changed */
  NoteChanged = 'NoteChanged',
  /** Image/drawing inserted */
  ImageInserted = 'ImageInserted',
  /** Image/drawing deleted */
  ImageDeleted = 'ImageDeleted',
  /** Image/drawing replaced */
  ImageReplaced = 'ImageReplaced',
}

/**
 * Word count statistics for a change.
 */
export interface WmlWordCount {
  /** Number of words deleted */
  deleted: number;
  /** Number of words inserted */
  inserted: number;
}

/**
 * Represents a single change between two Word documents.
 * This is the raw change data captured during comparison.
 */
export interface WmlChange {
  /** Type of change */
  changeType: WmlChangeType;
  /** Unique revision ID from the Word markup (w:id) */
  revisionId: number;
  /** Zero-based index of the paragraph containing this change */
  paragraphIndex?: number;
  /** For table changes, the row index */
  tableRowIndex?: number;
  /** For table changes, the cell index */
  tableCellIndex?: number;
  /** Original text (for deletions and replacements) */
  oldText?: string;
  /** New text (for insertions and replacements) */
  newText?: string;
  /** Word count statistics */
  wordCount?: WmlWordCount;
  /** For format changes, description of what changed */
  formatDescription?: string;
  /** Author who made the change */
  author?: string;
  /** Date/time of the change (ISO 8601) */
  dateTime?: string;
  /** Whether this is inside a footnote */
  inFootnote?: boolean;
  /** Whether this is inside an endnote */
  inEndnote?: boolean;
  /** Whether this is inside a table */
  inTable?: boolean;
  /** Whether this is inside a textbox */
  inTextbox?: boolean;
}

/**
 * Result of a Word document comparison with detailed change tracking.
 */
export interface WmlComparisonResult {
  /** The comparison result document as a buffer */
  document: Buffer;
  /** List of all individual changes detected */
  changes: WmlChange[];
  /** Number of insertions */
  insertions: number;
  /** Number of deletions */
  deletions: number;
  /** Number of format changes */
  formatChanges: number;
  /** Total number of revisions */
  revisionCount: number;
}

/**
 * UI-friendly representation of a change for display in a change list.
 */
export interface WmlChangeListItem {
  /** Unique identifier for this change list item */
  id: string;
  /** Type of change */
  changeType: WmlChangeType;
  /** Human-readable summary of the change */
  summary: string;
  /** Preview text showing what changed */
  previewText?: string;
  /** Word count statistics */
  wordCount?: WmlWordCount;
  /** Zero-based paragraph index for navigation */
  paragraphIndex?: number;
  /** Revision ID for navigation to the change in the document */
  revisionId?: number;
  /** Anchor string for navigation (e.g., "para-5" or "revision-12") */
  anchor?: string;
  /** Additional details about the change */
  details?: {
    /** Original text before change */
    oldText?: string;
    /** New text after change */
    newText?: string;
    /** Format change description */
    formatDescription?: string;
    /** Author of the change */
    author?: string;
    /** Date/time of change */
    dateTime?: string;
    /** Location context (e.g., "In footnote", "In table row 2") */
    locationContext?: string;
  };
}

/**
 * Options for building a change list from comparison results.
 */
export interface WmlChangeListOptions {
  /** 
   * Whether to group adjacent changes of the same type.
   * Default: true
   */
  groupAdjacentChanges?: boolean;
  /**
   * Whether to merge delete+insert pairs into "replaced" changes.
   * Default: true
   */
  mergeReplacements?: boolean;
  /**
   * Maximum length for preview text before truncation.
   * Default: 100
   */
  maxPreviewLength?: number;
}

/**
 * Settings for Word document comparison (extended).
 */
export interface WmlComparerSettings {
  /** Author name for tracked changes */
  author?: string;
  /** Date/time for tracked changes */
  dateTime?: Date;
  /** Threshold for paragraph matching (0-1) */
  detailThreshold?: number;
}
