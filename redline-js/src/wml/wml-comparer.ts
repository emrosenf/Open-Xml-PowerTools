/**
 * WmlComparer - Word document comparison
 *
 * Compares two Word documents and produces a result document with tracked changes
 * showing insertions and deletions.
 *
 * This is a TypeScript port of the C# WmlComparer from Open-Xml-PowerTools.
 */

import {
  loadWordDocument,
  extractParagraphs,
  extractParagraphText,
  getDocumentBody,
  type WordDocument,
} from './document';
import {
  createInsertion,
  createDeletion,
  createRun,
  createParagraph,
  countRevisions,
  resetRevisionIdCounter,
  type RevisionSettings,
  DEFAULT_REVISION_SETTINGS,
} from './revision';
import {
  computeCorrelation,
  CorrelationStatus,
  type Hashable,
  type CorrelatedSequence,
} from '../core/lcs';
import { hashString } from '../core/hash';
import {
  getTagName,
  getChildren,
  getTextContent,
  cloneNode,
  findNodes,
  type XmlNode,
} from '../core/xml';
import {
  openPackage,
  setPartFromXml,
  savePackage,
  type OoxmlPackage,
} from '../core/package';

/**
 * Settings for Word document comparison
 */
export interface WmlComparerSettings {
  /** Author name for tracked changes */
  author?: string;
  /** Date/time for tracked changes */
  dateTime?: Date;
  /** Threshold for paragraph matching (0-1) */
  detailThreshold?: number;
}

/**
 * Result of a Word document comparison
 */
export interface WmlComparisonResult {
  /** The comparison result document as a buffer */
  document: Buffer;
  /** Number of insertions */
  insertions: number;
  /** Number of deletions */
  deletions: number;
  /** Total number of revisions */
  revisionCount: number;
}

/**
 * A comparison unit representing a paragraph or table row with its hash
 */
interface ParagraphUnit extends Hashable {
  hash: string;
  node: XmlNode;
  text: string;
  /** If true, this represents a table row, not a single paragraph */
  isTableRow?: boolean;
  /** For table rows, the paragraphs within the row */
  rowParagraphs?: XmlNode[];
  /** For table rows, paragraphs grouped by cell */
  rowCells?: XmlNode[][];
}

/**
 * A comparison unit representing a word/token
 */
interface WordUnit extends Hashable {
  hash: string;
  text: string;
}

/**
 * Element types for comparison units.
 * Table rows are treated as single units, other paragraphs are individual units.
 */
interface ComparisonElement {
  type: 'paragraph' | 'tableRow';
  node: XmlNode;
  paragraphs: XmlNode[]; // For rows, all paragraphs in the row
  cells?: XmlNode[][]; // For rows, paragraphs grouped by cell
}

/**
 * Find all top-level comparison elements in a node.
 * - Paragraphs outside tables are returned individually
 * - Table rows are returned as single units (with all their paragraphs)
 * - Paragraphs inside textboxes are part of their containing paragraph
 */
function findTopLevelElements(node: XmlNode): ComparisonElement[] {
  const elements: ComparisonElement[] = [];

  function extractParagraphsFromNode(n: XmlNode, insideTextbox: boolean): XmlNode[] {
    const tagName = getTagName(n);
    const result: XmlNode[] = [];

    if (tagName === 'w:txbxContent') {
      insideTextbox = true;
    }

    if (tagName === 'w:p' && !insideTextbox) {
      result.push(n);
    }

    // Handle mc:AlternateContent - prefer mc:Fallback
    if (tagName === 'mc:AlternateContent') {
      const children = getChildren(n);
      const fallback = children.find((c) => getTagName(c) === 'mc:Fallback');
      if (fallback) {
        return extractParagraphsFromNode(fallback, insideTextbox);
      }
      const choice = children.find((c) => getTagName(c) === 'mc:Choice');
      if (choice) {
        return extractParagraphsFromNode(choice, insideTextbox);
      }
    }

    for (const child of getChildren(n)) {
      result.push(...extractParagraphsFromNode(child, insideTextbox));
    }

    return result;
  }

  function walk(n: XmlNode, insideTextbox: boolean): void {
    const tagName = getTagName(n);

    // Track textbox context
    if (tagName === 'w:txbxContent') {
      insideTextbox = true;
    }

    // Table row - extract as a single comparison unit
    if (tagName === 'w:tr') {
      const rowParas = extractParagraphsFromNode(n, insideTextbox);
      // Also extract paragraphs grouped by cell
      const cells: XmlNode[][] = [];
      for (const child of getChildren(n)) {
        if (getTagName(child) === 'w:tc') {
          const cellParas = extractParagraphsFromNode(child, insideTextbox);
          cells.push(cellParas);
        }
      }
      elements.push({
        type: 'tableRow',
        node: n,
        paragraphs: rowParas,
        cells,
      });
      return; // Don't recurse into table row children
    }

    // Regular paragraph outside table
    if (tagName === 'w:p' && !insideTextbox) {
      elements.push({
        type: 'paragraph',
        node: n,
        paragraphs: [n],
      });
      // Recurse into paragraph for nested content (textboxes)
      for (const child of getChildren(n)) {
        walk(child, insideTextbox);
      }
      return;
    }

    // Handle mc:AlternateContent
    if (tagName === 'mc:AlternateContent') {
      const children = getChildren(n);
      const fallback = children.find((c) => getTagName(c) === 'mc:Fallback');
      if (fallback) {
        walk(fallback, insideTextbox);
        return;
      }
      const choice = children.find((c) => getTagName(c) === 'mc:Choice');
      if (choice) {
        walk(choice, insideTextbox);
        return;
      }
    }

    // Recurse into children
    for (const child of getChildren(n)) {
      walk(child, insideTextbox);
    }
  }

  walk(node, false);
  return elements;
}

/**
 * Compare two Word documents and produce a result with tracked changes.
 *
 * @param source1 The original document
 * @param source2 The modified document
 * @param settings Comparison settings
 * @returns Comparison result with the marked-up document
 */
export async function compareDocuments(
  source1: Buffer | Uint8Array,
  source2: Buffer | Uint8Array,
  settings: WmlComparerSettings = {}
): Promise<WmlComparisonResult> {
  // Reset revision IDs for consistent output
  resetRevisionIdCounter();

  // Load both documents
  const doc1 = await loadWordDocument(source1);
  const doc2 = await loadWordDocument(source2);

  // Create revision settings
  const revisionSettings: RevisionSettings = {
    author: settings.author ?? 'redline-js',
    dateTime: (settings.dateTime ?? new Date()).toISOString(),
  };

  // Extract paragraphs from both documents
  const paras1 = extractParagraphUnits(doc1);
  const paras2 = extractParagraphUnits(doc2);

  // Compare at paragraph level first
  const paraCorrelation = computeCorrelation(paras1, paras2, {
    detailThreshold: settings.detailThreshold ?? 0.0,
  });

  // Build result paragraphs
  const resultParagraphs = buildResultParagraphs(paraCorrelation, revisionSettings);

  // Clone the second document as the base for the result
  const resultPkg = await openPackage(source2);

  // Get the document body and replace its content
  const docBody = getDocumentBody(doc2);
  if (docBody) {
    // Build new body content with the result paragraphs
    const newBody: XmlNode = {
      'w:body': resultParagraphs,
    };

    // Find and update the document in the package
    const mainDocXml = doc2.mainDocument;
    updateDocumentBody(mainDocXml, resultParagraphs);
    setPartFromXml(resultPkg, 'word/document.xml', mainDocXml);
  }

  // Save the result
  const resultBuffer = await savePackage(resultPkg);

  // Count revisions
  const counts = countRevisions(resultParagraphs);

  return {
    document: resultBuffer,
    insertions: counts.insertions,
    deletions: counts.deletions,
    revisionCount: counts.total,
  };
}

/**
 * Extract paragraph units from a document for comparison.
 * Uses extractParagraphText which properly handles existing tracked changes
 * by accepting revisions (skipping w:del content).
 * Includes paragraphs from main body, footnotes, and endnotes.
 *
 * IMPORTANT:
 * - Paragraphs inside textboxes are NOT extracted as separate units
 * - Table rows are treated as single comparison units
 */
function extractParagraphUnits(doc: WordDocument): ParagraphUnit[] {
  const units: ParagraphUnit[] = [];

  // Extract from main document body
  const body = getDocumentBody(doc);
  if (body) {
    // Find top-level elements (paragraphs and table rows)
    const elements = findTopLevelElements(body);
    for (const element of elements) {
      if (element.type === 'tableRow') {
        // Table row - combine all paragraph texts for hashing
        const texts = element.paragraphs.map((p) => extractParagraphText(p));
        const combinedText = texts.join(' ');
        units.push({
          hash: hashString(combinedText),
          node: cloneNode(element.node),
          text: combinedText,
          isTableRow: true,
          rowParagraphs: element.paragraphs.map(cloneNode),
          rowCells: element.cells?.map((cell) => cell.map(cloneNode)),
        });
      } else {
        // Regular paragraph
        const text = extractParagraphText(element.node);
        units.push({
          hash: hashString(text),
          node: cloneNode(element.node),
          text,
        });
      }
    }
  }

  // Extract from footnotes
  if (doc.footnotes) {
    const footnoteParas = extractFootnoteEndnoteParagraphs(doc.footnotes, 'w:footnote');
    units.push(...footnoteParas);
  }

  // Extract from endnotes
  if (doc.endnotes) {
    const endnoteParas = extractFootnoteEndnoteParagraphs(doc.endnotes, 'w:endnote');
    units.push(...endnoteParas);
  }

  return units;
}

/**
 * Extract paragraphs from footnotes or endnotes XML.
 * Skips separator and continuationSeparator notes.
 */
function extractFootnoteEndnoteParagraphs(
  xmlNodes: XmlNode[],
  noteTagName: string
): ParagraphUnit[] {
  const units: ParagraphUnit[] = [];

  for (const node of xmlNodes) {
    const tagName = getTagName(node);

    if (tagName === 'w:footnotes' || tagName === 'w:endnotes') {
      // Process children
      const children = getChildren(node);
      for (const child of children) {
        if (getTagName(child) === noteTagName) {
          // Skip separator and continuationSeparator
          const attrs = child[':@'] as Record<string, string> | undefined;
          const noteType = attrs?.['@_w:type'];
          if (noteType === 'separator' || noteType === 'continuationSeparator') {
            continue;
          }

          // Extract paragraphs from this note
          const paragraphs = findNodes(child, (n) => getTagName(n) === 'w:p');
          for (const para of paragraphs) {
            const text = extractParagraphText(para);
            if (text.trim()) {
              // Only include non-empty paragraphs
              units.push({
                hash: hashString(text),
                node: cloneNode(para),
                text,
              });
            }
          }
        }
      }
    }
  }

  return units;
}

/**
 * Build result paragraphs from correlation data
 */
function buildResultParagraphs(
  correlation: CorrelatedSequence<ParagraphUnit>[],
  settings: RevisionSettings
): XmlNode[] {
  const result: XmlNode[] = [];

  for (const seq of correlation) {
    switch (seq.status) {
      case CorrelationStatus.Equal:
        // For equal paragraphs, compare at word level for finer granularity
        if (seq.items1 && seq.items2) {
          for (let i = 0; i < seq.items1.length; i++) {
            const para1 = seq.items1[i];
            const para2 = seq.items2[i];

            // If the text is identical, just use the original paragraph
            if (para1.text === para2.text) {
              result.push(cloneNode(para2.node));
            } else {
              // Compare at word level
              const wordResult = compareWordsInParagraph(para1, para2, settings);
              result.push(wordResult);
            }
          }
        }
        break;

      case CorrelationStatus.Deleted:
        // Deleted paragraphs
        if (seq.items1) {
          for (const para of seq.items1) {
            const deletedPara = createDeletedParagraph(para, settings);
            result.push(deletedPara);
          }
        }
        break;

      case CorrelationStatus.Inserted:
        // Inserted paragraphs
        if (seq.items2) {
          for (const para of seq.items2) {
            const insertedPara = createInsertedParagraph(para, settings);
            result.push(insertedPara);
          }
        }
        break;
    }
  }

  return result;
}

/**
 * Compare words within paragraphs that have the same hash but different text
 */
function compareWordsInParagraph(
  para1: ParagraphUnit,
  para2: ParagraphUnit,
  settings: RevisionSettings
): XmlNode {
  // Tokenize both paragraphs
  const words1 = tokenize(para1.text);
  const words2 = tokenize(para2.text);

  // Compute word-level correlation
  const wordCorrelation = computeCorrelation(words1, words2);

  // Build runs from the correlation
  const runs: XmlNode[] = [];

  for (const seq of wordCorrelation) {
    switch (seq.status) {
      case CorrelationStatus.Equal:
        if (seq.items1) {
          const text = seq.items1.map((w) => w.text).join(' ');
          runs.push(createRun(text + ' '));
        }
        break;

      case CorrelationStatus.Deleted:
        if (seq.items1) {
          const text = seq.items1.map((w) => w.text).join(' ');
          const run = createRun(text + ' ');
          runs.push(createDeletion(run, settings));
        }
        break;

      case CorrelationStatus.Inserted:
        if (seq.items2) {
          const text = seq.items2.map((w) => w.text).join(' ');
          const run = createRun(text + ' ');
          runs.push(createInsertion(run, settings));
        }
        break;
    }
  }

  // Create paragraph with the runs
  // Preserve original paragraph properties if available
  const pPr = getChildren(para2.node).find((c) => getTagName(c) === 'w:pPr');
  return createParagraph(runs, pPr ? cloneNode(pPr) : undefined);
}

/**
 * Create a paragraph marked as deleted
 */
function createDeletedParagraph(para: ParagraphUnit, settings: RevisionSettings): XmlNode {
  const run = createRun(para.text);
  const deletion = createDeletion(run, settings);

  const pPr = getChildren(para.node).find((c) => getTagName(c) === 'w:pPr');
  return createParagraph([deletion], pPr ? cloneNode(pPr) : undefined);
}

/**
 * Create a paragraph marked as inserted
 */
function createInsertedParagraph(para: ParagraphUnit, settings: RevisionSettings): XmlNode {
  const run = createRun(para.text);
  const insertion = createInsertion(run, settings);

  const pPr = getChildren(para.node).find((c) => getTagName(c) === 'w:pPr');
  return createParagraph([insertion], pPr ? cloneNode(pPr) : undefined);
}

/**
 * Tokenize text into words for comparison.
 * Separates punctuation from words for finer-grained matching.
 */
function tokenize(text: string): WordUnit[] {
  // Split on whitespace, then further split on punctuation boundaries
  // This allows "12,34" to become ["12", ",", "34"] and "Test." to become ["Test", "."]
  const tokens: string[] = [];

  // First split on whitespace
  const parts = text.split(/\s+/).filter((p) => p.length > 0);

  for (const part of parts) {
    // Split each part, keeping punctuation as separate tokens
    // Match: word characters OR non-word characters (but not mixing)
    const subTokens = part.match(/\w+|[^\w\s]+/g) || [part];
    tokens.push(...subTokens);
  }

  // Use exact token text as hash for case-sensitive comparison
  // This ensures "Three" and "THree" are detected as different
  return tokens.map((token) => ({
    hash: token, // Case-sensitive matching
    text: token,
  }));
}

/**
 * Update the document body in the main document XML
 */
function updateDocumentBody(mainDocument: XmlNode[], newContent: XmlNode[]): void {
  // Find the w:document element
  for (const node of mainDocument) {
    if (getTagName(node) === 'w:document') {
      const docChildren = getChildren(node);
      // Find and update w:body
      for (let i = 0; i < docChildren.length; i++) {
        if (getTagName(docChildren[i]) === 'w:body') {
          // Replace body content
          docChildren[i] = {
            'w:body': newContent,
            ':@': docChildren[i][':@'],
          };
          // Update the document node
          (node['w:document'] as XmlNode[]) = docChildren;
          return;
        }
      }
    }
  }
}

/**
 * Simple comparison of two documents returning just the revision count.
 * Useful for quick validation without needing the full result document.
 *
 * Revision counting rules:
 * - A contiguous group of word-level changes within a paragraph = 1 revision
 * - An entire paragraph deleted = 1 revision
 * - An entire paragraph inserted = 1 revision
 * - Adjacent delete+insert at same position = 1 revision (replacement)
 */
export async function countDocumentRevisions(
  source1: Buffer | Uint8Array,
  source2: Buffer | Uint8Array,
  settings: WmlComparerSettings = {}
): Promise<{ insertions: number; deletions: number; total: number }> {
  const doc1 = await loadWordDocument(source1);
  const doc2 = await loadWordDocument(source2);

  const paras1 = extractParagraphUnits(doc1);
  const paras2 = extractParagraphUnits(doc2);

  // Use position-based comparison for paragraphs
  // This allows us to compare corresponding paragraphs even if hashes differ
  const result = compareAlignedParagraphs(paras1, paras2);

  return result;
}

/**
 * Compare paragraphs using LCS-based alignment.
 * This properly handles paragraph insertions and deletions.
 */
function compareAlignedParagraphs(
  paras1: ParagraphUnit[],
  paras2: ParagraphUnit[]
): { insertions: number; deletions: number; total: number } {
  let insertions = 0;
  let deletions = 0;

  // First, try LCS alignment at paragraph level
  const paraCorrelation = computeCorrelation(paras1, paras2);

  // Process correlation, looking for adjacent delete-insert pairs
  // that represent paragraph modifications rather than true deletions/insertions
  for (let i = 0; i < paraCorrelation.length; i++) {
    const seq = paraCorrelation[i];
    const nextSeq = paraCorrelation[i + 1];

    if (seq.status === CorrelationStatus.Deleted && seq.items1) {
      // Check if this is followed by an insert of same length (paragraph modification)
      if (
        nextSeq &&
        nextSeq.status === CorrelationStatus.Inserted &&
        nextSeq.items2 &&
        nextSeq.items2.length === seq.items1.length
      ) {
        // This is a modification - compare at word level
        for (let j = 0; j < seq.items1.length; j++) {
          const para1 = seq.items1[j];
          const para2 = nextSeq.items2[j];

          if (para1.text === para2.text) {
            // No actual change
            continue;
          }

          // For table rows, compare each cell paragraph separately
          if (para1.isTableRow && para1.rowCells && para2.isTableRow && para2.rowCells) {
            const revs = compareTableRowContent(para1.rowCells, para2.rowCells);
            insertions += revs.insertions;
            deletions += revs.deletions;
          } else {
            // Compare at word level for regular paragraphs
            const wordRevs = countWordRevisions(para1.text, para2.text);
            insertions += wordRevs.insertions;
            deletions += wordRevs.deletions;
          }
        }
        // Skip the next sequence since we processed it
        i++;
      } else {
        // True deletion - count each deleted paragraph as a revision
        // The C# comparer counts individual paragraph changes
        deletions += seq.items1.length;
      }
    } else if (seq.status === CorrelationStatus.Inserted && seq.items2) {
      // True insertion (not preceded by matching deletion)
      // Count each inserted paragraph as a revision
      insertions += seq.items2.length;
    } else if (seq.status === CorrelationStatus.Equal && seq.items1 && seq.items2) {
      // Paragraphs matched at hash level - check for word differences
      for (let j = 0; j < seq.items1.length; j++) {
        const para1 = seq.items1[j];
        const para2 = seq.items2[j];

        if (para1.text !== para2.text) {
          // For table rows, compare each cell paragraph separately
          if (para1.isTableRow && para1.rowCells && para2.isTableRow && para2.rowCells) {
            const revs = compareTableRowContent(para1.rowCells, para2.rowCells);
            insertions += revs.insertions;
            deletions += revs.deletions;
          } else {
            const wordRevs = countWordRevisions(para1.text, para2.text);
            insertions += wordRevs.insertions;
            deletions += wordRevs.deletions;
          }
        }
      }
    }
  }

  return { insertions, deletions, total: insertions + deletions };
}

/**
 * Compare content within table row cells.
 * For each cell, changes are grouped:
 * - All deletions within a cell = 1 revision
 * - All insertions within a cell = 1 revision
 * This allows multi-cell rows to count changes separately per cell.
 */
function compareTableRowContent(
  cells1: XmlNode[][],
  cells2: XmlNode[][]
): { insertions: number; deletions: number } {
  let insertions = 0;
  let deletions = 0;

  // Compare corresponding cells
  const maxCells = Math.max(cells1.length, cells2.length);
  for (let i = 0; i < maxCells; i++) {
    const cell1Paras = cells1[i] || [];
    const cell2Paras = cells2[i] || [];

    // Combine all text from each cell for comparison
    const text1 = cell1Paras.map((p) => extractParagraphText(p)).join(' ');
    const text2 = cell2Paras.map((p) => extractParagraphText(p)).join(' ');

    if (text1 === text2) {
      continue;
    }

    if (!text1.trim() && text2.trim()) {
      // Cell was empty, now has content
      insertions++;
    } else if (text1.trim() && !text2.trim()) {
      // Cell had content, now empty
      deletions++;
    } else {
      // Cell content changed - compare at word level and group by type
      const words1 = tokenize(text1);
      const words2 = tokenize(text2);
      const wordCorr = computeCorrelation(words1, words2);

      let hasInsertions = false;
      let hasDeletions = false;

      for (const wseq of wordCorr) {
        if (wseq.status === CorrelationStatus.Deleted) {
          hasDeletions = true;
        } else if (wseq.status === CorrelationStatus.Inserted) {
          hasInsertions = true;
        }
      }

      if (hasDeletions) deletions++;
      if (hasInsertions) insertions++;
    }
  }

  return { insertions, deletions };
}

/**
 * Count revisions at word level within a paragraph.
 * Adjacent changes of the same type are merged.
 * If there's only one type of change (pure insert or pure delete), count as 1.
 */
function countWordRevisions(
  text1: string,
  text2: string
): { insertions: number; deletions: number } {
  const words1 = tokenize(text1);
  const words2 = tokenize(text2);
  const wordCorr = computeCorrelation(words1, words2);

  let insertions = 0;
  let deletions = 0;
  let hasInsertions = false;
  let hasDeletions = false;
  let lastStatus: CorrelationStatus | null = null;

  for (const wseq of wordCorr) {
    if (wseq.status === CorrelationStatus.Deleted) {
      hasDeletions = true;
      if (lastStatus !== CorrelationStatus.Deleted) {
        deletions++;
      }
      lastStatus = CorrelationStatus.Deleted;
    } else if (wseq.status === CorrelationStatus.Inserted) {
      hasInsertions = true;
      if (lastStatus !== CorrelationStatus.Inserted) {
        insertions++;
      }
      lastStatus = CorrelationStatus.Inserted;
    } else {
      lastStatus = null;
    }
  }

  // If there's only one type of change (pure insert or pure delete),
  // treat as a single revision regardless of fragmentation
  if (hasInsertions && !hasDeletions) {
    return { insertions: 1, deletions: 0 };
  }
  if (hasDeletions && !hasInsertions) {
    return { insertions: 0, deletions: 1 };
  }

  return { insertions, deletions };
}
