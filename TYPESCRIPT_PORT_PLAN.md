# TypeScript Port Plan: Document Comparison Engines

## Executive Summary

This document provides a comprehensive plan for porting three document comparison engines (Word, Excel, PowerPoint) from C# to TypeScript, incorporating the best ideas from both the existing C# implementation and the JavaScript prototype in `redline-js/`.

**Current State:**
- C# implementation: ~14,500 lines across 3 comparers (WmlComparer: 8,835, SmlComparer: 2,982, PmlComparer: 2,708)
- JavaScript prototype: ~1,200 lines (Word-only, limited features)

**Target Benefits:**
- Browser-based processing (no server required)
- Cross-platform desktop apps (Electron/Tauri)
- Unified TypeScript codebase
- Modern async/streaming architecture

---

## Part 1: Algorithm Analysis Summary

### 1.1 WmlComparer (Word Documents) - 8,835 lines

**Core Algorithm: LCS-Based Paragraph Matching**

```
┌─────────────────────────────────────────────────────────────────────┐
│                        WML COMPARISON PIPELINE                       │
├─────────────────────────────────────────────────────────────────────┤
│  1. PREPROCESSING                                                    │
│     ├── Accept/reject existing tracked changes                       │
│     ├── Simplify markup (remove bookmarks, comments, etc.)          │
│     ├── Add unique IDs (Unid) to every element                      │
│     └── Compute correlation hashes for paragraphs/tables             │
│                                                                      │
│  2. ATOMIZATION                                                      │
│     ├── Decompose into ComparisonUnitAtoms (single chars/elements)  │
│     └── Preserve ancestor chain for each atom                        │
│                                                                      │
│  3. HIERARCHICAL GROUPING                                            │
│     ├── Atoms → Words (by word separators)                          │
│     ├── Words → Paragraphs                                           │
│     ├── Paragraphs → Cells → Rows → Tables                          │
│     └── Special: Textboxes as separate containers                    │
│                                                                      │
│  4. LCS MATCHING (with optimizations)                                │
│     ├── Hash-based paragraph correlation (pre-computed)              │
│     ├── Edge trimming (common prefix/suffix)                         │
│     ├── SHA1-based longest matching sequence                         │
│     └── Detail threshold filtering (ignore <15% matches)             │
│                                                                      │
│  5. TREE RECONSTRUCTION                                              │
│     ├── Group atoms by ancestor Unids                                │
│     ├── Wrap deleted content in <w:del>                              │
│     ├── Wrap inserted content in <w:ins>                             │
│     ├── Add <w:rPrChange> for formatting changes                     │
│     └── Coalesce adjacent runs with identical formatting             │
│                                                                      │
│  6. POST-PROCESSING                                                  │
│     ├── Fix footnote/endnote IDs and compare content                 │
│     ├── Merge adjacent paragraph marks                               │
│     ├── Renumber revision IDs                                        │
│     └── Fix drawing/shape IDs                                        │
└─────────────────────────────────────────────────────────────────────┘
```

**Key Data Structures:**
- `ComparisonUnitAtom`: Single character/element with full ancestor chain
- `ComparisonUnitWord`: Collection of atoms forming a word
- `ComparisonUnitGroup`: Paragraph/Table/Row/Cell/Textbox
- `CorrelatedSequence`: Matched/unmatched sequences with status

**Edge Cases Handled:**
- Tracked changes (existing insertions/deletions)
- Footnotes & endnotes (separate recursive comparison)
- Tables (merged cells, row insertion/deletion)
- Images & drawings (relationship-based comparison)
- Textboxes (VML content comparison)
- Math equations (atomic treatment)
- Formatting changes (optional tracking)

---

### 1.2 SmlComparer (Excel Spreadsheets) - 2,982 lines

**Core Algorithm: Sheet-Level Matching + LCS Row Alignment**

```
┌─────────────────────────────────────────────────────────────────────┐
│                        SML COMPARISON PIPELINE                       │
├─────────────────────────────────────────────────────────────────────┤
│  1. CANONICALIZATION                                                 │
│     ├── Resolve shared strings to actual values                      │
│     ├── Expand style indices to full format signatures               │
│     ├── Extract all cells with row/column indices                    │
│     ├── Extract Phase 3 features (comments, validations, etc.)       │
│     └── Compute row/column signatures for alignment                  │
│                                                                      │
│  2. SHEET MATCHING                                                   │
│     ├── Pass 1: Exact name matches                                   │
│     ├── Pass 2: Content hash matches (detect renames)                │
│     ├── Pass 3: Jaccard similarity for fuzzy rename detection       │
│     └── Remaining: Mark as added/deleted                             │
│                                                                      │
│  3. ROW ALIGNMENT (optional)                                         │
│     ├── Compute row signatures (sampled cell values)                 │
│     ├── LCS algorithm on row signature sequences                     │
│     └── Build alignment map: (oldRow, newRow)                        │
│                                                                      │
│  4. CELL COMPARISON                                                  │
│     ├── Compare values (with tolerance for numbers)                  │
│     ├── Compare formulas (exact string match)                        │
│     └── Compare formatting (24 properties)                           │
│                                                                      │
│  5. PHASE 3 COMPARISONS                                              │
│     ├── Named ranges (workbook level)                                │
│     ├── Comments (cell address → text/author)                        │
│     ├── Data validations (type, formulas, options)                   │
│     ├── Merged cells (region sets)                                   │
│     └── Hyperlinks (target, display text)                            │
│                                                                      │
│  6. MARKUP RENDERING                                                 │
│     ├── Add highlight fill styles to stylesheet                      │
│     ├── Apply style indices to changed cells                         │
│     ├── Add cell comments with change descriptions                   │
│     └── Create _DiffSummary sheet                                    │
└─────────────────────────────────────────────────────────────────────┘
```

**Key Data Structures:**
- `WorkbookSignature`: Collection of sheets + named ranges
- `WorksheetSignature`: Cells + comments + validations + hyperlinks
- `CellSignature`: Address, value, formula, format
- `CellFormatSignature`: 24 formatting properties
- `SmlChange`: Change record with all properties

**Edge Cases Handled:**
- Shared strings (resolve before comparison)
- Number formats (built-in codes 0-49 + custom)
- Indexed colors (64-color palette)
- Theme colors (reference-based)
- Empty cells & inline strings
- Multi-cell data validation ranges
- VML drawing parts (for comments display)

---

### 1.3 PmlComparer (PowerPoint Presentations) - 2,708 lines

**Core Algorithm: Slide Matching + Shape Matching**

```
┌─────────────────────────────────────────────────────────────────────┐
│                        PML COMPARISON PIPELINE                       │
├─────────────────────────────────────────────────────────────────────┤
│  1. CANONICALIZATION                                                 │
│     ├── Extract slide signatures (title, shapes, layout)            │
│     ├── Extract shape signatures (type, transform, content)         │
│     ├── Compute text body signatures (paragraphs, runs)             │
│     ├── Compute content hashes (images, tables, charts)             │
│     └── Build fingerprints for slide matching                        │
│                                                                      │
│  2. SLIDE MATCHING                                                   │
│     ├── Pass 1: Exact title text match                               │
│     ├── Pass 2: Fingerprint (content hash) match                     │
│     ├── Pass 3: LCS with weighted similarity scoring                 │
│     │     ├── Title weight: 3                                        │
│     │     ├── Shape types overlap: 1                                 │
│     │     ├── Shape names overlap: 2                                 │
│     │     └── Threshold: 0.4                                         │
│     └── Remaining: Mark as inserted/deleted                          │
│                                                                      │
│  3. SHAPE MATCHING (within matched slides)                           │
│     ├── Pass 1: Placeholder match (type + index)                     │
│     ├── Pass 2: Name + Type match                                    │
│     ├── Pass 3: Name only match                                      │
│     ├── Pass 4: Fuzzy match (position + content similarity)         │
│     │     ├── Same type required                                     │
│     │     ├── Position tolerance: 91,440 EMUs                        │
│     │     ├── Content similarity (text/image/hash)                   │
│     │     └── Threshold: 0.7                                         │
│     └── Remaining: Mark as inserted/deleted                          │
│                                                                      │
│  4. DIFF COMPUTATION                                                 │
│     ├── Slide-level: layout, background, notes, transitions          │
│     ├── Shape-level: position, size, rotation, z-order               │
│     ├── Content: text (plain + formatting), images, tables, charts   │
│     └── Style: fill, line, effects                                   │
│                                                                      │
│  5. MARKUP RENDERING                                                 │
│     ├── Create change overlay labels                                 │
│     ├── Add notes annotations with change list                       │
│     └── Create summary slide                                         │
└─────────────────────────────────────────────────────────────────────┘
```

**Key Data Structures:**
- `PresentationSignature`: Slides + theme + dimensions
- `SlideSignature`: Shapes + title + layout + fingerprint
- `ShapeSignature`: Transform + text body + content hash
- `TransformSignature`: X, Y, Cx, Cy, Rotation, Flip
- `TextBodySignature`: Paragraphs → Runs → Properties

**Edge Cases Handled:**
- Placeholder matching (title, body, subtitle, etc.)
- Shape type detection from element names
- Group shapes (recursive children)
- Charts (full XML hash comparison)
- Tables (cell text concatenation)
- Images (SHA256 of binary content)
- Position tolerance (EMU-based)
- Text similarity (Levenshtein distance)

---

### 1.4 redline-js Analysis (JavaScript Prototype)

**Valuable Ideas to Incorporate:**

| Idea | Description | Benefit |
|------|-------------|---------|
| **Dual Algorithm Strategy** | Myers for large docs, DP for small | Balance speed vs accuracy |
| **Diagonal Score Tracking** | Modified Myers tracks consecutive diagonals | Cleaner diffs |
| **Word-length Scoring** | DP scores by `1 + log(1 + wordLength)` | Semantic awareness |
| **Run Abstraction** | Treat Word runs as indexed strings | Efficient splitting |
| **Two-Phase Processing** | Mark additions first, then insert deletions | Simpler state |
| **Linked List Runs** | O(1) insertions between runs | Fast modifications |
| **Cursor-Based Traversal** | Sequential processing matching diff order | Natural flow |
| **Paragraph Boundary Tracking** | Aware of paragraph splits during insertion | Correct structure |

**Current Limitations:**
- Word documents only (no PPTX/XLSX)
- No headers/footers/footnotes
- No complex nested structures
- Hardcoded 2000 token threshold
- No existing tracked changes handling

---

## Part 2: Unified TypeScript Architecture

### 2.1 Package Structure

```
@openxml-compare/
├── core/                      # Shared infrastructure
│   ├── diff/                  # Diff algorithms
│   │   ├── myers.ts           # Myers O(ND) algorithm
│   │   ├── lcs.ts             # LCS dynamic programming
│   │   ├── semantic.ts        # Semantic cleanup
│   │   └── index.ts
│   ├── xml/                   # XML utilities
│   │   ├── parser.ts          # XML parsing wrapper
│   │   ├── builder.ts         # XML construction
│   │   ├── namespaces.ts      # OOXML namespace constants
│   │   └── traverse.ts        # Tree traversal helpers
│   ├── ooxml/                 # OOXML package handling
│   │   ├── package.ts         # ZIP file handling
│   │   ├── parts.ts           # Part management
│   │   ├── relationships.ts   # Relationship management
│   │   └── content-types.ts   # Content types
│   ├── hash/                  # Hashing utilities
│   │   ├── sha256.ts          # SHA256 implementation
│   │   ├── quick-hash.ts      # Fast 32-bit hash
│   │   └── similarity.ts      # Jaccard, Levenshtein
│   └── types/                 # Shared type definitions
│       ├── change.ts          # Change record types
│       ├── signature.ts       # Signature types
│       └── settings.ts        # Settings types
│
├── word/                      # Word document comparison
│   ├── canonicalize.ts        # Document normalization
│   ├── atomize.ts             # Decomposition to atoms
│   ├── group.ts               # Hierarchical grouping
│   ├── match.ts               # LCS matching
│   ├── reconstruct.ts         # XML tree reconstruction
│   ├── markup.ts              # Tracked changes markup
│   ├── footnotes.ts           # Footnote/endnote handling
│   ├── tables.ts              # Table comparison
│   ├── textboxes.ts           # Textbox handling
│   ├── WmlComparer.ts         # Main entry point
│   └── types.ts               # Word-specific types
│
├── excel/                     # Excel spreadsheet comparison
│   ├── canonicalize.ts        # Workbook normalization
│   ├── styles.ts              # Style expansion
│   ├── sheets.ts              # Sheet matching
│   ├── rows.ts                # Row alignment (LCS)
│   ├── cells.ts               # Cell comparison
│   ├── features.ts            # Comments, validations, etc.
│   ├── markup.ts              # Highlight rendering
│   ├── SmlComparer.ts         # Main entry point
│   └── types.ts               # Excel-specific types
│
├── powerpoint/                # PowerPoint comparison
│   ├── canonicalize.ts        # Presentation normalization
│   ├── slides.ts              # Slide matching
│   ├── shapes.ts              # Shape matching
│   ├── text.ts                # Text comparison
│   ├── transforms.ts          # Transform comparison
│   ├── content.ts             # Images, tables, charts
│   ├── markup.ts              # Annotation rendering
│   ├── PmlComparer.ts         # Main entry point
│   └── types.ts               # PowerPoint-specific types
│
└── index.ts                   # Public API exports
```

### 2.2 Shared Diff Engine

```typescript
// core/diff/types.ts
export enum DiffOperation {
  EQUAL = 0,
  DELETE = -1,
  INSERT = 1
}

export interface Diff {
  op: DiffOperation;
  text: string;
}

export interface DiffOptions {
  timeout?: number;           // Max time in ms
  checkLines?: boolean;       // Line-level preprocessing
  semanticThreshold?: number; // Cleanup threshold (0 = no cleanup)
  wordMode?: boolean;         // Tokenize by words
}

// core/diff/myers.ts
export function diffMyers(
  text1: string,
  text2: string,
  options?: DiffOptions
): Diff[];

// core/diff/lcs.ts
export function diffDP(
  text1: string,
  text2: string,
  equalityScore?: (i: number, j: number) => number
): Diff[];

// core/diff/index.ts
export function diff(
  text1: string,
  text2: string,
  options?: DiffOptions
): Diff[] {
  const numTokens = text1.length + text2.length;
  const threshold = options?.dpThreshold ?? 2000;

  if (numTokens > threshold) {
    return diffMyers(text1, text2, options);
  } else {
    return diffDP(text1, text2);
  }
}
```

### 2.3 XML Processing Layer

```typescript
// core/xml/namespaces.ts
export const NS = {
  // WordprocessingML
  W: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  WP: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',

  // SpreadsheetML
  S: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',

  // PresentationML
  P: 'http://schemas.openxmlformats.org/presentationml/2006/main',

  // DrawingML
  A: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  C: 'http://schemas.openxmlformats.org/drawingml/2006/chart',

  // Common
  R: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  MC: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  VML: 'urn:schemas-microsoft-com:vml',

  // Custom (for comparison)
  PT14: 'http://powertools.codeplex.com/2011'
} as const;

// core/xml/parser.ts
export interface XmlElement {
  name: string;
  prefix?: string;
  ns?: string;
  attributes: Record<string, string>;
  children: XmlNode[];
  parent?: XmlElement;
  text?: string;
}

export type XmlNode = XmlElement | string;

export interface XmlDocument {
  root: XmlElement;
  declaration?: { version: string; encoding: string };
}

export function parseXml(content: string): XmlDocument;
export function serializeXml(doc: XmlDocument): string;
```

### 2.4 OOXML Package Layer

```typescript
// core/ooxml/package.ts
export interface OoxmlPackage {
  type: 'word' | 'excel' | 'powerpoint';
  parts: Map<string, OoxmlPart>;
  relationships: Map<string, Relationship[]>;
  contentTypes: ContentType[];
}

export interface OoxmlPart {
  path: string;
  contentType: string;
  content: Uint8Array | XmlDocument;
}

export interface Relationship {
  id: string;
  type: string;
  target: string;
  targetMode?: 'External' | 'Internal';
}

export async function openPackage(data: ArrayBuffer): Promise<OoxmlPackage>;
export async function savePackage(pkg: OoxmlPackage): Promise<ArrayBuffer>;
export function getPart(pkg: OoxmlPackage, path: string): OoxmlPart | undefined;
export function getPartByRelId(pkg: OoxmlPackage, source: string, rId: string): OoxmlPart | undefined;
```

---

## Part 3: Implementation Details by Comparer

### 3.1 WmlComparer Implementation

**Phase 1: Preprocessing**

```typescript
// word/canonicalize.ts
export interface PreprocessOptions {
  acceptRevisions: boolean;
  removeBookmarks: boolean;
  removeComments: boolean;
  removeContentControls: boolean;
  removeHyperlinks: boolean;
}

export function preprocessDocument(
  doc: XmlDocument,
  options: PreprocessOptions
): XmlDocument {
  let result = doc;

  // 1. Handle markup compatibility (mc:AlternateContent)
  result = flattenAlternateContent(result, { preferVml: true });

  // 2. Accept or reject existing revisions
  if (options.acceptRevisions) {
    result = acceptRevisions(result);
  }

  // 3. Simplify markup
  result = simplifyMarkup(result, {
    removeBookmarks: options.removeBookmarks,
    removeComments: options.removeComments,
    // ... etc
  });

  // 4. Add unique identifiers
  result = addUnids(result);

  return result;
}

// Add unique IDs to every element
function addUnids(doc: XmlDocument): XmlDocument {
  let counter = 0;

  function addToElement(el: XmlElement): void {
    el.attributes[`${NS.PT14}:Unid`] = String(counter++);
    for (const child of el.children) {
      if (typeof child !== 'string') {
        addToElement(child);
      }
    }
  }

  addToElement(doc.root);
  return doc;
}
```

**Phase 2: Atomization**

```typescript
// word/atomize.ts
export interface ComparisonUnitAtom {
  contentElement: XmlElement;      // Single char or element
  ancestorElements: XmlElement[];  // Path from body to element
  ancestorUnids: string[];         // Unids for reconstruction
  sha1Hash: string;                // Content hash
  status: CorrelationStatus;
  formattingSignature?: string;    // For format change detection
  atomBefore?: ComparisonUnitAtom; // Reference for Equal atoms
}

export function createAtomList(
  doc: XmlDocument,
  part: OoxmlPart
): ComparisonUnitAtom[] {
  const atoms: ComparisonUnitAtom[] = [];
  const body = findElement(doc.root, 'w:body');

  function processElement(el: XmlElement, ancestors: XmlElement[]): void {
    if (el.name === 'w:t') {
      // Split text into individual characters
      const text = getTextContent(el);
      for (const char of text) {
        atoms.push({
          contentElement: createTextElement(char),
          ancestorElements: [...ancestors, el],
          ancestorUnids: [],
          sha1Hash: computeHash(char),
          status: CorrelationStatus.Unknown
        });
      }
    } else if (isAtomicElement(el)) {
      // Treat as single atom (br, drawing, pPr, etc.)
      atoms.push({
        contentElement: el,
        ancestorElements: ancestors,
        ancestorUnids: [],
        sha1Hash: computeHash(serializeElement(el)),
        status: CorrelationStatus.Unknown
      });
    } else {
      // Recurse into children
      for (const child of el.children) {
        if (typeof child !== 'string') {
          processElement(child, [...ancestors, el]);
        }
      }
    }
  }

  processElement(body, []);
  return atoms;
}

const ATOMIC_ELEMENTS = new Set([
  'w:pPr', 'w:rPr', 'w:br', 'w:tab', 'w:sym', 'w:cr',
  'w:drawing', 'w:pict', 'w:object', 'w:fldChar'
]);

function isAtomicElement(el: XmlElement): boolean {
  return ATOMIC_ELEMENTS.has(el.name);
}
```

**Phase 3: Grouping**

```typescript
// word/group.ts
export type ComparisonUnit =
  | ComparisonUnitAtom
  | ComparisonUnitWord
  | ComparisonUnitGroup;

export interface ComparisonUnitWord {
  type: 'word';
  atoms: ComparisonUnitAtom[];
  sha1Hash: string;
  status: CorrelationStatus;
}

export interface ComparisonUnitGroup {
  type: 'group';
  groupType: 'paragraph' | 'table' | 'row' | 'cell' | 'textbox';
  contents: ComparisonUnit[];
  sha1Hash: string;
  correlatedHash?: string; // From preprocessing
  status: CorrelationStatus;
}

const WORD_SEPARATORS = new Set([
  ' ', '\t', '\n', '\r', '.', ',', ';', ':', '!', '?',
  '(', ')', '[', ']', '{', '}', '"', "'", '-', '/'
]);

export function groupIntoUnits(
  atoms: ComparisonUnitAtom[]
): ComparisonUnit[] {
  // Step 1: Group into words
  const words = groupIntoWords(atoms);

  // Step 2: Group into paragraphs
  const paragraphs = groupIntoParagraphs(words);

  // Step 3: Group into tables/cells/rows
  const structured = groupIntoTables(paragraphs);

  // Step 4: Handle textboxes
  const final = groupIntoTextboxes(structured);

  return final;
}

function groupIntoWords(atoms: ComparisonUnitAtom[]): ComparisonUnit[] {
  const result: ComparisonUnit[] = [];
  let currentWord: ComparisonUnitAtom[] = [];

  for (const atom of atoms) {
    const isTextAtom = atom.contentElement.name === 'w:t';
    const text = isTextAtom ? getTextContent(atom.contentElement) : null;

    if (text && WORD_SEPARATORS.has(text)) {
      // End current word
      if (currentWord.length > 0) {
        result.push(createWord(currentWord));
        currentWord = [];
      }
      // Separator is its own word
      result.push(createWord([atom]));
    } else {
      currentWord.push(atom);
    }
  }

  if (currentWord.length > 0) {
    result.push(createWord(currentWord));
  }

  return result;
}
```

**Phase 4: LCS Matching**

```typescript
// word/match.ts
export interface CorrelatedSequence {
  status: CorrelationStatus;
  units1: ComparisonUnit[];
  units2: ComparisonUnit[];
}

export function matchUnits(
  units1: ComparisonUnit[],
  units2: ComparisonUnit[],
  settings: WmlComparerSettings
): CorrelatedSequence[] {
  const sequences: CorrelatedSequence[] = [{
    status: CorrelationStatus.Unknown,
    units1,
    units2
  }];

  let changed = true;
  while (changed) {
    changed = false;

    for (let i = 0; i < sequences.length; i++) {
      const seq = sequences[i];
      if (seq.status !== CorrelationStatus.Unknown) continue;

      // Try correlation hash matching
      let result = matchByCorrelationHash(seq);
      if (result) {
        sequences.splice(i, 1, ...result);
        changed = true;
        break;
      }

      // Try edge trimming
      result = trimCommonEdges(seq);
      if (result) {
        sequences.splice(i, 1, ...result);
        changed = true;
        break;
      }

      // LCS algorithm
      result = runLcsAlgorithm(seq, settings);
      if (result) {
        sequences.splice(i, 1, ...result);
        changed = true;
        break;
      }
    }
  }

  return sequences;
}

function runLcsAlgorithm(
  seq: CorrelatedSequence,
  settings: WmlComparerSettings
): CorrelatedSequence[] | null {
  const { units1, units2 } = seq;

  // Find longest matching sequence using SHA1 hashes
  let bestMatch: { start1: number; start2: number; length: number } | null = null;

  for (let i = 0; i < units1.length; i++) {
    const hash1 = units1[i].sha1Hash;

    for (let j = 0; j < units2.length; j++) {
      if (units2[j].sha1Hash !== hash1) continue;

      // Found potential match start, extend it
      let length = 1;
      while (
        i + length < units1.length &&
        j + length < units2.length &&
        units1[i + length].sha1Hash === units2[j + length].sha1Hash
      ) {
        length++;
      }

      // Check detail threshold
      const matchRatio = length / Math.max(units1.length, units2.length);
      if (matchRatio < settings.detailThreshold) continue;

      if (!bestMatch || length > bestMatch.length) {
        bestMatch = { start1: i, start2: j, length };
      }
    }
  }

  if (!bestMatch) {
    // No match found - mark as deleted + inserted
    return [
      { status: CorrelationStatus.Deleted, units1, units2: [] },
      { status: CorrelationStatus.Inserted, units1: [], units2 }
    ];
  }

  // Split into before, match, after
  const result: CorrelatedSequence[] = [];

  // Before match
  if (bestMatch.start1 > 0 || bestMatch.start2 > 0) {
    result.push({
      status: CorrelationStatus.Unknown,
      units1: units1.slice(0, bestMatch.start1),
      units2: units2.slice(0, bestMatch.start2)
    });
  }

  // The match
  result.push({
    status: CorrelationStatus.Equal,
    units1: units1.slice(bestMatch.start1, bestMatch.start1 + bestMatch.length),
    units2: units2.slice(bestMatch.start2, bestMatch.start2 + bestMatch.length)
  });

  // After match
  if (bestMatch.start1 + bestMatch.length < units1.length ||
      bestMatch.start2 + bestMatch.length < units2.length) {
    result.push({
      status: CorrelationStatus.Unknown,
      units1: units1.slice(bestMatch.start1 + bestMatch.length),
      units2: units2.slice(bestMatch.start2 + bestMatch.length)
    });
  }

  return result;
}
```

**Phase 5: Tree Reconstruction**

```typescript
// word/reconstruct.ts
export function reconstructDocument(
  atoms: ComparisonUnitAtom[],
  settings: WmlComparerSettings
): XmlDocument {
  // Group atoms by ancestor path
  const groups = groupByAncestorPath(atoms, 0);

  // Build tree recursively
  const body = buildElement('w:body', {},
    groups.map(g => coalesceGroup(g, 0, settings))
  );

  return createDocument(body);
}

function coalesceGroup(
  atoms: ComparisonUnitAtom[],
  level: number,
  settings: WmlComparerSettings
): XmlNode[] {
  if (atoms.length === 0) return [];

  // Group by: ancestorUnid | status | formatting
  const grouped = groupAdjacent(atoms, a => {
    const unid = a.ancestorUnids[level] ?? '';
    const status = a.status;
    const fmt = a.formattingSignature ?? '';
    return `${unid}|${status}|${fmt}`;
  });

  const result: XmlNode[] = [];

  for (const group of grouped) {
    const first = group[0];
    const status = first.status;

    if (level >= first.ancestorElements.length - 1) {
      // Leaf level - emit content with revision markup
      result.push(...wrapWithRevisionMarkup(group, status, settings));
    } else {
      // Intermediate level - recurse
      const template = first.ancestorElements[level];
      const children = coalesceGroup(group, level + 1, settings);
      const element = cloneElement(template, children);

      if (status === CorrelationStatus.Deleted && isRowElement(template)) {
        // Table row deletion - add marker to trPr
        addRowDeletionMarker(element, settings);
      }

      result.push(element);
    }
  }

  return result;
}

function wrapWithRevisionMarkup(
  atoms: ComparisonUnitAtom[],
  status: CorrelationStatus,
  settings: WmlComparerSettings
): XmlNode[] {
  switch (status) {
    case CorrelationStatus.Deleted:
      return [createDelElement(atoms, settings)];

    case CorrelationStatus.Inserted:
      return [createInsElement(atoms, settings)];

    case CorrelationStatus.FormatChanged:
      return [createRunWithRPrChange(atoms, settings)];

    case CorrelationStatus.Equal:
    default:
      return atoms.map(a => a.contentElement);
  }
}
```

---

### 3.2 SmlComparer Implementation

**Key differences from C# version:**

1. **Use JSZip for package handling** instead of OpenXML SDK
2. **Use fast-xml-parser** for XML parsing (faster than xml2js)
3. **WebCrypto API** for SHA256 in browser

```typescript
// excel/SmlComparer.ts
export interface SmlComparerSettings {
  compareValues: boolean;
  compareFormulas: boolean;
  compareFormatting: boolean;
  compareSheetStructure: boolean;
  caseInsensitiveValues: boolean;
  numericTolerance: number;
  enableRowAlignment: boolean;
  enableSheetRenameDetection: boolean;
  sheetRenameSimilarityThreshold: number;
  // Phase 3
  compareNamedRanges: boolean;
  compareComments: boolean;
  compareDataValidation: boolean;
  compareMergedCells: boolean;
  compareHyperlinks: boolean;
  // Output
  authorForChanges: string;
  highlightColors: HighlightColors;
}

export async function compare(
  older: ArrayBuffer,
  newer: ArrayBuffer,
  settings: SmlComparerSettings
): Promise<SmlComparisonResult> {
  // 1. Open packages
  const pkg1 = await openPackage(older);
  const pkg2 = await openPackage(newer);

  // 2. Canonicalize
  const sig1 = await canonicalize(pkg1, settings);
  const sig2 = await canonicalize(pkg2, settings);

  // 3. Match sheets
  const sheetMatches = matchSheets(sig1, sig2, settings);

  // 4. Compare matched sheets
  const result = new SmlComparisonResult();

  for (const match of sheetMatches) {
    if (match.type === 'added') {
      result.addChange({ type: SmlChangeType.SheetAdded, sheetName: match.newName });
    } else if (match.type === 'deleted') {
      result.addChange({ type: SmlChangeType.SheetDeleted, sheetName: match.oldName });
    } else if (match.type === 'renamed') {
      result.addChange({
        type: SmlChangeType.SheetRenamed,
        oldSheetName: match.oldName,
        sheetName: match.newName
      });
      compareSheets(sig1.sheets[match.oldName], sig2.sheets[match.newName], match.newName, settings, result);
    } else {
      compareSheets(sig1.sheets[match.oldName], sig2.sheets[match.newName], match.name, settings, result);
    }
  }

  // 5. Compare named ranges
  if (settings.compareNamedRanges) {
    compareNamedRanges(sig1.definedNames, sig2.definedNames, result);
  }

  return result;
}
```

---

### 3.3 PmlComparer Implementation

```typescript
// powerpoint/PmlComparer.ts
export interface PmlComparerSettings {
  compareSlideStructure: boolean;
  compareShapeStructure: boolean;
  compareTextContent: boolean;
  compareTextFormatting: boolean;
  compareShapeTransforms: boolean;
  compareImageContent: boolean;
  compareTables: boolean;
  compareCharts: boolean;
  // Matching
  enableFuzzyShapeMatching: boolean;
  slideSimilarityThreshold: number;
  shapeSimilarityThreshold: number;
  positionTolerance: number; // EMUs
  // Output
  authorForChanges: string;
  addSummarySlide: boolean;
  colors: ChangeColors;
}

export async function compare(
  older: ArrayBuffer,
  newer: ArrayBuffer,
  settings: PmlComparerSettings
): Promise<PmlComparisonResult> {
  // 1. Open packages
  const pkg1 = await openPackage(older);
  const pkg2 = await openPackage(newer);

  // 2. Canonicalize
  const sig1 = await canonicalize(pkg1, settings);
  const sig2 = await canonicalize(pkg2, settings);

  // 3. Match slides
  const slideMatches = matchSlides(sig1, sig2, settings);

  // 4. Compare
  const result = new PmlComparisonResult();

  for (const match of slideMatches) {
    switch (match.type) {
      case 'inserted':
        result.addChange({ type: PmlChangeType.SlideInserted, slideIndex: match.newIndex });
        break;
      case 'deleted':
        result.addChange({ type: PmlChangeType.SlideDeleted, oldSlideIndex: match.oldIndex });
        break;
      case 'matched':
        if (match.wasMoved) {
          result.addChange({
            type: PmlChangeType.SlideMoved,
            oldSlideIndex: match.oldIndex,
            slideIndex: match.newIndex
          });
        }
        compareSlideContents(match.oldSlide, match.newSlide, match.newIndex, settings, result);
        break;
    }
  }

  return result;
}

// powerpoint/slides.ts
export function matchSlides(
  sig1: PresentationSignature,
  sig2: PresentationSignature,
  settings: PmlComparerSettings
): SlideMatch[] {
  const matches: SlideMatch[] = [];
  const used1 = new Set<number>();
  const used2 = new Set<number>();

  // Pass 1: Title text match
  for (const slide1 of sig1.slides) {
    if (used1.has(slide1.index) || !slide1.titleText) continue;

    const match = sig2.slides.find(s =>
      !used2.has(s.index) && s.titleText === slide1.titleText
    );

    if (match) {
      matches.push({ type: 'matched', oldSlide: slide1, newSlide: match, similarity: 1.0 });
      used1.add(slide1.index);
      used2.add(match.index);
    }
  }

  // Pass 2: Fingerprint match
  for (const slide1 of sig1.slides) {
    if (used1.has(slide1.index)) continue;

    const fp1 = slide1.computeFingerprint();
    const match = sig2.slides.find(s =>
      !used2.has(s.index) && s.computeFingerprint() === fp1
    );

    if (match) {
      matches.push({ type: 'matched', oldSlide: slide1, newSlide: match, similarity: 1.0 });
      used1.add(slide1.index);
      used2.add(match.index);
    }
  }

  // Pass 3: Similarity-based matching
  const remaining1 = sig1.slides.filter(s => !used1.has(s.index));
  const remaining2 = sig2.slides.filter(s => !used2.has(s.index));

  for (const slide1 of remaining1) {
    let bestScore = 0;
    let bestMatch: SlideSignature | null = null;

    for (const slide2 of remaining2) {
      if (used2.has(slide2.index)) continue;

      const score = computeSlideSimilarity(slide1, slide2);
      if (score > bestScore && score >= settings.slideSimilarityThreshold) {
        bestScore = score;
        bestMatch = slide2;
      }
    }

    if (bestMatch) {
      matches.push({ type: 'matched', oldSlide: slide1, newSlide: bestMatch, similarity: bestScore });
      used1.add(slide1.index);
      used2.add(bestMatch.index);
    }
  }

  // Pass 4: Mark remaining as inserted/deleted
  for (const slide of sig1.slides.filter(s => !used1.has(s.index))) {
    matches.push({ type: 'deleted', oldSlide: slide });
  }
  for (const slide of sig2.slides.filter(s => !used2.has(s.index))) {
    matches.push({ type: 'inserted', newSlide: slide });
  }

  return matches.sort((a, b) =>
    (a.newSlide?.index ?? Infinity) - (b.newSlide?.index ?? Infinity)
  );
}
```

---

## Part 4: Testing Strategy (TDD Approach)

We use a **Test-Driven Development** approach, porting the ~200 existing C# tests first, then implementing the TypeScript code to make them pass. This ensures full compatibility with the C# implementation.

### 4.1 Golden File Generation

The C# `GoldenFileGenerator` project creates reference outputs:

```bash
# Generate golden files from C# implementation
cd redline-js && npm run generate-golden

# Or directly:
dotnet run --project GoldenFileGenerator
```

This creates:
- `redline-js/tests/golden/manifest.json` - Test case metadata
- `redline-js/tests/golden/wml/*.docx` - Word comparison outputs
- `redline-js/tests/golden/wml/*.document.xml` - Extracted document.xml for diffing
- `redline-js/tests/golden/pml/*.json` - PowerPoint comparison results

### 4.2 Test Directory Structure

```
redline-js/
├── tests/
│   ├── setup.ts                    # Test helpers and utilities
│   ├── wml-comparer.test.ts        # 100+ Word comparison tests
│   ├── sml-comparer.test.ts        # 50+ Excel comparison tests
│   ├── pml-comparer.test.ts        # 27+ PowerPoint comparison tests
│   ├── formatting-change.test.ts   # Formatting change tests
│   └── golden/
│       ├── manifest.json           # Test case metadata
│       ├── wml/                    # Word golden files
│       ├── sml/                    # Excel golden files
│       └── pml/                    # PowerPoint golden files
├── src/
│   ├── index.ts                    # Main exports
│   ├── types.ts                    # Type definitions
│   ├── wml/                        # Word comparison (to implement)
│   ├── sml/                        # Excel comparison (to implement)
│   └── pml/                        # PowerPoint comparison (to implement)
├── package.json
├── tsconfig.json
└── vitest.config.ts
```

### 4.3 Test Categories

Each test case validates three things:

1. **Revision Count**: Correct number of tracked changes detected
2. **Sanity Check 1**: `RejectRevisions(result)` equals original document
3. **Sanity Check 2**: `AcceptRevisions(result)` equals modified document

```typescript
describe.each([
  ['WC-1010', 'WC/WC001-Digits.docx', 'WC/WC001-Digits-Mod.docx', 4],
  ['WC-1040', 'WC/WC002-Unmodified.docx', 'WC/WC002-DiffInMiddle.docx', 2],
  // ... 100+ test cases from WmlComparerTests.cs
])('%s: %s vs %s', (testId, source1, source2, expectedRevisions) => {

  it(`produces ${expectedRevisions} revisions`, async () => {
    const doc1 = await loadDocument(source1);
    const doc2 = await loadDocument(source2);
    const result = await WmlComparer.compare(doc1, doc2);
    expect(result.revisions.length).toBe(expectedRevisions);
  });

  it('passes sanity check 1: reject produces original', async () => {
    const result = await WmlComparer.compare(doc1, doc2);
    const afterReject = await RevisionProcessor.rejectRevisions(result);
    const sanityCheck = await WmlComparer.compare(doc1, afterReject);
    expect(sanityCheck.revisions.length).toBe(0);
  });

  it('passes sanity check 2: accept produces modified', async () => {
    const result = await WmlComparer.compare(doc1, doc2);
    const afterAccept = await RevisionProcessor.acceptRevisions(result);
    const sanityCheck = await WmlComparer.compare(doc2, afterAccept);
    expect(sanityCheck.revisions.length).toBe(0);
  });

  it('matches golden file output', async () => {
    const result = await WmlComparer.compare(doc1, doc2);
    const goldenXml = await loadGoldenFile(`wml/${testId}.document.xml`);
    expect(extractDocumentXml(result)).toEqual(goldenXml);
  });
});
```

### 4.4 Test Execution

```bash
# Run all tests
npm test

# Run tests in watch mode
npm run test:watch

# Run with coverage
npm run test:coverage

# Run specific test file
npx vitest run tests/wml-comparer.test.ts
```

### 4.5 Implementation Order (TDD)

Implement features to make tests pass in this order:

| Phase | Test IDs | Feature | Tests |
|-------|----------|---------|-------|
| 1 | WC-1000 to WC-1030 | Basic text diff | 4 |
| 2 | WC-1040 to WC-1120 | Insert/delete at positions | 9 |
| 3 | WC-1140 to WC-1220 | Tables | 10 |
| 4 | WC-1230 to WC-1350 | Math, images, SmartArt | 15 |
| 5 | WC-1360 to WC-1430 | Fields, hyperlinks, math | 8 |
| 6 | WC-1400 to WC-1780 | Footnotes, endnotes, textboxes | 40+ |
| 7 | WC-1800+ | Edge cases, styles | 20+ |

### 4.6 Libraries Used

- **vitest**: Test runner (faster than Jest, native ESM support)
- **jszip**: ZIP extraction for fixture loading
- **fast-xml-parser**: XML parsing for assertions

---

## Part 5: Implementation Phases

### Phase 1: Core Infrastructure (2-3 weeks)

**Deliverables:**
- [ ] XML parsing/serialization with namespace support
- [ ] OOXML package handling (ZIP + parts + relationships)
- [ ] Diff algorithms (Myers + DP)
- [ ] Hashing utilities (SHA256, quick hash, similarity)
- [ ] Shared type definitions
- [ ] Basic test infrastructure

**Success Criteria:**
- Can open/read/modify/save .docx, .xlsx, .pptx files
- Diff algorithms produce correct output for text comparison
- All unit tests pass

### Phase 2: Word Comparer MVP (4-6 weeks)

**Deliverables:**
- [ ] Document preprocessing (accept revisions, simplify, add Unids)
- [ ] Atomization and hierarchical grouping
- [ ] LCS matching with hash-based optimization
- [ ] Tree reconstruction with revision markup
- [ ] Basic footnote handling
- [ ] Run coalescing

**Success Criteria:**
- Passes 50% of ported WmlComparer tests
- Output opens correctly in Microsoft Word
- Basic insertion/deletion tracking works

### Phase 3: Word Comparer Complete (3-4 weeks)

**Deliverables:**
- [ ] Table comparison (row insertion/deletion)
- [ ] Textbox handling
- [ ] Formatting change tracking
- [ ] Complete footnote/endnote support
- [ ] All post-processing (ID fixup, style sync)

**Success Criteria:**
- Passes 90%+ of ported WmlComparer tests
- Handles complex documents (tables, textboxes, footnotes)
- Output matches C# version for test fixtures

### Phase 4: Excel Comparer (3-4 weeks)

**Deliverables:**
- [ ] Workbook canonicalization (shared strings, styles)
- [ ] Sheet matching with rename detection
- [ ] Row alignment using LCS
- [ ] Cell comparison (values, formulas, formatting)
- [ ] Phase 3 features (comments, validations, etc.)
- [ ] Markup rendering (highlights, summary sheet)

**Success Criteria:**
- All documented features working
- Output opens correctly in Microsoft Excel
- Matches C# version behavior

### Phase 5: PowerPoint Comparer (3-4 weeks)

**Deliverables:**
- [ ] Presentation canonicalization
- [ ] Slide matching (multi-pass)
- [ ] Shape matching (placeholder, name, fuzzy)
- [ ] Content comparison (text, images, tables, charts)
- [ ] Transform comparison
- [ ] Markup rendering (overlays, notes, summary)

**Success Criteria:**
- All documented features working
- Output opens correctly in Microsoft PowerPoint
- Matches C# version behavior

### Phase 6: Browser Optimization (2-3 weeks)

**Deliverables:**
- [ ] Web Worker support for background processing
- [ ] Streaming for large documents
- [ ] Progress callbacks
- [ ] Memory optimization
- [ ] Bundle size optimization

**Success Criteria:**
- Works in modern browsers (Chrome, Firefox, Safari, Edge)
- Can handle 100+ page documents without crashes
- Responsive UI during comparison

### Phase 7: Polish & Documentation (2 weeks)

**Deliverables:**
- [ ] Complete API documentation
- [ ] Usage examples
- [ ] Migration guide from C# version
- [ ] Performance benchmarks
- [ ] npm package publishing

---

## Part 6: Dependencies

### Production Dependencies

```json
{
  "dependencies": {
    "jszip": "^3.10.1",
    "fast-xml-parser": "^4.3.0"
  }
}
```

### Development Dependencies

```json
{
  "devDependencies": {
    "typescript": "^5.3.0",
    "vitest": "^1.0.0",
    "@types/node": "^20.0.0",
    "esbuild": "^0.19.0",
    "prettier": "^3.0.0",
    "eslint": "^8.0.0"
  }
}
```

### Why These Choices?

| Library | Alternative | Reason |
|---------|-------------|--------|
| jszip | adm-zip | Better browser support, streaming |
| fast-xml-parser | xml2js | 10x faster, lower memory |
| vitest | jest | Faster, native TypeScript |
| esbuild | webpack | Faster builds, simpler config |

---

## Part 7: Risk Assessment

### High Risk

| Risk | Mitigation |
|------|------------|
| XML namespace handling complexity | Use established parser with namespace support |
| Large document memory usage | Implement streaming processing |
| Edge cases in revision markup | Port C# tests systematically |
| Browser compatibility issues | Use polyfills, test early |

### Medium Risk

| Risk | Mitigation |
|------|------------|
| Performance regression vs C# | Profile and optimize hot paths |
| OpenXML spec interpretation | Reference C# implementation |
| Test coverage gaps | Measure coverage, add edge cases |

### Low Risk

| Risk | Mitigation |
|------|------------|
| Library deprecation | Use well-maintained libraries |
| TypeScript version changes | Pin major version |

---

## Appendix A: XML Namespace Reference

```typescript
export const NAMESPACES = {
  // WordprocessingML
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  wpg: 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
  wpc: 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
  w14: 'http://schemas.microsoft.com/office/word/2010/wordml',
  w15: 'http://schemas.microsoft.com/office/word/2012/wordml',

  // SpreadsheetML
  s: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  x14: 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main',

  // PresentationML
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  p14: 'http://schemas.microsoft.com/office/powerpoint/2010/main',

  // DrawingML
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  a14: 'http://schemas.microsoft.com/office/drawing/2010/main',
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  dgm: 'http://schemas.openxmlformats.org/drawingml/2006/diagram',

  // Relationships
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  pr: 'http://schemas.openxmlformats.org/package/2006/relationships',

  // Other
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  vml: 'urn:schemas-microsoft-com:vml',
  o: 'urn:schemas-microsoft-com:office:office',
  v: 'urn:schemas-microsoft-com:vml',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',

  // Custom (for comparison engine)
  pt14: 'http://powertools.codeplex.com/2011'
};
```

---

## Appendix B: EMU Units Reference

```typescript
// Office uses EMUs (English Metric Units) for measurements
export const EMU = {
  PER_INCH: 914400,
  PER_CM: 360000,
  PER_POINT: 12700,

  // Common tolerances
  POSITION_TOLERANCE: 91440,  // ~0.1 inch
  SIZE_TOLERANCE: 91440,      // ~0.1 inch

  // Conversion helpers
  inchesToEmu: (inches: number) => Math.round(inches * 914400),
  emuToInches: (emu: number) => emu / 914400,
  pointsToEmu: (points: number) => Math.round(points * 12700),
  emuToPoints: (emu: number) => emu / 12700
};
```

---

*Document created: 2025-12-23*
*Last updated: 2025-12-23*
*Author: Claude Code*
