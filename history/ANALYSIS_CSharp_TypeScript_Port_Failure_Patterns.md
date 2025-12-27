# Analysis: Common Failure Patterns in C# to TypeScript Document Processing Ports

**Date**: December 27, 2025  
**Analysis Type**: Deep Multi-Agent Research (10+ parallel agents)  
**Scope**: OpenXML document processing code (DOCX/XLSX/PPTX comparison engines)

---

## Executive Summary

This analysis examines systematic failure patterns when porting document processing code from C# (OpenXML PowerTools) to TypeScript (redline-js). Based on comprehensive codebase analysis, specification research, and expert oracle consultation, we identify **three critical failure categories** and their root causes:

1. **Images & Relationships** - Package-level invariant violations
2. **Tracked Revisions** - Tree-rewrite complexity with bidirectional transformations
3. **Locale/Internationalization** - Culture-dependent text processing mismatches

**Key Finding**: These failures stem from **semantic gaps** between C# + OpenXML SDK (strongly-modeled part graph with implicit invariants) and TypeScript + raw XML/ZIP libraries (explicit invariant management required).

---

## 1. Images & Relationships: Package-Level Invariant Violations

### Root Causes

#### 1.1 Library Model Gap

**C# Approach**:
- OpenXML SDK models images/drawings as **typed parts** (`ImagePart`, `Drawing`, `RelationshipId`)
- Strong typing makes invalid packages **hard to create accidentally**
- Relationship graph managed automatically

**TypeScript Approach**:
- Manipulates **ZIP entries + XML strings** directly
- Package invariants are **easy to violate silently**
- Manual relationship management required

#### 1.2 Cross-File Invariants

Images depend on **consistent state across multiple files**:

```
word/document.xml         → Contains drawing markup referencing r:embed="rId3"
word/_rels/document.xml.rels → Maps rId3 → /word/media/image1.png
[Content_Types].xml       → Declares content type for PNG
word/media/image1.png     → Actual image binary
```

**Breaking any link causes "unreadable content" errors.**

#### 1.3 ID Allocation Assumptions

**C# Assumption**: IDs are stable and unique (relationship IDs, `docPr/@id`, `cNvPr/@id`)

**TypeScript Risk**: Naïve cloning/merging can **duplicate IDs**, causing:
- Word repair actions (deletes drawings)
- File rejection ("this file is corrupted")

### Failure Mechanisms

Word/PowerPoint reports "unreadable content" when:

| Failure Type | Example | Impact |
|--------------|---------|--------|
| **Broken relationship** | XML references `rId9` but `.rels` missing entry | Fatal |
| **Missing part** | Relationship exists but `media/image3.png` deleted | Fatal |
| **Wrong content type** | `[Content_Types].xml` lacks PNG declaration | Fatal |
| **Duplicate IDs** | Two drawings with same `docPr/@id` | Repair/deletion |
| **Namespace loss** | Rewriting removes `r:` prefix bindings | Fatal |

### TypeScript Port Status

**Fixes Applied** (commits 02591e9, 9515525):
1. ✅ **Structural preservation**: `extractStructuralElements()` preserves drawing XML instead of text tokens
2. ✅ **sectPr stripping**: `stripSectPrFromProperties()` removes orphan relationship sources
3. ✅ **docPr normalization**: Ignores `docPr/@id` and `@name` during comparison

**Remaining Risks**:
- ⚠️ **Coverage**: Handles `w:drawing` with `a:blip r:embed`, but what about:
  - Headers/footers, footnotes/endnotes, comments
  - VML (`v:imagedata`) in older documents
  - Charts with embedded drawings, SmartArt, group shapes
  - External relationships (linked images), theme images
- ⚠️ **Abstraction level**: Still "string-edit XML + hope rels stay valid"

### Architectural Recommendations

```typescript
// RECOMMENDED: Part-graph abstraction layer

interface PackageManager {
  // Relationship management
  allocateRelationshipId(sourcePart: string, targetPart: string, type: string): string;
  getRelationshipTarget(sourcePart: string, rId: string): string;
  
  // Content type registration
  registerContentType(extension: string, contentType: string): void;
  
  // Safe copy/move with relationship rewrites
  copyPartWithRelationships(source: string, dest: string): void;
}

// Validation gates (run after every transform)
interface PackageValidator {
  validateRelationships(): ValidationResult;  // Every rId has target
  validateContentTypes(): ValidationResult;   // Every part has content type
  validateDrawingIds(): ValidationResult;     // All docPr/@id unique
}
```

**Effort**: Medium (1-2 days)

---

## 2. Tracked Revisions: Tree-Rewrite Complexity

### Root Causes

#### 2.1 Tree Semantics, Not Plain Text

Tracked changes are **parallel structure embedded in XML tree**:
- `w:ins`, `w:del`, `w:moveFrom`, `w:moveTo`
- `w:delText`, `w:delInstrText` (special text elements)
- Change ranges, properties changes (`w:rPrChange`, `w:pPrChange`)

**Not just diff markers** - full semantic structures with rules.

#### 2.2 Strong Typing vs Weak Typing

**C# OpenXML SDK**:
- Nudges into correct constructs
- Well-known patterns for accept/reject
- Type errors prevent illegal nesting

**TypeScript**:
- Generic nodes → can create **illegal nesting** accidentally
- Can **split elements incorrectly** across revision boundaries

#### 2.3 Normalization Assumptions

**C# Implementation** (atom/word/paragraph hierarchy):
- Applies changes while preserving run boundaries
- Maintains properties and revision containers

**TypeScript Port** (word-level LCS + heuristics):
- Rewrites runs more aggressively
- **Collides with revision markup rules**

### Common Failure Patterns

#### Pattern 1: Illegal Nesting / Broken Ranges

```xml
<!-- ILLEGAL: Split run inside w:del -->
<w:del w:id="0">
  <w:r><w:t>Hel</w:t></w:r>
</w:del>
<w:r><w:t>lo</w:t></w:r>  <!-- Orphaned! -->
```

#### Pattern 2: Property Changes vs Content Changes

```xml
<!-- Paragraph properties change -->
<w:pPr>
  <w:numId w:val="2"/>  <!-- NEW VALUE -->
  <w:pPrChange w:id="0">
    <w:pPr>
      <w:numId w:val="1"/>  <!-- OLD VALUE -->
    </w:pPr>
  </w:pPrChange>
</w:pPr>
```

**Problem**: Must swap current/old when rejecting. Text-diff algorithms ignore these.

#### Pattern 3: Bidirectional Transformations

**Rejection = Reversal + Acceptance**:

```typescript
// Step 1: Reverse the sense
w:ins  → w:del
w:del  → w:ins
w:moveFrom ↔ w:moveTo

// Step 2: Accept the reversed revisions
// (Run acceptance algorithm on reversed markup)
```

**Any asymmetry causes corruption.**

#### Pattern 4: Context-Dependent Processing

```csharp
// Same element, different handling based on location
if (element.Name == W.del && parent.Name == W.p)
    return new XElement(W.ins, ...);  // Deleted run

if (element.Name == W.del && parent.Name == W.rPr && 
    parent.Parent.Name == W.pPr)
    return new XElement(W.ins);  // Deleted paragraph mark
```

#### Pattern 5: Range Matching Complexity

```xml
<w:moveFromRangeStart w:id="1" w:name="move478160808"/>
  <!-- Content spanning multiple paragraphs -->
<w:moveFromRangeEnd w:id="1"/>
```

**Algorithm**: O(n²) - must track elements between start/end pairs across entire document.

### Edge Cases (10+ Special Handlers Required)

| Element Type | Complexity | Example |
|--------------|------------|---------|
| **Field codes** | High | `w:fldChar`, `w:instrText` → `w:delInstrText` |
| **Paragraph marks** | High | Merge paragraphs, decide which properties to keep |
| **Table rows** | High | `w:trPr/w:del`, `cellDel`, `cellMerge`, `cellIns` |
| **Move operations** | Very High | Match by name, handle nested moves, range deletion |
| **Content controls** | Medium | Range markers crossing boundaries |
| **Math equations** | Medium | Different namespace (`m:`), different rules |
| **Textboxes** | Medium | Duplicate VML shapes with same ID |
| **Footnotes/Endnotes** | Medium | Revisions spanning across parts |

### TypeScript Port Status

**Implemented**:
- ✅ Basic revision markup generation (`w:ins`, `w:del`)
- ✅ Sequential ID assignment
- ✅ Footnote/endnote support (commit 21014e6)

**Missing/At Risk**:
- ⚠️ **Property change tracking** (`w:rPrChange`, `w:pPrChange`)
- ⚠️ **Move operations** (`w:moveFrom`, `w:moveTo`, range matching)
- ⚠️ **Table-specific revisions** (row/cell deletion, merging)
- ⚠️ **Field code revisions** (begin/instrText/end grouping)
- ⚠️ **Revision rejection** (reversal + acceptance pipeline)

**Test Coverage**: 104/104 tests passing, but:
- Most tests are **insertion/deletion only**
- Few tests for **property changes** or **move operations**
- No tests for **reject revisions**
- Revisions are "long tail" - real docs contain constructs unit tests don't

### Architectural Recommendations

```typescript
// RECOMMENDED: Revision-aware normalization pipeline

class RevisionProcessor {
  // Phase 1: Normalize input
  normalizeInput(doc: Document): Document {
    // Collapse trivial runs
    // Canonicalize whitespace
    // Normalize revision containers
  }
  
  // Phase 2: Diff on stable model
  generateDiff(doc1: Document, doc2: Document): DiffModel {
    // Diff on atoms/words with provenance
    // Track which atoms came from which revision
  }
  
  // Phase 3: Emit with constraints
  emitRevisions(diff: DiffModel): Document {
    // Don't split inside protected constructs
    // Treat revision containers as boundaries
    // Enforce schema-like constraints
  }
  
  // Phase 4: Re-normalize and validate
  validate(doc: Document): ValidationResult {
    // Check illegal nesting
    // Verify ID uniqueness
    // Ensure property changes have old/new values
  }
}

// Boundary rules: Unsplittable spans
const UNSPLITTABLE = [
  'field',          // Field codes (begin/instrText/end)
  'hyperlink',      // Hyperlink markup
  'revision',       // Existing revision nodes
  'footnoteRef',    // Footnote/endnote references
  'drawing',        // Drawing anchors
];
```

**Effort**: Large (3+ days for full revision surface area)

---

## 3. Locale/Internationalization: Culture-Dependent Text Processing

### Root Causes

#### 3.1 Different Culture Engines

**.NET**:
- `CultureInfo` + collation/casing rules
- Well-defined behavior per locale

**JavaScript**:
- ICU via `Intl` (varies by runtime)
- Node version affects behavior
- OS ICU data can differ

**Result**: Same operation, **different output** across platforms.

#### 3.2 Unicode Complexity Exposed by Simplification

**C# character-level atoms**: Complexity hidden in granular comparison

**TypeScript word-level tokens**: Increased dependence on:
- **Tokenization rules** (word boundaries)
- **Punctuation handling** (language-specific)
- **Whitespace rules** (varies by script)

**All are language-dependent.**

#### 3.3 Normalization and Equivalence

**Unicode Equivalence**:
- **NFC** (Canonical Composition): Precomposed characters
- **NFD** (Canonical Decomposition): Base + combining marks

```javascript
// Visually identical, byte-different
"café" === "\u00E9"         // Precomposed (NFC)
"café" === "e\u0301"        // Decomposed (NFD)

"café".length === 4  // NFC
"café".length === 5  // NFD
```

**C# and JS differ in defaults.**

### Failure Mechanisms

#### Failure 1: Locale-Specific List Formatting

**Evidence**: 6 locale implementations in C# codebase:

| Locale | Format | Example |
|--------|--------|---------|
| English | ordinal | "1st", "2nd", "3rd" |
| French | ordinal | "1er", "2e", "3e" |
| Chinese | counting | "一", "二", "三" |
| Russian | ordinal | **Incomplete (TODO)** |

**Comparison Failure**:
```
Document A (en-US): "1st item"
Document B (fr-FR): "1er item"
Result: ❌ MISMATCH (different text, same semantic)
```

#### Failure 2: Case Folding Differences

**Turkish i/İ Problem**:
```typescript
"i".toUpperCase()  // English: "I"
"i".toUpperCase()  // Turkish: "İ" (dotted capital I)

"I".toLowerCase()  // English: "i"
"I".toLowerCase()  // Turkish: "ı" (dotless small i)
```

**Locale-invariant compare may not match JS defaults.**

#### Failure 3: Word Segmentation

| Script | Segmentation |
|--------|--------------|
| English | Spaces work |
| Chinese | **No spaces** (dictionary-based) |
| Thai/Lao | **No spaces** (dictionary-based) |
| Mixed | **Context-dependent** |

**Simple `split(/\s+/)` fails for CJK languages.**

#### Failure 4: Invisible Unicode Marks

**RTL (Right-to-Left) Text**:
```xml
<w:lang w:val="en-US" w:bidi="ar-SA"/>
```

**Invisible control characters**:
- `U+200E` (LRM) - Left-to-Right Mark
- `U+200F` (RLM) - Right-to-Left Mark

**Comparison sees different bytes**:
```
Document A: "Hello"
Document B: "\u200EHello\u200E"
Result: ❌ MISMATCH
```

#### Failure 5: Decimal Separator Locale

```csharp
// English: 3.14 (period)
// French:  3,14 (comma)
// German:  3,14 (comma)

double.TryParse("3.14", NumberStyles.Float, 
                CultureInfo.InvariantCulture, out dv)  // ✅ Works

double.TryParse("3.14", NumberStyles.Float, 
                new CultureInfo("fr-FR"), out dv)      // ❌ Fails
```

#### Failure 6: Grapheme Clusters

**Emoji + Combining Marks**:
```javascript
"👨‍👩‍👧".length === 5  // Not 1!

// Character-level iteration breaks clusters:
[..."👨‍👩‍👧"]  // ['👨', '‍', '👩', '‍', '👧']
```

### TypeScript Port Status

**Implemented**:
- ✅ Basic word tokenization (English/Western European)
- ✅ `localeCompare` for sorting

**Missing/At Risk**:
- ❌ **Locale-specific list normalization**
- ❌ **Unicode normalization** (NFC/NFD)
- ❌ **RTL language support** (C# warns: "not fully engineered")
- ❌ **CJK word segmentation**
- ❌ **Grapheme-aware operations**

**Test Coverage**:
- Tests are **mostly English/Western European**
- No CJK tests
- No RTL tests
- No mixed-script tests

**Risk Level**: Medium-High for international documents

### Architectural Recommendations

```typescript
// RECOMMENDED: Explicit locale handling

interface TextProcessor {
  // Normalize before comparison
  normalize(text: string, mode: NormalizationMode): string;
  
  // Locale-aware or locale-invariant
  compare(text1: string, text2: string, locale?: string): number;
  
  // Segment into words/graphemes
  segment(text: string, type: 'word' | 'grapheme', locale?: string): string[];
}

enum NormalizationMode {
  NFC,           // Canonical composition
  NFD,           // Canonical decomposition
  StripBidi,     // Remove directional marks
  Lowercase,     // Locale-invariant lowercase
}

// Token model with metadata
interface Token {
  text: string;           // Original text
  normalized: string;     // NFC-normalized form
  script?: string;        // 'Latin' | 'Han' | 'Arabic' | 'Hebrew'
  locale?: string;        // Optional locale hint
}

// List item normalization
function normalizeListItem(text: string, locale: string): string {
  // Convert "1st" → "1", "1er" → "1", "一" → "1"
  // Use locale-specific rules from GetListItemText_*.cs
}
```

**Golden Tests Required**:
```typescript
describe('Internationalization', () => {
  it('handles Turkish casing', () => {
    // Test i/İ, I/ı differences
  });
  
  it('handles French accents (NFC/NFD)', () => {
    // Test "café" in both forms
  });
  
  it('handles Chinese text (no spaces)', () => {
    // Test word segmentation
  });
  
  it('handles Arabic RTL text', () => {
    // Test directional marks
  });
  
  it('handles emoji grapheme clusters', () => {
    // Test compound emoji
  });
});
```

**Effort**: Medium (1-2 days for basic support, larger for comprehensive CJK/RTL)

---

## 4. Type System and Collection Pattern Differences

### 4.1 LINQ to Array Methods

**C# LINQ** (lazy evaluation):
```csharp
var result = items
    .Where(x => x.IsValid)
    .Select(x => x.Name)
    .ToList();  // Evaluation happens here
```

**TypeScript** (eager evaluation):
```typescript
const result = items
    .filter(x => x.isValid)  // Executes immediately
    .map(x => x.name);       // Executes immediately
```

**Performance Impact**:
- C# can **short-circuit** and **avoid intermediate collections**
- TypeScript **creates intermediate arrays** for each operation

**Porting Challenge**: Complex LINQ chains with deferred execution need manual optimization in TypeScript.

### 4.2 Null Safety

**C#** (nullable reference types):
```csharp
string? nullableString = null;  // Explicit nullable
string nonNullString = "";      // Non-nullable (enforced)
```

**TypeScript**:
```typescript
let nullableString: string | null = null;
let undefinedString: string | undefined = undefined;
let bothString: string | null | undefined;  // Both!
```

**Porting Challenge**: C# has `null` only, TypeScript has `null` **and** `undefined`.

### 4.3 Collection Equality

**C# HashSet**:
```csharp
class MyClass {
    public override int GetHashCode() { ... }
    public override bool Equals(object obj) { ... }
}

var set = new HashSet<MyClass>();  // Uses custom equality
```

**JavaScript Set**:
```javascript
const set = new Set();
set.add({id: 1});
set.add({id: 1});  // Different object, both added!
```

**Porting Challenge**: JS Sets use **reference equality**, not value equality.

### 4.4 XML Namespace Handling

**C# LINQ-to-XML**:
```csharp
XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
var paragraphs = doc.Descendants(W + "p");  // Strongly-typed
```

**TypeScript**:
```typescript
// Manual namespace handling required
const paragraphs = findElementsByNamespaceAndTag(
  doc,
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  "p"
);
```

**Porting Challenge**: No standard TypeScript library provides LINQ-to-XML equivalent.

---

## 5. Synthesis: Cross-Cutting Patterns

### Pattern 1: Semantic Gap

| Aspect | C# + OpenXML SDK | TypeScript + Libraries |
|--------|------------------|------------------------|
| **Package Model** | Strongly-typed part graph | ZIP + XML strings |
| **Invariants** | Enforced by SDK | Must validate manually |
| **Relationships** | Automatic management | Manual tracking required |
| **Namespaces** | Built-in support | Custom implementation |
| **Culture** | `CultureInfo` comprehensive | `Intl` limited, varies by runtime |

### Pattern 2: Complexity Migration

**C# hides complexity** (in SDK and framework):
- Relationship management
- Content type registration
- Namespace handling
- Culture-aware text processing

**TypeScript exposes complexity** (must implement explicitly):
- Part graph validators
- Relationship rewrite logic
- Namespace normalization
- Locale-aware tokenization

### Pattern 3: Test Coverage Gaps

**C# test coverage**: Comprehensive (200+ tests)

**TypeScript test coverage**: 104 tests passing, but:
- ⚠️ Mostly English/Western European
- ⚠️ Limited RTL/CJK coverage
- ⚠️ Few property change tests
- ⚠️ No reject revisions tests
- ⚠️ Limited edge case coverage

**Recommendation**: Add 50+ international and edge case tests.

---

## 6. Recommendations by Priority

### P0: Critical (Must Implement)

| Recommendation | Effort | Impact |
|----------------|--------|--------|
| **Part-graph validation layer** | 1-2 days | Prevents "unreadable content" errors |
| **Relationship rewrite pipeline** | 1 day | Prevents rId corruption |
| **Revision boundary rules** | 2-3 days | Prevents illegal revision nesting |
| **Unicode NFC normalization** | 4 hours | Prevents false diffs on accented text |

**Total P0 Effort**: ~5-7 days

### P1: Important (Should Implement)

| Recommendation | Effort | Impact |
|----------------|--------|--------|
| **Property change tracking** | 1-2 days | Enables full revision support |
| **Move operation support** | 2 days | Handles complex change scenarios |
| **RTL language support** | 2-3 days | Enables Arabic/Hebrew documents |
| **CJK word segmentation** | 1-2 days | Enables Chinese/Japanese/Korean |

**Total P1 Effort**: ~6-9 days

### P2: Nice-to-Have (Future Work)

| Recommendation | Effort | Impact |
|----------------|--------|--------|
| **Full reject revisions** | 2-3 days | Enables bidirectional revision handling |
| **Table-specific revisions** | 1-2 days | Handles row/cell changes |
| **Field code revisions** | 1 day | Handles complex field changes |
| **Comprehensive intl tests** | 2-3 days | Ensures broad compatibility |

**Total P2 Effort**: ~6-9 days

---

## 7. Conclusion

### Key Findings

1. **Images fail due to package-level invariant violations**
   - Root cause: Raw XML/ZIP manipulation vs. OpenXML SDK part graph
   - Solution: Add part-graph abstraction with validators

2. **Revisions fail due to tree-rewrite complexity**
   - Root cause: Bidirectional transformations on context-dependent structures
   - Solution: Revision-aware normalization pipeline with boundary rules

3. **Locales fail due to culture-dependent text processing**
   - Root cause: .NET `CultureInfo` vs. JavaScript `Intl` differences
   - Solution: Explicit locale handling with Unicode normalization

### Success Criteria

TypeScript port achieves **adequate robustness** when:
- ✅ All package-level invariants validated before write
- ✅ Revision boundaries respected (no illegal nesting)
- ✅ Unicode normalized before comparison (NFC)
- ✅ Locale-aware or locale-invariant text operations (explicit choice)
- ✅ Test coverage includes international and edge cases

### Escalation Triggers

Consider **full architectural overhaul** if:
- Recurring "unreadable content" reports from real documents
- Revision tracking failures in complex documents
- Need to support broad multilingual corpora (CJK/Thai/RTL)
- Performance issues with 100+ page documents

---

## Appendix: Research Methodology

**Analysis conducted**: December 27, 2025

**Agents deployed**:
- 3 explore agents (C# codebase, TypeScript port, gap analysis)
- 3 librarian agents (OpenXML spec, revisions, internationalization)
- 3 oracle agents (architecture, performance, XML handling)
- 2 general analysis agents (type systems, collections)

**Total agent-hours**: ~11 agent-hours in parallel

**Files analyzed**:
- C#: WmlComparer.cs (8,835 lines), RevisionProcessor.cs, locale implementations
- TypeScript: wml-comparer.ts (1,962 lines), 104 test cases
- Specifications: ECMA-376, Unicode UAX #9, ICU documentation

**Evidence sources**:
- Direct code inspection (line-by-line comparison)
- Git commit analysis (bug fix patterns)
- OpenXML SDK documentation
- Unicode Technical Reports
- Real-world test case failures

---

**End of Analysis**
