# Rust Migration Plan: redline-rs

## Executive Summary

Port the TypeScript document comparison library to Rust for:
- **Native Tauri integration** (desktop apps)
- **WASM bundle** (browser/Node.js)
- **Python bindings** via PyO3 (optional future)

**Scope:** ~10,800 lines of TypeScript → estimated ~8,000-12,000 lines of Rust

---

## 1. Project Structure

```
redline-rs/
├── Cargo.toml                    # Workspace root
├── crates/
│   ├── redline-core/             # Pure Rust core library
│   │   ├── Cargo.toml
│   │   └── src/
│   │       ├── lib.rs
│   │       ├── lcs.rs            # LCS algorithm
│   │       ├── hash.rs           # Hashing utilities
│   │       ├── xml.rs            # XML abstraction trait
│   │       ├── package.rs        # OOXML package handling
│   │       ├── namespaces.rs     # XML namespace constants
│   │       ├── wml/              # Word document comparison
│   │       │   ├── mod.rs
│   │       │   ├── comparer.rs
│   │       │   ├── document.rs
│   │       │   ├── revision.rs
│   │       │   ├── revision_accepter.rs
│   │       │   └── types.rs
│   │       ├── sml/              # Excel comparison
│   │       │   ├── mod.rs
│   │       │   ├── comparer.rs
│   │       │   ├── canonicalize.rs
│   │       │   ├── diff.rs
│   │       │   ├── markup.rs
│   │       │   ├── sheets.rs
│   │       │   ├── cells.rs
│   │       │   ├── rows.rs
│   │       │   └── types.rs
│   │       └── pml/              # PowerPoint comparison
│   │           ├── mod.rs
│   │           ├── comparer.rs
│   │           ├── canonicalize.rs
│   │           ├── diff.rs
│   │           ├── markup.rs
│   │           ├── slide_match.rs
│   │           ├── shape_match.rs
│   │           └── types.rs
│   │
│   ├── redline-wasm/             # WASM bindings
│   │   ├── Cargo.toml
│   │   └── src/
│   │       └── lib.rs            # wasm-bindgen exports
│   │
│   └── redline-tauri/            # Tauri plugin (optional)
│       ├── Cargo.toml
│       └── src/
│           └── lib.rs            # Tauri commands
│
├── tests/                        # Integration tests
│   ├── golden/                   # Golden test files (copy from TS)
│   │   ├── wml/
│   │   ├── sml/
│   │   └── pml/
│   └── integration/
│
├── benches/                      # Benchmarks
│   └── comparison_bench.rs
│
└── examples/
    ├── compare_word.rs
    ├── compare_excel.rs
    └── compare_powerpoint.rs
```

---

## 2. Dependency Mapping

| TypeScript | Rust Crate | Notes |
|------------|------------|-------|
| `jszip` | `zip` (v2.x) | Sync API, works in WASM |
| `fast-xml-parser` | `quick-xml` | Fast, streaming XML parser |
| `js-sha256` | `sha2` | Part of RustCrypto |
| (Date handling) | `chrono` | ISO8601 formatting |
| (Buffer) | `bytes` or `Vec<u8>` | Native Rust |

### Cargo.toml (redline-core)

```toml
[package]
name = "redline-core"
version = "0.1.0"
edition = "2024"

[dependencies]
# XML parsing
quick-xml = { version = "0.37", features = ["serialize"] }

# ZIP handling
zip = { version = "2.2", default-features = false, features = ["deflate"] }

# Hashing
sha2 = "0.10"
hex = "0.4"

# Date/time
chrono = { version = "0.4", default-features = false, features = ["std"] }

# Error handling
thiserror = "2.0"

# Serialization (for types)
serde = { version = "1.0", features = ["derive"] }

[dev-dependencies]
insta = "1.41"  # Snapshot testing
pretty_assertions = "1.4"

[features]
default = []
wasm = []  # Feature flag for WASM-specific code
```

### Cargo.toml (redline-wasm)

```toml
[package]
name = "redline-wasm"
version = "0.1.0"
edition = "2024"

[lib]
crate-type = ["cdylib"]

[dependencies]
redline-core = { path = "../redline-core", features = ["wasm"] }
wasm-bindgen = "0.2"
js-sys = "0.3"
serde-wasm-bindgen = "0.6"

[dependencies.web-sys]
version = "0.3"
features = ["console"]
```

---

## 3. Phased Implementation Plan

### Phase 0: Project Setup (1 day)
- [ ] Initialize Cargo workspace
- [ ] Set up CI (GitHub Actions for native + WASM)
- [ ] Configure `wasm-pack` build
- [ ] Set up test infrastructure
- [ ] Copy golden test files from TypeScript

### Phase 1: Core Module (3-4 days)
Port foundation code with 100% test coverage.

| File | Lines | Priority | Complexity |
|------|-------|----------|------------|
| `core/lcs.rs` | ~340 | P0 | Medium |
| `core/hash.rs` | ~43 | P0 | Low |
| `core/namespaces.rs` | ~138 | P0 | Low |
| `core/xml.rs` | ~248 | P0 | Medium |
| `core/package.rs` | ~255 | P0 | Medium |
| `types.rs` | ~210 | P0 | Low |

**Deliverable:** Core library compiles, LCS tests pass.

### Phase 2: WML (Word) Module (5-7 days)
Largest and most complex module.

| File | Lines | Priority | Complexity |
|------|-------|----------|------------|
| `wml/types.rs` | ~177 | P0 | Low |
| `wml/document.rs` | ~513 | P0 | Medium |
| `wml/revision.rs` | ~402 | P0 | High |
| `wml/revision_accepter.rs` | ~193 | P1 | Medium |
| `wml/comparer.rs` | ~1962 | P0 | **Very High** |

**Deliverable:** Word comparison works, passes all WML golden tests.

### Phase 3: SML (Excel) Module (4-5 days)

| File | Lines | Priority | Complexity |
|------|-------|----------|------------|
| `sml/types.rs` | ~266 | P0 | Low |
| `sml/canonicalize.rs` | ~944 | P0 | High |
| `sml/sheets.rs` | ~185 | P0 | Medium |
| `sml/cells.rs` | ~175 | P0 | Medium |
| `sml/rows.rs` | ~41 | P0 | Low |
| `sml/diff.rs` | ~302 | P0 | Medium |
| `sml/markup.rs` | ~977 | P1 | High |
| `sml/comparer.rs` | ~578 | P0 | High |

**Deliverable:** Excel comparison works, passes all SML golden tests.

### Phase 4: PML (PowerPoint) Module (4-5 days)

| File | Lines | Priority | Complexity |
|------|-------|----------|------------|
| `pml/types.rs` | ~389 | P0 | Medium |
| `pml/canonicalize.rs` | ~567 | P0 | High |
| `pml/slide_match.rs` | ~301 | P0 | Medium |
| `pml/shape_match.rs` | ~283 | P0 | Medium |
| `pml/diff.rs` | ~349 | P0 | Medium |
| `pml/markup.rs` | ~597 | P1 | High |
| `pml/comparer.rs` | ~255 | P0 | Medium |

**Deliverable:** PowerPoint comparison works, passes all PML golden tests.

### Phase 5: WASM Bindings (2-3 days)
- [ ] Create `wasm-bindgen` exports
- [ ] Handle JS ↔ Rust type conversions
- [ ] Test in browser and Node.js
- [ ] Create npm package structure
- [ ] Write JavaScript/TypeScript wrapper

### Phase 6: Tauri Plugin (1-2 days)
- [ ] Create Tauri command handlers
- [ ] Test in Tauri app
- [ ] Document integration

### Phase 7: Polish & Documentation (2-3 days)
- [ ] Performance benchmarks vs TypeScript
- [ ] API documentation (rustdoc)
- [ ] Usage examples
- [ ] README and CHANGELOG
- [ ] Publish to crates.io

---

## 4. Type Mapping

### Core Types

```rust
// src/types.rs

use serde::{Deserialize, Serialize};

/// Represents a document as bytes
#[derive(Debug, Clone)]
pub struct Document {
    pub file_name: String,
    pub data: Vec<u8>,
}

/// Revision types for tracked changes
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub enum RevisionType {
    Insertion,
    Deletion,
    MoveFrom,
    MoveTo,
    ParagraphPropertiesChange,
    RunPropertiesChange,
    SectionPropertiesChange,
    StyleDefinitionChange,
    StyleInsertion,
    NumberingChange,
    CellDeletion,
    CellInsertion,
    CellMerge,
    CellPropertiesChange,
    TablePropertiesChange,
    TableGridChange,
}

/// A single revision/change in a document
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Revision {
    pub revision_type: RevisionType,
    pub author: Option<String>,
    pub date: Option<String>,
    pub text: Option<String>,
}
```

### LCS Types

```rust
// src/lcs.rs

/// Correlation status indicating how content relates between documents
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CorrelationStatus {
    Equal,
    Deleted,
    Inserted,
    Unknown,
}

/// Interface for items that can be compared using LCS
pub trait Hashable {
    fn hash(&self) -> &str;
}

/// A correlated sequence showing how portions of two arrays relate
#[derive(Debug, Clone)]
pub struct CorrelatedSequence<T> {
    pub status: CorrelationStatus,
    pub items1: Option<Vec<T>>,
    pub items2: Option<Vec<T>>,
}

/// Settings for the LCS algorithm
#[derive(Debug, Clone, Default)]
pub struct LcsSettings {
    pub min_match_length: usize,
    pub detail_threshold: f64,
}
```

### XML Abstraction

```rust
// src/xml.rs

use quick_xml::events::Event;
use quick_xml::Reader;

/// XML node representation
#[derive(Debug, Clone)]
pub enum XmlNode {
    Element {
        name: String,
        namespace: Option<String>,
        attributes: Vec<(String, String)>,
        children: Vec<XmlNode>,
    },
    Text(String),
    CData(String),
    Comment(String),
}

impl XmlNode {
    pub fn tag_name(&self) -> Option<&str> {
        match self {
            XmlNode::Element { name, .. } => Some(name),
            _ => None,
        }
    }

    pub fn children(&self) -> &[XmlNode] {
        match self {
            XmlNode::Element { children, .. } => children,
            _ => &[],
        }
    }

    pub fn get_attribute(&self, name: &str) -> Option<&str> {
        match self {
            XmlNode::Element { attributes, .. } => {
                attributes.iter()
                    .find(|(k, _)| k == name)
                    .map(|(_, v)| v.as_str())
            }
            _ => None,
        }
    }

    pub fn text_content(&self) -> String {
        match self {
            XmlNode::Text(s) | XmlNode::CData(s) => s.clone(),
            XmlNode::Element { children, .. } => {
                children.iter().map(|c| c.text_content()).collect()
            }
            _ => String::new(),
        }
    }
}

pub fn parse_xml(xml: &str) -> Result<Vec<XmlNode>, XmlError> {
    // Implementation using quick-xml
}

pub fn build_xml(nodes: &[XmlNode]) -> String {
    // Implementation using quick-xml Writer
}
```

---

## 5. Testing Strategy

### Test Categories

1. **Unit Tests** (inline `#[cfg(test)]` modules)
   - Every public function has tests
   - Edge cases: empty inputs, malformed XML, etc.

2. **Golden Tests** (snapshot testing with `insta`)
   - Copy all `.docx`, `.xlsx`, `.pptx` files from TypeScript `tests/golden/`
   - Compare output byte-for-byte or via normalized XML

3. **Property-Based Tests** (using `proptest`)
   - LCS algorithm: random inputs should produce valid correlations
   - XML round-trip: parse → build → parse should be idempotent

4. **Integration Tests**
   - Full document comparison workflows
   - Cross-platform consistency (native vs WASM)

### Coverage Requirements

```toml
# .cargo/config.toml
[build]
rustflags = ["-C", "instrument-coverage"]

# Run with:
# RUSTFLAGS="-C instrument-coverage" cargo test
# grcov . -s . --binary-path ./target/debug/ -t html -o ./coverage/
```

**Target: 100% line coverage for `redline-core`**

### Test File Structure

```
tests/
├── golden/
│   ├── wml/
│   │   ├── WC-1000.docx           # Input files
│   │   ├── WC-1000.document.xml   # Expected output
│   │   └── ...
│   ├── sml/
│   └── pml/
├── wml_comparison_test.rs
├── sml_comparison_test.rs
├── pml_comparison_test.rs
└── common/
    └── mod.rs                     # Shared test utilities
```

### Golden Test Pattern

```rust
// tests/wml_comparison_test.rs

use redline_core::wml::compare_documents;
use std::fs;

macro_rules! golden_test {
    ($name:ident, $test_id:expr) => {
        #[test]
        fn $name() {
            let older = fs::read(format!("tests/golden/wml/{}.docx", $test_id)).unwrap();
            let newer = fs::read(format!("tests/golden/wml/{}.docx", $test_id)).unwrap();
            
            let result = compare_documents(&older, &newer, Default::default()).unwrap();
            
            // Snapshot test the output
            insta::assert_snapshot!(
                format!("{}_output", $test_id),
                extract_document_xml(&result.document)
            );
        }
    };
}

golden_test!(wc_1000, "WC-1000");
golden_test!(wc_1010, "WC-1010");
// ... generate for all test cases
```

---

## 6. WASM Considerations

### Memory Management
- Use `wasm-bindgen` for automatic memory handling
- Return `JsValue` for complex types
- Use `serde-wasm-bindgen` for struct serialization

### API Surface

```rust
// crates/redline-wasm/src/lib.rs

use wasm_bindgen::prelude::*;
use redline_core::{wml, sml, pml};

#[wasm_bindgen]
pub fn compare_word_documents(
    older: &[u8],
    newer: &[u8],
    settings: JsValue,
) -> Result<JsValue, JsError> {
    let settings: wml::WmlComparerSettings = serde_wasm_bindgen::from_value(settings)?;
    let result = wml::compare_documents(older, newer, settings)?;
    Ok(serde_wasm_bindgen::to_value(&result)?)
}

#[wasm_bindgen]
pub fn compare_spreadsheets(
    older: &[u8],
    newer: &[u8],
    settings: JsValue,
) -> Result<JsValue, JsError> {
    let settings: sml::SmlComparerSettings = serde_wasm_bindgen::from_value(settings)?;
    let result = sml::compare(older, newer, settings)?;
    Ok(serde_wasm_bindgen::to_value(&result)?)
}

#[wasm_bindgen]
pub fn compare_presentations(
    older: &[u8],
    newer: &[u8],
    settings: JsValue,
) -> Result<JsValue, JsError> {
    let settings: pml::PmlComparerSettings = serde_wasm_bindgen::from_value(settings)?;
    let result = pml::compare_presentations(older, newer, settings)?;
    Ok(serde_wasm_bindgen::to_value(&result)?)
}
```

### Build Configuration

```toml
# wasm-pack build configuration
# Build with: wasm-pack build --target web

[package.metadata.wasm-pack.profile.release]
wasm-opt = ["-O4", "--enable-simd"]
```

---

## 7. Error Handling

```rust
// src/error.rs

use thiserror::Error;

#[derive(Error, Debug)]
pub enum RedlineError {
    #[error("Invalid OOXML package: {0}")]
    InvalidPackage(String),

    #[error("Missing required part: {0}")]
    MissingPart(String),

    #[error("XML parsing error: {0}")]
    XmlParse(#[from] quick_xml::Error),

    #[error("ZIP error: {0}")]
    Zip(#[from] zip::result::ZipError),

    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),

    #[error("UTF-8 encoding error: {0}")]
    Utf8(#[from] std::string::FromUtf8Error),
}

pub type Result<T> = std::result::Result<T, RedlineError>;
```

---

## 8. Timeline Estimate

| Phase | Duration | Dependencies |
|-------|----------|--------------|
| Phase 0: Setup | 1 day | None |
| Phase 1: Core | 3-4 days | Phase 0 |
| Phase 2: WML | 5-7 days | Phase 1 |
| Phase 3: SML | 4-5 days | Phase 1 |
| Phase 4: PML | 4-5 days | Phase 1 |
| Phase 5: WASM | 2-3 days | Phases 2-4 |
| Phase 6: Tauri | 1-2 days | Phases 2-4 |
| Phase 7: Polish | 2-3 days | All |

**Total: 22-30 working days** (~4-6 weeks)

Phases 2-4 can be parallelized if multiple developers are available.

---

## 9. Verification Checklist

Before considering each phase complete:

- [ ] All unit tests pass
- [ ] All golden tests pass
- [ ] Code coverage ≥ 100% for new code
- [ ] `cargo clippy` has no warnings
- [ ] `cargo fmt` applied
- [ ] Documentation complete (rustdoc)
- [ ] WASM build succeeds (after Phase 5)
- [ ] No `unsafe` code (or justified and audited)

---

## 10. Reference Materials

### Original C# Source
- Open-Xml-PowerTools: https://github.com/OfficeDev/Open-Xml-PowerTools

### Rust Crate Documentation
- quick-xml: https://docs.rs/quick-xml
- zip: https://docs.rs/zip
- wasm-bindgen: https://rustwasm.github.io/wasm-bindgen/

### OOXML Specifications
- ECMA-376: https://www.ecma-international.org/publications-and-standards/standards/ecma-376/

---

## Appendix A: File-by-File Porting Notes

### `wml/wml-comparer.ts` (1962 lines) — CRITICAL PATH

This is the largest and most complex file. Key functions:

1. `compareDocuments()` — Main entry point
2. `correlateSequences()` — Uses LCS to align paragraphs
3. `compareRuns()` — Character-level comparison
4. `produceMarkedDocument()` — Generate redlined output

**Porting strategy:**
- Start with types and signatures
- Port helper functions bottom-up
- Main comparison logic last
- Test each function in isolation

### `sml/canonicalize.ts` (944 lines)

Creates normalized signatures for comparison. Key challenges:
- Complex cell formatting normalization
- Shared string table handling
- Style inheritance resolution

### `pml/markup.ts` (597 lines)

Generates visual markup in PowerPoint. Key challenges:
- Shape geometry manipulation
- Text frame positioning
- Color/style application

---

## Appendix B: Migration Tracking

Use this checklist to track progress:

```
[ ] Phase 0: Setup
    [ ] Cargo workspace initialized
    [ ] CI configured
    [ ] Test infrastructure ready

[ ] Phase 1: Core
    [ ] lcs.rs - 100% coverage
    [ ] hash.rs - 100% coverage  
    [ ] namespaces.rs - 100% coverage
    [ ] xml.rs - 100% coverage
    [ ] package.rs - 100% coverage
    [ ] types.rs - 100% coverage

[ ] Phase 2: WML
    [ ] wml/types.rs
    [ ] wml/document.rs
    [ ] wml/revision.rs
    [ ] wml/revision_accepter.rs
    [ ] wml/comparer.rs
    [ ] All WML golden tests pass

[ ] Phase 3: SML
    [ ] sml/types.rs
    [ ] sml/canonicalize.rs
    [ ] sml/sheets.rs
    [ ] sml/cells.rs
    [ ] sml/rows.rs
    [ ] sml/diff.rs
    [ ] sml/markup.rs
    [ ] sml/comparer.rs
    [ ] All SML golden tests pass

[ ] Phase 4: PML
    [ ] pml/types.rs
    [ ] pml/canonicalize.rs
    [ ] pml/slide_match.rs
    [ ] pml/shape_match.rs
    [ ] pml/diff.rs
    [ ] pml/markup.rs
    [ ] pml/comparer.rs
    [ ] All PML golden tests pass

[ ] Phase 5: WASM
    [ ] wasm-bindgen exports
    [ ] npm package structure
    [ ] Browser tests pass
    [ ] Node.js tests pass

[ ] Phase 6: Tauri
    [ ] Tauri commands
    [ ] Integration tested

[ ] Phase 7: Polish
    [ ] Benchmarks complete
    [ ] Documentation complete
    [ ] Published to crates.io
```

---

## Appendix C: TypeScript Implementation Fidelity Notes

**CRITICAL:** The TypeScript port diverges from the original C# in several important ways.
These heuristics and bug fixes were discovered during the 104-test verification process
and MUST be replicated in the Rust port.

### C1. Image and Relationship Preservation (commit 02591e9)

The C# implementation had issues with embedded images and relationships. The TS fixes:

```typescript
// 1. extractStructuralElements() - Map drawing/picture elements during comparison
//    Instead of converting to DRAWING_ text tokens, preserve actual XML structure
function extractStructuralElements(node: XmlNode): Map<string, XmlNode>

// 2. stripSectPrFromProperties() - Remove sectPr from deleted paragraph properties
//    Prevents orphan rId references that cause "Word found unreadable content"
function stripSectPrFromProperties(pPr: XmlNode): XmlNode

// 3. wrapParagraphRunsWithRevision() - Preserve structural elements in revisions
//    When marking entire paragraphs as inserted/deleted, keep drawings intact
function wrapParagraphRunsWithRevision(runs: XmlNode[], type: 'ins' | 'del'): XmlNode[]
```

**Rust implementation:** Must handle drawing elements as structural, not text tokens.

### C2. Drawing ID Normalization (commit 9515525)

```typescript
// docPr/@id varies between documents even for identical drawings
// MUST be ignored during comparison, or SmartArt/images falsely detected as changed
function normalizeDrawing(drawing: XmlNode): string {
  // Remove/normalize docPr/@id and docPr/@name before hashing
}
```

### C3. Heuristic Algorithm Differences

The TypeScript implementation uses **word-level** comparison instead of the C#
**character-level atoms**. This is a fundamental architectural difference:

| Aspect | C# Original | TypeScript Port |
|--------|-------------|-----------------|
| Granularity | Character atoms | Word tokens |
| Grouping | Atom→Word→Para→Cell→Row→Table | Flat paragraph units |
| LCS Level | Recursive at each hierarchy | Single level + post-hoc |
| Threshold | DetailThreshold=0.15 | 0.4/0.5 similarity |

#### Key Heuristic Functions

```typescript
// HEURISTIC: Word-level Jaccard similarity with empirical thresholds
function calculateSimilarity(text1: string, text2: string): number
// Returns 0-1, use 0.4 for "similar enough", 0.5 for "likely same content"

// HEURISTIC: Detect footnote/endnote refs splitting words
function findSplittingReferences(tokens: Token[]): SplitInfo[]
// Example: "Vi"+"deo" → "Video" when split by footnote reference

// HEURISTIC: Categorize changes as meaningful/punctuation/reference
function classifySequence(tokens: Token[]): SequenceCategory
// Helps group related changes and ignore noise

// HEURISTIC: Group changes separated only by structural tokens
const STRUCTURAL_PREFIXES = ['DRAWING_', 'PICT_', 'MATH_', 'TXBX_']
// If changes are scattered ONLY by these, treat as single modification
```

### C4. Table Row Comparison (commits a516762, 2015168)

```typescript
// Tables are compared at ROW level, not cell-by-cell
// This produces cleaner diffs for row insertions/deletions
interface ParagraphUnit {
  isTableRow?: boolean;
  rowCells?: XmlNode[][];  // Paragraphs grouped by cell
}

// For footnotes/endnotes: positional table row comparison
function compareFootnoteTables(before: XmlNode[], after: XmlNode[]): Correlation[]
```

### C5. Paragraph Similarity Threshold (commit 6f53927)

```typescript
// If paragraph similarity < 0.4, treat as complete replacement
// rather than showing character-level changes
const COMPLETE_REPLACEMENT_THRESHOLD = 0.4;

if (calculateSimilarity(oldText, newText) < COMPLETE_REPLACEMENT_THRESHOLD) {
  // Mark entire old paragraph as deleted, entire new as inserted
  // Instead of showing confusing inline changes
}
```

### C6. Footnote/Endnote Support (commit 21014e6)

```typescript
// Extract and compare footnotes/endnotes separately
function extractFootnotes(doc: WordDocument): Map<string, XmlNode>
function extractEndnotes(doc: WordDocument): Map<string, XmlNode>

// Include empty paragraphs (commit 4c17382)
// They carry formatting and affect structure

// Handle splitting references (mentioned in C3)
```

### C7. Adjacent Delete-Insert Handling (commit 73bb3f9)

```typescript
// When delete immediately precedes insert, may need special grouping
// For revision counting: Adjacent delete+insert = 1 modification (not 2)
function groupAdjacentRevisions(correlations: Correlation[]): RevisionGroup[]
```

### C8. Format Change Detection (commits d76666a, e7a41fb)

```typescript
// Track formatting changes: bold, italic, underline, color, size, font, etc.
interface FormatChange {
  property: string;  // 'bold' | 'italic' | 'underline' | 'color' | 'sz' | ...
  oldValue: string | boolean;
  newValue: string | boolean;
}

function detectFormatChanges(runBefore: XmlNode, runAfter: XmlNode): FormatChange[]
```

### C9. Accept Tracked Changes Before Comparison (commit 883ca44)

```typescript
// If documents have existing tracked changes, accept them first
// This ensures we're comparing actual content, not revision markup
import { acceptRevisions } from './revision-accepter';

function compareDocuments(before: Buffer, after: Buffer): Result {
  const beforeClean = acceptRevisions(before);
  const afterClean = acceptRevisions(after);
  return compare(beforeClean, afterClean);
}
```

### C10. Metadata/Change List Generation (commits faa0004, 06195eb)

```typescript
// Generate UI-friendly change metadata alongside redlined document
interface WmlChangeListItem {
  type: 'insertion' | 'deletion' | 'modification' | 'formatting';
  preview: string;      // First ~50 chars
  wordCount: number;
  location: string;     // e.g., "Paragraph 3", "Table 1, Row 2"
}

function generateChangeList(result: ComparisonResult): WmlChangeListItem[]
```

---

## Appendix D: Regression Test Cases

The following commits added specific regression tests that MUST pass in Rust:

| Commit | Test Focus | Test Count |
|--------|------------|------------|
| 02591e9 | Image preservation, sectPr stripping | 8 tests |
| 9515525 | docPr id normalization | WC-1940 |
| 21014e6 | Footnote/endnote comparison | Multiple |
| a516762 | Table row comparison | Multiple |
| 6f53927 | Similarity threshold | Multiple |

All 104 WML golden tests from `tests/golden/wml/` must pass.
