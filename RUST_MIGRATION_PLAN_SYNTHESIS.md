# Rust Migration Plan: C# → Rust Port of OpenXML Document Comparers

## Executive Summary

**Objective:** Create a faithful, byte-for-byte compatible Rust port of the C# WmlComparer, SmlComparer, and PmlComparer modules that:
- Runs natively in **Tauri** desktop applications
- Compiles to **WebAssembly** for browser and Node.js
- Provides **python-wasmer** bindings for backend use

**Scope:** 22,727 lines of C# → estimated 18,000-25,000 lines of Rust

**Non-Negotiables:**
1. **Exact functional equivalence** to the current C# implementation
2. **No algorithm changes**, no heuristics, no output changes (XML structure, change counts, JSON shapes)
3. **Deterministic behavior** across native and WASM builds
4. **100% of specified tests pass** after porting
5. **Attribute order preservation** in all XML operations (critical for hash stability)

**Porting Methodology (CRITICAL):**
- **Translation-first**: Port C# code line-by-line, preserving structure and naming
- **No behavior inference**: Never guess what C# does from test outputs
- **Tests validate, not guide**: Run tests only after implementation is complete
- **When stuck**: Read C# source, not test expectations
- **Naming convention**: Rust functions should mirror C# method names (e.g., `CreateComparisonUnitAtomList` → `create_comparison_unit_atom_list`)

---

## 1. Source Code Inventory

### Primary Comparer Files

| File | Lines | Complexity | Hash Algorithm | Key Algorithms |
|------|-------|------------|----------------|----------------|
| `WmlComparer.cs` | 8,835 | **Very High** | SHA1 | LCS, hierarchical grouping, revision tracking |
| `SmlComparer.cs` | 2,982 | High | SHA256 | Cell signatures, row alignment, sheet matching |
| `PmlComparer.cs` | 2,708 | High | SHA256 | Multi-pass matching, shape/slide alignment |
| **Total Comparers** | **14,525** | | | |

### Required Supporting Files (ALL must be ported)

| File | Lines | Purpose | Critical For |
|------|-------|---------|--------------|
| `PtOpenXmlUtil.cs` | 6,014 | Namespaces (W, S, P, MC), XML utilities, packaging helpers | All comparers |
| `PtUtil.cs` | 1,431 | SHA1 hashing, DescendantsTrimmed, GroupAdjacent, Rollup | All comparers |
| `PtOpenXmlDocument.cs` | 757 | WmlDocument, SmlDocument, PmlDocument wrappers | All comparers |
| `RevisionProcessor.cs` | ~800 | Accept/Reject revisions, element-level accept | WmlComparer |
| `MarkupSimplifier.cs` | ~600 | Markup preprocessing used by WmlComparer | WmlComparer |
| `XlsxTables.cs` | ~400 | Range parsing, cell address math | SmlComparer |
| `SmlDataRetriever.cs` | ~500 | Styles, colors, indexed color table | SmlComparer |
| `ColorParser.cs` | ~200 | System.Drawing.Color name mapping for Consolidate | WmlComparer |
| **Total Supporting** | **~10,700** | | |

### Grand Total: ~25,000 lines of C# to port

### Test Files (617+ test cases)

| Test File | Test Cases | Type | Data Source |
|-----------|------------|------|-------------|
| `WmlComparerTests.cs` | 268 | InlineData Theory | External .docx files |
| `WmlComparerTests2.cs` | 251 | InlineData Theory | External .docx files |
| `FormattingChangeTests.cs` | 20 | InlineData Theory | External .docx files |
| `SmlComparerTests.cs` | 51 | Fact + Theory | Programmatic creation |
| `PmlComparerTests.cs` | 27 | Fact + Theory | Auto-generated .pptx |
| `PmlComparerTestFileGenerator.cs` | N/A | Support | Generates test fixtures |
| **Total** | **617+** | | |

### Test Data: 771 files in `TestFiles/` directory

---

## 2. Critical Technical Challenges

### 2.1 XML DOM Mutability Problem

**Challenge:** The C# code uses `System.Xml.Linq` (`XElement`) which is fully mutable. WmlComparer extensively uses:
- `ReplaceWith()`, `AddAfterSelf()`, `AddBeforeSelf()`
- `Remove()`, `SetAttributeValue()`
- Parent/ancestor traversal while modifying

**Recommended Solution:**
- Use `roxmltree` for **reading** (fast, DOM-like, parent/children access)
- Use `quick-xml` for **writing** output
- Implement a **custom mutable tree** using `indextree` (arena-based) for modifications
- Alternative: Build on `xmltree` crate with custom extensions

```rust
// Arena-based mutable XML tree
pub struct XmlArena {
    arena: indextree::Arena<XmlNodeData>,
    root: Option<indextree::NodeId>,
}

pub enum XmlNodeData {
    Element {
        name: XName,
        attributes: Vec<XAttribute>,  // Vec preserves order!
    },
    Text(String),
    CData(String),
    Comment(String),
}
```

### 2.2 Attribute Order Preservation

**Critical:** Hash values depend on XML serialization order. C# LINQ-to-XML preserves insertion order.

**Rule:** Always use `Vec<XAttribute>`, never `HashMap`. Serialize attributes in Vec order.

```rust
// CORRECT - preserves order
pub struct XmlElement {
    pub name: XName,
    pub attributes: Vec<(String, String)>,  // Order preserved
    pub children: Vec<XmlNode>,
}

// WRONG - loses order, breaks hashes
pub attributes: HashMap<String, String>,  // DO NOT USE
```

### 2.3 Culture-Sensitive Case Folding

**Challenge:** `WmlComparer` uses `ToUpper(CultureInfo)` when `CaseInsensitive = true`.

**Solution:** Use ICU4X for locale-aware uppercasing:
```rust
use icu::casemap::CaseMapper;
use icu::locid::locale;

fn to_upper_culture(s: &str, culture: &str) -> String {
    let mapper = CaseMapper::new();
    let locale = culture.parse().unwrap_or(locale!("en"));
    mapper.uppercase_to_string(s, &locale)
}
```

### 2.4 Space Conflation

**Behavior:** When `ConflateBreakingAndNonbreakingSpaces = true`, WmlComparer treats space (0x20) and NBSP (0xA0) as identical.

```rust
fn normalize_spaces(s: &str, conflate: bool) -> String {
    if conflate {
        s.replace('\u{00A0}', " ")
    } else {
        s.to_string()
    }
}
```

### 2.5 Color Name Mapping

**Challenge:** `ColorParser.cs` maps color names to `System.Drawing.Color` values for Consolidate.

**Solution:** Embed the .NET known-color table:
```rust
pub fn parse_color(name: &str) -> Option<(u8, u8, u8)> {
    match name.to_lowercase().as_str() {
        "red" => Some((255, 0, 0)),
        "blue" => Some((0, 0, 255)),
        "green" => Some((0, 128, 0)),  // Note: .NET Green is 0,128,0 not 0,255,0
        // ... full table from System.Drawing.KnownColor
        _ => None
    }
}
```

---

## 3. Project Structure

```
redline-rs/
├── Cargo.toml                      # Workspace root
├── crates/
│   ├── redline-core/               # Pure Rust core library (WASM-compatible)
│   │   ├── Cargo.toml
│   │   └── src/
│   │       ├── lib.rs
│   │       ├── error.rs            # Error types with thiserror
│   │       ├── types.rs            # Core types: Document, Revision, etc.
│   │       │
│   │       ├── xml/
│   │       │   ├── mod.rs
│   │       │   ├── arena.rs        # Arena-based mutable XML tree
│   │       │   ├── parser.rs       # XML parsing (roxmltree + quick-xml)
│   │       │   ├── builder.rs      # XML serialization (order-preserving)
│   │       │   ├── node.rs         # XmlNode type (mirrors XElement)
│   │       │   ├── xname.rs        # XName = namespace + local name
│   │       │   └── namespaces.rs   # All OOXML namespace constants
│   │       │
│   │       ├── hash/
│   │       │   ├── mod.rs
│   │       │   ├── sha1.rs         # SHA1 for WmlComparer
│   │       │   └── sha256.rs       # SHA256 for Sml/PmlComparer
│   │       │
│   │       ├── package/
│   │       │   ├── mod.rs
│   │       │   ├── ooxml.rs        # OOXML package (zip wrapper)
│   │       │   ├── relationships.rs # .rels file parsing
│   │       │   ├── content_types.rs # [Content_Types].xml
│   │       │   └── parts.rs        # Part abstraction, GetXDocument/PutXDocument
│   │       │
│   │       ├── util/
│   │       │   ├── mod.rs
│   │       │   ├── descendants.rs  # DescendantsTrimmed
│   │       │   ├── group.rs        # GroupAdjacent, Rollup
│   │       │   ├── strings.rs      # StringConcatenate, MakeValidXml
│   │       │   └── culture.rs      # Culture-sensitive operations (ICU4X)
│   │       │
│   │       ├── wml/                # Word document comparison
│   │       │   ├── mod.rs
│   │       │   ├── comparer.rs     # Main Compare() function
│   │       │   ├── consolidate.rs  # Consolidate() function
│   │       │   ├── revisions.rs    # GetRevisions(), revision types
│   │       │   ├── preprocess.rs   # PreProcessMarkup()
│   │       │   ├── accept.rs       # AcceptRevisions (from RevisionProcessor)
│   │       │   ├── reject.rs       # RejectRevisions (from RevisionProcessor)
│   │       │   ├── simplify.rs     # MarkupSimplifier logic
│   │       │   ├── coalesce.rs     # Coalesce() tree reconstruction
│   │       │   ├── comparison_unit.rs # ComparisonUnit hierarchy
│   │       │   ├── lcs.rs          # LCS algorithm (exact C# port)
│   │       │   ├── correlation.rs  # CorrelatedSequence, correlation logic
│   │       │   ├── formatting.rs   # Formatting change detection, rPrChange
│   │       │   ├── color.rs        # ColorParser port
│   │       │   ├── settings.rs     # WmlComparerSettings
│   │       │   ├── document.rs     # WmlDocument wrapper
│   │       │   └── order.rs        # WmlOrderElementsPerStandard
│   │       │
│   │       ├── sml/                # Excel comparison
│   │       │   ├── mod.rs
│   │       │   ├── comparer.rs     # Main Compare(), ProduceMarkedWorkbook()
│   │       │   ├── canonicalize.rs # Canonicalize(), signature creation
│   │       │   ├── signatures.rs   # WorkbookSignature, CellSignature, etc.
│   │       │   ├── styles.rs       # ExpandStyle, CellFormatSignature
│   │       │   ├── diff.rs         # SmlDiffEngine
│   │       │   ├── markup.rs       # SmlMarkupRenderer, _DiffSummary sheet
│   │       │   ├── sheets.rs       # Sheet matching, rename detection
│   │       │   ├── cells.rs        # Cell comparison, numeric tolerance
│   │       │   ├── rows.rs         # Row alignment (LCS)
│   │       │   ├── columns.rs      # Column alignment
│   │       │   ├── tables.rs       # XlsxTables port (cell address math)
│   │       │   ├── data_retriever.rs # SmlDataRetriever port
│   │       │   ├── settings.rs     # SmlComparerSettings
│   │       │   ├── document.rs     # SmlDocument wrapper
│   │       │   └── types.rs        # SmlChange, SmlChangeType, result types
│   │       │
│   │       └── pml/                # PowerPoint comparison
│   │           ├── mod.rs
│   │           ├── comparer.rs     # Main Compare(), ProduceMarkedPresentation()
│   │           ├── canonicalize.rs # Canonicalize(), signature creation
│   │           ├── signatures.rs   # SlideSignature, ShapeSignature, etc.
│   │           ├── diff.rs         # PmlDiffEngine
│   │           ├── markup.rs       # PmlMarkupRenderer, summary slide
│   │           ├── slide_match.rs  # Multi-pass slide matching (LCS)
│   │           ├── shape_match.rs  # Shape matching (exact + fuzzy)
│   │           ├── transform.rs    # TransformSignature, tolerance matching
│   │           ├── settings.rs     # PmlComparerSettings
│   │           ├── document.rs     # PmlDocument wrapper
│   │           └── types.rs        # PmlChange, PmlChangeType, result types
│   │
│   ├── redline-wasm/               # WASM bindings
│   │   ├── Cargo.toml
│   │   └── src/
│   │       ├── lib.rs              # wasm-bindgen exports
│   │       ├── wml.rs              # Word comparison exports
│   │       ├── sml.rs              # Excel comparison exports
│   │       └── pml.rs              # PowerPoint comparison exports
│   │
│   ├── redline-tauri/              # Tauri plugin
│   │   ├── Cargo.toml
│   │   └── src/
│   │       └── lib.rs              # Tauri IPC commands
│   │
│   └── redline-cli/                # CLI tool for testing
│       ├── Cargo.toml
│       └── src/
│           └── main.rs
│
├── golden-generator/               # C# tool to generate golden outputs
│   ├── Program.cs
│   └── golden-generator.csproj
│
├── tests/
│   ├── common/
│   │   ├── mod.rs                  # Test utilities
│   │   ├── normalizer.rs           # XML normalization for comparison
│   │   └── validator.rs            # OpenXML structure validation
│   ├── golden/                     # Golden test files (from C#)
│   │   ├── manifest.json           # Test metadata
│   │   ├── wml/
│   │   ├── sml/
│   │   └── pml/
│   ├── wml_tests.rs
│   ├── sml_tests.rs
│   └── pml_tests.rs
│
├── benches/
│   └── comparison_bench.rs
│
└── scripts/
    ├── generate_golden.sh          # Generate golden files from C#
    ├── verify_parity.sh            # Verify Rust matches C#
    └── run_determinism_check.sh    # Cross-platform determinism test
```

---

## 4. Dependencies

### Core Dependencies

| C# / .NET | Rust Crate | Notes |
|-----------|------------|-------|
| `System.IO.Packaging` | `zip = "2.2"` | WASM-compatible, sync API |
| `System.Xml.Linq` (read) | `roxmltree = "0.20"` | Fast, DOM-like, read-only |
| `System.Xml.Linq` (write) | `quick-xml = "0.37"` | Fast writer, streaming |
| `System.Xml.Linq` (mutate) | `indextree = "4.7"` | Arena-based mutable tree |
| `System.Security.Cryptography.SHA1` | `sha1 = "0.10"` | RustCrypto, WASM-safe |
| `System.Security.Cryptography.SHA256` | `sha2 = "0.10"` | RustCrypto, WASM-safe |
| `System.Globalization.CultureInfo` | `icu = "1.5"` | ICU4X for locale-aware ops |
| `System.Text.Encoding.UTF8` | `std::str` | Native Rust |
| `System.Buffers.SearchValues<char>` | `memchr = "2.7"` | SIMD character search |
| `System.Collections.Frozen` | `once_cell = "1.19"` | Lazy static initialization |
| `System.Text.Json` | `serde_json = "1.0"` | JSON serialization |
| `System.Drawing.Color` | Custom table | Embedded known-color map |

### Cargo.toml (redline-core)

```toml
[package]
name = "redline-core"
version = "0.1.0"
edition = "2021"
license = "MIT"
description = "OOXML document comparison engine"

[dependencies]
# XML parsing and manipulation
roxmltree = "0.20"
quick-xml = { version = "0.37", features = ["serialize"] }
indextree = "4.7"

# ZIP handling (WASM-compatible)
zip = { version = "2.2", default-features = false, features = ["deflate"] }

# Cryptographic hashing
sha1 = "0.10"
sha2 = "0.10"
hex = "0.4"

# Internationalization (culture-sensitive operations)
icu = "1.5"
icu_provider = "1.5"

# Error handling
thiserror = "2.0"

# Serialization
serde = { version = "1.0", features = ["derive"] }
serde_json = "1.0"

# Fast character search (SIMD)
memchr = "2.7"

# Lazy initialization
once_cell = "1.19"

[dev-dependencies]
insta = { version = "1.41", features = ["json"] }
pretty_assertions = "1.4"
criterion = "0.5"
proptest = "1.4"

[features]
default = []
wasm = []  # Feature flag for WASM-specific code
```

### Cargo.toml (redline-wasm)

```toml
[package]
name = "redline-wasm"
version = "0.1.0"
edition = "2021"

[lib]
crate-type = ["cdylib"]

[dependencies]
redline-core = { path = "../redline-core", features = ["wasm"] }
wasm-bindgen = "0.2"
js-sys = "0.3"
serde-wasm-bindgen = "0.6"
console_error_panic_hook = "0.1"

[dependencies.web-sys]
version = "0.3"
features = ["console"]
```

---

## 5. Phased Implementation Plan

### Phase 0: Baseline Capture and Parity Specification

**Goal:** Establish the "golden truth" from C# before writing any Rust code.

**Deliverables:**
- [ ] Golden file outputs for all 617+ test cases
- [ ] Parity specification document
- [ ] Normalization spec for verification
- [ ] Locked API surface inventory

**Tasks:**

1. **Generate Golden Outputs:**
   ```bash
   dotnet run --project golden-generator -- \
       --filter 'WcTests|FormattingChange|SmlComparer|PmlComparer' \
       --output tests/golden/
   ```

   Output structure:
   ```
   tests/golden/
   ├── manifest.json              # All test metadata + expected counts
   ├── wml/
   │   ├── WC001/
   │   │   ├── input1.docx
   │   │   ├── input2.docx
   │   │   ├── output.docx
   │   │   ├── document.xml       # Extracted for diffing
   │   │   └── metadata.json      # Revision count, rPrChange count
   │   └── ...
   ├── sml/
   │   ├── SC001/
   │   │   ├── result.json        # SmlComparisonResult.ToJson() output
   │   │   └── marked.xlsx        # ProduceMarkedWorkbook output
   │   └── ...
   └── pml/
       └── ...
   ```

2. **Define Normalization Spec:**
   - Ignore `pt:Unid`, `rsidR*`, `rsidP*`, `rsidRDefault` when diffing
   - Ignore timestamps in revision attributes
   - **DO NOT normalize Rust output** — normalization is for verification only

3. **Lock API Surface:**
   ```
   WmlComparer.Compare(WmlDocument, WmlDocument, WmlComparerSettings) -> WmlDocument
   WmlComparer.Consolidate(WmlDocument, WmlRevisedDocumentInfo[], WmlComparerConsolidateSettings) -> WmlDocument
   WmlComparer.GetRevisions(WmlDocument, WmlComparerSettings) -> List<WmlComparerRevision>

   SmlComparer.Compare(SmlDocument, SmlDocument, SmlComparerSettings) -> SmlComparisonResult
   SmlComparer.ProduceMarkedWorkbook(SmlDocument, SmlDocument, SmlComparerSettings) -> SmlDocument
   SmlComparer.Canonicalize(SmlDocument, SmlComparerSettings) -> WorkbookSignature

   PmlComparer.Compare(PmlDocument, PmlDocument, PmlComparerSettings) -> PmlComparisonResult
   PmlComparer.ProduceMarkedPresentation(PmlDocument, PmlDocument, PmlComparerSettings) -> PmlDocument
   PmlComparer.Canonicalize(PmlDocument, PmlComparerSettings) -> PresentationSignature
   ```

**Parallelizable:** All tasks in Phase 0

---

### Phase 1: Core XML and OOXML Substrate

**Goal:** Build the foundational modules that all comparers depend on.

**Status:** ✅ COMPLETE (as of 2025-12-26)

The following modules are implemented and working:
- `xml/xml_document.rs` - Arena-based mutable XML tree with indextree
- `xml/namespaces.rs` - All OOXML namespace constants (W, S, P, A, R, MC, etc.)
- `package/` - OOXML package read/write via zip crate
- `hash/` - SHA1/SHA256 matching C# output
- `util/lcs.rs` - Generic LCS algorithm
- `wml/document.rs` - WmlDocument wrapper with part access
- `error.rs` - Error types with thiserror

**Deliverables:**
- [x] Core crate builds
- [x] XML DOM with mutation support
- [x] OOXML package layer
- [x] All utility functions
- [x] Stable hash outputs verified against C#

#### 1.1 Error Handling (`error.rs`) — 1 day
```rust
#[derive(Error, Debug)]
pub enum RedlineError {
    #[error("Invalid OOXML package: {message}")]
    InvalidPackage { message: String },

    #[error("Missing required part '{part_path}' in {document_type} document")]
    MissingPart { part_path: String, document_type: String },

    #[error("XML parsing error at {location}: {message}")]
    XmlParse { message: String, location: String },

    #[error("Invalid relationship: {message}")]
    InvalidRelationship { message: String },

    #[error("Unsupported feature: {feature}")]
    UnsupportedFeature { feature: String },

    #[error(transparent)]
    Io(#[from] std::io::Error),

    #[error(transparent)]
    Zip(#[from] zip::result::ZipError),
}

pub type Result<T> = std::result::Result<T, RedlineError>;
```

#### 1.2 XML Namespaces (`xml/namespaces.rs`) — 1 day

Port all namespace constants from `PtOpenXmlUtil.cs`:
```rust
pub mod W {
    pub const NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    pub fn p() -> XName { XName::new(NS, "p") }
    pub fn r() -> XName { XName::new(NS, "r") }
    pub fn t() -> XName { XName::new(NS, "t") }
    pub fn rPr() -> XName { XName::new(NS, "rPr") }
    pub fn pPr() -> XName { XName::new(NS, "pPr") }
    // ... 100+ more elements
}

pub mod S { /* spreadsheetml */ }
pub mod P { /* presentationml */ }
pub mod A { /* drawingml */ }
pub mod R { /* relationships */ }
pub mod MC { /* markup-compatibility */ }
pub mod CP { /* core properties */ }
pub mod DC { /* dublin core */ }
// ... ~50 more namespaces
```

#### 1.3 XML DOM (`xml/`) — 4 days

**Critical implementation with arena-based mutation:**

```rust
// xml/xname.rs
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct XName {
    pub namespace: Option<String>,
    pub local_name: String,
}

// xml/arena.rs
use indextree::{Arena, NodeId};

pub struct XmlDocument {
    arena: Arena<XmlNodeData>,
    root: Option<NodeId>,
}

pub enum XmlNodeData {
    Element {
        name: XName,
        attributes: Vec<XAttribute>,  // ORDER PRESERVED
    },
    Text(String),
    CData(String),
    Comment(String),
    ProcessingInstruction { target: String, data: String },
}

#[derive(Clone, Debug)]
pub struct XAttribute {
    pub name: XName,
    pub value: String,
}

impl XmlDocument {
    // Parsing
    pub fn parse(xml: &str) -> Result<Self>;
    pub fn parse_bytes(bytes: &[u8]) -> Result<Self>;

    // Traversal (C# XElement equivalents)
    pub fn root(&self) -> Option<XmlNodeRef>;
    pub fn descendants(&self, name: &XName) -> impl Iterator<Item = XmlNodeRef>;
    pub fn elements(&self, parent: NodeId, name: &XName) -> impl Iterator<Item = XmlNodeRef>;

    // Mutation (critical for WmlComparer)
    pub fn add_child(&mut self, parent: NodeId, node: XmlNodeData) -> NodeId;
    pub fn add_before(&mut self, sibling: NodeId, node: XmlNodeData) -> NodeId;
    pub fn add_after(&mut self, sibling: NodeId, node: XmlNodeData) -> NodeId;
    pub fn replace(&mut self, node: NodeId, new_node: XmlNodeData);
    pub fn remove(&mut self, node: NodeId);
    pub fn set_attribute(&mut self, node: NodeId, name: &XName, value: &str);
    pub fn remove_attribute(&mut self, node: NodeId, name: &XName);

    // Serialization (ORDER PRESERVING)
    pub fn to_string(&self) -> String;
    pub fn to_bytes(&self) -> Vec<u8>;
}
```

#### 1.4 OOXML Package (`package/`) — 2 days

```rust
pub struct OoxmlPackage {
    parts: HashMap<String, Vec<u8>>,
    content_types: ContentTypes,
    relationships: HashMap<String, Vec<Relationship>>,
}

impl OoxmlPackage {
    pub fn open(bytes: &[u8]) -> Result<Self>;
    pub fn save(&self) -> Result<Vec<u8>>;

    // Part access
    pub fn get_part(&self, path: &str) -> Option<&[u8]>;
    pub fn get_xml_part(&self, path: &str) -> Result<XmlDocument>;
    pub fn set_part(&mut self, path: &str, content: Vec<u8>);
    pub fn put_xml_part(&mut self, path: &str, doc: &XmlDocument) -> Result<()>;
    pub fn delete_part(&mut self, path: &str);

    // Relationships (C# GetXDocument/PutXDocument pattern)
    pub fn get_relationships(&self, source: &str) -> &[Relationship];
    pub fn add_relationship(&mut self, source: &str, rel: Relationship);

    // Content types
    pub fn get_content_type(&self, path: &str) -> Option<&str>;
    pub fn set_content_type(&mut self, path: &str, content_type: &str);
}

pub struct Relationship {
    pub id: String,
    pub rel_type: String,
    pub target: String,
    pub target_mode: TargetMode,
}
```

#### 1.5 Hash Utilities (`hash/`) — 1 day

```rust
// MUST match C# output exactly
pub fn sha1_hash_string(s: &str) -> String {
    use sha1::{Sha1, Digest};
    let mut hasher = Sha1::new();
    hasher.update(s.as_bytes());  // UTF-8 encoding
    hex::encode(hasher.finalize())
}

pub fn sha1_hash_bytes(bytes: &[u8]) -> String {
    use sha1::{Sha1, Digest};
    let mut hasher = Sha1::new();
    hasher.update(bytes);
    hex::encode(hasher.finalize())
}

pub fn sha256_hash_string(s: &str) -> String {
    use sha2::{Sha256, Digest};
    let mut hasher = Sha256::new();
    hasher.update(s.as_bytes());
    hex::encode(hasher.finalize())
}

#[cfg(test)]
mod tests {
    #[test]
    fn sha1_matches_csharp() {
        // Verified against C# PtUtils.SHA1HashStringForUTF8String
        assert_eq!(sha1_hash_string("test"), "a94a8fe5ccb19ba61c4c0873d391e987982fbbd3");
        assert_eq!(sha1_hash_string(""), "da39a3ee5e6b4b0d3255bfef95601890afd80709");
    }
}
```

#### 1.6 Utility Functions (`util/`) — 2 days

Port from `PtUtil.cs`:

```rust
// util/descendants.rs
/// Port of DescendantsTrimmed - stops at elements matching predicate
pub fn descendants_trimmed<'a>(
    doc: &'a XmlDocument,
    node: NodeId,
    trim_predicate: impl Fn(&XmlNodeData) -> bool + 'a,
) -> impl Iterator<Item = NodeId> + 'a

// util/group.rs
/// Port of GroupAdjacent
pub fn group_adjacent<T, K: Eq>(
    items: impl Iterator<Item = T>,
    key_selector: impl Fn(&T) -> K,
) -> Vec<Vec<T>>

/// Port of Rollup
pub fn rollup<T, R>(
    items: impl Iterator<Item = T>,
    seed: R,
    folder: impl Fn(R, &T) -> R,
) -> Vec<R>

// util/strings.rs
/// Port of MakeValidXml - ensures string is valid XML content
pub fn make_valid_xml(s: &str) -> String

/// Port of StringConcatenate
pub fn string_concatenate<I: Iterator<Item = S>, S: AsRef<str>>(
    items: I,
    separator: &str,
) -> String

// util/culture.rs
/// Culture-aware uppercase (uses ICU4X)
pub fn to_upper_invariant(s: &str) -> String
pub fn to_upper_culture(s: &str, culture: &str) -> String
```

**Phase 1 Verification:**
- [ ] All core modules compile with no warnings
- [ ] Hash outputs match C# exactly (verify with test vectors)
- [ ] XML round-trip: parse → serialize → parse is idempotent
- [ ] Attribute order preserved through round-trip
- [ ] OOXML package can read/write all test files

---

### Phase 2: WmlComparer (Critical Path)

**Goal:** Port the complete WmlComparer with 100% test compatibility.

**Deliverables:**
- [ ] All 268 WmlComparerTests pass
- [ ] All 251 WmlComparerTests2 pass
- [ ] All 20 FormattingChangeTests pass
- [ ] Output documents open correctly in Microsoft Word
- [ ] Revision counts match C# exactly

#### 2.1 Types and Settings — 1 day

```rust
pub struct WmlComparerSettings {
    /// Characters that separate words. MUST match C# defaults exactly.
    pub word_separators: Vec<char>,
    word_separators_search: memchr::memmem::Finder<'static>,

    pub author_for_revisions: Option<String>,
    pub date_time_for_revisions: String,
    pub detail_threshold: f64,  // Default: 0.15
    pub case_insensitive: bool,
    pub conflate_breaking_and_nonbreaking_spaces: bool,
    pub track_formatting_changes: bool,
    pub culture_info: Option<String>,
    pub log_callback: Option<Box<dyn Fn(&str)>>,
    pub starting_id_for_footnotes_endnotes: i32,
}

impl Default for WmlComparerSettings {
    fn default() -> Self {
        Self {
            // EXACT copy of C# defaults including Chinese punctuation
            word_separators: vec![
                ' ', '-', ')', '(', ';', ',',
                '（', '）', '，', '、', '，', '；', '。', '：', '的',
            ],
            author_for_revisions: None,
            date_time_for_revisions: chrono::Utc::now().format("%Y-%m-%dT%H:%M:%S%.3fZ").to_string(),
            detail_threshold: 0.15,
            case_insensitive: false,
            conflate_breaking_and_nonbreaking_spaces: true,
            track_formatting_changes: true,
            culture_info: None,
            log_callback: None,
            starting_id_for_footnotes_endnotes: 1,
            // ...
        }
    }
}
```

#### 2.2 Comparison Units (`wml/comparison_unit.rs`) — 2 days

Port the `ComparisonUnit` hierarchy exactly:

```rust
pub trait ComparisonUnit {
    fn sha1_hash(&self, settings: &WmlComparerSettings) -> &str;
    fn contents(&self) -> &[ComparisonUnitContent];
    fn comparison_unit_type(&self) -> ComparisonUnitType;
}

pub struct ComparisonUnitAtom {
    pub content_element: NodeId,
    pub ancestor_unids: Vec<String>,
    pub part: PartType,
    cached_sha1_hash: OnceCell<String>,
}

pub struct ComparisonUnitWord {
    pub atoms: Vec<ComparisonUnitAtom>,
    cached_sha1_hash: OnceCell<String>,
}

pub struct ComparisonUnitGroup {
    pub group_type: ComparisonUnitGroupType,
    pub contents: Vec<Box<dyn ComparisonUnit>>,
    cached_sha1_hash: OnceCell<String>,
}

pub enum ComparisonUnitGroupType {
    Paragraph,
    Table,
    Row,
    Cell,
    Textbox,
    // ... all types from C#
}
```

#### 2.3 RevisionProcessor Port (`wml/accept.rs`, `wml/reject.rs`) — 2 days

```rust
/// Accept all tracked revisions in the document
/// Port of RevisionProcessor.AcceptRevisions
pub fn accept_revisions(doc: &mut WmlDocument) -> Result<()>

/// Reject all tracked revisions in the document
/// Port of RevisionProcessor.RejectRevisions
pub fn reject_revisions(doc: &mut WmlDocument) -> Result<()>

/// Accept revisions for a specific element
/// Port of RevisionProcessor.AcceptRevisionsForElement
/// NOTE: Keep incomplete behavior exactly as C#
pub fn accept_revisions_for_element(element: NodeId, doc: &mut XmlDocument) -> Result<()>
```

#### 2.4 MarkupSimplifier Port (`wml/simplify.rs`) — 1 day

```rust
/// Simplify markup before comparison
/// Port of MarkupSimplifier calls used by WmlComparer
pub fn simplify_markup(doc: &mut WmlDocument, settings: &SimplifySettings) -> Result<()>
```

#### 2.5 Preprocessing (`wml/preprocess.rs`) — 2 days

```rust
/// Port of PreProcessMarkup() from WmlComparer.cs
/// Adds unid attributes to all elements for tree reconstruction
pub fn preprocess_markup(doc: &mut WmlDocument, starting_id: i32) -> Result<()>

/// Port of AssignUnidToAllElements
fn assign_unid_to_all_elements(root: NodeId, doc: &mut XmlDocument, counter: &mut i32)

/// Port of HashBlockLevelContent
fn hash_block_level_content(doc: &XmlDocument, settings: &WmlComparerSettings) -> HashMap<NodeId, String>
```

#### 2.6 LCS Algorithm (`wml/lcs.rs`) — 2 days

**CRITICAL: Must match C# implementation exactly.**

```rust
/// Exact port of C# Lcs() function from WmlComparer.cs:5779
pub fn lcs<T: ComparisonUnit>(
    source1: &[T],
    source2: &[T],
    settings: &WmlComparerSettings,
) -> Vec<CorrelatedSequence<T>>

pub struct CorrelatedSequence<T> {
    pub status: CorrelationStatus,
    pub source1: Vec<T>,
    pub source2: Vec<T>,
}

pub enum CorrelationStatus {
    Equal,
    Deleted,
    Inserted,
    Unknown,
}
```

#### 2.7 Core Comparison (`wml/comparer.rs`) — 4 days

```rust
/// Main entry point - exact port of WmlComparer.Compare()
pub fn compare(
    source1: &WmlDocument,
    source2: &WmlDocument,
    settings: &WmlComparerSettings,
) -> Result<WmlDocument>

/// Internal comparison with preprocessing control
fn compare_internal(
    source1: WmlDocument,
    source2: WmlDocument,
    settings: &WmlComparerSettings,
    preprocess_original: bool,
) -> Result<WmlDocument>

/// Create comparison unit atoms from document part
/// Port of CreateComparisonUnitAtomList at WmlComparer.cs:7950
fn create_comparison_unit_atom_list(
    part: &OoxmlPart,
    content_parent: NodeId,
    settings: &WmlComparerSettings,
) -> Vec<ComparisonUnitAtom>

/// Group atoms into words, then paragraphs, then higher structures
fn create_comparison_unit_groups(
    atoms: &[ComparisonUnitAtom],
) -> Vec<ComparisonUnitGroup>

/// Correlate two sequences using LCS
fn correlate_sequences(
    groups1: &[ComparisonUnitGroup],
    groups2: &[ComparisonUnitGroup],
    settings: &WmlComparerSettings,
) -> Vec<CorrelatedSequence>
```

#### 2.8 Formatting Changes (`wml/formatting.rs`) — 1 day

```rust
/// Detect formatting changes between runs
/// Produces rPrChange elements exactly matching C#
pub fn detect_formatting_changes(
    run1: NodeId,
    run2: NodeId,
    doc: &XmlDocument,
    settings: &WmlComparerSettings,
) -> Option<FormattingChange>

/// Normalize run properties for comparison
/// Port of NormalizedRPr logic
fn normalize_rpr(rpr: NodeId, doc: &XmlDocument) -> NormalizedRPr
```

#### 2.9 Tree Reconstruction (`wml/coalesce.rs`) — 2 days

```rust
/// Reconstruct XML tree from flat list of comparison units
/// Uses unid attributes to restore original hierarchy
/// Port of Coalesce() at WmlComparer.cs:7977
pub fn coalesce(comparison_units: &[ComparisonUnitAtom]) -> Result<XmlDocument>
```

#### 2.10 Element Ordering (`wml/order.rs`) — 1 day

```rust
/// Ensure output XML element order matches OOXML standard
/// Port of WmlOrderElementsPerStandard and Order_pPr dictionaries
pub fn order_elements_per_standard(element: NodeId, doc: &mut XmlDocument)

// Order dictionaries from PtOpenXmlUtil.cs
static ORDER_PPR: Lazy<Vec<XName>> = Lazy::new(|| vec![...]);
static ORDER_RPR: Lazy<Vec<XName>> = Lazy::new(|| vec![...]);
```

#### 2.11 Consolidation (`wml/consolidate.rs`) — 1 day

```rust
/// Consolidate multiple revised documents
pub fn consolidate(
    original: &WmlDocument,
    revised_docs: &[WmlRevisedDocumentInfo],
    settings: &WmlComparerConsolidateSettings,
) -> Result<WmlDocument>
```

---

### Phase 3: SmlComparer

**Goal:** Port Excel comparison with 100% test compatibility.

**Can run in parallel with Phase 4 after Phase 1 completes.**

**Deliverables:**
- [ ] All 51 SmlComparerTests pass
- [ ] Output workbooks open correctly in Microsoft Excel
- [ ] `SmlComparisonResult.ToJson()` output matches C# exactly

#### 3.1 Types and Settings — 1 day

```rust
pub struct SmlComparerSettings {
    pub case_sensitive: bool,
    pub numeric_tolerance: Option<f64>,
    pub compare_formatting: bool,
    pub enable_row_alignment: bool,
    pub enable_column_alignment: bool,
    pub sheet_rename_threshold: f64,  // Default: 0.7
}

pub struct SmlComparisonResult {
    pub changes: Vec<SmlChange>,
    pub statistics: SmlStatistics,
}

impl SmlComparisonResult {
    /// MUST produce identical JSON to C#
    pub fn to_json(&self) -> String
}
```

#### 3.2 Signatures (`sml/signatures.rs`) — 2 days

```rust
pub struct WorkbookSignature {
    pub sheets: Vec<WorksheetSignature>,
    pub content_hash: String,
}

pub struct WorksheetSignature {
    pub name: String,
    pub cells: HashMap<CellAddress, CellSignature>,
    pub named_ranges: Vec<NamedRangeSignature>,
    pub merged_cells: Vec<MergedCellSignature>,
    pub comments: Vec<CommentSignature>,
    pub data_validations: Vec<DataValidationSignature>,
    pub hyperlinks: Vec<HyperlinkSignature>,
}

pub struct CellSignature {
    pub address: CellAddress,
    pub value: Option<String>,
    pub formula: Option<String>,
    pub format: CellFormatSignature,
    pub content_hash: String,
}

/// Port of ExpandStyle from SmlDataRetriever.cs
pub struct CellFormatSignature {
    pub font: FontInfo,
    pub fill: FillInfo,
    pub border: BorderInfo,
    pub number_format: String,
    // ... all fields from C#
}

impl CellFormatSignature {
    /// MUST match C# IEquatable implementation
    fn equals(&self, other: &Self) -> bool

    /// For debugging - GetDifferenceDescription
    fn difference_description(&self, other: &Self) -> String
}
```

#### 3.3 Canonicalization (`sml/canonicalize.rs`) — 2 days

```rust
/// Create normalized workbook signature for comparison
/// Exact port of SmlCanonicalizer
pub fn canonicalize(
    doc: &SmlDocument,
    settings: &SmlComparerSettings,
) -> Result<WorkbookSignature>
```

#### 3.4 Cell Address Math (`sml/tables.rs`) — 1 day

Port from `XlsxTables.cs`:
```rust
pub struct CellAddress {
    pub column: u32,
    pub row: u32,
}

impl CellAddress {
    pub fn parse(s: &str) -> Result<Self>;  // "A1" -> (1, 1)
    pub fn to_string(&self) -> String;       // (1, 1) -> "A1"
}

pub fn column_to_letter(col: u32) -> String;  // 1 -> "A", 27 -> "AA"
pub fn letter_to_column(s: &str) -> u32;      // "A" -> 1, "AA" -> 27
```

#### 3.5 Diff Engine (`sml/diff.rs`) — 2 days

```rust
/// Port of SmlDiffEngine
pub fn diff(
    older: &WorkbookSignature,
    newer: &WorkbookSignature,
    settings: &SmlComparerSettings,
) -> SmlComparisonResult

/// Sheet matching with rename detection
fn match_sheets(
    older: &[WorksheetSignature],
    newer: &[WorksheetSignature],
    settings: &SmlComparerSettings,
) -> Vec<SheetMatch>

/// Row alignment using LCS
fn align_rows(
    older: &WorksheetSignature,
    newer: &WorksheetSignature,
    settings: &SmlComparerSettings,
) -> Vec<RowAlignment>
```

#### 3.6 Markup Renderer (`sml/markup.rs`) — 2 days

```rust
/// Port of SmlMarkupRenderer
pub fn produce_marked_workbook(
    older: &SmlDocument,
    newer: &SmlDocument,
    result: &SmlComparisonResult,
    settings: &SmlComparerSettings,
) -> Result<SmlDocument>

/// Add highlight styles to workbook
fn add_highlight_styles(doc: &mut SmlDocument) -> HighlightStyles

/// Create _DiffSummary sheet
fn create_summary_sheet(
    doc: &mut SmlDocument,
    result: &SmlComparisonResult,
)
```

---

### Phase 4: PmlComparer

**Goal:** Port PowerPoint comparison with 100% test compatibility.

**Can run in parallel with Phase 3 after Phase 1 completes.**

**Deliverables:**
- [ ] All 27 PmlComparerTests pass
- [ ] Output presentations open correctly in Microsoft PowerPoint
- [ ] `PmlComparisonResult.ToJson()` output matches C# exactly

#### 4.1 Types and Settings — 1 day

```rust
pub struct PmlComparerSettings {
    pub compare_slide_structure: bool,
    pub compare_shape_structure: bool,
    pub compare_text_content: bool,
    pub compare_text_formatting: bool,
    pub compare_shape_transforms: bool,
    pub use_slide_alignment_lcs: bool,  // Default: true
    pub transform_tolerance: TransformTolerance,
}

pub struct TransformTolerance {
    pub position: f64,
    pub size: f64,
    pub rotation: f64,
}

pub struct PmlComparisonResult {
    pub changes: Vec<PmlChange>,
    pub slides_inserted: i32,
    pub slides_deleted: i32,
    pub shapes_inserted: i32,
    pub shapes_deleted: i32,
    pub shapes_moved: i32,
    pub shapes_resized: i32,
    pub text_changes: i32,
}

impl PmlComparisonResult {
    pub fn to_json(&self) -> String;
    pub fn get_changes_by_slide(&self) -> HashMap<usize, Vec<&PmlChange>>;
    pub fn get_changes_by_type(&self) -> HashMap<PmlChangeType, Vec<&PmlChange>>;
}
```

#### 4.2 Signatures (`pml/signatures.rs`) — 2 days

```rust
pub struct PresentationSignature {
    pub slides: Vec<SlideSignature>,
}

pub struct SlideSignature {
    pub index: usize,
    pub title: Option<String>,
    pub shapes: Vec<ShapeSignature>,
    pub content_fingerprint: String,
    pub geometry_hash: String,
    pub image_hash: Option<String>,
    pub table_hash: Option<String>,
    pub chart_hash: Option<String>,
}

pub struct ShapeSignature {
    pub id: String,
    pub name: Option<String>,
    pub placeholder_type: Option<String>,
    pub transform: TransformSignature,
    pub text_body: Option<TextBodySignature>,
    pub content_hash: String,
}

pub struct TransformSignature {
    pub offset_x: i64,
    pub offset_y: i64,
    pub extent_cx: i64,
    pub extent_cy: i64,
    pub rotation: Option<i32>,
}

impl TransformSignature {
    /// Check if transforms match within tolerance
    pub fn matches_with_tolerance(&self, other: &Self, tolerance: &TransformTolerance) -> bool
}
```

#### 4.3 Multi-pass Slide Matching (`pml/slide_match.rs`) — 2 days

```rust
/// Match slides using multiple passes (C# algorithm exactly)
/// Port of PmlSlideMatchEngine
pub fn match_slides(
    older: &[SlideSignature],
    newer: &[SlideSignature],
    settings: &PmlComparerSettings,
) -> Vec<SlideMatch>

// Pass 1: Match by title text (exact match)
fn match_by_title(older: &[SlideSignature], newer: &[SlideSignature]) -> Vec<SlideMatch>

// Pass 2: Match by content fingerprint (hash match)
fn match_by_fingerprint(older: &[SlideSignature], newer: &[SlideSignature]) -> Vec<SlideMatch>

// Pass 3: Match by LCS (if enabled)
fn match_by_lcs(older: &[SlideSignature], newer: &[SlideSignature]) -> Vec<SlideMatch>

// Pass 4: Match by position
fn match_by_position(older: &[SlideSignature], newer: &[SlideSignature]) -> Vec<SlideMatch>

// Pass 5: Fuzzy matching
fn match_fuzzy(older: &[SlideSignature], newer: &[SlideSignature]) -> Vec<SlideMatch>
```

#### 4.4 Shape Matching (`pml/shape_match.rs`) — 2 days

```rust
/// Match shapes within matched slides
/// Port of PmlShapeMatchEngine
pub fn match_shapes(
    older: &[ShapeSignature],
    newer: &[ShapeSignature],
    settings: &PmlComparerSettings,
) -> Vec<ShapeMatch>

// Match by placeholder type
fn match_by_placeholder(older: &[ShapeSignature], newer: &[ShapeSignature]) -> Vec<ShapeMatch>

// Match by name and type
fn match_by_name_and_type(older: &[ShapeSignature], newer: &[ShapeSignature]) -> Vec<ShapeMatch>

// Fuzzy similarity matching
fn match_fuzzy(older: &[ShapeSignature], newer: &[ShapeSignature]) -> Vec<ShapeMatch>
```

#### 4.5 Diff Engine and Markup (`pml/diff.rs`, `pml/markup.rs`) — 2 days

```rust
/// Port of PmlDiffEngine
pub fn diff(
    older: &PresentationSignature,
    newer: &PresentationSignature,
    settings: &PmlComparerSettings,
) -> PmlComparisonResult

/// Port of PmlMarkupRenderer
pub fn produce_marked_presentation(
    older: &PmlDocument,
    newer: &PmlDocument,
    result: &PmlComparisonResult,
    settings: &PmlComparerSettings,
) -> Result<PmlDocument>

/// Create summary slide
fn create_summary_slide(
    doc: &mut PmlDocument,
    result: &PmlComparisonResult,
)
```

---

### Phase 5: Test Port and Parity Harness

**Goal:** Rust tests mirror C# tests and all pass.

**Deliverables:**
- [ ] All 617+ test cases ported to Rust
- [ ] Golden file comparison infrastructure
- [ ] OpenXML validation on Rust outputs
- [ ] CI passes on all platforms

#### 5.1 Test Infrastructure

```rust
// tests/common/mod.rs

/// Load golden file metadata
pub fn load_golden_metadata(test_id: &str) -> TestMetadata;

/// Compare Rust output to golden output
pub fn verify_against_golden(
    output: &[u8],
    test_id: &str,
    normalizer: &Normalizer,
) -> Result<()>;

/// Validate OOXML structure
pub fn validate_ooxml(doc_bytes: &[u8]) -> ValidationResult;
```

#### 5.2 Normalization for Comparison Only

```rust
// tests/common/normalizer.rs

/// Normalize XML for comparison (NOT for output!)
pub struct Normalizer {
    ignore_attributes: HashSet<String>,
    ignore_elements: HashSet<String>,
}

impl Normalizer {
    pub fn new() -> Self {
        Self {
            ignore_attributes: hashset![
                "pt:Unid",
                "w:rsidR", "w:rsidRPr", "w:rsidP", "w:rsidRDefault",
            ],
            ignore_elements: hashset![],
        }
    }

    pub fn normalize(&self, xml: &str) -> String;
}
```

#### 5.3 Test Macro Pattern

```rust
// tests/wml_tests.rs

macro_rules! golden_test {
    ($name:ident, $test_id:expr, $expected_revisions:expr) => {
        #[test]
        fn $name() {
            let input1 = include_bytes!(concat!("golden/wml/", $test_id, "/input1.docx"));
            let input2 = include_bytes!(concat!("golden/wml/", $test_id, "/input2.docx"));

            let doc1 = WmlDocument::from_bytes(input1).unwrap();
            let doc2 = WmlDocument::from_bytes(input2).unwrap();
            let settings = WmlComparerSettings::default();

            let result = wml::compare(&doc1, &doc2, &settings).unwrap();

            // Verify revision count
            let revisions = wml::get_revisions(&result, &settings).unwrap();
            assert_eq!(revisions.len(), $expected_revisions);

            // Verify against golden output
            let normalizer = Normalizer::new();
            verify_against_golden(&result.to_bytes(), $test_id, &normalizer).unwrap();

            // Validate OOXML structure
            let validation = validate_ooxml(&result.to_bytes());
            assert!(validation.is_valid(), "Validation errors: {:?}", validation.errors);
        }
    };
}

// Generate from C# InlineData
golden_test!(wc001_digits, "WC001", 1);
golden_test!(wc002_paragraph_mod, "WC002", 2);
// ... 600+ more
```

---

### Phase 6: WASM and Tauri Integration

**Goal:** Create platform bindings for browser and desktop.

**Deliverables:**
- [ ] WASM builds and runs in browser
- [ ] WASM builds and runs in Node.js
- [ ] Tauri plugin works
- [ ] Stable JSON schema for python-wasmer

#### 6.1 WASM Exports

```rust
// crates/redline-wasm/src/lib.rs

use wasm_bindgen::prelude::*;

#[wasm_bindgen(start)]
pub fn init() {
    console_error_panic_hook::set_once();
}

#[wasm_bindgen]
pub fn compare_word_documents(
    older: &[u8],
    newer: &[u8],
    settings_json: &str,
) -> Result<Vec<u8>, JsError> {
    let settings: WmlComparerSettings = serde_json::from_str(settings_json)
        .map_err(|e| JsError::new(&e.to_string()))?;
    let older_doc = WmlDocument::from_bytes(older)
        .map_err(|e| JsError::new(&e.to_string()))?;
    let newer_doc = WmlDocument::from_bytes(newer)
        .map_err(|e| JsError::new(&e.to_string()))?;

    let result = wml::compare(&older_doc, &newer_doc, &settings)
        .map_err(|e| JsError::new(&e.to_string()))?;

    result.to_bytes().map_err(|e| JsError::new(&e.to_string()))
}

#[wasm_bindgen]
pub fn compare_spreadsheets(
    older: &[u8],
    newer: &[u8],
    settings_json: &str,
) -> Result<JsValue, JsError> {
    // ... returns JSON result
}

#[wasm_bindgen]
pub fn compare_presentations(
    older: &[u8],
    newer: &[u8],
    settings_json: &str,
) -> Result<JsValue, JsError> {
    // ... returns JSON result
}
```

#### 6.2 Tauri Plugin

```rust
// crates/redline-tauri/src/lib.rs

use tauri::plugin::{Builder, TauriPlugin};

#[tauri::command]
async fn compare_word(
    older_path: String,
    newer_path: String,
    output_path: String,
    settings: WmlComparerSettings,
) -> Result<(), String> {
    let older = tokio::fs::read(&older_path).await.map_err(|e| e.to_string())?;
    let newer = tokio::fs::read(&newer_path).await.map_err(|e| e.to_string())?;

    let older_doc = WmlDocument::from_bytes(&older).map_err(|e| e.to_string())?;
    let newer_doc = WmlDocument::from_bytes(&newer).map_err(|e| e.to_string())?;

    let result = wml::compare(&older_doc, &newer_doc, &settings).map_err(|e| e.to_string())?;

    tokio::fs::write(&output_path, result.to_bytes().map_err(|e| e.to_string())?).await
        .map_err(|e| e.to_string())?;

    Ok(())
}

pub fn init<R: tauri::Runtime>() -> TauriPlugin<R> {
    Builder::new("redline")
        .invoke_handler(tauri::generate_handler![
            compare_word,
            compare_spreadsheet,
            compare_presentation,
        ])
        .build()
}
```

#### 6.3 Python-Wasmer Interface

```python
# redline_py/redline.py

from wasmer import Store, Module, Instance, Memory
from wasmer_compiler_cranelift import Compiler
import json

class RedlineComparer:
    def __init__(self, wasm_path: str = None):
        store = Store(Compiler)
        if wasm_path is None:
            import pkg_resources
            wasm_path = pkg_resources.resource_filename(__name__, 'redline.wasm')

        with open(wasm_path, 'rb') as f:
            module = Module(store, f.read())

        self.instance = Instance(module)
        self.memory = self.instance.exports.memory

    def compare_word(
        self,
        older: bytes,
        newer: bytes,
        settings: dict = None
    ) -> bytes:
        settings_json = json.dumps(settings or {}).encode('utf-8')

        # Allocate memory
        older_ptr = self._allocate(older)
        newer_ptr = self._allocate(newer)
        settings_ptr = self._allocate(settings_json)

        # Call WASM
        result_ptr = self.instance.exports.compare_word_documents(
            older_ptr, len(older),
            newer_ptr, len(newer),
            settings_ptr, len(settings_json)
        )

        # Read result
        return self._read_result(result_ptr)

    def _allocate(self, data: bytes) -> int:
        ptr = self.instance.exports.alloc(len(data))
        mem_view = self.memory.uint8_view(ptr)
        for i, b in enumerate(data):
            mem_view[i] = b
        return ptr

    def _read_result(self, ptr: int) -> bytes:
        # Read length prefix, then data
        mem_view = self.memory.uint8_view(ptr)
        length = int.from_bytes(bytes(mem_view[0:4]), 'little')
        return bytes(mem_view[4:4+length])
```

---

### Phase 7: Determinism and Performance Hardening

**Goal:** Ensure identical output across all platforms.

**Deliverables:**
- [ ] Cross-platform determinism verified (macOS, Linux, Windows, WASM)
- [ ] Performance within 2x of C#
- [ ] Memory usage acceptable for large documents

#### 7.1 Determinism Tests

```rust
#[test]
fn determinism_across_runs() {
    let input1 = include_bytes!("fixtures/large_doc_1.docx");
    let input2 = include_bytes!("fixtures/large_doc_2.docx");

    let results: Vec<Vec<u8>> = (0..10)
        .map(|_| {
            let doc1 = WmlDocument::from_bytes(input1).unwrap();
            let doc2 = WmlDocument::from_bytes(input2).unwrap();
            wml::compare(&doc1, &doc2, &Default::default())
                .unwrap()
                .to_bytes()
                .unwrap()
        })
        .collect();

    // All runs must produce identical output
    for i in 1..results.len() {
        assert_eq!(results[0], results[i], "Run {} differs from run 0", i);
    }
}
```

#### 7.2 Performance Benchmarks

```rust
// benches/comparison_bench.rs

use criterion::{criterion_group, criterion_main, Criterion};

fn wml_comparison_benchmark(c: &mut Criterion) {
    let input1 = include_bytes!("../tests/fixtures/medium_doc_1.docx");
    let input2 = include_bytes!("../tests/fixtures/medium_doc_2.docx");

    c.bench_function("wml_compare_medium", |b| {
        b.iter(|| {
            let doc1 = WmlDocument::from_bytes(input1).unwrap();
            let doc2 = WmlDocument::from_bytes(input2).unwrap();
            wml::compare(&doc1, &doc2, &Default::default()).unwrap()
        })
    });
}

criterion_group!(benches, wml_comparison_benchmark);
criterion_main!(benches);
```

---

### Phase 8: Documentation and Release

**Goal:** Complete documentation and prepare for release.

**Deliverables:**
- [ ] rustdoc for all public APIs
- [ ] README with usage examples
- [ ] CHANGELOG
- [ ] Migration notes for existing users
- [ ] Published to crates.io

---

## 6. Parallelization and Dependencies

```
Phase 0 ──────────────────────────────────────────────────────► (All parallel)
         │
         ▼
Phase 1 ──────────────────────────────────────────────────────► (Sequential)
         │
         ├─────────────────────────────────────────────────────► Phase 2 (WML)
         │                                                        │
         ├─────────────────────────────────────────────────────► Phase 3 (SML) ◄─── Parallel
         │                                                        │
         └─────────────────────────────────────────────────────► Phase 4 (PML) ◄─── Parallel
                                                                  │
Phase 5 (Tests) ◄─────────────────────────────────────────────────┘
         │
         ├─────────────────────────────────────────────────────► Phase 6 (WASM)  ◄─── Parallel
         │                                                        │
         └─────────────────────────────────────────────────────► Phase 6 (Tauri) ◄─── Parallel
                                                                  │
Phase 7 (Hardening) ◄─────────────────────────────────────────────┘
         │
         ▼
Phase 8 (Release) ────────────────────────────────────────────────►
```

### Work Tracks (After Phase 1)

| Track | Focus | Dependencies |
|-------|-------|--------------|
| **Track A** | WML: RevisionProcessor, WmlComparer, FormattingChange | Phase 1 |
| **Track B** | SML: Canonicalizer, DiffEngine, MarkupRenderer | Phase 1 |
| **Track C** | PML: Canonicalizer, MatchEngines, MarkupRenderer | Phase 1 |
| **Track D** | Tests: Harness, validator, golden diffs | Phase 0 |
| **Track E** | Bindings: WASM, Tauri, Python | Phases 2-4 |

### Timeline Estimates

| Configuration | Duration |
|---------------|----------|
| 1 developer | 35-45 working days |
| 2 developers | 22-28 working days (A+B or A+C parallel) |
| 3 developers | 18-24 working days (A+B+C parallel) |
| 4 developers | 16-20 working days (A+B+C+D parallel) |

---

## 7. Risk Register

| Risk | Likelihood | Impact | Mitigation |
|------|------------|--------|------------|
| XML attribute order changes break hashes | High | Critical | Use Vec for attributes, never HashMap. Extensive hash verification tests. |
| Culture-sensitive case folding differs | Medium | High | Use ICU4X. Test with known C# culture outputs. |
| LCS algorithm produces different correlations | Medium | Critical | Port line-by-line. Test with identical inputs from C#. |
| WASM bundle too large | Medium | Medium | Use wasm-opt -O4, enable LTO, tree-shake unused code. |
| Memory leaks in WASM | Low | Medium | Use miri for testing. Explicit memory management tests. |
| Large document performance regression | Medium | Medium | Continuous benchmarking. Profile with flame graphs. |
| System.Drawing.Color mapping incomplete | Low | Low | Embed full .NET KnownColor table. |
| OpenXML part ordering differs | Medium | High | Explicit relationship handling matching C# logic. |
| Floating-point precision differs | Low | Medium | Use same rounding as C#. Test numeric edge cases. |

---

## 8. Milestone Gates

| Gate | Criteria | Blocking |
|------|----------|----------|
| **Gate 1** | XML/OOXML substrate complete. Hash verification passes. Attribute order preserved. | Phase 2-4 |
| **Gate 2** | WML parity complete. All 539 WML tests pass. | Phase 5 (WML tests) |
| **Gate 3** | SML parity complete. All 51 SML tests pass. JSON output matches. | Phase 5 (SML tests) |
| **Gate 4** | PML parity complete. All 27 PML tests pass. JSON output matches. | Phase 5 (PML tests) |
| **Gate 5** | WASM + Tauri integration complete. Cross-platform tests pass. | Phase 7 |
| **Gate 6** | Determinism verified. Performance acceptable. | Phase 8 |

---

## 9. Acceptance Criteria

1. **All 617+ Rust test equivalents pass**
2. **JSON outputs from `ToJson()` methods match C# exactly** (field order, formatting)
3. **WML output passes OpenXML validation** with same error whitelist as C#
4. **Marked workbooks and presentations open successfully** in Office applications
5. **Cross-platform determinism verified** (macOS, Linux, Windows, WASM)
6. **Performance within 2x of C#** on benchmark suite
7. **WASM bundle size < 2MB gzipped**
8. **No unsafe code** (or justified and audited)

---

## Appendix A: C# to Rust Type Mapping

| C# Type | Rust Type | Notes |
|---------|-----------|-------|
| `string` | `String` | UTF-8 in Rust vs UTF-16 in C# |
| `char` | `char` | Both are Unicode code points |
| `int` | `i32` | |
| `long` | `i64` | |
| `double` | `f64` | |
| `bool` | `bool` | |
| `byte[]` | `Vec<u8>` | |
| `List<T>` | `Vec<T>` | |
| `Dictionary<K, V>` | `HashMap<K, V>` | **Beware:** HashMap loses order |
| `HashSet<T>` | `HashSet<T>` | |
| `XElement` | `XmlNode` (custom) | Arena-based for mutation |
| `XDocument` | `XmlDocument` (custom) | |
| `XName` | `XName` (custom) | namespace + local_name |
| `XAttribute` | `XAttribute` (custom) | |
| `MemoryStream` | `Cursor<Vec<u8>>` | |
| `Func<T, R>` | `fn(T) -> R` or `Box<dyn Fn(T) -> R>` | |
| `Action<T>` | `fn(T)` or `Box<dyn Fn(T)>` | |
| `IEnumerable<T>` | `impl Iterator<Item = T>` | |
| `Lazy<T>` | `OnceCell<T>` | |
| `FrozenSet<T>` | `HashSet<T>` + `once_cell` | |
| `SearchValues<char>` | `memchr` | |
| `CultureInfo` | ICU4X `Locale` | |
| `System.Drawing.Color` | `(u8, u8, u8)` + lookup table | |

---

## Appendix B: Key Algorithm Locations in C#

| Algorithm | File:Line | Rust Module |
|-----------|-----------|-------------|
| LCS Core | `WmlComparer.cs:5779` | `wml/lcs.rs` |
| SHA1 Hash | `WmlComparer.cs:8771` | `hash/sha1.rs` |
| SHA256 Hash | `SmlComparer.cs`, `PmlComparer.cs` | `hash/sha256.rs` |
| Comparison Unit Creation | `WmlComparer.cs:7950` | `wml/comparison_unit.rs` |
| Coalesce | `WmlComparer.cs:7977` | `wml/coalesce.rs` |
| PreProcessMarkup | `WmlComparer.cs:~200` | `wml/preprocess.rs` |
| AcceptRevisions | `RevisionProcessor.cs` | `wml/accept.rs` |
| Element Ordering | `PtOpenXmlUtil.cs` (Order_pPr, etc.) | `wml/order.rs` |
| Slide Matching | `PmlComparer.cs` | `pml/slide_match.rs` |
| Cell Signature | `SmlComparer.cs` | `sml/signatures.rs` |

---

## Appendix C: Attribute Order Preservation Checklist

When porting any XML-handling code, verify:

- [ ] Attributes stored in `Vec`, not `HashMap`
- [ ] Serialization iterates Vec in order
- [ ] No sorting of attributes anywhere
- [ ] Namespace declarations preserved in order
- [ ] Child elements maintain insertion order
- [ ] Hash inputs use consistent serialization

---

*Document Version: 1.0 (Synthesis)*
*Sources: RUST_MIGRATION_PLAN_CC.md, RUST_MIGRATION_PLAN_FROM_CS_GMI.md, RUST_MIGRATION_PLAN_FROM_CS_CODEX.md*
*Target: 100% test compatibility with C# implementation*
