# redline-rs

Faithful Rust port of the C# OpenXML document comparers from OpenXmlPowerTools.

## Overview

This library provides byte-for-byte compatible document comparison for:
- **Word documents** (.docx) - WmlComparer
- **Excel workbooks** (.xlsx) - SmlComparer
- **PowerPoint presentations** (.pptx) - PmlComparer

## Target Platforms

- Native Rust (desktop applications, servers)
- WebAssembly (browser and Node.js)
- Tauri desktop applications
- Python via wasmer bindings

## CLI Usage

### Installation

```bash
# Build from source
cargo build --release

# The binary will be at target/release/redline
```

### Compare Documents

```bash
# Basic comparison - output filename auto-generated
redline compare original.docx modified.docx

# Specify output filename
redline compare original.docx modified.docx --output result.docx
redline compare original.docx modified.docx -o result.docx

# Get JSON output
redline compare original.docx modified.docx --json
```

### Auto-Generated Output Filename

When no `--output` is specified, the tool generates a filename with the format:

```
redline-YYYYMMDD-HHMMSS-COMMIT.docx
```

- `YYYYMMDD-HHMMSS`: Current date and time
- `COMMIT`: 8-character git commit hash of the redline binary

Example: `redline-20251231-104732-d25823b3.docx`

This makes it easy to track which version of the tool produced each output.

### Output Format

Default output (human-readable):
```
Comparison complete:
  Insertions: 42
  Deletions:  38
  Total:      80
  Output:     redline-20251231-104732-d25823b3.docx
  Commit:     d25823b3
```

JSON output (`--json` flag):
```json
{"insertions":42,"deletions":38,"total":80,"output":"redline-20251231-104732-d25823b3.docx","commit":"d25823b3"}
```

### Count Revisions (Without Generating Output)

```bash
# Count changes without writing output file
redline count original.docx modified.docx

# JSON output
redline count original.docx modified.docx --json
```

### Document Info

```bash
redline info document.docx
```

## Command Reference

### `compare`

Compare two documents and generate a redlined output.

```
redline compare <DOC1> <DOC2> [OPTIONS]

Arguments:
  <DOC1>  Original document (before changes)
  <DOC2>  Modified document (after changes)

Options:
  -o, --output <PATH>       Output document path (default: auto-generated)
  -t, --doc-type <TYPE>     Document type: auto, docx, xlsx, pptx [default: auto]
      --json                Output revision statistics as JSON
      --author <NAME>       Author name for revisions (default: doc2's LastModifiedBy or Creator)
      --date <ISO8601>      Date/time for revisions (default: doc2's modified date)
      --detail-threshold <FLOAT>  Comparison granularity (0.0-1.0) [default: 0.15]
```

### `count`

Count revisions between two documents without generating output.

```
redline count <DOC1> <DOC2> [OPTIONS]

Arguments:
  <DOC1>  Original document (before changes)
  <DOC2>  Modified document (after changes)

Options:
      --json    Output as JSON
```

### `info`

Display information about a document.

```
redline info <FILE>

Arguments:
  <FILE>  Document to analyze
```

## Comparison Granularity

The `--detail-threshold` option controls how granular the comparison is:

- **Higher values (e.g., 0.15)**: More coalesced, paragraph-level changes. Entire paragraphs may be shown as deleted/inserted when they differ significantly.
- **Lower values (e.g., 0.05 or 0.01)**: More granular, word-level changes. Individual words and phrases are tracked even when paragraphs are substantially different.

```bash
# Default comparison (threshold 0.15)
redline compare doc1.docx doc2.docx

# More detailed word-level comparison
redline compare doc1.docx doc2.docx --detail-threshold 0.05

# Maximum granularity
redline compare doc1.docx doc2.docx --detail-threshold 0.01
```

The threshold represents the minimum ratio of matching content required before the algorithm attempts word-level comparison within a paragraph. Lower thresholds produce more changes but may find spurious matches on common words.

## LCS Algorithm Tracing

For debugging the comparison algorithm, redline-rs includes optional tracing capabilities that capture detailed information about the Longest Common Subsequence (LCS) algorithm execution.

### Trace Options

```bash
# Trace a specific section
redline compare original.docx modified.docx \
  --trace-section "4.1" \
  --trace-output lcs-trace.json

# Trace paragraphs by text prefix
redline compare original.docx modified.docx \
  --trace-paragraph "The parties agree" \
  --trace-output lcs-trace.json
```

Options:
- `--trace-section <SECTION>`: Trace LCS for a specific section (e.g., "3.1", "(b)")
- `--trace-paragraph <PREFIX>`: Trace LCS for paragraphs starting with this text
- `--trace-output <PATH>`: Output file for trace JSON [default: lcs-trace.json]

### Trace Output Format

The trace JSON contains:
- `section_identifier`: The matched section
- `paragraph_text`: The paragraph text being traced
- `left_tokens` / `right_tokens`: Token sequences being compared
- `operations`: LCS algorithm operations (coalesced for readability)
- `final_correlations`: The resulting alignment

### Zero-Overhead Tracing

Tracing is implemented as a compile-time feature. When disabled, all tracing code is completely eliminated from the binary with zero runtime overhead.

```bash
# Build without trace support (smaller binary, no tracing overhead)
cargo build --release --no-default-features

# Build with trace support (default)
cargo build --release
```

## Revision Attribution

The tool automatically extracts metadata from the modified document (doc2) to attribute revisions:

1. **Author**: Uses `--author` flag if provided, otherwise falls back to:
   - `lastModifiedBy` from document properties
   - `creator` from document properties
   - "Redline" as final fallback

2. **Date**: Uses `--date` flag if provided (ISO 8601 format), otherwise:
   - `modified` date from document properties
   - Current time as final fallback

## Exit Codes

- `0`: Success
- `1`: Error (file not found, parse error, comparison failed, etc.)

## Library Usage

The comparison engine is available as a library crate (`redline-core`) for integration into other Rust projects:

```rust
use redline_core::{WmlDocument, WmlComparer, WmlComparerSettings};

let doc1 = WmlDocument::from_bytes(&bytes1)?;
let doc2 = WmlDocument::from_bytes(&bytes2)?;

let settings = WmlComparerSettings::default();
let result = WmlComparer::compare(&doc1, &doc2, Some(&settings))?;

// result.document contains the redlined DOCX bytes
// result.insertions and result.deletions contain counts
std::fs::write("output.docx", &result.document)?;
```

## Project Structure

```
redline-rs/
├── crates/
│   ├── redline-core/     # Pure Rust core library
│   ├── redline-wasm/     # WASM bindings
│   ├── redline-tauri/    # Tauri plugin
│   └── redline-cli/      # Command-line tool
├── tests/
│   ├── common/           # Test utilities
│   └── golden/           # Golden test files
└── benches/              # Performance benchmarks
```

## Supported Document Types

| Type | Extension | Status |
|------|-----------|--------|
| Word | `.docx` | Supported |
| Excel | `.xlsx` | Not yet implemented |
| PowerPoint | `.pptx` | Not yet implemented |

## Building

```bash
cargo build
cargo test
cargo clippy
```

## Status

This is a work-in-progress port. See RUST_MIGRATION_PLAN_SYNTHESIS.md for the detailed implementation plan.

### Phase Status

- [x] Phase 0: Baseline Capture and Project Setup
- [ ] Phase 1: Core XML and OOXML Substrate
- [ ] Phase 2: WmlComparer (Word)
- [ ] Phase 3: SmlComparer (Excel)
- [ ] Phase 4: PmlComparer (PowerPoint)
- [ ] Phase 5: Test Port and Parity Harness
- [ ] Phase 6: WASM and Tauri Integration
- [ ] Phase 7: Performance and Determinism Hardening
- [ ] Phase 8: Release Readiness

## License

MIT
