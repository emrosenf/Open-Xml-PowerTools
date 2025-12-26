# Consolidation Function Integration Report

## Task Summary

**Objective**: Integrate the `coalesce_adjacent_runs_with_identical_formatting` function into the WmlComparer pipeline.

**Expected Outcome**: The 48 failing WC-* tests should show improvement because revisions are being consolidated properly.

## What Was Implemented

### 1. Added `coalesce_document()` Function

Created a top-level wrapper function in `coalesce.rs` (line 696) that:
- Processes all paragraphs in a document (excluding those in txbxContent)
- Processes txbxContent paragraphs recursively
- Matches the C# `CoalesceAdjacentRunsWithIdenticalFormatting(XDocument)` signature at line 2336

```rust
pub fn coalesce_document(doc: &mut XmlDocument, doc_root: NodeId) {
    // Process main document paragraphs (excluding those in txbxContent)
    let paras: Vec<NodeId> = descendants_trimmed(doc, doc_root, |d| {
        d.name().map(|n| n == &W::txbxContent()).unwrap_or(false)
    })
    .filter(|&node| {
        doc.get(node)
            .and_then(|d| d.name())
            .map(|n| n == &W::p())
            .unwrap_or(false)
    })
    .collect();
    
    for para in paras {
        coalesce_adjacent_runs_with_identical_formatting(doc, para);
    }
    
    // Process txbxContent paragraphs recursively
    // ...
}
```

### 2. Exported Function

Added export in `mod.rs` line 40:
```rust
pub use coalesce::{
    produce_markup_from_atoms, mark_content_as_deleted_or_inserted, 
    reset_coalesce_revision_id, coalesce_adjacent_runs_with_identical_formatting,
    coalesce_document, CoalesceResult, pt_status, PT_STATUS_NS
};
```

### 3. Verification

**Compilation**: ✅ Success
```bash
$ cargo check --package redline-core
    Checking redline-core v0.1.0
    Finished `dev` profile [unoptimized + debuginfo] target(s) in 1.04s
```

**Unit Tests**: ✅ ALL PASSING (102/102)
```bash
$ cargo test --package redline-core --lib
running 102 tests
test result: ok. 102 passed; 0 failed; 0 ignored
```

**Integration Tests**: ⚠️ 56 PASSING, 48 FAILING
```bash
$ cargo test --package redline-core --test wml_tests
running 104 tests
test result: FAILED. 56 passed; 48 failed; 0 ignored
```

## Critical Finding: Integration Not Possible Yet

### The Problem

The current Rust `WmlComparer` implementation in `comparer.rs` is **NOT using the full atom-level comparison pipeline**. It's a simple stub that:

1. Does paragraph-level LCS comparison
2. Counts revisions based on paragraph differences
3. **Does NOT generate XML markup** with `w:ins`/`w:del` revision elements
4. Returns a `WmlComparisonResult` with just counts, not a modified document

### Why Consolidation Can't Be Called

The `coalesce_document()` function expects to operate on an XML document that contains revision marks (`w:ins`, `w:del` elements). However:

- The current implementation doesn't call `produce_markup_from_atoms()`
- It doesn't call `mark_content_as_deleted_or_inserted()`
- It doesn't generate any XML output with revision marks
- Therefore, there's nothing to consolidate

### C# vs Rust Pipeline

**C# Pipeline (WmlComparer.cs lines 2164-2174)**:
```csharp
// 1. Generate markup from correlated atoms
var newBodyChildren = ProduceNewWmlMarkupFromCorrelatedSequence(...);

// 2. Create document structure
XDocument newXDoc = new XDocument();
newXDoc.Add(new XElement(W.document, ...));

// 3. Wrap content in revision marks
MarkContentAsDeletedOrInserted(newXDoc, settings);           // Line 2173

// 4. CONSOLIDATE ADJACENT REVISIONS
CoalesceAdjacentRunsWithIdenticalFormatting(newXDoc);        // Line 2174

// 5. Cleanup
IgnorePt14Namespace(newXDoc.Root);
```

**Rust Current Implementation (comparer.rs lines 59-182)**:
```rust
pub fn compare(...) -> Result<WmlComparisonResult> {
    // Simple paragraph-level LCS
    let paras1 = find_paragraphs(&doc1, body1);
    let paras2 = find_paragraphs(&doc2, body2);
    let correlation = compute_correlation(&units1, &units2, ...);
    let (insertions, deletions) = count_revisions_smart(...);
    
    // Return counts only - NO XML GENERATION
    Ok(WmlComparisonResult {
        document: source2.to_bytes()?, // Original doc, not modified!
        changes: Vec::new(),
        insertions,
        deletions,
        ...
    })
}
```

## Test Failure Analysis

The 48 failing tests are failing because of **missing atom-level comparison**, not missing consolidation:

### Example Failures

| Test | Expected | Got | Reason |
|------|----------|-----|--------|
| WC-1000 (plain text) | 1 | 2 (1 ins, 1 del) | Paragraph-level treats modification as del+ins |
| WC-1160 (table) | 2 | 4 (2 ins, 2 del) | Missing atom-level granularity |
| WC-1420 (math) | 9 | 13 (7 ins, 6 del) | Complex content needs atom-level comparison |
| WC-1480 (simple table) | 4 | 2 (1 ins, 1 del) | Missing detailed correlation |

### Pattern Recognition

Most failures show one of these patterns:
- **Double counting**: Expected 1, got 2 (treating edit as delete+insert)
- **Over-counting**: Expected N, got N+K (missing consolidation would cause this IF pipeline existed)
- **Under-counting**: Expected N, got N-K (missing detailed atom-level comparison)

## What's Missing

To make the consolidation integration meaningful, the following needs to be implemented:

### 1. Complete Atom-Level Pipeline
```rust
// In WmlComparer::compare()

// 1. Create comparison unit atom lists
let atoms1 = create_comparison_unit_atom_list(&doc1, ...);
let atoms2 = create_comparison_unit_atom_list(&doc2, ...);

// 2. Build hierarchical comparison units
let units1 = get_comparison_unit_list(&atoms1, ...);
let units2 = get_comparison_unit_list(&atoms2, ...);

// 3. Perform multi-level LCS
let correlated = lcs(&units1, &units2, ...);

// 4. Flatten to atoms with status
let flattened_atoms = flatten_to_atoms(&correlated);

// 5. Generate result document
let result = produce_markup_from_atoms(&flattened_atoms, settings);

// 6. Mark content as deleted/inserted
mark_content_as_deleted_or_inserted(&mut result.document, result.root, settings);

// 7. CONSOLIDATE (this is where our new function goes!)
coalesce_document(&mut result.document, result.root);

// 8. Serialize and return
let result_bytes = result.document.to_bytes()?;
Ok(WmlComparisonResult { document: result_bytes, ... })
```

### 2. Update Integration Tests

Once the pipeline is complete, tests should pass because:
- Atom-level comparison provides correct granularity
- Consolidation merges adjacent identical revisions
- Result matches C# byte-for-byte

## Conclusion

### What Was Accomplished ✅

1. **Implemented `coalesce_document()`** - The consolidation function is ready and properly structured
2. **Verified compilation** - Code compiles without errors
3. **Documented integration point** - Clear understanding of where it fits in the pipeline
4. **Identified root cause** - Tests fail due to missing atom-level pipeline, not missing consolidation

### What's Still Needed ❌

1. **Implement full atom-level comparison pipeline** in `comparer.rs`
2. **Wire up the pipeline steps** as outlined above
3. **Call consolidation at the right point** (after mark_content_as_deleted_or_inserted)
4. **Verify byte-for-byte parity** with C# implementation

### Current Test Status

- **Unit tests**: 102/102 passing ✅
- **Integration tests**: 56/104 passing (53.8%)
- **Expected after full pipeline**: 104/104 passing (100%) 🎯

### Recommendation

The `coalesce_document()` function is **ready for integration** but cannot be used until the full atom-level comparison pipeline is implemented. The next phase should focus on:

1. Porting `CreateComparisonUnitAtomList`
2. Porting `GetComparisonUnitList`
3. Implementing multi-level LCS correlation
4. Integrating `produce_markup_from_atoms`, `mark_content_as_deleted_or_inserted`, and `coalesce_document` in sequence

This aligns with **Phase 2** of the RUST_MIGRATION_PLAN_SYNTHESIS.md document.

---

**Generated**: December 26, 2025  
**Files Modified**:
- `redline-rs/crates/redline-core/src/wml/coalesce.rs` (added `coalesce_document()`)
- `redline-rs/crates/redline-core/src/wml/mod.rs` (exported `coalesce_document`)
