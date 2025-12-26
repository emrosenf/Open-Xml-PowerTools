# Atom-Level Correlation Pipeline Analysis

## Executive Summary

**Root Cause**: The Rust WmlComparer fails 48 tests because it's missing the **CoalesceAdjacentRunsWithIdenticalFormatting** consolidation step that occurs at C# line 2174.

**Impact**: Without consolidation, one logical edit (same author/date/formatting) remains split into multiple adjacent `w:ins`/`w:del` elements, causing revision over-counting.

**Solution**: Implement consolidation after `mark_content_as_deleted_or_inserted`, matching C#'s exact grouping keys and eligibility rules.

---

## C# Pipeline (Source of Truth)

### Complete Flow (WmlComparer.cs lines 2164-2174)

```csharp
// 1. Generate markup from correlated atoms
var newBodyChildren = ProduceNewWmlMarkupFromCorrelatedSequence(
    wDocWithRevisions.MainDocumentPart,
    listOfComparisonUnitAtoms, 
    settings);

// 2. Create document structure  
XDocument newXDoc = new XDocument();
newXDoc.Add(new XElement(W.document,
    rootNamespaceAttributes,
    new XElement(W.body, newBodyChildren)));

// 3. Wrap content in revision marks (w:ins/w:del)
MarkContentAsDeletedOrInserted(newXDoc, settings);           // Line 2173

// 4. CONSOLIDATE ADJACENT REVISIONS ← MISSING IN RUST!
CoalesceAdjacentRunsWithIdenticalFormatting(newXDoc);        // Line 2174

// 5. Cleanup
IgnorePt14Namespace(newXDoc.Root);
```

---

## What Rust Currently Has

### ✅ Implemented
- `produce_markup_from_atoms` → ProduceNewWmlMarkupFromCorrelatedSequence
- `mark_content_as_deleted_or_inserted` → MarkContentAsDeletedOrInserted  
- Paragraph-level correlation via LCS
- Basic revision counting

### ❌ Missing
- `coalesce_adjacent_runs_with_identical_formatting` → CoalesceAdjacentRunsWithIdenticalFormatting
- Full atom-level correlation pipeline
- Hierarchical ComparisonUnit correlation
- Markup generation and document output

---

## Consolidation Algorithm (PtOpenXmlUtil.cs lines 799-991)

### Purpose
Merge adjacent `w:r`, `w:ins`, or `w:del` elements that represent a single logical change.

### Grouping Keys (CRITICAL - Must Match Exactly!)

#### w:r (Plain Text Run)
```
Key = "Wt" + rPrString
```
- Only merges runs with identical `w:rPr` formatting
- Must have exactly one non-`w:rPr` child
- Child must be `w:t` (text)

#### w:r (Instruction Text)
```
Key = "WinstrText" + rPrString  
```
- Same rules as text, but for `w:instrText`

#### w:ins (Insertion)
```
Key = "Wins2" + author + date + w:id + rPrString
```
- **Includes `w:id`** → Different IDs will NOT merge
- Must match: `w:ins/w:r/w:t` shape
- No nested `w:del` allowed

#### w:del (Deletion)  
```
Key = "Wdel" + author + date + rPrString
```
- **NO `w:id`** → Different IDs CAN merge (CRITICAL ASYMMETRY!)
- Must match: `w:del/w:r/w:delText` shape

### Eligible Shapes (Don't Consolidate Others!)

#### ✅ Safe to Merge
```xml
<!-- w:r with exactly one w:t -->
<w:r>
  <w:rPr>...</w:rPr>
  <w:t>text</w:t>
</w:r>

<!-- w:ins with exactly one w:r/w:t -->
<w:ins w:author="..." w:date="..." w:id="...">
  <w:r>
    <w:rPr>...</w:rPr>
    <w:t>inserted text</w:t>
  </w:r>
</w:ins>

<!-- w:del with exactly one w:r/w:delText -->
<w:del w:author="..." w:date="...">
  <w:r>
    <w:rPr>...</w:rPr>
    <w:delText>deleted text</w:delText>
  </w:r>
</w:del>
```

#### ❌ Don't Consolidate
```xml
<!-- Multiple children -->
<w:r>
  <w:t>text</w:t>
  <w:tab/>
</w:r>

<!-- Nested revisions -->
<w:ins>
  <w:del>...</w:del>
</w:ins>

<!-- List metadata -->
<w:r PtOpenXml.AbstractNumId="...">
  <w:t>...</w:t>
</w:r>

<!-- Complex content -->
<w:r>
  <w:drawing/>
</w:r>
```

### Recursive Processing

```csharp
// 1. Main document paragraphs (excluding txbxContent)
var paras = xDoc.Root.DescendantsTrimmed(W.txbxContent).Where(d => d.Name == W.p);
foreach (var para in paras)
{
    var newPara = WordprocessingMLUtil.CoalesceAdjacentRunsWithIdenticalFormatting(para);
    para.ReplaceNodes(newPara.Nodes());
}

// 2. txbxContent paragraphs (CRITICAL for round-trip!)
foreach (var txbx in xDoc.Root.Descendants(W.txbxContent))
{
    foreach (var txbxPara in txbx.DescendantsTrimmed(W.txbxContent).Where(d => d.Name == W.p))
    {
        var newPara = WordprocessingMLUtil.CoalesceAdjacentRunsWithIdenticalFormatting(txbxPara);
        txbxPara.ReplaceNodes(newPara.Nodes());
    }
}
```

**Why txbxContent is special:**
- Text boxes are nested "sub-documents"  
- Often duplicated via `mc:AlternateContent` (DrawingML + VML fallback)
- Leaving fragmented runs inside breaks reconstruction and Word round-trip

---

## Edge Cases (From Oracle Analysis)

### 1. **ID Asymmetry**
- `w:ins`: Different IDs → **separate revisions**
- `w:del`: Different IDs → **can merge**
- **Rust Impact**: If you include ID in deletion keys, you'll over-count

### 2. **Nested Revisions**
- `w:ins` containing `w:del` → **don't consolidate**
- Prevents destroying revision history

### 3. **Field Boundaries**
- Don't merge across `w:fldChar`, `w:instrText`, `w:fldSimple`
- Field structure must remain intact

### 4. **Container Boundaries**  
- Don't merge across: `w:p`, `w:tbl`, `w:tr`, `w:tc`
- Don't merge across: `w:hyperlink`, `w:sdtContent`, `w:moveFrom`, `w:moveTo`

### 5. **Whitespace Handling**
- When concatenating text, recompute `xml:space="preserve"`
- Needed when merged text has leading/trailing whitespace

### 6. **List/Numbering Runs**
- Runs with `PtOpenXml.AbstractNumId` → **never merge**
- Prevents breaking list reconstruction

---

## Performance Characteristics

### Algorithm Complexity
- GroupAdjacent: **O(n)** time, **O(k)** space (k = max group size)
- Key computation: **depends on rPr serialization**

### C# Bottlenecks
1. **String-based keys**: `rPr.ToString()` + concatenation
2. **Multiple DOM passes**: Emit → traverse → rebuild → traverse again
3. **LINQ allocations**: Iterator/collection overhead

### Rust Optimization Opportunities  

#### 1. Compact Key Representation
```rust
struct ConsolidationKey {
    kind: RevisionKind,        // enum: Ins/Del/None
    author_id: u32,            // interned ID
    date_id: u32,              // interned/normalized
    rev_id: Option<u32>,       // only for insertions
    rpr_sig_id: u64,           // hash or interned sig
}
```
**Benefit**: Integer comparisons instead of string ops

#### 2. Signature Interning
- Compute `rPr` signature once at atom creation
- Store as hash (u64) or interned ID
- Avoid repeated XML serialization

#### 3. Streaming Merge
```rust
// Instead of: build fragments → consolidate
// Do: merge during emission
while iterating atoms {
    if same_key_as_open_wrapper {
        append_to_current;
    } else {
        close_current;
        open_new_wrapper;
    }
}
```
**Benefit**: Zero-allocation consolidation

#### 4. Single-Pass Processing
- Correlate → Coalesce+Emit → Optional fixup
- Not: Emit → Traverse → Rebuild → Traverse → Rebuild

---

## Implementation Plan

### Phase 1: Minimal Fix (Short-term, 1-4h)

**Goal**: Make revision counts match C# without full markup generation

**Approach**: Logical consolidation during counting

```rust
fn count_revisions_with_consolidation(
    correlation: &[CorrelatedSequence],
) -> (usize, usize) {
    // Group adjacent segments by consolidation key
    let consolidated = consolidate_correlation_segments(correlation);
    
    // Count consolidated segments
    count_from_consolidated(&consolidated)
}
```

**Requirements**:
- Match C# grouping keys (including ID asymmetry)
- Respect same adjacency boundaries  
- Handle txbxContent grouping ("TXBX" marker)

### Phase 2: Full Pipeline (Medium-term, 1-2d)

**Goal**: Byte-for-byte compatible markup generation

**Steps**:
1. Implement full atom-level correlation
2. Port `ProduceNewWmlMarkupFromCorrelatedSequence`
3. Port `MarkContentAsDeletedOrInserted`  
4. **Port `CoalesceAdjacentRunsWithIdenticalFormatting`** ← Critical step
5. Port fixup methods (IDs, footnotes, styles)

**Structure**:
```rust
// redline-core/src/wml/consolidation.rs
pub fn coalesce_adjacent_runs(
    doc: &mut XmlDocument,
    root: NodeId,
    settings: &WmlComparerSettings,
) {
    // Process main paragraphs
    for para in find_paragraphs_excluding_txbx(doc, root) {
        coalesce_paragraph_runs(doc, para, settings);
    }
    
    // Process txbxContent paragraphs recursively
    for txbx in find_all_txbx_content(doc, root) {
        for para in find_paragraphs_in_txbx(doc, txbx) {
            coalesce_paragraph_runs(doc, para, settings);
        }
    }
}
```

---

## Testing Strategy

### Unit Tests
```rust
#[test]
fn consolidates_adjacent_insertions() {
    // Adjacent w:ins with same author/date/id/rPr → merge
}

#[test]
fn does_not_consolidate_different_ids() {
    // Adjacent w:ins with different ids → keep separate
}

#[test]
fn consolidates_deletions_despite_different_ids() {
    // Adjacent w:del with different ids → merge (asymmetry!)
}

#[test]
fn does_not_consolidate_nested_revisions() {
    // w:ins containing w:del → don't merge
}

#[test]
fn handles_txbx_content_recursively() {
    // Consolidation inside text boxes
}
```

### Integration Tests  
- Compare against C# WmlComparer golden files
- Verify revision counts match exactly
- Test round-trip: accept/reject all → should restore original

---

## Blockers & Risks

### Blockers
1. **XML Serialization**: Need deterministic `rPr.ToString()` equivalent
   - C# uses `ToString(SaveOptions.None)`
   - Must match attribute order, namespace prefixes
   
2. **DOM Manipulation**: Need efficient child replacement
   - Current `XmlDocument` API may need extensions

3. **GroupAdjacent Implementation**: Need LINQ-equivalent
   - Already exists in `util::lcs` as `group_adjacent`
   - May need adaptation for DOM nodes

### Risks
1. **Subtle Serialization Differences**
   - Different `rPr` serialization → keys don't match → no merge → over-counting
   
2. **Container Boundary Mismatches**
   - If Rust groups differently than C# → different consolidation → test failures

3. **Performance Regression**
   - Naive implementation could be slower than C# LINQ
   - Need to measure and optimize

---

## References

### C# Source Files
- `WmlComparer.cs` lines 2027-2238 (ProduceDocumentWithTrackedRevisions)
- `WmlComparer.cs` lines 2646-2824 (MarkContentAsDeletedOrInserted, CoalesceAdjacentRuns)
- `PtOpenXmlUtil.cs` lines 799-991 (CoalesceAdjacentRunsWithIdenticalFormatting)

### Rust Files
- `redline-core/src/wml/comparer.rs` (current implementation)
- `redline-core/src/wml/coalesce.rs` (markup generation, mark_content_as_deleted_or_inserted)
- `redline-core/src/wml/correlation.rs` (CorrelatedSequence types)

### Key Insights from Oracle Analysis
1. **Consolidation is not optional** - it defines "a revision"
2. **Must happen after markup generation** - needs real w:author/date/id attributes
3. **Two coalescing stages**: atom grouping + final run consolidation
4. **Rust can optimize** - compact keys, streaming merge, single-pass

---

## Next Steps

1. ✅ **Completed**: Deep analysis of C# consolidation algorithm
2. **Next**: Choose implementation path (Phase 1 vs Phase 2)
3. **Then**: Implement `coalesce_adjacent_runs_with_identical_formatting`
4. **Finally**: Run tests and verify revision counts match C#

