# OOXML Table Revision Markup Fix

**Date**: 2025-12-30  
**Epic ID**: cell-gpsl8v-mjsuzmz41tj  
**Status**: In Progress

## Executive Summary

This document describes a critical regression in the redline-rs WmlComparer where table rows and cells are incorrectly wrapped in `<w:ins>`/`<w:del>` elements, causing MS Word to report document corruption. The fix requires implementing property-based revision tracking for table structures per ECMA-376 specification.

---

## Problem Statement

### Symptoms
- **Commit 31f4d40**: MS Word opens correctly, but 39 wml_tests fail
- **HEAD**: 9 wml_tests fail, but MS Word reports "The file is corrupt and cannot be opened"

### Root Cause
Commit `b6f5f32` introduced "uniform status wrapping" in `reconstruct_element()` that wraps entire table rows (`<w:tr>`) and cells (`<w:tc>`) in `<w:ins>`/`<w:del>` elements. This violates the OOXML (ECMA-376) schema.

---

## OOXML Schema Requirements

### Revision Markup by Element Type

| Element Type | Revision Location | Correct Pattern |
|--------------|-------------------|-----------------|
| **Run (w:r)** | Wrapper | `<w:ins><w:r>...</w:r></w:ins>` |
| **Paragraph (w:p)** | Wrapper (at body level) | `<w:ins><w:p>...</w:p></w:ins>` |
| **Table Row (w:tr)** | **Property** | `<w:trPr><w:ins .../></w:trPr>` |
| **Table Cell (w:tc)** | **Property** | `<w:tcPr><w:cellIns .../></w:tcPr>` |
| **Math (m:oMath)** | Wrapper | `<w:ins><m:oMath>...</m:oMath></w:ins>` |

### Correct Table Row Revision XML

```xml
<!-- CORRECT: Inserted table row -->
<w:tr>
  <w:trPr>
    <w:ins w:id="1" w:author="User" w:date="2025-12-30T12:00:00Z" />
  </w:trPr>
  <w:tc>
    <w:tcPr>...</w:tcPr>
    <w:p>...</w:p>
  </w:tc>
</w:tr>

<!-- INCORRECT: Causes corruption -->
<w:ins w:id="1" w:author="User" w:date="2025-12-30T12:00:00Z">
  <w:tr>
    <w:tc>...</w:tc>
  </w:tr>
</w:ins>
```

### Correct Table Cell Revision XML

```xml
<!-- CORRECT: Inserted table cell -->
<w:tc>
  <w:tcPr>
    <w:cellIns w:id="2" w:author="User" w:date="2025-12-30T12:00:00Z" />
  </w:tcPr>
  <w:p>...</w:p>
</w:tc>

<!-- Note: Cells use w:cellIns/w:cellDel, NOT w:ins/w:del -->
```

---

## C# Reference Implementation

The C# OpenXmlPowerTools WmlComparer uses different strategies for different elements:

### Table Rows (Property-Based)
```csharp
// From WmlComparer.cs - MarkRowsAsDeletedOrInserted
XElement tr = firstContentAtom.AncestorElements.Reverse().FirstOrDefault(a => a.Name == W.tr);
XElement trPr = tr.Element(W.trPr);
if (trPr == null) {
    trPr = new XElement(W.trPr);
    tr.AddFirst(trPr);
}
XName revTrackElementName = (status == CorrelationStatus.Deleted) ? W.del : W.ins;
trPr.Add(new XElement(revTrackElementName,
    new XAttribute(W.author, settings.AuthorForRevisions),
    new XAttribute(W.id, _maxId++),
    new XAttribute(W.date, settings.DateTimeForRevisions)));
```

### Runs (Wrapper-Based)
```csharp
// From WmlComparer.cs - MarkContentAsDeletedOrInsertedTransform
return new XElement(W.ins,  // Wrapping is OK for runs
    new XAttribute(W.author, settings.AuthorForRevisions),
    new XAttribute(W.id, _maxId++),
    new XAttribute(W.date, settings.DateTimeForRevisions),
    new XElement(W.r, ...));
```

---

## Current Rust Bug

### Location
`crates/redline-core/src/wml/coalesce.rs`, lines 1699-1725

### Problematic Code
```rust
fn reconstruct_element(doc: &mut XmlDocument, parent: NodeId, group_key: &str, 
    ancestor: &AncestorElementInfo, _props_names: &[&str], 
    group_atoms: &[ComparisonUnitAtom], level: usize, part: Option<()>, 
    settings: &WmlComparerSettings) 
{
    // BUG: This wrapping logic is applied to ALL elements including tr/tc
    let uniform_status = get_uniform_status(group_atoms);
    
    let (container, atoms_vec) = if let Some(status) = uniform_status {
        let wrapper = create_revision_wrapper(doc, parent, status, settings).unwrap();
        // Creates <w:ins><w:tr>...</w:tr></w:ins> - INVALID!
        // ...
    };
    // ...
}
```

### Why It's Wrong
The `reconstruct_element()` function is used for `tbl`, `tr`, `tc`, `sdt`, `ruby`, and other elements. The wrapper-based approach is only valid for some of these (like `oMath`), but NOT for table structures.

---

## Implementation Plan

### Phase 1: Create Element-Specific Functions

#### 1.1 Create `reconstruct_table_row()`
```rust
fn reconstruct_table_row(
    doc: &mut XmlDocument, 
    parent: NodeId, 
    group_key: &str,
    ancestor: &AncestorElementInfo, 
    atoms: &[ComparisonUnitAtom], 
    grouped_children: &[(String, usize, usize)],
    level: usize, 
    part: Option<()>, 
    settings: &WmlComparerSettings
) {
    let uniform_status = get_uniform_status(atoms);
    
    // Create tr element directly under parent (NO wrapper)
    let mut attrs = ancestor.attributes.clone();
    attrs.retain(|a| a.name.namespace.as_deref() != Some(PT_STATUS_NS));
    let tr = doc.add_child(parent, XmlNodeData::element_with_attrs(W::tr(), attrs));
    
    // If uniform status, add revision to trPr (property-based)
    if let Some(status) = uniform_status {
        add_revision_to_tr_properties(doc, tr, status, settings);
    }
    
    // Recurse for children with suppressed status if needed
    let atoms_for_children = if uniform_status.is_some() {
        suppress_inner_revisions(atoms)
    } else {
        atoms.to_vec()
    };
    
    // Reconstruct trPr and cell children...
}

fn add_revision_to_tr_properties(
    doc: &mut XmlDocument, 
    tr: NodeId, 
    status: ComparisonCorrelationStatus, 
    settings: &WmlComparerSettings
) {
    // Find or create w:trPr as first child
    let tr_pr = ensure_first_child_is_tr_pr(doc, tr);
    
    let rev_name = match status {
        ComparisonCorrelationStatus::Inserted => W::ins(),
        ComparisonCorrelationStatus::Deleted => W::del(),
        _ => return,
    };
    
    // CRITICAL: w:id must be FIRST attribute per ECMA-376
    let id_str = next_revision_id().to_string();
    let attrs = vec![
        XAttribute::new(W::id(), &id_str),
        XAttribute::new(W::author(), settings.author_for_revisions.as_deref().unwrap_or("Unknown")),
        XAttribute::new(W::date(), settings.date_time_for_revisions.as_deref().unwrap_or("1970-01-01T00:00:00Z")),
        XAttribute::new(W16DU::dateUtc(), settings.date_time_for_revisions.as_deref().unwrap_or("1970-01-01T00:00:00Z")),
    ];
    
    doc.add_child(tr_pr, XmlNodeData::element_with_attrs(rev_name, attrs));
}
```

#### 1.2 Create `reconstruct_table_cell()`
```rust
fn reconstruct_table_cell(
    doc: &mut XmlDocument, 
    parent: NodeId, 
    group_key: &str,
    ancestor: &AncestorElementInfo, 
    atoms: &[ComparisonUnitAtom], 
    grouped_children: &[(String, usize, usize)],
    level: usize, 
    part: Option<()>, 
    settings: &WmlComparerSettings
) {
    let uniform_status = get_uniform_status(atoms);
    
    // Create tc element directly under parent (NO wrapper)
    let mut attrs = ancestor.attributes.clone();
    attrs.retain(|a| a.name.namespace.as_deref() != Some(PT_STATUS_NS));
    let tc = doc.add_child(parent, XmlNodeData::element_with_attrs(W::tc(), attrs));
    
    // If uniform status, add cell revision to tcPr (property-based)
    if let Some(status) = uniform_status {
        add_revision_to_tc_properties(doc, tc, status, settings);
    }
    
    // Recurse for children...
}

fn add_revision_to_tc_properties(
    doc: &mut XmlDocument, 
    tc: NodeId, 
    status: ComparisonCorrelationStatus, 
    settings: &WmlComparerSettings
) {
    let tc_pr = ensure_first_child_is_tc_pr(doc, tc);
    
    // NOTE: Cells use w:cellIns / w:cellDel, NOT w:ins / w:del
    let rev_name = match status {
        ComparisonCorrelationStatus::Inserted => W::cellIns(),
        ComparisonCorrelationStatus::Deleted => W::cellDel(),
        _ => return,
    };
    
    let id_str = next_revision_id().to_string();
    let attrs = vec![
        XAttribute::new(W::id(), &id_str),
        XAttribute::new(W::author(), settings.author_for_revisions.as_deref().unwrap_or("Unknown")),
        XAttribute::new(W::date(), settings.date_time_for_revisions.as_deref().unwrap_or("1970-01-01T00:00:00Z")),
    ];
    
    doc.add_child(tc_pr, XmlNodeData::element_with_attrs(rev_name, attrs));
}
```

### Phase 2: Update Dispatch Logic

Update `coalesce_recurse()` to use the new functions:

```rust
match ancestor_name.as_str() {
    "p" => reconstruct_paragraph(...),      // Keep - wrapper OK
    "r" => reconstruct_run(...),            // Keep - wrapper OK  
    "t" => reconstruct_text_elements(...),  // Keep
    "drawing" => reconstruct_drawing_elements(...),  // Keep
    "oMath" | "oMathPara" => reconstruct_element(...),  // Keep - wrapper OK for math
    
    // NEW: Table-specific functions with property-based revisions
    "tr" => reconstruct_table_row(doc, parent, &group_key, &ancestor_being_constructed, group_atoms, &grouped_children, level, _part, settings),
    "tc" => reconstruct_table_cell(doc, parent, &group_key, &ancestor_being_constructed, group_atoms, &grouped_children, level, _part, settings),
    "tbl" => reconstruct_table(doc, parent, &group_key, &ancestor_being_constructed, group_atoms, &grouped_children, level, _part, settings),
    
    // Generic fallback (no wrapping for unknown elements)
    _ => reconstruct_element_no_wrap(...),
}
```

### Phase 3: Ensure Namespace Definitions Exist

Check `crates/redline-core/src/xml/namespaces.rs` for:
- `W::cellIns()` 
- `W::cellDel()`
- `W::trPr()`
- `W::tcPr()`

---

## Affected Tests

| Test ID | Test Name | Expected | Current | Issue |
|---------|-----------|----------|---------|-------|
| WC-1450 | table_4_row_image | 7 | 4 | Over-coalescing |
| WC-1470 | table2 | 7 | 4 | Over-coalescing |
| WC-1480 | simple_table | 4 | 3 | Over-coalescing |
| WC-1500 | long_table | 2 | 3 | Extra wrapper |
| WC-1660 | footnote_with_table | 5 | 7 | Extra wrappers |
| WC-1670 | footnote_with_table_reverse | 5 | 7 | Extra wrappers |
| WC-1770 | textbox | 2 | 3 | Extra wrapper |
| WC-1830 | table_5 | 2 | 5 | Extra wrappers |
| WC-1840 | table_5_2 | 2 | 4 | Extra wrappers |

---

## Verification Steps

1. **Build**: `cargo build --release`
2. **Test with real documents**:
   ```bash
   ./target/release/redline compare \
     --source1 "test1.docx" \
     --source2 "test2.docx" \
     --output /tmp/test-result.docx
   ```
3. **Open in MS Word** - must not show corruption error
4. **Run test suite**: `cargo test --package redline-core --test wml_tests`
5. **Check for regressions** in the 95 previously passing tests

---

## References

- **ECMA-376 Part 1**: Office Open XML File Formats - Fundamentals and Markup Language Reference
- **ISO/IEC 29500-1 §17.13.5.17**: Table Row Insertion/Deletion
- **ISO/IEC 29500-1 §17.13.5.2**: Cell Insertion (w:cellIns)
- **C# OpenXmlPowerTools**: `WmlComparer.cs` - `MarkRowsAsDeletedOrInserted` function

---

## Rollback Plan

If this fix introduces new issues:
1. Revert to commit `31f4d40` (known MS Word compatible state)
2. Cherry-pick individual fixes from later commits
3. Re-approach table revision tracking with smaller incremental changes

---

## Change Log

| Date | Author | Change |
|------|--------|--------|
| 2025-12-30 | Analysis | Initial documentation and epic creation |
