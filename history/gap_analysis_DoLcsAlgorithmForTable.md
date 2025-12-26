# Gap Analysis: DoLcsAlgorithmForTable (C# to Rust Port)

## Executive Summary

**Status: INCOMPLETE - Critical Missing Implementation**

The Rust port of `do_lcs_algorithm_for_table` is missing crucial logic for detecting merged cells in tables. The helper function `check_table_has_merged_cells` is a stub that always returns `false`, which means the algorithm will fail to correctly handle tables with merged cells (vertical merges via `vMerge` or horizontal merges via `gridSpan`).

**Gap Size:** Medium-High (one critical helper function missing implementation, ~10-20 lines of XML traversal logic needed)

---

## File Locations

### C# Source
- **File:** `OpenXmlPowerTools/WmlComparer.cs`
- **Line Range:** 7145-7255 (111 lines total)
- **Method Signature:** 
  ```csharp
  private static List<CorrelatedSequence> DoLcsAlgorithmForTable(
      CorrelatedSequence unknown, 
      WmlComparerSettings settings)
  ```

### Rust Source
- **File:** `redline-rs/crates/redline-core/src/wml/lcs_algorithm.rs`
- **Line Range:** 1345-1428 (84 lines total, including helper function)
- **Function Signature:**
  ```rust
  fn do_lcs_algorithm_for_table(
      units1: &[ComparisonUnit],
      units2: &[ComparisonUnit],
      _settings: &WmlComparerSettings,
  ) -> Option<Vec<CorrelatedSequence>>
  ```

---

## Algorithm Overview

The `DoLcsAlgorithmForTable` function implements table-specific LCS (Longest Common Subsequence) logic with special handling for merged cells. The algorithm has three distinct branches:

1. **Equal row count with matching hashes:** If tables have same number of rows and all corresponding rows have matching `CorrelatedSHA1Hash`, create paired unknown sequences
2. **Merged cells with matching structure:** If tables contain merged cells but have identical `StructureSHA1Hash`, pair rows by position
3. **Merged cells with different structure:** If tables contain merged cells but different structures, flatten to deleted + inserted sequences

---

## Line-by-Line Mapping

| C# Lines | Rust Lines | Section | Status | Notes |
|----------|------------|---------|--------|-------|
| 7145-7146 | 1345-1349 | Function signature | ✅ Complete | Rust uses slices instead of CorrelatedSequence parameter |
| 7147 | - | Variable declaration | ✅ Complete | Rust builds result inline |
| 7153-7154 | 1350-1357 | Extract table groups | ✅ Complete | Rust includes type validation |
| 7155-7179 | 1368-1390 | **Branch 1:** Equal rows, matching hashes | ✅ Complete | Logic identical |
| 7157-7161 | 1369-1372 | Zip rows together | ✅ Complete | Rust uses iterator zip |
| 7162-7164 | 1369-1375 | Check all hashes match | ✅ Complete | Rust uses `all()` combinator |
| 7165-7178 | 1377-1389 | Create paired sequences | ✅ Complete | Rust uses `map().collect()` |
| 7181-7195 | **MISSING** | Navigate to XML table elements | ❌ **MISSING** | C# walks ancestor elements to find `<w:tbl>` |
| 7197-7199 | 1392 (stub) | Check left table for merged cells | ⚠️ **STUB** | Rust helper always returns `false` |
| 7201-7203 | 1393 (stub) | Check right table for merged cells | ⚠️ **STUB** | Rust helper always returns `false` |
| 7205-7229 | 1395-1410 | **Branch 2:** Merged cells, matching structure | ⚠️ Conditional | Works only if stub is fixed |
| 7209-7211 | 1396-1398 | Check structure hashes match | ✅ Complete | Logic matches |
| 7213-7228 | 1399-1409 | Create paired sequences | ✅ Complete | Logic matches |
| 7231-7252 | 1412-1424 | **Branch 3:** Flatten to deleted + inserted | ✅ Complete | Logic matches |
| 7233-7240 | 1412-1415, 1422 | Create deleted sequence | ✅ Complete | Rust flattens rows correctly |
| 7242-7250 | 1416-1419, 1423 | Create inserted sequence | ✅ Complete | Rust flattens rows correctly |
| 7254 | 1427 | Return null/None | ✅ Complete | Return when no special handling needed |
| - | 1434-1436 | Helper: check_table_has_merged_cells | ❌ **STUB** | Always returns `false` |

---

## Critical Missing Logic

### 1. XML Table Element Navigation (C# lines 7181-7195)

**C# Implementation:**
```csharp
var firstContentAtom1 = tblGroup1.DescendantContentAtoms().FirstOrDefault();
if (firstContentAtom1 == null)
    throw new OpenXmlPowerToolsException("Internal error");
var tblElement1 = firstContentAtom1
    .AncestorElements
    .Reverse()
    .FirstOrDefault(a => a.Name == W.tbl);

var firstContentAtom2 = tblGroup2.DescendantContentAtoms().FirstOrDefault();
if (firstContentAtom2 == null)
    throw new OpenXmlPowerToolsException("Internal error");
var tblElement2 = firstContentAtom2
    .AncestorElements
    .Reverse()
    .FirstOrDefault(a => a.Name == W.tbl);
```

**Rust Status:** ❌ **COMPLETELY MISSING**

The Rust implementation does not navigate from `ComparisonUnitGroup` to the underlying XML `<w:tbl>` element. This is problematic because:
- The merged cell detection requires access to the actual XML structure
- The helper function receives a `ComparisonUnitGroup` but needs to access XML elements
- Without this, the helper function cannot examine descendants for `<w:vMerge>` or `<w:gridSpan>`

---

### 2. Merged Cell Detection (C# lines 7197-7203)

**C# Implementation:**
```csharp
var leftContainsMerged = tblElement1
    .Descendants()
    .Any(d => d.Name == W.vMerge || d.Name == W.gridSpan);

var rightContainsMerged = tblElement2
    .Descendants()
    .Any(d => d.Name == W.vMerge || d.Name == W.gridSpan);
```

**Rust Implementation (STUB):**
```rust
fn check_table_has_merged_cells(_table_group: &super::comparison_unit::ComparisonUnitGroup) -> bool {
    false  // ⚠️ ALWAYS RETURNS FALSE - NOT IMPLEMENTED
}
```

**Impact:**
- **HIGH SEVERITY:** The function will never detect merged cells
- Tables with `<w:vMerge>` (vertical merge) or `<w:gridSpan>` (horizontal merge) will be incorrectly processed
- Branch 2 and Branch 3 logic will never execute, even when they should
- This could lead to incorrect comparison results for complex tables

**What's Needed:**
1. Add method to `ComparisonUnitGroup` (or helper) to access underlying XML elements
2. Implement XML descendant traversal to find `<w:vMerge>` or `<w:gridSpan>` elements
3. Return `true` if any such elements exist

---

## Missing Sections Detail

### Section 1: XML Element Access
**C# Lines:** 7181-7195  
**Rust Lines:** None  
**Description:** Navigate from ComparisonUnitGroup → ComparisonUnitAtom → XML ancestor elements to find `<w:tbl>` element

**Required Implementation:**
- Add `descendant_content_atoms()` method to `ComparisonUnitGroup` (or equivalent)
- Add `ancestor_elements` field/method to `ComparisonUnitAtom` 
- Implement reverse iteration and element name matching

---

### Section 2: Merged Cell Detection Helper
**C# Lines:** 7197-7203  
**Rust Lines:** 1434-1436 (stub only)  
**Description:** Check if a table contains merged cells by examining XML descendants for `<w:vMerge>` or `<w:gridSpan>` elements

**Required Implementation:**
- Accept `ComparisonUnitGroup` or XML element as parameter
- Traverse all descendant XML elements
- Check for element names matching `W.vMerge` or `W.gridSpan`
- Return `true` if any found, `false` otherwise

**Estimated Complexity:** ~10-15 lines of Rust code, requires XML element access infrastructure

---

## Completeness Assessment

### ✅ Complete Sections
1. **Function structure and control flow** - All three branches present
2. **Branch 1: Equal rows with matching hashes** - Full implementation
3. **Branch 2: Merged cells with matching structure** - Logic complete (but depends on broken helper)
4. **Branch 3: Flatten to deleted/inserted** - Logic complete (but depends on broken helper)
5. **Return value handling** - Correct use of `Option<Vec<CorrelatedSequence>>`

### ❌ Incomplete Sections
1. **XML element navigation** - Completely missing (14 lines in C#)
2. **Merged cell detection** - Stub implementation (7 lines in C#, needs ~10-15 in Rust)

### ⚠️ Conditionally Working
- Branches 2 and 3 have correct logic but **will never execute** because the helper function always returns `false`
- The algorithm works correctly ONLY for tables without merged cells
- Any table with `<w:vMerge>` or `<w:gridSpan>` will be incorrectly processed

---

## Impact Analysis

### Functional Impact
- **Critical:** Tables with merged cells will not be compared correctly
- **Risk:** Silent failures - function returns `None` when it should handle merged cell cases
- **Coverage:** Approximately 30-40% of the algorithm logic is effectively disabled

### Test Coverage Impact
- Current tests likely pass because they don't use tables with merged cells
- Need tests with:
  - Vertically merged cells (`<w:vMerge>`)
  - Horizontally merged cells (`<w:gridSpan>`)
  - Tables with matching structure but merged cells
  - Tables with different structure and merged cells

---

## Recommendations

### Priority 1: Critical Fix
1. **Implement `check_table_has_merged_cells` helper function**
   - Add XML element access to `ComparisonUnitGroup` or related types
   - Implement descendant traversal
   - Check for `<w:vMerge>` and `<w:gridSpan>` elements
   - Add unit tests for detection

### Priority 2: Infrastructure
2. **Add XML element navigation support**
   - Implement `descendant_content_atoms()` on `ComparisonUnitGroup`
   - Add `ancestor_elements` to `ComparisonUnitAtom`
   - Ensure XML element references are available throughout comparison pipeline

### Priority 3: Testing
3. **Add comprehensive tests**
   - Test tables with vertical merges
   - Test tables with horizontal merges
   - Test tables with both merge types
   - Test edge cases (empty cells, complex nesting)

### Priority 4: Documentation
4. **Update code comments**
   - Document the stub nature of the current implementation
   - Add TODO comments pointing to this gap analysis
   - Mark function as incomplete in API documentation

---

## Conclusion

**The Rust port of `do_lcs_algorithm_for_table` is INCOMPLETE.**

While the overall control flow and algorithm structure are correctly ported, the critical merged cell detection logic is missing. The function will work correctly for simple tables without merged cells, but will fail silently or produce incorrect results for tables containing `<w:vMerge>` or `<w:gridSpan>` elements.

**Estimated Effort to Complete:**
- Small (1-2 hours): If XML element access infrastructure already exists elsewhere in the codebase
- Medium (4-8 hours): If XML element access needs to be added to comparison unit types
- Includes: Implementation (15-20 lines) + Tests (50-100 lines) + Integration testing

**Blocking Factor:**
This gap may indicate that earlier phases of the Rust port (XML element access, ancestor tracking) are also incomplete. The completion of this function depends on having proper XML element references available in the comparison unit data structures.

---

## Appendix: Full Code Comparison

### C# Full Implementation (Lines 7145-7255)
```csharp
private static List<CorrelatedSequence> DoLcsAlgorithmForTable(CorrelatedSequence unknown, WmlComparerSettings settings)
{
    List<CorrelatedSequence> newListOfCorrelatedSequence = new List<CorrelatedSequence>();

    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    // if we have a table with the same number of rows, and all rows have equal CorrelatedSHA1Hash, then we can flatten and compare every corresponding row.
    // This is true regardless of whether there are horizontally or vertically merged cells, since that characteristic is incorporated into the CorrespondingSHA1Hash.
    // This is probably not very common, but it will never do any harm.
    var tblGroup1 = unknown.ComparisonUnitArray1.First() as ComparisonUnitGroup;
    var tblGroup2 = unknown.ComparisonUnitArray2.First() as ComparisonUnitGroup;
    if (tblGroup1.Contents.Count() == tblGroup2.Contents.Count()) // if there are the same number of rows
    {
        var zipped = tblGroup1.Contents.Zip(tblGroup2.Contents, (r1, r2) => new
        {
            Row1 = r1 as ComparisonUnitGroup,
            Row2 = r2 as ComparisonUnitGroup,
        });
        var canCollapse = true;
        if (zipped.Any(z => z.Row1.CorrelatedSHA1Hash != z.Row2.CorrelatedSHA1Hash))
            canCollapse = false;
        if (canCollapse)
        {
            newListOfCorrelatedSequence = zipped
                .Select(z =>
                {
                    var unknownCorrelatedSequence = new CorrelatedSequence();
                    unknownCorrelatedSequence.ComparisonUnitArray1 = new[] { z.Row1 };
                    unknownCorrelatedSequence.ComparisonUnitArray2 = new[] { z.Row2 };
                    unknownCorrelatedSequence.CorrelationStatus = CorrelationStatus.Unknown;
                    return unknownCorrelatedSequence;
                })
                .ToList();
            return newListOfCorrelatedSequence;
        }
    }

    var firstContentAtom1 = tblGroup1.DescendantContentAtoms().FirstOrDefault();
    if (firstContentAtom1 == null)
        throw new OpenXmlPowerToolsException("Internal error");
    var tblElement1 = firstContentAtom1
        .AncestorElements
        .Reverse()
        .FirstOrDefault(a => a.Name == W.tbl);

    var firstContentAtom2 = tblGroup2.DescendantContentAtoms().FirstOrDefault();
    if (firstContentAtom2 == null)
        throw new OpenXmlPowerToolsException("Internal error");
    var tblElement2 = firstContentAtom2
        .AncestorElements
        .Reverse()
        .FirstOrDefault(a => a.Name == W.tbl);

    var leftContainsMerged = tblElement1
        .Descendants()
        .Any(d => d.Name == W.vMerge || d.Name == W.gridSpan);

    var rightContainsMerged = tblElement2
        .Descendants()
        .Any(d => d.Name == W.vMerge || d.Name == W.gridSpan);

    if (leftContainsMerged || rightContainsMerged)
    {
        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        // If StructureSha1Hash is the same for both tables, then we know that the structure of the tables is identical, so we can break into correlated sequences for rows.
        if (tblGroup1.StructureSHA1Hash != null &&
            tblGroup2.StructureSHA1Hash != null &&
            tblGroup1.StructureSHA1Hash == tblGroup2.StructureSHA1Hash)
        {
            var zipped = tblGroup1.Contents.Zip(tblGroup2.Contents, (r1, r2) => new
            {
                Row1 = r1 as ComparisonUnitGroup,
                Row2 = r2 as ComparisonUnitGroup,
            });
            newListOfCorrelatedSequence = zipped
                .Select(z =>
                {
                    var unknownCorrelatedSequence = new CorrelatedSequence();
                    unknownCorrelatedSequence.ComparisonUnitArray1 = new[] { z.Row1 };
                    unknownCorrelatedSequence.ComparisonUnitArray2 = new[] { z.Row2 };
                    unknownCorrelatedSequence.CorrelationStatus = CorrelationStatus.Unknown;
                    return unknownCorrelatedSequence;
                })
                .ToList();
            return newListOfCorrelatedSequence;
        }

        // otherwise flatten to rows
        var deletedCorrelatedSequence = new CorrelatedSequence();
        deletedCorrelatedSequence.ComparisonUnitArray1 = unknown
            .ComparisonUnitArray1
            .Select(z => z.Contents)
            .SelectMany(m => m)
            .ToArray();
        deletedCorrelatedSequence.ComparisonUnitArray2 = null;
        deletedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Deleted;
        newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);

        var insertedCorrelatedSequence = new CorrelatedSequence();
        insertedCorrelatedSequence.ComparisonUnitArray1 = null;
        insertedCorrelatedSequence.ComparisonUnitArray2 = unknown
            .ComparisonUnitArray2
            .Select(z => z.Contents)
            .SelectMany(m => m)
            .ToArray();
        insertedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Inserted;
        newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);

        return newListOfCorrelatedSequence;
    }
    return null;
}
```

### Rust Current Implementation (Lines 1345-1436)
```rust
fn do_lcs_algorithm_for_table(
    units1: &[ComparisonUnit],
    units2: &[ComparisonUnit],
    _settings: &WmlComparerSettings,
) -> Option<Vec<CorrelatedSequence>> {
    let tbl_group1 = units1.first()?.as_group()?;
    let tbl_group2 = units2.first()?.as_group()?;

    if tbl_group1.group_type != ComparisonUnitGroupType::Table
        || tbl_group2.group_type != ComparisonUnitGroupType::Table
    {
        return None;
    }

    let rows1 = match &tbl_group1.contents {
        super::comparison_unit::ComparisonUnitGroupContents::Groups(groups) => groups,
        _ => return None,
    };
    let rows2 = match &tbl_group2.contents {
        super::comparison_unit::ComparisonUnitGroupContents::Groups(groups) => groups,
        _ => return None,
    };

    if rows1.len() == rows2.len() {
        let all_rows_match = rows1
            .iter()
            .zip(rows2.iter())
            .all(|(r1, r2)| {
                r1.correlated_sha1_hash.is_some()
                    && r1.correlated_sha1_hash == r2.correlated_sha1_hash
            });

        if all_rows_match {
            let sequences: Vec<_> = rows1
                .iter()
                .zip(rows2.iter())
                .map(|(r1, r2)| {
                    CorrelatedSequence::unknown(
                        vec![ComparisonUnit::Group(r1.clone())],
                        vec![ComparisonUnit::Group(r2.clone())],
                    )
                })
                .collect();
            return Some(sequences);
        }
    }

    let left_contains_merged = check_table_has_merged_cells(tbl_group1);
    let right_contains_merged = check_table_has_merged_cells(tbl_group2);

    if left_contains_merged || right_contains_merged {
        if tbl_group1.structure_sha1_hash.is_some()
            && tbl_group1.structure_sha1_hash == tbl_group2.structure_sha1_hash
        {
            let sequences: Vec<_> = rows1
                .iter()
                .zip(rows2.iter())
                .map(|(r1, r2)| {
                    CorrelatedSequence::unknown(
                        vec![ComparisonUnit::Group(r1.clone())],
                        vec![ComparisonUnit::Group(r2.clone())],
                    )
                })
                .collect();
            return Some(sequences);
        }

        let flattened1: Vec<_> = rows1
            .iter()
            .map(|r| ComparisonUnit::Group(r.clone()))
            .collect();
        let flattened2: Vec<_> = rows2
            .iter()
            .map(|r| ComparisonUnit::Group(r.clone()))
            .collect();

        return Some(vec![
            CorrelatedSequence::deleted(flattened1),
            CorrelatedSequence::inserted(flattened2),
        ]);
    }

    None
}

/// Check if a table contains merged cells
///
/// This is a simplified check - in the full implementation, this would
/// examine the table's XML structure for vMerge or gridSpan elements.
fn check_table_has_merged_cells(_table_group: &super::comparison_unit::ComparisonUnitGroup) -> bool {
    false  // ⚠️ STUB - NOT IMPLEMENTED
}
```

---

**Analysis Date:** 2025-12-26  
**Analyst:** AI Code Review Agent  
**Task ID:** cell--tsysh-mjndvw6nvuy
