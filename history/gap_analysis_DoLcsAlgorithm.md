# DoLcsAlgorithm Gap Analysis: C# vs Rust

**Analysis Date:** December 26, 2025  
**Analyst:** Deep Analysis Mode (10+ parallel agents)  
**Scope:** C# WmlComparer.cs lines 6148-7144 vs Rust lcs_algorithm.rs lines 637-847  

---

## Executive Summary

### Total Gap Size
- **C# Implementation:** 997 lines (6148-7144)
- **Rust Implementation:** 211 lines (637-847) in `do_lcs_algorithm` + ~600 lines in helper functions
- **Missing from Rust:** ~400-500 lines of critical logic
- **Severity:** **P0 CRITICAL** - Core algorithm completeness gap

### Key Findings
The Rust port implements approximately **40-50%** of the C# DoLcsAlgorithm functionality. Critical missing sections include:

1. **Paragraph boundary preservation logic** (C# lines 6927-7130, ~200 lines) - **P0**
2. **Advanced content type grouping** (C# lines 6370-6625, ~255 lines) - **P0**  
3. **Table-specific comparison algorithm** (C# lines 7145-7255, ~110 lines) - **P0**
4. **Paragraph mark priority ordering** (C# lines 6827-6910, ~80 lines) - **P1**
5. **Cell content flattening** (C# lines 6771-6824, ~50 lines) - **P1**

---

## Section 1: What IS Ported (Line-by-Line Mapping)

| Functionality | C# Lines | Rust Lines | Completeness | Notes |
|---------------|----------|------------|--------------|-------|
| **Empty array handling** | 6155-6178 | 644-671 | ✅ 100% | Fully ported with pattern matching |
| **LCS computation** | 6180-6221 | 676-703 | ✅ 100% | Hash-based LCS with same algorithm |
| **Paragraph mark filtering** | 6223-6257 | 705-722 | ✅ 95% | Core logic ported, minor differences |
| **Single paragraph mark check** | 6259-6278 | 724-732 | ✅ 100% | Fully ported |
| **Single space filtering** | 6280-6293 | 734-743 | ✅ 100% | Fully ported |
| **Word break character filtering** | 6295-6332 | 745-785 | ✅ 95% | Ported with CJK handling |
| **Detail threshold** | 6334-6350 | 787-801 | ✅ 100% | Fully ported |
| **Basic no-match handling** | 6352-6369 | 804-806 | ✅ 80% | Delegates to helper, partial |
| **Content type counting** | 6354-6368 | 858-936 | ✅ 100% | Ported in `count_group_types` |
| **Basic result construction** | 6992-7082 | 813-846 | ⚠️ 40% | Simple prefix/match/suffix only |

**Total Ported:** ~350 lines (35% of C# implementation)

---

## Section 2: What is MISSING (Grouped by Functionality)

### **GROUP A: Paragraph Boundary Preservation (P0 - CRITICAL)**

**C# Lines:** 6927-7130 (204 lines)  
**Impact:** Core correctness for paragraph-level comparison  
**Status:** ❌ NOT PORTED

#### Missing Logic:

**1. Paragraph-Aware LCS Splitting (lines 6936-6987)**
```csharp
// Detect paragraph marks in common sequence
if (commonSeq.Any(cu => /* contains paragraph mark */)) {
    remainingInLeftParagraph = unknown.ComparisonUnitArray1
        .Take(currentI1)
        .Reverse()
        .TakeWhile(cu => /* not paragraph mark */)
        .Count();
    // Calculate how much content before LCS belongs to same paragraph
}
```

**Purpose:** When LCS starts mid-paragraph, group preceding content from the same paragraph into a separate Unknown sequence for proper comparison.

**Why Critical:** Without this, paragraph boundaries are not respected, leading to:
- Incorrect correlation of content across paragraph boundaries
- Lost formatting when paragraphs merge
- Broken document structure in output

**Rust Gap:** The current Rust implementation at lines 813-846 does simple prefix/match/suffix splitting without paragraph awareness.

**2. Before-Paragraph Content Handling (lines 6989-7028)**
```csharp
var countBeforeCurrentParagraphLeft = currentI1 - remainingInLeftParagraph;
var countBeforeCurrentParagraphRight = currentI2 - remainingInRightParagraph;

if (countBeforeCurrentParagraphLeft > 0 && countBeforeCurrentParagraphRight == 0) {
    // Create Deleted sequence for content before current paragraph
}
else if (countBeforeCurrentParagraphLeft == 0 && countBeforeCurrentParagraphRight > 0) {
    // Create Inserted sequence
}
else if (countBeforeCurrentParagraphLeft > 0 && countBeforeCurrentParagraphRight > 0) {
    // Create Unknown sequence for pre-paragraph content
}
```

**Purpose:** Handle content that appears before the current paragraph boundary on one or both sides.

**Why Critical:** Ensures proper correlation status (Deleted/Inserted/Unknown) for content before paragraph marks.

**3. Within-Paragraph Remainder Handling (lines 7030-7069)**
```csharp
if (remainingInLeftParagraph > 0 && remainingInRightParagraph == 0) {
    var deletedCorrelatedSequence = new CorrelatedSequence();
    deletedCorrelatedSequence.ComparisonUnitArray1 = cul1
        .Skip(countBeforeCurrentParagraphLeft)
        .Take(remainingInLeftParagraph)
        .ToArray();
    // Mark as Deleted
}
// Similar for Inserted and Unknown cases
```

**Purpose:** Handle content within the same paragraph as the LCS match but before the LCS starts.

**Why Critical:** Preserves paragraph coherence in diff output.

**4. Post-LCS Paragraph Handling (lines 7095-7122)**
```csharp
var leftCuw = middleEqual.ComparisonUnitArray1[middleEqual.ComparisonUnitArray1.Length - 1] as ComparisonUnitWord;
if (leftCuw != null) {
    var lastContentAtom = leftCuw.DescendantContentAtoms().LastOrDefault();
    // If the middleEqual did not end with a paragraph mark
    if (lastContentAtom != null && lastContentAtom.ContentElement.Name != W.pPr) {
        int idx1 = FindIndexOfNextParaMark(remaining1);
        int idx2 = FindIndexOfNextParaMark(remaining2);
        
        // Create Unknown sequence up to next paragraph mark
        var unknownCorrelatedSequenceRemaining = new CorrelatedSequence();
        unknownCorrelatedSequenceRemaining.ComparisonUnitArray1 = remaining1.Take(idx1).ToArray();
        unknownCorrelatedSequenceRemaining.ComparisonUnitArray2 = remaining2.Take(idx2).ToArray();
        
        // Create Unknown sequence for content after paragraph mark
        var unknownCorrelatedSequenceAfter = new CorrelatedSequence();
        unknownCorrelatedSequenceAfter.ComparisonUnitArray1 = remaining1.Skip(idx1).ToArray();
        unknownCorrelatedSequenceAfter.ComparisonUnitArray2 = remaining2.Skip(idx2).ToArray();
    }
}
```

**Purpose:** When the LCS match doesn't end on a paragraph boundary, find the next paragraph mark and create separate Unknown sequences for content within the paragraph vs after.

**Why Critical:** 
- Ensures content within same paragraph is compared together
- Prevents cross-paragraph correlation errors
- Critical for maintaining document structure integrity

**5. Helper Method: FindIndexOfNextParaMark (lines 7133-7143)**
```csharp
private static int FindIndexOfNextParaMark(ComparisonUnit[] cul)
{
    for (int i = 0; i < cul.Length; i++)
    {
        var cuw = cul[i] as ComparisonUnitWord;
        var lastAtom = cuw.DescendantContentAtoms().LastOrDefault();
        if (lastAtom.ContentElement.Name == W.pPr)
            return i;
    }
    return cul.Length;
}
```

**Purpose:** Find the index of the next paragraph mark in a unit array.

**Why Critical:** Used by post-LCS paragraph handling logic.

**Rust Implementation Status:** ❌ NOT IMPLEMENTED

---

### **GROUP B: Advanced Content Type Grouping (P0 - CRITICAL)**

**C# Lines:** 6370-6625 (256 lines)  
**Impact:** Core correctness for mixed content types  
**Status:** ⚠️ PARTIALLY PORTED (basic version only)

#### Missing Logic:

**1. Mixed Words/Rows/Textboxes with Paragraph Mark Priority (lines 6383-6524)**

**Current Rust (lines 972-1019):** Basic grouping and pairing
```rust
fn handle_mixed_words_rows_textboxes(
    units1: &[ComparisonUnit],
    units2: &[ComparisonUnit],
) -> Vec<CorrelatedSequence> {
    // Basic: Group by type, pair matching types, mark mismatches as deleted/inserted
    // MISSING: Paragraph mark priority logic
}
```

**Missing C# Logic (lines 6447-6471):**
```csharp
// Special case: word group followed by row group
else if (leftGrouped[iLeft].Key == "Word" &&
    leftGrouped[iLeft].Select(lg => lg.DescendantContentAtoms()).SelectMany(m => m).Last().ContentElement.Name != W.pPr &&
    rightGrouped[iRight].Key == "Row")
{
    // If the word group does NOT end with paragraph mark, insert the row BEFORE deleting the word
    var insertedCorrelatedSequence = new CorrelatedSequence();
    insertedCorrelatedSequence.ComparisonUnitArray1 = null;
    insertedCorrelatedSequence.ComparisonUnitArray2 = rightGrouped[iRight].ToArray();
    insertedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Inserted;
    newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
    ++iRight;
}
```

**Purpose:** When Word content doesn't end with a paragraph mark and is followed by a table row, ensure correct ordering (insert table before deleting text, not after).

**Why Critical:** Prevents incorrect diff output where table insertion appears in wrong location relative to text deletion.

**Rust Gap:** Lines 994-1003 handle type mismatches but don't check for paragraph marks or apply priority ordering.

**2. Mixed Tables/Paragraphs Grouping (lines 6526-6625)**

**Current Rust (lines 1052-1096):** Basic table/para grouping
```rust
fn handle_mixed_tables_paragraphs(
    units1: &[ComparisonUnit],
    units2: &[ComparisonUnit],
) -> Vec<CorrelatedSequence> {
    // Basic: Group by "Table" vs "Para", pair same types
    // MISSING: Only processes when both sides have tables AND paragraphs
}
```

**Missing C# Logic:**
- Condition check at line 6527-6529 requires `leftLength > 1 || rightLength > 1`
- Rust version doesn't enforce this condition
- Could cause incorrect behavior for single-element sequences

**Impact:** Medium - edge case handling

---

### **GROUP C: Table-Specific Comparison (P0 - CRITICAL)**

**C# Lines:** 7145-7255 (111 lines)  
**Function:** `DoLcsAlgorithmForTable`  
**Impact:** Core correctness for table comparison  
**Status:** ⚠️ PARTIALLY PORTED (basic skeleton only)

#### Current Rust Implementation (lines 1346-1441):
```rust
fn do_lcs_algorithm_for_table(
    units1: &[ComparisonUnit],
    units2: &[ComparisonUnit],
    _settings: &WmlComparerSettings,
) -> Option<Vec<CorrelatedSequence>> {
    // Basic structure check
    let table1 = units1.first()?.as_group()?;
    let table2 = units2.first()?.as_group()?;
    
    // IMPLEMENTED: CorrelatedSHA1Hash matching (C# lines 7150-7179)
    // IMPLEMENTED: StructureSHA1Hash matching for merged cells (C# lines 7207-7229)
    
    // MISSING: Merged cell detection logic
    // MISSING: Fallback to row-level flattening
    // MISSING: Proper null handling
    
    Some(vec![/* placeholder */])
}
```

#### Missing Logic:

**1. Merged Cell Detection (lines 7197-7204)**
```csharp
var leftContainsMerged = tblElement1
    .Descendants()
    .Any(d => d.Name == W.vMerge || d.Name == W.gridSpan);

var rightContainsMerged = tblElement2
    .Descendants()
    .Any(d => d.Name == W.vMerge || d.Name == W.gridSpan);
```

**Purpose:** Detect if tables contain merged cells (horizontal via `gridSpan`, vertical via `vMerge`).

**Why Critical:** Determines which comparison strategy to use (row-by-row vs flatten-and-mark).

**Rust Status:** ❌ NOT IMPLEMENTED - Need access to XML elements from ComparisonUnitAtom ancestors

**2. Ancestor Element Extraction (lines 7181-7195)**
```csharp
var firstContentAtom1 = tblGroup1.DescendantContentAtoms().FirstOrDefault();
if (firstContentAtom1 == null)
    throw new OpenXmlPowerToolsException("Internal error");
var tblElement1 = firstContentAtom1
    .AncestorElements
    .Reverse()
    .FirstOrDefault(a => a.Name == W.tbl);
```

**Purpose:** Get the actual `<w:tbl>` XML element from the comparison unit to inspect structure.

**Why Critical:** Required to check for `vMerge` and `gridSpan` child elements.

**Rust Status:** ⚠️ PARTIALLY AVAILABLE - `ComparisonUnitAtom` has `ancestors` field but not easily accessible from `ComparisonUnitGroup`

**3. Fallback for Mismatched Merged Cell Tables (lines 7231-7252)**
```csharp
if (leftContainsMerged || rightContainsMerged) {
    if (tblGroup1.StructureSHA1Hash != tblGroup2.StructureSHA1Hash) {
        // Different structures with merged cells - cannot correlate safely
        // Flatten to rows and mark entire table as deleted + inserted
        var deletedCorrelatedSequence = new CorrelatedSequence();
        deletedCorrelatedSequence.ComparisonUnitArray1 = unknown
            .ComparisonUnitArray1
            .Select(z => z.Contents)
            .SelectMany(m => m)
            .ToArray();
        deletedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Deleted;
        
        var insertedCorrelatedSequence = new CorrelatedSequence();
        insertedCorrelatedSequence.ComparisonUnitArray2 = unknown
            .ComparisonUnitArray2
            .Select(z => z.Contents)
            .SelectMany(m => m)
            .ToArray();
        insertedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Inserted;
        
        return newListOfCorrelatedSequence;
    }
}
```

**Purpose:** When both tables have merged cells but different structures, give up on correlation and mark the entire left table as deleted and right table as inserted (after flattening to rows).

**Why Critical:** Prevents incorrect cell-to-cell correlation that would produce wrong diff output.

**Rust Status:** ❌ NOT IMPLEMENTED

**4. Null Return Handling (line 7254)**
```csharp
return null;
```

**Purpose:** Return `null` when table doesn't meet special case criteria, signaling caller to use regular LCS algorithm.

**Why Critical:** Allows fallback to standard comparison for simple tables.

**Rust Status:** ✅ IMPLEMENTED - Uses `Option<Vec<CorrelatedSequence>>` with `None` return

---

### **GROUP D: Paragraph Mark Priority Ordering (P1 - IMPORTANT)**

**C# Lines:** 6827-6910 (84 lines)  
**Impact:** Correctness for mixed Word/Row content  
**Status:** ❌ NOT PORTED

#### Missing Logic:

**1. Word-Ending-Without-ParaMark Before Row Insertion (lines 6827-6848)**
```csharp
if (unknown.ComparisonUnitArray1.Any() && unknown.ComparisonUnitArray2.Any())
{
    var left = unknown.ComparisonUnitArray1.First() as ComparisonUnitWord;
    var right = unknown.ComparisonUnitArray2.First() as ComparisonUnitGroup;
    if (left != null &&
        right != null &&
        right.ComparisonUnitGroupType == ComparisonUnitGroupType.Row)
    {
        // If left side has word content (not ending in paragraph mark) and right has table row
        // Insert the row FIRST, then delete the word
        // This preserves logical document order
        var insertedCorrelatedSequence3 = new CorrelatedSequence();
        insertedCorrelatedSequence3.CorrelationStatus = CorrelationStatus.Inserted;
        insertedCorrelatedSequence3.ComparisonUnitArray2 = unknown.ComparisonUnitArray2;
        newListOfCorrelatedSequence.Add(insertedCorrelatedSequence3);

        var deletedCorrelatedSequence3 = new CorrelatedSequence();
        deletedCorrelatedSequence3.CorrelationStatus = CorrelationStatus.Deleted;
        deletedCorrelatedSequence3.ComparisonUnitArray1 = unknown.ComparisonUnitArray1;
        newListOfCorrelatedSequence.Add(deletedCorrelatedSequence3);

        return newListOfCorrelatedSequence;
    }
}
```

**Purpose:** When left side has word content and right side has a table row, ensure the row insertion appears before the word deletion in the output.

**Why Important:** Maintains logical document flow in diff output - tables should appear before orphaned text.

**Rust Gap:** This logic is not present in any Rust function.

**2. Symmetric Case: Row Before Word-Without-ParaMark (lines 6850-6869)**

Same logic but with sides reversed (row on left, word on right).

**3. Paragraph Mark Asymmetry Handling (lines 6871-6910)**
```csharp
var lastContentAtomLeft = unknown.ComparisonUnitArray1.Select(cu => cu.DescendantContentAtoms().Last()).LastOrDefault();
var lastContentAtomRight = unknown.ComparisonUnitArray2.Select(cu => cu.DescendantContentAtoms().Last()).LastOrDefault();

if (lastContentAtomLeft != null && lastContentAtomRight != null)
{
    if (lastContentAtomLeft.ContentElement.Name == W.pPr &&
        lastContentAtomRight.ContentElement.Name != W.pPr)
    {
        // Left ends with paragraph mark, right doesn't
        // Insert right content FIRST, then delete left
        // (Opposite order from normal to preserve paragraph structure)
    }
    else if (lastContentAtomLeft.ContentElement.Name != W.pPr &&
        lastContentAtomRight.ContentElement.Name == W.pPr)
    {
        // Symmetric case
    }
}
```

**Purpose:** When one side ends with a paragraph mark and the other doesn't, order the Inserted/Deleted sequences appropriately to maintain document structure.

**Why Important:** Ensures paragraph marks are handled correctly when content is asymmetric.

**Rust Gap:** Not implemented.

---

### **GROUP E: Row and Cell Content Handling (P1 - IMPORTANT)**

**C# Lines:** 6663-6826 (164 lines)  
**Impact:** Correctness for table row/cell comparison  
**Status:** ✅ MOSTLY PORTED

#### Current Rust Implementation (lines 1142-1206):
```rust
fn handle_matching_rows(
    units1: &[ComparisonUnit],
    units2: &[ComparisonUnit],
) -> Vec<CorrelatedSequence> {
    // Extract cells from first row on each side
    // Zip cells with padding for length mismatch
    // Create Unknown sequences for each cell pair
    // Handle remaining rows recursively
}
```

**Status:** ✅ Core logic ported

**Minor Missing:** Lines 6760-6767 debug output (not critical)

#### Cell Content Handling (lines 6771-6824)
```csharp
if (firstLeft.ComparisonUnitGroupType == ComparisonUnitGroupType.Cell &&
    firstRight.ComparisonUnitGroupType == ComparisonUnitGroupType.Cell)
{
    // Flatten cell contents and create Unknown sequence
    var left = firstLeft.Contents.ToArray();
    var right = firstRight.Contents.ToArray();
    
    // Create Unknown for cell contents
    // Handle remainder rows
}
```

**Rust Status:** ❌ NOT IMPLEMENTED - Need to add cell-specific handling

**Impact:** Cells are not compared optimally when both sides start with cell group.

---

### **GROUP F: Content Type Analysis and Flattening (P1 - IMPORTANT)**

**C# Lines:** 6636-6661 (26 lines)  
**Impact:** Optimization for hierarchical content  
**Status:** ✅ PORTED (lines 1099-1139)

#### Logic:
```csharp
if (leftOnlyParasTablesTextboxes && rightOnlyParasTablesTextboxes)
{
    // Flatten paragraphs and tables, and iterate
    var left = unknown.ComparisonUnitArray1
        .Select(cu => cu.Contents)
        .SelectMany(m => m)
        .ToArray();
    
    var right = unknown.ComparisonUnitArray2
        .Select(cu => cu.Contents)
        .SelectMany(m => m)
        .ToArray();
    
    // Create single Unknown sequence with flattened content
}
```

**Rust Implementation:**
```rust
fn flatten_and_create_unknown(
    units1: &[ComparisonUnit],
    units2: &[ComparisonUnit],
) -> Vec<CorrelatedSequence> {
    let flattened1: Vec<_> = units1.iter().flat_map(|u| /* extract contents */).collect();
    let flattened2: Vec<_> = units2.iter().flat_map(|u| /* extract contents */).collect();
    vec![CorrelatedSequence::unknown(flattened1, flattened2)]
}
```

**Status:** ✅ FULLY PORTED

---

## Section 3: Priority Ranking of Missing Sections

### **P0: Critical (Must Port for Correctness)**

| Section | C# Lines | Complexity | Impact | Estimated Effort |
|---------|----------|------------|--------|------------------|
| Paragraph boundary preservation | 6927-7130 (204) | High | Broken paragraph correlation | 3-4 days |
| Table merged cell detection | 7197-7204, 7231-7252 | Medium | Wrong table diffs | 2-3 days |
| Paragraph mark priority (mixed content) | 6447-6471 | Medium | Wrong sequence ordering | 2 days |
| FindIndexOfNextParaMark helper | 7133-7143 | Low | Required by para boundary logic | 4 hours |

**Total P0 Effort:** ~8-10 days

### **P1: Important (Should Port for Full Fidelity)**

| Section | C# Lines | Complexity | Impact | Estimated Effort |
|---------|----------|------------|--------|------------------|
| Paragraph mark asymmetry handling | 6871-6910 | Medium | Suboptimal diff ordering | 1-2 days |
| Cell content handling | 6771-6824 | Low | Missing optimization | 1 day |
| Mixed content length check | 6527-6529 | Low | Edge case handling | 2 hours |

**Total P1 Effort:** ~2-3 days

### **P2: Nice-to-Have (Optimizations)**

| Section | C# Lines | Complexity | Impact | Estimated Effort |
|---------|----------|------------|--------|------------------|
| Debug output | 6760-6767 | Low | Development aid | 1 hour |
| Error messages | various | Low | Better diagnostics | 2 hours |

**Total P2 Effort:** ~3 hours

---

## Section 4: Detailed Missing Line Ranges with Explanations

### **4.1 Empty Content Optimization (PORTED ✅)**
- **C# Lines:** 6155-6178
- **Rust Lines:** 644-671
- **Status:** Fully ported

### **4.2 LCS Core Algorithm (PORTED ✅)**
- **C# Lines:** 6180-6221
- **Rust Lines:** 676-703
- **Status:** Fully ported

### **4.3 Paragraph Mark Filtering (PORTED ✅)**
- **C# Lines:** 6223-6257
- **Rust Lines:** 705-722
- **Status:** Fully ported

### **4.4 Single Paragraph Mark Detection (PORTED ✅)**
- **C# Lines:** 6259-6278
- **Rust Lines:** 724-732
- **Status:** Fully ported

### **4.5 Single Space Filtering (PORTED ✅)**
- **C# Lines:** 6280-6293
- **Rust Lines:** 734-743
- **Status:** Fully ported

### **4.6 Word Break Character Filtering (PORTED ✅)**
- **C# Lines:** 6295-6332
- **Rust Lines:** 745-785
- **Status:** Fully ported with CJK support

### **4.7 Detail Threshold (PORTED ✅)**
- **C# Lines:** 6334-6350
- **Rust Lines:** 787-801
- **Status:** Fully ported

### **4.8 Content Type Counting (PORTED ✅)**
- **C# Lines:** 6354-6368
- **Rust Lines:** 858-936 (`count_group_types`)
- **Status:** Fully ported

### **4.9 Mixed Words/Rows/Textboxes Grouping (PARTIAL ⚠️)**
- **C# Lines:** 6381-6524 (144 lines)
- **Rust Lines:** 972-1019 (48 lines)
- **Missing:**
  - Paragraph mark priority checks (lines 6447-6471)
  - Better error handling
- **Impact:** P0 - Affects mixed content correctness

### **4.10 Mixed Tables/Paragraphs Grouping (PARTIAL ⚠️)**
- **C# Lines:** 6526-6625 (100 lines)
- **Rust Lines:** 1052-1096 (45 lines)
- **Missing:**
  - Length > 1 condition check
- **Impact:** P1 - Edge case handling

### **4.11 Single Table Delegation (PORTED ✅)**
- **C# Lines:** 6627-6634
- **Rust Lines:** 882-886
- **Status:** Fully ported

### **4.12 Flatten Paras/Tables/Textboxes (PORTED ✅)**
- **C# Lines:** 6636-6661
- **Rust Lines:** 1099-1139
- **Status:** Fully ported

### **4.13 Row Content Handling (PORTED ✅)**
- **C# Lines:** 6663-6770
- **Rust Lines:** 1142-1206
- **Status:** Core logic ported (debug output skipped)

### **4.14 Cell Content Handling (MISSING ❌)**
- **C# Lines:** 6771-6826 (56 lines)
- **Rust Lines:** N/A
- **Missing:** Entire section
- **Impact:** P1 - Missing optimization for cell-level comparison

### **4.15 Paragraph Mark Priority for Word/Row (MISSING ❌)**
- **C# Lines:** 6827-6869 (43 lines)
- **Rust Lines:** N/A
- **Missing:** Entire section
- **Impact:** P1 - Wrong sequence ordering in mixed content

### **4.16 Paragraph Mark Asymmetry (MISSING ❌)**
- **C# Lines:** 6871-6910 (40 lines)
- **Rust Lines:** N/A
- **Missing:** Entire section
- **Impact:** P1 - Suboptimal diff ordering

### **4.17 Final Fallback: Mark as Deleted+Inserted (PORTED ✅)**
- **C# Lines:** 6912-6924
- **Rust Lines:** 907-911
- **Status:** Fully ported

### **4.18 Paragraph Boundary Preservation (MISSING ❌)**
- **C# Lines:** 6927-7130 (204 lines)
- **Rust Lines:** N/A
- **Missing:** Entire critical section
- **Impact:** P0 - Core correctness issue

#### Breakdown of 4.18:
- **Before-paragraph content** (6989-7028): 40 lines - MISSING ❌
- **Within-paragraph remainder** (7030-7069): 40 lines - MISSING ❌
- **Equal sequence creation** (7071-7081): 11 lines - PORTED ✅ (lines 826-829)
- **Post-LCS paragraph handling** (7095-7122): 28 lines - MISSING ❌
- **Final remainder Unknown** (7124-7128): 5 lines - PORTED ✅ (lines 840-843)

### **4.19 FindIndexOfNextParaMark Helper (MISSING ❌)**
- **C# Lines:** 7133-7143 (11 lines)
- **Rust Lines:** N/A
- **Missing:** Entire helper function
- **Impact:** P0 - Required by paragraph boundary logic

### **4.20 DoLcsAlgorithmForTable (PARTIAL ⚠️)**
- **C# Lines:** 7145-7255 (111 lines)
- **Rust Lines:** 1346-1441 (96 lines)
- **Missing:**
  - Merged cell detection (lines 7197-7204)
  - Ancestor element extraction (lines 7181-7195)
  - Fallback for mismatched structures (lines 7231-7252)
- **Impact:** P0 - Wrong diffs for tables with merged cells

---

## Section 5: Recommendations

### **5.1 Immediate Actions (P0 - Week 1-2)**

1. **Implement Paragraph Boundary Preservation**
   - Port lines 6927-7130 from C#
   - Add `FindIndexOfNextParaMark` helper (lines 7133-7143)
   - Add tests for paragraph-spanning changes
   - **Acceptance Criteria:** Diffs respect paragraph boundaries, no cross-paragraph correlation

2. **Complete Table Merged Cell Handling**
   - Port merged cell detection (lines 7197-7204)
   - Implement ancestor element extraction (lines 7181-7195)
   - Add fallback for structure mismatch (lines 7231-7252)
   - **Acceptance Criteria:** Tables with merged cells produce correct diffs

3. **Add Paragraph Mark Priority to Mixed Content**
   - Port paragraph mark checks in word/row grouping (lines 6447-6471)
   - **Acceptance Criteria:** Mixed word/row content produces correctly ordered diffs

### **5.2 Short-Term Enhancements (P1 - Week 3-4)**

1. **Implement Paragraph Mark Asymmetry Handling**
   - Port lines 6871-6910
   - Add tests for asymmetric paragraph endings
   - **Acceptance Criteria:** Diffs with asymmetric paragraph marks produce optimal ordering

2. **Add Cell Content Handling**
   - Port lines 6771-6824
   - **Acceptance Criteria:** Cell-level comparisons are optimized

3. **Add Length Condition Check**
   - Port condition from lines 6527-6529 to mixed table/para handler
   - **Acceptance Criteria:** Edge case handled correctly

### **5.3 Long-Term Improvements (P2 - Future)**

1. **Add Debug Output** (optional)
   - Port debug output from lines 6760-6767
   - Use Rust feature flags to enable/disable

2. **Enhance Error Messages**
   - Add context to error messages
   - Use `thiserror` or `anyhow` for better error types

### **5.4 Testing Strategy**

For each ported section, add:

1. **Unit Tests:**
   - Test the specific function in isolation
   - Cover edge cases (empty arrays, single elements, etc.)

2. **Integration Tests:**
   - Create DOCX test files with specific scenarios
   - Run full comparison and verify output
   - Use golden file testing (compare to C# output)

3. **Property-Based Tests:**
   - Use `proptest` to generate random inputs
   - Verify invariants (e.g., paragraph boundaries preserved)

4. **Regression Tests:**
   - Port existing C# test cases from `WmlComparerTests.cs`
   - Ensure byte-for-byte compatibility where possible

---

## Section 6: Complexity Metrics

### **C# DoLcsAlgorithm Complexity:**
- **Total Lines:** 997
- **Cyclomatic Complexity:** ~45 (estimated from nesting depth)
- **Max Nesting Depth:** 6 levels
- **Number of Branches:** 28 major branches
- **Number of Loops:** 8 loops (for, while, foreach)
- **LINQ Chains:** 25+ complex LINQ expressions
- **Helper Functions:** 1 (FindIndexOfNextParaMark)

### **Rust do_lcs_algorithm Complexity:**
- **Total Lines:** 211 (main function) + ~600 (helpers) = ~811 lines total
- **Cyclomatic Complexity:** ~25 (estimated)
- **Max Nesting Depth:** 4 levels
- **Number of Branches:** 15 major branches
- **Number of Loops:** 4 loops
- **Iterator Chains:** 10+ iterator chains
- **Helper Functions:** 8 (count_group_types, handle_mixed_words_rows_textboxes, etc.)

### **Gap Metrics:**
- **Missing Lines:** ~400-500 lines (40-50% of total logic)
- **Missing Branches:** ~13 major branches
- **Missing Functions:** 1 helper function
- **Complexity Reduction:** ~44% (fewer branches, simpler nesting)

**Analysis:** The Rust port has simplified some complex C# LINQ chains into separate helper functions, reducing complexity per function but requiring more total functions. The missing logic represents critical path sections that cannot be omitted.

---

## Section 7: Conclusion

### **Summary:**
The Rust port of DoLcsAlgorithm has successfully implemented the core LCS computation and basic filtering logic (~350 lines, 35%), but is **missing critical sections totaling ~400-500 lines (40-50%)** that handle:

1. Paragraph boundary preservation (P0)
2. Table merged cell handling (P0)
3. Content type priority ordering (P0-P1)
4. Cell content optimization (P1)

### **Critical Path Forward:**
To achieve correctness parity with the C# implementation, the **P0 sections must be ported within 2 weeks**. This represents approximately **8-10 days of engineering effort** plus testing.

### **Risk Assessment:**
- **HIGH RISK:** Shipping without paragraph boundary preservation will produce incorrect diffs for paragraph-spanning changes
- **HIGH RISK:** Shipping without merged cell handling will produce incorrect diffs for complex tables
- **MEDIUM RISK:** Shipping without priority ordering will produce suboptimal (but possibly correct) diffs

### **Next Steps:**
1. Review this analysis with the team
2. Prioritize P0 sections for immediate implementation
3. Create tracking issues for each missing section
4. Begin porting starting with paragraph boundary preservation
5. Add comprehensive test coverage for each ported section

---

---

## APPENDIX A: Enhanced Line-by-Line Cross-Reference

*Added by comprehensive deep analysis pass (December 26, 2025)*

This appendix provides an enhanced, section-by-section mapping between C# and Rust with exact line references and detailed notes on what's ported vs missing.

### A.1 Core Algorithm Structure Comparison

| Phase | C# Lines | Rust Lines | Status | Notes |
|-------|----------|------------|--------|-------|
| **Empty case handling** | 6155-6178 (24 lines) | 644-671 (28 lines) | ✅ 100% | Rust uses pattern matching, more verbose but clearer |
| **LCS discovery loop** | 6180-6221 (42 lines) | 676-703 (28 lines) | ✅ 100% | Identical algorithm, Rust more concise |
| **Paragraph mark start filter** | 6223-6257 (35 lines) | 705-722 (18 lines) | ✅ 100% | Rust uses `is_paragraph_mark()` helper |
| **Only para mark check** | 6259-6278 (20 lines) | 724-732 (9 lines) | ✅ 100% | Simplified in Rust |
| **Single space filter** | 6280-6293 (14 lines) | 734-743 (10 lines) | ✅ 100% | Direct port |
| **Word break filter** | 6295-6332 (38 lines) | 745-785 (41 lines) | ✅ 98% | Includes CJK 0x4e00-0x9fff check |
| **Detail threshold** | 6334-6350 (17 lines) | 787-801 (15 lines) | ✅ 100% | Direct port |
| **No match delegation** | 6352 | 804-806 (3 lines) | ✅ | Calls `handle_no_match_cases()` |

**Subtotal: Core Algorithm - 190 C# lines → 152 Rust lines (✅ 100% coverage)**

---

### A.2 No-Match Edge Cases (Critical Gap Area)

| Subcase | C# Lines | Rust Lines | Status | Notes |
|---------|----------|------------|--------|-------|
| **Type counting** | 6354-6368 (15 lines) | 858-936 (79 lines) | ✅ 100% | Extracted to `count_group_types()` |
| **Words/Rows/Textboxes: Basic grouping** | 6387-6424 (38 lines) | 976-977 (2 lines) | ✅ 80% | Uses `group_units_by_type()` |
| **Words/Rows/Textboxes: Para mark checks** | 6447-6471 (25 lines) | ❌ MISSING | ❌ 0% | **CRITICAL GAP** |
| **Words/Rows/Textboxes: Basic matching** | 6473-6491 (19 lines) | 983-1003 (21 lines) | ⚠️ 70% | Simplified logic |
| **Words/Rows/Textboxes: Cleanup** | 6493-6523 (31 lines) | 1006-1018 (13 lines) | ✅ 90% | Equivalent |
| **Tables/Paras: Grouping** | 6531-6552 (22 lines) | 1056-1057 (2 lines) | ✅ 90% | Uses `group_units_table_para()` |
| **Tables/Paras: Matching loop** | 6562-6593 (32 lines) | 1063-1081 (19 lines) | ✅ 95% | Direct port |
| **Tables/Paras: Cleanup** | 6594-6624 (31 lines) | 1083-1095 (13 lines) | ✅ 95% | Equivalent |
| **Table delegation** | 6627-6634 (8 lines) | 882-886 (5 lines) | ✅ 100% | Calls helper |
| **Flatten paras/tables** | 6636-6661 (26 lines) | 1099-1139 (41 lines) | ✅ 100% | More verbose in Rust |
| **Row matching** | 6663-6770 (108 lines) | 1142-1206 (65 lines) | ⚠️ 85% | **Missing null padding** |
| **Cell matching** | 6771-6824 (54 lines) | ❌ MISSING | ❌ 0% | **COMPLETE GAP** |
| **Word/Row conflict** | 6827-6869 (43 lines) | ❌ MISSING | ❌ 0% | **COMPLETE GAP** |
| **Para mark asymmetry** | 6871-6910 (40 lines) | ❌ MISSING | ❌ 0% | **COMPLETE GAP** |
| **Final fallback** | 6912-6924 (13 lines) | 907-911 (5 lines) | ✅ 100% | Direct port |

**Subtotal: No-Match Cases - 505 C# lines → ~268 Rust lines (⚠️ ~53% coverage, critical gaps)**

---

### A.3 Match Found - Result Construction (CRITICAL GAP)

| Section | C# Lines | Rust Lines | Status | Notes |
|---------|----------|------------|--------|-------|
| **Calculate remaining in para (left)** | 6936-6970 (35 lines) | ❌ MISSING | ❌ 0% | **P0 CRITICAL** |
| **Calculate remaining in para (right)** | 6971-6984 (14 lines) | ❌ MISSING | ❌ 0% | **P0 CRITICAL** |
| **Before-para prefix handling** | 6989-7028 (40 lines) | ❌ MISSING | ❌ 0% | **P0 CRITICAL** |
| **Within-para remainder** | 7030-7069 (40 lines) | ❌ MISSING | ❌ 0% | **P0 CRITICAL** |
| **Middle equal sequence** | 7071-7081 (11 lines) | 826-829 (4 lines) | ✅ 100% | Simplified |
| **Calculate remainder arrays** | 7084-7093 (10 lines) | 832-833 (2 lines) | ✅ 100% | Direct port |
| **Post-LCS para boundary check** | 7095-7122 (28 lines) | ❌ MISSING | ❌ 0% | **P0 CRITICAL** |
| **Final remainder unknown** | 7124-7128 (5 lines) | 839-843 (5 lines) | ⚠️ 50% | No para boundary split |

**Subtotal: Result Construction - 183 C# lines → 11 Rust lines (❌ ~6% coverage, CRITICAL)**

**Rust Implementation (Simplified):**
```rust
// Lines 813-846 - Simple prefix/match/suffix split
// MISSING: All paragraph boundary awareness
if best_i1 > 0 && best_i2 > 0 {
    result.push(CorrelatedSequence::unknown(
        units1[..best_i1].to_vec(),
        units2[..best_i2].to_vec(),
    )); // NO paragraph boundary detection
}
result.push(CorrelatedSequence::equal(/* match */));
if end_i1 < units1.len() && end_i2 < units2.len() {
    result.push(CorrelatedSequence::unknown(
        units1[end_i1..].to_vec(),
        units2[end_i2..].to_vec(),
    )); // NO paragraph boundary detection
}
```

---

### A.4 Helper Functions

| Helper | C# Lines | Rust Lines | Status | Notes |
|--------|----------|------------|--------|-------|
| **FindIndexOfNextParaMark** | 7133-7143 (11 lines) | ❌ MISSING | ❌ 0% | **P0 - Required by para boundary logic** |
| **DoLcsAlgorithmForTable** | 7145-7255 (111 lines) | 1346-1441 (96 lines) | ⚠️ 60% | Stub merged cell detection |

**DoLcsAlgorithmForTable Breakdown:**

| Subsection | C# Lines | Rust Lines | Status |
|------------|----------|------------|--------|
| Extract table groups | 7147-7150 | 1350-1357 | ✅ 100% |
| Row count/hash check | 7150-7179 | 1368-1388 | ✅ 100% |
| Get table XML elements | 7181-7195 | ❌ MISSING | ❌ 0% |
| Merged cell detection | 7197-7204 | 1392-1393 | ❌ STUB |
| Structure hash check | 7207-7229 | 1396-1409 | ✅ 100% |
| Fallback flatten | 7231-7252 | 1412-1424 | ⚠️ 70% |
| Null return | 7254 | 1427 | ✅ 100% |

**Subtotal: Helpers - 122 C# lines → 96 Rust lines (⚠️ ~55% coverage)**

---

### A.5 Complete Coverage Summary

| Major Section | C# Lines | Rust Lines | Coverage | P0 Issues |
|---------------|----------|------------|----------|-----------|
| Core Algorithm | 190 | 152 | ✅ 100% | None |
| No-Match Cases | 505 | 268 | ⚠️ 53% | Para mark checks, cell handling, word/row conflicts |
| Result Construction | 183 | 11 | ❌ 6% | **Entire paragraph boundary preservation** |
| Helpers | 122 | 96 | ⚠️ 55% | FindIndexOfNextParaMark, merged cells |
| **TOTAL** | **1,000** | **527** | **⚠️ 53%** | **~473 lines missing (47%)** |

---

## APPENDIX B: Critical Missing Code Blocks

*These code blocks represent P0 functionality that MUST be ported.*

### B.1 Paragraph Alignment - Calculate Remaining (C# 6936-6987)

**C# Code:**
```csharp
int remainingInLeftParagraph = 0;
int remainingInRightParagraph = 0;
if (currentLongestCommonSequenceLength != 0)
{
    var commonSeq = unknown
        .ComparisonUnitArray1
        .Skip(currentI1)
        .Take(currentLongestCommonSequenceLength)
        .ToList();
    var firstOfCommonSeq = commonSeq.First();
    if (firstOfCommonSeq is ComparisonUnitWord)
    {
        // are there any paragraph marks in the common seq at end?
        if (commonSeq.Any(cu =>
        {
            var firstComparisonUnitAtom = cu.Contents.OfType<ComparisonUnitAtom>().FirstOrDefault();
            if (firstComparisonUnitAtom == null)
                return false;
            return firstComparisonUnitAtom.ContentElement.Name == W.pPr;
        }))
        {
            remainingInLeftParagraph = unknown
                .ComparisonUnitArray1
                .Take(currentI1)
                .Reverse()
                .TakeWhile(cu =>
                {
                    if (!(cu is ComparisonUnitWord))
                        return false;
                    var firstComparisonUnitAtom = cu.Contents.OfType<ComparisonUnitAtom>().FirstOrDefault();
                    if (firstComparisonUnitAtom == null)
                        return true;
                    return firstComparisonUnitAtom.ContentElement.Name != W.pPr;
                })
                .Count();
            remainingInRightParagraph = unknown
                .ComparisonUnitArray2
                .Take(currentI2)
                .Reverse()
                .TakeWhile(cu =>
                {
                    if (!(cu is ComparisonUnitWord))
                        return false;
                    var firstComparisonUnitAtom = cu.Contents.OfType<ComparisonUnitAtom>().FirstOrDefault();
                    if (firstComparisonUnitAtom == null)
                        return true;
                    return firstComparisonUnitAtom.ContentElement.Name != W.pPr;
                })
                .Count();
        }
    }
}
```

**Required Rust Implementation:**
```rust
fn calculate_remaining_in_paragraph(
    units1: &[ComparisonUnit],
    units2: &[ComparisonUnit],
    best_i1: usize,
    best_i2: usize,
    best_length: usize,
) -> (usize, usize) {
    let mut remaining_in_left = 0;
    let mut remaining_in_right = 0;
    
    if best_length == 0 {
        return (0, 0);
    }
    
    // Get common sequence
    let common_seq: Vec<_> = units1[best_i1..best_i1 + best_length].to_vec();
    
    // Check if first element is a word
    if let Some(first_word) = common_seq.first().and_then(|u| u.as_word()) {
        // Check if any unit in common_seq has paragraph mark
        let has_para_mark = common_seq.iter().any(|cu| {
            if let Some(word) = cu.as_word() {
                if let Some(first_atom) = word.atoms.first() {
                    return matches!(first_atom.content_element, ContentElement::ParagraphProperties);
                }
            }
            false
        });
        
        if has_para_mark {
            // Count units before best_i1 that are in same paragraph (reverse scan until para mark)
            remaining_in_left = units1[..best_i1]
                .iter()
                .rev()
                .take_while(|cu| {
                    if let Some(word) = cu.as_word() {
                        if let Some(first_atom) = word.atoms.first() {
                            return !matches!(first_atom.content_element, ContentElement::ParagraphProperties);
                        }
                        return true; // No atoms means continue
                    }
                    false // Not a word means stop
                })
                .count();
            
            // Same for right side
            remaining_in_right = units2[..best_i2]
                .iter()
                .rev()
                .take_while(|cu| {
                    if let Some(word) = cu.as_word() {
                        if let Some(first_atom) = word.atoms.first() {
                            return !matches!(first_atom.content_element, ContentElement::ParagraphProperties);
                        }
                        return true;
                    }
                    false
                })
                .count();
        }
    }
    
    (remaining_in_left, remaining_in_right)
}
```

---

### B.2 FindIndexOfNextParaMark Helper (C# 7133-7143)

**C# Code:**
```csharp
private static int FindIndexOfNextParaMark(ComparisonUnit[] cul)
{
    for (int i = 0; i < cul.Length; i++)
    {
        var cuw = cul[i] as ComparisonUnitWord;
        var lastAtom = cuw.DescendantContentAtoms().LastOrDefault();
        if (lastAtom.ContentElement.Name == W.pPr)
            return i;
    }
    return cul.Length;
}
```

**Required Rust Implementation:**
```rust
fn find_index_of_next_para_mark(units: &[ComparisonUnit]) -> usize {
    for (i, unit) in units.iter().enumerate() {
        if let Some(word) = unit.as_word() {
            if let Some(last_atom) = word.descendant_atoms().last() {
                if matches!(last_atom.content_element, ContentElement::ParagraphProperties) {
                    return i;
                }
            }
        }
    }
    units.len()
}
```

---

### B.3 Mixed Content Paragraph Mark Priority (C# 6447-6471)

**C# Code:**
```csharp
else if (leftGrouped[iLeft].Key == "Word" &&
    leftGrouped[iLeft].Select(lg => lg.DescendantContentAtoms()).SelectMany(m => m).Last().ContentElement.Name != W.pPr &&
    rightGrouped[iRight].Key == "Row")
{
    var insertedCorrelatedSequence = new CorrelatedSequence();
    insertedCorrelatedSequence.ComparisonUnitArray1 = null;
    insertedCorrelatedSequence.ComparisonUnitArray2 = rightGrouped[iRight].ToArray();
    insertedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Inserted;
    newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
    ++iRight;
}
else if (rightGrouped[iRight].Key == "Word" &&
    rightGrouped[iRight].Select(lg => lg.DescendantContentAtoms()).SelectMany(m => m).Last().ContentElement.Name != W.pPr &&
    leftGrouped[iLeft].Key == "Row")
{
    var insertedCorrelatedSequence = new CorrelatedSequence();
    insertedCorrelatedSequence.ComparisonUnitArray1 = null;
    insertedCorrelatedSequence.ComparisonUnitArray2 = leftGrouped[iLeft].ToArray();
    insertedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Inserted;
    newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
    ++iLeft;
}
```

**Required Rust Enhancement to handle_mixed_words_rows_textboxes():**
```rust
// Add after line 990 in handle_mixed_words_rows_textboxes()
// Check for paragraph mark priority when types don't match

// Word (left) without para mark vs Row (right)
else if *key1 == "Word" && *key2 == "Row" {
    // Check if word group ends with paragraph mark
    let ends_with_para_mark = items1.iter()
        .flat_map(|u| u.descendant_atoms())
        .last()
        .map(|atom| matches!(atom.content_element, ContentElement::ParagraphProperties))
        .unwrap_or(false);
    
    if !ends_with_para_mark {
        // Insert row BEFORE processing word
        result.push(CorrelatedSequence::inserted(items2.clone()));
        i2 += 1;
        continue; // Don't increment i1, process word on next iteration
    }
    
    // Otherwise fall through to normal word deletion
    result.push(CorrelatedSequence::deleted(items1.clone()));
    i1 += 1;
}
// Symmetric case: Row (left) vs Word (right) without para mark
else if *key1 == "Row" && *key2 == "Word" {
    let ends_with_para_mark = items2.iter()
        .flat_map(|u| u.descendant_atoms())
        .last()
        .map(|atom| matches!(atom.content_element, ContentElement::ParagraphProperties))
        .unwrap_or(false);
    
    if !ends_with_para_mark {
        // Insert row BEFORE processing word
        result.push(CorrelatedSequence::inserted(items1.clone()));
        i1 += 1;
        continue;
    }
    
    result.push(CorrelatedSequence::inserted(items2.clone()));
    i2 += 1;
}
```

---

## APPENDIX C: Updated Effort Estimates

*Based on enhanced line-by-line analysis*

### P0 Critical Implementation (Must Complete)

| Task | Lines to Port | Complexity | Estimated Effort |
|------|---------------|------------|------------------|
| Paragraph boundary preservation core | ~204 lines | High | 4-5 days |
| - Calculate remaining in paragraph | ~50 lines | Medium | 1 day |
| - Before-paragraph handling | ~40 lines | Medium | 1 day |
| - Within-paragraph handling | ~40 lines | Medium | 1 day |
| - Post-LCS boundary detection | ~30 lines | High | 1-2 days |
| FindIndexOfNextParaMark helper | ~11 lines | Low | 0.5 days |
| Mixed content para mark priority | ~25 lines | Medium | 1-2 days |
| Table merged cell detection | ~50 lines | High | 2-3 days |
| **TOTAL P0** | **~340 lines** | - | **9-12 days** |

### P1 Important Implementation (Should Complete)

| Task | Lines to Port | Complexity | Estimated Effort |
|------|---------------|------------|------------------|
| Cell content handling | ~54 lines | Medium | 1-2 days |
| Word/Row conflict resolution | ~43 lines | Medium | 1 day |
| Para mark asymmetry | ~40 lines | Medium | 1 day |
| Row cell padding | ~15 lines | Low | 0.5 days |
| **TOTAL P1** | **~152 lines** | - | **3.5-4.5 days** |

### Combined Effort
- **P0 + P1 Total:** ~492 lines, **12.5-16.5 days** of implementation
- **Testing & Integration:** +5-7 days
- **Total Project Time:** **17.5-23.5 days** (~3.5-5 weeks)

---

**Document Version:** 2.0 (Enhanced)  
**Last Updated:** December 26, 2025 - 21:45 UTC  
**Prepared by:** Swarm Analysis System (10+ parallel research agents) + Deep Analysis Enhancement
