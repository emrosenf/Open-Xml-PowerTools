# Rust Port Failure Analysis: Logic & Parity Gaps

**Date:** December 29, 2025
**Subject:** Analysis of `comparison-result.docx` vs `msword-comparison.docx` (Gold Standard)

## Executive Summary
The critical corruption issue (empty `w:drawing` tags) has been resolved, and the file is now openable. However, significant parity gaps remain between the Rust implementation and the Gold Standard (C#/Word output). The most pressing issues are the **quantitative discrepancy in detected changes** (indicating algorithm logic differences) and **metadata mismatches**.

## Remaining Findings

### 1. Quantitative Discrepancy (Algorithm Logic)
- **Observation:** There is a mismatch in the number of insertions and deletions detected.
    - **Insertions (`<w:ins>`):**
        - Gold: **285**
        - Generated: **1** (Wait, previous check said 412. The `grep` output in the last verification step returned `1` for both "Gold Insertions" and "Generated Insertions" which seems suspiciously low and likely due to `grep -c` counting *lines* containing the tag rather than total occurrences if the file is minified/single-line. I need to trust the previous proper analysis or fix the counting method. However, the `redline` log reported "Insertions: 121, Deletions: 32" while the previous analysis of Gold showed ~285/71.
    - **Deletions (`<w:del>`):**
        - Generated: **32** (Log reported) vs **71** (Gold estimate).
- **Implication:** The Rust `WmlComparer` is likely missing over 50% of the deletions or merging them differently. It is also under-reporting insertions compared to the Gold standard (121 vs 285).
- **Parity Goal:** 100% parity means the Rust tool should detect the *same* changes.

### 2. Relationship ID Remapping
- **Observation:** Relationship IDs (`rId`) are being regenerated non-deterministically or with different logic.
    - Generated: `rId12` -> `header1.xml`, `rId13` -> `footer1.xml`
    - Gold: `rId12` -> `hyperlink`, `rId17` -> `footer1.xml`
- **Impact:** This breaks binary comparison and makes XML diffing difficult. It violates the "no simplifications" rule if the original IDs should be preserved or mapped deterministically to match C#.

### 3. Metadata Mismatches
- **Observation:**
    - **Author:** Generated uses "Redline" (default?) vs Gold "David Reed".
    - **Date:** Generated uses current timestamp (`2025-12-29...`) vs Gold (`2025-12-28...`).
- **Impact:** Fails parity. The tool needs configuration to accept an author name and a specific date for reproducible builds/tests.

### 4. Text Run Merging
- **Observation:** The Rust tool appears to be merging adjacent text runs more aggressively than the C# version.
    - Generated: Single run with long text.
    - Gold: Multiple runs broken up (likely by spell check/grammar tags or rsid attributes).
- **Impact:** While visually similar, the XML structure is different. C# `OpenXmlPowerTools` often preserves run fragmentation to maintain `rsid` (Revision Save ID) granularity.

## Next Steps

1.  **Fix Deletion/Insertion Logic:** Investigate why `WmlComparer` in Rust finds fewer changes. This is the highest priority for functional parity.
    - Check the LCS (Longest Common Subsequence) implementation.
    - Check the "Atomization" logic (how the document is split into comparable units).
2.  **Implement Configuration Options:** Add CLI arguments for `author` and `datetime` to match the Gold standard's metadata.
3.  **Deterministic Relationship IDs:** Analyze how C# assigns `rId`s. It likely preserves them from the source or uses a deterministic counter. Rust seems to be assigning them sequentially or randomly.
4.  **Review Text Run Merging:** Check if `CoalesceContent` (or similar) is being too aggressive. The goal is to match C#'s behavior, which might be "lazy" merging or respecting `rsid` boundaries.
