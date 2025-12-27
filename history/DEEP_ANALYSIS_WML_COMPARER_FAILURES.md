# DEEP ANALYSIS REPORT: WmlComparer Test Failure Categories

**Date**: December 26, 2025
**Scope**: Image/Drawing failures, Revision Counting edge cases, French Locale failures.

---

## 1. Executive Summary

A deep, multi-agent analysis has identified that current failures in the WmlComparer ports (Rust and TypeScript) are not merely isolated bugs but symptoms of an incomplete architectural pipeline. To achieve 100% parity with the C# reference and pass the remaining 104 golden tests, the implementation must move beyond "translation-first" to "invariant-first" by completing the normalization, event classification, and locale abstraction layers.

---

## 2. Detailed Findings

### **A. Image/Drawing Failures (wc_1260, wc_1450, wc_1940)**
*   **Root Cause**: Missing `CoalesceAdjacentRunsWithIdenticalFormatting` step in the Rust port.
*   **Mechanism**: In C#, images (`w:drawing`, `w:pict`) are treated as atomic units with SHA1 hashes. The pipeline ensures a consolidated run structure before correlation. Without this consolidation, adjacent revisions with identical formatting are treated as separate edits, leading to inflated revision counts (e.g., 2 revisions instead of 1).
*   **Technical Gap**: The Rust port has `ContentElement::Drawing` but lacks the downstream markup generation pipeline (`produce_markup_from_atoms`, `mark_content_as_deleted_or_inserted`, and the consolidation pass).

### **B. Revision Counting Edge Cases (wc_1180, wc_2040)**
*   **Root Cause**: Incomplete change classification logic (missing similarity thresholds and formatting tracking).
*   **Mechanism**:
    *   **wc_1180**: Rust splits "deleted at beginning of paragraph" into separate insert/delete events. C# uses similarity thresholds (0.4/0.5) to group these as a single "modification".
    *   **wc_2040**: Rust misses formatting-only revisions. It currently ignores `ComparisonCorrelationStatus::FormatChanged` in the counting loop and doesn't extract `rPrChange` elements.
*   **Technical Gap**: The counting loop in `comparer.rs` needs to be upgraded to a state-aware classification system that groups adjacent changes by author/date/type and applies similarity gates.

### **C. French Locale Failures (wc_1970, wc_1980)**
*   **Root Cause**: Missing locale-specific list item text generation in the TypeScript port.
*   **Mechanism**: C# implements `GetListItemText_fr_FR.cs` which handles complex French numbering (e.g., "soixante-dix" for 70, pluralization of "quatre-vingts"). The TypeScript port currently falls back to English numbering, causing mismatches in documents using `cardinalText` or `ordinalText`.
*   **Technical Gap**: Lack of a `LocaleProvider` abstraction in `redline-js` to route specific locale identifiers to their respective numbering implementations.

---

## 3. Oracle Strategic Analysis

### **Architecture Assessment**
The system currently compares **partially-normalized structures**. This creates representational ambiguity where identical content can yield different atom streams depending on incidental run boundaries.

### **Key Invariants to Preserve**
1.  **Canonicalization**: Equivalent OpenXML markup must map to the same internal atom stream.
2.  **Stable Atom Identity**: Non-text objects must use stable hashes of canonical XML.
3.  **Partition of Change Categories**: Every change must belong to exactly one category (Insert, Delete, Replace, Format).
4.  **Locale Isolation**: Formatting logic must be isolated behind provider interfaces.

### **Performance Considerations**
*   **Linear Scaling**: All fixes can be implemented in O(n) passes.
*   **Risk**: Avoid deep XML equality checks in hot loops; use interning for formatting keys and SHA1 content hashing.
*   **Caching**: Caching image hashes and memoizing locale text is essential for documents >100 pages.

---

## 4. Recommended Action Plan (Minimal Architecture-First Path)

| Step | Task | Target | Effort |
| :--- | :--- | :--- | :--- |
| 1 | **Canonical Atom Model** | redline-rs | 0.5d |
| 2 | **Normalization Stage** (CoalesceAdjacentRuns) | redline-rs | 1.5d |
| 3 | **Drawing Canonical Identity** (SHA1 Hash) | redline-rs | 1d |
| 4 | **Explicit ChangeEvents** (FormatChange support) | redline-rs | 1d |
| 5 | **Similarity Classification** (Grouping logic) | redline-rs | 1.5d |
| 6 | **Locale Provider + French Logic** | redline-js | 1.5d |

**Total Estimated Effort**: 7-9 days for full production-ready stability.

---

## 5. Verification Strategy
*   **Differential Testing**: Use the C# WmlComparer as a "live oracle" to compare normalized output counts.
*   **Metamorphic Testing**: Verify that AlternateContent swaps and relationship renames do not change the diff outcome.
*   **Boundary Testing**: Validate French 70/80/90 rules and nested textbox structures.

---

*Report synthesized by SilverStar Research Agent via multi-agent swarm analysis.*
