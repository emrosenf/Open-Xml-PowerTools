# XML Whitespace Discrepancies Investigation

## Summary
Investigated XML differences between MS Word gold standard (msword-comparison-2025-12-28.docx) and Rust output (comparison-result.docx) for sections 3.4, 4, and 4.1.

**Key Finding:** Rust output is missing markup/elements that MS Word includes, resulting in simpler but less faithful reconstruction.

---

## Finding #1: Section 3.4 - Extra Empty Paragraphs in Rust

### Issue
After "3.4 Intentionally Deleted", Rust generates **4 extra empty paragraphs** (`<w:p>` elements) that are NOT present in the gold standard.

### RUST Output (section 3.4)
```xml
<w:t>3.4  Intentionally Deleted</w:t>
</w:r>
<w:r>
  <w:rPr><w:iCs/><w:szCs w:val="24"/></w:rPr>
  <w:t xml:space="preserve">.  </w:t>
</w:r>
</w:p>

<!-- HERE: 4 EXTRA EMPTY PARAGRAPHS -->
<w:p w14:paraId="6E472136" w14:textId="2622D11A" w:rsidR="00163273" w:rsidDel="008448C0" w:rsidRDefault="00163273" w:rsidP="000F7098"/>
<w:p w14:paraId="1B5C0135" w14:textId="06EE693C" w:rsidR="00163273" w:rsidDel="008448C0" w:rsidRDefault="00163273" w:rsidP="000F7098"/>
<w:p w14:paraId="3B65750C" w14:textId="0347D155" w:rsidR="00163273" w:rsidDel="008448C0" w:rsidRDefault="00163273" w:rsidP="000F7098"/>
<w:p w14:paraId="5A96C747" w14:textId="1A1D5211" w:rsidR="00873DE2" w:rsidDel="008448C0" w:rsidRDefault="00163273" w:rsidP="00163273"/>

<w:p w14:paraId="6567C615" w14:textId="77777777" w:rsidR="00AF48AA" w:rsidRPr="003157DF" w:rsidRDefault="00AF48AA" w:rsidP="00163273">
  <w:pPr><w:jc w:val="both"/><w:rPr><w:iCs/><w:szCs w:val="24"/></w:rPr></w:pPr>
</w:p>
```

### GOLD Output (section 3.4)
```xml
<w:t>3.4  </w:t>
</w:r>
<w:r w:rsidR="008448C0">
  <w:rPr><w:b/><w:bCs/><w:iCs/><w:szCs w:val="24"/></w:rPr>
  <w:t>Intentionally</w:t>
</w:r>
<w:proofErr w:type="gramEnd"/>
<w:r w:rsidR="008448C0">
  <w:rPr><w:b/><w:bCs/><w:iCs/><w:szCs w:val="24"/></w:rPr>
  <w:t xml:space="preserve"> Deleted</w:t>
</w:r>
<w:r>
  <w:rPr><w:iCs/><w:szCs w:val="24"/></w:rPr>
  <w:t xml:space="preserve">.  </w:t>
</w:r>
</w:p>

<!-- NO EXTRA PARAGRAPHS - Goes directly to section 3.5 -->
<w:p w14:paraId="6567C615" w14:textId="77777777" w:rsidR="00AF48AA" w:rsidRPr="003157DF" w:rsidRDefault="00AF48AA" w:rsidP="00163273">
  <w:pPr><w:jc w:val="both"/><w:rPr><w:iCs/><w:szCs w:val="24"/></w:rPr></w:pPr>
</w:p>
```

**Additional differences noticed:**
- GOLD splits "Intentionally Deleted" into separate runs with different formatting
- GOLD includes `<w:proofErr>` elements (grammar checking markup)
- RUST merges all text into single run "3.4  Intentionally Deleted"

---

## Finding #2: Missing Track Changes Markup

### Count Discrepancy
- **GOLD:** 62 deletion (`<w:del>`) elements
- **RUST:** 34 deletion elements

**This means Rust is missing 28 deletions (45% of tracked changes lost!)**

---

## Finding #3: Section 3.5 - Text Splitting Differences  

### RUST Output (section 3.5)
```xml
<w:r w:rsidR="00356315" w:rsidRPr="000F7098">
  <w:t>.  If any payments, rights or obligations hereunder (whether relating to payment of Rent, Taxes, insurance, other impositions, or to any other provision of this Lease) relate to a period in part before the Commencement Date or in part after the date of expiration or termination of the Term, appropriate adjustments and prorations shall be made.</w:t>
</w:r>
```

### GOLD Output (section 3.5)
```xml
<w:r w:rsidR="00356315" w:rsidRPr="000F7098">
  <w:t xml:space="preserve">.  If any payments, rights or obligations hereunder (whether relating to payment of </w:t>
</w:r>
<w:r w:rsidR="00241C96">
  <w:t>R</w:t>
</w:r>
<w:r w:rsidR="00356315" w:rsidRPr="000F7098">
  <w:t xml:space="preserve">ent, </w:t>
</w:r>
<w:r w:rsidR="00241C96">
  <w:t>T</w:t>
</w:r>
<w:r w:rsidR="00356315" w:rsidRPr="000F7098">
  <w:t xml:space="preserve">axes, insurance, other impositions, or to any other provision of this Lease) relate to a period in part before the Commencement Date or in part after the date of expiration or termination of the </w:t>
</w:r>
<w:r w:rsidR="00461784">
  <w:t>Term</w:t>
</w:r>
<w:r w:rsidR="00356315" w:rsidRPr="000F7098">
  <w:t>, appropriate adjustments and prorations shall be made.</w:t>
</w:r>
```

**Notice:** GOLD preserves individual character edits (R, T, Term in separate runs). RUST merges everything.

---

## Root Cause Analysis

### Likely Culprit: `coalesce.rs`

The coalescing logic is aggressively merging runs that should be kept separate. From the context:

**Location:** `crates/redline-core/src/wml/coalesce.rs`

**Problem Areas:**
1. **Lines 1543-1557:** Whitespace reconstruction from atoms
   - May be dropping `<w:br/>` elements (creating extra paragraphs instead)
   - May be incorrectly handling `<w:tab/>` elements
   
2. **Run Merging Logic:** Too aggressive
   - Merging runs with different `rsidR` values
   - Not preserving `<w:proofErr>` elements
   - Not preserving fine-grained text splits

### Secondary Issue: `atom_list.rs`

**Location:** `crates/redline-core/src/wml/atom_list.rs`

**Lines 632-633:**
```rust
// Extract atoms - may be losing information about deleted content
line 632: w:br→Break atom
line 633: w:tab→Tab atom
```

The atom extraction may be:
- Dropping deleted paragraphs (the 4 empty `<w:p>` with `w:rsidDel` attribute)
- Not preserving `<w:proofErr>` markup
- Losing revision tracking metadata

---

## Recommendations

1. **Preserve Deleted Paragraphs:**
   - Empty paragraphs with `w:rsidDel` attribute should be kept
   - These represent formatting/structure changes that MS Word tracks
   
2. **Don't Merge Runs with Different Revision IDs:**
   - Check `rsidR` attributes before coalescing
   - Preserve even single-character runs if they have different revision metadata
   
3. **Preserve All Track Changes Elements:**
   - Ensure all `<w:del>` elements are reconstructed
   - Currently losing 28 out of 62 deletions
   
4. **Keep Proof-Checking Elements:**
   - `<w:proofErr>` should pass through unchanged
   - These are part of Word's document model

---

## Next Steps

1. Review `coalesce.rs` run-merging logic (check rsidR matching)
2. Check `atom_list.rs` for deleted paragraph handling
3. Verify `<w:del>` reconstruction in output generation
4. Add test cases for:
   - Empty paragraphs with rsidDel
   - Multi-run text with different rsidR values
   - Track changes preservation

