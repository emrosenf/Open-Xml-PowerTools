# Continuation Prompt for Next Session

## Context

You are working on `redline-rs`, a Rust port of Open-Xml-PowerTools WmlComparer for comparing Word documents. The project is located at:
```
/Users/evan/development/openxml-worktree/rust-port-phase0/redline-rs
```

## Critical Bug: MS Word Cannot Open Output

**Priority: P0 - Blocking**

MS Word refuses to open the comparison output file. This is a **pre-existing bug** (exists in HEAD before any changes from this session).

### Test Command
```bash
cd /Users/evan/development/openxml-worktree/rust-port-phase0/redline-rs
cargo build --release
./target/release/redline compare \
  --source1 "/Users/evan/Dropbox/Real Estate Deals/Acquired/115 Industrial Ct/Lease Up/Republic/Lease/Version 20699_2 Lease (11-19-25) - LL 1.11.docx" \
  --source2 "/Users/evan/Dropbox/Real Estate Deals/Acquired/115 Industrial Ct/Lease Up/Republic/Lease/Lease (BFI Rev 12-11-25).docx" \
  --output /tmp/comparison-result.docx
```

### Gold Standard (works in Word)
```
/Users/evan/Dropbox/Real Estate Deals/Acquired/115 Industrial Ct/Lease Up/Republic/Lease/msword-comparison-2025-12-28.docx
```

### What We Verified Is NOT The Problem
1. All XML files pass Python XML parser validation
2. rId references in document.xml match the rels file
3. Content_Types.xml structure is correct
4. ZIP package structure is valid
5. File sizes are similar to gold standard (~1.7MB)

### Suspected Issues (Need Investigation)
1. **Dangling comment references**: comments.xml has 10 comments (IDs: 24,42,43,72,98,129,130,149,155,157) but document.xml has ZERO:
   - No `commentRangeStart` elements
   - No `commentRangeEnd` elements  
   - No `commentReference` elements
   
2. **paraId mismatches**: commentsExtended.xml references paraIds like `0DC80402` that don't appear in document.xml's paragraph paraIds

3. **Possible revision markup issues**: w:ins/w:del structure might be malformed

### Next Steps
1. Ask user what error Word shows (repair dialog? specific message?)
2. Try creating a test docx WITHOUT comments - does it open?
3. Compare document.xml structure against gold standard in detail
4. Use Open XML SDK Productivity Tool or similar OOXML validator
5. Check C# reference implementation for how comments are handled:
   `/Users/evan/development/openxml-worktree/rust-port-phase0/OpenXmlPowerTools/WmlComparer.cs`

### Uncommitted Changes From This Session
The following files have uncommitted changes that ADD functionality (comment handling improvements):
- `crates/redline-core/src/wml/atom_list.rs` - Added CommentRangeStart/End handling
- `crates/redline-core/src/wml/coalesce.rs` - Added comment range element creation
- `crates/redline-core/src/wml/comments.rs` - Added annotationRef and styles
- `crates/redline-core/src/wml/comparison_unit.rs` - Added CommentRangeStart/End enum variants
- `crates/redline-core/src/xml/namespaces.rs` - Added p_style helper

These changes improve comment handling but do NOT fix the Word opening issue (bug exists in HEAD too).

### Hive Issue Tracking
Run `bd ready` to see open issues. Key issue: `cell-gpsl8v-mjrku28ggvf`

### Quick Diagnostic Commands
```bash
# Check comment markers in output vs gold
unzip -p /tmp/comparison-result.docx word/document.xml | grep -c "commentRangeStart"
unzip -p "/Users/evan/Dropbox/Real Estate Deals/Acquired/115 Industrial Ct/Lease Up/Republic/Lease/msword-comparison-2025-12-28.docx" word/document.xml | grep -c "commentRangeStart"

# Compare file sizes
ls -la /tmp/comparison-result.docx
ls -la "/Users/evan/Dropbox/Real Estate Deals/Acquired/115 Industrial Ct/Lease Up/Republic/Lease/msword-comparison-2025-12-28.docx"

# Validate XML
python3 -c "import xml.etree.ElementTree as ET; ET.parse('/tmp/doc.xml'); print('Valid')"
```

## Start Here
1. First, ask the user: "What error does Word show when you try to open the file? Does it offer to repair?"
2. Based on the answer, either:
   - If repair works: Check what Word changed during repair
   - If no repair option: Deep dive into OOXML structure comparison with gold standard
