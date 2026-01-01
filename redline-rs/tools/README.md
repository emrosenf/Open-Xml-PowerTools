# DOCX Debugging Tools

## trace_lcs.py

Trace the LCS (Longest Common Subsequence) algorithm step-by-step for debugging redline comparison.

### Usage

```bash
# Compare same section from two documents
./tools/trace_lcs.py doc1.docx doc2.docx "section 3.1"

# Compare with raw text (no files needed)
./tools/trace_lcs.py --text "the quick brown fox" "the slow brown dog"

# Show full LCS matrix (for small inputs)
./tools/trace_lcs.py --matrix --text "hello world" "hello there world"

# Character-level diff instead of word-level
./tools/trace_lcs.py --chars --text "hello" "hallo"

# Show backtrack trace through the matrix
./tools/trace_lcs.py --trace --text "old text here" "new text here"

# Show raw edits without coalescing
./tools/trace_lcs.py --no-coalesce --text "a b c" "a x c"
```

### Output

The tool shows:
- **LCS Matrix**: The dynamic programming table (with `--matrix`)
- **Backtrack Trace**: Step-by-step path through the matrix (with `--trace`)
- **Edit Script**: The sequence of EQUAL/DELETE/INSERT operations
- **Coalescing Analysis**: How consecutive operations group into `<w:ins>`/`<w:del>` blocks
- **Side-by-Side Alignment**: Visual comparison of the two sequences

### Example

```
$ ./tools/trace_lcs.py --text "Months 1-12 $TBD" "Months 1-4 $0.00"

=== COALESCED EDIT SCRIPT ===
  = Months
  - 1-12
  + 1-4
  =
  - $TBD
  + $0.00

Revision groups that would be created:
  <w:del> #1: 1-12
  <w:ins> #2: 1-4
  <w:del> #4: $TBD
  <w:ins> #5: $0.00
```

---

## extract_section.py

Extract and compare sections from DOCX files for debugging redline output.

### Requirements

Python 3.8+ (uses only standard library)

### Usage

```bash
# Extract a numbered section (continues until next section at same/higher level)
./tools/extract_section.py doc.docx "section 3.1"
./tools/extract_section.py doc.docx "section (b)"

# Extract paragraph(s) starting with specific text
./tools/extract_section.py doc.docx "para 'Rent Commencement'"
./tools/extract_section.py doc.docx "paragraph 'The quick brown fox'"

# Extract a footnote or endnote by ID
./tools/extract_section.py doc.docx "footnote 5"
./tools/extract_section.py doc.docx "endnote 3"

# Extract the most specific element containing text
./tools/extract_section.py doc.docx "element 'unique phrase'"

# Compare two files side-by-side
./tools/extract_section.py gold.docx rust-output.docx "section 4.1"

# Compare two files with unified diff
./tools/extract_section.py --diff gold.docx rust-output.docx "section 4.1"

# Write output to file
./tools/extract_section.py doc.docx "section 3.1" -o section3.1.xml
```

### Query Types

| Query | Description |
|-------|-------------|
| `section X` | Find section by number (e.g., `3.1`, `(a)`, `(ii)`, `a.`) and include content until next section at same or higher level |
| `para 'text'` | Find paragraph starting with text, continue until next numbered section |
| `paragraph 'text'` | Same as `para` |
| `footnote N` | Extract footnote by ID from footnotes.xml |
| `endnote N` | Extract endnote by ID from endnotes.xml |
| `element 'text'` | Find smallest element containing text (single element, no continuation) |

### Section Number Formats

The tool recognizes these section numbering styles:

- Numeric: `1`, `1.1`, `1.1.1`, `2.`
- Parenthetical alpha: `(a)`, `(b)`, `(A)`, `(B)`
- Parenthetical roman: `(i)`, `(ii)`, `(iii)`
- Dotted alpha: `a.`, `b.`, `A.`
- Dotted roman: `i.`, `ii.`, `I.`

### Output

Pretty-printed XML suitable for diffing. Extracts from `word/document.xml` (or `word/footnotes.xml` / `word/endnotes.xml` for notes).

### Examples

Compare rent section between MS Word redline and Rust redline:
```bash
./tools/extract_section.py --diff \
  /path/to/msword-comparison.docx \
  /tmp/rust-redline-output.docx \
  "section 4.1"
```

Extract a specific clause for inspection:
```bash
./tools/extract_section.py lease.docx "section (b)" | less
```

Find where a specific phrase appears:
```bash
./tools/extract_section.py contract.docx "element 'force majeure'"
```
