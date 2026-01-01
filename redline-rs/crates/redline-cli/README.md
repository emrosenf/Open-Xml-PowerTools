# Redline CLI

Command-line interface for the Redline OOXML comparison engine.

## Installation

```bash
cargo install redline-cli
```

## Usage

### Comparing Documents

Compare two documents and generate a redlined output file.

```bash
redline compare original.docx modified.docx -o diff.docx
```

The tool automatically detects file types (`.docx`, `.xlsx`, `.pptx`). You can force a specific type using `-t` / `--doc-type`:

```bash
redline compare old.xlsx new.xlsx --doc-type xlsx
```

### Options

| Option | Description | Default |
|--------|-------------|---------|
| `-o, --output <PATH>` | Output document path | `redline-DATETIME-COMMIT.ext` |
| `--author <NAME>` | Author name for revisions | Modified document's author or "Redline" |
| `--json` | Output statistics as JSON (useful for scripting) | false |
| `--date <ISO8601>` | Date/time for revisions | Current time |

### Word-Specific Options

| Option | Description | Default |
|--------|-------------|---------|
| `--detail-threshold <0.0-1.0>` | Granularity of comparison. Lower = finer (word/char), Higher = coarser (paragraph). | 0.15 |
| `--trace-section <SEC>` | Trace LCS algorithm for specific section (debug) | - |
| `--trace-paragraph <TEXT>` | Trace LCS algorithm for specific paragraph (debug) | - |

## Examples

**Basic comparison:**
```bash
redline compare v1.docx v2.docx
```

**JSON output for CI/CD:**
```bash
redline compare v1.docx v2.docx --json > stats.json
```

**Custom author:**
```bash
redline compare v1.docx v2.docx --author "Compliance Bot"
```

## Supported Formats

- **Word (.docx)**: Full support. Tracks text, formatting, tables, numbering, and styles.
- **Excel (.xlsx)**: Full support. Tracks cell values, formulas, structure, and formatting. Output includes cell comments for changes.
- **PowerPoint (.pptx)**: Full support. Tracks slides, shapes, text, and geometry. Output includes visual overlays for changes.
