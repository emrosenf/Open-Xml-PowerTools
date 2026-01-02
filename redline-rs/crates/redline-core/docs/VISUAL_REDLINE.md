# Visual Redline Feature

Transform OOXML tracked changes into visual formatting for "Litera-style" document redlining.

## Overview

The visual redline feature converts Microsoft Word's native track changes (using OOXML revision elements like `w:ins`, `w:del`, and `w:rPrChange`) into visual formatting with colored text and formatting indicators. This produces a document that visually shows changes without requiring Word's track changes viewer.

### Visual Formatting Applied

| Change Type | Color (Default) | Formatting |
|-------------|-----------------|------------|
| Insertions | Blue (`0000FF`) | Double underline |
| Deletions | Red (`FF0000`) | Strikethrough |
| Moves | Green (`008000`) | Double underline (source) / Strikethrough (destination) |

## Usage

### CLI

```bash
# Basic visual redline
redline compare doc1.docx doc2.docx -o output.docx --visual-redline

# Custom colors (hex RGB)
redline compare doc1.docx doc2.docx -o output.docx --visual-redline \
  --insertion-color 0066CC \
  --deletion-color CC0000 \
  --move-color 009900

# Without summary table
redline compare doc1.docx doc2.docx -o output.docx --visual-redline --no-summary-table
```

### CLI Options

| Option | Default | Description |
|--------|---------|-------------|
| `--visual-redline` | `false` | Enable visual redline transformation |
| `--insertion-color` | `0000FF` | Hex RGB color for insertions |
| `--deletion-color` | `FF0000` | Hex RGB color for deletions |
| `--move-color` | `008000` | Hex RGB color for moved content |
| `--no-summary-table` | `false` | Skip the summary table |

### Rust API

```rust
use redline_core::wml::{render_visual_redline, VisualRedlineSettings, WmlDocument};

// Load a document with tracked changes
let doc = WmlDocument::from_file("compared.docx")?;

// Configure settings
let settings = VisualRedlineSettings {
    insertion_color: "0000FF".to_string(),  // Blue
    deletion_color: "FF0000".to_string(),   // Red
    move_color: "008000".to_string(),       // Green
    move_detection_min_words: 5,            // Min words to detect as move
    add_summary_table: true,
    older_filename: Some("original.docx".to_string()),
    newer_filename: Some("modified.docx".to_string()),
};

// Transform to visual redline
let result = render_visual_redline(&doc, &settings)?;

// result.document contains the transformed DOCX bytes
// result.insertions - count of insertions processed
// result.deletions - count of deletions processed
// result.moves - count of move pairs detected
// result.format_changes_removed - count of w:rPrChange elements removed

std::fs::write("visual_output.docx", &result.document)?;
```

### WASM API

```javascript
import init, {
    render_visual_redline,
    compare_word_documents_visual
} from 'redline-wasm';

await init();

// Option 1: Transform existing document with tracked changes
const result = render_visual_redline(documentBytes, JSON.stringify({
    insertion_color: "0000FF",
    deletion_color: "FF0000",
    move_color: "008000",
    move_detection_min_words: 5,
    add_summary_table: true,
    older_filename: "original.docx",
    newer_filename: "modified.docx"
}));

console.log(`Processed: ${result.insertions} insertions, ${result.deletions} deletions, ${result.moves} moves`);

// Option 2: Compare and render visual in one step
const visualResult = compare_word_documents_visual(
    olderDocBytes,
    newerDocBytes,
    JSON.stringify({ author_for_revisions: "Reviewer" }),  // Compare settings
    JSON.stringify({                                        // Visual settings
        insertion_color: "0066CC",
        deletion_color: "CC0000",
        add_summary_table: true
    })
);
```

## Settings Reference

### VisualRedlineSettings

| Field | Type | Default | Description |
|-------|------|---------|-------------|
| `insertion_color` | String | `"0000FF"` | Hex RGB color for inserted content |
| `deletion_color` | String | `"FF0000"` | Hex RGB color for deleted content |
| `move_color` | String | `"008000"` | Hex RGB color for moved content |
| `move_detection_min_words` | usize | `5` | Minimum word count to consider as a move |
| `add_summary_table` | bool | `true` | Add summary statistics table at document end |
| `older_filename` | Option<String> | `None` | Original filename for summary table |
| `newer_filename` | Option<String> | `None` | Modified filename for summary table |

## How It Works

### Transformation Process

1. **Collect Revisions**: Scan document for all `w:ins` (insertions) and `w:del` (deletions) elements, extracting their text content and word counts.

2. **Detect Moves**: Match insertions with deletions that have identical normalized text (case-insensitive, whitespace-collapsed) with at least `move_detection_min_words` words. Matched pairs are marked as "moves" and displayed in the move color.

3. **Transform Insertions**: For each `w:ins` element:
   - Apply color and double underline to all contained runs
   - Unwrap the `w:ins` element (remove wrapper, keep content)

4. **Transform Deletions**: For each `w:del` element:
   - Apply color and strikethrough to all contained runs
   - Convert `w:delText` elements to `w:t` (regular text)
   - Unwrap the `w:del` element

5. **Remove Format Changes**: Delete all `w:rPrChange` elements while preserving the current run properties.

6. **Add Summary Table**: If enabled, append a formatted table showing:
   - Report header with timestamp
   - Filenames (if provided)
   - Change statistics with visual indicators
   - Total count

### Move Detection Algorithm

The move detection algorithm identifies content that was moved from one location to another:

```
For each insertion with word_count >= min_words:
    Normalize text (lowercase, collapse whitespace)
    Look for matching deletion with same normalized text
    If found and not already paired:
        Mark both as moves
        Apply move color instead of insert/delete colors
```

This ensures that large blocks of text that were repositioned appear in green (or configured move color) rather than as separate insertions and deletions.

## Summary Table Format

When `add_summary_table` is enabled, the output document includes a formatted table:

```
┌─────────────────────────────────────────────┐
│           Summary Report                     │
├─────────────────────────────────────────────┤
│ Document comparison performed on [datetime]  │
├─────────────────────────────────────────────┤
│ Original filename: doc1.docx                 │  (if provided)
│ Modified filename: doc2.docx                 │  (if provided)
├─────────────────────────────────────────────┤
│ Changes                                      │
├─────────────────────────────────────────────┤
│ Insertions (blue, underlined)          42   │
│ Deletions (red, strikethrough)         18   │
│ Moves (green)                           3   │
├─────────────────────────────────────────────┤
│ Total                                  63   │
└─────────────────────────────────────────────┘
```

## Color Reference

Common color codes (hex RGB):

| Color | Hex Code | Preview |
|-------|----------|---------|
| Blue (default insert) | `0000FF` | Standard blue |
| Red (default delete) | `FF0000` | Standard red |
| Green (default move) | `008000` | Dark green |
| Navy | `000080` | Dark blue |
| Maroon | `800000` | Dark red |
| Purple | `800080` | Purple |
| Teal | `008080` | Teal |
| Orange | `FF6600` | Orange |
| Dark Gray | `404040` | Dark gray |

## Limitations

1. **Move Detection**: Only exact text matches are detected as moves. Minor edits within moved content will not be recognized.

2. **Complex Formatting**: Very complex nested formatting may not preserve all styling attributes.

3. **Comments**: Document comments are preserved but not modified by visual redline.

4. **Headers/Footers**: Changes in headers and footers are transformed but not included in move detection.

## Examples

### Basic Workflow

```bash
# 1. Compare two documents
redline compare original.docx modified.docx -o compared.docx

# 2. Convert to visual redline (separate step)
redline compare original.docx modified.docx -o visual.docx --visual-redline

# The --visual-redline flag combines comparison + transformation
```

### Custom Branding Colors

```bash
# Corporate blue and orange theme
redline compare doc1.docx doc2.docx -o output.docx --visual-redline \
  --insertion-color 0066B3 \
  --deletion-color E65100 \
  --move-color 2E7D32
```

### Programmatic Pipeline

```rust
use redline_core::wml::{WmlComparer, WmlComparerSettings, render_visual_redline, VisualRedlineSettings};

// Compare documents
let mut comparer = WmlComparer::new(WmlComparerSettings::default());
let result = comparer.compare_files("doc1.docx", "doc2.docx")?;

// Apply visual redline
let compared_doc = WmlDocument::from_bytes(&result.document)?;
let visual_result = render_visual_redline(&compared_doc, &VisualRedlineSettings::default())?;

// Write output
std::fs::write("visual_output.docx", &visual_result.document)?;
```
