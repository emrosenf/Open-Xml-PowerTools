# Redline WASM

WebAssembly bindings for the Redline OOXML comparison engine. Compare Word, Excel, and PowerPoint documents directly in the browser or Node.js.

## Installation

```bash
npm install redline-wasm
```

*Note: Package name may vary depending on publication status.*

## Usage

### Word Documents (.docx)

```javascript
import init, { compare_word_documents_with_changes } from 'redline-wasm';

await init();

const oldBytes = new Uint8Array(...); // Load file
const newBytes = new Uint8Array(...);

const result = compare_word_documents_with_changes(
    oldBytes, 
    newBytes, 
    JSON.stringify({ author_for_revisions: "Me" })
);

console.log(`Changes: ${result.total_revisions}`);
console.log(result.changes); // Detailed change list

// Save result.document (Uint8Array) as .docx
```

### Excel Spreadsheets (.xlsx)

```javascript
import { compare_spreadsheets } from 'redline-wasm';

const result = compare_spreadsheets(oldBytes, newBytes, null);
console.log(`${result.insertions} cells added`);
```

### PowerPoint Presentations (.pptx)

```javascript
import { compare_presentations } from 'redline-wasm';

const result = compare_presentations(oldBytes, newBytes, null);
console.log(`${result.revision_count} total changes`);
```

## API Reference

### `compare_word_documents_with_changes(older, newer, settings_json)`
Returns `CompareResultWithChanges`:
- `document`: `Uint8Array` (redlined .docx)
- `changes`: `WmlChange[]`
- `insertions`: `number`
- `deletions`: `number`
- `total_revisions`: `number`

### `compare_spreadsheets(older, newer, settings_json)`
Returns `SmlCompareResultWithChanges`:
- `document`: `Uint8Array` (marked .xlsx with comments)
- `changes`: `SmlChange[]`
- `insertions`: `number` (cells added)
- `deletions`: `number` (cells deleted)
- `revision_count`: `number`

### `compare_presentations(older, newer, settings_json)`
Returns `PmlCompareResultWithChanges`:
- `document`: `Uint8Array` (marked .pptx with overlays)
- `changes`: `PmlChange[]`
- `insertions`: `number` (slides/shapes inserted)
- `deletions`: `number` (slides/shapes deleted)
- `revision_count`: `number`

### `build_change_list(changes, options)`
Transform raw `WmlChange[]` into a UI-friendly structure.

### `build_sml_change_list(changes, options)`
Transform raw `SmlChange[]` into a UI-friendly structure (groups adjacent cells).

### `build_pml_change_list(changes, options)`
Transform raw `PmlChange[]` into a UI-friendly structure (groups by slide).

## Building Locally

```bash
wasm-pack build --target web
```

Run the demo:
1. `cd demo`
2. Serve directory (e.g., `python3 -m http.server`)
3. Open `index.html`
