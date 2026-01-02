# Redline Document Comparison Viewer

A production-grade document comparison interface built with Vue 3, showcasing the `redline-wasm` OOXML comparison engine. This demo provides a beautiful, professional UI for comparing Word, Excel, and PowerPoint documents.

## Features

### 🎯 Core Functionality
- **Real-time document comparison** using WASM bindings
- **Change categorization** (Insertions, Deletions, Replacements, Moves)
- **Interactive change list** with filtering and search
- **Document viewer** with change highlighting
- **Accept/Reject changes** with visual feedback
- **Change details panel** showing authorship and metadata

### ✨ Design Highlights
- **Editorial aesthetic**: Professional, trustworthy design
- **Dark theme** optimized for long reading sessions
- **Smooth animations** and transitions
- **Responsive layout** with adaptive sidebars
- **Keyboard navigation** (↑/↓ arrow keys, Enter/Escape)
- **Color-coded changes** for visual distinction
- **Immersive document viewing** with proper typography

### 🎮 User Experience
- **Instant feedback** on interactions
- **Smooth scrolling** to change locations
- **Status indicators** for accepted/rejected changes
- **Author and date tracking** for all changes
- **Statistics dashboard** showing change counts
- **Search and filter** changes by content or author

## Architecture

### Component Structure
```
App (Main)
├── Toolbar (Document controls)
├── ChangesList (Left sidebar)
│   ├── ChangeItem (Individual change)
│   ├── Search/Filter
│   └── Statistics
├── DocumentViewer (Center)
│   ├── DocumentHeader
│   ├── RenderedContent
│   └── PageNavigation
└── ChangeDetails (Right sidebar)
    ├── ChangeMetadata
    └── AcceptRejectButtons
```

### Technology Stack
- **Vue 3** - Reactive UI framework
- **Tailwind CSS** - Utility-first styling
- **Custom CSS** - Editorial aesthetic with animations
- **WASM Bindings** - redline-wasm for document comparison

## Integration with redline-wasm

### Real WASM Bindings Integration

The demo has been updated to use the real `redline-wasm` bindings. The integration is handled through the `redlineWasmBindings` object which:

1. **Initializes WASM** on component mount
2. **Handles comparison** for Word, Excel, and PowerPoint documents
3. **Falls back to demo data** if WASM is unavailable

### Setup Options

#### Option 1: Local Build (Development)

Copy the built WASM package to your web server:

```bash
# After building redline-wasm
cp -r crates/redline-wasm/pkg/* /path/to/webroot/
```

Then uncomment the local import in `index.html`:

```javascript
import init, {
    compare_word_documents_with_changes,
    compare_spreadsheets,
    compare_presentations
} from './redline_wasm.js';
```

#### Option 2: NPM Package

Install from npm and use dynamic import:

```bash
npm install redline-wasm
```

Update the WASM initialization:

```javascript
async init() {
    const module = await import('redline-wasm');
    await module.default();
    this.wasmModule = module;
    this.isInitialized = true;
}
```

#### Option 3: CDN (Once Published)

```javascript
import init, {
    compare_word_documents_with_changes,
    compare_spreadsheets,
    compare_presentations
} from 'https://cdn.jsdelivr.net/npm/redline-wasm@latest';

await init();
```

### API Usage

The demo uses these WASM functions:

```javascript
// Word document comparison
compare_word_documents_with_changes(
    olderDocBytes: Uint8Array,
    newerDocBytes: Uint8Array,
    settings_json?: string  // {"author_for_revisions": "Author Name"}
): object  // Returns comparison result with changes

// Visual redline (transforms tracked changes to visual formatting)
render_visual_redline(
    documentBytes: Uint8Array,  // Document with tracked changes
    settings_json?: string      // Visual redline settings
): object  // Returns document with visual formatting

// Convenience function: compare and render visual in one step
compare_word_documents_visual(
    olderDocBytes: Uint8Array,
    newerDocBytes: Uint8Array,
    compare_settings_json?: string,
    visual_settings_json?: string
): object  // Returns visually formatted document

// Spreadsheet comparison
compare_spreadsheets(
    olderSheetBytes: Uint8Array,
    newerSheetBytes: Uint8Array,
    settings_json?: string
): object

// Presentation comparison
compare_presentations(
    olderPresBytes: Uint8Array,
    newerPresBytes: Uint8Array,
    settings_json?: string
): object
```

### File Handling

The demo converts files to `Uint8Array` for WASM processing:

```javascript
async readFileAsBytes(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const buffer = e.target?.result;
            if (buffer instanceof ArrayBuffer) {
                resolve(new Uint8Array(buffer));
            } else {
                reject(new Error('Failed to read file'));
            }
        };
        reader.onerror = () => reject(new Error('File read error'));
        reader.readAsArrayBuffer(file);
    });
}
```

### Result Parsing

The WASM returns comparison results that the demo parses into a normalized format:

```javascript
{
    insertions: number,
    deletions: number,
    replacements: number,
    moves: number,
    changes: [
        {
            id: number,
            type: 'insert' | 'delete' | 'replace' | 'move',
            content: string,
            author: string,
            date: string,
            location: string,
            count: number
        }
    ]
}
```

### Complete Example

```javascript
const { createApp } = Vue;

const app = createApp({
    data() {
        return {
            wasmReady: false,
            comparisonResult: null,
            changes: [],
            documentType: 'word' // 'word', 'spreadsheet', 'presentation'
        };
    },

    async mounted() {
        // Initialize WASM
        const wasmModule = await import('redline-wasm');
        await wasmModule.default();
        this.wasmReady = true;
    },

    async methods: {
        async compareDocuments(file1Bytes, file2Bytes) {
            const wasmModule = await import('redline-wasm');

            if (this.documentType === 'word') {
                const result = wasmModule.compare_word_documents_with_changes(
                    file1Bytes,
                    file2Bytes,
                    JSON.stringify({ author_for_revisions: "Me" })
                );

                this.changes = result.changes;
                this.comparisonResult = {
                    insertions: result.insertions,
                    deletions: result.deletions,
                    total: result.total_revisions,
                    document: result.document
                };
            } else if (this.documentType === 'spreadsheet') {
                const result = wasmModule.compare_spreadsheets(
                    file1Bytes,
                    file2Bytes
                );

                this.changes = result.changes;
                this.comparisonResult = {
                    insertions: result.insertions,
                    deletions: result.deletions,
                    total: result.revision_count,
                    document: result.document
                };
            } else if (this.documentType === 'presentation') {
                const result = wasmModule.compare_presentations(
                    file1Bytes,
                    file2Bytes
                );

                this.changes = result.changes;
                this.comparisonResult = {
                    insertions: result.insertions,
                    deletions: result.deletions,
                    total: result.revision_count,
                    document: result.document
                };
            }
        },

        async acceptChange(changeId) {
            // In real app, this would use accept_revisions_by_id for Word docs
            const change = this.changes.find(c => c.id === changeId);
            if (change) change.status = 'accepted';
        },

        async rejectChange(changeId) {
            // In real app, this would use reject_revisions_by_id for Word docs
            const change = this.changes.find(c => c.id === changeId);
            if (change) change.status = 'rejected';
        }
    }
});
```

## Styling & Customization

### CSS Variables
The design uses CSS custom properties for easy theming:

```css
:root {
    --primary-dark: #0f172a;
    --primary-darker: #050f1f;
    --accent-insert: #10b981;
    --accent-delete: #ef4444;
    --accent-replace: #f59e0b;
    --accent-move: #6366f1;
    --surface-1: #1e293b;
    --surface-2: #334155;
    --text-primary: #f1f5f9;
    --text-secondary: #cbd5e1;
    --border-color: #475569;
}
```

### Customizing Colors
```css
:root {
    --accent-insert: #06b6d4;  /* Cyan */
    --accent-delete: #f87171;  /* Pink */
    --accent-replace: #a78bfa; /* Purple */
}
```

### Adding Light Mode
```css
@media (prefers-color-scheme: light) {
    :root {
        --primary-dark: #ffffff;
        --text-primary: #1e293b;
        --text-secondary: #475569;
        /* ... etc */
    }
}
```

## Animation Details

The UI includes several animation patterns:

### Entrance Animations
- `slideInLeft` - Left sidebar slides in from left
- `slideInRight` - Right sidebar slides in from right
- `fadeIn` - Content fades in smoothly

### Interactive Animations
- Hover effects on change items with shimmer
- Button elevation and glow on hover
- Smooth color transitions
- Pulsing glow on active changes

### Custom Animation Classes
```css
.animate-slide-in-left { animation: slideInLeft 0.3s ease-out; }
.animate-slide-in-right { animation: slideInRight 0.3s ease-out; }
.animate-fade-in { animation: fadeIn 0.3s ease-out; }
.active-change { animation: pulse-glow 2s infinite; }
```

## Keyboard Shortcuts

| Key | Action |
|-----|--------|
| ↑ Arrow Up | Previous change |
| ↓ Arrow Down | Next change |
| Enter | Accept current change |
| Escape | Reject current change |

## File Structure

```
redline-demo/
├── index.html          # Complete self-contained demo
├── README.md          # This file
├── INTEGRATION.md     # Detailed WASM integration guide
└── STYLING.md         # Design system documentation
```

## Quick Start

### 1. Build WASM (Required for Real Document Comparison)

```bash
# From the redline-wasm directory
./build.sh
```

This creates:
- `pkg/redline_wasm_bg.wasm` - Uncompressed WASM (1.5 MB)
- `pkg/redline_wasm_bg.wasm.br` - Brotli compressed (430 KB)
- `pkg/redline_wasm.js` - JavaScript bindings
- And other supporting files

### 2. Run the Demo

```bash
# From the demo directory
cd demo

# Option A: Using Python (Recommended)
python3 -m http.server 8000

# Option B: Using Node http-server
npx http-server .

# Visit http://localhost:8000
```

### 3. Demo Features

**Demo Mode (without real files):**
- Click on the interface to load sample comparison data
- All features work with mock data (insertions, deletions, replacements)
- No files needed to test the UI

**Real Document Comparison:**
- Open the file upload modal (when ready)
- Select document type: Word (.docx), Excel (.xlsx), or PowerPoint (.pptx)
- Upload two documents to compare
- WASM engine performs the comparison
- Results displayed with full change tracking

### 4. Integrate into Existing Project

Copy the Vue 3 application code into your project:

```html
<div id="app"></div>
<script src="https://cdn.jsdelivr.net/npm/vue@3/dist/vue.global.js"></script>
<script src="index.html" type="module"></script>
<!-- And include redline-wasm WASM files -->
<script type="module">
  import init, { compare_word_documents_with_changes } from './redline_wasm.js';
  // Initialize and use WASM
</script>
```

## Performance Optimization

### WASM Loading & Compression

The redline-wasm module is built with Brotli compression, reducing download size by 72.8%:

- **Uncompressed**: 1.5 MB
- **Brotli**: 430 KB (72.8% savings)
- **Gzip**: 589 KB (62.7% savings)

For deployment, see the [DEPLOYMENT.md](../DEPLOYMENT.md) guide for server configuration examples (Nginx, Apache, Vercel, Cloudflare, AWS CloudFront, Fastly, jsDelivr).

```javascript
// Lazy-load WASM only when needed
let wasmModule = null;

async function ensureWasmLoaded() {
    if (!wasmModule) {
        wasmModule = await import('redline-wasm');
        await wasmModule.default();
    }
    return wasmModule;
}
```

### Document Rendering
- Virtual scrolling for large documents
- Lazy-load change highlighting
- Debounce search queries
- Cache comparison results

### Memory Management
```javascript
// Clear large objects when not needed
this.comparisonResult = null;
this.documentBytes = null;
this.file1Bytes = null;
this.file2Bytes = null;
```

## Browser Support

| Browser | Support |
|---------|---------|
| Chrome | ✅ Full support |
| Firefox | ✅ Full support |
| Safari | ✅ Full support |
| Edge | ✅ Full support |
| IE 11 | ❌ Not supported (WASM required) |

## Accessibility

The interface includes:
- Keyboard navigation support
- Semantic HTML structure
- ARIA labels for screen readers
- High contrast color scheme
- Focus indicators on all interactive elements

## Known Limitations

1. **Document Rendering**: Currently uses text rendering. Full document rendering would require:
   - PDF.js for visual document display
   - Custom Word/Excel/PowerPoint parsers
   - Layout engine for proper formatting

2. **Large Files**: Performance optimizations needed for:
   - Documents with 1000+ changes
   - File sizes > 100MB
   - Real-time updates

3. **Advanced Features**: Not yet implemented:
   - Change comments/notes
   - Document properties comparison
   - Track changes history
   - Collaborative review mode

## Future Enhancements

- [ ] PDF rendering with change overlays
- [ ] Full-fidelity Word/Excel/PowerPoint rendering
- [ ] Batch comparison of multiple documents
- [ ] Change comments and discussion threads
- [ ] Real-time collaboration
- [ ] Document diff metrics and analytics
- [ ] Integration with document management systems
- [ ] Mobile-responsive design
- [ ] Dark/Light mode toggle
- [ ] Export comparison reports

## Contributing

This demo is designed to showcase redline-wasm capabilities. To extend it:

1. Modify the Vue 3 app in `index.html`
2. Update styling in the `<style>` section
3. Integrate real WASM bindings (see INTEGRATION.md)
4. Test with actual documents

## License

MIT License - Use freely in your projects

## Support

For issues or questions about:
- **redline-wasm**: See the main redline-rs repository
- **This demo**: Refer to INTEGRATION.md for technical details
- **Design patterns**: Check STYLING.md for customization guide
