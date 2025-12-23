# OpenXmlPowerTools WASI Experiment

This directory contains an experimental WASI (WebAssembly System Interface) build of document comparison functionality.

## Status: Experimental / Incomplete

This was an experiment to run document comparison in WASM without bundling the full .NET runtime. The experiment successfully proved that:

1. **WASI builds work** - Using componentize-dotnet (NativeAOT-LLVM), we can compile .NET to WASM
2. **Document I/O works** - SharpCompress-based ZIP handling works in WASI (System.IO.Compression does not)
3. **Basic comparison works** - Simple paragraph-level comparison produces valid DOCX output

However, this approach was **abandoned** because:

- The simplified `WasiComparer` doesn't use any of the 8,800 lines of edge-case handling from `WmlComparer`
- Refactoring `WmlComparer` to use an abstraction layer over `WordprocessingDocument` would be substantial work
- A TypeScript port may be more practical for desktop/browser use cases

## What Was Built

### OpenXmlPowerTools.Packaging/

SharpCompress-based OPC (Open Packaging Convention) implementation:

- `SharpCompressPackage.cs` - ZIP-based package handling (replaces System.IO.Packaging)
- `SharpCompressPart.cs` - Package part implementation
- `SharpCompressRelationshipCollection.cs` - Relationship management
- `WasiDocument.cs` - `WasiWordDocument` and `WasiPart` adapters for document access
- `WasiComparer.cs` - Simplified paragraph-level document comparer
- `IDocumentPart.cs` - Abstraction interfaces (`IDocumentPart`, `IWordDocument`)

### OpenXmlPowerTools.Wasi/

CLI application for testing WASI functionality:

- `Program.cs` - Commands: `info`, `extract-text`, `list-parts`, `compare`, `test`

### Dockerfile.componentize

Docker build file for componentize-dotnet (NativeAOT-LLVM to WASM).

## Build Instructions

```bash
# Build WASM module (requires Docker)
docker build --platform linux/amd64 -t openxml-componentize -f Dockerfile.componentize .

# Extract WASM file
docker run --rm -v $(pwd)/output:/out openxml-componentize cp /output/native/OpenXmlPowerTools.Wasi.wasm /out/

# Run with wasmtime
wasmtime run -S cli -S http --dir . OpenXmlPowerTools.Wasi.wasm test
wasmtime run -S cli -S http --dir . OpenXmlPowerTools.Wasi.wasm compare original.docx modified.docx result.docx
```

## Bundle Size

| Format | Size |
|--------|------|
| Uncompressed WASM | 16 MB |
| Brotli compressed | 3.4 MB |

## Key Learnings

1. **Mono WASI is broken** - .NET 10's `wasi-experimental` workload crashes during GC initialization ([dotnet/runtime#117848](https://github.com/dotnet/runtime/issues/117848))

2. **componentize-dotnet works** - NativeAOT-LLVM compilation to WASM works correctly with the `BytecodeAlliance.Componentize.DotNet.Wasm.SDK` package

3. **System.IO.Compression doesn't work** - Requires native zlib. Use SharpCompress instead.

4. **System.Security.Cryptography doesn't work** - SHA256 not available. Use alternative hash algorithms (FNV-1a, etc.)

5. **OpenXML SDK is too tightly coupled** - `WmlComparer` extensively uses SDK-specific types (`MainDocumentPart`, `StyleDefinitionsPart`, etc.) making abstraction difficult

## Conclusion

For a desktop app needing document comparison without bundling .NET:

- **Recommended**: Port WmlComparer logic to TypeScript (runs natively in browser/Node.js)
- **Not recommended**: Continue with WASI approach (bleeding-edge tooling, abstraction complexity)

The value of WmlComparer is in its algorithms and edge-case handling, not in C#-specific features. Those algorithms can be ported to TypeScript with the same test documents validating correctness.
