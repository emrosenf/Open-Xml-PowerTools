# redline-rs

Faithful Rust port of the C# OpenXML document comparers from OpenXmlPowerTools.

## Overview

This library provides byte-for-byte compatible document comparison for:
- **Word documents** (.docx) - WmlComparer
- **Excel workbooks** (.xlsx) - SmlComparer  
- **PowerPoint presentations** (.pptx) - PmlComparer

## Target Platforms

- Native Rust (desktop applications, servers)
- WebAssembly (browser and Node.js)
- Tauri desktop applications
- Python via wasmer bindings

## Project Structure

```
redline-rs/
├── crates/
│   ├── redline-core/     # Pure Rust core library
│   ├── redline-wasm/     # WASM bindings
│   ├── redline-tauri/    # Tauri plugin
│   └── redline-cli/      # Command-line tool
├── tests/
│   ├── common/           # Test utilities
│   └── golden/           # Golden test files
└── benches/              # Performance benchmarks
```

## Building

```bash
cargo build
cargo test
cargo clippy
```

## Status

This is a work-in-progress port. See RUST_MIGRATION_PLAN_SYNTHESIS.md for the detailed implementation plan.

### Phase Status

- [x] Phase 0: Baseline Capture and Project Setup
- [ ] Phase 1: Core XML and OOXML Substrate
- [ ] Phase 2: WmlComparer (Word)
- [ ] Phase 3: SmlComparer (Excel)
- [ ] Phase 4: PmlComparer (PowerPoint)
- [ ] Phase 5: Test Port and Parity Harness
- [ ] Phase 6: WASM and Tauri Integration
- [ ] Phase 7: Performance and Determinism Hardening
- [ ] Phase 8: Release Readiness

## License

MIT
