#!/bin/bash

# Build script for redline-wasm with brotli compression

set -e

echo "Building redline-wasm..."
wasm-pack build --release --target web

# Compress with brotli
echo "Compressing WASM with brotli..."
brotli -k -q 11 pkg/redline_wasm_bg.wasm

# Also create gzip for comparison
if command -v gzip &> /dev/null; then
    gzip -k -9 pkg/redline_wasm_bg.wasm
fi

# Print compression stats
echo ""
echo "=== Compression Results ==="
echo "Original: $(ls -lh pkg/redline_wasm_bg.wasm | awk '{print $5}')"
if [ -f "pkg/redline_wasm_bg.wasm.br" ]; then
    echo "Brotli:   $(ls -lh pkg/redline_wasm_bg.wasm.br | awk '{print $5}')"
fi
if [ -f "pkg/redline_wasm_bg.wasm.gz" ]; then
    echo "Gzip:     $(ls -lh pkg/redline_wasm_bg.wasm.gz | awk '{print $5}')"
fi

echo ""
echo "✓ Build complete!"
