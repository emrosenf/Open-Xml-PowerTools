#!/bin/bash
set -e

# Check if wasm-pack is installed
if ! command -v wasm-pack &> /dev/null; then
    echo "Error: wasm-pack is not installed."
    echo "Please install it using: cargo install wasm-pack"
    exit 1
fi

echo "Building WASM module..."
wasm-pack build --target web

echo "Build complete."
echo ""
echo "To run the demo:"
echo "1. python3 -m http.server"
echo "2. Open http://localhost:8000/demo/"
