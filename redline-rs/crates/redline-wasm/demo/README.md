# Redline WASM Demo

This demo allows you to compare documents entirely in the browser using the compiled WebAssembly module.

## Quick Start

1. **Install Prerequisites**:
   Ensure you have Rust and `wasm-pack` installed:
   ```bash
   cargo install wasm-pack
   ```

2. **Run Setup**:
   Navigate to `crates/redline-wasm` and run:
   ```bash
   ./setup_demo.sh
   ```
   
   Or manually:
   ```bash
   wasm-pack build --target web
   python3 -m http.server
   ```

3. **Open Browser**:
   Go to [http://localhost:8000/demo/](http://localhost:8000/demo/)

   **Important**: You MUST run the server from the `redline-wasm` directory, NOT the `demo` directory, so that it can serve the `pkg/` files (which are siblings to `demo/`).

## Troubleshooting

- **404 for redline_wasm.js**: 
  - Did you run `wasm-pack build`? The `pkg/` directory must exist.
  - Are you running the server from `crates/redline-wasm`? If you run it from `demo/`, it cannot access the `../pkg` folder.
  
- **Drag and drop opens the file**: 
  - The JavaScript failed to load or initialize. Check the console.
  - Ensure you are serving via HTTP (`http://localhost...`), not `file://`.
