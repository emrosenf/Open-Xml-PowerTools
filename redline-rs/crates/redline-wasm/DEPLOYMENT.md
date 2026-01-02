# Redline WASM Deployment Guide

## Compression & Size Optimization

The redline-wasm package includes pre-compressed variants to minimize download size:

| Format | Size | Compression |
|--------|------|-------------|
| Uncompressed (.wasm) | 1.5 MB | Baseline |
| Gzip (.wasm.gz) | 380 KB | 74.7% savings |
| Brotli (.wasm.br) | 430 KB | 72.8% savings |

### Recommended Approach: HTTP Content-Encoding

The most efficient method is using **HTTP Content-Encoding**, where the web server/CDN automatically compresses the WASM with Brotli and the browser automatically decompresses it.

**Benefits:**
- Single WASM file to serve
- Automatic compression/decompression
- Works with all modern browsers
- CDN-optimized

#### Nginx Configuration

```nginx
server {
    # Enable Brotli compression (requires brotli module)
    brotli on;
    brotli_types application/wasm application/javascript text/css text/xml;
    
    # Or use gzip if brotli unavailable
    gzip on;
    gzip_types application/wasm application/javascript text/css text/xml;
    gzip_comp_level 9;
    
    location /wasm/ {
        alias /path/to/redline-wasm/pkg/;
        add_header Cache-Control "public, max-age=31536000, immutable";
    }
}
```

#### Apache Configuration

```apache
<FilesMatch "\.wasm$">
    # Enable brotli if available
    <IfModule mod_brotli.c>
        SetEnvIfNoCase Request_URI "\.wasm$" no-brotli
    </IfModule>
    
    # Or use deflate
    <IfModule mod_deflate.c>
        AddOutputFilter DEFLATE
    </IfModule>
    
    Header set Cache-Control "public, max-age=31536000, immutable"
</FilesMatch>
```

#### Vercel Configuration

```json
{
  "builds": [
    {
      "src": "public/**",
      "use": "@vercel/static"
    }
  ],
  "headers": [
    {
      "source": "/wasm/(.*)",
      "headers": [
        {
          "key": "Cache-Control",
          "value": "public, max-age=31536000, immutable"
        }
      ]
    }
  ]
}
```

#### Cloudflare Workers

```javascript
addEventListener('fetch', event => {
  event.respondWith(handleRequest(event.request));
});

async function handleRequest(request) {
  let response = await fetch(request);
  
  // Cloudflare automatically handles brotli compression
  // Just ensure correct MIME types
  if (request.url.endsWith('.wasm')) {
    response = new Response(response.body, {
      ...response,
      headers: new Headers(response.headers)
    });
    response.headers.set('Content-Type', 'application/wasm');
    response.headers.set('Cache-Control', 'public, max-age=31536000, immutable');
  }
  
  return response;
}
```

### Alternative Approach: Pre-Compressed Variants

If you need more control over which variant is served (e.g., for older browser support), use pre-compressed files:

#### Webpack Configuration

```javascript
export default {
  // ... other config
  experiments: {
    asyncWebAssembly: true,
    topLevelAwait: true,
  },
  plugins: [
    new WasmPackPlugin({
      crateDirectory: path.resolve(__dirname, 'crates/redline-wasm'),
      // Webpack will automatically optimize WASM
    }),
  ],
};
```

#### Custom Loader

```javascript
async function loadWasm() {
  // Detect what the browser supports
  const supportsBrotli = await checkBrotliSupport();
  
  let wasmPath = 'redline_wasm_bg.wasm';
  let contentEncoding = null;
  
  if (supportsBrotli) {
    wasmPath = 'redline_wasm_bg.wasm.br';
    contentEncoding = 'br';
  } else if (await checkGzipSupport()) {
    wasmPath = 'redline_wasm_bg.wasm.gz';
    contentEncoding = 'gzip';
  }
  
  const response = await fetch(wasmPath);
  
  let buffer = await response.arrayBuffer();
  if (contentEncoding === 'br') {
    buffer = await decompressBrotli(buffer);
  } else if (contentEncoding === 'gzip') {
    buffer = await decompressGzip(buffer);
  }
  
  return WebAssembly.instantiate(buffer);
}
```

## CDN Recommendations

### AWS CloudFront

1. Create S3 bucket with WASM files
2. Configure CloudFront distribution:
   - Enable "Compress Objects Automatically"
   - Set Cache-Control headers in S3
   - Enable GZip/Brotli compression in viewer settings

### Fastly

```vcl
sub vcl_deliver {
  if (req.url ~ "\.wasm$") {
    set resp.http.Cache-Control = "public, max-age=31536000, immutable";
  }
}
```

### jsDelivr

Automatically serves with optimal compression:
```html
<script type="module">
  import init from 'https://cdn.jsdelivr.net/npm/redline-wasm@latest/redline_wasm.js';
  await init();
</script>
```

## Performance Optimization Tips

1. **Use immutable cache headers** - WASM rarely changes, cache forever with version in URL
2. **Enable HTTP/2 Server Push** - Push WASM alongside HTML
3. **Consider code splitting** - Load WASM on-demand if not needed immediately
4. **Monitor Core Web Vitals** - Track impact on LCP/FID

```javascript
// Load WASM on-demand
async function loadRedlineWasm() {
  const module = await import('redline-wasm');
  await module.default();
  return module;
}

// Only load when needed
document.getElementById('compare-btn')?.addEventListener('click', 
  async () => {
    const redline = await loadRedlineWasm();
    // Use redline API
  }
);
```

## Verification

After deployment, verify compression is working:

```bash
# Check Brotli compression
curl -I -H "Accept-Encoding: br" https://your-domain/wasm/redline_wasm_bg.wasm
# Should show: Content-Encoding: br

# Check file size on wire
curl -w "Downloaded %{size_download} bytes\n" -o /dev/null -s \
  -H "Accept-Encoding: br" \
  https://your-domain/wasm/redline_wasm_bg.wasm
```

## Build Instructions

To rebuild WASM with compression:

```bash
cd crates/redline-wasm
./build.sh
```

This will:
1. Build WASM with wasm-pack
2. Create brotli compression (.br)
3. Create gzip compression (.gz)
4. Display compression results

## Summary

| Deployment Method | File Size | Implementation |
|-------------------|-----------|-----------------|
| HTTP Content-Encoding (Brotli) | 430 KB | ⭐⭐⭐ Recommended |
| HTTP Content-Encoding (Gzip) | 380 KB | ⭐⭐⭐ Recommended |
| Pre-compressed variants | 430-1500 KB | ⭐⭐ Good |
| Uncompressed | 1.5 MB | ⭐ Not recommended |

Choose HTTP Content-Encoding for the best balance of simplicity and performance.
