# Rust Determinism Best Practices for XML Tree Traversal and Hashing

**Context**: OpenXML document comparison requires deterministic behavior for:
- Reproducible hashing (SHA1/SHA256) across runs
- Consistent test results and golden file validation
- Cross-platform compatibility (Windows, macOS, Linux, WASM)

---

## 1. HashMap/HashSet Iteration Nondeterminism

### The Problem

**`HashMap` and `HashSet` use randomized hashing by default** to prevent DoS attacks. Iteration order is **nondeterministic** across runs:

```rust
use std::collections::HashMap;

let mut map = HashMap::new();
map.insert("a", 1);
map.insert("b", 2);
map.insert("c", 3);

// ❌ NONDETERMINISTIC: Order changes between runs
for (k, v) in &map {
    println!("{}: {}", k, v);
}
// Run 1: b:2, a:1, c:3
// Run 2: c:3, b:2, a:1
```

### Solutions

#### Option 1: BTreeMap (Sorted Order)

**Use when**: You need sorted keys or deterministic iteration.

```rust
use std::collections::BTreeMap;

let mut map = BTreeMap::new();
map.insert("c", 3);
map.insert("a", 1);
map.insert("b", 2);

// ✅ DETERMINISTIC: Always sorted by key
for (k, v) in &map {
    println!("{}: {}", k, v);
}
// Always: a:1, b:2, c:3
```

**Tradeoffs**:
- ✅ Deterministic iteration (sorted)
- ✅ No extra dependencies
- ❌ Slower than HashMap (O(log n) vs O(1))
- ❌ Requires `Ord` trait on keys

#### Option 2: IndexMap (Insertion Order)

**Use when**: You need insertion order preservation (like C# `Dictionary` or Python `dict` 3.7+).

```rust
use indexmap::IndexMap;

let mut map = IndexMap::new();
map.insert("c", 3);
map.insert("a", 1);
map.insert("b", 2);

// ✅ DETERMINISTIC: Preserves insertion order
for (k, v) in &map {
    println!("{}: {}", k, v);
}
// Always: c:3, a:1, b:2
```

**Tradeoffs**:
- ✅ Deterministic iteration (insertion order)
- ✅ Fast lookups (O(1) like HashMap)
- ✅ Familiar to C#/Python developers
- ❌ Requires `indexmap` crate
- ❌ Slightly more memory overhead

**Add to Cargo.toml**:
```toml
[dependencies]
indexmap = "2.0"
```

#### Option 3: Collect and Sort Keys

**Use when**: You rarely iterate and want to keep HashMap performance.

```rust
use std::collections::HashMap;

let mut map = HashMap::new();
map.insert("c", 3);
map.insert("a", 1);
map.insert("b", 2);

// ✅ DETERMINISTIC: Sort keys before iteration
let mut keys: Vec<_> = map.keys().collect();
keys.sort();
for k in keys {
    println!("{}: {}", k, map[k]);
}
// Always: a:1, b:2, c:3
```

**Tradeoffs**:
- ✅ No extra dependencies
- ✅ HashMap performance for lookups
- ❌ Allocation overhead on every iteration
- ❌ Verbose

---

## 2. Controlling HashMap Random State

### The Problem

Even with the same insertion order, `HashMap` uses a **random seed** per process:

```rust
use std::collections::HashMap;

// ❌ Different hash values across runs
let mut map1 = HashMap::new();
map1.insert("key", "value");
// Run 1: internal hash = 0x1234abcd
// Run 2: internal hash = 0x9876fedc
```

### Solution: Fixed Hasher (Testing Only)

**Use when**: You need reproducible hashes for testing (NOT production).

```rust
use std::collections::HashMap;
use std::hash::BuildHasherDefault;
use std::collections::hash_map::DefaultHasher;

// ✅ Fixed hasher (deterministic within same build)
type DeterministicHashMap<K, V> = HashMap<K, V, BuildHasherDefault<DefaultHasher>>;

let mut map = DeterministicHashMap::default();
map.insert("key", "value");
```

**⚠️ WARNING**: `DefaultHasher` is **not stable across Rust versions**. Use only for tests, never for persistent storage.

### Solution: ahash (Fast + Deterministic Option)

**Use when**: You want fast hashing with optional determinism.

```rust
use ahash::AHashMap;

// ❌ NONDETERMINISTIC (default, secure)
let mut map1 = AHashMap::new();

// ✅ DETERMINISTIC (with fixed seed)
use ahash::RandomState;
let hasher = RandomState::with_seeds(1, 2, 3, 4);
let mut map2 = AHashMap::with_hasher(hasher);
```

**Add to Cargo.toml**:
```toml
[dependencies]
ahash = "0.8"
```

**Tradeoffs**:
- ✅ Faster than `std::collections::HashMap`
- ✅ Optional determinism via fixed seeds
- ❌ Requires extra dependency
- ❌ Fixed seeds reduce DoS protection

---

## 3. Iterator Ordering Guarantees

### Deterministic Iterators

| Collection | Iteration Order | Deterministic? |
|------------|----------------|----------------|
| `Vec<T>` | Index order | ✅ Yes |
| `BTreeMap<K, V>` | Sorted by key | ✅ Yes |
| `BTreeSet<T>` | Sorted | ✅ Yes |
| `IndexMap<K, V>` | Insertion order | ✅ Yes |
| `IndexSet<T>` | Insertion order | ✅ Yes |
| `HashMap<K, V>` | **Random** | ❌ No |
| `HashSet<T>` | **Random** | ❌ No |

### XML Attribute Ordering

**Problem**: XML attributes are unordered by spec, but hashing requires consistent serialization.

```rust
// ❌ NONDETERMINISTIC: HashMap iteration
use std::collections::HashMap;

fn serialize_attributes(attrs: &HashMap<String, String>) -> String {
    let mut result = String::new();
    for (k, v) in attrs {
        result.push_str(&format!(" {}=\"{}\"", k, v));
    }
    result
}
// Run 1: " id=\"1\" class=\"foo\""
// Run 2: " class=\"foo\" id=\"1\""
```

**Solution**: Sort attributes before serialization.

```rust
// ✅ DETERMINISTIC: Sort keys
use std::collections::HashMap;

fn serialize_attributes(attrs: &HashMap<String, String>) -> String {
    let mut result = String::new();
    let mut keys: Vec<_> = attrs.keys().collect();
    keys.sort(); // Alphabetical order
    for k in keys {
        result.push_str(&format!(" {}=\"{}\"", k, attrs[k]));
    }
    result
}
// Always: " class=\"foo\" id=\"1\""
```

**Better**: Use `BTreeMap` for attributes from the start.

```rust
// ✅ DETERMINISTIC: BTreeMap always sorted
use std::collections::BTreeMap;

fn serialize_attributes(attrs: &BTreeMap<String, String>) -> String {
    let mut result = String::new();
    for (k, v) in attrs {
        result.push_str(&format!(" {}=\"{}\"", k, v));
    }
    result
}
```

---

## 4. Recursion vs Iteration for XML Trees

### The Problem: Stack Overflow

**Recursive traversal** can overflow the stack on deeply nested XML:

```rust
// ❌ STACK OVERFLOW RISK: Deep recursion
fn traverse_recursive(doc: &XmlDocument, node: NodeId) {
    if let Some(data) = doc.get(node) {
        process(data);
        for child in doc.children(node) {
            traverse_recursive(doc, child); // Recursive call
        }
    }
}
```

**Typical stack sizes**:
- Linux/macOS: 8 MB (default)
- Windows: 1 MB (default)
- WASM: 1 MB (configurable)

**Depth limits**:
- ~10,000 levels on Linux/macOS
- ~1,000 levels on Windows
- Varies by frame size

### Solution 1: Iterative Traversal with Explicit Stack

**Use when**: You need full control and want to avoid stack overflow.

```rust
// ✅ STACK SAFE: Iterative with Vec stack
fn traverse_iterative(doc: &XmlDocument, root: NodeId) {
    let mut stack = vec![root];
    
    while let Some(node) = stack.pop() {
        if let Some(data) = doc.get(node) {
            process(data);
            
            // Push children in reverse order for depth-first
            let children: Vec<_> = doc.children(node).collect();
            for child in children.into_iter().rev() {
                stack.push(child);
            }
        }
    }
}
```

**Tradeoffs**:
- ✅ No stack overflow risk
- ✅ Heap allocation (grows as needed)
- ✅ Explicit control over traversal order
- ❌ More verbose than recursion
- ❌ Slight performance overhead (heap allocation)

### Solution 2: Iterator-Based Traversal

**Use when**: You want clean code and the library provides iterators.

```rust
// ✅ STACK SAFE: Library-provided iterator
fn traverse_with_iterator(doc: &XmlDocument, root: NodeId) {
    for node in doc.descendants(root) {
        if let Some(data) = doc.get(node) {
            process(data);
        }
    }
}
```

**Example from `redline-rs`** (see `util/descendants.rs`):

```rust
// Iterative implementation under the hood
pub fn descendants_trimmed<'a, F>(
    doc: &'a XmlDocument,
    node: NodeId,
    trim_predicate: F,
) -> impl Iterator<Item = NodeId> + 'a
where
    F: Fn(&XmlNodeData) -> bool + 'a,
{
    DescendantsTrimmedIter::new(doc, node, trim_predicate)
}

struct DescendantsTrimmedIter<'a, F> {
    doc: &'a XmlDocument,
    stack: Vec<NodeId>, // ✅ Heap-allocated stack
    trim_predicate: F,
}
```

**Tradeoffs**:
- ✅ Clean, idiomatic Rust
- ✅ No stack overflow risk
- ✅ Composable with other iterators
- ❌ Requires library support

### Solution 3: Tail Recursion (Limited Use)

**Use when**: You have simple tail-recursive patterns (rare in tree traversal).

```rust
// ✅ TAIL RECURSIVE: Compiler may optimize
fn count_nodes_tail(doc: &XmlDocument, node: NodeId, acc: usize) -> usize {
    let children: Vec<_> = doc.children(node).collect();
    if children.is_empty() {
        acc + 1
    } else {
        // Tail call (last operation)
        count_nodes_tail(doc, children[0], acc + 1)
    }
}
```

**⚠️ WARNING**: Rust does **not guarantee** tail call optimization. Use explicit iteration for safety.

---

## 5. Hashing for XML Trees

### Deterministic Hashing Requirements

For document comparison, hashes must be:
1. **Reproducible**: Same input → same hash
2. **Stable**: Same across platforms and runs
3. **Collision-resistant**: Different inputs → different hashes

### Cryptographic Hashing (SHA1/SHA256)

**Use when**: You need strong collision resistance and cross-platform stability.

```rust
use sha1::{Digest, Sha1};

// ✅ DETERMINISTIC: SHA1 is stable
fn hash_xml_content(xml: &str) -> String {
    let mut hasher = Sha1::new();
    hasher.update(xml.as_bytes());
    let result = hasher.finalize();
    format!("{:x}", result)
}
```

**Example from `redline-rs`** (see `wml/block_hash.rs`):

```rust
pub fn compute_block_hash(
    doc: &XmlDocument,
    node: NodeId,
    settings: &HashingSettings,
) -> String {
    let xml_string = clone_block_level_content_for_hashing(doc, node, settings);
    let mut hasher = Sha1::new();
    hasher.update(xml_string.as_bytes());
    let result = hasher.finalize();
    format!("{:x}", result) // ✅ Deterministic hex string
}
```

**Tradeoffs**:
- ✅ Stable across platforms and Rust versions
- ✅ Strong collision resistance
- ✅ Suitable for persistent storage
- ❌ Slower than non-cryptographic hashes
- ❌ Overkill for in-memory deduplication

### Non-Cryptographic Hashing (ahash, FxHash)

**Use when**: You need fast hashing for in-memory data structures (NOT for persistence).

```rust
use ahash::AHasher;
use std::hash::{Hash, Hasher};

// ❌ NONDETERMINISTIC: Random seed by default
fn hash_fast<T: Hash>(value: &T) -> u64 {
    let mut hasher = AHasher::default();
    value.hash(&mut hasher);
    hasher.finish()
}

// ✅ DETERMINISTIC: Fixed seed
use ahash::RandomState;

fn hash_fast_deterministic<T: Hash>(value: &T) -> u64 {
    let state = RandomState::with_seeds(1, 2, 3, 4);
    let mut hasher = state.build_hasher();
    value.hash(&mut hasher);
    hasher.finish()
}
```

**⚠️ WARNING**: Non-cryptographic hashes are **not stable** across:
- Different `ahash` versions
- Different platforms (endianness)
- Different Rust versions

**Use only for**:
- In-memory deduplication
- Temporary caching
- Performance-critical hot paths

---

## 6. XML-Specific Determinism Patterns

### Pattern 1: Attribute Normalization

**Problem**: Attributes can appear in any order in source XML.

```rust
// ✅ DETERMINISTIC: Sort attributes before hashing
use std::collections::BTreeMap;

fn normalize_attributes(attrs: &[(String, String)]) -> BTreeMap<String, String> {
    attrs.iter().cloned().collect() // BTreeMap auto-sorts
}
```

**Example from `redline-rs`** (see `wml/block_hash.rs:173-182`):

```rust
for attr in attributes {
    if is_rsid_attribute(&attr.name) || is_pt_namespace(&attr.name) {
        continue; // Skip nondeterministic attributes
    }
    output.push(' ');
    output.push_str(&attr.name.local_name);
    output.push_str("=\"");
    output.push_str(&escape_xml_attr(&attr.value));
    output.push('"');
}
```

**Note**: Attributes are already stored in a `Vec` in `redline-rs`, so order is preserved from parsing. For hashing, we iterate in storage order (deterministic if parser is deterministic).

### Pattern 2: Namespace Prefix Normalization

**Problem**: XML allows different prefixes for the same namespace.

```xml
<!-- Semantically identical, different prefixes -->
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
<word:p xmlns:word="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
```

**Solution**: Normalize to canonical prefixes before hashing.

```rust
// ✅ DETERMINISTIC: Canonical prefix mapping
fn get_prefix_for_namespace(ns: Option<&str>) -> Option<&'static str> {
    match ns {
        Some("http://schemas.openxmlformats.org/wordprocessingml/2006/main") => Some("w"),
        Some("urn:schemas-microsoft-com:vml") => Some("v"),
        Some("urn:schemas-microsoft-com:office:office") => Some("o"),
        _ => None,
    }
}
```

**Example from `redline-rs`** (see `wml/block_hash.rs:370-381`).

### Pattern 3: Whitespace Normalization

**Problem**: Insignificant whitespace can vary.

```rust
// ✅ DETERMINISTIC: Normalize whitespace
fn normalize_text(text: &str, settings: &HashingSettings) -> String {
    let mut result = text.to_string();
    if settings.case_insensitive {
        result = result.to_uppercase();
    }
    if settings.conflate_spaces {
        result = result.replace(' ', "\u{00a0}"); // Non-breaking space
    }
    result
}
```

**Example from `redline-rs`** (see `wml/block_hash.rs:202-210`).

### Pattern 4: Skipping Nondeterministic Metadata

**Problem**: Some XML attributes are nondeterministic (timestamps, IDs, revision IDs).

```rust
// ✅ DETERMINISTIC: Skip RSID attributes
fn is_rsid_attribute(name: &XName) -> bool {
    if let Some(ns) = &name.namespace {
        if ns == W::NS {
            let local = &name.local_name;
            return local == "rsid"
                || local == "rsidDel"
                || local == "rsidR"
                || local == "rsidRDefault"
                || local == "rsidRPr"
                || local == "rsidSect"
                || local == "rsidTr";
        }
    }
    false
}
```

**Example from `redline-rs`** (see `wml/block_hash.rs:15-30`).

---

## 7. Recommendations for `redline-rs`

### Current State Analysis

**Good**:
- ✅ Uses `Vec` for attributes (deterministic order from parser)
- ✅ Uses `indextree::Arena` for tree structure (deterministic traversal)
- ✅ Uses SHA1 for block hashing (stable, reproducible)
- ✅ Iterative traversal with explicit stack (see `descendants.rs`)
- ✅ Normalizes attributes and namespaces before hashing

**Potential Issues**:
- ⚠️ No `HashMap`/`HashSet` usage detected in current code (good!)
- ⚠️ If adding caching, use `BTreeMap` or `IndexMap`

### Recommendations

#### 1. For Attribute Storage

**Current**: `Vec<XAttribute>` (good for small attribute counts)

**If optimizing lookups**:
```rust
// Option A: BTreeMap (sorted, deterministic)
use std::collections::BTreeMap;
pub type AttributeMap = BTreeMap<XName, String>;

// Option B: IndexMap (insertion order, deterministic)
use indexmap::IndexMap;
pub type AttributeMap = IndexMap<XName, String>;
```

**Recommendation**: Keep `Vec<XAttribute>` unless profiling shows lookup performance issues. Sorting before hashing is cheap.

#### 2. For Caching/Memoization

**If adding hash caches**:
```rust
// ✅ DETERMINISTIC: BTreeMap for caches
use std::collections::BTreeMap;

pub struct HashCache {
    cache: BTreeMap<NodeId, String>, // NodeId is Copy, Ord
}
```

**Avoid**:
```rust
// ❌ NONDETERMINISTIC: HashMap for caches
use std::collections::HashMap;
pub struct HashCache {
    cache: HashMap<NodeId, String>, // Iteration order varies
}
```

#### 3. For Tree Traversal

**Current**: Iterative with `Vec` stack (excellent!)

**Keep using**:
```rust
// ✅ STACK SAFE: Explicit stack
let mut stack = vec![root];
while let Some(node) = stack.pop() {
    // Process node
    for child in doc.children(node).rev() {
        stack.push(child);
    }
}
```

**Avoid**:
```rust
// ❌ STACK OVERFLOW RISK: Deep recursion
fn traverse(node: NodeId) {
    for child in doc.children(node) {
        traverse(child); // Recursive
    }
}
```

#### 4. For Testing

**Add determinism checks**:
```rust
#[test]
fn test_hash_determinism() {
    let doc = load_test_document();
    let hash1 = compute_block_hash(&doc, root, &settings);
    let hash2 = compute_block_hash(&doc, root, &settings);
    assert_eq!(hash1, hash2, "Hash must be deterministic");
}

#[test]
fn test_hash_stability_across_runs() {
    let doc = load_test_document();
    let hash = compute_block_hash(&doc, root, &settings);
    // Golden hash from previous run
    assert_eq!(hash, "a1b2c3d4e5f6...", "Hash must be stable");
}
```

---

## 8. Quick Reference Table

| Use Case | Recommended Type | Deterministic? | Performance |
|----------|------------------|----------------|-------------|
| Attribute storage (small) | `Vec<XAttribute>` | ✅ Yes | O(n) lookup |
| Attribute storage (large) | `BTreeMap<XName, String>` | ✅ Yes | O(log n) lookup |
| Insertion-order map | `IndexMap<K, V>` | ✅ Yes | O(1) lookup |
| Sorted map | `BTreeMap<K, V>` | ✅ Yes | O(log n) lookup |
| Fast in-memory cache | `HashMap<K, V>` | ❌ No | O(1) lookup |
| Tree traversal | Iterative with `Vec` stack | ✅ Yes | O(n) |
| Cryptographic hash | `sha1::Sha1`, `sha2::Sha256` | ✅ Yes | Slow |
| Fast hash (temp) | `ahash` with fixed seed | ⚠️ Conditional | Fast |

---

## 9. Common Pitfalls

### Pitfall 1: Assuming HashMap Iteration Order

```rust
// ❌ WRONG: Assumes order
let mut map = HashMap::new();
map.insert("a", 1);
map.insert("b", 2);
let keys: Vec<_> = map.keys().collect();
assert_eq!(keys, vec![&"a", &"b"]); // FAILS randomly
```

**Fix**:
```rust
// ✅ CORRECT: Sort before comparing
let mut keys: Vec<_> = map.keys().collect();
keys.sort();
assert_eq!(keys, vec![&"a", &"b"]);
```

### Pitfall 2: Using `Hash` for Persistence

```rust
use std::hash::{Hash, Hasher};
use std::collections::hash_map::DefaultHasher;

// ❌ WRONG: Hash value changes across Rust versions
fn persist_hash<T: Hash>(value: &T) -> u64 {
    let mut hasher = DefaultHasher::new();
    value.hash(&mut hasher);
    hasher.finish() // NOT STABLE
}
```

**Fix**:
```rust
use sha2::{Sha256, Digest};

// ✅ CORRECT: Use cryptographic hash
fn persist_hash(value: &str) -> String {
    let mut hasher = Sha256::new();
    hasher.update(value.as_bytes());
    format!("{:x}", hasher.finalize())
}
```

### Pitfall 3: Deep Recursion Without Limits

```rust
// ❌ WRONG: No depth limit
fn traverse(node: NodeId) {
    for child in children(node) {
        traverse(child); // Stack overflow on deep trees
    }
}
```

**Fix**:
```rust
// ✅ CORRECT: Iterative traversal
fn traverse(root: NodeId) {
    let mut stack = vec![root];
    while let Some(node) = stack.pop() {
        for child in children(node) {
            stack.push(child);
        }
    }
}
```

---

## 10. Summary Checklist

- [ ] Use `BTreeMap`/`BTreeSet` or `IndexMap`/`IndexSet` instead of `HashMap`/`HashSet` when iteration order matters
- [ ] Sort keys before iterating `HashMap` if you must use it
- [ ] Use SHA1/SHA256 for persistent hashes, not `std::hash::Hash`
- [ ] Use iterative traversal with explicit `Vec` stack for deep trees
- [ ] Normalize XML attributes, namespaces, and whitespace before hashing
- [ ] Skip nondeterministic metadata (RSIDs, timestamps) in hashes
- [ ] Add determinism tests (same input → same output, multiple runs)
- [ ] Document any intentional nondeterminism (e.g., UUIDs)

---

## References

- [Rust `std::collections` docs](https://doc.rust-lang.org/std/collections/)
- [`indexmap` crate](https://docs.rs/indexmap/)
- [`ahash` crate](https://docs.rs/ahash/)
- [SHA1 crate](https://docs.rs/sha1/)
- [SHA2 crate](https://docs.rs/sha2/)
- [`indextree` crate](https://docs.rs/indextree/) (used in `redline-rs`)

---

**Document Version**: 1.0  
**Last Updated**: 2025-12-26  
**Project**: `redline-rs` (OpenXML PowerTools Rust port)
