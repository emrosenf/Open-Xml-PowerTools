# Investigation Summary: C# Text Content Handling

**Date:** 2025-12-29  
**Task:** Find how C# WmlComparer outputs text content in w:t elements

## The Critical Line

**File:** `OpenXmlPowerTools/WmlComparer.cs`  
**Line:** 5406

```csharp
var textOfTextElement = gc.Select(gce => gce.ContentElement.Value).StringConcatenate();
```

## What This Does

1. Takes a group of `ComparisonUnitAtom` objects (`gc`)
2. For each atom, extracts `ContentElement.Value` (the text character)
3. Concatenates all values together using `.StringConcatenate()`
4. Uses the result as the text content when creating `new XElement(W.t, textOfTextElement)`

## Code Flow

```
CreateComparisonUnitAtomListRecurse (line 8081-8093)
├─ Splits text into character atoms
└─ Each atom: new ComparisonUnitAtom(new XElement(W.t, ch), ...)
    ↓
ProduceDocumentWithTrackedRevisions (line 2027)
    ↓
ProduceNewWmlMarkupFromCorrelatedSequence (line 5014)
    ↓
CoalesceRecurse (line 5161)
├─ Groups atoms by ancestor unid
├─ Recurses down tree hierarchy
└─ When ancestor is W.t (line 5401):
    ├─ Line 5406: Extract & concatenate text from ContentElement.Value
    └─ Lines 5409-5422: Create XElement(W.t, concatenatedText)
```

## Example

**Input atoms:**
```
Atom1: ContentElement = <w:t>H</w:t>
Atom2: ContentElement = <w:t>e</w:t>
Atom3: ContentElement = <w:t>l</w:t>
Atom4: ContentElement = <w:t>l</w:t>
Atom5: ContentElement = <w:t>o</w:t>
```

**Line 5406 executes:**
```csharp
var textOfTextElement = gc.Select(gce => gce.ContentElement.Value).StringConcatenate();
// → ["H", "e", "l", "l", "o"].join() → "Hello"
```

**Lines 5415-5418 create:**
```csharp
new XElement(W.t, GetXmlSpaceAttribute("Hello"), "Hello")
```

**Output:**
```xml
<w:t>Hello</w:t>
```

## What the Rust Code Needs to Do

1. **Store XML elements in atoms** - `content_element` must contain the actual XML element
2. **Extract text values** - Read the text content from `content_element` (like `.Value` in C#)
3. **Concatenate text** - Join all character values together
4. **Create text nodes** - Create `<w:t>text</w:t>` with the concatenated string, NOT `<w:t/>`

## Most Likely Bug Location

The Rust equivalent of `CoalesceRecurse` at the section handling `W.t` elements is probably:
- Missing the text extraction step (line 5406 equivalent)
- Missing the concatenation logic
- Creating empty elements instead of elements with text content

## Quick Fix Checklist

- [ ] Find the Rust equivalent of `CoalesceRecurse`
- [ ] Locate the section that handles `w:t` elements
- [ ] Verify it extracts text from `content_element`
- [ ] Verify it concatenates the text values
- [ ] Verify it creates a text node inside the `w:t` element
- [ ] Test with a simple "Hello" text comparison

## See Also

- `C_SHARP_TEXT_HANDLING_ANALYSIS.md` - Complete detailed analysis with all line numbers and code snippets
