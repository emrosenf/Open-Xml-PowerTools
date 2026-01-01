# C# WmlComparer Text Content Handling - Investigation Report

**Investigation Date:** 2025-12-29  
**Agent:** csharp-investigator  
**Task:** Investigate how C# OpenXmlPowerTools WmlComparer handles text content during comparison output reconstruction

## Executive Summary

The bug in the Rust port (missing `w:t` elements) is caused by not properly extracting and using `ContentElement.Value` during reconstruction. The C# code explicitly reads the **text content** from `ContentElement.Value` and concatenates it when building `w:t` elements.

**Critical Finding:** The C# code creates `w:t` elements with text content by:
1. Extracting `ContentElement.Value` from each `ComparisonUnitAtom`
2. Concatenating these text values together
3. Creating a new `XElement(W.t, textContent)` with the concatenated text

## Key Architecture Components

### 1. ComparisonUnitAtom Structure

**File:** `WmlComparer.cs`  
**Location:** Lines 8280-8479

```csharp
public class ComparisonUnitAtom : ComparisonUnit
{
    public XElement[] AncestorElements;
    public string[] AncestorUnids;
    public XElement ContentElement;        // ← THIS HOLDS THE CONTENT
    public XElement ContentElementBefore;
    public ComparisonUnitAtom ComparisonUnitAtomBefore;
    public OpenXmlPart Part;
    // ... more fields
}
```

**Key Points:**
- `ContentElement` is an `XElement` containing the atom's content
- For text atoms, `ContentElement` is created as `new XElement(W.t, ch)` where `ch` is a single character
- The text value is accessed via `ContentElement.Value` property

### 2. Atom Creation from Source Documents

**File:** `WmlComparer.cs`  
**Location:** Lines 8033-8118  
**Function:** `CreateComparisonUnitAtomListRecurse`

This is where atoms are created from the source XML. For text elements:

```csharp
if (element.Name == W.t || element.Name == W.delText)
{
    var val = element.Value;  // Get the text content
    foreach (var ch in val)   // Split into individual characters
    {
        ComparisonUnitAtom sr = new ComparisonUnitAtom(
            new XElement(element.Name, ch),  // ← Create XElement with single char
            element.AncestorsAndSelf()
                .TakeWhile(a => a.Name != W.body && ...)
                .Reverse()
                .ToArray(),
            part,
            settings);
        comparisonUnitAtomList.Add(sr);
    }
    return;
}
```

**Key Insight:**  
- Text elements (`w:t`) are broken down into **individual character atoms**
- Each atom stores the character in `ContentElement` as `new XElement(W.t, ch)`
- This means `ContentElement.Value` will return that single character

### 3. Document Reconstruction Entry Point

**File:** `WmlComparer.cs`  
**Location:** Lines 2027-2238  
**Function:** `ProduceDocumentWithTrackedRevisions`

The main reconstruction flow:

```csharp
// Line 2114: Flatten comparison tree to atom list
var listOfComparisonUnitAtoms = FlattenToComparisonUnitAtomList(correlatedSequence, settings);

// Line 2139: Set up ancestor unids for tree reconstruction
AssembleAncestorUnidsInOrderToRebuildXmlTreeProperly(listOfComparisonUnitAtoms);

// Line 2165-2166: THE CRITICAL CALL - produce new XML from atoms
var newBodyChildren = ProduceNewWmlMarkupFromCorrelatedSequence(
    wDocWithRevisions.MainDocumentPart,
    listOfComparisonUnitAtoms, 
    settings);
```

### 4. Reconstruction Core - ProduceNewWmlMarkupFromCorrelatedSequence

**File:** `WmlComparer.cs`  
**Location:** Lines 5014-5022  
**Function:** `ProduceNewWmlMarkupFromCorrelatedSequence`

```csharp
private static object ProduceNewWmlMarkupFromCorrelatedSequence(
    OpenXmlPart part,
    IEnumerable<ComparisonUnitAtom> comparisonUnitAtomList,
    WmlComparerSettings settings)
{
    // fabricate new MainDocumentPart from correlatedSequence
    s_MaxId = 0;
    var newBodyChildren = CoalesceRecurse(part, comparisonUnitAtomList, 0, settings);
    return newBodyChildren;
}
```

This delegates to `CoalesceRecurse`, which is the heart of reconstruction.

### 5. THE CRITICAL FUNCTION: CoalesceRecurse

**File:** `WmlComparer.cs`  
**Location:** Lines 5161-5599  
**Function:** `CoalesceRecurse`

This recursive function rebuilds the XML tree from atoms. The critical section for `w:t` elements:

#### Text Element Reconstruction (Lines 5401-5426)

```csharp
if (ancestorBeingConstructed.Name == W.t)
{
    var newChildElements = groupedChildren
        .Select(gc =>
        {
            // ═══════════════════════════════════════════════════════════
            // THIS IS THE CRITICAL LINE - LINE 5406
            // ═══════════════════════════════════════════════════════════
            var textOfTextElement = gc.Select(gce => gce.ContentElement.Value)
                                      .StringConcatenate();
            
            var del = gc.First().CorrelationStatus == CorrelationStatus.Deleted;
            var ins = gc.First().CorrelationStatus == CorrelationStatus.Inserted;
            
            if (del)
                return (object)(new XElement(W.delText,
                    new XAttribute(PtOpenXml.Status, "Deleted"),
                    GetXmlSpaceAttribute(textOfTextElement),
                    textOfTextElement));  // ← TEXT CONTENT HERE
            else if (ins)
                return (object)(new XElement(W.t,
                    new XAttribute(PtOpenXml.Status, "Inserted"),
                    GetXmlSpaceAttribute(textOfTextElement),
                    textOfTextElement));  // ← TEXT CONTENT HERE
            else
                return (object)(new XElement(W.t,
                    GetXmlSpaceAttribute(textOfTextElement),
                    textOfTextElement));  // ← TEXT CONTENT HERE
        })
        .ToList();
    return newChildElements;
}
```

**THE KEY LINE - Line 5406:**
```csharp
var textOfTextElement = gc.Select(gce => gce.ContentElement.Value).StringConcatenate();
```

**What This Does:**
1. `gc` is a group of `ComparisonUnitAtom` objects that should form one text element
2. For each atom `gce`, it extracts `gce.ContentElement.Value` (the text character)
3. It concatenates all these values together using `.StringConcatenate()`
4. The result `textOfTextElement` is the actual text string (e.g., "Hello")
5. This string is then used as the **content** of the new `XElement(W.t, textOfTextElement)`

### 6. Run Element Reconstruction (Lines 5334-5399)

```csharp
if (ancestorBeingConstructed.Name == W.r)
{
    var newChildElements = groupedChildren
        .Select(gc =>
        {
            var spl = gc.Key.Split('|');
            if (spl[0] == "")
                return (object)gc.Select(gcc =>
                {
                    // For run children (not recursing deeper)
                    var contentElement = (isInsideVml && gcc.ContentElementBefore != null)
                        ? gcc.ContentElementBefore
                        : gcc.ContentElement;
                    var dup = new XElement(contentElement);  // Clone the element
                    if (spl[1] == "Deleted")
                        dup.Add(new XAttribute(PtOpenXml.Status, "Deleted"));
                    else if (spl[1] == "Inserted")
                        dup.Add(new XAttribute(PtOpenXml.Status, "Inserted"));
                    return dup;
                });
            else
            {
                return CoalesceRecurse(part, gc, level + 1, settings);  // Recurse
            }
        })
        .ToList();

    // Create the run with rPr and children
    var newRun = new XElement(W.r,
        ancestorBeingConstructed.Attributes()...,
        rPr,
        newChildElements);  // ← Children include w:t elements

    return newRun;
}
```

**Key Point:** When reconstructing a run, if there are deeper children (like `w:t`), it recurses down. The recursion will hit the `w:t` case above, which extracts the text content.

### 7. Paragraph Element Reconstruction (Lines 5292-5332)

```csharp
if (ancestorBeingConstructed.Name == W.p)
{
    var newChildElements = groupedChildren
        .Select(gc =>
        {
            var spl = gc.Key.Split('|');
            if (spl[0] == "")  // No deeper ancestors - these are direct children
                return (object)gc.Select(gcc =>
                {
                    var contentElement = (isInsideVml && gcc.ContentElementBefore != null)
                        ? gcc.ContentElementBefore
                        : gcc.ContentElement;
                    var dup = new XElement(contentElement);
                    // ... status attributes
                    return dup;
                }).Where(e => e != null);
            else  // Has deeper ancestors - recurse
            {
                return CoalesceRecurse(part, gc, level + 1, settings);
            }
        })
        .ToList();

    var newPara = new XElement(W.p,
        ancestorBeingConstructed.Attributes()...,
        new XAttribute(PtOpenXml.Unid, g.Key),
        newChildElements);  // ← Children include runs

    return newPara;
}
```

## Complete Code Flow for Text Output

### Flow Diagram

```
1. ProduceDocumentWithTrackedRevisions (line 2027)
   ↓
2. ProduceNewWmlMarkupFromCorrelatedSequence (line 5014)
   ↓
3. CoalesceRecurse (line 5161) - level=0 (body children)
   ├─ Groups atoms by ancestor unid at level 0
   ├─ For each group, gets ancestorBeingConstructed (e.g., W.p)
   └─ Handles W.p case (line 5292)
      ├─ Groups children by next level ancestor
      └─ Recurses with level=1
         ↓
4. CoalesceRecurse (line 5161) - level=1 (paragraph children)
   ├─ Groups atoms by ancestor unid at level 1
   ├─ Gets ancestorBeingConstructed (e.g., W.r)
   └─ Handles W.r case (line 5334)
      ├─ Groups children by next level ancestor
      └─ Recurses with level=2
         ↓
5. CoalesceRecurse (line 5161) - level=2 (run children)
   ├─ Groups atoms by ancestor unid at level 2
   ├─ Gets ancestorBeingConstructed (e.g., W.t)
   └─ Handles W.t case (line 5401) ← THE CRITICAL SECTION
      ├─ Line 5406: Extract text from each atom's ContentElement.Value
      ├─ Concatenate all text values together
      └─ Create new XElement(W.t, concatenatedText)
```

### Concrete Example

**Input atoms** (simplified):
```
Atom1: ContentElement = <w:t>H</w:t>, AncestorUnids = ["body-1", "p-1", "r-1", "t-1"]
Atom2: ContentElement = <w:t>e</w:t>, AncestorUnids = ["body-1", "p-1", "r-1", "t-1"]
Atom3: ContentElement = <w:t>l</w:t>, AncestorUnids = ["body-1", "p-1", "r-1", "t-1"]
Atom4: ContentElement = <w:t>l</w:t>, AncestorUnids = ["body-1", "p-1", "r-1", "t-1"]
Atom5: ContentElement = <w:t>o</w:t>, AncestorUnids = ["body-1", "p-1", "r-1", "t-1"]
```

**At level 2 (run children), ancestor="t-1":**

Line 5406 executes:
```csharp
var textOfTextElement = gc.Select(gce => gce.ContentElement.Value).StringConcatenate();
// gc contains [Atom1, Atom2, Atom3, Atom4, Atom5]
// .Select(gce => gce.ContentElement.Value) → ["H", "e", "l", "l", "o"]
// .StringConcatenate() → "Hello"
```

Line 5415-5418 creates:
```csharp
return (object)(new XElement(W.t,
    GetXmlSpaceAttribute("Hello"),
    "Hello"));  // ← The actual text string
```

**Output XML:**
```xml
<w:t>Hello</w:t>
```

## Critical Differences to Look for in Rust Implementation

### 1. **ContentElement Must Be an XElement (or equivalent)**

The C# code stores `ContentElement` as an `XElement`, which has a `.Value` property that returns the text content.

**Rust equivalent check:**
- Does `ComparisonUnitAtom.content_element` store the actual XML element?
- Can you extract the text value from it?

### 2. **Text Value Extraction (Line 5406)**

The C# code explicitly calls `gce.ContentElement.Value` to get the text.

**Rust equivalent check:**
- Does the Rust code extract the text value from `content_element`?
- Or does it just clone/use the element without reading its text?

### 3. **String Concatenation**

The C# code uses `.StringConcatenate()` to join all the character strings.

**Rust equivalent check:**
- Does the Rust code concatenate the text values?
- Or does it just collect the elements without merging text?

### 4. **Creating w:t with Text Content**

The C# code creates: `new XElement(W.t, textOfTextElement)` where `textOfTextElement` is a **string**.

**Rust equivalent check:**
- Does the Rust code create a text node inside the `w:t` element?
- The XElement constructor `XElement(name, content)` creates an element with text content
- Rust needs to create: `<w:t>Hello</w:t>`, not just `<w:t/>`

## Specific Line Numbers and Code Snippets

### Line 5406 - The Critical Text Extraction
```csharp
var textOfTextElement = gc.Select(gce => gce.ContentElement.Value).StringConcatenate();
```

### Lines 5409-5422 - Creating w:t Elements with Text
```csharp
if (del)
    return (object)(new XElement(W.delText,
        new XAttribute(PtOpenXml.Status, "Deleted"),
        GetXmlSpaceAttribute(textOfTextElement),
        textOfTextElement));  // ← Text as XElement content
else if (ins)
    return (object)(new XElement(W.t,
        new XAttribute(PtOpenXml.Status, "Inserted"),
        GetXmlSpaceAttribute(textOfTextElement),
        textOfTextElement));  // ← Text as XElement content
else
    return (object)(new XElement(W.t,
        GetXmlSpaceAttribute(textOfTextElement),
        textOfTextElement));  // ← Text as XElement content
```

### Line 8086-8091 - Creating Character Atoms
```csharp
ComparisonUnitAtom sr = new ComparisonUnitAtom(
    new XElement(element.Name, ch),  // ← Single character in XElement
    element.AncestorsAndSelf()...,
    part,
    settings);
```

## Debugging Checklist for Rust Code

1. **Verify atom creation:**
   - Check how text elements are split into atoms
   - Verify that `content_element` stores the text character

2. **Verify CoalesceRecurse logic:**
   - Find the equivalent of line 5401 (`if ancestorBeingConstructed.Name == W.t`)
   - Check if it extracts text from `content_element`

3. **Verify text concatenation:**
   - Look for the equivalent of line 5406
   - Check if text values are being concatenated

4. **Verify XElement creation:**
   - Check how `w:t` elements are created
   - Verify that text content is passed to the element constructor

5. **Check for missing text node creation:**
   - The most likely bug: creating `<w:t/>` instead of `<w:t>Hello</w:t>`
   - This happens if the code creates an empty element instead of adding text content

## Expected Fix Location (Hypothesis)

Based on this analysis, the Rust bug is most likely in the equivalent of `CoalesceRecurse` at the section handling `W.t` elements. Specifically:

1. **Missing text extraction:** Not calling the equivalent of `gce.ContentElement.Value`
2. **Missing concatenation:** Not joining the character values together
3. **Missing text node:** Creating an empty `<w:t/>` element instead of `<w:t>text</w:t>`

The fix should:
1. Extract text from each atom's `content_element` 
2. Concatenate the text values
3. Create a text node inside the `w:t` element with the concatenated text

## Related Functions

- `GetXmlSpaceAttribute`: Returns `xml:space="preserve"` if needed (lines 2774-2779)
- `StringConcatenate`: Extension method to concatenate strings (likely in PtUtils or similar)
- `MarkContentAsDeletedOrInserted`: Wraps content in `w:ins`/`w:del` elements (line 2173)
- `CoalesceAdjacentRunsWithIdenticalFormatting`: Merges adjacent runs (line 2174)

## Conclusion

The C# WmlComparer creates `w:t` elements with text content by:

1. **Storing text in atoms:** Each `ComparisonUnitAtom` has a `ContentElement` (XElement) containing a single character
2. **Extracting text values:** Using `ContentElement.Value` to get the text string
3. **Concatenating text:** Joining all character values together
4. **Creating w:t with text:** Passing the concatenated string as content to `new XElement(W.t, text)`

The Rust port must replicate this exact pattern, especially the critical line 5406 that extracts and concatenates text values from `ContentElement.Value`.
