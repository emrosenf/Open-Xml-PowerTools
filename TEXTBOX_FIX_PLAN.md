# TextBox Comparison Round-Trip Fix Plan

## Problem Statement

TextBox-related tests in WmlComparer are failing sanity checks. The sanity checks verify that:

1. **Sanity Check #1**: Compare(A, B) -> reject revisions -> Compare with A = 0 differences
2. **Sanity Check #2**: Compare(A, B) -> accept revisions -> Compare with B = 0 differences

**Root Cause**: VML (Vector Markup Language) property elements are lost during document reconstruction.

### Technical Details

TextBoxes in Word documents use VML markup with this structure:
```xml
<w:pict>
  <v:shapetype id="_x0000_t202" ...>
    <v:stroke joinstyle="miter"/>
    <v:path gradientshapeok="t" o:connecttype="rect"/>
  </v:shapetype>
  <v:shape id="_x0000_s1026" type="#_x0000_t202" style="...">
    <v:fill color2="#d9d9d9"/>
    <v:stroke color="#969696"/>
    <v:shadow on="t" color="black" opacity="26214f"/>
    <v:textbox>
      <w:txbxContent>
        <w:p>
          <w:r><w:t>Text content</w:t></w:r>
        </w:p>
      </w:txbxContent>
    </v:textbox>
    <o:lock v:ext="edit" aspectratio="f"/>
  </v:shape>
</w:pict>
```

**The problem**: Elements like `v:fill`, `v:stroke`, `v:shadow`, `o:lock`, `v:path` have no text content. WmlComparer creates `ComparisonUnitAtom` instances only from text runs (`w:t`). During reconstruction:

1. `CoalesceRecurse` returns only elements that had atoms
2. `ReconstructElement` preserves only explicitly listed property children
3. VML property elements (not listed) are silently dropped

Current `RecursionInfo` configuration (WmlComparer.cs lines 7019-7033):
```csharp
{ ElementName = VML.shape, ChildElementPropertyNames = null }  // No properties preserved!
{ ElementName = VML.rect, ChildElementPropertyNames = null }   // No properties preserved!
{ ElementName = VML.group, ChildElementPropertyNames = null }  // No properties preserved!
{ ElementName = VML.textbox, ChildElementPropertyNames = null } // No properties preserved!
```

### Previous Failed Attempt

A `ReconstructVmlElement` function was added to preserve ALL non-recursive children. This failed because:
- It preserved elements that needed unique IDs (causing duplicate ID validation errors)
- It was too aggressive, preserving things that shouldn't be preserved

---

## Avenue 1: Update RecursionInfo with Specific VML Property Names

**Approach**: Add explicit lists of VML property children to preserve in `RecursionInfo`.

**Implementation**:
```csharp
new RecursionInfo()
{
    ElementName = VML.shape,
    ChildElementPropertyNames = new[] {
        VML.fill, VML.stroke, VML.shadow, VML.textpath, VML.path,
        VML.formulas, VML.handles, VML.imagedata,
        O._lock, O.extrusion, O.callout, O.signatureline,
        W10.wrap, W10.anchorlock
    },
},
new RecursionInfo()
{
    ElementName = VML.rect,
    ChildElementPropertyNames = new[] {
        VML.fill, VML.stroke, VML.shadow, VML.textpath, VML.path,
        O._lock, O.extrusion
    },
},
// Similar for VML.group, VML.oval, VML.line, etc.
```

**Pros**:
- Follows existing pattern (like W.tbl preserving W.tblPr)
- Surgical fix - minimal code changes
- Easy to understand and maintain

**Cons**:
- Must enumerate ALL VML property elements (easy to miss some)
- Different VML shape types may have different valid children
- Elements with relationship IDs (VML.fill, VML.imagedata, VML.stroke) may need relationship fixup

**Risk Level**: Low
**Complexity**: Low
**Files to Modify**: WmlComparer.cs (RecursionInfo array)

---

## Avenue 2: Preserve Non-Content VML Children Dynamically

**Approach**: Modify reconstruction logic to automatically preserve VML children that:
- Are NOT in the recursive content list (not containers like v:textbox, w:txbxContent)
- Have no text descendants OR are known property elements

**Implementation**:
```csharp
private static readonly HashSet<XName> VmlContentElements = new HashSet<XName>
{
    VML.textbox, W.txbxContent, VML.shape, VML.rect, VML.group,
    VML.oval, VML.line, VML.arc, VML.curve, VML.polyline, VML.roundrect
};

private static XElement ReconstructVmlElement(...)
{
    var newChildElements = CoalesceRecurse(part, g, level + 1, settings);

    // Preserve children that are NOT content containers
    var propertyChildren = ancestorBeingConstructed.Elements()
        .Where(e => !VmlContentElements.Contains(e.Name))
        .Where(e => !e.Descendants().Any(d => d.Name == W.t)); // No text content

    return new XElement(ancestorBeingConstructed.Name,
        ancestorBeingConstructed.Attributes(),
        propertyChildren,
        newChildElements);
}
```

**Pros**:
- Automatically handles all VML property types (no enumeration needed)
- Future-proof against new VML elements

**Cons**:
- May preserve unwanted elements
- Need to carefully define "content" vs "property" elements
- Previous similar attempt failed due to ID conflicts

**Risk Level**: Medium
**Complexity**: Medium
**Files to Modify**: WmlComparer.cs (CoalesceRecurse, new helper method)

---

## Avenue 3: Track VML Properties on ComparisonUnitAtom

**Approach**: When creating atoms from VML content, store the VML container's properties as metadata on the atoms. Restore during reconstruction.

**Implementation**:
```csharp
// In ComparisonUnitAtom class:
public List<XElement> VmlPropertyElements { get; set; }

// When creating atoms in CreateComparisonUnitAtomList:
if (ancestor.Name == VML.shape || ancestor.Name == VML.rect || ...)
{
    var vmlProps = ancestor.Elements()
        .Where(e => IsVmlPropertyElement(e))
        .Select(e => new XElement(e))
        .ToList();
    atom.VmlPropertyElements = vmlProps;
}

// During reconstruction in CoalesceRecurse:
if (g.Key contains VML shape ancestor)
{
    var vmlProps = g.First().VmlPropertyElements;
    // Include vmlProps in reconstructed element
}
```

**Pros**:
- Properties travel with their content through the entire comparison pipeline
- No chance of mismatched properties
- Clean conceptual model

**Cons**:
- Significant changes to ComparisonUnitAtom class
- Memory overhead storing property copies on every atom
- Complex reconstruction logic to merge properties from multiple atoms

**Risk Level**: Medium-High
**Complexity**: High
**Files to Modify**: WmlComparer.cs (ComparisonUnitAtom class, CreateComparisonUnitAtomList, CoalesceRecurse)

---

## Avenue 4: Pre/Post-Process VML Elements

**Approach**: Handle VML properties as a separate canonicalization step outside the main comparison logic.

**Implementation**:
```csharp
// Before comparison - extract and store VML properties
Dictionary<string, List<XElement>> vmlPropertiesById = new();

void ExtractVmlProperties(XElement doc)
{
    foreach (var shape in doc.Descendants(VML.shape))
    {
        var id = shape.Attribute("id")?.Value ?? GenerateId(shape);
        var props = shape.Elements()
            .Where(e => IsVmlPropertyElement(e))
            .Select(e => new XElement(e))
            .ToList();
        vmlPropertiesById[id] = props;
    }
}

// After comparison - restore VML properties
void RestoreVmlProperties(XElement doc)
{
    foreach (var shape in doc.Descendants(VML.shape))
    {
        var id = shape.Attribute("id")?.Value;
        if (id != null && vmlPropertiesById.TryGetValue(id, out var props))
        {
            // Add missing properties back to shape
            foreach (var prop in props)
            {
                if (!shape.Elements(prop.Name).Any())
                    shape.AddFirst(prop);
            }
        }
    }
}
```

**Pros**:
- Completely separates VML handling from comparison logic
- Easy to debug and test independently
- Can handle complex VML scenarios without affecting core algorithm

**Cons**:
- ID matching may fail if IDs change during comparison
- Doesn't handle cases where VML structure changes (shape split/merged)
- Two-pass approach adds complexity
- May restore properties to wrong shapes if document structure changes significantly

**Risk Level**: Medium
**Complexity**: Medium
**Files to Modify**: WmlComparer.cs (Compare method, new helper class/methods)

---

## Recommendation

**Start with Avenue 1** (Update RecursionInfo) because:
1. It's the least invasive change
2. Follows established patterns in the codebase
3. Lowest risk of breaking other tests
4. Easy to verify and rollback

If Avenue 1 fails due to relationship ID issues with VML.fill/VML.imagedata/VML.stroke, consider:
- Adding relationship fixup logic similar to `MoveRelatedPartsToDestination`
- Or falling back to Avenue 2 with careful element filtering

## Special Considerations

### Elements with Relationship IDs
These VML elements can contain relationship IDs to external resources:
- `VML.fill` - can reference image fills
- `VML.imagedata` - references image data
- `VML.stroke` - can reference dash patterns

These may need the same relationship fixup that's done for other elements (see `MoveRelatedPartsToDestination`).

### ID Uniqueness
VML shapes have IDs (e.g., `_x0000_s1026`). The existing `FixUpShapeIds`, `FixUpGroupIds`, `FixUpRectIds` functions handle ID deduplication. The fix should work with this existing infrastructure.

### Test Files
Skipped tests to re-enable after fix:
- WC-1770: WC037-Textbox-Before.docx / WC037-Textbox-After1.docx
- WC-1860: WC044-Text-Box.docx / WC044-Text-Box-Mod.docx
- WC-1870: WC045-Text-Box.docx / WC045-Text-Box-Mod.docx
- WC-1880: WC046-Two-Text-Box.docx / WC046-Two-Text-Box-Mod.docx
- WC-1890: WC047-Two-Text-Box.docx / WC047-Two-Text-Box-Mod.docx
- WC-1900: WC048-Text-Box-in-Cell.docx / WC048-Text-Box-in-Cell-Mod.docx
- WC-1910: WC049-Text-Box-in-Cell.docx / WC049-Text-Box-in-Cell-Mod.docx
- WC-1920: WC050-Table-in-Text-Box.docx / WC050-Table-in-Text-Box-Mod.docx
- WC-1930: WC051-Table-in-Text-Box.docx / WC051-Table-in-Text-Box-Mod.docx
- WC-1990: WC057-Table-Merged-Cell.docx / WC057-Table-Merged-Cell-Mod.docx
- WC-2080: WC065-Textbox.docx / WC065-Textbox-Mod.docx
- WC-2100: WC067-Textbox-Image.docx / WC067-Textbox-Image-Mod.docx
