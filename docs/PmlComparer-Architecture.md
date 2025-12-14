# PmlComparer Architecture Design

## Executive Summary

PmlComparer is a PowerPoint presentation comparison tool for Open-Xml-PowerTools, following the established patterns of SmlComparer (Excel) and WmlComparer (Word). Unlike Word documents, **PowerPoint has no native tracked changes markup** (no `w:ins`/`w:del` equivalents). Microsoft's Compare/Merge feature is being retired in Microsoft 365, making a standards-based comparison tool valuable.

This design adopts a **semantic model approach** with **signature-based comparison** (similar to SmlComparer) rather than raw XML diffing, which would be too noisy due to ID churn, relationship IDs, and reserialization differences.

---

## Table of Contents

1. [Design Principles](#1-design-principles)
2. [Architecture Overview](#2-architecture-overview)
3. [Semantic Model (Canonicalization)](#3-semantic-model-canonicalization)
4. [Matching Engine](#4-matching-engine)
5. [Diff Engine](#5-diff-engine)
6. [Renderers](#6-renderers)
7. [Public API](#7-public-api)
8. [Implementation Phases](#8-implementation-phases)
9. [Test Strategy](#9-test-strategy)

---

## 1. Design Principles

### 1.1 Follow Established Patterns

- **SmlComparer pattern**: Signature-based canonicalization, structured change results, optional marked output
- **Separation of concerns**: Canonicalizer → Matcher → DiffEngine → Renderer
- **Consistent API style**: Static `Compare()` method, settings class, result class with statistics

### 1.2 Semantic Model Over Raw XML

Raw XML diffs are unsuitable because PPTX contains significant non-semantic churn:
- Shape IDs and relationship IDs change across saves
- `extLst` extensions vary by PowerPoint version
- Reserialization produces different but semantically identical XML
- Z-order indices may be renumbered

Instead, build a **canonical semantic model** that captures what humans care about.

### 1.3 No Native Tracked Changes

PowerPoint lacks Word-style inline revision markup. Microsoft's `revisionInfo` part ([MS-PPTX §2.1.27](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/ca43e072-32cf-47db-99e2-f751fa624118)) contains collaborative session metadata, not granular change tracking.

**Implication**: We define our own change representation and rendering strategies.

### 1.4 Multiple Output Modes

Support various use cases:
1. **Structured comparison result** (JSON-serializable, for programmatic use)
2. **Markup presentation** (visual diff overlays in a standard PPTX)
3. **Merge engine** (apply changes from B onto A)

---

## 2. Architecture Overview

```
┌─────────────────────────────────────────────────────────────────────────┐
│                            PmlComparer                                   │
│  ┌─────────────┐    ┌──────────────┐    ┌─────────────┐                │
│  │ PmlDocument │───▶│ Canonicalizer│───▶│ Presentation│                │
│  │   (older)   │    │              │    │  Signature  │                │
│  └─────────────┘    └──────────────┘    └──────┬──────┘                │
│                                                 │                       │
│  ┌─────────────┐    ┌──────────────┐    ┌──────▼──────┐    ┌─────────┐ │
│  │ PmlDocument │───▶│ Canonicalizer│───▶│ Presentation│───▶│ Match   │ │
│  │   (newer)   │    │              │    │  Signature  │    │ Engine  │ │
│  └─────────────┘    └──────────────┘    └─────────────┘    └────┬────┘ │
│                                                                  │      │
│                     ┌──────────────┐    ┌─────────────┐         │      │
│                     │    Diff      │◀───│   Matched   │◀────────┘      │
│                     │   Engine     │    │   Slides    │                │
│                     └──────┬───────┘    └─────────────┘                │
│                            │                                            │
│                     ┌──────▼───────┐                                   │
│                     │  Comparison  │                                   │
│                     │    Result    │                                   │
│                     └──────┬───────┘                                   │
│                            │                                            │
│          ┌─────────────────┼─────────────────┐                         │
│          ▼                 ▼                 ▼                         │
│   ┌─────────────┐   ┌─────────────┐   ┌─────────────┐                 │
│   │   Markup    │   │   Two-Up    │   │    Merge    │                 │
│   │  Renderer   │   │  Renderer   │   │   Applier   │                 │
│   └─────────────┘   └─────────────┘   └─────────────┘                 │
└─────────────────────────────────────────────────────────────────────────┘
```

### 2.1 Component Responsibilities

| Component | Responsibility |
|-----------|----------------|
| `PmlCanonicalizer` | Convert PresentationDocument to semantic signatures |
| `PmlMatchEngine` | Match slides between presentations, shapes within slides |
| `PmlDiffEngine` | Compute changes between matched elements |
| `PmlMarkupRenderer` | Produce PPTX with visual change overlays |
| `PmlTwoUpRenderer` | Produce side-by-side before/after slides |
| `PmlMergeApplier` | Apply changes to produce merged presentation |

---

## 3. Semantic Model (Canonicalization)

### 3.1 Signature Hierarchy

```
PresentationSignature
├── SlideSize (cx, cy)
├── List<SlideSignature> Slides (ordered by p:sldIdLst)
├── List<SlideMasterSignature> Masters
├── List<SlideLayoutSignature> Layouts
├── Dictionary<string, string> CustomProperties
└── string ThemeHash

SlideSignature
├── int Index (1-based position in deck)
├── string RelationshipId
├── string LayoutRelationshipId
├── SlideBackgroundSignature Background
├── List<ShapeSignature> Shapes (ordered by z-order)
├── string NotesText (from notes slide, if present)
├── TransitionSignature Transition
├── string ContentHash (for similarity matching)
└── string TitleText (extracted from title placeholder)

ShapeSignature
├── string Name (from cNvPr/@name - primary identity)
├── uint Id (from cNvPr/@id - secondary, may change)
├── ShapeType Type (TextBox, Picture, Table, Chart, Group, SmartArt, Connector, etc.)
├── PlaceholderInfo Placeholder (type + idx, if placeholder)
├── TransformSignature Transform (x, y, cx, cy, rot, flipH, flipV)
├── int ZOrder (position in spTree)
├── string GeometryHash (prstGeom or custGeom signature)
├── ShapeStyleSignature Style (fill, line, effects)
├── TextBodySignature TextBody (if has text)
├── string ImageHash (if picture - hash of media binary)
├── TableSignature Table (if table)
├── string ChartHash (if chart - hash of chart XML + embedded workbook)
└── List<ShapeSignature> Children (if group)

TextBodySignature
├── List<ParagraphSignature> Paragraphs
├── BodyPropertiesSignature BodyProperties
└── string PlainText (concatenated, for quick comparison)

ParagraphSignature
├── List<RunSignature> Runs
├── ParagraphPropertiesSignature Properties (alignment, bullet, spacing)
└── string PlainText

RunSignature
├── string Text
├── RunPropertiesSignature Properties (bold, italic, underline, color, size, font)
└── string ContentHash
```

### 3.2 Shape Identity Strategy

Shape matching is the hard problem. We use a **multi-key identity strategy**:

```csharp
internal class ShapeIdentity
{
    // Primary keys (stable across edits when present)
    public string Name { get; set; }           // cNvPr/@name (user-assigned or auto-generated)
    public PlaceholderInfo Placeholder { get; set; } // Placeholder type + index

    // Secondary keys (for fuzzy matching)
    public ShapeType Type { get; set; }
    public TransformSignature Transform { get; set; }
    public string ContentHash { get; set; }    // Text or image hash
    public string GeometryHash { get; set; }
}
```

**Identity resolution priority**:
1. Exact placeholder match (type + idx) → definite match
2. Same name + same type → high confidence match
3. Same name (different type) → possible match (shape converted)
4. Fuzzy match (type + position + content similarity) → candidate match with score

### 3.3 Transform Normalization

```csharp
internal class TransformSignature
{
    public long X { get; set; }      // EMUs from a:off/@x
    public long Y { get; set; }      // EMUs from a:off/@y
    public long Cx { get; set; }     // EMUs from a:ext/@cx
    public long Cy { get; set; }     // EMUs from a:ext/@cy
    public int Rotation { get; set; } // 60000ths of a degree from @rot
    public bool FlipH { get; set; }
    public bool FlipV { get; set; }

    // Computed for comparison
    public (long X, long Y, long Cx, long Cy) BoundingBox =>
        ComputeRotatedBoundingBox();

    // Position tolerance for "same location" (e.g., 1% of slide width)
    public bool IsNear(TransformSignature other, long tolerance) { ... }
}
```

### 3.4 Text Extraction

Extract text respecting paragraph/run structure for granular diffing:

```csharp
internal static class PmlTextExtractor
{
    public static TextBodySignature ExtractTextBody(XElement txBody)
    {
        var sig = new TextBodySignature();

        foreach (var p in txBody.Elements(A.p))
        {
            var para = new ParagraphSignature();

            foreach (var r in p.Elements(A.r))
            {
                var run = new RunSignature
                {
                    Text = r.Element(A.t)?.Value ?? "",
                    Properties = ExtractRunProperties(r.Element(A.rPr))
                };
                para.Runs.Add(run);
            }

            // Handle field codes (a:fld) and line breaks (a:br)
            para.Properties = ExtractParagraphProperties(p.Element(A.pPr));
            para.PlainText = string.Join("", para.Runs.Select(r => r.Text));
            sig.Paragraphs.Add(para);
        }

        sig.PlainText = string.Join("\n", sig.Paragraphs.Select(p => p.PlainText));
        return sig;
    }
}
```

### 3.5 Media Hashing

```csharp
internal static class PmlMediaHasher
{
    public static string HashImagePart(ImagePart imagePart)
    {
        using var stream = imagePart.GetStream();
        var bytes = SHA256.HashData(stream);
        return Convert.ToBase64String(bytes);
    }

    public static string HashChartPart(ChartPart chartPart)
    {
        // Hash chart XML + embedded spreadsheet (if present)
        var sb = new StringBuilder();
        sb.Append(chartPart.GetXDocument().ToString(SaveOptions.DisableFormatting));

        if (chartPart.EmbeddedPackagePart != null)
        {
            using var stream = chartPart.EmbeddedPackagePart.GetStream();
            var bytes = SHA256.HashData(stream);
            sb.Append(Convert.ToBase64String(bytes));
        }

        return PtUtils.SHA256HashStringForUTF8String(sb.ToString());
    }
}
```

---

## 4. Matching Engine

### 4.1 Slide Matching

```csharp
internal static class PmlSlideMatchEngine
{
    public static List<SlideMatch> MatchSlides(
        PresentationSignature sig1,
        PresentationSignature sig2,
        PmlComparerSettings settings)
    {
        var matches = new List<SlideMatch>();
        var used1 = new HashSet<int>();
        var used2 = new HashSet<int>();

        // Pass 1: Match by relationship ID (if stable)
        if (settings.MatchByRelationshipId)
        {
            MatchByRelationshipId(sig1, sig2, matches, used1, used2);
        }

        // Pass 2: Match by title text (exact match)
        MatchByTitleText(sig1, sig2, matches, used1, used2);

        // Pass 3: Match by content fingerprint (for reordered slides)
        MatchByContentFingerprint(sig1, sig2, matches, used1, used2, settings);

        // Pass 4: Use LCS on remaining slides for optimal alignment
        if (settings.UseSlideAlignmentLCS)
        {
            AlignRemainingWithLCS(sig1, sig2, matches, used1, used2, settings);
        }

        // Remaining unmatched = inserted/deleted
        AddUnmatchedAsInsertedDeleted(sig1, sig2, matches, used1, used2);

        return matches;
    }
}

internal class SlideMatch
{
    public SlideMatchType MatchType { get; set; }
    public int? OldIndex { get; set; }
    public int? NewIndex { get; set; }
    public SlideSignature OldSlide { get; set; }
    public SlideSignature NewSlide { get; set; }
    public double Similarity { get; set; }
    public bool WasMoved => OldIndex != NewIndex && MatchType == SlideMatchType.Matched;
}

internal enum SlideMatchType
{
    Matched,    // Same slide in both presentations
    Inserted,   // New slide in newer presentation
    Deleted,    // Slide removed from older presentation
}
```

### 4.2 Slide Fingerprinting

```csharp
internal class SlideFingerprint
{
    public string TitleText { get; set; }
    public string BodyTextHash { get; set; }
    public int ShapeCount { get; set; }
    public int PictureCount { get; set; }
    public int TableCount { get; set; }
    public int ChartCount { get; set; }
    public string LayoutType { get; set; }

    public double ComputeSimilarity(SlideFingerprint other)
    {
        double score = 0;
        double maxScore = 0;

        // Title match (high weight)
        maxScore += 3;
        if (TitleText == other.TitleText && !string.IsNullOrEmpty(TitleText))
            score += 3;
        else if (LevenshteinSimilarity(TitleText, other.TitleText) > 0.8)
            score += 2;

        // Body text similarity
        maxScore += 2;
        if (BodyTextHash == other.BodyTextHash)
            score += 2;

        // Shape counts (lower weight)
        maxScore += 1;
        if (ShapeCount == other.ShapeCount &&
            PictureCount == other.PictureCount &&
            TableCount == other.TableCount)
            score += 1;

        // Layout match
        maxScore += 1;
        if (LayoutType == other.LayoutType)
            score += 1;

        return score / maxScore;
    }
}
```

### 4.3 Shape Matching Within Slides

```csharp
internal static class PmlShapeMatchEngine
{
    public static List<ShapeMatch> MatchShapes(
        SlideSignature slide1,
        SlideSignature slide2,
        PmlComparerSettings settings)
    {
        var matches = new List<ShapeMatch>();
        var used1 = new HashSet<string>();
        var used2 = new HashSet<string>();

        // Pass 1: Exact placeholder match (type + idx)
        MatchByPlaceholder(slide1, slide2, matches, used1, used2);

        // Pass 2: Exact name match + same type
        MatchByNameAndType(slide1, slide2, matches, used1, used2);

        // Pass 3: Same name (possibly converted type)
        MatchByNameOnly(slide1, slide2, matches, used1, used2);

        // Pass 4: Fuzzy match (type + position + content)
        if (settings.EnableFuzzyShapeMatching)
        {
            FuzzyMatch(slide1, slide2, matches, used1, used2, settings);
        }

        // Remaining = inserted/deleted
        AddUnmatchedAsInsertedDeleted(slide1, slide2, matches, used1, used2);

        return matches;
    }

    private static double ComputeShapeMatchScore(
        ShapeSignature s1,
        ShapeSignature s2,
        PmlComparerSettings settings)
    {
        double score = 0;

        // Same type (required for match)
        if (s1.Type != s2.Type) return 0;
        score += 0.2;

        // Position similarity
        if (s1.Transform.IsNear(s2.Transform, settings.PositionTolerance))
            score += 0.3;
        else if (s1.Transform.BoundingBox.Intersects(s2.Transform.BoundingBox))
            score += 0.1;

        // Content similarity
        if (s1.Type == ShapeType.Picture)
        {
            if (s1.ImageHash == s2.ImageHash)
                score += 0.5;
        }
        else if (s1.TextBody != null && s2.TextBody != null)
        {
            var textSim = ComputeTextSimilarity(s1.TextBody.PlainText, s2.TextBody.PlainText);
            score += textSim * 0.5;
        }

        // Geometry similarity
        if (s1.GeometryHash == s2.GeometryHash)
            score += 0.1;

        return score;
    }
}

internal class ShapeMatch
{
    public ShapeMatchType MatchType { get; set; }
    public ShapeSignature OldShape { get; set; }
    public ShapeSignature NewShape { get; set; }
    public double Score { get; set; }
    public ShapeMatchMethod Method { get; set; }
}

internal enum ShapeMatchType
{
    Matched,
    Inserted,
    Deleted,
}

internal enum ShapeMatchMethod
{
    Placeholder,
    NameAndType,
    NameOnly,
    Fuzzy,
}
```

---

## 5. Diff Engine

### 5.1 Change Types

```csharp
public enum PmlChangeType
{
    // Presentation-level
    SlideSizeChanged,
    ThemeChanged,

    // Slide-level structure
    SlideInserted,
    SlideDeleted,
    SlideMoved,
    SlideLayoutChanged,
    SlideBackgroundChanged,
    SlideTransitionChanged,
    SlideNotesChanged,

    // Shape-level structure
    ShapeInserted,
    ShapeDeleted,
    ShapeMoved,
    ShapeResized,
    ShapeRotated,
    ShapeZOrderChanged,
    ShapeTypeChanged,        // e.g., rectangle → oval

    // Shape content
    TextChanged,
    TextFormattingChanged,
    ImageReplaced,
    TableContentChanged,
    TableStructureChanged,   // rows/cols added/removed
    ChartDataChanged,
    ChartFormatChanged,

    // Shape style
    ShapeFillChanged,
    ShapeLineChanged,
    ShapeEffectsChanged,

    // Group-specific
    GroupMembershipChanged,  // shapes added/removed from group
}
```

### 5.2 Change Model

```csharp
public class PmlChange
{
    public PmlChangeType ChangeType { get; set; }

    // Location
    public int? SlideIndex { get; set; }        // 1-based
    public int? OldSlideIndex { get; set; }     // For moved slides
    public string ShapeName { get; set; }
    public string ShapeId { get; set; }

    // Details (type-specific)
    public string OldValue { get; set; }
    public string NewValue { get; set; }
    public TransformSignature OldTransform { get; set; }
    public TransformSignature NewTransform { get; set; }

    // Text changes (for granular text diffs)
    public List<TextChange> TextChanges { get; set; }

    // Metadata
    public ShapeMatchMethod MatchMethod { get; set; }
    public double MatchConfidence { get; set; }

    public string GetDescription()
    {
        return ChangeType switch
        {
            PmlChangeType.SlideInserted => $"Slide {SlideIndex} inserted",
            PmlChangeType.SlideDeleted => $"Slide {OldSlideIndex} deleted",
            PmlChangeType.SlideMoved => $"Slide moved from position {OldSlideIndex} to {SlideIndex}",
            PmlChangeType.ShapeInserted => $"Shape '{ShapeName}' inserted on slide {SlideIndex}",
            PmlChangeType.ShapeDeleted => $"Shape '{ShapeName}' deleted from slide {SlideIndex}",
            PmlChangeType.ShapeMoved => $"Shape '{ShapeName}' moved on slide {SlideIndex}",
            PmlChangeType.ShapeResized => $"Shape '{ShapeName}' resized on slide {SlideIndex}",
            PmlChangeType.TextChanged => $"Text changed in '{ShapeName}' on slide {SlideIndex}",
            PmlChangeType.ImageReplaced => $"Image replaced in '{ShapeName}' on slide {SlideIndex}",
            _ => $"{ChangeType} on slide {SlideIndex}"
        };
    }
}

public class TextChange
{
    public TextChangeType Type { get; set; }  // Insert, Delete, Replace, FormatOnly
    public int ParagraphIndex { get; set; }
    public int RunIndex { get; set; }
    public string OldText { get; set; }
    public string NewText { get; set; }
    public RunPropertiesSignature OldFormat { get; set; }
    public RunPropertiesSignature NewFormat { get; set; }
}
```

### 5.3 Comparison Result

```csharp
public class PmlComparisonResult
{
    public List<PmlChange> Changes { get; } = new List<PmlChange>();

    // Statistics
    public int TotalChanges => Changes.Count;
    public int SlidesInserted => Changes.Count(c => c.ChangeType == PmlChangeType.SlideInserted);
    public int SlidesDeleted => Changes.Count(c => c.ChangeType == PmlChangeType.SlideDeleted);
    public int SlidesMoved => Changes.Count(c => c.ChangeType == PmlChangeType.SlideMoved);
    public int ShapesInserted => Changes.Count(c => c.ChangeType == PmlChangeType.ShapeInserted);
    public int ShapesDeleted => Changes.Count(c => c.ChangeType == PmlChangeType.ShapeDeleted);
    public int ShapesMoved => Changes.Count(c => c.ChangeType == PmlChangeType.ShapeMoved);
    public int TextChanges => Changes.Count(c => c.ChangeType == PmlChangeType.TextChanged);
    public int ImagesReplaced => Changes.Count(c => c.ChangeType == PmlChangeType.ImageReplaced);

    // Queries
    public IEnumerable<PmlChange> GetChangesBySlide(int slideIndex)
        => Changes.Where(c => c.SlideIndex == slideIndex);

    public IEnumerable<PmlChange> GetChangesByType(PmlChangeType type)
        => Changes.Where(c => c.ChangeType == type);

    public IEnumerable<PmlChange> GetChangesByShape(string shapeName)
        => Changes.Where(c => c.ShapeName == shapeName);

    // Serialization
    public string ToJson()
    {
        var options = new JsonSerializerOptions { WriteIndented = true };
        return JsonSerializer.Serialize(new
        {
            Summary = new
            {
                TotalChanges,
                SlidesInserted,
                SlidesDeleted,
                SlidesMoved,
                ShapesInserted,
                ShapesDeleted,
                ShapesMoved,
                TextChanges,
                ImagesReplaced
            },
            Changes = Changes.Select(c => new
            {
                c.ChangeType,
                c.SlideIndex,
                c.OldSlideIndex,
                c.ShapeName,
                Description = c.GetDescription()
            })
        }, options);
    }
}
```

### 5.4 Diff Engine Implementation

```csharp
internal static class PmlDiffEngine
{
    public static PmlComparisonResult ComputeDiff(
        PresentationSignature sig1,
        PresentationSignature sig2,
        List<SlideMatch> slideMatches,
        PmlComparerSettings settings)
    {
        var result = new PmlComparisonResult();

        // Presentation-level changes
        if (settings.ComparePresentationProperties)
        {
            ComparePresentationLevel(sig1, sig2, result);
        }

        // Process slide matches
        foreach (var slideMatch in slideMatches)
        {
            switch (slideMatch.MatchType)
            {
                case SlideMatchType.Inserted:
                    result.Changes.Add(new PmlChange
                    {
                        ChangeType = PmlChangeType.SlideInserted,
                        SlideIndex = slideMatch.NewIndex
                    });
                    break;

                case SlideMatchType.Deleted:
                    result.Changes.Add(new PmlChange
                    {
                        ChangeType = PmlChangeType.SlideDeleted,
                        OldSlideIndex = slideMatch.OldIndex
                    });
                    break;

                case SlideMatchType.Matched:
                    // Check if moved
                    if (slideMatch.WasMoved)
                    {
                        result.Changes.Add(new PmlChange
                        {
                            ChangeType = PmlChangeType.SlideMoved,
                            SlideIndex = slideMatch.NewIndex,
                            OldSlideIndex = slideMatch.OldIndex
                        });
                    }

                    // Compare slide contents
                    CompareSlideContents(
                        slideMatch.OldSlide,
                        slideMatch.NewSlide,
                        slideMatch.NewIndex.Value,
                        settings,
                        result);
                    break;
            }
        }

        return result;
    }

    private static void CompareSlideContents(
        SlideSignature slide1,
        SlideSignature slide2,
        int slideIndex,
        PmlComparerSettings settings,
        PmlComparisonResult result)
    {
        // Compare slide-level properties
        if (settings.CompareSlideProperties)
        {
            if (slide1.LayoutRelationshipId != slide2.LayoutRelationshipId)
            {
                result.Changes.Add(new PmlChange
                {
                    ChangeType = PmlChangeType.SlideLayoutChanged,
                    SlideIndex = slideIndex
                });
            }

            // Background, transition, notes...
        }

        // Match and compare shapes
        var shapeMatches = PmlShapeMatchEngine.MatchShapes(slide1, slide2, settings);

        foreach (var shapeMatch in shapeMatches)
        {
            switch (shapeMatch.MatchType)
            {
                case ShapeMatchType.Inserted:
                    result.Changes.Add(new PmlChange
                    {
                        ChangeType = PmlChangeType.ShapeInserted,
                        SlideIndex = slideIndex,
                        ShapeName = shapeMatch.NewShape.Name,
                        MatchConfidence = shapeMatch.Score
                    });
                    break;

                case ShapeMatchType.Deleted:
                    result.Changes.Add(new PmlChange
                    {
                        ChangeType = PmlChangeType.ShapeDeleted,
                        SlideIndex = slideIndex,
                        ShapeName = shapeMatch.OldShape.Name,
                        MatchConfidence = shapeMatch.Score
                    });
                    break;

                case ShapeMatchType.Matched:
                    CompareMatchedShapes(
                        shapeMatch.OldShape,
                        shapeMatch.NewShape,
                        slideIndex,
                        shapeMatch,
                        settings,
                        result);
                    break;
            }
        }
    }

    private static void CompareMatchedShapes(
        ShapeSignature shape1,
        ShapeSignature shape2,
        int slideIndex,
        ShapeMatch match,
        PmlComparerSettings settings,
        PmlComparisonResult result)
    {
        // Transform changes
        if (settings.CompareShapeTransforms)
        {
            var t1 = shape1.Transform;
            var t2 = shape2.Transform;

            if (t1.X != t2.X || t1.Y != t2.Y)
            {
                result.Changes.Add(new PmlChange
                {
                    ChangeType = PmlChangeType.ShapeMoved,
                    SlideIndex = slideIndex,
                    ShapeName = shape2.Name,
                    OldTransform = t1,
                    NewTransform = t2,
                    MatchMethod = match.Method,
                    MatchConfidence = match.Score
                });
            }

            if (t1.Cx != t2.Cx || t1.Cy != t2.Cy)
            {
                result.Changes.Add(new PmlChange
                {
                    ChangeType = PmlChangeType.ShapeResized,
                    SlideIndex = slideIndex,
                    ShapeName = shape2.Name,
                    OldTransform = t1,
                    NewTransform = t2
                });
            }

            if (t1.Rotation != t2.Rotation)
            {
                result.Changes.Add(new PmlChange
                {
                    ChangeType = PmlChangeType.ShapeRotated,
                    SlideIndex = slideIndex,
                    ShapeName = shape2.Name,
                    OldValue = t1.Rotation.ToString(),
                    NewValue = t2.Rotation.ToString()
                });
            }
        }

        // Z-order changes
        if (shape1.ZOrder != shape2.ZOrder)
        {
            result.Changes.Add(new PmlChange
            {
                ChangeType = PmlChangeType.ShapeZOrderChanged,
                SlideIndex = slideIndex,
                ShapeName = shape2.Name,
                OldValue = shape1.ZOrder.ToString(),
                NewValue = shape2.ZOrder.ToString()
            });
        }

        // Content changes based on type
        switch (shape1.Type)
        {
            case ShapeType.TextBox:
            case ShapeType.AutoShape:
                if (shape1.TextBody != null && shape2.TextBody != null)
                {
                    CompareTextContent(shape1, shape2, slideIndex, settings, result);
                }
                break;

            case ShapeType.Picture:
                if (shape1.ImageHash != shape2.ImageHash)
                {
                    result.Changes.Add(new PmlChange
                    {
                        ChangeType = PmlChangeType.ImageReplaced,
                        SlideIndex = slideIndex,
                        ShapeName = shape2.Name
                    });
                }
                break;

            case ShapeType.Table:
                CompareTableContent(shape1, shape2, slideIndex, settings, result);
                break;

            case ShapeType.Chart:
                if (shape1.ChartHash != shape2.ChartHash)
                {
                    result.Changes.Add(new PmlChange
                    {
                        ChangeType = PmlChangeType.ChartDataChanged,
                        SlideIndex = slideIndex,
                        ShapeName = shape2.Name
                    });
                }
                break;
        }

        // Style changes
        if (settings.CompareShapeStyles)
        {
            CompareShapeStyles(shape1, shape2, slideIndex, result);
        }
    }

    private static void CompareTextContent(
        ShapeSignature shape1,
        ShapeSignature shape2,
        int slideIndex,
        PmlComparerSettings settings,
        PmlComparisonResult result)
    {
        var text1 = shape1.TextBody;
        var text2 = shape2.TextBody;

        // Quick check: if plain text is identical, check formatting only
        if (text1.PlainText == text2.PlainText)
        {
            if (settings.CompareTextFormatting)
            {
                // Compare paragraph-by-paragraph formatting
                CompareTextFormatting(shape1, shape2, slideIndex, result);
            }
            return;
        }

        // Text content differs
        var textChanges = ComputeTextDiff(text1, text2, settings);

        result.Changes.Add(new PmlChange
        {
            ChangeType = PmlChangeType.TextChanged,
            SlideIndex = slideIndex,
            ShapeName = shape2.Name,
            OldValue = text1.PlainText,
            NewValue = text2.PlainText,
            TextChanges = textChanges
        });
    }
}
```

---

## 6. Renderers

### 6.1 Markup Renderer (Primary)

Produces a standard PPTX with visual change overlays that opens in any PowerPoint version.

```csharp
public static class PmlMarkupRenderer
{
    public static PmlDocument RenderMarkedPresentation(
        PmlDocument newerDoc,
        PmlComparisonResult result,
        PmlComparerSettings settings)
    {
        using var ms = new MemoryStream();
        ms.Write(newerDoc.DocumentByteArray, 0, newerDoc.DocumentByteArray.Length);

        using var pDoc = PresentationDocument.Open(ms, true);

        foreach (var slideChange in result.Changes.GroupBy(c => c.SlideIndex))
        {
            if (slideChange.Key == null) continue;

            var slidePart = GetSlidePart(pDoc, slideChange.Key.Value);
            if (slidePart == null) continue;

            var slideXDoc = slidePart.GetXDocument();
            var spTree = slideXDoc.Root.Element(P.cSld)?.Element(P.spTree);
            if (spTree == null) continue;

            foreach (var change in slideChange)
            {
                AddChangeOverlay(spTree, change, settings);
            }

            // Add change summary to speaker notes
            AddChangeSummaryToNotes(slidePart, slideChange, settings);

            slidePart.PutXDocument();
        }

        // Add summary slide at the end
        if (settings.AddSummarySlide)
        {
            AddSummarySlide(pDoc, result, settings);
        }

        pDoc.Save();
        return new PmlDocument(newerDoc.FileName, ms.ToArray());
    }

    private static void AddChangeOverlay(
        XElement spTree,
        PmlChange change,
        PmlComparerSettings settings)
    {
        switch (change.ChangeType)
        {
            case PmlChangeType.ShapeInserted:
                // Add green bounding box around new shape
                AddBoundingBox(spTree, change, settings.InsertedColor, "New");
                break;

            case PmlChangeType.ShapeDeleted:
                // Add red "ghost" indicator where shape was
                AddDeletedIndicator(spTree, change, settings.DeletedColor);
                break;

            case PmlChangeType.ShapeMoved:
                // Add arrow from old position to new
                AddMoveIndicator(spTree, change, settings.MovedColor);
                break;

            case PmlChangeType.TextChanged:
                // Add callout with before/after text
                AddTextChangeCallout(spTree, change, settings);
                break;

            case PmlChangeType.ImageReplaced:
                // Add "Image Replaced" label
                AddChangeLabel(spTree, change, "Image Replaced", settings.ModifiedColor);
                break;
        }
    }

    private static void AddBoundingBox(
        XElement spTree,
        PmlChange change,
        string color,
        string label)
    {
        // Create a rectangle shape with no fill, colored stroke
        var boundingBox = new XElement(P.sp,
            new XElement(P.nvSpPr,
                new XElement(P.cNvPr,
                    new XAttribute("id", GetNextShapeId(spTree)),
                    new XAttribute("name", $"Change: {label} - {change.ShapeName}")),
                new XElement(P.cNvSpPr),
                new XElement(P.nvPr)),
            new XElement(P.spPr,
                new XElement(A.xfrm,
                    new XElement(A.off,
                        new XAttribute("x", change.NewTransform?.X ?? 0),
                        new XAttribute("y", change.NewTransform?.Y ?? 0)),
                    new XElement(A.ext,
                        new XAttribute("cx", change.NewTransform?.Cx ?? 914400),
                        new XAttribute("cy", change.NewTransform?.Cy ?? 914400))),
                new XElement(A.prstGeom,
                    new XAttribute("prst", "rect"),
                    new XElement(A.avLst)),
                new XElement(A.noFill),
                new XElement(A.ln,
                    new XAttribute("w", "38100"), // 3pt
                    new XElement(A.solidFill,
                        new XElement(A.srgbClr,
                            new XAttribute("val", color))),
                    new XElement(A.prstDash,
                        new XAttribute("val", "dash")))));

        spTree.Add(boundingBox);

        // Add label
        AddLabel(spTree, change.NewTransform, label, color);
    }
}
```

### 6.2 Two-Up Renderer

Creates side-by-side comparison slides:

```csharp
public static class PmlTwoUpRenderer
{
    public static PmlDocument RenderTwoUpPresentation(
        PmlDocument olderDoc,
        PmlDocument newerDoc,
        PmlComparisonResult result,
        PmlComparerSettings settings)
    {
        // For each changed slide, create a new slide showing:
        // Left side: scaled "before" version
        // Right side: scaled "after" version
        // Connectors between matched shapes
        // Change annotations

        // Implementation follows PresentationBuilder patterns
        // ...
    }
}
```

### 6.3 Merge Applier

Applies changes from source B onto source A:

```csharp
public static class PmlMergeApplier
{
    public static PmlDocument ApplyChanges(
        PmlDocument baseDoc,
        PmlDocument changedDoc,
        PmlComparisonResult result,
        PmlMergeSettings mergeSettings)
    {
        // Useful for:
        // - "Accept all changes" workflow
        // - Selective merge of specific changes
        // - Three-way merge (future)

        using var ms = new MemoryStream();
        ms.Write(baseDoc.DocumentByteArray, 0, baseDoc.DocumentByteArray.Length);

        using var pDoc = PresentationDocument.Open(ms, true);

        foreach (var change in result.Changes)
        {
            if (mergeSettings.ShouldApply(change))
            {
                ApplyChange(pDoc, change, changedDoc);
            }
        }

        pDoc.Save();
        return new PmlDocument(baseDoc.FileName, ms.ToArray());
    }
}
```

---

## 7. Public API

### 7.1 Settings

```csharp
public class PmlComparerSettings
{
    // === Comparison Scope ===

    /// <summary>Compare slide structure (added/deleted/reordered slides).</summary>
    public bool CompareSlideStructure { get; set; } = true;

    /// <summary>Compare shape structure within slides.</summary>
    public bool CompareShapeStructure { get; set; } = true;

    /// <summary>Compare text content within shapes.</summary>
    public bool CompareTextContent { get; set; } = true;

    /// <summary>Compare text formatting (bold, italic, color, etc.).</summary>
    public bool CompareTextFormatting { get; set; } = true;

    /// <summary>Compare shape transforms (position, size, rotation).</summary>
    public bool CompareShapeTransforms { get; set; } = true;

    /// <summary>Compare shape styles (fill, line, effects).</summary>
    public bool CompareShapeStyles { get; set; } = false; // Off by default, can be noisy

    /// <summary>Compare images by content hash.</summary>
    public bool CompareImageContent { get; set; } = true;

    /// <summary>Compare chart data and formatting.</summary>
    public bool CompareCharts { get; set; } = true;

    /// <summary>Compare tables.</summary>
    public bool CompareTables { get; set; } = true;

    /// <summary>Compare slide notes.</summary>
    public bool CompareNotes { get; set; } = false;

    /// <summary>Compare slide transitions.</summary>
    public bool CompareTransitions { get; set; } = false;

    /// <summary>Compare slide masters and layouts.</summary>
    public bool CompareMasters { get; set; } = false;

    // === Matching Settings ===

    /// <summary>Use relationship IDs for initial slide matching.</summary>
    public bool MatchByRelationshipId { get; set; } = true;

    /// <summary>Enable fuzzy shape matching when exact matches fail.</summary>
    public bool EnableFuzzyShapeMatching { get; set; } = true;

    /// <summary>Minimum similarity score (0.0-1.0) for fuzzy slide matching.</summary>
    public double SlideSimilarityThreshold { get; set; } = 0.6;

    /// <summary>Minimum similarity score (0.0-1.0) for fuzzy shape matching.</summary>
    public double ShapeSimilarityThreshold { get; set; } = 0.7;

    /// <summary>Position tolerance in EMUs for "same location" matching.</summary>
    public long PositionTolerance { get; set; } = 91440; // ~0.1 inch

    /// <summary>Use LCS algorithm for slide alignment.</summary>
    public bool UseSlideAlignmentLCS { get; set; } = true;

    // === Output Settings ===

    /// <summary>Author name for change annotations.</summary>
    public string AuthorForChanges { get; set; } = "Open-Xml-PowerTools";

    /// <summary>Add a summary slide at the end of marked presentations.</summary>
    public bool AddSummarySlide { get; set; } = true;

    /// <summary>Add change summary to speaker notes.</summary>
    public bool AddNotesAnnotations { get; set; } = true;

    // === Colors (RRGGBB hex) ===

    public string InsertedColor { get; set; } = "00AA00";  // Green
    public string DeletedColor { get; set; } = "FF0000";   // Red
    public string ModifiedColor { get; set; } = "FFA500";  // Orange
    public string MovedColor { get; set; } = "0000FF";     // Blue
    public string FormattingColor { get; set; } = "9932CC"; // Purple

    // === Logging ===

    /// <summary>Optional callback for logging/debugging.</summary>
    public Action<string> LogCallback { get; set; }
}
```

### 7.2 Main API

```csharp
/// <summary>
/// Compares PowerPoint presentations and produces structured change results
/// and/or marked presentations showing differences.
/// </summary>
public static class PmlComparer
{
    /// <summary>
    /// Compare two presentations and return a structured list of changes.
    /// </summary>
    /// <param name="older">The original/older presentation.</param>
    /// <param name="newer">The revised/newer presentation.</param>
    /// <param name="settings">Comparison settings.</param>
    /// <returns>A result object containing all detected changes.</returns>
    public static PmlComparisonResult Compare(
        PmlDocument older,
        PmlDocument newer,
        PmlComparerSettings settings = null)
    {
        if (older == null) throw new ArgumentNullException(nameof(older));
        if (newer == null) throw new ArgumentNullException(nameof(newer));
        settings ??= new PmlComparerSettings();

        Log(settings, "PmlComparer.Compare: Starting comparison");

        // 1. Canonicalize both presentations
        var sig1 = PmlCanonicalizer.Canonicalize(older, settings);
        var sig2 = PmlCanonicalizer.Canonicalize(newer, settings);

        Log(settings, $"Canonicalized older: {sig1.Slides.Count} slides");
        Log(settings, $"Canonicalized newer: {sig2.Slides.Count} slides");

        // 2. Match slides
        var slideMatches = PmlSlideMatchEngine.MatchSlides(sig1, sig2, settings);

        Log(settings, $"Matched {slideMatches.Count(m => m.MatchType == SlideMatchType.Matched)} slides");

        // 3. Compute diff
        var result = PmlDiffEngine.ComputeDiff(sig1, sig2, slideMatches, settings);

        Log(settings, $"Found {result.TotalChanges} changes");

        return result;
    }

    /// <summary>
    /// Produce a marked presentation highlighting all differences.
    /// The output is based on the newer presentation with visual change overlays.
    /// </summary>
    public static PmlDocument ProduceMarkedPresentation(
        PmlDocument older,
        PmlDocument newer,
        PmlComparerSettings settings = null)
    {
        settings ??= new PmlComparerSettings();

        var result = Compare(older, newer, settings);
        return PmlMarkupRenderer.RenderMarkedPresentation(newer, result, settings);
    }

    /// <summary>
    /// Produce a two-up comparison presentation showing before/after side-by-side.
    /// </summary>
    public static PmlDocument ProduceTwoUpPresentation(
        PmlDocument older,
        PmlDocument newer,
        PmlComparerSettings settings = null)
    {
        settings ??= new PmlComparerSettings();

        var result = Compare(older, newer, settings);
        return PmlTwoUpRenderer.RenderTwoUpPresentation(older, newer, result, settings);
    }

    /// <summary>
    /// Get the internal canonical signature of a presentation (for advanced use).
    /// </summary>
    public static PresentationSignature Canonicalize(
        PmlDocument doc,
        PmlComparerSettings settings = null)
    {
        return PmlCanonicalizer.Canonicalize(doc, settings ?? new PmlComparerSettings());
    }

    private static void Log(PmlComparerSettings settings, string message)
    {
        settings?.LogCallback?.Invoke(message);
    }
}
```

---

## 8. Implementation Phases

### Phase 1: MVP (Core Comparison)

**Goal**: Basic slide and shape comparison with structured results.

- [ ] `PmlComparerSettings` class
- [ ] `PmlChangeType` enum (core types only)
- [ ] `PmlChange` and `PmlComparisonResult` classes
- [ ] `PresentationSignature`, `SlideSignature`, `ShapeSignature` (basic)
- [ ] `PmlCanonicalizer` - extract basic structure
  - Slide list with indices
  - Shape names, types, and transforms
  - Plain text extraction
- [ ] `PmlSlideMatchEngine` - basic matching
  - Match by position (same index)
  - Match by title text
  - Detect inserted/deleted slides
- [ ] `PmlShapeMatchEngine` - basic matching
  - Match by name
  - Match by placeholder
  - Detect inserted/deleted shapes
- [ ] `PmlDiffEngine` - basic comparison
  - Slide insert/delete
  - Shape insert/delete
  - Text content changes (plain text diff)
- [ ] Unit tests for basic scenarios
- [ ] `PmlComparer.Compare()` public API

**Deliverables**:
- Can compare two presentations and return structured change list
- JSON output for programmatic consumption

### Phase 2: Enhanced Matching & Text Comparison

**Goal**: Robust matching and granular text diffs.

- [ ] Enhanced slide matching
  - Content fingerprinting
  - LCS-based alignment for reordered slides
  - Slide move detection
- [ ] Enhanced shape matching
  - Fuzzy matching by position + type
  - Match scoring and confidence
- [ ] Transform comparison
  - Move detection
  - Resize detection
  - Rotation detection
- [ ] Granular text comparison
  - Paragraph-level diffs
  - Run-level diffs
  - Formatting-only changes
- [ ] Image comparison by hash
- [ ] Extended test coverage

### Phase 3: Visual Markup Renderer

**Goal**: Produce marked PPTX with visual change overlays.

- [ ] `PmlMarkupRenderer`
  - Bounding boxes for changed shapes
  - Change labels/callouts
  - Deleted shape indicators
  - Move arrows
- [ ] Speaker notes annotations
- [ ] Summary slide generation
- [ ] `PmlComparer.ProduceMarkedPresentation()` public API

### Phase 4: Advanced Content Types

**Goal**: Tables, charts, and complex content.

- [ ] Table comparison
  - Cell content changes
  - Row/column structure changes
- [ ] Chart comparison
  - Data changes (via embedded workbook)
  - Format changes
- [ ] Group shape handling
  - Group membership changes
  - Recursive shape comparison within groups
- [ ] SmartArt detection (high-level)
- [ ] Media comparison (video, audio)

### Phase 5: Two-Up Renderer & Merge

**Goal**: Alternative output modes.

- [ ] `PmlTwoUpRenderer` - side-by-side comparison slides
- [ ] `PmlMergeApplier` - apply changes programmatically
- [ ] Selective merge support

### Phase 6: Advanced Features

**Goal**: Masters, layouts, and optimizations.

- [ ] Slide master/layout comparison
- [ ] Theme comparison
- [ ] Animation/transition comparison
- [ ] Performance optimization for large presentations
- [ ] Custom XML part for storing diff metadata (for add-in integration)

---

## 9. Test Strategy

### 9.1 Test Categories

```csharp
public class PmlComparerTests
{
    // === Slide Structure Tests ===

    [Fact]
    public void Compare_IdenticalPresentations_NoChanges() { }

    [Fact]
    public void Compare_SlideAdded_DetectsInsertion() { }

    [Fact]
    public void Compare_SlideDeleted_DetectsDeletion() { }

    [Fact]
    public void Compare_SlidesReordered_DetectsMove() { }

    [Fact]
    public void Compare_MultipleSlideChanges_DetectsAll() { }

    // === Shape Structure Tests ===

    [Fact]
    public void Compare_ShapeAdded_DetectsInsertion() { }

    [Fact]
    public void Compare_ShapeDeleted_DetectsDeletion() { }

    [Fact]
    public void Compare_ShapeMoved_DetectsMove() { }

    [Fact]
    public void Compare_ShapeResized_DetectsResize() { }

    [Fact]
    public void Compare_ShapeRotated_DetectsRotation() { }

    // === Text Content Tests ===

    [Fact]
    public void Compare_TextChanged_DetectsChange() { }

    [Fact]
    public void Compare_TextFormattingChanged_DetectsFormatChange() { }

    [Fact]
    public void Compare_TextAddedToShape_DetectsAddition() { }

    // === Image Tests ===

    [Fact]
    public void Compare_ImageReplaced_DetectsReplacement() { }

    [Fact]
    public void Compare_SameImage_NoChange() { }

    // === Matching Tests ===

    [Fact]
    public void ShapeMatching_ByPlaceholder_MatchesCorrectly() { }

    [Fact]
    public void ShapeMatching_ByName_MatchesCorrectly() { }

    [Fact]
    public void ShapeMatching_Fuzzy_MatchesSimilarShapes() { }

    [Fact]
    public void SlideMatching_ByTitle_MatchesCorrectly() { }

    [Fact]
    public void SlideMatching_ByFingerprint_MatchesReorderedSlides() { }

    // === Edge Cases ===

    [Fact]
    public void Compare_EmptyPresentation_NoErrors() { }

    [Fact]
    public void Compare_SingleSlide_Works() { }

    [Fact]
    public void Compare_LargePresentation_PerformanceAcceptable() { }

    [Fact]
    public void Compare_GroupedShapes_HandlesCorrectly() { }

    // === Renderer Tests ===

    [Fact]
    public void ProduceMarkedPresentation_AddsOverlays() { }

    [Fact]
    public void ProduceMarkedPresentation_OpensInPowerPoint() { }
}
```

### 9.2 Test Helpers

```csharp
internal static class PmlTestHelpers
{
    public static PmlDocument CreateTestPresentation(Action<PresentationDocument> configure)
    {
        using var ms = new MemoryStream();
        using (var pDoc = PresentationDocument.Create(ms, PresentationDocumentType.Presentation))
        {
            // Initialize minimal structure
            var presentationPart = pDoc.AddPresentationPart();
            presentationPart.Presentation = new Presentation();

            configure(pDoc);

            pDoc.Save();
        }

        return new PmlDocument("test.pptx", ms.ToArray());
    }

    public static void AddSlideWithTextBox(
        PresentationDocument pDoc,
        string text,
        string shapeName = "TextBox 1")
    {
        // Add slide with a text box containing the specified text
        // ...
    }

    public static void AddSlideWithImage(
        PresentationDocument pDoc,
        byte[] imageBytes,
        string shapeName = "Picture 1")
    {
        // Add slide with an image
        // ...
    }
}
```

---

## Appendix A: PPTX Structure Reference

### Key Parts and Relationships

```
/ppt/presentation.xml          - Main presentation part
  └── p:sldIdLst               - Ordered list of slide references

/ppt/slides/slide1.xml         - Slide content
  └── p:cSld
      └── p:spTree             - Shape tree (z-ordered)
          ├── p:nvGrpSpPr      - Non-visual group shape properties
          ├── p:grpSpPr        - Group shape properties
          └── p:sp, p:pic, ... - Shapes

/ppt/slideMasters/slideMaster1.xml
/ppt/slideLayouts/slideLayout1.xml
/ppt/theme/theme1.xml
/ppt/media/image1.png          - Embedded media
```

### Shape Elements

| Element | Description |
|---------|-------------|
| `p:sp` | AutoShape or text box |
| `p:pic` | Picture |
| `p:graphicFrame` | Table, chart, diagram, or OLE object |
| `p:grpSp` | Group shape |
| `p:cxnSp` | Connector |

### Key Attributes

| Path | Description |
|------|-------------|
| `p:nvSpPr/p:cNvPr/@id` | Shape ID (may change) |
| `p:nvSpPr/p:cNvPr/@name` | Shape name (more stable) |
| `p:nvSpPr/p:nvPr/p:ph/@type` | Placeholder type |
| `p:nvSpPr/p:nvPr/p:ph/@idx` | Placeholder index |
| `p:spPr/a:xfrm/a:off/@x,@y` | Position (EMUs) |
| `p:spPr/a:xfrm/a:ext/@cx,@cy` | Size (EMUs) |
| `p:spPr/a:xfrm/@rot` | Rotation (60000ths of degree) |

---

## Appendix B: Design Decisions

### Why Signature-Based (like SmlComparer) vs Atom-Based (like WmlComparer)?

1. **Slides are discrete units**: Unlike Word's flowing paragraphs, slides are independent. No need for complex document reconstruction.

2. **No native revision tracking**: Word's atom approach enables tracked-changes output. PowerPoint has no such target format.

3. **Cleaner mental model**: Presentation → Slides → Shapes → Content is a natural hierarchy for comparison.

4. **Performance**: Signature comparison is faster than building atom chains for every text run.

### Why Not Raw XML Diff?

1. **ID churn**: Shape IDs, relationship IDs change across saves
2. **Extension noise**: `extLst` elements vary by PowerPoint version
3. **Reserialization variance**: Semantically identical XML may serialize differently
4. **Poor user experience**: Raw XML diffs are unreadable

### Why Multiple Renderers?

Different use cases need different outputs:
- **Markup deck**: Quick visual review, works offline, any PowerPoint version
- **Two-up**: Detailed side-by-side for careful review
- **Merge**: Programmatic change application
- **JSON result**: API integration, custom UIs

---

## Appendix C: Future Considerations

### Three-Way Merge

For collaborative scenarios: merge changes from two branches against a common ancestor.

### Office Add-in Integration

Store diff metadata in a Custom XML part. Build a taskpane add-in that:
- Shows revisions list
- Allows accept/reject per change
- Applies changes to live presentation

### Real-Time Collaboration Diff

Compare against collaboration session revision history (if accessible).

### AI-Assisted Matching

Use embedding similarity for fuzzy text matching when exact matches fail.
