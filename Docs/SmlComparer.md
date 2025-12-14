# SmlComparer - Excel Spreadsheet Comparison

SmlComparer provides functionality to compare two Excel spreadsheet (.xlsx) files and identify differences, similar to how WmlComparer works for Word documents.

## Overview

The comparison process involves three main stages:
1. **Canonicalization** - Convert workbooks to normalized internal representations
2. **Diff Computation** - Compare the canonical forms and identify changes
3. **Rendering** - Produce a marked workbook with highlighted differences

## Quick Start

```csharp
using OpenXmlPowerTools;

// Load two workbooks
var older = new SmlDocument("original.xlsx");
var newer = new SmlDocument("modified.xlsx");

// Compare and get results
var settings = new SmlComparerSettings();
var result = SmlComparer.Compare(older, newer, settings);

// Get summary
Console.WriteLine($"Total Changes: {result.TotalChanges}");
Console.WriteLine($"Value Changes: {result.ValueChanges}");
Console.WriteLine($"Formula Changes: {result.FormulaChanges}");

// Export to JSON
string json = result.ToJson();

// Produce a marked workbook with highlights
var marked = SmlComparer.ProduceMarkedWorkbook(older, newer, settings);
marked.SaveAs("comparison-result.xlsx");
```

## Architecture

### Class Diagram

```
┌─────────────────────────────────────────────────────────────────┐
│                        SmlComparer                               │
│  ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐  │
│  │ Compare()       │  │ ProduceMarked   │  │ SmlComparer     │  │
│  │                 │  │ Workbook()      │  │ Settings        │  │
│  └────────┬────────┘  └────────┬────────┘  └─────────────────┘  │
│           │                    │                                 │
│           ▼                    ▼                                 │
│  ┌─────────────────────────────────────────────────────────────┐│
│  │                   SmlCanonicalizer                          ││
│  │  - Canonicalize workbooks to WorkbookSignature              ││
│  │  - Resolve shared strings, expand styles                    ││
│  │  - Extract Phase 3 features (comments, validations, etc.)   ││
│  └─────────────────────────────────────────────────────────────┘│
│           │                                                      │
│           ▼                                                      │
│  ┌─────────────────────────────────────────────────────────────┐│
│  │                    SmlDiffEngine                            ││
│  │  - Match sheets (including rename detection)                ││
│  │  - Row/column alignment using LCS algorithm                 ││
│  │  - Cell-by-cell comparison with tolerance options           ││
│  │  - Phase 3 comparisons (named ranges, comments, etc.)       ││
│  └─────────────────────────────────────────────────────────────┘│
│           │                                                      │
│           ▼                                                      │
│  ┌─────────────────────────────────────────────────────────────┐│
│  │                  SmlMarkupRenderer                          ││
│  │  - Apply cell highlighting based on change type             ││
│  │  - Add comments explaining changes                          ││
│  │  - Create _DiffSummary sheet                                ││
│  └─────────────────────────────────────────────────────────────┘│
└─────────────────────────────────────────────────────────────────┘
```

### Internal Data Structures

```csharp
// Canonical workbook representation
internal class WorkbookSignature
{
    Dictionary<string, WorksheetSignature> Sheets;
    Dictionary<string, string> DefinedNames;  // Named ranges
}

// Canonical worksheet representation
internal class WorksheetSignature
{
    string Name;
    Dictionary<string, CellSignature> Cells;

    // Phase 2: Row alignment
    SortedSet<int> PopulatedRows;
    Dictionary<int, string> RowSignatures;

    // Phase 3: Additional features
    Dictionary<string, CommentSignature> Comments;
    Dictionary<string, DataValidationSignature> DataValidations;
    HashSet<string> MergedCellRanges;
    Dictionary<string, HyperlinkSignature> Hyperlinks;
}

// Canonical cell representation
internal class CellSignature
{
    string Address;
    int Row, Column;
    string ResolvedValue;
    string Formula;
    CellFormatSignature Format;
    string ContentHash;
}
```

## Settings Reference

### Basic Comparison Settings

| Setting | Default | Description |
|---------|---------|-------------|
| `CompareValues` | `true` | Compare cell values |
| `CompareFormulas` | `true` | Compare cell formulas |
| `CompareFormatting` | `true` | Compare cell formatting (bold, colors, etc.) |
| `CompareSheetStructure` | `true` | Detect added/deleted sheets |
| `CaseInsensitiveValues` | `false` | Ignore case in string comparisons |
| `NumericTolerance` | `0.0` | Tolerance for numeric comparisons |

### Phase 2: Alignment Settings

| Setting | Default | Description |
|---------|---------|-------------|
| `EnableRowAlignment` | `true` | Use LCS algorithm for row alignment |
| `EnableColumnAlignment` | `false` | Use LCS algorithm for column alignment |
| `EnableSheetRenameDetection` | `true` | Detect renamed sheets by content similarity |
| `SheetRenameSimilarityThreshold` | `0.7` | Minimum Jaccard similarity for rename detection |
| `RowSignatureSampleSize` | `10` | Number of cells to sample per row for hashing |

### Phase 3: Extended Feature Settings

| Setting | Default | Description |
|---------|---------|-------------|
| `CompareNamedRanges` | `true` | Compare defined names |
| `CompareComments` | `true` | Compare cell comments/notes |
| `CompareDataValidation` | `true` | Compare data validation rules |
| `CompareMergedCells` | `true` | Compare merged cell regions |
| `CompareHyperlinks` | `true` | Compare cell hyperlinks |
| `CompareConditionalFormatting` | `true` | Compare conditional formatting rules |

### Highlight Colors

| Setting | Default | Description |
|---------|---------|-------------|
| `AddedCellColor` | `"90EE90"` | Light green for added cells |
| `DeletedCellColor` | `"FFCCCB"` | Light red for deleted cells |
| `ModifiedValueColor` | `"FFD700"` | Gold for value changes |
| `ModifiedFormulaColor` | `"87CEEB"` | Sky blue for formula changes |
| `ModifiedFormatColor` | `"E6E6FA"` | Lavender for format changes |
| `InsertedRowColor` | `"E0FFFF"` | Light cyan for inserted rows |
| `DeletedRowColor` | `"FFE4E1"` | Misty rose for deleted rows |
| `NamedRangeChangeColor` | `"DDA0DD"` | Light purple for named range changes |
| `CommentChangeColor` | `"FFFACD"` | Light yellow for comment changes |
| `DataValidationChangeColor` | `"FFDAB9"` | Light orange for validation changes |

## How the Renderer Works

### Overview

The `SmlMarkupRenderer` class transforms comparison results into a visual representation by:
1. Adding highlight fill styles to the workbook's stylesheet
2. Applying style indices to changed cells
3. Adding cell comments with change details
4. Creating a `_DiffSummary` sheet

### Rendering Pipeline

```
┌────────────────────────────────────────────────────────────────┐
│                    SmlMarkupRenderer                            │
├────────────────────────────────────────────────────────────────┤
│  1. RenderMarkedWorkbook(source, result, settings)             │
│     │                                                          │
│     ├─► 2. AddHighlightStyles(styleXDoc, settings)             │
│     │      - Add fill patterns for each change type            │
│     │      - Add cellXf entries referencing the fills          │
│     │      - Return HighlightStyles with style IDs             │
│     │                                                          │
│     ├─► 3. For each sheet with changes:                        │
│     │      │                                                   │
│     │      ├─► ApplyCellHighlight(wsXDoc, change, styles)      │
│     │      │   - Find or create cell element                   │
│     │      │   - Set style index (s attribute)                 │
│     │      │                                                   │
│     │      └─► AddCommentsForChanges(worksheetPart, changes)   │
│     │          - Create/update WorksheetCommentsPart           │
│     │          - Add VML drawing for comment display           │
│     │                                                          │
│     └─► 4. AddDiffSummarySheet(sDoc, result, settings)         │
│          - Create new worksheet with summary statistics        │
│          - List all changes with descriptions                  │
└────────────────────────────────────────────────────────────────┘
```

### Style Application

The renderer modifies the workbook's `styles.xml` to add highlight fills:

```xml
<!-- Added to <fills> element -->
<fill>
  <patternFill patternType="solid">
    <fgColor rgb="FF90EE90"/>  <!-- Light green for added cells -->
    <bgColor indexed="64"/>
  </patternFill>
</fill>

<!-- Added to <cellXfs> element -->
<xf numFmtId="0" fontId="0" fillId="5" borderId="0" applyFill="1"/>
```

Then cells are marked with the appropriate style index:

```xml
<c r="A1" s="5">  <!-- s="5" references the highlight style -->
  <v>New Value</v>
</c>
```

### Comment Structure

Comments are added via `WorksheetCommentsPart`:

```xml
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors>
    <author>Open-Xml-PowerTools</author>
  </authors>
  <commentList>
    <comment ref="A1" authorId="0">
      <text>
        <r><t>[ValueChanged]
Old value: Hello
New value: World</t></r>
      </text>
    </comment>
  </commentList>
</comments>
```

VML drawings are required for Excel to display comments (legacy format).

## Extending the Renderer

### Custom Highlight Colors

The simplest customization is changing the highlight colors:

```csharp
var settings = new SmlComparerSettings
{
    // Use corporate brand colors
    AddedCellColor = "00FF00",      // Bright green
    ModifiedValueColor = "0000FF",  // Blue
    DeletedCellColor = "FF0000"     // Red
};
```

### Custom Change Processing

To add custom processing for specific change types, you can post-process the result:

```csharp
var result = SmlComparer.Compare(older, newer, settings);

// Add custom metadata to changes
foreach (var change in result.Changes)
{
    if (change.ChangeType == SmlChangeType.ValueChanged)
    {
        // Calculate percentage change for numeric values
        if (double.TryParse(change.OldValue, out var oldVal) &&
            double.TryParse(change.NewValue, out var newVal) &&
            oldVal != 0)
        {
            var pctChange = (newVal - oldVal) / oldVal * 100;
            // Store or use pctChange as needed
        }
    }
}
```

### Creating a Custom Renderer

For more control, implement a custom renderer:

```csharp
public class CustomMarkupRenderer
{
    public static SmlDocument RenderMarkedWorkbook(
        SmlDocument source,
        SmlComparisonResult result,
        SmlComparerSettings settings)
    {
        using var ms = new MemoryStream();
        ms.Write(source.DocumentByteArray, 0, source.DocumentByteArray.Length);

        using (var sDoc = SpreadsheetDocument.Open(ms, true))
        {
            var workbookPart = sDoc.WorkbookPart;

            // 1. Add your custom styles
            AddCustomStyles(workbookPart.WorkbookStylesPart);

            // 2. Apply cell highlighting
            foreach (var change in result.Changes.Where(c => c.CellAddress != null))
            {
                ApplyCustomHighlight(workbookPart, change);
            }

            // 3. Add custom summary (e.g., charts, pivot tables)
            AddCustomSummary(sDoc, result);
        }

        return new SmlDocument("marked.xlsx", ms.ToArray());
    }

    private static void AddCustomStyles(WorkbookStylesPart stylesPart)
    {
        // Add gradient fills, borders, etc.
    }

    private static void ApplyCustomHighlight(WorkbookPart workbookPart, SmlChange change)
    {
        // Apply conditional formatting instead of static fills
        // Add data bars for numeric changes
        // Add icons for different change severities
    }

    private static void AddCustomSummary(SpreadsheetDocument sDoc, SmlComparisonResult result)
    {
        // Add charts showing change distribution
        // Add pivot table for change analysis
    }
}
```

### Adding New Change Types

To add support for new comparison features:

1. **Add the change type enum**:
```csharp
public enum SmlChangeType
{
    // ... existing types ...

    // Your new type
    ConditionalFormatChanged,
    PivotTableChanged,
    ChartChanged
}
```

2. **Add properties to SmlChange**:
```csharp
public class SmlChange
{
    // ... existing properties ...

    public string ChartName { get; set; }
    public string OldChartType { get; set; }
    public string NewChartType { get; set; }
}
```

3. **Update the canonicalizer** to extract the new features:
```csharp
private static void ExtractCharts(WorksheetPart worksheetPart, WorksheetSignature signature)
{
    var drawingPart = worksheetPart.DrawingsPart;
    if (drawingPart == null) return;

    // Extract chart information
}
```

4. **Update the diff engine** to compare the new features:
```csharp
private static void CompareCharts(
    WorksheetSignature ws1,
    WorksheetSignature ws2,
    string sheetName,
    SmlComparisonResult result)
{
    // Compare charts and add changes
}
```

5. **Update the renderer** to visualize changes:
```csharp
private static void HighlightChartChanges(...)
{
    // Add visual indicators for chart changes
}
```

## Future Work

### Planned Features

1. **Conditional Formatting Comparison** (infrastructure ready)
   - Compare CF rules and priorities
   - Detect added/deleted/modified rules

2. **Chart Comparison**
   - Detect chart type changes
   - Compare data ranges
   - Track axis/legend modifications

3. **Pivot Table Comparison**
   - Compare pivot table structure
   - Track field changes
   - Detect filter modifications

4. **Performance Improvements**
   - Parallel sheet processing
   - Streaming for large workbooks
   - Memory-mapped file support

5. **Enhanced Visualization**
   - Side-by-side comparison view
   - Change heatmaps
   - Timeline of changes (for version history)

### Extension Points

| Area | How to Extend |
|------|---------------|
| **New Features** | Add signature classes, extraction methods, comparison logic |
| **Custom Styling** | Implement custom renderer with advanced Excel styling |
| **Output Formats** | Add exporters for HTML, PDF, or other formats |
| **Filtering** | Add change filters to result processing |
| **Validation** | Add semantic validation of changes |

## Troubleshooting

### Common Issues

**Row alignment produces unexpected results**
- Try adjusting `RowSignatureSampleSize` for wide spreadsheets
- Set `EnableRowAlignment = false` for simple cell-by-cell comparison

**Sheet rename not detected**
- Lower `SheetRenameSimilarityThreshold` (default 0.7)
- Ensure sheets have sufficient content for similarity calculation

**Performance with large files**
- Disable features you don't need (`CompareFormatting = false`)
- Use `EnableColumnAlignment = false` (expensive operation)
- Consider comparing specific sheets only

### Debugging

Enable logging to trace comparison:

```csharp
var settings = new SmlComparerSettings
{
    LogCallback = message => Console.WriteLine($"[SmlComparer] {message}")
};
```

## API Reference

### SmlComparer

```csharp
public static class SmlComparer
{
    // Compare two workbooks and return structured results
    public static SmlComparisonResult Compare(
        SmlDocument older,
        SmlDocument newer,
        SmlComparerSettings settings);

    // Produce a marked workbook with visual highlights
    public static SmlDocument ProduceMarkedWorkbook(
        SmlDocument older,
        SmlDocument newer,
        SmlComparerSettings settings);
}
```

### SmlComparisonResult

```csharp
public class SmlComparisonResult
{
    // All detected changes
    List<SmlChange> Changes { get; }

    // Statistics
    int TotalChanges { get; }
    int ValueChanges { get; }
    int FormulaChanges { get; }
    int FormatChanges { get; }
    int CellsAdded { get; }
    int CellsDeleted { get; }
    int SheetsAdded { get; }
    int SheetsDeleted { get; }
    int SheetsRenamed { get; }
    int RowsInserted { get; }
    int RowsDeleted { get; }
    // ... Phase 3 statistics ...

    // Filtering
    IEnumerable<SmlChange> GetChangesBySheet(string sheetName);
    IEnumerable<SmlChange> GetChangesByType(SmlChangeType type);

    // Export
    string ToJson();
}
```

### SmlChange

```csharp
public class SmlChange
{
    SmlChangeType ChangeType { get; set; }
    string SheetName { get; set; }
    string CellAddress { get; set; }

    // Values
    string OldValue { get; set; }
    string NewValue { get; set; }
    string OldFormula { get; set; }
    string NewFormula { get; set; }

    // Formatting
    CellFormatSignature OldFormat { get; set; }
    CellFormatSignature NewFormat { get; set; }

    // Phase 3 properties
    string NamedRangeName { get; set; }
    string OldComment { get; set; }
    string NewComment { get; set; }
    // ... etc ...

    // Human-readable description
    string GetDescription();
}
```
