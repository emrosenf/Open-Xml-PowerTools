// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    /// <summary>
    /// Settings for controlling spreadsheet comparison behavior.
    /// </summary>
    public class SmlComparerSettings
    {
        /// <summary>Whether to compare cell values.</summary>
        public bool CompareValues = true;

        /// <summary>Whether to compare cell formulas.</summary>
        public bool CompareFormulas = true;

        /// <summary>Whether to compare cell formatting.</summary>
        public bool CompareFormatting = true;

        /// <summary>Whether to compare sheet structure (added/removed sheets).</summary>
        public bool CompareSheetStructure = true;

        /// <summary>Whether value comparison should be case-insensitive.</summary>
        public bool CaseInsensitiveValues = false;

        /// <summary>Tolerance for numeric comparison (0.0 for exact match).</summary>
        public double NumericTolerance = 0.0;

        /// <summary>Author name for change annotations.</summary>
        public string AuthorForChanges = "Open-Xml-PowerTools";

        /// <summary>Optional callback for logging.</summary>
        public Action<string> LogCallback = null;

        // Highlight colors (ARGB hex without #)
        /// <summary>Fill color for added cells (default: light green).</summary>
        public string AddedCellColor = "90EE90";

        /// <summary>Fill color for deleted cells in summary (default: light red).</summary>
        public string DeletedCellColor = "FFCCCB";

        /// <summary>Fill color for value changes (default: gold).</summary>
        public string ModifiedValueColor = "FFD700";

        /// <summary>Fill color for formula changes (default: sky blue).</summary>
        public string ModifiedFormulaColor = "87CEEB";

        /// <summary>Fill color for format-only changes (default: lavender).</summary>
        public string ModifiedFormatColor = "E6E6FA";
    }

    /// <summary>
    /// Types of changes detected during spreadsheet comparison.
    /// </summary>
    public enum SmlChangeType
    {
        // Workbook structure
        SheetAdded,
        SheetDeleted,

        // Cell content
        CellAdded,
        CellDeleted,
        ValueChanged,
        FormulaChanged,
        FormatChanged,
    }

    /// <summary>
    /// Represents a single change between two spreadsheets.
    /// </summary>
    public class SmlChange
    {
        public SmlChangeType ChangeType { get; set; }
        public string SheetName { get; set; }
        public string CellAddress { get; set; }

        public string OldValue { get; set; }
        public string NewValue { get; set; }
        public string OldFormula { get; set; }
        public string NewFormula { get; set; }
        public CellFormatSignature OldFormat { get; set; }
        public CellFormatSignature NewFormat { get; set; }

        /// <summary>
        /// Returns a human-readable description of this change.
        /// </summary>
        public string GetDescription()
        {
            return ChangeType switch
            {
                SmlChangeType.SheetAdded => $"Sheet '{SheetName}' was added",
                SmlChangeType.SheetDeleted => $"Sheet '{SheetName}' was deleted",
                SmlChangeType.CellAdded => $"Cell {SheetName}!{CellAddress} was added with value '{NewValue}'",
                SmlChangeType.CellDeleted => $"Cell {SheetName}!{CellAddress} was deleted (had value '{OldValue}')",
                SmlChangeType.ValueChanged => $"Cell {SheetName}!{CellAddress} value changed from '{OldValue}' to '{NewValue}'",
                SmlChangeType.FormulaChanged => $"Cell {SheetName}!{CellAddress} formula changed from '{OldFormula}' to '{NewFormula}'",
                SmlChangeType.FormatChanged => $"Cell {SheetName}!{CellAddress} formatting changed",
                _ => $"Unknown change at {SheetName}!{CellAddress}"
            };
        }
    }

    /// <summary>
    /// Result of comparing two spreadsheets, containing all detected changes.
    /// </summary>
    public class SmlComparisonResult
    {
        public List<SmlChange> Changes { get; } = new List<SmlChange>();

        // Computed statistics
        public int TotalChanges => Changes.Count;
        public int ValueChanges => Changes.Count(c => c.ChangeType == SmlChangeType.ValueChanged);
        public int FormulaChanges => Changes.Count(c => c.ChangeType == SmlChangeType.FormulaChanged);
        public int FormatChanges => Changes.Count(c => c.ChangeType == SmlChangeType.FormatChanged);
        public int CellsAdded => Changes.Count(c => c.ChangeType == SmlChangeType.CellAdded);
        public int CellsDeleted => Changes.Count(c => c.ChangeType == SmlChangeType.CellDeleted);
        public int SheetsAdded => Changes.Count(c => c.ChangeType == SmlChangeType.SheetAdded);
        public int SheetsDeleted => Changes.Count(c => c.ChangeType == SmlChangeType.SheetDeleted);

        public int StructuralChanges => CellsAdded + CellsDeleted + SheetsAdded + SheetsDeleted;

        /// <summary>
        /// Get all changes for a specific sheet.
        /// </summary>
        public IEnumerable<SmlChange> GetChangesBySheet(string sheetName)
        {
            return Changes.Where(c => c.SheetName == sheetName);
        }

        /// <summary>
        /// Get all changes of a specific type.
        /// </summary>
        public IEnumerable<SmlChange> GetChangesByType(SmlChangeType type)
        {
            return Changes.Where(c => c.ChangeType == type);
        }

        /// <summary>
        /// Export the comparison result to JSON.
        /// </summary>
        public string ToJson()
        {
            var sb = new StringBuilder();
            sb.AppendLine("{");
            sb.AppendLine("  \"Summary\": {");
            sb.AppendLine($"    \"TotalChanges\": {TotalChanges},");
            sb.AppendLine($"    \"ValueChanges\": {ValueChanges},");
            sb.AppendLine($"    \"FormulaChanges\": {FormulaChanges},");
            sb.AppendLine($"    \"FormatChanges\": {FormatChanges},");
            sb.AppendLine($"    \"CellsAdded\": {CellsAdded},");
            sb.AppendLine($"    \"CellsDeleted\": {CellsDeleted},");
            sb.AppendLine($"    \"SheetsAdded\": {SheetsAdded},");
            sb.AppendLine($"    \"SheetsDeleted\": {SheetsDeleted}");
            sb.AppendLine("  },");
            sb.AppendLine("  \"Changes\": [");

            for (int i = 0; i < Changes.Count; i++)
            {
                var c = Changes[i];
                sb.AppendLine("    {");
                sb.AppendLine($"      \"ChangeType\": \"{c.ChangeType}\",");
                sb.AppendLine($"      \"SheetName\": {JsonEscape(c.SheetName)},");
                sb.AppendLine($"      \"CellAddress\": {JsonEscape(c.CellAddress)},");
                sb.AppendLine($"      \"OldValue\": {JsonEscape(c.OldValue)},");
                sb.AppendLine($"      \"NewValue\": {JsonEscape(c.NewValue)},");
                sb.AppendLine($"      \"OldFormula\": {JsonEscape(c.OldFormula)},");
                sb.AppendLine($"      \"NewFormula\": {JsonEscape(c.NewFormula)},");
                sb.AppendLine($"      \"Description\": {JsonEscape(c.GetDescription())}");
                sb.Append("    }");
                if (i < Changes.Count - 1) sb.Append(",");
                sb.AppendLine();
            }

            sb.AppendLine("  ]");
            sb.AppendLine("}");
            return sb.ToString();
        }

        private static string JsonEscape(string value)
        {
            if (value == null) return "null";
            return "\"" + value
                .Replace("\\", "\\\\")
                .Replace("\"", "\\\"")
                .Replace("\n", "\\n")
                .Replace("\r", "\\r")
                .Replace("\t", "\\t") + "\"";
        }
    }

    /// <summary>
    /// Represents the expanded formatting of a cell for comparison purposes.
    /// Style indices are resolved to actual formatting properties.
    /// </summary>
    public class CellFormatSignature : IEquatable<CellFormatSignature>
    {
        // Number format
        public string NumberFormatCode { get; set; }

        // Font
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public bool Strikethrough { get; set; }
        public string FontName { get; set; }
        public double? FontSize { get; set; }
        public string FontColor { get; set; }

        // Fill
        public string FillPattern { get; set; }
        public string FillForegroundColor { get; set; }
        public string FillBackgroundColor { get; set; }

        // Border
        public string BorderLeftStyle { get; set; }
        public string BorderLeftColor { get; set; }
        public string BorderRightStyle { get; set; }
        public string BorderRightColor { get; set; }
        public string BorderTopStyle { get; set; }
        public string BorderTopColor { get; set; }
        public string BorderBottomStyle { get; set; }
        public string BorderBottomColor { get; set; }

        // Alignment
        public string HorizontalAlignment { get; set; }
        public string VerticalAlignment { get; set; }
        public bool WrapText { get; set; }
        public int? Indent { get; set; }

        public bool Equals(CellFormatSignature other)
        {
            if (other == null) return false;
            return NumberFormatCode == other.NumberFormatCode &&
                   Bold == other.Bold &&
                   Italic == other.Italic &&
                   Underline == other.Underline &&
                   Strikethrough == other.Strikethrough &&
                   FontName == other.FontName &&
                   FontSize == other.FontSize &&
                   FontColor == other.FontColor &&
                   FillPattern == other.FillPattern &&
                   FillForegroundColor == other.FillForegroundColor &&
                   FillBackgroundColor == other.FillBackgroundColor &&
                   BorderLeftStyle == other.BorderLeftStyle &&
                   BorderLeftColor == other.BorderLeftColor &&
                   BorderRightStyle == other.BorderRightStyle &&
                   BorderRightColor == other.BorderRightColor &&
                   BorderTopStyle == other.BorderTopStyle &&
                   BorderTopColor == other.BorderTopColor &&
                   BorderBottomStyle == other.BorderBottomStyle &&
                   BorderBottomColor == other.BorderBottomColor &&
                   HorizontalAlignment == other.HorizontalAlignment &&
                   VerticalAlignment == other.VerticalAlignment &&
                   WrapText == other.WrapText &&
                   Indent == other.Indent;
        }

        public override bool Equals(object obj) => Equals(obj as CellFormatSignature);

        public override int GetHashCode()
        {
            unchecked
            {
                int hash = 17;
                hash = hash * 31 + (NumberFormatCode?.GetHashCode() ?? 0);
                hash = hash * 31 + Bold.GetHashCode();
                hash = hash * 31 + Italic.GetHashCode();
                hash = hash * 31 + Underline.GetHashCode();
                hash = hash * 31 + Strikethrough.GetHashCode();
                hash = hash * 31 + (FontName?.GetHashCode() ?? 0);
                hash = hash * 31 + (FontSize?.GetHashCode() ?? 0);
                hash = hash * 31 + (FontColor?.GetHashCode() ?? 0);
                hash = hash * 31 + (FillPattern?.GetHashCode() ?? 0);
                hash = hash * 31 + (FillForegroundColor?.GetHashCode() ?? 0);
                hash = hash * 31 + (FillBackgroundColor?.GetHashCode() ?? 0);
                hash = hash * 31 + (HorizontalAlignment?.GetHashCode() ?? 0);
                hash = hash * 31 + (VerticalAlignment?.GetHashCode() ?? 0);
                hash = hash * 31 + WrapText.GetHashCode();
                return hash;
            }
        }

        /// <summary>
        /// Returns a human-readable description of the differences between this format and another.
        /// </summary>
        public string GetDifferenceDescription(CellFormatSignature other)
        {
            if (other == null) return "Format added";
            if (Equals(other)) return "No difference";

            var diffs = new List<string>();

            if (NumberFormatCode != other.NumberFormatCode)
                diffs.Add($"Number format: '{other.NumberFormatCode}' → '{NumberFormatCode}'");
            if (Bold != other.Bold)
                diffs.Add(Bold ? "Made bold" : "Removed bold");
            if (Italic != other.Italic)
                diffs.Add(Italic ? "Made italic" : "Removed italic");
            if (Underline != other.Underline)
                diffs.Add(Underline ? "Added underline" : "Removed underline");
            if (Strikethrough != other.Strikethrough)
                diffs.Add(Strikethrough ? "Added strikethrough" : "Removed strikethrough");
            if (FontName != other.FontName)
                diffs.Add($"Font: '{other.FontName}' → '{FontName}'");
            if (FontSize != other.FontSize)
                diffs.Add($"Size: {other.FontSize} → {FontSize}");
            if (FontColor != other.FontColor)
                diffs.Add($"Font color: {other.FontColor} → {FontColor}");
            if (FillForegroundColor != other.FillForegroundColor)
                diffs.Add($"Fill color: {other.FillForegroundColor} → {FillForegroundColor}");
            if (HorizontalAlignment != other.HorizontalAlignment)
                diffs.Add($"Horizontal align: {other.HorizontalAlignment} → {HorizontalAlignment}");
            if (VerticalAlignment != other.VerticalAlignment)
                diffs.Add($"Vertical align: {other.VerticalAlignment} → {VerticalAlignment}");
            if (WrapText != other.WrapText)
                diffs.Add(WrapText ? "Enabled wrap text" : "Disabled wrap text");

            return diffs.Count > 0 ? string.Join("; ", diffs) : "Minor formatting change";
        }

        public static CellFormatSignature Default => new CellFormatSignature
        {
            NumberFormatCode = "General",
            Bold = false,
            Italic = false,
            Underline = false,
            Strikethrough = false,
            FontName = "Calibri",
            FontSize = 11,
            HorizontalAlignment = "general",
            VerticalAlignment = "bottom",
            WrapText = false
        };
    }

    /// <summary>
    /// Compares two Excel spreadsheets and produces a marked workbook showing differences.
    /// Similar to Microsoft Spreadsheet Compare functionality.
    /// </summary>
    public static class SmlComparer
    {
        /// <summary>
        /// Compare two workbooks and return a structured list of changes.
        /// </summary>
        /// <param name="older">The original/older workbook.</param>
        /// <param name="newer">The revised/newer workbook.</param>
        /// <param name="settings">Comparison settings.</param>
        /// <returns>A result object containing all detected changes.</returns>
        public static SmlComparisonResult Compare(SmlDocument older, SmlDocument newer, SmlComparerSettings settings)
        {
            if (older == null) throw new ArgumentNullException(nameof(older));
            if (newer == null) throw new ArgumentNullException(nameof(newer));
            settings ??= new SmlComparerSettings();

            Log(settings, "SmlComparer.Compare: Starting comparison");

            // Canonicalize both workbooks
            var sig1 = SmlCanonicalizer.Canonicalize(older, settings);
            var sig2 = SmlCanonicalizer.Canonicalize(newer, settings);

            Log(settings, $"SmlComparer.Compare: Canonicalized older workbook: {sig1.Sheets.Count} sheets");
            Log(settings, $"SmlComparer.Compare: Canonicalized newer workbook: {sig2.Sheets.Count} sheets");

            // Compute diff
            var result = SmlDiffEngine.ComputeDiff(sig1, sig2, settings);

            Log(settings, $"SmlComparer.Compare: Found {result.TotalChanges} changes");

            return result;
        }

        /// <summary>
        /// Produce a marked workbook highlighting all differences between two workbooks.
        /// The output is based on the newer workbook with highlights and comments showing changes.
        /// </summary>
        /// <param name="older">The original/older workbook.</param>
        /// <param name="newer">The revised/newer workbook.</param>
        /// <param name="settings">Comparison settings.</param>
        /// <returns>A new workbook with changes highlighted.</returns>
        public static SmlDocument ProduceMarkedWorkbook(SmlDocument older, SmlDocument newer, SmlComparerSettings settings)
        {
            if (older == null) throw new ArgumentNullException(nameof(older));
            if (newer == null) throw new ArgumentNullException(nameof(newer));
            settings ??= new SmlComparerSettings();

            Log(settings, "SmlComparer.ProduceMarkedWorkbook: Starting");

            // First compute the diff
            var result = Compare(older, newer, settings);

            // Then render the marked workbook
            var markedWorkbook = SmlMarkupRenderer.RenderMarkedWorkbook(newer, result, settings);

            Log(settings, "SmlComparer.ProduceMarkedWorkbook: Complete");

            return markedWorkbook;
        }

        private static void Log(SmlComparerSettings settings, string message)
        {
            settings?.LogCallback?.Invoke(message);
        }
    }

    #region Internal Implementation Classes

    /// <summary>
    /// Internal canonical representation of a workbook for comparison.
    /// </summary>
    internal class WorkbookSignature
    {
        public Dictionary<string, WorksheetSignature> Sheets { get; } = new Dictionary<string, WorksheetSignature>();
        public Dictionary<string, string> DefinedNames { get; } = new Dictionary<string, string>();
    }

    /// <summary>
    /// Internal canonical representation of a worksheet for comparison.
    /// </summary>
    internal class WorksheetSignature
    {
        public string Name { get; set; }
        public string RelationshipId { get; set; }
        public Dictionary<string, CellSignature> Cells { get; } = new Dictionary<string, CellSignature>();
    }

    /// <summary>
    /// Internal canonical representation of a cell for comparison.
    /// </summary>
    internal class CellSignature
    {
        public string Address { get; set; }
        public int Row { get; set; }
        public int Column { get; set; }
        public string ResolvedValue { get; set; }
        public string Formula { get; set; }
        public string ContentHash { get; set; }
        public CellFormatSignature Format { get; set; }

        public static string ComputeHash(string value, string formula)
        {
            var content = $"{value ?? ""}|{formula ?? ""}";
            using var sha256 = SHA256.Create();
            var bytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(content));
            return Convert.ToBase64String(bytes);
        }
    }

    /// <summary>
    /// Canonicalizes spreadsheets into a normalized form for comparison.
    /// Resolves shared strings, expands style indices to actual formatting.
    /// </summary>
    internal static class SmlCanonicalizer
    {
        public static WorkbookSignature Canonicalize(SmlDocument doc, SmlComparerSettings settings)
        {
            var signature = new WorkbookSignature();

            using var ms = new MemoryStream();
            ms.Write(doc.DocumentByteArray, 0, doc.DocumentByteArray.Length);

            using var sDoc = SpreadsheetDocument.Open(ms, false);
            var workbookPart = sDoc.WorkbookPart;
            var wbXDoc = workbookPart.GetXDocument();

            // Get shared string table
            var sharedStrings = GetSharedStrings(workbookPart);

            // Get styles
            var styleInfo = GetStyleInfo(workbookPart);

            // Process each sheet
            var sheets = wbXDoc.Root.Elements(S.sheets).Elements(S.sheet);
            foreach (var sheet in sheets)
            {
                var sheetName = (string)sheet.Attribute("name");
                var rId = (string)sheet.Attribute(R.id);

                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(rId);
                var wsSignature = CanonicalizeWorksheet(worksheetPart, sheetName, rId, sharedStrings, styleInfo, settings);
                signature.Sheets[sheetName] = wsSignature;
            }

            // Get defined names
            var definedNames = wbXDoc.Root.Elements(S.definedNames).Elements(S.definedName);
            foreach (var dn in definedNames)
            {
                var name = (string)dn.Attribute("name");
                var value = (string)dn;
                if (!string.IsNullOrEmpty(name))
                    signature.DefinedNames[name] = value;
            }

            return signature;
        }

        private static WorksheetSignature CanonicalizeWorksheet(
            WorksheetPart worksheetPart,
            string sheetName,
            string rId,
            List<string> sharedStrings,
            StyleInfo styleInfo,
            SmlComparerSettings settings)
        {
            var signature = new WorksheetSignature
            {
                Name = sheetName,
                RelationshipId = rId
            };

            var wsXDoc = worksheetPart.GetXDocument();
            var sheetData = wsXDoc.Root.Element(S.sheetData);
            if (sheetData == null) return signature;

            foreach (var row in sheetData.Elements(S.row))
            {
                var rowIndex = (int?)row.Attribute("r") ?? 0;

                foreach (var cell in row.Elements(S.c))
                {
                    var cellRef = (string)cell.Attribute("r");
                    if (string.IsNullOrEmpty(cellRef)) continue;

                    var cellSig = CanonicalizeCell(cell, cellRef, sharedStrings, styleInfo, settings);
                    signature.Cells[cellRef] = cellSig;
                }
            }

            return signature;
        }

        private static CellSignature CanonicalizeCell(
            XElement cell,
            string cellRef,
            List<string> sharedStrings,
            StyleInfo styleInfo,
            SmlComparerSettings settings)
        {
            // Parse cell reference
            var (col, row) = ParseCellReference(cellRef);

            // Get value
            var resolvedValue = ResolveValue(cell, sharedStrings);

            // Get formula
            var formula = (string)cell.Element(S.f);

            // Get format
            var styleIndex = (int?)cell.Attribute("s") ?? 0;
            var format = ExpandStyle(styleIndex, styleInfo);

            var sig = new CellSignature
            {
                Address = cellRef,
                Row = row,
                Column = col,
                ResolvedValue = resolvedValue,
                Formula = formula,
                Format = format,
                ContentHash = CellSignature.ComputeHash(resolvedValue, formula)
            };

            return sig;
        }

        private static string ResolveValue(XElement cell, List<string> sharedStrings)
        {
            var cellType = (string)cell.Attribute("t");
            var valueElement = cell.Element(S.v);
            var rawValue = (string)valueElement;

            if (string.IsNullOrEmpty(rawValue))
            {
                // Check for inline string
                var inlineStr = cell.Element(S._is);
                if (inlineStr != null)
                {
                    return inlineStr.Descendants(S.t).Select(t => (string)t).StringConcatenate();
                }
                return null;
            }

            return cellType switch
            {
                "s" => ResolveSharedString(rawValue, sharedStrings),
                "str" => rawValue,
                "b" => rawValue == "1" ? "TRUE" : "FALSE",
                "e" => rawValue, // Error value
                _ => NormalizeNumeric(rawValue)
            };
        }

        private static string ResolveSharedString(string indexStr, List<string> sharedStrings)
        {
            if (int.TryParse(indexStr, out var index) && index >= 0 && index < sharedStrings.Count)
            {
                return sharedStrings[index];
            }
            return indexStr;
        }

        private static string NormalizeNumeric(string value)
        {
            if (string.IsNullOrEmpty(value)) return value;

            if (decimal.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var d))
            {
                // Normalize to consistent representation
                return d.ToString("G", CultureInfo.InvariantCulture);
            }
            return value;
        }

        private static (int column, int row) ParseCellReference(string cellRef)
        {
            var col = 0;
            var row = 0;
            var i = 0;

            // Parse column letters
            while (i < cellRef.Length && char.IsLetter(cellRef[i]))
            {
                col = col * 26 + (char.ToUpper(cellRef[i]) - 'A' + 1);
                i++;
            }

            // Parse row number
            if (i < cellRef.Length)
            {
                int.TryParse(cellRef.Substring(i), out row);
            }

            return (col, row);
        }

        private static List<string> GetSharedStrings(WorkbookPart workbookPart)
        {
            var result = new List<string>();
            var ssPart = workbookPart.SharedStringTablePart;
            if (ssPart == null) return result;

            var ssXDoc = ssPart.GetXDocument();
            foreach (var si in ssXDoc.Root.Elements(S.si))
            {
                var text = si.Descendants(S.t).Select(t => (string)t).StringConcatenate();
                result.Add(text);
            }

            return result;
        }

        private static StyleInfo GetStyleInfo(WorkbookPart workbookPart)
        {
            var info = new StyleInfo();
            var stylesPart = workbookPart.WorkbookStylesPart;
            if (stylesPart == null) return info;

            var styleXDoc = stylesPart.GetXDocument();

            // Get number formats
            var numFmts = styleXDoc.Root.Element(S.numFmts);
            if (numFmts != null)
            {
                foreach (var numFmt in numFmts.Elements(S.numFmt))
                {
                    var id = (int?)numFmt.Attribute("numFmtId") ?? 0;
                    var code = (string)numFmt.Attribute("formatCode");
                    info.NumberFormats[id] = code;
                }
            }

            // Get fonts
            var fonts = styleXDoc.Root.Element(S.fonts);
            if (fonts != null)
            {
                foreach (var font in fonts.Elements(S.font))
                {
                    info.Fonts.Add(ParseFont(font));
                }
            }

            // Get fills
            var fills = styleXDoc.Root.Element(S.fills);
            if (fills != null)
            {
                foreach (var fill in fills.Elements(S.fill))
                {
                    info.Fills.Add(ParseFill(fill));
                }
            }

            // Get borders
            var borders = styleXDoc.Root.Element(S.borders);
            if (borders != null)
            {
                foreach (var border in borders.Elements(S.border))
                {
                    info.Borders.Add(ParseBorder(border));
                }
            }

            // Get cellXfs (cell formats)
            var cellXfs = styleXDoc.Root.Element(S.cellXfs);
            if (cellXfs != null)
            {
                foreach (var xf in cellXfs.Elements(S.xf))
                {
                    info.CellFormats.Add(ParseCellXf(xf));
                }
            }

            return info;
        }

        private static FontInfo ParseFont(XElement font)
        {
            var info = new FontInfo();
            info.Bold = font.Element(S.b) != null;
            info.Italic = font.Element(S.i) != null;
            info.Underline = font.Element(S.u) != null;
            info.Strikethrough = font.Element(S.strike) != null;

            var sz = font.Element(S.sz);
            if (sz != null) info.Size = (double?)sz.Attribute("val");

            var name = font.Element(S.name);
            if (name != null) info.Name = (string)name.Attribute("val");

            var color = font.Element(S.color);
            if (color != null) info.Color = GetColorValue(color);

            return info;
        }

        private static FillInfo ParseFill(XElement fill)
        {
            var info = new FillInfo();
            var patternFill = fill.Element(S.patternFill);
            if (patternFill != null)
            {
                info.Pattern = (string)patternFill.Attribute("patternType");
                var fgColor = patternFill.Element(S.fgColor);
                if (fgColor != null) info.ForegroundColor = GetColorValue(fgColor);
                var bgColor = patternFill.Element(S.bgColor);
                if (bgColor != null) info.BackgroundColor = GetColorValue(bgColor);
            }
            return info;
        }

        private static BorderInfo ParseBorder(XElement border)
        {
            var info = new BorderInfo();

            var left = border.Element(S.left);
            if (left != null)
            {
                info.LeftStyle = (string)left.Attribute("style");
                var color = left.Element(S.color);
                if (color != null) info.LeftColor = GetColorValue(color);
            }

            var right = border.Element(S.right);
            if (right != null)
            {
                info.RightStyle = (string)right.Attribute("style");
                var color = right.Element(S.color);
                if (color != null) info.RightColor = GetColorValue(color);
            }

            var top = border.Element(S.top);
            if (top != null)
            {
                info.TopStyle = (string)top.Attribute("style");
                var color = top.Element(S.color);
                if (color != null) info.TopColor = GetColorValue(color);
            }

            var bottom = border.Element(S.bottom);
            if (bottom != null)
            {
                info.BottomStyle = (string)bottom.Attribute("style");
                var color = bottom.Element(S.color);
                if (color != null) info.BottomColor = GetColorValue(color);
            }

            return info;
        }

        private static CellXfInfo ParseCellXf(XElement xf)
        {
            var info = new CellXfInfo
            {
                NumFmtId = (int?)xf.Attribute("numFmtId") ?? 0,
                FontId = (int?)xf.Attribute("fontId") ?? 0,
                FillId = (int?)xf.Attribute("fillId") ?? 0,
                BorderId = (int?)xf.Attribute("borderId") ?? 0,
                ApplyNumberFormat = (string)xf.Attribute("applyNumberFormat") == "1",
                ApplyFont = (string)xf.Attribute("applyFont") == "1",
                ApplyFill = (string)xf.Attribute("applyFill") == "1",
                ApplyBorder = (string)xf.Attribute("applyBorder") == "1",
                ApplyAlignment = (string)xf.Attribute("applyAlignment") == "1"
            };

            var alignment = xf.Element(S.alignment);
            if (alignment != null)
            {
                info.HorizontalAlignment = (string)alignment.Attribute("horizontal");
                info.VerticalAlignment = (string)alignment.Attribute("vertical");
                info.WrapText = (string)alignment.Attribute("wrapText") == "1";
                info.Indent = (int?)alignment.Attribute("indent");
            }

            return info;
        }

        private static string GetColorValue(XElement colorElement)
        {
            // Try RGB first
            var rgb = (string)colorElement.Attribute("rgb");
            if (!string.IsNullOrEmpty(rgb)) return rgb;

            // Try indexed color
            var indexed = (int?)colorElement.Attribute("indexed");
            if (indexed.HasValue && indexed.Value < SmlDataRetriever.IndexedColors.Length)
            {
                return SmlDataRetriever.IndexedColors[indexed.Value];
            }

            // Try theme color
            var theme = (int?)colorElement.Attribute("theme");
            if (theme.HasValue)
            {
                return $"theme:{theme.Value}";
            }

            return null;
        }

        private static CellFormatSignature ExpandStyle(int styleIndex, StyleInfo styleInfo)
        {
            var format = new CellFormatSignature();

            if (styleIndex < 0 || styleIndex >= styleInfo.CellFormats.Count)
            {
                return CellFormatSignature.Default;
            }

            var xf = styleInfo.CellFormats[styleIndex];

            // Number format
            if (styleInfo.NumberFormats.TryGetValue(xf.NumFmtId, out var numFmt))
            {
                format.NumberFormatCode = numFmt;
            }
            else
            {
                format.NumberFormatCode = GetBuiltInNumberFormat(xf.NumFmtId);
            }

            // Font
            if (xf.FontId >= 0 && xf.FontId < styleInfo.Fonts.Count)
            {
                var font = styleInfo.Fonts[xf.FontId];
                format.Bold = font.Bold;
                format.Italic = font.Italic;
                format.Underline = font.Underline;
                format.Strikethrough = font.Strikethrough;
                format.FontName = font.Name;
                format.FontSize = font.Size;
                format.FontColor = font.Color;
            }

            // Fill
            if (xf.FillId >= 0 && xf.FillId < styleInfo.Fills.Count)
            {
                var fill = styleInfo.Fills[xf.FillId];
                format.FillPattern = fill.Pattern;
                format.FillForegroundColor = fill.ForegroundColor;
                format.FillBackgroundColor = fill.BackgroundColor;
            }

            // Border
            if (xf.BorderId >= 0 && xf.BorderId < styleInfo.Borders.Count)
            {
                var border = styleInfo.Borders[xf.BorderId];
                format.BorderLeftStyle = border.LeftStyle;
                format.BorderLeftColor = border.LeftColor;
                format.BorderRightStyle = border.RightStyle;
                format.BorderRightColor = border.RightColor;
                format.BorderTopStyle = border.TopStyle;
                format.BorderTopColor = border.TopColor;
                format.BorderBottomStyle = border.BottomStyle;
                format.BorderBottomColor = border.BottomColor;
            }

            // Alignment
            format.HorizontalAlignment = xf.HorizontalAlignment;
            format.VerticalAlignment = xf.VerticalAlignment;
            format.WrapText = xf.WrapText;
            format.Indent = xf.Indent;

            return format;
        }

        private static string GetBuiltInNumberFormat(int numFmtId)
        {
            // Built-in number formats per ECMA-376
            return numFmtId switch
            {
                0 => "General",
                1 => "0",
                2 => "0.00",
                3 => "#,##0",
                4 => "#,##0.00",
                9 => "0%",
                10 => "0.00%",
                11 => "0.00E+00",
                12 => "# ?/?",
                13 => "# ??/??",
                14 => "mm-dd-yy",
                15 => "d-mmm-yy",
                16 => "d-mmm",
                17 => "mmm-yy",
                18 => "h:mm AM/PM",
                19 => "h:mm:ss AM/PM",
                20 => "h:mm",
                21 => "h:mm:ss",
                22 => "m/d/yy h:mm",
                37 => "#,##0 ;(#,##0)",
                38 => "#,##0 ;[Red](#,##0)",
                39 => "#,##0.00;(#,##0.00)",
                40 => "#,##0.00;[Red](#,##0.00)",
                45 => "mm:ss",
                46 => "[h]:mm:ss",
                47 => "mmss.0",
                48 => "##0.0E+0",
                49 => "@",
                _ => "General"
            };
        }
    }

    // Internal style info classes
    internal class StyleInfo
    {
        public Dictionary<int, string> NumberFormats { get; } = new Dictionary<int, string>();
        public List<FontInfo> Fonts { get; } = new List<FontInfo>();
        public List<FillInfo> Fills { get; } = new List<FillInfo>();
        public List<BorderInfo> Borders { get; } = new List<BorderInfo>();
        public List<CellXfInfo> CellFormats { get; } = new List<CellXfInfo>();
    }

    internal class FontInfo
    {
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public bool Strikethrough { get; set; }
        public string Name { get; set; }
        public double? Size { get; set; }
        public string Color { get; set; }
    }

    internal class FillInfo
    {
        public string Pattern { get; set; }
        public string ForegroundColor { get; set; }
        public string BackgroundColor { get; set; }
    }

    internal class BorderInfo
    {
        public string LeftStyle { get; set; }
        public string LeftColor { get; set; }
        public string RightStyle { get; set; }
        public string RightColor { get; set; }
        public string TopStyle { get; set; }
        public string TopColor { get; set; }
        public string BottomStyle { get; set; }
        public string BottomColor { get; set; }
    }

    internal class CellXfInfo
    {
        public int NumFmtId { get; set; }
        public int FontId { get; set; }
        public int FillId { get; set; }
        public int BorderId { get; set; }
        public bool ApplyNumberFormat { get; set; }
        public bool ApplyFont { get; set; }
        public bool ApplyFill { get; set; }
        public bool ApplyBorder { get; set; }
        public bool ApplyAlignment { get; set; }
        public string HorizontalAlignment { get; set; }
        public string VerticalAlignment { get; set; }
        public bool WrapText { get; set; }
        public int? Indent { get; set; }
    }

    /// <summary>
    /// Computes the diff between two canonicalized workbooks.
    /// </summary>
    internal static class SmlDiffEngine
    {
        public static SmlComparisonResult ComputeDiff(
            WorkbookSignature sig1,
            WorkbookSignature sig2,
            SmlComparerSettings settings)
        {
            var result = new SmlComparisonResult();

            if (settings.CompareSheetStructure)
            {
                // Sheet-level diff
                var sheets1 = sig1.Sheets.Keys.ToHashSet();
                var sheets2 = sig2.Sheets.Keys.ToHashSet();

                foreach (var added in sheets2.Except(sheets1))
                {
                    result.Changes.Add(new SmlChange
                    {
                        ChangeType = SmlChangeType.SheetAdded,
                        SheetName = added
                    });
                }

                foreach (var deleted in sheets1.Except(sheets2))
                {
                    result.Changes.Add(new SmlChange
                    {
                        ChangeType = SmlChangeType.SheetDeleted,
                        SheetName = deleted
                    });
                }
            }

            // For matched sheets, compare cells
            var commonSheets = sig1.Sheets.Keys.Intersect(sig2.Sheets.Keys);
            foreach (var sheetName in commonSheets)
            {
                var ws1 = sig1.Sheets[sheetName];
                var ws2 = sig2.Sheets[sheetName];

                CompareWorksheets(ws1, ws2, sheetName, settings, result);
            }

            return result;
        }

        private static void CompareWorksheets(
            WorksheetSignature ws1,
            WorksheetSignature ws2,
            string sheetName,
            SmlComparerSettings settings,
            SmlComparisonResult result)
        {
            // Get union of all cell addresses
            var allAddresses = ws1.Cells.Keys.Union(ws2.Cells.Keys);

            foreach (var addr in allAddresses)
            {
                var has1 = ws1.Cells.TryGetValue(addr, out var cell1);
                var has2 = ws2.Cells.TryGetValue(addr, out var cell2);

                if (!has1 && has2)
                {
                    // Cell added
                    result.Changes.Add(new SmlChange
                    {
                        ChangeType = SmlChangeType.CellAdded,
                        SheetName = sheetName,
                        CellAddress = addr,
                        NewValue = cell2.ResolvedValue,
                        NewFormula = cell2.Formula,
                        NewFormat = cell2.Format
                    });
                }
                else if (has1 && !has2)
                {
                    // Cell deleted
                    result.Changes.Add(new SmlChange
                    {
                        ChangeType = SmlChangeType.CellDeleted,
                        SheetName = sheetName,
                        CellAddress = addr,
                        OldValue = cell1.ResolvedValue,
                        OldFormula = cell1.Formula,
                        OldFormat = cell1.Format
                    });
                }
                else if (has1 && has2)
                {
                    // Compare cells
                    CompareCells(cell1, cell2, sheetName, settings, result);
                }
            }
        }

        private static void CompareCells(
            CellSignature cell1,
            CellSignature cell2,
            string sheetName,
            SmlComparerSettings settings,
            SmlComparisonResult result)
        {
            // Quick check via content hash (value + formula)
            if (cell1.ContentHash == cell2.ContentHash &&
                (!settings.CompareFormatting || Equals(cell1.Format, cell2.Format)))
            {
                return; // No changes
            }

            // Check value change
            if (settings.CompareValues)
            {
                var val1 = cell1.ResolvedValue ?? "";
                var val2 = cell2.ResolvedValue ?? "";

                bool valuesEqual;
                if (settings.CaseInsensitiveValues)
                {
                    valuesEqual = string.Equals(val1, val2, StringComparison.OrdinalIgnoreCase);
                }
                else if (settings.NumericTolerance > 0 &&
                         double.TryParse(val1, NumberStyles.Float, CultureInfo.InvariantCulture, out var d1) &&
                         double.TryParse(val2, NumberStyles.Float, CultureInfo.InvariantCulture, out var d2))
                {
                    valuesEqual = Math.Abs(d1 - d2) <= settings.NumericTolerance;
                }
                else
                {
                    valuesEqual = val1 == val2;
                }

                if (!valuesEqual)
                {
                    result.Changes.Add(new SmlChange
                    {
                        ChangeType = SmlChangeType.ValueChanged,
                        SheetName = sheetName,
                        CellAddress = cell1.Address,
                        OldValue = cell1.ResolvedValue,
                        NewValue = cell2.ResolvedValue,
                        OldFormula = cell1.Formula,
                        NewFormula = cell2.Formula
                    });
                    return; // Don't report formula change if value changed
                }
            }

            // Check formula change
            if (settings.CompareFormulas)
            {
                var formula1 = cell1.Formula ?? "";
                var formula2 = cell2.Formula ?? "";

                if (formula1 != formula2)
                {
                    result.Changes.Add(new SmlChange
                    {
                        ChangeType = SmlChangeType.FormulaChanged,
                        SheetName = sheetName,
                        CellAddress = cell1.Address,
                        OldFormula = cell1.Formula,
                        NewFormula = cell2.Formula,
                        OldValue = cell1.ResolvedValue,
                        NewValue = cell2.ResolvedValue
                    });
                    return; // Don't report format change if formula changed
                }
            }

            // Check format change
            if (settings.CompareFormatting && !Equals(cell1.Format, cell2.Format))
            {
                result.Changes.Add(new SmlChange
                {
                    ChangeType = SmlChangeType.FormatChanged,
                    SheetName = sheetName,
                    CellAddress = cell1.Address,
                    OldFormat = cell1.Format,
                    NewFormat = cell2.Format,
                    OldValue = cell1.ResolvedValue,
                    NewValue = cell2.ResolvedValue
                });
            }
        }
    }

    /// <summary>
    /// Renders a marked workbook showing differences.
    /// </summary>
    internal static class SmlMarkupRenderer
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
                var stylesPart = workbookPart.WorkbookStylesPart;
                var styleXDoc = stylesPart.GetXDocument();

                // Add highlight fill styles
                var highlightStyles = AddHighlightStyles(styleXDoc, settings);
                stylesPart.PutXDocument();

                // Group changes by sheet
                var changesBySheet = result.Changes
                    .Where(c => c.CellAddress != null)
                    .GroupBy(c => c.SheetName);

                foreach (var sheetGroup in changesBySheet)
                {
                    var sheetName = sheetGroup.Key;
                    var worksheetPart = GetWorksheetPart(workbookPart, sheetName);
                    if (worksheetPart == null) continue;

                    var wsXDoc = worksheetPart.GetXDocument();

                    foreach (var change in sheetGroup)
                    {
                        ApplyCellHighlight(wsXDoc, change, highlightStyles);
                    }

                    worksheetPart.PutXDocument();

                    // Add comments for changes
                    AddCommentsForChanges(worksheetPart, sheetGroup.ToList(), settings);
                }

                // Add summary sheet
                AddDiffSummarySheet(sDoc, result, settings);
            }

            return new SmlDocument("compared.xlsx", ms.ToArray());
        }

        private static HighlightStyles AddHighlightStyles(XDocument styleXDoc, SmlComparerSettings settings)
        {
            var styles = new HighlightStyles();

            // Find or create fills element
            var fills = styleXDoc.Root.Element(S.fills);
            if (fills == null)
            {
                fills = new XElement(S.fills, new XAttribute("count", "0"));
                var fonts = styleXDoc.Root.Element(S.fonts);
                if (fonts != null)
                    fonts.AddAfterSelf(fills);
                else
                    styleXDoc.Root.AddFirst(fills);
            }

            var fillCount = (int?)fills.Attribute("count") ?? fills.Elements(S.fill).Count();

            // Add highlight fills
            styles.AddedFillId = fillCount++;
            fills.Add(CreateSolidFill(settings.AddedCellColor));

            styles.ModifiedValueFillId = fillCount++;
            fills.Add(CreateSolidFill(settings.ModifiedValueColor));

            styles.ModifiedFormulaFillId = fillCount++;
            fills.Add(CreateSolidFill(settings.ModifiedFormulaColor));

            styles.ModifiedFormatFillId = fillCount++;
            fills.Add(CreateSolidFill(settings.ModifiedFormatColor));

            fills.SetAttributeValue("count", fillCount);

            // Find or create cellXfs element
            var cellXfs = styleXDoc.Root.Element(S.cellXfs);
            if (cellXfs == null)
            {
                cellXfs = new XElement(S.cellXfs, new XAttribute("count", "0"));
                var cellStyleXfs = styleXDoc.Root.Element(S.cellStyleXfs);
                if (cellStyleXfs != null)
                    cellStyleXfs.AddAfterSelf(cellXfs);
                else
                    styleXDoc.Root.Add(cellXfs);
            }

            var xfCount = (int?)cellXfs.Attribute("count") ?? cellXfs.Elements(S.xf).Count();

            // Add cell formats that use the highlight fills
            styles.AddedStyleId = xfCount++;
            cellXfs.Add(CreateXfWithFill(styles.AddedFillId));

            styles.ModifiedValueStyleId = xfCount++;
            cellXfs.Add(CreateXfWithFill(styles.ModifiedValueFillId));

            styles.ModifiedFormulaStyleId = xfCount++;
            cellXfs.Add(CreateXfWithFill(styles.ModifiedFormulaFillId));

            styles.ModifiedFormatStyleId = xfCount++;
            cellXfs.Add(CreateXfWithFill(styles.ModifiedFormatFillId));

            cellXfs.SetAttributeValue("count", xfCount);

            return styles;
        }

        private static XElement CreateSolidFill(string color)
        {
            return new XElement(S.fill,
                new XElement(S.patternFill,
                    new XAttribute("patternType", "solid"),
                    new XElement(S.fgColor, new XAttribute("rgb", "FF" + color)),
                    new XElement(S.bgColor, new XAttribute("indexed", "64"))));
        }

        private static XElement CreateXfWithFill(int fillId)
        {
            return new XElement(S.xf,
                new XAttribute("numFmtId", "0"),
                new XAttribute("fontId", "0"),
                new XAttribute("fillId", fillId),
                new XAttribute("borderId", "0"),
                new XAttribute("applyFill", "1"));
        }

        private static void ApplyCellHighlight(XDocument wsXDoc, SmlChange change, HighlightStyles styles)
        {
            var sheetData = wsXDoc.Root.Element(S.sheetData);
            if (sheetData == null) return;

            // Find or create the cell
            var (colIndex, rowIndex) = ParseCellRef(change.CellAddress);
            var row = sheetData.Elements(S.row)
                .FirstOrDefault(r => (int?)r.Attribute("r") == rowIndex);

            if (row == null)
            {
                // Create row if needed
                row = new XElement(S.row, new XAttribute("r", rowIndex));
                // Insert in correct position
                var insertAfter = sheetData.Elements(S.row)
                    .Where(r => (int?)r.Attribute("r") < rowIndex)
                    .LastOrDefault();
                if (insertAfter != null)
                    insertAfter.AddAfterSelf(row);
                else
                    sheetData.AddFirst(row);
            }

            var cell = row.Elements(S.c)
                .FirstOrDefault(c => (string)c.Attribute("r") == change.CellAddress);

            if (cell == null)
            {
                // Create cell if needed
                cell = new XElement(S.c, new XAttribute("r", change.CellAddress));
                // Insert in correct position
                var insertAfter = row.Elements(S.c)
                    .Where(c => GetColumnIndex((string)c.Attribute("r")) < colIndex)
                    .LastOrDefault();
                if (insertAfter != null)
                    insertAfter.AddAfterSelf(cell);
                else
                    row.AddFirst(cell);
            }

            // Apply style based on change type
            var styleId = change.ChangeType switch
            {
                SmlChangeType.CellAdded => styles.AddedStyleId,
                SmlChangeType.ValueChanged => styles.ModifiedValueStyleId,
                SmlChangeType.FormulaChanged => styles.ModifiedFormulaStyleId,
                SmlChangeType.FormatChanged => styles.ModifiedFormatStyleId,
                _ => -1
            };

            if (styleId >= 0)
            {
                cell.SetAttributeValue("s", styleId);
            }
        }

        private static (int col, int row) ParseCellRef(string cellRef)
        {
            var col = 0;
            var i = 0;

            while (i < cellRef.Length && char.IsLetter(cellRef[i]))
            {
                col = col * 26 + (char.ToUpper(cellRef[i]) - 'A' + 1);
                i++;
            }

            int.TryParse(cellRef.Substring(i), out var row);
            return (col, row);
        }

        private static int GetColumnIndex(string cellRef)
        {
            var col = 0;
            foreach (var c in cellRef)
            {
                if (!char.IsLetter(c)) break;
                col = col * 26 + (char.ToUpper(c) - 'A' + 1);
            }
            return col;
        }

        private static void AddCommentsForChanges(
            WorksheetPart worksheetPart,
            List<SmlChange> changes,
            SmlComparerSettings settings)
        {
            if (!changes.Any()) return;

            // Get or create comments part
            var commentsPart = worksheetPart.WorksheetCommentsPart;
            XDocument commentsXDoc;

            if (commentsPart == null)
            {
                commentsPart = worksheetPart.AddNewPart<WorksheetCommentsPart>();
                commentsXDoc = new XDocument(
                    new XElement(S.comments,
                        new XAttribute(XNamespace.Xmlns + "x", S.s.NamespaceName),
                        new XElement(S.authors,
                            new XElement(S.author, settings.AuthorForChanges)),
                        new XElement(S.commentList)));
            }
            else
            {
                commentsXDoc = commentsPart.GetXDocument();
            }

            var commentList = commentsXDoc.Root.Element(S.commentList);

            foreach (var change in changes)
            {
                var commentText = BuildCommentText(change);

                var comment = new XElement(S.comment,
                    new XAttribute("ref", change.CellAddress),
                    new XAttribute("authorId", "0"),
                    new XElement(S.text,
                        new XElement(S.r,
                            new XElement(S.t, commentText))));

                commentList.Add(comment);
            }

            commentsPart.PutXDocument();

            // Add VML drawing part for comment display (required for comments to show)
            AddVmlDrawingForComments(worksheetPart, changes);
        }

        private static string BuildCommentText(SmlChange change)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"[{change.ChangeType}]");

            switch (change.ChangeType)
            {
                case SmlChangeType.CellAdded:
                    sb.AppendLine($"New value: {change.NewValue}");
                    if (!string.IsNullOrEmpty(change.NewFormula))
                        sb.AppendLine($"Formula: ={change.NewFormula}");
                    break;

                case SmlChangeType.ValueChanged:
                    sb.AppendLine($"Old value: {change.OldValue}");
                    sb.AppendLine($"New value: {change.NewValue}");
                    break;

                case SmlChangeType.FormulaChanged:
                    sb.AppendLine($"Old formula: ={change.OldFormula}");
                    sb.AppendLine($"New formula: ={change.NewFormula}");
                    break;

                case SmlChangeType.FormatChanged:
                    if (change.NewFormat != null && change.OldFormat != null)
                    {
                        sb.AppendLine(change.NewFormat.GetDifferenceDescription(change.OldFormat));
                    }
                    break;
            }

            return sb.ToString().TrimEnd();
        }

        private static void AddVmlDrawingForComments(WorksheetPart worksheetPart, List<SmlChange> changes)
        {
            // VML is required for Excel to display comments
            var vmlPart = worksheetPart.VmlDrawingParts.FirstOrDefault();
            if (vmlPart == null)
            {
                vmlPart = worksheetPart.AddNewPart<VmlDrawingPart>();

                // Add relationship to worksheet
                var wsXDoc = worksheetPart.GetXDocument();
                var legacyDrawing = wsXDoc.Root.Element(S.s + "legacyDrawing");
                if (legacyDrawing == null)
                {
                    var rId = worksheetPart.GetIdOfPart(vmlPart);
                    legacyDrawing = new XElement(S.s + "legacyDrawing",
                        new XAttribute(R.id, rId));
                    wsXDoc.Root.Add(legacyDrawing);
                    worksheetPart.PutXDocument();
                }
            }

            // Build VML content
            var vmlBuilder = new StringBuilder();
            vmlBuilder.AppendLine("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\">");
            vmlBuilder.AppendLine("<o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"1\"/></o:shapelayout>");
            vmlBuilder.AppendLine("<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\">");
            vmlBuilder.AppendLine("<v:stroke joinstyle=\"miter\"/><v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/>");
            vmlBuilder.AppendLine("</v:shapetype>");

            var shapeId = 1024;
            foreach (var change in changes)
            {
                var (col, row) = ParseCellRef(change.CellAddress);
                vmlBuilder.AppendLine($"<v:shape id=\"_x0000_s{shapeId++}\" type=\"#_x0000_t202\" style=\"position:absolute;margin-left:80pt;margin-top:5pt;width:120pt;height:60pt;z-index:1;visibility:hidden\" fillcolor=\"#ffffe1\" o:insetmode=\"auto\">");
                vmlBuilder.AppendLine("<v:fill color2=\"#ffffe1\"/>");
                vmlBuilder.AppendLine("<v:shadow on=\"t\" color=\"black\" obscured=\"t\"/>");
                vmlBuilder.AppendLine("<v:path o:connecttype=\"none\"/>");
                vmlBuilder.AppendLine("<v:textbox style=\"mso-direction-alt:auto\"/>");
                vmlBuilder.AppendLine($"<x:ClientData ObjectType=\"Note\"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>{col - 1}, 0, {row - 1}, 0, {col + 1}, 0, {row + 3}, 0</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>{row - 1}</x:Row><x:Column>{col - 1}</x:Column></x:ClientData>");
                vmlBuilder.AppendLine("</v:shape>");
            }

            vmlBuilder.AppendLine("</xml>");

            using var stream = vmlPart.GetStream(FileMode.Create);
            using var writer = new StreamWriter(stream);
            writer.Write(vmlBuilder.ToString());
        }

        private static WorksheetPart GetWorksheetPart(WorkbookPart workbookPart, string sheetName)
        {
            var wbXDoc = workbookPart.GetXDocument();
            var sheet = wbXDoc.Root.Elements(S.sheets).Elements(S.sheet)
                .FirstOrDefault(s => (string)s.Attribute("name") == sheetName);

            if (sheet == null) return null;

            var rId = (string)sheet.Attribute(R.id);
            return (WorksheetPart)workbookPart.GetPartById(rId);
        }

        private static void AddDiffSummarySheet(
            SpreadsheetDocument sDoc,
            SmlComparisonResult result,
            SmlComparerSettings settings)
        {
            var workbookPart = sDoc.WorkbookPart;
            var wbXDoc = workbookPart.GetXDocument();

            // Create a new worksheet part
            var newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();

            // Build the worksheet content
            var wsXDoc = new XDocument(
                new XElement(S.worksheet,
                    new XAttribute(XNamespace.Xmlns + "x", S.s.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "r", R.r.NamespaceName),
                    new XElement(S.sheetData)));

            var sheetData = wsXDoc.Root.Element(S.sheetData);

            // Add summary header
            var rowNum = 1;
            AddRow(sheetData, rowNum++, new[] { "Spreadsheet Comparison Summary" });
            AddRow(sheetData, rowNum++, new[] { "" });
            AddRow(sheetData, rowNum++, new[] { "Total Changes:", result.TotalChanges.ToString() });
            AddRow(sheetData, rowNum++, new[] { "Value Changes:", result.ValueChanges.ToString() });
            AddRow(sheetData, rowNum++, new[] { "Formula Changes:", result.FormulaChanges.ToString() });
            AddRow(sheetData, rowNum++, new[] { "Format Changes:", result.FormatChanges.ToString() });
            AddRow(sheetData, rowNum++, new[] { "Cells Added:", result.CellsAdded.ToString() });
            AddRow(sheetData, rowNum++, new[] { "Cells Deleted:", result.CellsDeleted.ToString() });
            AddRow(sheetData, rowNum++, new[] { "Sheets Added:", result.SheetsAdded.ToString() });
            AddRow(sheetData, rowNum++, new[] { "Sheets Deleted:", result.SheetsDeleted.ToString() });
            AddRow(sheetData, rowNum++, new[] { "" });

            // Add change details header
            AddRow(sheetData, rowNum++, new[] { "Change Type", "Sheet", "Cell", "Old Value", "New Value", "Description" });

            // Add each change
            foreach (var change in result.Changes)
            {
                AddRow(sheetData, rowNum++, new[]
                {
                    change.ChangeType.ToString(),
                    change.SheetName ?? "",
                    change.CellAddress ?? "",
                    change.OldValue ?? change.OldFormula ?? "",
                    change.NewValue ?? change.NewFormula ?? "",
                    change.GetDescription()
                });
            }

            newWorksheetPart.PutXDocument(wsXDoc);

            // Add sheet to workbook
            var sheets = wbXDoc.Root.Element(S.sheets);
            var newSheetId = sheets.Elements(S.sheet)
                .Select(s => (uint?)s.Attribute("sheetId") ?? 0)
                .DefaultIfEmpty(0u)
                .Max() + 1;

            var rId = workbookPart.GetIdOfPart(newWorksheetPart);

            sheets.Add(new XElement(S.sheet,
                new XAttribute("name", "_DiffSummary"),
                new XAttribute("sheetId", newSheetId),
                new XAttribute(R.id, rId)));

            workbookPart.PutXDocument();
        }

        private static void AddRow(XElement sheetData, int rowNum, string[] values)
        {
            var row = new XElement(S.row, new XAttribute("r", rowNum));

            for (int i = 0; i < values.Length; i++)
            {
                var colLetter = GetColumnLetter(i + 1);
                var cellRef = $"{colLetter}{rowNum}";

                var cell = new XElement(S.c,
                    new XAttribute("r", cellRef),
                    new XAttribute("t", "inlineStr"),
                    new XElement(S._is,
                        new XElement(S.t, values[i])));

                row.Add(cell);
            }

            sheetData.Add(row);
        }

        private static string GetColumnLetter(int columnNumber)
        {
            var result = "";
            while (columnNumber > 0)
            {
                columnNumber--;
                result = (char)('A' + columnNumber % 26) + result;
                columnNumber /= 26;
            }
            return result;
        }
    }

    internal class HighlightStyles
    {
        public int AddedFillId { get; set; }
        public int ModifiedValueFillId { get; set; }
        public int ModifiedFormulaFillId { get; set; }
        public int ModifiedFormatFillId { get; set; }

        public int AddedStyleId { get; set; }
        public int ModifiedValueStyleId { get; set; }
        public int ModifiedFormulaStyleId { get; set; }
        public int ModifiedFormatStyleId { get; set; }
    }

    #endregion
}
