// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace OpenXmlPowerTools
{
    #region Settings

    /// <summary>
    /// Settings for controlling PowerPoint presentation comparison behavior.
    /// </summary>
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
        public bool CompareShapeStyles { get; set; } = false;

        /// <summary>Compare images by content hash.</summary>
        public bool CompareImageContent { get; set; } = true;

        /// <summary>Compare chart data.</summary>
        public bool CompareCharts { get; set; } = true;

        /// <summary>Compare tables.</summary>
        public bool CompareTables { get; set; } = true;

        /// <summary>Compare slide notes.</summary>
        public bool CompareNotes { get; set; } = false;

        /// <summary>Compare slide transitions.</summary>
        public bool CompareTransitions { get; set; } = false;

        // === Matching Settings ===

        /// <summary>Enable fuzzy shape matching when exact matches fail.</summary>
        public bool EnableFuzzyShapeMatching { get; set; } = true;

        /// <summary>Minimum similarity score (0.0-1.0) for fuzzy slide matching.</summary>
        public double SlideSimilarityThreshold { get; set; } = 0.6;

        /// <summary>Minimum similarity score (0.0-1.0) for fuzzy shape matching.</summary>
        public double ShapeSimilarityThreshold { get; set; } = 0.7;

        /// <summary>Position tolerance in EMUs for "same location" matching (default ~0.1 inch).</summary>
        public long PositionTolerance { get; set; } = 91440;

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

        /// <summary>Color for inserted elements.</summary>
        public string InsertedColor { get; set; } = "00AA00";

        /// <summary>Color for deleted elements.</summary>
        public string DeletedColor { get; set; } = "FF0000";

        /// <summary>Color for modified elements.</summary>
        public string ModifiedColor { get; set; } = "FFA500";

        /// <summary>Color for moved elements.</summary>
        public string MovedColor { get; set; } = "0000FF";

        /// <summary>Color for formatting-only changes.</summary>
        public string FormattingColor { get; set; } = "9932CC";

        // === Logging ===

        /// <summary>Optional callback for logging/debugging.</summary>
        public Action<string> LogCallback { get; set; }
    }

    #endregion

    #region Change Types

    /// <summary>
    /// Types of changes detected during presentation comparison.
    /// </summary>
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
        ShapeTypeChanged,

        // Shape content
        TextChanged,
        TextFormattingChanged,
        ImageReplaced,
        TableContentChanged,
        TableStructureChanged,
        ChartDataChanged,
        ChartFormatChanged,

        // Shape style
        ShapeFillChanged,
        ShapeLineChanged,
        ShapeEffectsChanged,

        // Group-specific
        GroupMembershipChanged,
    }

    /// <summary>
    /// Types of text changes within a shape.
    /// </summary>
    public enum TextChangeType
    {
        Insert,
        Delete,
        Replace,
        FormatOnly
    }

    /// <summary>
    /// Represents a single text change within a shape.
    /// </summary>
    public class PmlTextChange
    {
        public TextChangeType Type { get; set; }
        public int ParagraphIndex { get; set; }
        public int RunIndex { get; set; }
        public string OldText { get; set; }
        public string NewText { get; set; }
    }

    /// <summary>
    /// Represents a single change between two presentations.
    /// </summary>
    public class PmlChange
    {
        public PmlChangeType ChangeType { get; set; }

        // Location
        public int? SlideIndex { get; set; }
        public int? OldSlideIndex { get; set; }
        public string ShapeName { get; set; }
        public string ShapeId { get; set; }

        // Details
        public string OldValue { get; set; }
        public string NewValue { get; set; }
        public long? OldX { get; set; }
        public long? OldY { get; set; }
        public long? OldCx { get; set; }
        public long? OldCy { get; set; }
        public long? NewX { get; set; }
        public long? NewY { get; set; }
        public long? NewCx { get; set; }
        public long? NewCy { get; set; }

        // Text changes
        public List<PmlTextChange> TextChanges { get; set; }

        // Matching info
        public double MatchConfidence { get; set; }

        /// <summary>
        /// Returns a human-readable description of this change.
        /// </summary>
        public string GetDescription()
        {
            return ChangeType switch
            {
                PmlChangeType.SlideInserted => $"Slide {SlideIndex} inserted",
                PmlChangeType.SlideDeleted => $"Slide {OldSlideIndex} deleted",
                PmlChangeType.SlideMoved => $"Slide moved from position {OldSlideIndex} to {SlideIndex}",
                PmlChangeType.SlideLayoutChanged => $"Slide {SlideIndex} layout changed",
                PmlChangeType.SlideBackgroundChanged => $"Slide {SlideIndex} background changed",
                PmlChangeType.SlideNotesChanged => $"Slide {SlideIndex} notes changed",
                PmlChangeType.ShapeInserted => $"Shape '{ShapeName}' inserted on slide {SlideIndex}",
                PmlChangeType.ShapeDeleted => $"Shape '{ShapeName}' deleted from slide {SlideIndex}",
                PmlChangeType.ShapeMoved => $"Shape '{ShapeName}' moved on slide {SlideIndex}",
                PmlChangeType.ShapeResized => $"Shape '{ShapeName}' resized on slide {SlideIndex}",
                PmlChangeType.ShapeRotated => $"Shape '{ShapeName}' rotated on slide {SlideIndex}",
                PmlChangeType.ShapeZOrderChanged => $"Shape '{ShapeName}' z-order changed on slide {SlideIndex}",
                PmlChangeType.TextChanged => $"Text changed in '{ShapeName}' on slide {SlideIndex}",
                PmlChangeType.TextFormattingChanged => $"Text formatting changed in '{ShapeName}' on slide {SlideIndex}",
                PmlChangeType.ImageReplaced => $"Image replaced in '{ShapeName}' on slide {SlideIndex}",
                PmlChangeType.TableContentChanged => $"Table content changed in '{ShapeName}' on slide {SlideIndex}",
                PmlChangeType.ChartDataChanged => $"Chart data changed in '{ShapeName}' on slide {SlideIndex}",
                _ => $"{ChangeType} on slide {SlideIndex}"
            };
        }
    }

    /// <summary>
    /// Result of comparing two presentations, containing all detected changes.
    /// </summary>
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
        public int ShapesResized => Changes.Count(c => c.ChangeType == PmlChangeType.ShapeResized);
        public int TextChanges => Changes.Count(c => c.ChangeType == PmlChangeType.TextChanged);
        public int FormattingChanges => Changes.Count(c => c.ChangeType == PmlChangeType.TextFormattingChanged);
        public int ImagesReplaced => Changes.Count(c => c.ChangeType == PmlChangeType.ImageReplaced);

        /// <summary>
        /// Get all changes for a specific slide.
        /// </summary>
        public IEnumerable<PmlChange> GetChangesBySlide(int slideIndex)
            => Changes.Where(c => c.SlideIndex == slideIndex);

        /// <summary>
        /// Get all changes of a specific type.
        /// </summary>
        public IEnumerable<PmlChange> GetChangesByType(PmlChangeType type)
            => Changes.Where(c => c.ChangeType == type);

        /// <summary>
        /// Get all changes for a specific shape.
        /// </summary>
        public IEnumerable<PmlChange> GetChangesByShape(string shapeName)
            => Changes.Where(c => c.ShapeName == shapeName);

        /// <summary>
        /// Export the comparison result to JSON.
        /// </summary>
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
                    ShapesResized,
                    TextChanges,
                    FormattingChanges,
                    ImagesReplaced
                },
                Changes = Changes.Select(c => new
                {
                    c.ChangeType,
                    c.SlideIndex,
                    c.OldSlideIndex,
                    c.ShapeName,
                    c.OldValue,
                    c.NewValue,
                    Description = c.GetDescription()
                })
            }, options);
        }
    }

    #endregion

    #region Signature Classes

    /// <summary>
    /// Shape types for comparison.
    /// </summary>
    public enum PmlShapeType
    {
        Unknown,
        TextBox,
        AutoShape,
        Picture,
        Table,
        Chart,
        SmartArt,
        Group,
        Connector,
        OleObject,
        Media
    }

    /// <summary>
    /// Placeholder information for a shape.
    /// </summary>
    public class PlaceholderInfo
    {
        public string Type { get; set; }
        public uint? Index { get; set; }

        public override bool Equals(object obj)
        {
            if (obj is PlaceholderInfo other)
                return Type == other.Type && Index == other.Index;
            return false;
        }

        public override int GetHashCode() => HashCode.Combine(Type, Index);
    }

    /// <summary>
    /// Transform (position and size) information for a shape.
    /// </summary>
    public class TransformSignature
    {
        public long X { get; set; }
        public long Y { get; set; }
        public long Cx { get; set; }
        public long Cy { get; set; }
        public int Rotation { get; set; }
        public bool FlipH { get; set; }
        public bool FlipV { get; set; }

        public bool IsNear(TransformSignature other, long tolerance)
        {
            if (other == null) return false;
            return Math.Abs(X - other.X) <= tolerance &&
                   Math.Abs(Y - other.Y) <= tolerance;
        }

        public bool IsSameSize(TransformSignature other, long tolerance)
        {
            if (other == null) return false;
            return Math.Abs(Cx - other.Cx) <= tolerance &&
                   Math.Abs(Cy - other.Cy) <= tolerance;
        }
    }

    /// <summary>
    /// Run (text span) properties for comparison.
    /// </summary>
    public class RunPropertiesSignature
    {
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public bool Strikethrough { get; set; }
        public string FontName { get; set; }
        public int? FontSize { get; set; }
        public string FontColor { get; set; }

        public override bool Equals(object obj)
        {
            if (obj is RunPropertiesSignature other)
            {
                return Bold == other.Bold &&
                       Italic == other.Italic &&
                       Underline == other.Underline &&
                       Strikethrough == other.Strikethrough &&
                       FontName == other.FontName &&
                       FontSize == other.FontSize &&
                       FontColor == other.FontColor;
            }
            return false;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(Bold, Italic, Underline, Strikethrough, FontName, FontSize, FontColor);
        }
    }

    /// <summary>
    /// Run (text span) signature for comparison.
    /// </summary>
    internal class RunSignature
    {
        public string Text { get; set; }
        public RunPropertiesSignature Properties { get; set; }
        public string ContentHash { get; set; }
    }

    /// <summary>
    /// Paragraph signature for comparison.
    /// </summary>
    internal class ParagraphSignature
    {
        public List<RunSignature> Runs { get; } = new List<RunSignature>();
        public string PlainText { get; set; }
        public string Alignment { get; set; }
        public bool HasBullet { get; set; }
    }

    /// <summary>
    /// Text body signature for comparison.
    /// </summary>
    internal class TextBodySignature
    {
        public List<ParagraphSignature> Paragraphs { get; } = new List<ParagraphSignature>();
        public string PlainText { get; set; }
    }

    /// <summary>
    /// Shape signature for comparison.
    /// </summary>
    internal class ShapeSignature
    {
        public string Name { get; set; }
        public uint Id { get; set; }
        public PmlShapeType Type { get; set; }
        public PlaceholderInfo Placeholder { get; set; }
        public TransformSignature Transform { get; set; }
        public int ZOrder { get; set; }
        public string GeometryHash { get; set; }
        public TextBodySignature TextBody { get; set; }
        public string ImageHash { get; set; }
        public string TableHash { get; set; }
        public string ChartHash { get; set; }
        public List<ShapeSignature> Children { get; set; }
        public string ContentHash { get; set; }
    }

    /// <summary>
    /// Slide signature for comparison.
    /// </summary>
    internal class SlideSignature
    {
        public int Index { get; set; }
        public string RelationshipId { get; set; }
        public string LayoutRelationshipId { get; set; }
        public string LayoutHash { get; set; }
        public List<ShapeSignature> Shapes { get; } = new List<ShapeSignature>();
        public string NotesText { get; set; }
        public string TitleText { get; set; }
        public string ContentHash { get; set; }
        public string BackgroundHash { get; set; }

        public string ComputeFingerprint()
        {
            var sb = new StringBuilder();
            sb.Append(TitleText ?? "");
            sb.Append("|");
            foreach (var shape in Shapes.OrderBy(s => s.ZOrder))
            {
                sb.Append(shape.Name ?? "");
                sb.Append(":");
                sb.Append(shape.Type);
                sb.Append(":");
                sb.Append(shape.TextBody?.PlainText ?? "");
                sb.Append("|");
            }
            return PmlHasher.ComputeHash(sb.ToString());
        }
    }

    /// <summary>
    /// Presentation signature for comparison.
    /// </summary>
    internal class PresentationSignature
    {
        public long SlideCx { get; set; }
        public long SlideCy { get; set; }
        public List<SlideSignature> Slides { get; } = new List<SlideSignature>();
        public string ThemeHash { get; set; }
    }

    #endregion

    #region Matching Classes

    internal enum SlideMatchType
    {
        Matched,
        Inserted,
        Deleted
    }

    internal class SlideMatch
    {
        public SlideMatchType MatchType { get; set; }
        public int? OldIndex { get; set; }
        public int? NewIndex { get; set; }
        public SlideSignature OldSlide { get; set; }
        public SlideSignature NewSlide { get; set; }
        public double Similarity { get; set; }
        public bool WasMoved => OldIndex.HasValue && NewIndex.HasValue && OldIndex != NewIndex;
    }

    internal enum ShapeMatchType
    {
        Matched,
        Inserted,
        Deleted
    }

    internal enum ShapeMatchMethod
    {
        Placeholder,
        NameAndType,
        NameOnly,
        Fuzzy
    }

    internal class ShapeMatch
    {
        public ShapeMatchType MatchType { get; set; }
        public ShapeSignature OldShape { get; set; }
        public ShapeSignature NewShape { get; set; }
        public double Score { get; set; }
        public ShapeMatchMethod Method { get; set; }
    }

    #endregion

    #region Utility Classes

    internal static class PmlHasher
    {
        public static string ComputeHash(string content)
        {
            if (string.IsNullOrEmpty(content))
                return "";
            var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(content));
            return Convert.ToBase64String(bytes);
        }

        public static string ComputeHash(Stream stream)
        {
            stream.Position = 0;
            var bytes = SHA256.HashData(stream);
            return Convert.ToBase64String(bytes);
        }
    }

    #endregion

    #region Canonicalizer

    /// <summary>
    /// Extracts semantic signatures from presentations.
    /// </summary>
    internal static class PmlCanonicalizer
    {
        public static PresentationSignature Canonicalize(PmlDocument doc, PmlComparerSettings settings)
        {
            var signature = new PresentationSignature();

            using var ms = new MemoryStream();
            ms.Write(doc.DocumentByteArray, 0, doc.DocumentByteArray.Length);
            ms.Position = 0;

            using var pDoc = PresentationDocument.Open(ms, false);
            var presentationPart = pDoc.PresentationPart;
            if (presentationPart == null)
                return signature;

            var presentationXDoc = presentationPart.GetXDocument();
            var presentationRoot = presentationXDoc.Root;

            // Get slide size
            var sldSz = presentationRoot.Element(P.sldSz);
            if (sldSz != null)
            {
                signature.SlideCx = (long?)sldSz.Attribute("cx") ?? 0;
                signature.SlideCy = (long?)sldSz.Attribute("cy") ?? 0;
            }

            // Get slide list
            var sldIdLst = presentationRoot.Element(P.sldIdLst);
            if (sldIdLst == null)
                return signature;

            var slideIds = sldIdLst.Elements(P.sldId).ToList();
            int slideIndex = 1;

            foreach (var sldId in slideIds)
            {
                var rId = (string)sldId.Attribute(R.id);
                if (string.IsNullOrEmpty(rId))
                    continue;

                try
                {
                    var slidePart = (SlidePart)presentationPart.GetPartById(rId);
                    var slideSignature = CanonicalizeSlide(slidePart, slideIndex, rId, settings);
                    signature.Slides.Add(slideSignature);
                }
                catch
                {
                    // Skip invalid slide references
                }

                slideIndex++;
            }

            return signature;
        }

        private static SlideSignature CanonicalizeSlide(
            SlidePart slidePart,
            int index,
            string rId,
            PmlComparerSettings settings)
        {
            var signature = new SlideSignature
            {
                Index = index,
                RelationshipId = rId
            };

            var slideXDoc = slidePart.GetXDocument();
            var slideRoot = slideXDoc.Root;

            // Get layout reference
            if (slidePart.SlideLayoutPart != null)
            {
                signature.LayoutRelationshipId = slidePart.GetIdOfPart(slidePart.SlideLayoutPart);
                // Compute layout hash based on layout type, not relationship ID or full content
                // (full content hash can differ due to XML serialization differences)
                var layoutXDoc = slidePart.SlideLayoutPart.GetXDocument();
                var layoutRoot = layoutXDoc.Root;
                var layoutType = (string)layoutRoot?.Attribute("type") ?? "custom";
                signature.LayoutHash = PmlHasher.ComputeHash(layoutType);
            }

            // Get common slide data
            var cSld = slideRoot.Element(P.cSld);
            if (cSld == null)
                return signature;

            // Get background hash
            var bg = cSld.Element(P.bg);
            if (bg != null)
            {
                signature.BackgroundHash = PmlHasher.ComputeHash(bg.ToString(SaveOptions.DisableFormatting));
            }

            // Get shape tree
            var spTree = cSld.Element(P.spTree);
            if (spTree == null)
                return signature;

            int zOrder = 0;
            foreach (var element in spTree.Elements())
            {
                if (element.Name == P.sp || element.Name == P.pic ||
                    element.Name == P.graphicFrame || element.Name == P.grpSp ||
                    element.Name == P.cxnSp)
                {
                    var shapeSignature = CanonicalizeShape(element, slidePart, zOrder++, settings);
                    if (shapeSignature != null)
                    {
                        signature.Shapes.Add(shapeSignature);

                        // Extract title text
                        if (shapeSignature.Placeholder?.Type == "title" ||
                            shapeSignature.Placeholder?.Type == "ctrTitle")
                        {
                            signature.TitleText = shapeSignature.TextBody?.PlainText;
                        }
                    }
                }
            }

            // Get notes text
            if (settings.CompareNotes && slidePart.NotesSlidePart != null)
            {
                signature.NotesText = ExtractNotesText(slidePart.NotesSlidePart);
            }

            // Compute content hash (includes shape names and types for consistency)
            var contentBuilder = new StringBuilder();
            contentBuilder.Append(signature.TitleText ?? "");
            foreach (var shape in signature.Shapes)
            {
                contentBuilder.Append("|");
                contentBuilder.Append(shape.Name ?? "");
                contentBuilder.Append(":");
                contentBuilder.Append(shape.Type);
                contentBuilder.Append(":");
                contentBuilder.Append(shape.TextBody?.PlainText ?? "");
            }
            signature.ContentHash = PmlHasher.ComputeHash(contentBuilder.ToString());

            return signature;
        }

        private static ShapeSignature CanonicalizeShape(
            XElement element,
            SlidePart slidePart,
            int zOrder,
            PmlComparerSettings settings)
        {
            var signature = new ShapeSignature
            {
                ZOrder = zOrder
            };

            // Determine shape type
            if (element.Name == P.sp)
            {
                signature.Type = PmlShapeType.AutoShape;
            }
            else if (element.Name == P.pic)
            {
                signature.Type = PmlShapeType.Picture;
            }
            else if (element.Name == P.graphicFrame)
            {
                // Could be table, chart, or diagram
                var graphic = element.Element(A.graphic);
                var graphicData = graphic?.Element(A.graphicData);
                var uri = (string)graphicData?.Attribute("uri");

                if (uri == "http://schemas.openxmlformats.org/drawingml/2006/table")
                    signature.Type = PmlShapeType.Table;
                else if (uri == "http://schemas.openxmlformats.org/drawingml/2006/chart")
                    signature.Type = PmlShapeType.Chart;
                else if (uri == "http://schemas.openxmlformats.org/drawingml/2006/diagram")
                    signature.Type = PmlShapeType.SmartArt;
                else
                    signature.Type = PmlShapeType.OleObject;
            }
            else if (element.Name == P.grpSp)
            {
                signature.Type = PmlShapeType.Group;
            }
            else if (element.Name == P.cxnSp)
            {
                signature.Type = PmlShapeType.Connector;
            }
            else
            {
                signature.Type = PmlShapeType.Unknown;
            }

            // Get non-visual properties
            var nvSpPr = element.Element(P.nvSpPr) ??
                         element.Element(P.nvPicPr) ??
                         element.Element(P.nvGraphicFramePr) ??
                         element.Element(P.nvGrpSpPr) ??
                         element.Element(P.nvCxnSpPr);

            if (nvSpPr != null)
            {
                var cNvPr = nvSpPr.Element(P.cNvPr);
                if (cNvPr != null)
                {
                    signature.Name = (string)cNvPr.Attribute("name") ?? "";
                    signature.Id = (uint?)cNvPr.Attribute("id") ?? 0;
                }

                // Get placeholder info
                var nvPr = nvSpPr.Element(P.nvPr);
                var ph = nvPr?.Element(P.ph);
                if (ph != null)
                {
                    signature.Placeholder = new PlaceholderInfo
                    {
                        Type = (string)ph.Attribute("type") ?? "body",
                        Index = (uint?)ph.Attribute("idx")
                    };
                }
            }

            // Get transform
            var spPr = element.Element(P.spPr) ??
                       element.Element(P.grpSpPr);

            if (spPr != null)
            {
                var xfrm = spPr.Element(A.xfrm);
                if (xfrm != null)
                {
                    signature.Transform = ExtractTransform(xfrm);
                }

                // Get geometry hash
                var prstGeom = spPr.Element(A.prstGeom);
                var custGeom = spPr.Element(A.custGeom);
                if (prstGeom != null)
                {
                    signature.GeometryHash = (string)prstGeom.Attribute("prst");
                }
                else if (custGeom != null)
                {
                    signature.GeometryHash = PmlHasher.ComputeHash(custGeom.ToString(SaveOptions.DisableFormatting));
                }
            }

            // For groups, check grpSpPr for transform
            if (element.Name == P.grpSp)
            {
                var grpSpPr = element.Element(P.grpSpPr);
                var xfrm = grpSpPr?.Element(A.xfrm);
                if (xfrm != null && signature.Transform == null)
                {
                    signature.Transform = ExtractTransform(xfrm);
                }
            }

            // Get text body
            var txBody = element.Element(P.txBody);
            if (txBody != null)
            {
                signature.TextBody = ExtractTextBody(txBody);
                if (signature.Type == PmlShapeType.AutoShape &&
                    !string.IsNullOrEmpty(signature.TextBody?.PlainText))
                {
                    signature.Type = PmlShapeType.TextBox;
                }
            }

            // Get image hash for pictures
            if (signature.Type == PmlShapeType.Picture)
            {
                signature.ImageHash = ExtractImageHash(element, slidePart);
            }

            // Get table hash
            if (signature.Type == PmlShapeType.Table)
            {
                signature.TableHash = ExtractTableHash(element);
            }

            // Get chart hash
            if (signature.Type == PmlShapeType.Chart)
            {
                signature.ChartHash = ExtractChartHash(element, slidePart);
            }

            // Handle group children
            if (signature.Type == PmlShapeType.Group)
            {
                signature.Children = new List<ShapeSignature>();
                int childZOrder = 0;
                foreach (var child in element.Elements())
                {
                    if (child.Name == P.sp || child.Name == P.pic ||
                        child.Name == P.graphicFrame || child.Name == P.grpSp ||
                        child.Name == P.cxnSp)
                    {
                        var childSig = CanonicalizeShape(child, slidePart, childZOrder++, settings);
                        if (childSig != null)
                        {
                            signature.Children.Add(childSig);
                        }
                    }
                }
            }

            // Compute content hash
            var contentBuilder = new StringBuilder();
            contentBuilder.Append(signature.Type);
            contentBuilder.Append("|");
            contentBuilder.Append(signature.TextBody?.PlainText ?? "");
            contentBuilder.Append("|");
            contentBuilder.Append(signature.ImageHash ?? "");
            contentBuilder.Append("|");
            contentBuilder.Append(signature.TableHash ?? "");
            contentBuilder.Append("|");
            contentBuilder.Append(signature.ChartHash ?? "");
            signature.ContentHash = PmlHasher.ComputeHash(contentBuilder.ToString());

            return signature;
        }

        private static TransformSignature ExtractTransform(XElement xfrm)
        {
            var off = xfrm.Element(A.off);
            var ext = xfrm.Element(A.ext);

            return new TransformSignature
            {
                X = (long?)off?.Attribute("x") ?? 0,
                Y = (long?)off?.Attribute("y") ?? 0,
                Cx = (long?)ext?.Attribute("cx") ?? 0,
                Cy = (long?)ext?.Attribute("cy") ?? 0,
                Rotation = (int?)xfrm.Attribute("rot") ?? 0,
                FlipH = (bool?)xfrm.Attribute("flipH") ?? false,
                FlipV = (bool?)xfrm.Attribute("flipV") ?? false
            };
        }

        private static TextBodySignature ExtractTextBody(XElement txBody)
        {
            var signature = new TextBodySignature();
            var plainTextBuilder = new StringBuilder();

            foreach (var p in txBody.Elements(A.p))
            {
                var para = new ParagraphSignature();
                var paraTextBuilder = new StringBuilder();

                // Get paragraph properties
                var pPr = p.Element(A.pPr);
                if (pPr != null)
                {
                    para.Alignment = (string)pPr.Attribute("algn");
                    para.HasBullet = pPr.Element(A.buChar) != null ||
                                     pPr.Element(A.buAutoNum) != null;
                }

                // Get runs
                foreach (var r in p.Elements(A.r))
                {
                    var run = new RunSignature();
                    var t = r.Element(A.t);
                    run.Text = t?.Value ?? "";
                    paraTextBuilder.Append(run.Text);

                    // Get run properties
                    var rPr = r.Element(A.rPr);
                    if (rPr != null)
                    {
                        run.Properties = ExtractRunProperties(rPr);
                    }

                    run.ContentHash = PmlHasher.ComputeHash(run.Text);
                    para.Runs.Add(run);
                }

                // Handle field codes
                foreach (var fld in p.Elements(A.fld))
                {
                    var t = fld.Element(A.t);
                    var text = t?.Value ?? "";
                    paraTextBuilder.Append(text);

                    var run = new RunSignature { Text = text };
                    para.Runs.Add(run);
                }

                para.PlainText = paraTextBuilder.ToString();
                if (plainTextBuilder.Length > 0)
                    plainTextBuilder.Append("\n");
                plainTextBuilder.Append(para.PlainText);
                signature.Paragraphs.Add(para);
            }

            signature.PlainText = plainTextBuilder.ToString();
            return signature;
        }

        private static RunPropertiesSignature ExtractRunProperties(XElement rPr)
        {
            var props = new RunPropertiesSignature();

            props.Bold = (bool?)rPr.Attribute("b") ?? false;
            props.Italic = (bool?)rPr.Attribute("i") ?? false;
            props.Underline = rPr.Attribute("u") != null && (string)rPr.Attribute("u") != "none";
            props.Strikethrough = rPr.Attribute("strike") != null && (string)rPr.Attribute("strike") != "noStrike";
            props.FontSize = (int?)rPr.Attribute("sz");

            // Get font name
            var latin = rPr.Element(A.latin);
            if (latin != null)
            {
                props.FontName = (string)latin.Attribute("typeface");
            }

            // Get font color
            var solidFill = rPr.Element(A.solidFill);
            if (solidFill != null)
            {
                var srgbClr = solidFill.Element(A.srgbClr);
                if (srgbClr != null)
                {
                    props.FontColor = (string)srgbClr.Attribute("val");
                }
            }

            return props;
        }

        private static string ExtractImageHash(XElement element, SlidePart slidePart)
        {
            var blipFill = element.Element(P.blipFill);
            var blip = blipFill?.Element(A.blip);
            var embed = (string)blip?.Attribute(R.embed);

            if (string.IsNullOrEmpty(embed))
                return null;

            try
            {
                var imagePart = slidePart.GetPartById(embed);
                using var stream = imagePart.GetStream();
                return PmlHasher.ComputeHash(stream);
            }
            catch
            {
                return null;
            }
        }

        private static string ExtractTableHash(XElement element)
        {
            var graphic = element.Element(A.graphic);
            var graphicData = graphic?.Element(A.graphicData);
            var tbl = graphicData?.Element(A.tbl);

            if (tbl == null)
                return null;

            // Hash table content
            var contentBuilder = new StringBuilder();
            foreach (var tr in tbl.Elements(A.tr))
            {
                foreach (var tc in tr.Elements(A.tc))
                {
                    var txBody = tc.Element(A.txBody);
                    if (txBody != null)
                    {
                        var text = ExtractTextBody(txBody);
                        contentBuilder.Append(text.PlainText);
                        contentBuilder.Append("|");
                    }
                }
                contentBuilder.Append("||");
            }

            return PmlHasher.ComputeHash(contentBuilder.ToString());
        }

        private static string ExtractChartHash(XElement element, SlidePart slidePart)
        {
            var graphic = element.Element(A.graphic);
            var graphicData = graphic?.Element(A.graphicData);
            var chartRef = graphicData?.Element(C.chart);
            var rId = (string)chartRef?.Attribute(R.id);

            if (string.IsNullOrEmpty(rId))
                return null;

            try
            {
                var chartPart = slidePart.GetPartById(rId) as ChartPart;
                if (chartPart == null)
                    return null;

                var chartXDoc = chartPart.GetXDocument();
                return PmlHasher.ComputeHash(chartXDoc.ToString(SaveOptions.DisableFormatting));
            }
            catch
            {
                return null;
            }
        }

        private static string ExtractNotesText(NotesSlidePart notesPart)
        {
            var xDoc = notesPart.GetXDocument();
            var spTree = xDoc.Root?.Element(P.cSld)?.Element(P.spTree);
            if (spTree == null)
                return null;

            var textBuilder = new StringBuilder();
            foreach (var sp in spTree.Elements(P.sp))
            {
                var txBody = sp.Element(P.txBody);
                if (txBody != null)
                {
                    var text = ExtractTextBody(txBody);
                    if (!string.IsNullOrEmpty(text.PlainText))
                    {
                        if (textBuilder.Length > 0)
                            textBuilder.Append("\n");
                        textBuilder.Append(text.PlainText);
                    }
                }
            }

            return textBuilder.ToString();
        }
    }

    #endregion

    #region Slide Match Engine

    /// <summary>
    /// Matches slides between two presentations.
    /// </summary>
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

            // Pass 1: Match by title text (exact match)
            MatchByTitleText(sig1, sig2, matches, used1, used2);

            // Pass 2: Match by content fingerprint
            MatchByFingerprint(sig1, sig2, matches, used1, used2, settings);

            // Pass 3: Match by position (remaining slides)
            if (settings.UseSlideAlignmentLCS)
            {
                MatchByLCS(sig1, sig2, matches, used1, used2, settings);
            }
            else
            {
                MatchByPosition(sig1, sig2, matches, used1, used2);
            }

            // Remaining unmatched = inserted/deleted
            AddUnmatchedAsInsertedDeleted(sig1, sig2, matches, used1, used2);

            // Sort by new index for consistent ordering
            return matches.OrderBy(m => m.NewIndex ?? int.MaxValue)
                          .ThenBy(m => m.OldIndex ?? int.MaxValue)
                          .ToList();
        }

        private static void MatchByTitleText(
            PresentationSignature sig1,
            PresentationSignature sig2,
            List<SlideMatch> matches,
            HashSet<int> used1,
            HashSet<int> used2)
        {
            foreach (var slide1 in sig1.Slides)
            {
                if (used1.Contains(slide1.Index))
                    continue;

                if (string.IsNullOrEmpty(slide1.TitleText))
                    continue;

                var match = sig2.Slides.FirstOrDefault(s2 =>
                    !used2.Contains(s2.Index) &&
                    s2.TitleText == slide1.TitleText);

                if (match != null)
                {
                    matches.Add(new SlideMatch
                    {
                        MatchType = SlideMatchType.Matched,
                        OldIndex = slide1.Index,
                        NewIndex = match.Index,
                        OldSlide = slide1,
                        NewSlide = match,
                        Similarity = 1.0
                    });
                    used1.Add(slide1.Index);
                    used2.Add(match.Index);
                }
            }
        }

        private static void MatchByFingerprint(
            PresentationSignature sig1,
            PresentationSignature sig2,
            List<SlideMatch> matches,
            HashSet<int> used1,
            HashSet<int> used2,
            PmlComparerSettings settings)
        {
            var fingerprints1 = sig1.Slides
                .Where(s => !used1.Contains(s.Index))
                .ToDictionary(s => s.Index, s => s.ComputeFingerprint());

            var fingerprints2 = sig2.Slides
                .Where(s => !used2.Contains(s.Index))
                .ToDictionary(s => s.Index, s => s.ComputeFingerprint());

            // Exact fingerprint match
            foreach (var slide1 in sig1.Slides.Where(s => !used1.Contains(s.Index)))
            {
                var fp1 = fingerprints1[slide1.Index];
                var match = sig2.Slides.FirstOrDefault(s2 =>
                    !used2.Contains(s2.Index) &&
                    fingerprints2.TryGetValue(s2.Index, out var fp2) &&
                    fp1 == fp2);

                if (match != null)
                {
                    matches.Add(new SlideMatch
                    {
                        MatchType = SlideMatchType.Matched,
                        OldIndex = slide1.Index,
                        NewIndex = match.Index,
                        OldSlide = slide1,
                        NewSlide = match,
                        Similarity = 1.0
                    });
                    used1.Add(slide1.Index);
                    used2.Add(match.Index);
                }
            }
        }

        private static void MatchByLCS(
            PresentationSignature sig1,
            PresentationSignature sig2,
            List<SlideMatch> matches,
            HashSet<int> used1,
            HashSet<int> used2,
            PmlComparerSettings settings)
        {
            var remaining1 = sig1.Slides.Where(s => !used1.Contains(s.Index)).ToList();
            var remaining2 = sig2.Slides.Where(s => !used2.Contains(s.Index)).ToList();

            if (remaining1.Count == 0 || remaining2.Count == 0)
                return;

            // Compute similarity matrix
            var similarities = new double[remaining1.Count, remaining2.Count];
            for (int i = 0; i < remaining1.Count; i++)
            {
                for (int j = 0; j < remaining2.Count; j++)
                {
                    similarities[i, j] = ComputeSlideSimilarity(remaining1[i], remaining2[j]);
                }
            }

            // Greedy matching by highest similarity
            var matched1 = new HashSet<int>();
            var matched2 = new HashSet<int>();

            while (matched1.Count < remaining1.Count && matched2.Count < remaining2.Count)
            {
                double bestSim = 0;
                int bestI = -1, bestJ = -1;

                for (int i = 0; i < remaining1.Count; i++)
                {
                    if (matched1.Contains(i)) continue;
                    for (int j = 0; j < remaining2.Count; j++)
                    {
                        if (matched2.Contains(j)) continue;
                        if (similarities[i, j] > bestSim)
                        {
                            bestSim = similarities[i, j];
                            bestI = i;
                            bestJ = j;
                        }
                    }
                }

                if (bestI < 0 || bestSim < settings.SlideSimilarityThreshold)
                    break;

                matches.Add(new SlideMatch
                {
                    MatchType = SlideMatchType.Matched,
                    OldIndex = remaining1[bestI].Index,
                    NewIndex = remaining2[bestJ].Index,
                    OldSlide = remaining1[bestI],
                    NewSlide = remaining2[bestJ],
                    Similarity = bestSim
                });
                used1.Add(remaining1[bestI].Index);
                used2.Add(remaining2[bestJ].Index);
                matched1.Add(bestI);
                matched2.Add(bestJ);
            }
        }

        private static void MatchByPosition(
            PresentationSignature sig1,
            PresentationSignature sig2,
            List<SlideMatch> matches,
            HashSet<int> used1,
            HashSet<int> used2)
        {
            var remaining1 = sig1.Slides.Where(s => !used1.Contains(s.Index)).OrderBy(s => s.Index).ToList();
            var remaining2 = sig2.Slides.Where(s => !used2.Contains(s.Index)).OrderBy(s => s.Index).ToList();

            int count = Math.Min(remaining1.Count, remaining2.Count);
            for (int i = 0; i < count; i++)
            {
                matches.Add(new SlideMatch
                {
                    MatchType = SlideMatchType.Matched,
                    OldIndex = remaining1[i].Index,
                    NewIndex = remaining2[i].Index,
                    OldSlide = remaining1[i],
                    NewSlide = remaining2[i],
                    Similarity = ComputeSlideSimilarity(remaining1[i], remaining2[i])
                });
                used1.Add(remaining1[i].Index);
                used2.Add(remaining2[i].Index);
            }
        }

        private static void AddUnmatchedAsInsertedDeleted(
            PresentationSignature sig1,
            PresentationSignature sig2,
            List<SlideMatch> matches,
            HashSet<int> used1,
            HashSet<int> used2)
        {
            // Deleted slides
            foreach (var slide in sig1.Slides.Where(s => !used1.Contains(s.Index)))
            {
                matches.Add(new SlideMatch
                {
                    MatchType = SlideMatchType.Deleted,
                    OldIndex = slide.Index,
                    OldSlide = slide
                });
            }

            // Inserted slides
            foreach (var slide in sig2.Slides.Where(s => !used2.Contains(s.Index)))
            {
                matches.Add(new SlideMatch
                {
                    MatchType = SlideMatchType.Inserted,
                    NewIndex = slide.Index,
                    NewSlide = slide
                });
            }
        }

        private static double ComputeSlideSimilarity(SlideSignature s1, SlideSignature s2)
        {
            double score = 0;
            double maxScore = 0;

            // Title match (high weight)
            maxScore += 3;
            if (!string.IsNullOrEmpty(s1.TitleText) && s1.TitleText == s2.TitleText)
                score += 3;
            else if (!string.IsNullOrEmpty(s1.TitleText) && !string.IsNullOrEmpty(s2.TitleText))
            {
                var similarity = ComputeTextSimilarity(s1.TitleText, s2.TitleText);
                score += similarity * 2;
            }

            // Content hash match
            maxScore += 2;
            if (s1.ContentHash == s2.ContentHash)
                score += 2;

            // Shape count similarity
            maxScore += 1;
            var shapeCount1 = s1.Shapes.Count;
            var shapeCount2 = s2.Shapes.Count;
            if (shapeCount1 == shapeCount2)
                score += 1;
            else if (Math.Abs(shapeCount1 - shapeCount2) <= 2)
                score += 0.5;

            // Shape types match
            maxScore += 1;
            var types1 = s1.Shapes.Select(s => s.Type).OrderBy(t => t).ToList();
            var types2 = s2.Shapes.Select(s => s.Type).OrderBy(t => t).ToList();
            if (types1.SequenceEqual(types2))
                score += 1;

            return maxScore > 0 ? score / maxScore : 0;
        }

        private static double ComputeTextSimilarity(string s1, string s2)
        {
            if (string.IsNullOrEmpty(s1) || string.IsNullOrEmpty(s2))
                return 0;

            if (s1 == s2)
                return 1;

            // Simple Jaccard similarity on words
            var words1 = s1.ToLower().Split(' ', StringSplitOptions.RemoveEmptyEntries).ToHashSet();
            var words2 = s2.ToLower().Split(' ', StringSplitOptions.RemoveEmptyEntries).ToHashSet();

            var intersection = words1.Intersect(words2).Count();
            var union = words1.Union(words2).Count();

            return union > 0 ? (double)intersection / union : 0;
        }
    }

    #endregion

    #region Shape Match Engine

    /// <summary>
    /// Matches shapes within slides.
    /// </summary>
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

            // Pass 1: Match by placeholder
            MatchByPlaceholder(slide1, slide2, matches, used1, used2);

            // Pass 2: Match by name and type
            MatchByNameAndType(slide1, slide2, matches, used1, used2);

            // Pass 3: Match by name only
            MatchByNameOnly(slide1, slide2, matches, used1, used2);

            // Pass 4: Fuzzy matching
            if (settings.EnableFuzzyShapeMatching)
            {
                FuzzyMatch(slide1, slide2, matches, used1, used2, settings);
            }

            // Remaining unmatched
            AddUnmatchedAsInsertedDeleted(slide1, slide2, matches, used1, used2);

            return matches;
        }

        private static string GetShapeKey(ShapeSignature shape)
        {
            return $"{shape.Id}:{shape.Name}";
        }

        private static void MatchByPlaceholder(
            SlideSignature slide1,
            SlideSignature slide2,
            List<ShapeMatch> matches,
            HashSet<string> used1,
            HashSet<string> used2)
        {
            var placeholders1 = slide1.Shapes
                .Where(s => s.Placeholder != null)
                .ToList();

            foreach (var shape1 in placeholders1)
            {
                var key1 = GetShapeKey(shape1);
                if (used1.Contains(key1))
                    continue;

                var match = slide2.Shapes.FirstOrDefault(s2 =>
                    s2.Placeholder != null &&
                    !used2.Contains(GetShapeKey(s2)) &&
                    s2.Placeholder.Equals(shape1.Placeholder));

                if (match != null)
                {
                    var key2 = GetShapeKey(match);
                    matches.Add(new ShapeMatch
                    {
                        MatchType = ShapeMatchType.Matched,
                        OldShape = shape1,
                        NewShape = match,
                        Score = 1.0,
                        Method = ShapeMatchMethod.Placeholder
                    });
                    used1.Add(key1);
                    used2.Add(key2);
                }
            }
        }

        private static void MatchByNameAndType(
            SlideSignature slide1,
            SlideSignature slide2,
            List<ShapeMatch> matches,
            HashSet<string> used1,
            HashSet<string> used2)
        {
            foreach (var shape1 in slide1.Shapes)
            {
                var key1 = GetShapeKey(shape1);
                if (used1.Contains(key1))
                    continue;

                if (string.IsNullOrEmpty(shape1.Name))
                    continue;

                var match = slide2.Shapes.FirstOrDefault(s2 =>
                    !used2.Contains(GetShapeKey(s2)) &&
                    s2.Name == shape1.Name &&
                    s2.Type == shape1.Type);

                if (match != null)
                {
                    var key2 = GetShapeKey(match);
                    matches.Add(new ShapeMatch
                    {
                        MatchType = ShapeMatchType.Matched,
                        OldShape = shape1,
                        NewShape = match,
                        Score = 0.95,
                        Method = ShapeMatchMethod.NameAndType
                    });
                    used1.Add(key1);
                    used2.Add(key2);
                }
            }
        }

        private static void MatchByNameOnly(
            SlideSignature slide1,
            SlideSignature slide2,
            List<ShapeMatch> matches,
            HashSet<string> used1,
            HashSet<string> used2)
        {
            foreach (var shape1 in slide1.Shapes)
            {
                var key1 = GetShapeKey(shape1);
                if (used1.Contains(key1))
                    continue;

                if (string.IsNullOrEmpty(shape1.Name))
                    continue;

                var match = slide2.Shapes.FirstOrDefault(s2 =>
                    !used2.Contains(GetShapeKey(s2)) &&
                    s2.Name == shape1.Name);

                if (match != null)
                {
                    var key2 = GetShapeKey(match);
                    matches.Add(new ShapeMatch
                    {
                        MatchType = ShapeMatchType.Matched,
                        OldShape = shape1,
                        NewShape = match,
                        Score = 0.8,
                        Method = ShapeMatchMethod.NameOnly
                    });
                    used1.Add(key1);
                    used2.Add(key2);
                }
            }
        }

        private static void FuzzyMatch(
            SlideSignature slide1,
            SlideSignature slide2,
            List<ShapeMatch> matches,
            HashSet<string> used1,
            HashSet<string> used2,
            PmlComparerSettings settings)
        {
            var remaining1 = slide1.Shapes.Where(s => !used1.Contains(GetShapeKey(s))).ToList();
            var remaining2 = slide2.Shapes.Where(s => !used2.Contains(GetShapeKey(s))).ToList();

            foreach (var shape1 in remaining1)
            {
                var key1 = GetShapeKey(shape1);
                if (used1.Contains(key1))
                    continue;

                double bestScore = 0;
                ShapeSignature bestMatch = null;

                foreach (var shape2 in remaining2)
                {
                    var key2 = GetShapeKey(shape2);
                    if (used2.Contains(key2))
                        continue;

                    var score = ComputeShapeMatchScore(shape1, shape2, settings);
                    if (score > bestScore && score >= settings.ShapeSimilarityThreshold)
                    {
                        bestScore = score;
                        bestMatch = shape2;
                    }
                }

                if (bestMatch != null)
                {
                    var key2 = GetShapeKey(bestMatch);
                    matches.Add(new ShapeMatch
                    {
                        MatchType = ShapeMatchType.Matched,
                        OldShape = shape1,
                        NewShape = bestMatch,
                        Score = bestScore,
                        Method = ShapeMatchMethod.Fuzzy
                    });
                    used1.Add(key1);
                    used2.Add(key2);
                }
            }
        }

        private static void AddUnmatchedAsInsertedDeleted(
            SlideSignature slide1,
            SlideSignature slide2,
            List<ShapeMatch> matches,
            HashSet<string> used1,
            HashSet<string> used2)
        {
            // Deleted shapes
            foreach (var shape in slide1.Shapes.Where(s => !used1.Contains(GetShapeKey(s))))
            {
                matches.Add(new ShapeMatch
                {
                    MatchType = ShapeMatchType.Deleted,
                    OldShape = shape,
                    Score = 0
                });
            }

            // Inserted shapes
            foreach (var shape in slide2.Shapes.Where(s => !used2.Contains(GetShapeKey(s))))
            {
                matches.Add(new ShapeMatch
                {
                    MatchType = ShapeMatchType.Inserted,
                    NewShape = shape,
                    Score = 0
                });
            }
        }

        private static double ComputeShapeMatchScore(
            ShapeSignature s1,
            ShapeSignature s2,
            PmlComparerSettings settings)
        {
            double score = 0;

            // Same type (required)
            if (s1.Type != s2.Type)
                return 0;
            score += 0.2;

            // Position similarity
            if (s1.Transform != null && s2.Transform != null)
            {
                if (s1.Transform.IsNear(s2.Transform, settings.PositionTolerance))
                    score += 0.3;
                else
                {
                    // Partial credit for nearby positions
                    var distance = Math.Sqrt(
                        Math.Pow(s1.Transform.X - s2.Transform.X, 2) +
                        Math.Pow(s1.Transform.Y - s2.Transform.Y, 2));
                    if (distance < settings.PositionTolerance * 5)
                        score += 0.1;
                }
            }

            // Content similarity
            if (s1.Type == PmlShapeType.Picture)
            {
                if (!string.IsNullOrEmpty(s1.ImageHash) && s1.ImageHash == s2.ImageHash)
                    score += 0.5;
            }
            else if (s1.TextBody != null && s2.TextBody != null)
            {
                if (s1.TextBody.PlainText == s2.TextBody.PlainText)
                    score += 0.5;
                else
                {
                    var textSim = ComputeTextSimilarity(s1.TextBody.PlainText, s2.TextBody.PlainText);
                    score += textSim * 0.5;
                }
            }
            else if (s1.ContentHash == s2.ContentHash)
            {
                score += 0.5;
            }

            return score;
        }

        private static double ComputeTextSimilarity(string s1, string s2)
        {
            if (string.IsNullOrEmpty(s1) && string.IsNullOrEmpty(s2))
                return 1;
            if (string.IsNullOrEmpty(s1) || string.IsNullOrEmpty(s2))
                return 0;
            if (s1 == s2)
                return 1;

            // Levenshtein-based similarity
            var maxLen = Math.Max(s1.Length, s2.Length);
            if (maxLen == 0)
                return 1;

            var distance = LevenshteinDistance(s1, s2);
            return 1.0 - ((double)distance / maxLen);
        }

        private static int LevenshteinDistance(string s1, string s2)
        {
            var m = s1.Length;
            var n = s2.Length;
            var d = new int[m + 1, n + 1];

            for (int i = 0; i <= m; i++)
                d[i, 0] = i;
            for (int j = 0; j <= n; j++)
                d[0, j] = j;

            for (int j = 1; j <= n; j++)
            {
                for (int i = 1; i <= m; i++)
                {
                    var cost = s1[i - 1] == s2[j - 1] ? 0 : 1;
                    d[i, j] = Math.Min(
                        Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1),
                        d[i - 1, j - 1] + cost);
                }
            }

            return d[m, n];
        }
    }

    #endregion

    #region Diff Engine

    /// <summary>
    /// Computes differences between matched presentations.
    /// </summary>
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
            if (sig1.SlideCx != sig2.SlideCx || sig1.SlideCy != sig2.SlideCy)
            {
                result.Changes.Add(new PmlChange
                {
                    ChangeType = PmlChangeType.SlideSizeChanged,
                    OldValue = $"{sig1.SlideCx}x{sig1.SlideCy}",
                    NewValue = $"{sig2.SlideCx}x{sig2.SlideCy}"
                });
            }

            // Process slide matches
            foreach (var slideMatch in slideMatches)
            {
                switch (slideMatch.MatchType)
                {
                    case SlideMatchType.Inserted:
                        if (settings.CompareSlideStructure)
                        {
                            result.Changes.Add(new PmlChange
                            {
                                ChangeType = PmlChangeType.SlideInserted,
                                SlideIndex = slideMatch.NewIndex
                            });
                        }
                        break;

                    case SlideMatchType.Deleted:
                        if (settings.CompareSlideStructure)
                        {
                            result.Changes.Add(new PmlChange
                            {
                                ChangeType = PmlChangeType.SlideDeleted,
                                OldSlideIndex = slideMatch.OldIndex
                            });
                        }
                        break;

                    case SlideMatchType.Matched:
                        // Check if moved
                        if (settings.CompareSlideStructure && slideMatch.WasMoved)
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
            // Compare layout (use content hash, not relationship ID)
            if (slide1.LayoutHash != slide2.LayoutHash)
            {
                result.Changes.Add(new PmlChange
                {
                    ChangeType = PmlChangeType.SlideLayoutChanged,
                    SlideIndex = slideIndex
                });
            }

            // Compare background
            if (slide1.BackgroundHash != slide2.BackgroundHash)
            {
                result.Changes.Add(new PmlChange
                {
                    ChangeType = PmlChangeType.SlideBackgroundChanged,
                    SlideIndex = slideIndex
                });
            }

            // Compare notes
            if (settings.CompareNotes && slide1.NotesText != slide2.NotesText)
            {
                result.Changes.Add(new PmlChange
                {
                    ChangeType = PmlChangeType.SlideNotesChanged,
                    SlideIndex = slideIndex,
                    OldValue = slide1.NotesText,
                    NewValue = slide2.NotesText
                });
            }

            // Match and compare shapes
            if (settings.CompareShapeStructure)
            {
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
                                ShapeId = shapeMatch.NewShape.Id.ToString(),
                                MatchConfidence = shapeMatch.Score
                            });
                            break;

                        case ShapeMatchType.Deleted:
                            result.Changes.Add(new PmlChange
                            {
                                ChangeType = PmlChangeType.ShapeDeleted,
                                SlideIndex = slideIndex,
                                ShapeName = shapeMatch.OldShape.Name,
                                ShapeId = shapeMatch.OldShape.Id.ToString(),
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
            if (settings.CompareShapeTransforms && shape1.Transform != null && shape2.Transform != null)
            {
                var t1 = shape1.Transform;
                var t2 = shape2.Transform;

                // Position change
                if (!t1.IsNear(t2, settings.PositionTolerance))
                {
                    result.Changes.Add(new PmlChange
                    {
                        ChangeType = PmlChangeType.ShapeMoved,
                        SlideIndex = slideIndex,
                        ShapeName = shape2.Name,
                        ShapeId = shape2.Id.ToString(),
                        OldX = t1.X,
                        OldY = t1.Y,
                        NewX = t2.X,
                        NewY = t2.Y,
                        MatchConfidence = match.Score
                    });
                }

                // Size change
                if (!t1.IsSameSize(t2, settings.PositionTolerance))
                {
                    result.Changes.Add(new PmlChange
                    {
                        ChangeType = PmlChangeType.ShapeResized,
                        SlideIndex = slideIndex,
                        ShapeName = shape2.Name,
                        ShapeId = shape2.Id.ToString(),
                        OldCx = t1.Cx,
                        OldCy = t1.Cy,
                        NewCx = t2.Cx,
                        NewCy = t2.Cy,
                        MatchConfidence = match.Score
                    });
                }

                // Rotation change
                if (t1.Rotation != t2.Rotation)
                {
                    result.Changes.Add(new PmlChange
                    {
                        ChangeType = PmlChangeType.ShapeRotated,
                        SlideIndex = slideIndex,
                        ShapeName = shape2.Name,
                        ShapeId = shape2.Id.ToString(),
                        OldValue = t1.Rotation.ToString(),
                        NewValue = t2.Rotation.ToString(),
                        MatchConfidence = match.Score
                    });
                }
            }

            // Z-order change
            if (shape1.ZOrder != shape2.ZOrder)
            {
                result.Changes.Add(new PmlChange
                {
                    ChangeType = PmlChangeType.ShapeZOrderChanged,
                    SlideIndex = slideIndex,
                    ShapeName = shape2.Name,
                    ShapeId = shape2.Id.ToString(),
                    OldValue = shape1.ZOrder.ToString(),
                    NewValue = shape2.ZOrder.ToString()
                });
            }

            // Content changes based on type
            switch (shape1.Type)
            {
                case PmlShapeType.TextBox:
                case PmlShapeType.AutoShape:
                    if (settings.CompareTextContent)
                    {
                        CompareTextContent(shape1, shape2, slideIndex, settings, result);
                    }
                    break;

                case PmlShapeType.Picture:
                    if (settings.CompareImageContent && shape1.ImageHash != shape2.ImageHash)
                    {
                        result.Changes.Add(new PmlChange
                        {
                            ChangeType = PmlChangeType.ImageReplaced,
                            SlideIndex = slideIndex,
                            ShapeName = shape2.Name,
                            ShapeId = shape2.Id.ToString()
                        });
                    }
                    break;

                case PmlShapeType.Table:
                    if (settings.CompareTables && shape1.TableHash != shape2.TableHash)
                    {
                        result.Changes.Add(new PmlChange
                        {
                            ChangeType = PmlChangeType.TableContentChanged,
                            SlideIndex = slideIndex,
                            ShapeName = shape2.Name,
                            ShapeId = shape2.Id.ToString()
                        });
                    }
                    break;

                case PmlShapeType.Chart:
                    if (settings.CompareCharts && shape1.ChartHash != shape2.ChartHash)
                    {
                        result.Changes.Add(new PmlChange
                        {
                            ChangeType = PmlChangeType.ChartDataChanged,
                            SlideIndex = slideIndex,
                            ShapeName = shape2.Name,
                            ShapeId = shape2.Id.ToString()
                        });
                    }
                    break;
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

            if (text1 == null && text2 == null)
                return;

            if (text1 == null || text2 == null)
            {
                result.Changes.Add(new PmlChange
                {
                    ChangeType = PmlChangeType.TextChanged,
                    SlideIndex = slideIndex,
                    ShapeName = shape2.Name,
                    ShapeId = shape2.Id.ToString(),
                    OldValue = text1?.PlainText ?? "",
                    NewValue = text2?.PlainText ?? ""
                });
                return;
            }

            // Compare plain text first
            if (text1.PlainText != text2.PlainText)
            {
                result.Changes.Add(new PmlChange
                {
                    ChangeType = PmlChangeType.TextChanged,
                    SlideIndex = slideIndex,
                    ShapeName = shape2.Name,
                    ShapeId = shape2.Id.ToString(),
                    OldValue = text1.PlainText,
                    NewValue = text2.PlainText
                });
            }
            else if (settings.CompareTextFormatting)
            {
                // Text is same, check formatting
                if (HasFormattingChanges(text1, text2))
                {
                    result.Changes.Add(new PmlChange
                    {
                        ChangeType = PmlChangeType.TextFormattingChanged,
                        SlideIndex = slideIndex,
                        ShapeName = shape2.Name,
                        ShapeId = shape2.Id.ToString()
                    });
                }
            }
        }

        private static bool HasFormattingChanges(TextBodySignature text1, TextBodySignature text2)
        {
            if (text1.Paragraphs.Count != text2.Paragraphs.Count)
                return true;

            for (int i = 0; i < text1.Paragraphs.Count; i++)
            {
                var p1 = text1.Paragraphs[i];
                var p2 = text2.Paragraphs[i];

                if (p1.Alignment != p2.Alignment || p1.HasBullet != p2.HasBullet)
                    return true;

                if (p1.Runs.Count != p2.Runs.Count)
                    return true;

                for (int j = 0; j < p1.Runs.Count; j++)
                {
                    var r1 = p1.Runs[j];
                    var r2 = p2.Runs[j];

                    if (r1.Properties != null && r2.Properties != null)
                    {
                        if (!r1.Properties.Equals(r2.Properties))
                            return true;
                    }
                    else if (r1.Properties != null || r2.Properties != null)
                    {
                        return true;
                    }
                }
            }

            return false;
        }
    }

    #endregion

    #region Markup Renderer

    /// <summary>
    /// Produces marked presentations with visual change overlays.
    /// </summary>
    public static class PmlMarkupRenderer
    {
        public static PmlDocument RenderMarkedPresentation(
            PmlDocument newerDoc,
            PmlComparisonResult result,
            PmlComparerSettings settings)
        {
            if (result.TotalChanges == 0)
                return newerDoc;

            using var ms = new MemoryStream();
            ms.Write(newerDoc.DocumentByteArray, 0, newerDoc.DocumentByteArray.Length);
            ms.Position = 0;

            using (var pDoc = PresentationDocument.Open(ms, true))
            {
                var presentationPart = pDoc.PresentationPart;
                if (presentationPart == null)
                    return newerDoc;

                // Group changes by slide
                var changesBySlide = result.Changes
                    .Where(c => c.SlideIndex.HasValue)
                    .GroupBy(c => c.SlideIndex.Value)
                    .ToDictionary(g => g.Key, g => g.ToList());

                // Process each slide
                var slideIds = presentationPart.Presentation.SlideIdList?.Elements<SlideId>().ToList() ?? new List<SlideId>();

                for (int i = 0; i < slideIds.Count; i++)
                {
                    var slideIndex = i + 1;
                    if (!changesBySlide.TryGetValue(slideIndex, out var slideChanges))
                        continue;

                    var slideId = slideIds[i];
                    var rId = slideId.RelationshipId;
                    if (string.IsNullOrEmpty(rId))
                        continue;

                    try
                    {
                        var slidePart = (SlidePart)presentationPart.GetPartById(rId);
                        AddChangeOverlays(slidePart, slideChanges, settings);

                        if (settings.AddNotesAnnotations)
                        {
                            AddNotesAnnotations(slidePart, slideChanges, settings);
                        }
                    }
                    catch
                    {
                        // Skip slides that can't be processed
                    }
                }

                // Add summary slide
                if (settings.AddSummarySlide && result.TotalChanges > 0)
                {
                    AddSummarySlide(presentationPart, result, settings);
                }

                pDoc.Save();
            }

            return new PmlDocument(newerDoc.FileName, ms.ToArray());
        }

        private static void AddChangeOverlays(
            SlidePart slidePart,
            List<PmlChange> changes,
            PmlComparerSettings settings)
        {
            var slideXDoc = slidePart.GetXDocument();
            var spTree = slideXDoc.Root?.Element(P.cSld)?.Element(P.spTree);
            if (spTree == null)
                return;

            uint nextId = GetNextShapeId(spTree);

            foreach (var change in changes)
            {
                switch (change.ChangeType)
                {
                    case PmlChangeType.ShapeInserted:
                        AddChangeLabel(spTree, change, "NEW", settings.InsertedColor, ref nextId);
                        break;

                    case PmlChangeType.ShapeMoved:
                        AddChangeLabel(spTree, change, "MOVED", settings.MovedColor, ref nextId);
                        break;

                    case PmlChangeType.ShapeResized:
                        AddChangeLabel(spTree, change, "RESIZED", settings.ModifiedColor, ref nextId);
                        break;

                    case PmlChangeType.TextChanged:
                        AddChangeLabel(spTree, change, "TEXT CHANGED", settings.ModifiedColor, ref nextId);
                        break;

                    case PmlChangeType.ImageReplaced:
                        AddChangeLabel(spTree, change, "IMAGE REPLACED", settings.ModifiedColor, ref nextId);
                        break;

                    case PmlChangeType.TableContentChanged:
                        AddChangeLabel(spTree, change, "TABLE CHANGED", settings.ModifiedColor, ref nextId);
                        break;

                    case PmlChangeType.ChartDataChanged:
                        AddChangeLabel(spTree, change, "CHART CHANGED", settings.ModifiedColor, ref nextId);
                        break;
                }
            }

            slidePart.PutXDocument();
        }

        private static uint GetNextShapeId(XElement spTree)
        {
            uint maxId = 0;
            foreach (var nvPr in spTree.Descendants(P.cNvPr))
            {
                var id = (uint?)nvPr.Attribute("id") ?? 0;
                if (id > maxId)
                    maxId = id;
            }
            return maxId + 1;
        }

        private static void AddChangeLabel(
            XElement spTree,
            PmlChange change,
            string labelText,
            string color,
            ref uint nextId)
        {
            // Create a small label shape
            long x = change.NewX ?? 0;
            long y = change.NewY ?? 0;

            // Position label above the shape if we have coordinates
            if (y > 200000)
                y -= 200000;

            var label = new XElement(P.sp,
                new XElement(P.nvSpPr,
                    new XElement(P.cNvPr,
                        new XAttribute("id", nextId++),
                        new XAttribute("name", $"Change Label: {change.ShapeName}")),
                    new XElement(P.cNvSpPr,
                        new XElement(A.spLocks,
                            new XAttribute("noGrp", "1"))),
                    new XElement(P.nvPr)),
                new XElement(P.spPr,
                    new XElement(A.xfrm,
                        new XElement(A.off,
                            new XAttribute("x", x),
                            new XAttribute("y", y)),
                        new XElement(A.ext,
                            new XAttribute("cx", "1500000"),
                            new XAttribute("cy", "300000"))),
                    new XElement(A.prstGeom,
                        new XAttribute("prst", "rect"),
                        new XElement(A.avLst)),
                    new XElement(A.solidFill,
                        new XElement(A.srgbClr,
                            new XAttribute("val", color))),
                    new XElement(A.ln,
                        new XAttribute("w", "12700"),
                        new XElement(A.solidFill,
                            new XElement(A.srgbClr,
                                new XAttribute("val", "000000"))))),
                new XElement(P.txBody,
                    new XElement(A.bodyPr,
                        new XAttribute("wrap", "square"),
                        new XAttribute("rtlCol", "0"),
                        new XAttribute("anchor", "ctr")),
                    new XElement(A.lstStyle),
                    new XElement(A.p,
                        new XElement(A.pPr,
                            new XAttribute("algn", "ctr")),
                        new XElement(A.r,
                            new XElement(A.rPr,
                                new XAttribute("lang", "en-US"),
                                new XAttribute("sz", "1000"),
                                new XAttribute("b", "1"),
                                new XElement(A.solidFill,
                                    new XElement(A.srgbClr,
                                        new XAttribute("val", "FFFFFF")))),
                            new XElement(A.t, labelText)),
                        new XElement(A.endParaRPr,
                            new XAttribute("lang", "en-US")))));

            spTree.Add(label);
        }

        private static void AddNotesAnnotations(
            SlidePart slidePart,
            List<PmlChange> changes,
            PmlComparerSettings settings)
        {
            // Get or create notes slide part
            NotesSlidePart notesPart;
            if (slidePart.NotesSlidePart == null)
            {
                notesPart = slidePart.AddNewPart<NotesSlidePart>();
                notesPart.PutXDocument(CreateEmptyNotesSlide());
            }
            else
            {
                notesPart = slidePart.NotesSlidePart;
            }

            var notesXDoc = notesPart.GetXDocument();
            var spTree = notesXDoc.Root?.Element(P.cSld)?.Element(P.spTree);
            if (spTree == null)
                return;

            // Find or create the notes text shape
            var notesShape = spTree.Elements(P.sp)
                .FirstOrDefault(sp =>
                {
                    var ph = sp.Element(P.nvSpPr)?.Element(P.nvPr)?.Element(P.ph);
                    return ph != null && (string)ph.Attribute("type") == "body";
                });

            if (notesShape == null)
                return;

            var txBody = notesShape.Element(P.txBody);
            if (txBody == null)
                return;

            // Add change summary
            var summaryPara = new XElement(A.p,
                new XElement(A.r,
                    new XElement(A.rPr,
                        new XAttribute("lang", "en-US"),
                        new XAttribute("b", "1")),
                    new XElement(A.t, $"--- Changes ({changes.Count}) ---")));
            txBody.Add(summaryPara);

            foreach (var change in changes.Take(10)) // Limit to 10 changes
            {
                var changePara = new XElement(A.p,
                    new XElement(A.r,
                        new XElement(A.rPr,
                            new XAttribute("lang", "en-US")),
                        new XElement(A.t, $"- {change.GetDescription()}")));
                txBody.Add(changePara);
            }

            if (changes.Count > 10)
            {
                var morePara = new XElement(A.p,
                    new XElement(A.r,
                        new XElement(A.rPr,
                            new XAttribute("lang", "en-US")),
                        new XElement(A.t, $"... and {changes.Count - 10} more changes")));
                txBody.Add(morePara);
            }

            notesPart.PutXDocument();
        }

        private static XDocument CreateEmptyNotesSlide()
        {
            return new XDocument(
                new XDeclaration("1.0", "UTF-8", "yes"),
                new XElement(P.notes,
                    new XAttribute(XNamespace.Xmlns + "a", A.a),
                    new XAttribute(XNamespace.Xmlns + "p", P.p),
                    new XAttribute(XNamespace.Xmlns + "r", R.r),
                    new XElement(P.cSld,
                        new XElement(P.spTree,
                            new XElement(P.nvGrpSpPr,
                                new XElement(P.cNvPr,
                                    new XAttribute("id", "1"),
                                    new XAttribute("name", "")),
                                new XElement(P.cNvGrpSpPr),
                                new XElement(P.nvPr)),
                            new XElement(P.grpSpPr),
                            new XElement(P.sp,
                                new XElement(P.nvSpPr,
                                    new XElement(P.cNvPr,
                                        new XAttribute("id", "2"),
                                        new XAttribute("name", "Notes Placeholder")),
                                    new XElement(P.cNvSpPr),
                                    new XElement(P.nvPr,
                                        new XElement(P.ph,
                                            new XAttribute("type", "body"),
                                            new XAttribute("idx", "1")))),
                                new XElement(P.spPr),
                                new XElement(P.txBody,
                                    new XElement(A.bodyPr),
                                    new XElement(A.lstStyle),
                                    new XElement(A.p,
                                        new XElement(A.endParaRPr,
                                            new XAttribute("lang", "en-US")))))))));
        }

        private static void AddSummarySlide(
            PresentationPart presentationPart,
            PmlComparisonResult result,
            PmlComparerSettings settings)
        {
            // Create a new slide part
            var slidePart = presentationPart.AddNewPart<SlidePart>();

            var slideXDoc = new XDocument(
                new XDeclaration("1.0", "UTF-8", "yes"),
                new XElement(P.sld,
                    new XAttribute(XNamespace.Xmlns + "a", A.a),
                    new XAttribute(XNamespace.Xmlns + "p", P.p),
                    new XAttribute(XNamespace.Xmlns + "r", R.r),
                    new XElement(P.cSld,
                        new XElement(P.spTree,
                            new XElement(P.nvGrpSpPr,
                                new XElement(P.cNvPr,
                                    new XAttribute("id", "1"),
                                    new XAttribute("name", "")),
                                new XElement(P.cNvGrpSpPr),
                                new XElement(P.nvPr)),
                            new XElement(P.grpSpPr),
                            CreateTitleShape("Comparison Summary", 2),
                            CreateSummaryContentShape(result, settings, 3)))));

            slidePart.PutXDocument(slideXDoc);

            // Add slide to presentation
            var slideIdList = presentationPart.Presentation.SlideIdList;
            if (slideIdList == null)
            {
                slideIdList = new SlideIdList();
                presentationPart.Presentation.SlideIdList = slideIdList;
            }

            uint maxId = 256;
            foreach (var slideId in slideIdList.Elements<SlideId>())
            {
                if (slideId.Id > maxId)
                    maxId = slideId.Id;
            }

            var newSlideId = new SlideId
            {
                Id = maxId + 1,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            };
            slideIdList.Append(newSlideId);

            presentationPart.Presentation.Save();
        }

        private static XElement CreateTitleShape(string title, uint id)
        {
            return new XElement(P.sp,
                new XElement(P.nvSpPr,
                    new XElement(P.cNvPr,
                        new XAttribute("id", id),
                        new XAttribute("name", "Title")),
                    new XElement(P.cNvSpPr),
                    new XElement(P.nvPr)),
                new XElement(P.spPr,
                    new XElement(A.xfrm,
                        new XElement(A.off,
                            new XAttribute("x", "457200"),
                            new XAttribute("y", "274638")),
                        new XElement(A.ext,
                            new XAttribute("cx", "8229600"),
                            new XAttribute("cy", "1143000"))),
                    new XElement(A.prstGeom,
                        new XAttribute("prst", "rect"),
                        new XElement(A.avLst))),
                new XElement(P.txBody,
                    new XElement(A.bodyPr),
                    new XElement(A.lstStyle),
                    new XElement(A.p,
                        new XElement(A.r,
                            new XElement(A.rPr,
                                new XAttribute("lang", "en-US"),
                                new XAttribute("sz", "4400"),
                                new XAttribute("b", "1")),
                            new XElement(A.t, title)))));
        }

        private static XElement CreateSummaryContentShape(PmlComparisonResult result, PmlComparerSettings settings, uint id)
        {
            var content = new StringBuilder();
            content.AppendLine($"Total Changes: {result.TotalChanges}");
            content.AppendLine();
            content.AppendLine($"Slides Inserted: {result.SlidesInserted}");
            content.AppendLine($"Slides Deleted: {result.SlidesDeleted}");
            content.AppendLine($"Slides Moved: {result.SlidesMoved}");
            content.AppendLine();
            content.AppendLine($"Shapes Inserted: {result.ShapesInserted}");
            content.AppendLine($"Shapes Deleted: {result.ShapesDeleted}");
            content.AppendLine($"Shapes Moved: {result.ShapesMoved}");
            content.AppendLine($"Shapes Resized: {result.ShapesResized}");
            content.AppendLine();
            content.AppendLine($"Text Changes: {result.TextChanges}");
            content.AppendLine($"Formatting Changes: {result.FormattingChanges}");
            content.AppendLine($"Images Replaced: {result.ImagesReplaced}");

            var txBody = new XElement(P.txBody,
                new XElement(A.bodyPr),
                new XElement(A.lstStyle));

            foreach (var line in content.ToString().Split('\n'))
            {
                txBody.Add(new XElement(A.p,
                    new XElement(A.r,
                        new XElement(A.rPr,
                            new XAttribute("lang", "en-US"),
                            new XAttribute("sz", "2000")),
                        new XElement(A.t, line.TrimEnd()))));
            }

            return new XElement(P.sp,
                new XElement(P.nvSpPr,
                    new XElement(P.cNvPr,
                        new XAttribute("id", id),
                        new XAttribute("name", "Content")),
                    new XElement(P.cNvSpPr),
                    new XElement(P.nvPr)),
                new XElement(P.spPr,
                    new XElement(A.xfrm,
                        new XElement(A.off,
                            new XAttribute("x", "457200"),
                            new XAttribute("y", "1600200")),
                        new XElement(A.ext,
                            new XAttribute("cx", "8229600"),
                            new XAttribute("cy", "4525963"))),
                    new XElement(A.prstGeom,
                        new XAttribute("prst", "rect"),
                        new XElement(A.avLst))),
                txBody);
        }
    }

    #endregion

    #region Main Comparer Class

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

            var matchedCount = slideMatches.Count(m => m.MatchType == SlideMatchType.Matched);
            Log(settings, $"Matched {matchedCount} slides");

            // 3. Compute diff
            var result = PmlDiffEngine.ComputeDiff(sig1, sig2, slideMatches, settings);

            Log(settings, $"Found {result.TotalChanges} changes");

            return result;
        }

        /// <summary>
        /// Produce a marked presentation highlighting all differences.
        /// The output is based on the newer presentation with visual change overlays.
        /// </summary>
        /// <param name="older">The original/older presentation.</param>
        /// <param name="newer">The revised/newer presentation.</param>
        /// <param name="settings">Comparison settings.</param>
        /// <returns>A new presentation with changes highlighted.</returns>
        public static PmlDocument ProduceMarkedPresentation(
            PmlDocument older,
            PmlDocument newer,
            PmlComparerSettings settings = null)
        {
            settings ??= new PmlComparerSettings();

            Log(settings, "PmlComparer.ProduceMarkedPresentation: Starting");

            var result = Compare(older, newer, settings);
            var marked = PmlMarkupRenderer.RenderMarkedPresentation(newer, result, settings);

            Log(settings, "PmlComparer.ProduceMarkedPresentation: Complete");

            return marked;
        }

        /// <summary>
        /// Get the internal canonical signature of a presentation (for advanced use).
        /// </summary>
        /// <param name="doc">The presentation document.</param>
        /// <param name="settings">Comparison settings.</param>
        /// <returns>The canonical signature of the presentation.</returns>
        public static object Canonicalize(
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

    #endregion
}
