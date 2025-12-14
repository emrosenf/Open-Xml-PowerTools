// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Draw = DocumentFormat.OpenXml.Drawing;

namespace OxPt
{
    /// <summary>
    /// Generates test PPTX files with known differences for PmlComparer testing.
    /// </summary>
    public static class PmlComparerTestFileGenerator
    {
        private static readonly string TestFilesDir = "../../../../TestFiles/";

        /// <summary>
        /// Ensures all PmlComparer test files exist. Creates them if they don't.
        /// </summary>
        public static void EnsureTestFilesExist()
        {
            Directory.CreateDirectory(TestFilesDir);

            // Base presentation with 2 slides
            CreateIfNotExists("PmlComparer-Base.pptx", () => CreateBasePresentation());

            // Same as base (for identical comparison)
            CreateIfNotExists("PmlComparer-Identical.pptx", () => CreateBasePresentation());

            // One extra slide added
            CreateIfNotExists("PmlComparer-SlideAdded.pptx", () => CreatePresentationWithExtraSlide());

            // One slide removed
            CreateIfNotExists("PmlComparer-SlideDeleted.pptx", () => CreatePresentationWithDeletedSlide());

            // Extra shape on slide 1
            CreateIfNotExists("PmlComparer-ShapeAdded.pptx", () => CreatePresentationWithExtraShape());

            // Shape removed from slide 1
            CreateIfNotExists("PmlComparer-ShapeDeleted.pptx", () => CreatePresentationWithDeletedShape());

            // Text modified in shape
            CreateIfNotExists("PmlComparer-TextChanged.pptx", () => CreatePresentationWithChangedText());

            // Shape moved to different position
            CreateIfNotExists("PmlComparer-ShapeMoved.pptx", () => CreatePresentationWithMovedShape());

            // Shape resized
            CreateIfNotExists("PmlComparer-ShapeResized.pptx", () => CreatePresentationWithResizedShape());
        }

        private static void CreateIfNotExists(string fileName, System.Func<byte[]> generator)
        {
            var path = Path.Combine(TestFilesDir, fileName);
            if (!File.Exists(path))
            {
                var bytes = generator();
                File.WriteAllBytes(path, bytes);
            }
        }

        /// <summary>
        /// Creates base presentation with 2 slides, each with a title and content shape.
        /// </summary>
        private static byte[] CreateBasePresentation()
        {
            return CreatePresentation(new[]
            {
                new SlideContent("Slide 1 Title", "Slide 1 Content", 457200, 1600200, 8229600, 4525963),
                new SlideContent("Slide 2 Title", "Slide 2 Content", 457200, 1600200, 8229600, 4525963)
            });
        }

        /// <summary>
        /// Creates presentation with 3 slides (one extra).
        /// </summary>
        private static byte[] CreatePresentationWithExtraSlide()
        {
            return CreatePresentation(new[]
            {
                new SlideContent("Slide 1 Title", "Slide 1 Content", 457200, 1600200, 8229600, 4525963),
                new SlideContent("Slide 2 Title", "Slide 2 Content", 457200, 1600200, 8229600, 4525963),
                new SlideContent("Slide 3 Title", "Slide 3 Content - NEW", 457200, 1600200, 8229600, 4525963)
            });
        }

        /// <summary>
        /// Creates presentation with 1 slide (one deleted).
        /// </summary>
        private static byte[] CreatePresentationWithDeletedSlide()
        {
            return CreatePresentation(new[]
            {
                new SlideContent("Slide 1 Title", "Slide 1 Content", 457200, 1600200, 8229600, 4525963)
            });
        }

        /// <summary>
        /// Creates presentation with extra shape on slide 1.
        /// </summary>
        private static byte[] CreatePresentationWithExtraShape()
        {
            return CreatePresentation(new[]
            {
                new SlideContent("Slide 1 Title", "Slide 1 Content", 457200, 1600200, 8229600, 4525963,
                    extraShape: new ShapeInfo("Extra Shape", "Extra Content", 457200, 5000000, 3000000, 1000000)),
                new SlideContent("Slide 2 Title", "Slide 2 Content", 457200, 1600200, 8229600, 4525963)
            });
        }

        /// <summary>
        /// Creates presentation with content shape removed from slide 1.
        /// </summary>
        private static byte[] CreatePresentationWithDeletedShape()
        {
            return CreatePresentation(new[]
            {
                new SlideContent("Slide 1 Title", null, 457200, 1600200, 8229600, 4525963), // No content shape
                new SlideContent("Slide 2 Title", "Slide 2 Content", 457200, 1600200, 8229600, 4525963)
            });
        }

        /// <summary>
        /// Creates presentation with modified text in slide 1.
        /// </summary>
        private static byte[] CreatePresentationWithChangedText()
        {
            return CreatePresentation(new[]
            {
                new SlideContent("Slide 1 Title", "MODIFIED Content", 457200, 1600200, 8229600, 4525963),
                new SlideContent("Slide 2 Title", "Slide 2 Content", 457200, 1600200, 8229600, 4525963)
            });
        }

        /// <summary>
        /// Creates presentation with shape moved to different position.
        /// </summary>
        private static byte[] CreatePresentationWithMovedShape()
        {
            return CreatePresentation(new[]
            {
                new SlideContent("Slide 1 Title", "Slide 1 Content", 1000000, 2000000, 8229600, 4525963), // Different X, Y
                new SlideContent("Slide 2 Title", "Slide 2 Content", 457200, 1600200, 8229600, 4525963)
            });
        }

        /// <summary>
        /// Creates presentation with shape resized.
        /// </summary>
        private static byte[] CreatePresentationWithResizedShape()
        {
            return CreatePresentation(new[]
            {
                new SlideContent("Slide 1 Title", "Slide 1 Content", 457200, 1600200, 6000000, 3000000), // Different Cx, Cy
                new SlideContent("Slide 2 Title", "Slide 2 Content", 457200, 1600200, 8229600, 4525963)
            });
        }

        private class SlideContent
        {
            public string Title { get; }
            public string Content { get; }
            public long X { get; }
            public long Y { get; }
            public long Cx { get; }
            public long Cy { get; }
            public ShapeInfo ExtraShape { get; }

            public SlideContent(string title, string content, long x, long y, long cx, long cy, ShapeInfo extraShape = null)
            {
                Title = title;
                Content = content;
                X = x;
                Y = y;
                Cx = cx;
                Cy = cy;
                ExtraShape = extraShape;
            }
        }

        private class ShapeInfo
        {
            public string Name { get; }
            public string Text { get; }
            public long X { get; }
            public long Y { get; }
            public long Cx { get; }
            public long Cy { get; }

            public ShapeInfo(string name, string text, long x, long y, long cx, long cy)
            {
                Name = name;
                Text = text;
                X = x;
                Y = y;
                Cx = cx;
                Cy = cy;
            }
        }

        private static byte[] CreatePresentation(SlideContent[] slides)
        {
            using var ms = new MemoryStream();
            using (var doc = PresentationDocument.Create(ms, PresentationDocumentType.Presentation))
            {
                var presentationPart = doc.AddPresentationPart();
                presentationPart.Presentation = new Presentation(
                    new SlideSize { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 },
                    new NotesSize { Cx = 6858000, Cy = 9144000 }
                );

                // Create slide master and layout
                var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
                slideMasterPart.SlideMaster = CreateSlideMaster();

                var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
                slideLayoutPart.SlideLayout = CreateSlideLayout();

                slideMasterPart.SlideMaster.SlideLayoutIdList = new SlideLayoutIdList(
                    new SlideLayoutId { Id = 2147483649U, RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart) }
                );

                // Add theme
                var themePart = slideMasterPart.AddNewPart<ThemePart>();
                themePart.Theme = CreateMinimalTheme();

                // Create master id list
                presentationPart.Presentation.SlideMasterIdList = new SlideMasterIdList(
                    new SlideMasterId { Id = 2147483648U, RelationshipId = presentationPart.GetIdOfPart(slideMasterPart) }
                );

                // Create slides
                presentationPart.Presentation.SlideIdList = new SlideIdList();
                uint slideId = 256;

                foreach (var slideContent in slides)
                {
                    var slidePart = presentationPart.AddNewPart<SlidePart>();
                    slidePart.Slide = CreateSlide(slideContent);
                    slidePart.AddPart(slideLayoutPart);

                    presentationPart.Presentation.SlideIdList.Append(new SlideId
                    {
                        Id = slideId++,
                        RelationshipId = presentationPart.GetIdOfPart(slidePart)
                    });
                }

                presentationPart.Presentation.Save();
            }

            return ms.ToArray();
        }

        private static SlideMaster CreateSlideMaster()
        {
            return new SlideMaster(
                new CommonSlideData(
                    new ShapeTree(
                        new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties { Id = 1U, Name = "" },
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new Draw.GroupShapeProperties())));
        }

        private static SlideLayout CreateSlideLayout()
        {
            return new SlideLayout(
                new CommonSlideData(
                    new ShapeTree(
                        new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties { Id = 1U, Name = "" },
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new Draw.GroupShapeProperties())))
            { Type = SlideLayoutValues.Title };
        }

        private static Draw.Theme CreateMinimalTheme()
        {
            return new Draw.Theme(
                new Draw.ThemeElements(
                    new Draw.ColorScheme(
                        new Draw.Dark1Color(new Draw.SystemColor { Val = Draw.SystemColorValues.WindowText }),
                        new Draw.Light1Color(new Draw.SystemColor { Val = Draw.SystemColorValues.Window }),
                        new Draw.Dark2Color(new Draw.RgbColorModelHex { Val = "1F497D" }),
                        new Draw.Light2Color(new Draw.RgbColorModelHex { Val = "EEECE1" }),
                        new Draw.Accent1Color(new Draw.RgbColorModelHex { Val = "4F81BD" }),
                        new Draw.Accent2Color(new Draw.RgbColorModelHex { Val = "C0504D" }),
                        new Draw.Accent3Color(new Draw.RgbColorModelHex { Val = "9BBB59" }),
                        new Draw.Accent4Color(new Draw.RgbColorModelHex { Val = "8064A2" }),
                        new Draw.Accent5Color(new Draw.RgbColorModelHex { Val = "4BACC6" }),
                        new Draw.Accent6Color(new Draw.RgbColorModelHex { Val = "F79646" }),
                        new Draw.Hyperlink(new Draw.RgbColorModelHex { Val = "0000FF" }),
                        new Draw.FollowedHyperlinkColor(new Draw.RgbColorModelHex { Val = "800080" }))
                    { Name = "Office" },
                    new Draw.FontScheme(
                        new Draw.MajorFont(
                            new Draw.LatinFont { Typeface = "Calibri" },
                            new Draw.EastAsianFont { Typeface = "" },
                            new Draw.ComplexScriptFont { Typeface = "" }),
                        new Draw.MinorFont(
                            new Draw.LatinFont { Typeface = "Calibri" },
                            new Draw.EastAsianFont { Typeface = "" },
                            new Draw.ComplexScriptFont { Typeface = "" }))
                    { Name = "Office" },
                    new Draw.FormatScheme(
                        new Draw.FillStyleList(
                            new Draw.SolidFill(new Draw.SchemeColor { Val = Draw.SchemeColorValues.PhColor }),
                            new Draw.SolidFill(new Draw.SchemeColor { Val = Draw.SchemeColorValues.PhColor }),
                            new Draw.SolidFill(new Draw.SchemeColor { Val = Draw.SchemeColorValues.PhColor })),
                        new Draw.LineStyleList(
                            new Draw.Outline(),
                            new Draw.Outline(),
                            new Draw.Outline()),
                        new Draw.EffectStyleList(
                            new Draw.EffectStyle(new Draw.EffectList()),
                            new Draw.EffectStyle(new Draw.EffectList()),
                            new Draw.EffectStyle(new Draw.EffectList())),
                        new Draw.BackgroundFillStyleList(
                            new Draw.SolidFill(new Draw.SchemeColor { Val = Draw.SchemeColorValues.PhColor }),
                            new Draw.SolidFill(new Draw.SchemeColor { Val = Draw.SchemeColorValues.PhColor }),
                            new Draw.SolidFill(new Draw.SchemeColor { Val = Draw.SchemeColorValues.PhColor })))
                    { Name = "Office" }))
            { Name = "Office Theme" };
        }

        private static Slide CreateSlide(SlideContent content)
        {
            var shapeTree = new ShapeTree(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties { Id = 1U, Name = "" },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new Draw.GroupShapeProperties(
                    new Draw.TransformGroup(
                        new Draw.Offset { X = 0, Y = 0 },
                        new Draw.Extents { Cx = 0, Cy = 0 },
                        new Draw.ChildOffset { X = 0, Y = 0 },
                        new Draw.ChildExtents { Cx = 0, Cy = 0 })));

            uint shapeId = 2;

            // Add title shape
            if (!string.IsNullOrEmpty(content.Title))
            {
                shapeTree.Append(CreateTextShape(shapeId++, "Title", content.Title, 457200, 274638, 8229600, 1143000));
            }

            // Add content shape
            if (!string.IsNullOrEmpty(content.Content))
            {
                shapeTree.Append(CreateTextShape(shapeId++, "Content", content.Content, content.X, content.Y, content.Cx, content.Cy));
            }

            // Add extra shape if specified
            if (content.ExtraShape != null)
            {
                var extra = content.ExtraShape;
                shapeTree.Append(CreateTextShape(shapeId++, extra.Name, extra.Text, extra.X, extra.Y, extra.Cx, extra.Cy));
            }

            return new Slide(new CommonSlideData(shapeTree));
        }

        private static Shape CreateTextShape(uint id, string name, string text, long x, long y, long cx, long cy)
        {
            return new Shape(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = id, Name = name },
                    new NonVisualShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new ShapeProperties(
                    new Draw.Transform2D(
                        new Draw.Offset { X = x, Y = y },
                        new Draw.Extents { Cx = cx, Cy = cy }),
                    new Draw.PresetGeometry(new Draw.AdjustValueList()) { Preset = Draw.ShapeTypeValues.Rectangle }),
                new TextBody(
                    new Draw.BodyProperties(),
                    new Draw.ListStyle(),
                    new Draw.Paragraph(
                        new Draw.Run(
                            new Draw.RunProperties { Language = "en-US" },
                            new Draw.Text(text)),
                        new Draw.EndParagraphRunProperties { Language = "en-US" })));
        }
    }
}
