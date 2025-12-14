// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OpenXmlPowerTools;
using Xunit;

// Aliases to resolve ambiguous references (Pres/Draw to avoid conflict with OpenXmlPowerTools.P/A)
using Pres = DocumentFormat.OpenXml.Presentation;
using Draw = DocumentFormat.OpenXml.Drawing;

#if !ELIDE_XUNIT_TESTS

namespace OxPt
{
    public class PmlComparerTests
    {
        #region Test Helpers

        /// <summary>
        /// Helper to create a simple test presentation with specified slides.
        /// Each slide contains a title and optional content shapes.
        /// </summary>
        private static PmlDocument CreateTestPresentation(List<SlideData> slides)
        {
            using var ms = new MemoryStream();
            using (var doc = PresentationDocument.Create(ms, PresentationDocumentType.Presentation))
            {
                // Create presentation part
                var presentationPart = doc.AddPresentationPart();
                presentationPart.Presentation = new Presentation();

                // Set slide size
                presentationPart.Presentation.SlideSize = new SlideSize
                {
                    Cx = 9144000,
                    Cy = 6858000,
                    Type = SlideSizeValues.Screen4x3
                };

                presentationPart.Presentation.NotesSize = new NotesSize
                {
                    Cx = 6858000,
                    Cy = 9144000
                };

                // Create slide master and layout
                var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
                slideMasterPart.SlideMaster = CreateSlideMaster();

                var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
                slideLayoutPart.SlideLayout = CreateSlideLayout();

                // Link layout to master
                slideMasterPart.SlideMaster.SlideLayoutIdList = new SlideLayoutIdList(
                    new SlideLayoutId
                    {
                        Id = 2147483649U,
                        RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart)
                    });

                // Add theme
                var themePart = slideMasterPart.AddNewPart<ThemePart>();
                themePart.Theme = CreateTheme();

                // Create slide id list and master id list
                presentationPart.Presentation.SlideMasterIdList = new SlideMasterIdList(
                    new SlideMasterId
                    {
                        Id = 2147483648U,
                        RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
                    });

                presentationPart.Presentation.SlideIdList = new SlideIdList();

                uint slideId = 256;
                foreach (var slideData in slides)
                {
                    var slidePart = presentationPart.AddNewPart<SlidePart>();
                    slidePart.Slide = CreateSlide(slideData);

                    // Link slide to layout
                    slidePart.AddPart(slideLayoutPart);

                    presentationPart.Presentation.SlideIdList.Append(new SlideId
                    {
                        Id = slideId++,
                        RelationshipId = presentationPart.GetIdOfPart(slidePart)
                    });
                }

                presentationPart.Presentation.Save();
            }

            return new PmlDocument("test.pptx", ms.ToArray());
        }

        private static SlideMaster CreateSlideMaster()
        {
            return new SlideMaster(
                new CommonSlideData(
                    new ShapeTree(
                        new Pres.NonVisualGroupShapeProperties(
                            new Pres.NonVisualDrawingProperties { Id = 1U, Name = "" },
                            new Pres.NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new Draw.GroupShapeProperties())));
        }

        private static SlideLayout CreateSlideLayout()
        {
            return new SlideLayout(
                new CommonSlideData(
                    new ShapeTree(
                        new Pres.NonVisualGroupShapeProperties(
                            new Pres.NonVisualDrawingProperties { Id = 1U, Name = "" },
                            new Pres.NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new Draw.GroupShapeProperties())))
            {
                Type = SlideLayoutValues.Title
            };
        }

        private static Draw.Theme CreateTheme()
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
                            new Draw.GradientFill(new Draw.GradientStopList()),
                            new Draw.GradientFill(new Draw.GradientStopList())),
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
                            new Draw.GradientFill(new Draw.GradientStopList()),
                            new Draw.GradientFill(new Draw.GradientStopList())))
                    { Name = "Office" }))
            { Name = "Office Theme" };
        }

        private static Slide CreateSlide(SlideData data)
        {
            var shapeTree = new ShapeTree(
                new Pres.NonVisualGroupShapeProperties(
                    new Pres.NonVisualDrawingProperties { Id = 1U, Name = "" },
                    new Pres.NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new Draw.GroupShapeProperties(
                    new Draw.TransformGroup(
                        new Draw.Offset { X = 0, Y = 0 },
                        new Draw.Extents { Cx = 0, Cy = 0 },
                        new Draw.ChildOffset { X = 0, Y = 0 },
                        new Draw.ChildExtents { Cx = 0, Cy = 0 })));

            uint shapeId = 2;

            // Add title shape
            if (!string.IsNullOrEmpty(data.Title))
            {
                shapeTree.Append(CreateTextShape(
                    shapeId++,
                    "Title",
                    data.Title,
                    457200, 274638, 8229600, 1143000,
                    "title"));
            }

            // Add content shapes
            if (data.Shapes != null)
            {
                foreach (var shape in data.Shapes)
                {
                    shapeTree.Append(CreateTextShape(
                        shapeId++,
                        shape.Name,
                        shape.Text,
                        shape.X,
                        shape.Y,
                        shape.Cx,
                        shape.Cy,
                        shape.PlaceholderType));
                }
            }

            return new Slide(new CommonSlideData(shapeTree));
        }

        private static Pres.Shape CreateTextShape(
            uint id,
            string name,
            string text,
            long x,
            long y,
            long cx,
            long cy,
            string placeholderType = null)
        {
            var nvSpPr = new Pres.NonVisualShapeProperties(
                new Pres.NonVisualDrawingProperties { Id = id, Name = name },
                new Pres.NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties());

            if (!string.IsNullOrEmpty(placeholderType))
            {
                nvSpPr.ApplicationNonVisualDrawingProperties.Append(
                    new PlaceholderShape { Type = GetPlaceholderType(placeholderType) });
            }

            var spPr = new Pres.ShapeProperties(
                new Draw.Transform2D(
                    new Draw.Offset { X = x, Y = y },
                    new Draw.Extents { Cx = cx, Cy = cy }),
                new Draw.PresetGeometry(new Draw.AdjustValueList()) { Preset = Draw.ShapeTypeValues.Rectangle });

            var txBody = new Pres.TextBody(
                new Draw.BodyProperties(),
                new Draw.ListStyle(),
                new Draw.Paragraph(
                    new Draw.Run(
                        new Draw.RunProperties { Language = "en-US" },
                        new Draw.Text(text)),
                    new Draw.EndParagraphRunProperties { Language = "en-US" }));

            return new Pres.Shape(nvSpPr, spPr, txBody);
        }

        private static PlaceholderValues GetPlaceholderType(string type)
        {
            return type?.ToLower() switch
            {
                "title" => PlaceholderValues.Title,
                "ctrtitle" => PlaceholderValues.CenteredTitle,
                "body" => PlaceholderValues.Body,
                "subtitle" => PlaceholderValues.SubTitle,
                _ => PlaceholderValues.Body
            };
        }

        private class SlideData
        {
            public string Title { get; set; }
            public List<ShapeData> Shapes { get; set; }
        }

        private class ShapeData
        {
            public string Name { get; set; }
            public string Text { get; set; }
            public long X { get; set; } = 457200;
            public long Y { get; set; } = 1600200;
            public long Cx { get; set; } = 8229600;
            public long Cy { get; set; } = 4525963;
            public string PlaceholderType { get; set; }
        }

        #endregion

        #region Basic Comparison Tests

        [Fact]
        public void PC001_IdenticalPresentations_NoChanges()
        {
            // Arrange
            var slides = new List<SlideData>
            {
                new SlideData
                {
                    Title = "Test Slide",
                    Shapes = new List<ShapeData>
                    {
                        new ShapeData { Name = "Content", Text = "Hello World" }
                    }
                }
            };

            var doc1 = CreateTestPresentation(slides);
            var doc2 = CreateTestPresentation(slides);
            var settings = new PmlComparerSettings();

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.NotNull(result);
            Assert.Equal(0, result.TotalChanges);
        }

        [Fact]
        public void PC002_SlideAdded_DetectsInsertion()
        {
            // Arrange
            var slides1 = new List<SlideData>
            {
                new SlideData { Title = "Slide 1" }
            };

            var slides2 = new List<SlideData>
            {
                new SlideData { Title = "Slide 1" },
                new SlideData { Title = "Slide 2" }
            };

            var doc1 = CreateTestPresentation(slides1);
            var doc2 = CreateTestPresentation(slides2);
            var settings = new PmlComparerSettings();

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.NotNull(result);
            Assert.Equal(1, result.SlidesInserted);
            Assert.Contains(result.Changes, c => c.ChangeType == PmlChangeType.SlideInserted);
        }

        [Fact]
        public void PC003_SlideDeleted_DetectsDeletion()
        {
            // Arrange
            var slides1 = new List<SlideData>
            {
                new SlideData { Title = "Slide 1" },
                new SlideData { Title = "Slide 2" }
            };

            var slides2 = new List<SlideData>
            {
                new SlideData { Title = "Slide 1" }
            };

            var doc1 = CreateTestPresentation(slides1);
            var doc2 = CreateTestPresentation(slides2);
            var settings = new PmlComparerSettings();

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.NotNull(result);
            Assert.Equal(1, result.SlidesDeleted);
            Assert.Contains(result.Changes, c => c.ChangeType == PmlChangeType.SlideDeleted);
        }

        [Fact]
        public void PC004_SlidesReordered_DetectsMove()
        {
            // Arrange
            var slides1 = new List<SlideData>
            {
                new SlideData { Title = "First Slide" },
                new SlideData { Title = "Second Slide" }
            };

            var slides2 = new List<SlideData>
            {
                new SlideData { Title = "Second Slide" },
                new SlideData { Title = "First Slide" }
            };

            var doc1 = CreateTestPresentation(slides1);
            var doc2 = CreateTestPresentation(slides2);
            var settings = new PmlComparerSettings();

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.NotNull(result);
            Assert.True(result.SlidesMoved > 0, "Should detect moved slides");
        }

        #endregion

        #region Shape Tests

        [Fact]
        public void PC005_ShapeAdded_DetectsInsertion()
        {
            // Arrange
            var slides1 = new List<SlideData>
            {
                new SlideData
                {
                    Title = "Test Slide",
                    Shapes = new List<ShapeData>
                    {
                        new ShapeData { Name = "Shape1", Text = "Original" }
                    }
                }
            };

            var slides2 = new List<SlideData>
            {
                new SlideData
                {
                    Title = "Test Slide",
                    Shapes = new List<ShapeData>
                    {
                        new ShapeData { Name = "Shape1", Text = "Original" },
                        new ShapeData { Name = "Shape2", Text = "New Shape" }
                    }
                }
            };

            var doc1 = CreateTestPresentation(slides1);
            var doc2 = CreateTestPresentation(slides2);
            var settings = new PmlComparerSettings();

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.NotNull(result);
            Assert.Equal(1, result.ShapesInserted);
            var insertedChange = result.Changes.FirstOrDefault(c => c.ChangeType == PmlChangeType.ShapeInserted);
            Assert.NotNull(insertedChange);
            Assert.Equal("Shape2", insertedChange.ShapeName);
        }

        [Fact]
        public void PC006_ShapeDeleted_DetectsDeletion()
        {
            // Arrange
            var slides1 = new List<SlideData>
            {
                new SlideData
                {
                    Title = "Test Slide",
                    Shapes = new List<ShapeData>
                    {
                        new ShapeData { Name = "Shape1", Text = "Keep" },
                        new ShapeData { Name = "Shape2", Text = "Delete Me" }
                    }
                }
            };

            var slides2 = new List<SlideData>
            {
                new SlideData
                {
                    Title = "Test Slide",
                    Shapes = new List<ShapeData>
                    {
                        new ShapeData { Name = "Shape1", Text = "Keep" }
                    }
                }
            };

            var doc1 = CreateTestPresentation(slides1);
            var doc2 = CreateTestPresentation(slides2);
            var settings = new PmlComparerSettings();

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.NotNull(result);
            Assert.Equal(1, result.ShapesDeleted);
            var deletedChange = result.Changes.FirstOrDefault(c => c.ChangeType == PmlChangeType.ShapeDeleted);
            Assert.NotNull(deletedChange);
            Assert.Equal("Shape2", deletedChange.ShapeName);
        }

        [Fact]
        public void PC007_ShapeMoved_DetectsMove()
        {
            // Arrange
            var slides1 = new List<SlideData>
            {
                new SlideData
                {
                    Title = "Test Slide",
                    Shapes = new List<ShapeData>
                    {
                        new ShapeData { Name = "MovingShape", Text = "Content", X = 100000, Y = 100000 }
                    }
                }
            };

            var slides2 = new List<SlideData>
            {
                new SlideData
                {
                    Title = "Test Slide",
                    Shapes = new List<ShapeData>
                    {
                        new ShapeData { Name = "MovingShape", Text = "Content", X = 500000, Y = 500000 }
                    }
                }
            };

            var doc1 = CreateTestPresentation(slides1);
            var doc2 = CreateTestPresentation(slides2);
            var settings = new PmlComparerSettings();

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.NotNull(result);
            Assert.Equal(1, result.ShapesMoved);
            var movedChange = result.Changes.FirstOrDefault(c => c.ChangeType == PmlChangeType.ShapeMoved);
            Assert.NotNull(movedChange);
            Assert.Equal("MovingShape", movedChange.ShapeName);
        }

        [Fact]
        public void PC008_ShapeResized_DetectsResize()
        {
            // Arrange
            var slides1 = new List<SlideData>
            {
                new SlideData
                {
                    Title = "Test Slide",
                    Shapes = new List<ShapeData>
                    {
                        new ShapeData { Name = "ResizingShape", Text = "Content", Cx = 1000000, Cy = 500000 }
                    }
                }
            };

            var slides2 = new List<SlideData>
            {
                new SlideData
                {
                    Title = "Test Slide",
                    Shapes = new List<ShapeData>
                    {
                        new ShapeData { Name = "ResizingShape", Text = "Content", Cx = 2000000, Cy = 1000000 }
                    }
                }
            };

            var doc1 = CreateTestPresentation(slides1);
            var doc2 = CreateTestPresentation(slides2);
            var settings = new PmlComparerSettings();

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.NotNull(result);
            Assert.Equal(1, result.ShapesResized);
            var resizedChange = result.Changes.FirstOrDefault(c => c.ChangeType == PmlChangeType.ShapeResized);
            Assert.NotNull(resizedChange);
            Assert.Equal("ResizingShape", resizedChange.ShapeName);
        }

        #endregion

        #region Text Content Tests

        [Fact]
        public void PC009_TextChanged_DetectsChange()
        {
            // Arrange
            var slides1 = new List<SlideData>
            {
                new SlideData
                {
                    Title = "Test Slide",
                    Shapes = new List<ShapeData>
                    {
                        new ShapeData { Name = "TextShape", Text = "Original Text" }
                    }
                }
            };

            var slides2 = new List<SlideData>
            {
                new SlideData
                {
                    Title = "Test Slide",
                    Shapes = new List<ShapeData>
                    {
                        new ShapeData { Name = "TextShape", Text = "Modified Text" }
                    }
                }
            };

            var doc1 = CreateTestPresentation(slides1);
            var doc2 = CreateTestPresentation(slides2);
            var settings = new PmlComparerSettings();

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.NotNull(result);
            Assert.Equal(1, result.TextChanges);
            var textChange = result.Changes.FirstOrDefault(c => c.ChangeType == PmlChangeType.TextChanged);
            Assert.NotNull(textChange);
            Assert.Equal("Original Text", textChange.OldValue);
            Assert.Equal("Modified Text", textChange.NewValue);
        }

        [Fact]
        public void PC010_TitleChanged_DetectsChange()
        {
            // Arrange
            var slides1 = new List<SlideData>
            {
                new SlideData { Title = "Original Title" }
            };

            var slides2 = new List<SlideData>
            {
                new SlideData { Title = "New Title" }
            };

            var doc1 = CreateTestPresentation(slides1);
            var doc2 = CreateTestPresentation(slides2);
            var settings = new PmlComparerSettings();

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.NotNull(result);
            // Slides won't match by title, so we should see changes
            Assert.True(result.TotalChanges > 0);
        }

        #endregion

        #region Multiple Changes Tests

        [Fact]
        public void PC011_MultipleChanges_DetectsAll()
        {
            // Arrange
            var slides1 = new List<SlideData>
            {
                new SlideData
                {
                    Title = "Slide 1",
                    Shapes = new List<ShapeData>
                    {
                        new ShapeData { Name = "Shape1", Text = "Text 1" },
                        new ShapeData { Name = "Shape2", Text = "Text 2" }
                    }
                },
                new SlideData { Title = "Slide 2" }
            };

            var slides2 = new List<SlideData>
            {
                new SlideData
                {
                    Title = "Slide 1",
                    Shapes = new List<ShapeData>
                    {
                        new ShapeData { Name = "Shape1", Text = "Modified Text 1" },
                        new ShapeData { Name = "Shape3", Text = "New Shape" }
                    }
                },
                new SlideData { Title = "Slide 2" },
                new SlideData { Title = "Slide 3" }
            };

            var doc1 = CreateTestPresentation(slides1);
            var doc2 = CreateTestPresentation(slides2);
            var settings = new PmlComparerSettings();

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.NotNull(result);
            Assert.True(result.TotalChanges >= 3, $"Expected at least 3 changes, got {result.TotalChanges}");
            Assert.Equal(1, result.SlidesInserted);  // Slide 3 added
            Assert.Equal(1, result.ShapesDeleted);   // Shape2 deleted
            Assert.Equal(1, result.ShapesInserted);  // Shape3 added
        }

        #endregion

        #region Settings Tests

        [Fact]
        public void PC012_CompareShapeStructureDisabled_IgnoresShapeChanges()
        {
            // Arrange
            var slides1 = new List<SlideData>
            {
                new SlideData
                {
                    Title = "Test Slide",
                    Shapes = new List<ShapeData>
                    {
                        new ShapeData { Name = "Shape1", Text = "Original" }
                    }
                }
            };

            var slides2 = new List<SlideData>
            {
                new SlideData
                {
                    Title = "Test Slide",
                    Shapes = new List<ShapeData>
                    {
                        new ShapeData { Name = "Shape1", Text = "Original" },
                        new ShapeData { Name = "Shape2", Text = "New" }
                    }
                }
            };

            var doc1 = CreateTestPresentation(slides1);
            var doc2 = CreateTestPresentation(slides2);
            var settings = new PmlComparerSettings
            {
                CompareShapeStructure = false
            };

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.NotNull(result);
            Assert.Equal(0, result.ShapesInserted);
            Assert.Equal(0, result.ShapesDeleted);
        }

        [Fact]
        public void PC013_CompareTextContentDisabled_IgnoresTextChanges()
        {
            // Arrange
            var slides1 = new List<SlideData>
            {
                new SlideData
                {
                    Title = "Test Slide",
                    Shapes = new List<ShapeData>
                    {
                        new ShapeData { Name = "TextShape", Text = "Original" }
                    }
                }
            };

            var slides2 = new List<SlideData>
            {
                new SlideData
                {
                    Title = "Test Slide",
                    Shapes = new List<ShapeData>
                    {
                        new ShapeData { Name = "TextShape", Text = "Modified" }
                    }
                }
            };

            var doc1 = CreateTestPresentation(slides1);
            var doc2 = CreateTestPresentation(slides2);
            var settings = new PmlComparerSettings
            {
                CompareTextContent = false
            };

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.NotNull(result);
            Assert.Equal(0, result.TextChanges);
        }

        #endregion

        #region Edge Cases

        [Fact]
        public void PC014_EmptyPresentation_NoErrors()
        {
            // Arrange
            var doc1 = CreateTestPresentation(new List<SlideData>());
            var doc2 = CreateTestPresentation(new List<SlideData>());
            var settings = new PmlComparerSettings();

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.NotNull(result);
            Assert.Equal(0, result.TotalChanges);
        }

        [Fact]
        public void PC015_SingleSlide_Works()
        {
            // Arrange
            var slides = new List<SlideData>
            {
                new SlideData { Title = "Only Slide" }
            };

            var doc1 = CreateTestPresentation(slides);
            var doc2 = CreateTestPresentation(slides);
            var settings = new PmlComparerSettings();

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.NotNull(result);
            Assert.Equal(0, result.TotalChanges);
        }

        [Fact]
        public void PC016_NullSettings_UsesDefaults()
        {
            // Arrange
            var slides = new List<SlideData>
            {
                new SlideData { Title = "Test" }
            };

            var doc1 = CreateTestPresentation(slides);
            var doc2 = CreateTestPresentation(slides);

            // Act
            var result = PmlComparer.Compare(doc1, doc2, null);

            // Assert
            Assert.NotNull(result);
        }

        [Fact]
        public void PC017_NullDocument_ThrowsException()
        {
            // Arrange
            var doc = CreateTestPresentation(new List<SlideData>
            {
                new SlideData { Title = "Test" }
            });

            // Act & Assert
            Assert.Throws<ArgumentNullException>(() => PmlComparer.Compare(null, doc));
            Assert.Throws<ArgumentNullException>(() => PmlComparer.Compare(doc, null));
        }

        #endregion

        #region JSON Output Tests

        [Fact]
        public void PC018_ToJson_ProducesValidJson()
        {
            // Arrange
            var slides1 = new List<SlideData>
            {
                new SlideData { Title = "Slide 1" }
            };

            var slides2 = new List<SlideData>
            {
                new SlideData { Title = "Slide 1" },
                new SlideData { Title = "Slide 2" }
            };

            var doc1 = CreateTestPresentation(slides1);
            var doc2 = CreateTestPresentation(slides2);
            var settings = new PmlComparerSettings();

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);
            var json = result.ToJson();

            // Assert
            Assert.NotNull(json);
            Assert.Contains("TotalChanges", json);
            Assert.Contains("SlidesInserted", json);
            Assert.Contains("Changes", json);
        }

        #endregion

        #region Marked Presentation Tests

        [Fact]
        public void PC019_ProduceMarkedPresentation_CreatesValidDocument()
        {
            // Arrange
            var slides1 = new List<SlideData>
            {
                new SlideData
                {
                    Title = "Test Slide",
                    Shapes = new List<ShapeData>
                    {
                        new ShapeData { Name = "Shape1", Text = "Original" }
                    }
                }
            };

            var slides2 = new List<SlideData>
            {
                new SlideData
                {
                    Title = "Test Slide",
                    Shapes = new List<ShapeData>
                    {
                        new ShapeData { Name = "Shape1", Text = "Modified" },
                        new ShapeData { Name = "Shape2", Text = "New Shape" }
                    }
                }
            };

            var doc1 = CreateTestPresentation(slides1);
            var doc2 = CreateTestPresentation(slides2);
            var settings = new PmlComparerSettings
            {
                AddSummarySlide = true,
                AddNotesAnnotations = true
            };

            // Act
            var marked = PmlComparer.ProduceMarkedPresentation(doc1, doc2, settings);

            // Assert
            Assert.NotNull(marked);
            Assert.NotNull(marked.DocumentByteArray);
            Assert.True(marked.DocumentByteArray.Length > 0);

            // Verify it's a valid presentation
            using var ms = new MemoryStream(marked.DocumentByteArray);
            using var pDoc = PresentationDocument.Open(ms, false);
            Assert.NotNull(pDoc.PresentationPart);
            Assert.NotNull(pDoc.PresentationPart.Presentation);

            // Should have at least 2 slides (original + summary)
            var slideCount = pDoc.PresentationPart.Presentation.SlideIdList?.Count() ?? 0;
            Assert.True(slideCount >= 2, $"Expected at least 2 slides, got {slideCount}");
        }

        [Fact]
        public void PC020_ProduceMarkedPresentation_NoChanges_ReturnsOriginal()
        {
            // Arrange
            var slides = new List<SlideData>
            {
                new SlideData { Title = "Test Slide" }
            };

            var doc1 = CreateTestPresentation(slides);
            var doc2 = CreateTestPresentation(slides);
            var settings = new PmlComparerSettings();

            // Act
            var marked = PmlComparer.ProduceMarkedPresentation(doc1, doc2, settings);

            // Assert
            Assert.NotNull(marked);
            // When no changes, should return the original document
            Assert.Equal(doc2.DocumentByteArray.Length, marked.DocumentByteArray.Length);
        }

        #endregion

        #region Logging Tests

        [Fact]
        public void PC021_LogCallback_ReceivesMessages()
        {
            // Arrange
            var logMessages = new List<string>();
            var slides1 = new List<SlideData>
            {
                new SlideData { Title = "Slide 1" }
            };

            var slides2 = new List<SlideData>
            {
                new SlideData { Title = "Slide 2" }
            };

            var doc1 = CreateTestPresentation(slides1);
            var doc2 = CreateTestPresentation(slides2);
            var settings = new PmlComparerSettings
            {
                LogCallback = msg => logMessages.Add(msg)
            };

            // Act
            PmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.NotEmpty(logMessages);
            Assert.Contains(logMessages, m => m.Contains("Starting comparison"));
        }

        #endregion

        #region Query Tests

        [Fact]
        public void PC022_GetChangesBySlide_ReturnsCorrectChanges()
        {
            // Arrange
            var slides1 = new List<SlideData>
            {
                new SlideData
                {
                    Title = "Slide 1",
                    Shapes = new List<ShapeData>
                    {
                        new ShapeData { Name = "Shape1", Text = "Text" }
                    }
                },
                new SlideData
                {
                    Title = "Slide 2",
                    Shapes = new List<ShapeData>
                    {
                        new ShapeData { Name = "Shape2", Text = "Text" }
                    }
                }
            };

            var slides2 = new List<SlideData>
            {
                new SlideData
                {
                    Title = "Slide 1",
                    Shapes = new List<ShapeData>
                    {
                        new ShapeData { Name = "Shape1", Text = "Modified" }
                    }
                },
                new SlideData
                {
                    Title = "Slide 2",
                    Shapes = new List<ShapeData>
                    {
                        new ShapeData { Name = "Shape2", Text = "Text" }
                    }
                }
            };

            var doc1 = CreateTestPresentation(slides1);
            var doc2 = CreateTestPresentation(slides2);
            var settings = new PmlComparerSettings();

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);
            var slide1Changes = result.GetChangesBySlide(1).ToList();
            var slide2Changes = result.GetChangesBySlide(2).ToList();

            // Assert
            Assert.True(slide1Changes.Count > 0, "Slide 1 should have changes");
            Assert.True(slide2Changes.Count == 0, "Slide 2 should have no changes");
        }

        [Fact]
        public void PC023_GetChangesByType_ReturnsCorrectChanges()
        {
            // Arrange
            var slides1 = new List<SlideData>
            {
                new SlideData
                {
                    Title = "Test Slide",
                    Shapes = new List<ShapeData>
                    {
                        new ShapeData { Name = "Shape1", Text = "Original" }
                    }
                }
            };

            var slides2 = new List<SlideData>
            {
                new SlideData
                {
                    Title = "Test Slide",
                    Shapes = new List<ShapeData>
                    {
                        new ShapeData { Name = "Shape1", Text = "Original" },
                        new ShapeData { Name = "NewShape", Text = "New" }
                    }
                }
            };

            var doc1 = CreateTestPresentation(slides1);
            var doc2 = CreateTestPresentation(slides2);
            var settings = new PmlComparerSettings();

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);
            var insertedChanges = result.GetChangesByType(PmlChangeType.ShapeInserted).ToList();
            var deletedChanges = result.GetChangesByType(PmlChangeType.ShapeDeleted).ToList();

            // Assert
            Assert.Single(insertedChanges);
            Assert.Empty(deletedChanges);
        }

        #endregion

        #region Placeholder Tests

        [Fact]
        public void PC024_PlaceholderMatching_MatchesByPlaceholder()
        {
            // Arrange
            var slides1 = new List<SlideData>
            {
                new SlideData
                {
                    Title = "Test",
                    Shapes = new List<ShapeData>
                    {
                        new ShapeData { Name = "Content1", Text = "Body Text", PlaceholderType = "body" }
                    }
                }
            };

            var slides2 = new List<SlideData>
            {
                new SlideData
                {
                    Title = "Test",
                    Shapes = new List<ShapeData>
                    {
                        new ShapeData { Name = "Content1", Text = "Modified Body", PlaceholderType = "body" }
                    }
                }
            };

            var doc1 = CreateTestPresentation(slides1);
            var doc2 = CreateTestPresentation(slides2);
            var settings = new PmlComparerSettings();

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.NotNull(result);
            // Should detect text change, not shape insert/delete
            Assert.Equal(0, result.ShapesInserted);
            Assert.Equal(0, result.ShapesDeleted);
            Assert.Equal(1, result.TextChanges);
        }

        #endregion
    }
}

#endif
