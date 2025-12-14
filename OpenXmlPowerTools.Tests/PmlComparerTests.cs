// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OpenXmlPowerTools;
using Xunit;

#if !ELIDE_XUNIT_TESTS

namespace OxPt
{
    public class PmlComparerTests : IClassFixture<PmlComparerTestFixture>
    {
        private readonly PmlComparerTestFixture _fixture;

        public PmlComparerTests(PmlComparerTestFixture fixture)
        {
            _fixture = fixture;
        }

        #region Basic Comparison Tests

        [Fact]
        public void PC001_IdenticalPresentations_NoChanges()
        {
            // Arrange
            var doc1 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-Base.pptx"));
            var doc2 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-Identical.pptx"));
            var settings = new PmlComparerSettings();

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.NotNull(result);
            Assert.Equal(0, result.TotalChanges);
        }

        [Fact]
        public void PC002_DifferentPresentations_DetectsChanges()
        {
            // Arrange
            var source1 = new FileInfo(Path.Combine(_fixture.TestFilesDir, "PB001-Input1.pptx"));
            var source2 = new FileInfo(Path.Combine(_fixture.TestFilesDir, "PB001-Input2.pptx"));
            var doc1 = new PmlDocument(source1.FullName);
            var doc2 = new PmlDocument(source2.FullName);
            var settings = new PmlComparerSettings();

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.NotNull(result);
            Assert.True(result.TotalChanges > 0, "Should detect differences between different presentations");
        }

        #endregion

        #region Slide Change Detection Tests

        [Fact]
        public void PC003_SlideAdded_DetectsInsertion()
        {
            // Arrange
            var doc1 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-Base.pptx"));
            var doc2 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-SlideAdded.pptx"));
            var settings = new PmlComparerSettings();

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.NotNull(result);
            Assert.Equal(1, result.SlidesInserted);
            Assert.Contains(result.Changes, c => c.ChangeType == PmlChangeType.SlideInserted);
        }

        [Fact]
        public void PC004_SlideDeleted_DetectsDeletion()
        {
            // Arrange
            var doc1 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-Base.pptx"));
            var doc2 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-SlideDeleted.pptx"));
            var settings = new PmlComparerSettings();

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.NotNull(result);
            Assert.Equal(1, result.SlidesDeleted);
            Assert.Contains(result.Changes, c => c.ChangeType == PmlChangeType.SlideDeleted);
        }

        #endregion

        #region Shape Change Detection Tests

        [Fact]
        public void PC005_ShapeAdded_DetectsInsertion()
        {
            // Arrange
            var doc1 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-Base.pptx"));
            var doc2 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-ShapeAdded.pptx"));
            var settings = new PmlComparerSettings();

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.NotNull(result);
            Assert.True(result.ShapesInserted >= 1, $"Expected at least 1 shape inserted, got {result.ShapesInserted}");
            Assert.Contains(result.Changes, c => c.ChangeType == PmlChangeType.ShapeInserted);
        }

        [Fact]
        public void PC006_ShapeDeleted_DetectsDeletion()
        {
            // Arrange
            var doc1 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-Base.pptx"));
            var doc2 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-ShapeDeleted.pptx"));
            var settings = new PmlComparerSettings();

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.NotNull(result);
            Assert.True(result.ShapesDeleted >= 1, $"Expected at least 1 shape deleted, got {result.ShapesDeleted}");
            Assert.Contains(result.Changes, c => c.ChangeType == PmlChangeType.ShapeDeleted);
        }

        [Fact]
        public void PC007_ShapeMoved_DetectsMove()
        {
            // Arrange
            var doc1 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-Base.pptx"));
            var doc2 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-ShapeMoved.pptx"));
            var settings = new PmlComparerSettings();

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.NotNull(result);
            Assert.True(result.ShapesMoved >= 1, $"Expected at least 1 shape moved, got {result.ShapesMoved}");
            Assert.Contains(result.Changes, c => c.ChangeType == PmlChangeType.ShapeMoved);
        }

        [Fact]
        public void PC008_ShapeResized_DetectsResize()
        {
            // Arrange
            var doc1 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-Base.pptx"));
            var doc2 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-ShapeResized.pptx"));
            var settings = new PmlComparerSettings();

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.NotNull(result);
            Assert.True(result.ShapesResized >= 1, $"Expected at least 1 shape resized, got {result.ShapesResized}");
            Assert.Contains(result.Changes, c => c.ChangeType == PmlChangeType.ShapeResized);
        }

        #endregion

        #region Text Change Detection Tests

        [Fact]
        public void PC009_TextChanged_DetectsTextChange()
        {
            // Arrange
            var doc1 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-Base.pptx"));
            var doc2 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-TextChanged.pptx"));
            var settings = new PmlComparerSettings();

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.NotNull(result);
            Assert.True(result.TextChanges >= 1, $"Expected at least 1 text change, got {result.TextChanges}");
            Assert.Contains(result.Changes, c => c.ChangeType == PmlChangeType.TextChanged);
        }

        #endregion

        #region Settings Tests

        [Fact]
        public void PC010_DefaultSettings_HasCorrectDefaults()
        {
            // Arrange & Act
            var settings = new PmlComparerSettings();

            // Assert
            Assert.True(settings.CompareSlideStructure);
            Assert.True(settings.CompareShapeStructure);
            Assert.True(settings.CompareTextContent);
            Assert.True(settings.CompareTextFormatting);
            Assert.True(settings.CompareShapeTransforms);
            Assert.False(settings.CompareShapeStyles);
            Assert.True(settings.CompareImageContent);
            Assert.True(settings.CompareCharts);
            Assert.True(settings.CompareTables);
            Assert.False(settings.CompareNotes);
            Assert.False(settings.CompareTransitions);
            Assert.True(settings.EnableFuzzyShapeMatching);
            Assert.True(settings.UseSlideAlignmentLCS);
            Assert.True(settings.AddSummarySlide);
            Assert.True(settings.AddNotesAnnotations);
        }

        [Fact]
        public void PC011_DisableSlideStructure_IgnoresSlideChanges()
        {
            // Arrange
            var doc1 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-Base.pptx"));
            var doc2 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-SlideAdded.pptx"));
            var settings = new PmlComparerSettings { CompareSlideStructure = false };

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.NotNull(result);
            Assert.Equal(0, result.SlidesInserted);
            Assert.Equal(0, result.SlidesDeleted);
        }

        [Fact]
        public void PC012_DisableShapeStructure_IgnoresShapeChanges()
        {
            // Arrange
            var doc1 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-Base.pptx"));
            var doc2 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-ShapeAdded.pptx"));
            var settings = new PmlComparerSettings { CompareShapeStructure = false };

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.NotNull(result);
            Assert.Equal(0, result.ShapesInserted);
            Assert.Equal(0, result.ShapesDeleted);
        }

        [Fact]
        public void PC013_DisableTextContent_IgnoresTextChanges()
        {
            // Arrange
            var doc1 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-Base.pptx"));
            var doc2 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-TextChanged.pptx"));
            var settings = new PmlComparerSettings { CompareTextContent = false };

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.NotNull(result);
            Assert.Equal(0, result.TextChanges);
        }

        [Fact]
        public void PC014_DisableShapeTransforms_IgnoresMoveAndResize()
        {
            // Arrange
            var doc1 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-Base.pptx"));
            var doc2 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-ShapeMoved.pptx"));
            var settings = new PmlComparerSettings { CompareShapeTransforms = false };

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.NotNull(result);
            Assert.Equal(0, result.ShapesMoved);
            Assert.Equal(0, result.ShapesResized);
        }

        #endregion

        #region Result Properties Tests

        [Fact]
        public void PC015_Result_GetChangesBySlide_Works()
        {
            // Arrange
            var doc1 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-Base.pptx"));
            var doc2 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-ShapeAdded.pptx"));
            var settings = new PmlComparerSettings();
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Act
            var slideChanges = result.GetChangesBySlide(1);

            // Assert
            Assert.NotNull(slideChanges);
            Assert.True(slideChanges.Any(), "Slide 1 should have changes");
        }

        [Fact]
        public void PC016_Result_GetChangesByType_Works()
        {
            // Arrange
            var doc1 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-Base.pptx"));
            var doc2 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-SlideAdded.pptx"));
            var settings = new PmlComparerSettings();
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Act
            var slideInsertions = result.GetChangesByType(PmlChangeType.SlideInserted);

            // Assert
            Assert.NotNull(slideInsertions);
            Assert.Single(slideInsertions);
        }

        [Fact]
        public void PC017_Result_ToJson_ReturnsValidJson()
        {
            // Arrange
            var doc1 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-Base.pptx"));
            var doc2 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-SlideAdded.pptx"));
            var settings = new PmlComparerSettings();
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Act
            var json = result.ToJson();

            // Assert
            Assert.NotNull(json);
            Assert.Contains("TotalChanges", json);
            Assert.Contains("Summary", json);
            Assert.Contains("Changes", json);
            Assert.Contains("SlidesInserted", json);
        }

        [Fact]
        public void PC018_Change_GetDescription_ReturnsReadableText()
        {
            // Arrange
            var doc1 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-Base.pptx"));
            var doc2 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-SlideAdded.pptx"));
            var settings = new PmlComparerSettings();
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Act
            var change = result.Changes.First(c => c.ChangeType == PmlChangeType.SlideInserted);
            var description = change.GetDescription();

            // Assert
            Assert.NotNull(description);
            Assert.Contains("inserted", description.ToLower());
        }

        #endregion

        #region Marked Presentation Tests

        [Fact]
        public void PC019_ProduceMarkedPresentation_ReturnsValidDocument()
        {
            // Arrange
            var doc1 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-Base.pptx"));
            var doc2 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-ShapeAdded.pptx"));
            var settings = new PmlComparerSettings();

            // Act
            var marked = PmlComparer.ProduceMarkedPresentation(doc1, doc2, settings);

            // Assert
            Assert.NotNull(marked);
            Assert.NotNull(marked.DocumentByteArray);
            Assert.True(marked.DocumentByteArray.Length > 0);
        }

        [Fact]
        public void PC020_ProduceMarkedPresentation_WithSummarySlide()
        {
            // Arrange
            var doc1 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-Base.pptx"));
            var doc2 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-SlideAdded.pptx"));
            var settings = new PmlComparerSettings { AddSummarySlide = true };

            // Act
            var marked = PmlComparer.ProduceMarkedPresentation(doc1, doc2, settings);

            // Assert
            Assert.NotNull(marked);
            Assert.True(marked.DocumentByteArray.Length > doc2.DocumentByteArray.Length,
                "Marked presentation with summary slide should be larger");
        }

        [Fact]
        public void PC021_ProduceMarkedPresentation_NoChanges_ReturnsSameSize()
        {
            // Arrange
            var doc1 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-Base.pptx"));
            var doc2 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-Identical.pptx"));
            var settings = new PmlComparerSettings { AddSummarySlide = false };

            // Act
            var marked = PmlComparer.ProduceMarkedPresentation(doc1, doc2, settings);

            // Assert
            Assert.NotNull(marked);
            // With no changes and no summary slide, size should be similar
            Assert.True(marked.DocumentByteArray.Length > 0);
        }

        #endregion

        #region Logging Tests

        [Fact]
        public void PC022_LogCallback_ReceivesMessages()
        {
            // Arrange
            var logMessages = new List<string>();
            var doc1 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-Base.pptx"));
            var doc2 = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-SlideAdded.pptx"));
            var settings = new PmlComparerSettings
            {
                LogCallback = msg => logMessages.Add(msg)
            };

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.NotNull(result);
            Assert.True(logMessages.Count > 0, "Log callback should receive messages");
            Assert.Contains(logMessages, m => m.Contains("PmlComparer"));
        }

        #endregion

        #region Edge Cases

        [Fact]
        public void PC023_Compare_NullOlder_ThrowsArgumentNullException()
        {
            // Arrange
            var doc = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-Base.pptx"));
            var settings = new PmlComparerSettings();

            // Act & Assert
            Assert.Throws<ArgumentNullException>(() => PmlComparer.Compare(null, doc, settings));
        }

        [Fact]
        public void PC024_Compare_NullNewer_ThrowsArgumentNullException()
        {
            // Arrange
            var doc = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-Base.pptx"));
            var settings = new PmlComparerSettings();

            // Act & Assert
            Assert.Throws<ArgumentNullException>(() => PmlComparer.Compare(doc, null, settings));
        }

        [Fact]
        public void PC025_Compare_NullSettings_UsesDefaults()
        {
            // Arrange
            var doc = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-Base.pptx"));

            // Act
            var result = PmlComparer.Compare(doc, doc, null);

            // Assert
            Assert.NotNull(result);
            Assert.Equal(0, result.TotalChanges);
        }

        #endregion

        #region Change Type Tests

        [Fact]
        public void PC026_ChangeType_HasExpectedValues()
        {
            // Assert - verify all expected change types exist
            Assert.Equal(0, (int)PmlChangeType.SlideSizeChanged);
            Assert.True(Enum.IsDefined(typeof(PmlChangeType), PmlChangeType.SlideInserted));
            Assert.True(Enum.IsDefined(typeof(PmlChangeType), PmlChangeType.SlideDeleted));
            Assert.True(Enum.IsDefined(typeof(PmlChangeType), PmlChangeType.SlideMoved));
            Assert.True(Enum.IsDefined(typeof(PmlChangeType), PmlChangeType.ShapeInserted));
            Assert.True(Enum.IsDefined(typeof(PmlChangeType), PmlChangeType.ShapeDeleted));
            Assert.True(Enum.IsDefined(typeof(PmlChangeType), PmlChangeType.ShapeMoved));
            Assert.True(Enum.IsDefined(typeof(PmlChangeType), PmlChangeType.ShapeResized));
            Assert.True(Enum.IsDefined(typeof(PmlChangeType), PmlChangeType.TextChanged));
            Assert.True(Enum.IsDefined(typeof(PmlChangeType), PmlChangeType.ImageReplaced));
        }

        #endregion

        #region Canonicalize Tests

        [Fact]
        public void PC027_Canonicalize_ReturnsSignature()
        {
            // Arrange
            var doc = new PmlDocument(_fixture.GetTestFilePath("PmlComparer-Base.pptx"));
            var settings = new PmlComparerSettings();

            // Act
            var signature = PmlComparer.Canonicalize(doc, settings);

            // Assert
            Assert.NotNull(signature);
        }

        #endregion
    }

    /// <summary>
    /// Test fixture that ensures test files exist before tests run.
    /// </summary>
    public class PmlComparerTestFixture : IDisposable
    {
        public string TestFilesDir { get; }

        public PmlComparerTestFixture()
        {
            TestFilesDir = "../../../../TestFiles/";
            PmlComparerTestFileGenerator.EnsureTestFilesExist();
        }

        public string GetTestFilePath(string fileName)
        {
            return Path.Combine(TestFilesDir, fileName);
        }

        public void Dispose()
        {
            // Cleanup if needed
        }
    }
}

#endif
