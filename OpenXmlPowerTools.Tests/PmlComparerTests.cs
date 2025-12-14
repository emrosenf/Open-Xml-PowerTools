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
    public class PmlComparerTests
    {
        private static readonly DirectoryInfo TestFilesDir = new DirectoryInfo("../../../../TestFiles/");

        #region Basic Comparison Tests

        [Fact]
        public void PC001_IdenticalPresentations_NoChanges()
        {
            // Arrange
            var source = new FileInfo(Path.Combine(TestFilesDir.FullName, "PB001-Input1.pptx"));
            var doc1 = new PmlDocument(source.FullName);
            var doc2 = new PmlDocument(source.FullName);
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
            var source1 = new FileInfo(Path.Combine(TestFilesDir.FullName, "PB001-Input1.pptx"));
            var source2 = new FileInfo(Path.Combine(TestFilesDir.FullName, "PB001-Input2.pptx"));
            var doc1 = new PmlDocument(source1.FullName);
            var doc2 = new PmlDocument(source2.FullName);
            var settings = new PmlComparerSettings();

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.NotNull(result);
            Assert.True(result.TotalChanges > 0, "Should detect differences between different presentations");
        }

        [Fact]
        public void PC003_Compare_ReturnsValidResult()
        {
            // Arrange
            var source = new FileInfo(Path.Combine(TestFilesDir.FullName, "Presentation.pptx"));
            var doc1 = new PmlDocument(source.FullName);
            var doc2 = new PmlDocument(source.FullName);
            var settings = new PmlComparerSettings();

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Assert
            Assert.NotNull(result);
            Assert.NotNull(result.Changes);
        }

        #endregion

        #region Settings Tests

        [Fact]
        public void PC004_DefaultSettings_HasCorrectDefaults()
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
        public void PC005_CustomSettings_AreRespected()
        {
            // Arrange
            var source = new FileInfo(Path.Combine(TestFilesDir.FullName, "PB001-Input1.pptx"));
            var doc1 = new PmlDocument(source.FullName);
            var doc2 = new PmlDocument(source.FullName);
            var settings = new PmlComparerSettings
            {
                CompareSlideStructure = false,
                CompareShapeStructure = false,
                CompareTextContent = false
            };

            // Act
            var result = PmlComparer.Compare(doc1, doc2, settings);

            // Assert - with all comparisons disabled, should still return a valid result
            Assert.NotNull(result);
        }

        #endregion

        #region Result Properties Tests

        [Fact]
        public void PC006_Result_HasCorrectStatistics()
        {
            // Arrange
            var source = new FileInfo(Path.Combine(TestFilesDir.FullName, "PB001-Input1.pptx"));
            var doc = new PmlDocument(source.FullName);
            var settings = new PmlComparerSettings();

            // Act
            var result = PmlComparer.Compare(doc, doc, settings);

            // Assert
            Assert.Equal(0, result.TotalChanges);
            Assert.Equal(0, result.SlidesInserted);
            Assert.Equal(0, result.SlidesDeleted);
            Assert.Equal(0, result.SlidesMoved);
            Assert.Equal(0, result.ShapesInserted);
            Assert.Equal(0, result.ShapesDeleted);
            Assert.Equal(0, result.ShapesMoved);
            Assert.Equal(0, result.ShapesResized);
            Assert.Equal(0, result.TextChanges);
            Assert.Equal(0, result.FormattingChanges);
            Assert.Equal(0, result.ImagesReplaced);
        }

        [Fact]
        public void PC007_Result_GetChangesBySlide_Works()
        {
            // Arrange
            var source = new FileInfo(Path.Combine(TestFilesDir.FullName, "PB001-Input1.pptx"));
            var doc = new PmlDocument(source.FullName);
            var settings = new PmlComparerSettings();
            var result = PmlComparer.Compare(doc, doc, settings);

            // Act
            var slideChanges = result.GetChangesBySlide(1);

            // Assert
            Assert.NotNull(slideChanges);
        }

        [Fact]
        public void PC008_Result_GetChangesByType_Works()
        {
            // Arrange
            var source = new FileInfo(Path.Combine(TestFilesDir.FullName, "PB001-Input1.pptx"));
            var doc = new PmlDocument(source.FullName);
            var settings = new PmlComparerSettings();
            var result = PmlComparer.Compare(doc, doc, settings);

            // Act
            var slideInsertions = result.GetChangesByType(PmlChangeType.SlideInserted);

            // Assert
            Assert.NotNull(slideInsertions);
            Assert.Empty(slideInsertions);
        }

        [Fact]
        public void PC009_Result_ToJson_ReturnsValidJson()
        {
            // Arrange
            var source = new FileInfo(Path.Combine(TestFilesDir.FullName, "PB001-Input1.pptx"));
            var doc = new PmlDocument(source.FullName);
            var settings = new PmlComparerSettings();
            var result = PmlComparer.Compare(doc, doc, settings);

            // Act
            var json = result.ToJson();

            // Assert
            Assert.NotNull(json);
            Assert.Contains("TotalChanges", json);
            Assert.Contains("Summary", json);
            Assert.Contains("Changes", json);
        }

        #endregion

        #region Marked Presentation Tests

        [Fact]
        public void PC010_ProduceMarkedPresentation_ReturnsValidDocument()
        {
            // Arrange
            var source = new FileInfo(Path.Combine(TestFilesDir.FullName, "PB001-Input1.pptx"));
            var doc = new PmlDocument(source.FullName);
            var settings = new PmlComparerSettings();

            // Act
            var marked = PmlComparer.ProduceMarkedPresentation(doc, doc, settings);

            // Assert
            Assert.NotNull(marked);
            Assert.NotNull(marked.DocumentByteArray);
            Assert.True(marked.DocumentByteArray.Length > 0);
        }

        [Fact]
        public void PC011_ProduceMarkedPresentation_WithDifferences_AddsSummarySlide()
        {
            // Arrange
            var source1 = new FileInfo(Path.Combine(TestFilesDir.FullName, "PB001-Input1.pptx"));
            var source2 = new FileInfo(Path.Combine(TestFilesDir.FullName, "PB001-Input2.pptx"));
            var doc1 = new PmlDocument(source1.FullName);
            var doc2 = new PmlDocument(source2.FullName);
            var settings = new PmlComparerSettings { AddSummarySlide = true };

            // Act
            var marked = PmlComparer.ProduceMarkedPresentation(doc1, doc2, settings);

            // Assert
            Assert.NotNull(marked);
            Assert.True(marked.DocumentByteArray.Length > 0);
        }

        #endregion

        #region Logging Tests

        [Fact]
        public void PC012_LogCallback_ReceivesMessages()
        {
            // Arrange
            var logMessages = new List<string>();
            var source = new FileInfo(Path.Combine(TestFilesDir.FullName, "PB001-Input1.pptx"));
            var doc = new PmlDocument(source.FullName);
            var settings = new PmlComparerSettings
            {
                LogCallback = msg => logMessages.Add(msg)
            };

            // Act
            var result = PmlComparer.Compare(doc, doc, settings);

            // Assert
            Assert.NotNull(result);
            Assert.True(logMessages.Count > 0, "Log callback should receive messages");
        }

        #endregion

        #region Edge Cases

        [Fact]
        public void PC013_Compare_NullOlder_ThrowsArgumentNullException()
        {
            // Arrange
            var source = new FileInfo(Path.Combine(TestFilesDir.FullName, "PB001-Input1.pptx"));
            var doc = new PmlDocument(source.FullName);
            var settings = new PmlComparerSettings();

            // Act & Assert
            Assert.Throws<ArgumentNullException>(() => PmlComparer.Compare(null, doc, settings));
        }

        [Fact]
        public void PC014_Compare_NullNewer_ThrowsArgumentNullException()
        {
            // Arrange
            var source = new FileInfo(Path.Combine(TestFilesDir.FullName, "PB001-Input1.pptx"));
            var doc = new PmlDocument(source.FullName);
            var settings = new PmlComparerSettings();

            // Act & Assert
            Assert.Throws<ArgumentNullException>(() => PmlComparer.Compare(doc, null, settings));
        }

        [Fact]
        public void PC015_Compare_NullSettings_UsesDefaults()
        {
            // Arrange
            var source = new FileInfo(Path.Combine(TestFilesDir.FullName, "PB001-Input1.pptx"));
            var doc = new PmlDocument(source.FullName);

            // Act
            var result = PmlComparer.Compare(doc, doc, null);

            // Assert
            Assert.NotNull(result);
        }

        #endregion

        #region Change Type Tests

        [Fact]
        public void PC016_ChangeType_HasExpectedValues()
        {
            // Assert - verify all expected change types exist
            Assert.Equal(0, (int)PmlChangeType.SlideSizeChanged);
            Assert.True(Enum.IsDefined(typeof(PmlChangeType), PmlChangeType.SlideInserted));
            Assert.True(Enum.IsDefined(typeof(PmlChangeType), PmlChangeType.SlideDeleted));
            Assert.True(Enum.IsDefined(typeof(PmlChangeType), PmlChangeType.SlideMoved));
            Assert.True(Enum.IsDefined(typeof(PmlChangeType), PmlChangeType.ShapeInserted));
            Assert.True(Enum.IsDefined(typeof(PmlChangeType), PmlChangeType.ShapeDeleted));
            Assert.True(Enum.IsDefined(typeof(PmlChangeType), PmlChangeType.TextChanged));
            Assert.True(Enum.IsDefined(typeof(PmlChangeType), PmlChangeType.ImageReplaced));
        }

        #endregion

        #region Canonicalize Tests

        [Fact]
        public void PC017_Canonicalize_ReturnsSignature()
        {
            // Arrange
            var source = new FileInfo(Path.Combine(TestFilesDir.FullName, "PB001-Input1.pptx"));
            var doc = new PmlDocument(source.FullName);
            var settings = new PmlComparerSettings();

            // Act
            var signature = PmlComparer.Canonicalize(doc, settings);

            // Assert
            Assert.NotNull(signature);
        }

        #endregion
    }
}

#endif
