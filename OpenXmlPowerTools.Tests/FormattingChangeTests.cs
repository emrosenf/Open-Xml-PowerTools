// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OpenXmlPowerTools;
using Xunit;

#if !ELIDE_XUNIT_TESTS

namespace OxPt
{
    /// <summary>
    /// Tests for formatting change detection (rPrChange elements).
    /// These tests verify that WmlComparer correctly detects and tracks
    /// formatting changes like bold, italic, underline, font changes, etc.
    /// </summary>
    public class FormattingChangeTests
    {
        private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        /// <summary>
        /// Count w:rPrChange elements in a document to verify formatting changes are tracked.
        /// </summary>
        private static int CountRPrChanges(WmlDocument wmlDoc)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(wmlDoc.DocumentByteArray, 0, wmlDoc.DocumentByteArray.Length);
                ms.Position = 0;
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, false))
                {
                    var mainPart = wDoc.MainDocumentPart;
                    var doc = XDocument.Load(mainPart.GetStream());
                    return doc.Descendants(W + "rPrChange").Count();
                }
            }
        }

        // ===================================================================================
        // Basic formatting tests - single property changes
        // ===================================================================================

        [Theory]
        [InlineData("FC-0010", "FC/bold_added_before.docx", "FC/bold_added_after.docx", 1, "Bold added to word")]
        [InlineData("FC-0020", "FC/bold_removed_before.docx", "FC/bold_removed_after.docx", 1, "Bold removed from word")]
        [InlineData("FC-0030", "FC/italic_added_before.docx", "FC/italic_added_after.docx", 1, "Italic added to word")]
        [InlineData("FC-0040", "FC/underline_added_before.docx", "FC/underline_added_after.docx", 1, "Underline added to word")]
        [InlineData("FC-0050", "FC/underline_removed_before.docx", "FC/underline_removed_after.docx", 1, "Underline removed from word")]

        // ===================================================================================
        // Multiple formatting changes
        // ===================================================================================

        [InlineData("FC-0100", "FC/bold_to_italic_before.docx", "FC/bold_to_italic_after.docx", 2, "Bold changed to italic")]
        [InlineData("FC-0110", "FC/toggle_runs_before.docx", "FC/toggle_runs_after.docx", 2, "Multiple bold/italic toggles")]
        [InlineData("FC-0120", "FC/multi_changes_before.docx", "FC/multi_changes_after.docx", 2, "Multiple different changes")]
        [InlineData("FC-0130", "FC/adjacent_different_changes_before.docx", "FC/adjacent_different_changes_after.docx", 2, "Different changes on adjacent words")]

        // ===================================================================================
        // Advanced formatting tests - font properties
        // ===================================================================================

        [InlineData("FC-0200", "FC/color_change_before.docx", "FC/color_change_after.docx", 1, "Font color changed")]
        [InlineData("FC-0210", "FC/font_size_change_before.docx", "FC/font_size_change_after.docx", 1, "Font size changed")]
        [InlineData("FC-0220", "FC/font_family_change_before.docx", "FC/font_family_change_after.docx", 1, "Font family changed")]
        [InlineData("FC-0230", "FC/highlight_change_before.docx", "FC/highlight_change_after.docx", 1, "Text highlight added")]

        // ===================================================================================
        // Special formatting tests
        // ===================================================================================

        [InlineData("FC-0300", "FC/caps_change_before.docx", "FC/caps_change_after.docx", 1, "All caps applied")]
        [InlineData("FC-0310", "FC/small_caps_change_before.docx", "FC/small_caps_change_after.docx", 1, "Small caps applied")]
        [InlineData("FC-0320", "FC/strikethrough_change_before.docx", "FC/strikethrough_change_after.docx", 1, "Strikethrough added")]
        [InlineData("FC-0330", "FC/underline_style_change_before.docx", "FC/underline_style_change_after.docx", 1, "Underline style changed")]
        [InlineData("FC-0340", "FC/multi_property_change_before.docx", "FC/multi_property_change_after.docx", 1, "Multiple properties changed together")]

        // ===================================================================================
        // Edge cases
        // ===================================================================================

        [InlineData("FC-0400", "FC/mid_word_before.docx", "FC/mid_word_after.docx", 1, "Mid-word formatting change")]
        [InlineData("FC-0410", "FC/landlord_fee_swap_before.docx", "FC/landlord_fee_swap_after.docx", 2, "Bold moved between words")]

        public void FC001_FormattingChange_CountRPrChanges(string testId, string name1, string name2, int expectedRPrChangeCount, string description)
        {
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/");
            FileInfo source1Docx = new FileInfo(Path.Combine(sourceDir.FullName, name1));
            FileInfo source2Docx = new FileInfo(Path.Combine(sourceDir.FullName, name2));

            if (!source1Docx.Exists)
            {
                // Skip test if fixtures not yet copied
                return;
            }

            var rootTempDir = TestUtil.TempDir;
            var thisTestTempDir = new DirectoryInfo(Path.Combine(rootTempDir.FullName, testId));
            if (thisTestTempDir.Exists)
                Assert.True(false, "Duplicate test id: " + testId);
            else
                thisTestTempDir.Create();

            // Load source documents
            WmlDocument source1Wml = new WmlDocument(source1Docx.FullName);
            WmlDocument source2Wml = new WmlDocument(source2Docx.FullName);

            // Compare documents
            WmlComparerSettings settings = new WmlComparerSettings();
            settings.DebugTempFileDi = thisTestTempDir;
            WmlDocument comparedWml = WmlComparer.Compare(source1Wml, source2Wml, settings);

            // Save output for debugging
            var outputFi = new FileInfo(Path.Combine(thisTestTempDir.FullName, "compared-output.docx"));
            comparedWml.SaveAs(outputFi.FullName);

            // Count rPrChange elements
            int actualCount = CountRPrChanges(comparedWml);

            // Validate the output document
            ValidateDocument(comparedWml);

            // Assert expected rPrChange count
            Assert.True(actualCount == expectedRPrChangeCount,
                $"{description}: Expected {expectedRPrChangeCount} rPrChange elements, but found {actualCount}");
        }

        private static void ValidateDocument(WmlDocument wmlToValidate)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(wmlToValidate.DocumentByteArray, 0, wmlToValidate.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    OpenXmlValidator validator = new OpenXmlValidator();
                    var errors = validator.Validate(wDoc).Where(e => !WcTests.ExpectedErrors.Contains(e.Description));
                    if (errors.Count() != 0)
                    {
                        var errorMessages = string.Join("\n", errors.Select(e => $"  {e.ErrorType}: {e.Description}"));
                        Assert.True(false, $"Document validation failed:\n{errorMessages}");
                    }
                }
            }
        }
    }
}

#endif
