// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Xml.Linq;
using Xunit;

namespace OpenXmlPowerTools.Packaging.Tests;

public class WasiDocumentTests
{
    private static readonly string TestFilesDir = Path.Combine(
        AppDomain.CurrentDomain.BaseDirectory,
        "..", "..", "..", "..", "TestFiles", "WC");

    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Fact]
    public void CanOpenWordDocument()
    {
        var docxPath = Path.Combine(TestFilesDir, "WC001-Digits.docx");
        var docxBytes = File.ReadAllBytes(docxPath);

        using var doc = WasiWordDocument.Open(docxBytes);

        Assert.NotNull(doc.MainDocumentPart);
    }

    [Fact]
    public void CanReadMainDocumentXml()
    {
        var docxPath = Path.Combine(TestFilesDir, "WC001-Digits.docx");
        var docxBytes = File.ReadAllBytes(docxPath);

        using var doc = WasiWordDocument.Open(docxBytes);
        var xdoc = doc.MainDocumentPart.GetXDocument();

        Assert.NotNull(xdoc);
        Assert.NotNull(xdoc.Root);
        Assert.Equal("document", xdoc.Root.Name.LocalName);
    }

    [Fact]
    public void CanAccessStylesPart()
    {
        var docxPath = Path.Combine(TestFilesDir, "WC001-Digits.docx");
        var docxBytes = File.ReadAllBytes(docxPath);

        using var doc = WasiWordDocument.Open(docxBytes);
        var stylesPart = doc.StyleDefinitionsPart;

        Assert.NotNull(stylesPart);

        var xdoc = stylesPart.GetXDocument();
        Assert.NotNull(xdoc.Root);
        Assert.Equal("styles", xdoc.Root.Name.LocalName);
    }

    [Fact]
    public void CanModifyAndSaveDocument()
    {
        var docxPath = Path.Combine(TestFilesDir, "WC001-Digits.docx");
        var docxBytes = File.ReadAllBytes(docxPath);

        byte[] modifiedBytes;

        // Open, modify, save
        using (var doc = WasiWordDocument.Open(docxBytes, isEditable: true))
        {
            var xdoc = doc.MainDocumentPart.GetXDocument();

            // Add a custom attribute to the root
            xdoc.Root!.SetAttributeValue("customAttr", "testValue");

            doc.MainDocumentPart.PutXDocument();
            modifiedBytes = doc.ToByteArray();
        }

        // Re-open and verify modification
        using (var doc = WasiWordDocument.Open(modifiedBytes))
        {
            var xdoc = doc.MainDocumentPart.GetXDocument();
            var customAttr = xdoc.Root!.Attribute("customAttr");

            Assert.NotNull(customAttr);
            Assert.Equal("testValue", customAttr.Value);
        }
    }

    [Fact]
    public void CanGetAllParts()
    {
        var docxPath = Path.Combine(TestFilesDir, "WC001-Digits.docx");
        var docxBytes = File.ReadAllBytes(docxPath);

        using var doc = WasiWordDocument.Open(docxBytes);
        var parts = doc.GetAllParts().ToList();

        Assert.NotEmpty(parts);

        // Should have main document part
        Assert.Contains(parts, p => p.Uri.OriginalString.Contains("document.xml"));

        // Should have styles
        Assert.Contains(parts, p => p.Uri.OriginalString.Contains("styles.xml"));
    }

    [Fact]
    public void CanGetPartByUri()
    {
        var docxPath = Path.Combine(TestFilesDir, "WC001-Digits.docx");
        var docxBytes = File.ReadAllBytes(docxPath);

        using var doc = WasiWordDocument.Open(docxBytes);

        var stylesPart = doc.GetPart("/word/styles.xml");
        Assert.NotNull(stylesPart);

        var xdoc = stylesPart.GetXDocument();
        Assert.Equal("styles", xdoc.Root!.Name.LocalName);
    }

    [Fact]
    public void CanExtractTextContent()
    {
        var docxPath = Path.Combine(TestFilesDir, "WC001-Digits.docx");
        var docxBytes = File.ReadAllBytes(docxPath);

        using var doc = WasiWordDocument.Open(docxBytes);
        var xdoc = doc.MainDocumentPart.GetXDocument();

        // Get all text elements
        var textElements = xdoc.Descendants(W + "t").ToList();

        Assert.NotEmpty(textElements);
    }

    [Fact]
    public void RoundTripPreservesContent()
    {
        var docxPath = Path.Combine(TestFilesDir, "WC001-Digits.docx");
        var originalBytes = File.ReadAllBytes(docxPath);

        // Round-trip the document
        byte[] roundTrippedBytes;
        using (var doc = WasiWordDocument.Open(originalBytes, isEditable: true))
        {
            roundTrippedBytes = doc.ToByteArray();
        }

        // Compare content
        using var original = WasiWordDocument.Open(originalBytes);
        using var roundTripped = WasiWordDocument.Open(roundTrippedBytes);

        var originalXml = original.MainDocumentPart.GetXDocument().ToString();
        var roundTrippedXml = roundTripped.MainDocumentPart.GetXDocument().ToString();

        Assert.Equal(originalXml, roundTrippedXml);
    }

    [Fact]
    public void CanAccessHeadersAndFooters()
    {
        // Use a document that has headers/footers
        var docxPath = Path.Combine(TestFilesDir, "WC001-Digits.docx");
        var docxBytes = File.ReadAllBytes(docxPath);

        using var doc = WasiWordDocument.Open(docxBytes);

        // These may or may not exist depending on the test document
        var headers = doc.HeaderParts.ToList();
        var footers = doc.FooterParts.ToList();

        // Just verify we can enumerate them without errors
        Assert.NotNull(headers);
        Assert.NotNull(footers);
    }
}
