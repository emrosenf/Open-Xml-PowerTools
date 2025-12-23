// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Xml.Linq;
using OpenXmlPowerTools.Packaging;

namespace OpenXmlPowerTools.Wasi;

/// <summary>
/// WASI component entry point for document operations.
/// Uses SharpCompress-based packaging for WASI compatibility (no System.IO.Compression).
/// </summary>
public static class Program
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public static void Main(string[] args)
    {
        Console.WriteLine("OpenXmlPowerTools WASI Test");
        Console.WriteLine("===========================");

        if (args.Length < 1)
        {
            Console.WriteLine("Usage: program <command> [args...]");
            Console.WriteLine("Commands:");
            Console.WriteLine("  info <file.docx>                              - Show document info");
            Console.WriteLine("  extract-text <file.docx>                      - Extract text content");
            Console.WriteLine("  list-parts <file.docx>                        - List all parts");
            Console.WriteLine("  compare <original.docx> <modified.docx> [out] - Compare documents");
            Console.WriteLine("  test                                          - Run built-in test");
            return;
        }

        var command = args[0].ToLowerInvariant();

        try
        {
            switch (command)
            {
                case "info":
                    if (args.Length < 2)
                    {
                        Console.Error.WriteLine("Usage: info <file.docx>");
                        return;
                    }
                    ShowDocumentInfo(args[1]);
                    break;

                case "extract-text":
                    if (args.Length < 2)
                    {
                        Console.Error.WriteLine("Usage: extract-text <file.docx>");
                        return;
                    }
                    ExtractText(args[1]);
                    break;

                case "list-parts":
                    if (args.Length < 2)
                    {
                        Console.Error.WriteLine("Usage: list-parts <file.docx>");
                        return;
                    }
                    ListParts(args[1]);
                    break;

                case "compare":
                    if (args.Length < 3)
                    {
                        Console.Error.WriteLine("Usage: compare <original.docx> <modified.docx> [output.docx]");
                        return;
                    }
                    var outputPath = args.Length > 3 ? args[3] : "comparison-result.docx";
                    CompareDocuments(args[1], args[2], outputPath);
                    break;

                case "test":
                    RunBuiltInTest();
                    break;

                default:
                    Console.Error.WriteLine($"Unknown command: {command}");
                    break;
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
            Console.Error.WriteLine(ex.StackTrace);
        }
    }

    /// <summary>
    /// Show basic document information.
    /// </summary>
    private static void ShowDocumentInfo(string path)
    {
        Console.WriteLine($"Document: {path}");

        var bytes = File.ReadAllBytes(path);
        Console.WriteLine($"File size: {bytes.Length} bytes");

        using var doc = WasiWordDocument.Open(bytes);

        Console.WriteLine("\nParts found:");
        Console.WriteLine($"  - Main Document: {doc.MainDocumentPart.Uri}");

        if (doc.StyleDefinitionsPart != null)
            Console.WriteLine($"  - Styles: {doc.StyleDefinitionsPart.Uri}");

        if (doc.NumberingDefinitionsPart != null)
            Console.WriteLine($"  - Numbering: {doc.NumberingDefinitionsPart.Uri}");

        if (doc.FootnotesPart != null)
            Console.WriteLine($"  - Footnotes: {doc.FootnotesPart.Uri}");

        if (doc.EndnotesPart != null)
            Console.WriteLine($"  - Endnotes: {doc.EndnotesPart.Uri}");

        var headerCount = doc.HeaderParts.Count();
        var footerCount = doc.FooterParts.Count();
        Console.WriteLine($"  - Headers: {headerCount}");
        Console.WriteLine($"  - Footers: {footerCount}");

        // Count paragraphs
        var mainXml = doc.MainDocumentPart.GetXDocument();
        var paragraphs = mainXml.Descendants(W + "p").Count();
        var runs = mainXml.Descendants(W + "r").Count();
        var textElements = mainXml.Descendants(W + "t").Count();

        Console.WriteLine($"\nContent statistics:");
        Console.WriteLine($"  - Paragraphs: {paragraphs}");
        Console.WriteLine($"  - Runs: {runs}");
        Console.WriteLine($"  - Text elements: {textElements}");
    }

    /// <summary>
    /// Extract and display text content.
    /// </summary>
    private static void ExtractText(string path)
    {
        var bytes = File.ReadAllBytes(path);
        using var doc = WasiWordDocument.Open(bytes);

        var mainXml = doc.MainDocumentPart.GetXDocument();
        var textElements = mainXml.Descendants(W + "t");

        Console.WriteLine("Document text:");
        Console.WriteLine("==============");

        foreach (var para in mainXml.Descendants(W + "p"))
        {
            var paraText = string.Concat(para.Descendants(W + "t").Select(t => t.Value));
            if (!string.IsNullOrWhiteSpace(paraText))
            {
                Console.WriteLine(paraText);
            }
        }
    }

    /// <summary>
    /// Compare two documents and produce a result with tracked changes.
    /// </summary>
    private static void CompareDocuments(string originalPath, string modifiedPath, string outputPath)
    {
        Console.WriteLine($"Comparing documents...");
        Console.WriteLine($"  Original: {originalPath}");
        Console.WriteLine($"  Modified: {modifiedPath}");

        var originalBytes = File.ReadAllBytes(originalPath);
        var modifiedBytes = File.ReadAllBytes(modifiedPath);

        var settings = new WasiComparerSettings
        {
            AuthorForRevisions = "WasiComparer",
            DateTimeForRevisions = DateTime.UtcNow
        };

        var result = WasiComparer.Compare(originalBytes, modifiedBytes, settings);

        Console.WriteLine($"\nComparison Results:");
        Console.WriteLine($"  Insertions: {result.Insertions}");
        Console.WriteLine($"  Deletions: {result.Deletions}");
        Console.WriteLine($"  Identical: {result.AreIdentical}");

        File.WriteAllBytes(outputPath, result.ResultDocument);
        Console.WriteLine($"\nOutput saved to: {outputPath}");
        Console.WriteLine($"Output size: {result.ResultDocument.Length} bytes");
    }

    /// <summary>
    /// List all parts in the package.
    /// </summary>
    private static void ListParts(string path)
    {
        var bytes = File.ReadAllBytes(path);
        using var doc = WasiWordDocument.Open(bytes);

        Console.WriteLine("All parts in package:");
        Console.WriteLine("=====================");

        foreach (var part in doc.GetAllParts().OrderBy(p => p.Uri.OriginalString))
        {
            Console.WriteLine($"  {part.Uri} [{part.ContentType}]");
        }
    }

    /// <summary>
    /// Run a built-in test to verify the adapter works.
    /// </summary>
    private static void RunBuiltInTest()
    {
        Console.WriteLine("Running built-in test...");
        Console.WriteLine();

        // Create a minimal DOCX in memory
        Console.WriteLine("1. Creating minimal DOCX in memory...");

        using var ms = new MemoryStream();
        using (var package = SharpCompressPackage.Create(ms))
        {
            // Create main document part
            var docUri = new Uri("/word/document.xml", UriKind.Relative);
            var docPart = package.CreatePart(docUri,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                System.IO.Packaging.CompressionOption.Normal);

            var docXml = new XDocument(
                new XElement(W + "document",
                    new XAttribute(XNamespace.Xmlns + "w", W.NamespaceName),
                    new XElement(W + "body",
                        new XElement(W + "p",
                            new XElement(W + "r",
                                new XElement(W + "t", "Hello from WASI!"))),
                        new XElement(W + "p",
                            new XElement(W + "r",
                                new XElement(W + "t", "This document was created using SharpCompress."))))));

            using (var partStream = docPart.GetStream(FileMode.Create, FileAccess.Write))
            using (var writer = new StreamWriter(partStream))
            {
                docXml.Save(writer);
            }

            // Create package relationship to document
            package.Relationships.Create(
                new Uri("word/document.xml", UriKind.Relative),
                System.IO.Packaging.TargetMode.Internal,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
                "rId1");

            package.Save();
        }

        Console.WriteLine($"   Created document: {ms.Length} bytes");

        // Re-open and read
        Console.WriteLine("2. Re-opening and reading document...");
        ms.Position = 0;
        var docBytes = ms.ToArray();

        using (var doc = WasiWordDocument.Open(docBytes))
        {
            var mainXml = doc.MainDocumentPart.GetXDocument();
            var paragraphs = mainXml.Descendants(W + "p").ToList();

            Console.WriteLine($"   Found {paragraphs.Count} paragraphs:");
            foreach (var para in paragraphs)
            {
                var text = string.Concat(para.Descendants(W + "t").Select(t => t.Value));
                Console.WriteLine($"   - \"{text}\"");
            }
        }

        // Modify and save
        Console.WriteLine("3. Modifying document...");
        using (var doc = WasiWordDocument.Open(docBytes, isEditable: true))
        {
            var mainXml = doc.MainDocumentPart.GetXDocument();
            var body = mainXml.Descendants(W + "body").First();

            // Add a new paragraph
            body.Add(new XElement(W + "p",
                new XElement(W + "r",
                    new XElement(W + "t", "Added by WASI at " + DateTime.UtcNow.ToString("o")))));

            doc.MainDocumentPart.PutXDocument();
            var modifiedBytes = doc.ToByteArray();

            Console.WriteLine($"   Modified document: {modifiedBytes.Length} bytes");

            // Verify modification
            using var verifyDoc = WasiWordDocument.Open(modifiedBytes);
            var verifyXml = verifyDoc.MainDocumentPart.GetXDocument();
            var verifyParas = verifyXml.Descendants(W + "p").Count();
            Console.WriteLine($"   Verified: {verifyParas} paragraphs (was 2, now 3)");
        }

        Console.WriteLine();
        Console.WriteLine("All tests passed!");
    }
}
