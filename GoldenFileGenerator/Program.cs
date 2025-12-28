// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

// Golden File Generator for TypeScript Port TDD
// Generates reference outputs from the C# WmlComparer, SmlComparer, and PmlComparer
// to validate TypeScript implementations against.

using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

class Program
{
    // Paths work both locally (from GoldenFileGenerator dir) and in Docker (from /app)
    static readonly string TestFilesDir = Directory.Exists("TestFiles") ? "TestFiles" : "../TestFiles";
    static readonly string OutputDir = Directory.Exists("redline-js") ? "redline-js/tests/golden" : "../redline-js/tests/golden";
    static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    static async Task Main(string[] args)
    {
        Console.WriteLine("Golden File Generator for TypeScript Port");
        Console.WriteLine("==========================================\n");

        // Ensure output directory exists
        Directory.CreateDirectory(OutputDir);
        Directory.CreateDirectory(Path.Combine(OutputDir, "wml"));
        Directory.CreateDirectory(Path.Combine(OutputDir, "sml"));
        Directory.CreateDirectory(Path.Combine(OutputDir, "pml"));

        var manifest = new TestManifest
        {
            GeneratedAt = DateTime.UtcNow.ToString("o"),
            WmlTests = new List<WmlTestCase>(),
            SmlTests = new List<SmlTestCase>(),
            PmlTests = new List<PmlTestCase>()
        };

        // Generate WML golden files
        Console.WriteLine("Generating WML (Word) golden files...");
        await GenerateWmlGoldenFiles(manifest);

        // Generate SML golden files
        Console.WriteLine("\nGenerating SML (Excel) golden files...");
        await GenerateSmlGoldenFiles(manifest);

        // Generate PML golden files
        Console.WriteLine("\nGenerating PML (PowerPoint) golden files...");
        await GeneratePmlGoldenFiles(manifest);

        // Write manifest
        var manifestPath = Path.Combine(OutputDir, "manifest.json");
        var options = new JsonSerializerOptions { WriteIndented = true };
        await File.WriteAllTextAsync(manifestPath, JsonSerializer.Serialize(manifest, options));

        Console.WriteLine($"\n==========================================");
        Console.WriteLine($"Generated {manifest.WmlTests.Count} WML test cases");
        Console.WriteLine($"Generated {manifest.SmlTests.Count} SML test cases");
        Console.WriteLine($"Generated {manifest.PmlTests.Count} PML test cases");
        Console.WriteLine($"Manifest written to: {manifestPath}");
    }

    static async Task GenerateWmlGoldenFiles(TestManifest manifest)
    {
        // WmlComparer test cases from WmlComparerTests.cs WC003_Compare
        var testCases = new List<(string TestId, string File1, string File2, int ExpectedRevisions)>
        {
            ("WC-1000", "CA/CA001-Plain.docx", "CA/CA001-Plain-Mod.docx", 1),
            ("WC-1010", "WC/WC001-Digits.docx", "WC/WC001-Digits-Mod.docx", 4),
            ("WC-1020", "WC/WC001-Digits.docx", "WC/WC001-Digits-Deleted-Paragraph.docx", 1),
            ("WC-1030", "WC/WC001-Digits-Deleted-Paragraph.docx", "WC/WC001-Digits.docx", 1),
            ("WC-1040", "WC/WC002-Unmodified.docx", "WC/WC002-DiffInMiddle.docx", 2),
            ("WC-1050", "WC/WC002-Unmodified.docx", "WC/WC002-DiffAtBeginning.docx", 2),
            ("WC-1060", "WC/WC002-Unmodified.docx", "WC/WC002-DeleteAtBeginning.docx", 1),
            ("WC-1070", "WC/WC002-Unmodified.docx", "WC/WC002-InsertAtBeginning.docx", 1),
            ("WC-1080", "WC/WC002-Unmodified.docx", "WC/WC002-InsertAtEnd.docx", 1),
            ("WC-1090", "WC/WC002-Unmodified.docx", "WC/WC002-DeleteAtEnd.docx", 1),
            ("WC-1100", "WC/WC002-Unmodified.docx", "WC/WC002-DeleteInMiddle.docx", 1),
            ("WC-1110", "WC/WC002-Unmodified.docx", "WC/WC002-InsertInMiddle.docx", 1),
            ("WC-1120", "WC/WC002-DeleteInMiddle.docx", "WC/WC002-Unmodified.docx", 1),
            ("WC-1140", "WC/WC006-Table.docx", "WC/WC006-Table-Delete-Row.docx", 1),
            ("WC-1150", "WC/WC006-Table-Delete-Row.docx", "WC/WC006-Table.docx", 1),
            ("WC-1160", "WC/WC006-Table.docx", "WC/WC006-Table-Delete-Contests-of-Row.docx", 2),
            ("WC-1170", "WC/WC007-Unmodified.docx", "WC/WC007-Longest-At-End.docx", 2),
            ("WC-1180", "WC/WC007-Unmodified.docx", "WC/WC007-Deleted-at-Beginning-of-Para.docx", 1),
            ("WC-1190", "WC/WC007-Unmodified.docx", "WC/WC007-Moved-into-Table.docx", 2),
            ("WC-1200", "WC/WC009-Table-Unmodified.docx", "WC/WC009-Table-Cell-1-1-Mod.docx", 1),
            ("WC-1210", "WC/WC010-Para-Before-Table-Unmodified.docx", "WC/WC010-Para-Before-Table-Mod.docx", 3),
            ("WC-1220", "WC/WC011-Before.docx", "WC/WC011-After.docx", 2),
            ("WC-1230", "WC/WC012-Math-Before.docx", "WC/WC012-Math-After.docx", 2),
            ("WC-1240", "WC/WC013-Image-Before.docx", "WC/WC013-Image-After.docx", 2),
            ("WC-1250", "WC/WC013-Image-Before.docx", "WC/WC013-Image-After2.docx", 2),
            ("WC-1260", "WC/WC013-Image-Before2.docx", "WC/WC013-Image-After2.docx", 2),
            ("WC-1270", "WC/WC014-SmartArt-Before.docx", "WC/WC014-SmartArt-After.docx", 2),
            ("WC-1280", "WC/WC014-SmartArt-With-Image-Before.docx", "WC/WC014-SmartArt-With-Image-After.docx", 2),
            ("WC-1310", "WC/WC014-SmartArt-With-Image-Before.docx", "WC/WC014-SmartArt-With-Image-Deleted-After.docx", 3),
            ("WC-1320", "WC/WC014-SmartArt-With-Image-Before.docx", "WC/WC014-SmartArt-With-Image-Deleted-After2.docx", 1),
            ("WC-1330", "WC/WC015-Three-Paragraphs.docx", "WC/WC015-Three-Paragraphs-After.docx", 3),
            ("WC-1340", "WC/WC016-Para-Image-Para.docx", "WC/WC016-Para-Image-Para-w-Deleted-Image.docx", 1),
            ("WC-1350", "WC/WC017-Image.docx", "WC/WC017-Image-After.docx", 3),
            ("WC-1360", "WC/WC018-Field-Simple-Before.docx", "WC/WC018-Field-Simple-After-1.docx", 2),
            ("WC-1370", "WC/WC018-Field-Simple-Before.docx", "WC/WC018-Field-Simple-After-2.docx", 3),
            ("WC-1380", "WC/WC019-Hyperlink-Before.docx", "WC/WC019-Hyperlink-After-1.docx", 3),
            ("WC-1390", "WC/WC019-Hyperlink-Before.docx", "WC/WC019-Hyperlink-After-2.docx", 5),
            ("WC-1400", "WC/WC020-FootNote-Before.docx", "WC/WC020-FootNote-After-1.docx", 3),
            ("WC-1410", "WC/WC020-FootNote-Before.docx", "WC/WC020-FootNote-After-2.docx", 5),
            ("WC-1420", "WC/WC021-Math-Before-1.docx", "WC/WC021-Math-After-1.docx", 9),
            ("WC-1430", "WC/WC021-Math-Before-2.docx", "WC/WC021-Math-After-2.docx", 6),
            ("WC-1440", "WC/WC022-Image-Math-Para-Before.docx", "WC/WC022-Image-Math-Para-After.docx", 10),
            ("WC-1450", "WC/WC023-Table-4-Row-Image-Before.docx", "WC/WC023-Table-4-Row-Image-After-Delete-1-Row.docx", 7),
            ("WC-1460", "WC/WC024-Table-Before.docx", "WC/WC024-Table-After.docx", 1),
            ("WC-1470", "WC/WC024-Table-Before.docx", "WC/WC024-Table-After2.docx", 7),
            ("WC-1480", "WC/WC025-Simple-Table-Before.docx", "WC/WC025-Simple-Table-After.docx", 4),
            ("WC-1500", "WC/WC026-Long-Table-Before.docx", "WC/WC026-Long-Table-After-1.docx", 2),
            ("WC-1510", "WC/WC027-Twenty-Paras-Before.docx", "WC/WC027-Twenty-Paras-After-1.docx", 2),
            ("WC-1520", "WC/WC027-Twenty-Paras-After-1.docx", "WC/WC027-Twenty-Paras-Before.docx", 2),
            ("WC-1530", "WC/WC027-Twenty-Paras-Before.docx", "WC/WC027-Twenty-Paras-After-2.docx", 4),
            ("WC-1540", "WC/WC030-Image-Math-Before.docx", "WC/WC030-Image-Math-After.docx", 2),
            ("WC-1550", "WC/WC031-Two-Maths-Before.docx", "WC/WC031-Two-Maths-After.docx", 4),
            ("WC-1560", "WC/WC032-Para-with-Para-Props.docx", "WC/WC032-Para-with-Para-Props-After.docx", 3),
            ("WC-1570", "WC/WC033-Merged-Cells-Before.docx", "WC/WC033-Merged-Cells-After1.docx", 2),
            ("WC-1580", "WC/WC033-Merged-Cells-Before.docx", "WC/WC033-Merged-Cells-After2.docx", 4),
            ("WC-1600", "WC/WC034-Footnotes-Before.docx", "WC/WC034-Footnotes-After1.docx", 1),
            ("WC-1610", "WC/WC034-Footnotes-Before.docx", "WC/WC034-Footnotes-After2.docx", 4),
            ("WC-1620", "WC/WC034-Footnotes-Before.docx", "WC/WC034-Footnotes-After3.docx", 3),
            ("WC-1630", "WC/WC034-Footnotes-After3.docx", "WC/WC034-Footnotes-Before.docx", 3),
            ("WC-1640", "WC/WC035-Footnote-Before.docx", "WC/WC035-Footnote-After.docx", 2),
            ("WC-1650", "WC/WC035-Footnote-After.docx", "WC/WC035-Footnote-Before.docx", 2),
            ("WC-1660", "WC/WC036-Footnote-With-Table-Before.docx", "WC/WC036-Footnote-With-Table-After.docx", 5),
            ("WC-1670", "WC/WC036-Footnote-With-Table-After.docx", "WC/WC036-Footnote-With-Table-Before.docx", 5),
            ("WC-1680", "WC/WC034-Endnotes-Before.docx", "WC/WC034-Endnotes-After1.docx", 1),
            ("WC-1700", "WC/WC034-Endnotes-Before.docx", "WC/WC034-Endnotes-After2.docx", 4),
            ("WC-1710", "WC/WC034-Endnotes-Before.docx", "WC/WC034-Endnotes-After3.docx", 7),
            ("WC-1720", "WC/WC034-Endnotes-After3.docx", "WC/WC034-Endnotes-Before.docx", 7),
            ("WC-1730", "WC/WC035-Endnote-Before.docx", "WC/WC035-Endnote-After.docx", 2),
            ("WC-1740", "WC/WC035-Endnote-After.docx", "WC/WC035-Endnote-Before.docx", 2),
            ("WC-1750", "WC/WC036-Endnote-With-Table-Before.docx", "WC/WC036-Endnote-With-Table-After.docx", 6),
            ("WC-1760", "WC/WC036-Endnote-With-Table-After.docx", "WC/WC036-Endnote-With-Table-Before.docx", 6),
            ("WC-1770", "WC/WC037-Textbox-Before.docx", "WC/WC037-Textbox-After1.docx", 2),
            ("WC-1780", "WC/WC038-Document-With-BR-Before.docx", "WC/WC038-Document-With-BR-After.docx", 2),
            ("WC-1800", "RC/RC001-Before.docx", "RC/RC001-After1.docx", 2),
            ("WC-1810", "RC/RC002-Image.docx", "RC/RC002-Image-After1.docx", 1),
            ("WC-1820", "WC/WC039-Break-In-Row.docx", "WC/WC039-Break-In-Row-After1.docx", 1),
            ("WC-1830", "WC/WC041-Table-5.docx", "WC/WC041-Table-5-Mod.docx", 2),
            ("WC-1840", "WC/WC042-Table-5.docx", "WC/WC042-Table-5-Mod.docx", 2),
            ("WC-1850", "WC/WC043-Nested-Table.docx", "WC/WC043-Nested-Table-Mod.docx", 2),
            ("WC-1860", "WC/WC044-Text-Box.docx", "WC/WC044-Text-Box-Mod.docx", 2),
            ("WC-1870", "WC/WC045-Text-Box.docx", "WC/WC045-Text-Box-Mod.docx", 2),
            ("WC-1880", "WC/WC046-Two-Text-Box.docx", "WC/WC046-Two-Text-Box-Mod.docx", 2),
            ("WC-1890", "WC/WC047-Two-Text-Box.docx", "WC/WC047-Two-Text-Box-Mod.docx", 2),
            ("WC-1900", "WC/WC048-Text-Box-in-Cell.docx", "WC/WC048-Text-Box-in-Cell-Mod.docx", 6),
            ("WC-1910", "WC/WC049-Text-Box-in-Cell.docx", "WC/WC049-Text-Box-in-Cell-Mod.docx", 5),
            ("WC-1920", "WC/WC050-Table-in-Text-Box.docx", "WC/WC050-Table-in-Text-Box-Mod.docx", 8),
            ("WC-1930", "WC/WC051-Table-in-Text-Box.docx", "WC/WC051-Table-in-Text-Box-Mod.docx", 9),
            ("WC-1940", "WC/WC052-SmartArt-Same.docx", "WC/WC052-SmartArt-Same-Mod.docx", 2),
            ("WC-1950", "WC/WC053-Text-in-Cell.docx", "WC/WC053-Text-in-Cell-Mod.docx", 2),
            ("WC-1960", "WC/WC054-Text-in-Cell.docx", "WC/WC054-Text-in-Cell-Mod.docx", 0),
            ("WC-1970", "WC/WC055-French.docx", "WC/WC055-French-Mod.docx", 0),
            ("WC-1980", "WC/WC056-French.docx", "WC/WC056-French-Mod.docx", 0),
            ("WC-2000", "WC/WC058-Table-Merged-Cell.docx", "WC/WC058-Table-Merged-Cell-Mod.docx", 6),
            ("WC-2010", "WC/WC059-Footnote.docx", "WC/WC059-Footnote-Mod.docx", 5),
            ("WC-2020", "WC/WC060-Endnote.docx", "WC/WC060-Endnote-Mod.docx", 3),
            ("WC-2030", "WC/WC061-Style-Added.docx", "WC/WC061-Style-Added-Mod.docx", 1),
            ("WC-2040", "WC/WC062-New-Char-Style-Added.docx", "WC/WC062-New-Char-Style-Added-Mod.docx", 3),
            ("WC-2050", "WC/WC063-Footnote.docx", "WC/WC063-Footnote-Mod.docx", 1),
            ("WC-2060", "WC/WC063-Footnote-Mod.docx", "WC/WC063-Footnote.docx", 1),
            ("WC-2070", "WC/WC064-Footnote.docx", "WC/WC064-Footnote-Mod.docx", 0),
            ("WC-2080", "WC/WC065-Textbox.docx", "WC/WC065-Textbox-Mod.docx", 2),
            ("WC-2090", "WC/WC066-Textbox-Before-Ins.docx", "WC/WC066-Textbox-Before-Ins-Mod.docx", 1),
            ("WC-2092", "WC/WC066-Textbox-Before-Ins-Mod.docx", "WC/WC066-Textbox-Before-Ins.docx", 1),
            ("WC-2100", "WC/WC067-Textbox-Image.docx", "WC/WC067-Textbox-Image-Mod.docx", 2),
        };

        var wmlOutputDir = Path.Combine(OutputDir, "wml");

        foreach (var (testId, file1, file2, expectedRevisions) in testCases)
        {
            try
            {
                var source1Path = Path.Combine(TestFilesDir, file1);
                var source2Path = Path.Combine(TestFilesDir, file2);

                if (!File.Exists(source1Path) || !File.Exists(source2Path))
                {
                    Console.WriteLine($"  SKIP {testId}: Missing fixture files");
                    continue;
                }

                var source1Wml = new WmlDocument(source1Path);
                var source2Wml = new WmlDocument(source2Path);
                var settings = new WmlComparerSettings();

                // Run comparison
                var comparedWml = WmlComparer.Compare(source1Wml, source2Wml, settings);

                // Count revisions
                var revisions = WmlComparer.GetRevisions(comparedWml, settings);
                var actualRevisions = revisions.Count();

                // Extract revision details
                var revisionDetails = revisions.Select(r => new RevisionDetail
                {
                    Type = r.RevisionType.ToString(),
                    Text = r.Text,
                    Author = r.Author
                }).ToList();

                // Perform sanity checks
                bool sanityCheck1Pass = true;
                bool sanityCheck2Pass = true;
                string? sanityCheck1Error = null;
                string? sanityCheck2Error = null;

                try
                {
                    var afterRejecting = RevisionProcessor.RejectRevisions(comparedWml);
                    var sanityCheck1Result = WmlComparer.Compare(source1Wml, afterRejecting, settings);
                    var sanityCheck1Revisions = WmlComparer.GetRevisions(sanityCheck1Result, settings);
                    if (sanityCheck1Revisions.Any())
                    {
                        sanityCheck1Pass = false;
                        sanityCheck1Error = $"Found {sanityCheck1Revisions.Count()} revisions after reject";
                    }
                }
                catch (Exception ex)
                {
                    sanityCheck1Pass = false;
                    sanityCheck1Error = ex.Message;
                }

                try
                {
                    var afterAccepting = RevisionProcessor.AcceptRevisions(comparedWml);
                    var sanityCheck2Result = WmlComparer.Compare(source2Wml, afterAccepting, settings);
                    var sanityCheck2Revisions = WmlComparer.GetRevisions(sanityCheck2Result, settings);
                    if (sanityCheck2Revisions.Any())
                    {
                        sanityCheck2Pass = false;
                        sanityCheck2Error = $"Found {sanityCheck2Revisions.Count()} revisions after accept";
                    }
                }
                catch (Exception ex)
                {
                    sanityCheck2Pass = false;
                    sanityCheck2Error = ex.Message;
                }

                // Save output document
                var outputPath = Path.Combine(wmlOutputDir, $"{testId}.docx");
                comparedWml.SaveAs(outputPath);

                // Extract and save document.xml for easier diffing
                var documentXml = ExtractDocumentXml(comparedWml);
                var xmlOutputPath = Path.Combine(wmlOutputDir, $"{testId}.document.xml");
                await File.WriteAllTextAsync(xmlOutputPath, documentXml);

                // Add to manifest
                manifest.WmlTests.Add(new WmlTestCase
                {
                    TestId = testId,
                    Source1 = file1,
                    Source2 = file2,
                    ExpectedRevisions = expectedRevisions,
                    ActualRevisions = actualRevisions,
                    Revisions = revisionDetails,
                    SanityCheck1Pass = sanityCheck1Pass,
                    SanityCheck1Error = sanityCheck1Error,
                    SanityCheck2Pass = sanityCheck2Pass,
                    SanityCheck2Error = sanityCheck2Error,
                    OutputFile = $"wml/{testId}.docx",
                    DocumentXmlFile = $"wml/{testId}.document.xml"
                });

                var status = actualRevisions == expectedRevisions ? "OK" : "MISMATCH";
                var sanityStatus = sanityCheck1Pass && sanityCheck2Pass ? "" : " [SANITY FAIL]";
                Console.WriteLine($"  {status} {testId}: {actualRevisions} revisions (expected {expectedRevisions}){sanityStatus}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  ERROR {testId}: {ex.Message}");
                manifest.WmlTests.Add(new WmlTestCase
                {
                    TestId = testId,
                    Source1 = file1,
                    Source2 = file2,
                    ExpectedRevisions = expectedRevisions,
                    Error = ex.Message
                });
            }
        }
    }

    static async Task GenerateSmlGoldenFiles(TestManifest manifest)
    {
        // SmlComparer generates workbooks programmatically in tests
        // For golden files, we'll create a few representative test cases
        Console.WriteLine("  SML golden file generation requires programmatic workbook creation");
        Console.WriteLine("  These will be generated by the TypeScript test setup");

        // Add placeholder entries for the manifest
        manifest.SmlTests.Add(new SmlTestCase
        {
            TestId = "SC-INFO",
            Description = "SML tests use programmatically generated workbooks - see TypeScript test setup"
        });
    }

    static async Task GeneratePmlGoldenFiles(TestManifest manifest)
    {
        var testCases = new List<(string TestId, string File1, string File2, string Description)>
        {
            ("PC-001", "PmlComparer-Base.pptx", "PmlComparer-Identical.pptx", "Identical presentations"),
            ("PC-003", "PmlComparer-Base.pptx", "PmlComparer-SlideAdded.pptx", "Slide added"),
            ("PC-004", "PmlComparer-Base.pptx", "PmlComparer-SlideDeleted.pptx", "Slide deleted"),
            ("PC-005", "PmlComparer-Base.pptx", "PmlComparer-ShapeAdded.pptx", "Shape added"),
            ("PC-006", "PmlComparer-Base.pptx", "PmlComparer-ShapeDeleted.pptx", "Shape deleted"),
            ("PC-007", "PmlComparer-Base.pptx", "PmlComparer-ShapeMoved.pptx", "Shape moved"),
            ("PC-008", "PmlComparer-Base.pptx", "PmlComparer-ShapeResized.pptx", "Shape resized"),
            ("PC-009", "PmlComparer-Base.pptx", "PmlComparer-TextChanged.pptx", "Text changed"),
        };

        var pmlOutputDir = Path.Combine(OutputDir, "pml");

        foreach (var (testId, file1, file2, description) in testCases)
        {
            try
            {
                var source1Path = Path.Combine(TestFilesDir, file1);
                var source2Path = Path.Combine(TestFilesDir, file2);

                if (!File.Exists(source1Path) || !File.Exists(source2Path))
                {
                    Console.WriteLine($"  SKIP {testId}: Missing fixture files ({file1} or {file2})");
                    continue;
                }

                var source1Pml = new PmlDocument(source1Path);
                var source2Pml = new PmlDocument(source2Path);
                var settings = new PmlComparerSettings();

                // Run comparison
                var result = PmlComparer.Compare(source1Pml, source2Pml, settings);

                // Save result JSON
                var jsonOutputPath = Path.Combine(pmlOutputDir, $"{testId}.json");
                await File.WriteAllTextAsync(jsonOutputPath, result.ToJson());

                // Add to manifest
                manifest.PmlTests.Add(new PmlTestCase
                {
                    TestId = testId,
                    Source1 = file1,
                    Source2 = file2,
                    Description = description,
                    TotalChanges = result.TotalChanges,
                    SlidesInserted = result.SlidesInserted,
                    SlidesDeleted = result.SlidesDeleted,
                    ShapesInserted = result.ShapesInserted,
                    ShapesDeleted = result.ShapesDeleted,
                    ShapesMoved = result.ShapesMoved,
                    ShapesResized = result.ShapesResized,
                    TextChanges = result.TextChanges,
                    OutputFile = $"pml/{testId}.json"
                });

                Console.WriteLine($"  OK {testId}: {result.TotalChanges} changes ({description})");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  ERROR {testId}: {ex.Message}");
                manifest.PmlTests.Add(new PmlTestCase
                {
                    TestId = testId,
                    Source1 = file1,
                    Source2 = file2,
                    Description = description,
                    Error = ex.Message
                });
            }
        }
    }

    static string ExtractDocumentXml(WmlDocument wmlDoc)
    {
        using var ms = new MemoryStream(wmlDoc.DocumentByteArray);
        using var wDoc = WordprocessingDocument.Open(ms, false);
        var mainPart = wDoc.MainDocumentPart;
        if (mainPart == null) return "";

        var doc = XDocument.Load(mainPart.GetStream());
        return doc.ToString();
    }
}

// Manifest classes
class TestManifest
{
    public string GeneratedAt { get; set; } = "";
    public List<WmlTestCase> WmlTests { get; set; } = new();
    public List<SmlTestCase> SmlTests { get; set; } = new();
    public List<PmlTestCase> PmlTests { get; set; } = new();
}

class WmlTestCase
{
    public string TestId { get; set; } = "";
    public string Source1 { get; set; } = "";
    public string Source2 { get; set; } = "";
    public int ExpectedRevisions { get; set; }
    public int ActualRevisions { get; set; }
    public List<RevisionDetail> Revisions { get; set; } = new();
    public bool SanityCheck1Pass { get; set; }
    public string? SanityCheck1Error { get; set; }
    public bool SanityCheck2Pass { get; set; }
    public string? SanityCheck2Error { get; set; }
    public string? OutputFile { get; set; }
    public string? DocumentXmlFile { get; set; }
    public string? Error { get; set; }
}

class RevisionDetail
{
    public string Type { get; set; } = "";
    public string? Text { get; set; }
    public string? Author { get; set; }
}

class SmlTestCase
{
    public string TestId { get; set; } = "";
    public string? Description { get; set; }
    public string? Source1 { get; set; }
    public string? Source2 { get; set; }
    public int TotalChanges { get; set; }
    public string? OutputFile { get; set; }
    public string? Error { get; set; }
}

class PmlTestCase
{
    public string TestId { get; set; } = "";
    public string Source1 { get; set; } = "";
    public string Source2 { get; set; } = "";
    public string? Description { get; set; }
    public int TotalChanges { get; set; }
    public int SlidesInserted { get; set; }
    public int SlidesDeleted { get; set; }
    public int ShapesInserted { get; set; }
    public int ShapesDeleted { get; set; }
    public int ShapesMoved { get; set; }
    public int ShapesResized { get; set; }
    public int TextChanges { get; set; }
    public string? OutputFile { get; set; }
    public string? Error { get; set; }
}
