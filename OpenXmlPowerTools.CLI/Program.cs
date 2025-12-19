using System.CommandLine;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

var rootCommand = new RootCommand("OpenXmlPowerTools CLI - Document comparison and manipulation tools");

// Compare command
var compareCommand = new Command("compare", "Compare two Office documents (docx, pptx, or xlsx) and produce a comparison result");
var doc1Argument = new Argument<FileInfo>("document1", "Path to the original document");
var doc2Argument = new Argument<FileInfo>("document2", "Path to the modified document");
var outputOption = new Option<FileInfo?>(
    aliases: new[] { "-o", "--output" },
    description: "Output path for the comparison result (defaults to comparison-result.[ext])");
var authorOption = new Option<string>(
    aliases: new[] { "-a", "--author" },
    getDefaultValue: () => "Comparer",
    description: "Author name for revision tracking (Word only)");

compareCommand.AddArgument(doc1Argument);
compareCommand.AddArgument(doc2Argument);
compareCommand.AddOption(outputOption);
compareCommand.AddOption(authorOption);

compareCommand.SetHandler(async (doc1, doc2, output, author) =>
{
    if (!doc1.Exists)
    {
        Console.Error.WriteLine($"Error: Document not found: {doc1.FullName}");
        Environment.Exit(1);
    }
    if (!doc2.Exists)
    {
        Console.Error.WriteLine($"Error: Document not found: {doc2.FullName}");
        Environment.Exit(1);
    }

    var ext1 = Path.GetExtension(doc1.Name).ToLowerInvariant();
    var ext2 = Path.GetExtension(doc2.Name).ToLowerInvariant();

    if (ext1 != ext2)
    {
        Console.Error.WriteLine($"Error: Both documents must have the same file type. Got {ext1} and {ext2}");
        Environment.Exit(1);
    }

    try
    {
        switch (ext1)
        {
            case ".docx":
                CompareWord(doc1, doc2, output, author);
                break;
            case ".pptx":
                ComparePowerPoint(doc1, doc2, output);
                break;
            case ".xlsx":
                CompareExcel(doc1, doc2, output);
                break;
            default:
                Console.Error.WriteLine($"Error: Unsupported file type: {ext1}");
                Console.Error.WriteLine("Supported types: .docx, .pptx, .xlsx");
                Environment.Exit(1);
                break;
        }
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"Error during comparison: {ex.Message}");
        if (ex.InnerException != null)
            Console.Error.WriteLine($"  Inner: {ex.InnerException.Message}");
        Environment.Exit(1);
    }
}, doc1Argument, doc2Argument, outputOption, authorOption);

rootCommand.AddCommand(compareCommand);

return await rootCommand.InvokeAsync(args);

void CompareWord(FileInfo doc1, FileInfo doc2, FileInfo? output, string author)
{
    var outputPath = output?.FullName ?? "comparison-result.docx";

    Console.WriteLine($"Comparing Word documents...");
    Console.WriteLine($"  Original: {doc1.FullName}");
    Console.WriteLine($"  Modified: {doc2.FullName}");
    Console.WriteLine($"  Output:   {outputPath}");
    Console.WriteLine();

    var source1 = new WmlDocument(doc1.FullName);
    var source2 = new WmlDocument(doc2.FullName);

    var settings = new WmlComparerSettings
    {
        AuthorForRevisions = author,
        DateTimeForRevisions = DateTime.Now.ToString("o"),
    };

    var result = WmlComparer.Compare(source1, source2, settings);
    result.SaveAs(outputPath);

    Console.WriteLine($"Comparison complete! Output saved to: {outputPath}");

    // Get revision count
    using var resultDoc = WordprocessingDocument.Open(outputPath, false);
    var body = resultDoc.MainDocumentPart?.Document?.Body;
    if (body != null)
    {
        var insCount = body.Descendants<DocumentFormat.OpenXml.Wordprocessing.InsertedRun>().Count();
        var delCount = body.Descendants<DocumentFormat.OpenXml.Wordprocessing.DeletedRun>().Count();
        Console.WriteLine($"  Insertions: {insCount}");
        Console.WriteLine($"  Deletions:  {delCount}");
    }
}

void ComparePowerPoint(FileInfo doc1, FileInfo doc2, FileInfo? output)
{
    var outputPath = output?.FullName ?? "comparison-result.pptx";

    Console.WriteLine($"Comparing PowerPoint presentations...");
    Console.WriteLine($"  Original: {doc1.FullName}");
    Console.WriteLine($"  Modified: {doc2.FullName}");
    Console.WriteLine($"  Output:   {outputPath}");
    Console.WriteLine();

    var source1 = new PmlDocument(doc1.FullName);
    var source2 = new PmlDocument(doc2.FullName);

    var settings = new PmlComparerSettings();

    var comparisonResult = PmlComparer.Compare(source1, source2, settings);

    // Produce marked presentation
    var markedPresentation = PmlComparer.ProduceMarkedPresentation(source1, source2, settings);
    markedPresentation.SaveAs(outputPath);

    Console.WriteLine($"Comparison complete! Output saved to: {outputPath}");
    Console.WriteLine($"  Total changes: {comparisonResult.TotalChanges}");
    Console.WriteLine($"  Slides inserted: {comparisonResult.SlidesInserted}");
    Console.WriteLine($"  Slides deleted: {comparisonResult.SlidesDeleted}");
    Console.WriteLine($"  Shapes inserted: {comparisonResult.ShapesInserted}");
    Console.WriteLine($"  Shapes deleted: {comparisonResult.ShapesDeleted}");
    Console.WriteLine($"  Text changes: {comparisonResult.TextChanges}");
}

void CompareExcel(FileInfo doc1, FileInfo doc2, FileInfo? output)
{
    var outputPath = output?.FullName ?? "comparison-result.xlsx";

    Console.WriteLine($"Comparing Excel workbooks...");
    Console.WriteLine($"  Original: {doc1.FullName}");
    Console.WriteLine($"  Modified: {doc2.FullName}");
    Console.WriteLine($"  Output:   {outputPath}");
    Console.WriteLine();

    var source1 = new SmlDocument(doc1.FullName);
    var source2 = new SmlDocument(doc2.FullName);

    var settings = new SmlComparerSettings();

    var comparisonResult = SmlComparer.Compare(source1, source2, settings);

    // Produce marked workbook
    var markedWorkbook = SmlComparer.ProduceMarkedWorkbook(source1, source2, settings);
    markedWorkbook.SaveAs(outputPath);

    Console.WriteLine($"Comparison complete! Output saved to: {outputPath}");
    Console.WriteLine($"  Total changes: {comparisonResult.TotalChanges}");
    Console.WriteLine($"  Sheets added: {comparisonResult.SheetsAdded}");
    Console.WriteLine($"  Sheets deleted: {comparisonResult.SheetsDeleted}");
    Console.WriteLine($"  Sheets renamed: {comparisonResult.SheetsRenamed}");
    Console.WriteLine($"  Rows inserted: {comparisonResult.RowsInserted}");
    Console.WriteLine($"  Rows deleted: {comparisonResult.RowsDeleted}");
    Console.WriteLine($"  Value changes: {comparisonResult.ValueChanges}");
}
