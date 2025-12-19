using System.CommandLine;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

var rootCommand = new RootCommand("OpenXmlPowerTools CLI - Document comparison and manipulation tools");

// Compare command
var compareCommand = new Command("compare", "Compare two Word documents and produce a redline");
var doc1Argument = new Argument<FileInfo>("document1", "Path to the original document");
var doc2Argument = new Argument<FileInfo>("document2", "Path to the modified document");
var outputOption = new Option<FileInfo?>(
    aliases: new[] { "-o", "--output" },
    description: "Output path for the comparison result (defaults to comparison-result.docx)");
var authorOption = new Option<string>(
    aliases: new[] { "-a", "--author" },
    getDefaultValue: () => "WmlComparer",
    description: "Author name for revision tracking");

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

    var outputPath = output?.FullName ?? "comparison-result.docx";

    Console.WriteLine($"Comparing documents...");
    Console.WriteLine($"  Original: {doc1.FullName}");
    Console.WriteLine($"  Modified: {doc2.FullName}");
    Console.WriteLine($"  Output:   {outputPath}");
    Console.WriteLine();

    try
    {
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
