// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Text;
using System.Xml.Linq;

namespace OpenXmlPowerTools.Packaging;

/// <summary>
/// XML namespace constants for Office Open XML.
/// </summary>
public static class WasiNamespaces
{
    public static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    public static readonly XNamespace W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
    public static readonly XNamespace W15 = "http://schemas.microsoft.com/office/word/2012/wordml";
    public static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    public static readonly XNamespace WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    public static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    public static readonly XNamespace MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
}

/// <summary>
/// Settings for document comparison.
/// </summary>
public class WasiComparerSettings
{
    /// <summary>
    /// The author name to use for revision marks.
    /// </summary>
    public string AuthorForRevisions { get; set; } = "WasiComparer";

    /// <summary>
    /// The date/time to use for revision marks.
    /// </summary>
    public DateTime DateTimeForRevisions { get; set; } = DateTime.UtcNow;

    /// <summary>
    /// Whether to include formatting changes in the comparison.
    /// </summary>
    public bool CompareFormatting { get; set; } = true;
}

/// <summary>
/// Result of comparing two documents.
/// </summary>
public class WasiCompareResult
{
    /// <summary>
    /// The comparison result document as bytes.
    /// </summary>
    public byte[] ResultDocument { get; set; } = Array.Empty<byte>();

    /// <summary>
    /// Number of insertions found.
    /// </summary>
    public int Insertions { get; set; }

    /// <summary>
    /// Number of deletions found.
    /// </summary>
    public int Deletions { get; set; }

    /// <summary>
    /// Whether the documents are identical.
    /// </summary>
    public bool AreIdentical => Insertions == 0 && Deletions == 0;
}

/// <summary>
/// WASI-compatible document comparer.
/// Provides simplified document comparison that works in WASM/NativeAOT environments.
/// </summary>
public static class WasiComparer
{
    private static readonly XNamespace W = WasiNamespaces.W;

    /// <summary>
    /// Compares two Word documents and produces a result with tracked changes.
    /// </summary>
    public static WasiCompareResult Compare(byte[] source1, byte[] source2, WasiComparerSettings? settings = null)
    {
        settings ??= new WasiComparerSettings();

        using var doc1 = WasiWordDocument.Open(source1);
        using var doc2 = WasiWordDocument.Open(source2);

        return CompareDocuments(doc1, doc2, settings);
    }

    /// <summary>
    /// Compares two Word documents and produces a result with tracked changes.
    /// </summary>
    public static WasiCompareResult CompareDocuments(WasiWordDocument doc1, WasiWordDocument doc2, WasiComparerSettings settings)
    {
        var result = new WasiCompareResult();

        // Get the main document XML from both
        var xml1 = doc1.MainDocumentPart.GetXDocument();
        var xml2 = doc2.MainDocumentPart.GetXDocument();

        // Extract paragraphs from both documents
        var paras1 = ExtractParagraphs(xml1);
        var paras2 = ExtractParagraphs(xml2);

        // Compute the longest common subsequence
        var lcs = ComputeLCS(paras1, paras2);

        // Build the result document with revision marks
        var resultXml = BuildResultDocument(xml1, paras1, paras2, lcs, settings, out int insertions, out int deletions);

        result.Insertions = insertions;
        result.Deletions = deletions;

        // Create result document
        result.ResultDocument = CreateResultDocument(doc1, resultXml);

        return result;
    }

    /// <summary>
    /// Extracts paragraph content for comparison.
    /// </summary>
    private static List<ParagraphInfo> ExtractParagraphs(XDocument doc)
    {
        var paragraphs = new List<ParagraphInfo>();
        var body = doc.Root?.Element(W + "body");
        if (body == null) return paragraphs;

        foreach (var para in body.Elements(W + "p"))
        {
            var info = new ParagraphInfo
            {
                Element = para,
                Text = GetParagraphText(para),
                Hash = ComputeHash(para)
            };
            paragraphs.Add(info);
        }

        return paragraphs;
    }

    /// <summary>
    /// Gets the text content of a paragraph.
    /// </summary>
    private static string GetParagraphText(XElement para)
    {
        var sb = new StringBuilder();
        foreach (var t in para.Descendants(W + "t"))
        {
            sb.Append(t.Value);
        }
        return sb.ToString();
    }

    /// <summary>
    /// Computes a hash for a paragraph element using FNV-1a algorithm.
    /// (SHA256 is not available in WASI NativeAOT)
    /// </summary>
    private static string ComputeHash(XElement element)
    {
        // Create a normalized representation for hashing
        var sb = new StringBuilder();
        AppendNormalizedContent(element, sb);
        var content = sb.ToString();

        // Use FNV-1a hash algorithm (works in WASI)
        ulong hash = 14695981039346656037UL; // FNV offset basis
        const ulong fnvPrime = 1099511628211UL;

        foreach (char c in content)
        {
            hash ^= c;
            hash *= fnvPrime;
        }

        return hash.ToString("X16");
    }

    private static void AppendNormalizedContent(XElement element, StringBuilder sb)
    {
        foreach (var node in element.Nodes())
        {
            if (node is XText text)
            {
                sb.Append(text.Value);
            }
            else if (node is XElement child)
            {
                if (child.Name == W + "t")
                {
                    sb.Append(child.Value);
                }
                else if (child.Name != W + "rPr" && child.Name != W + "pPr")
                {
                    // Skip formatting properties but recurse into other elements
                    AppendNormalizedContent(child, sb);
                }
            }
        }
    }

    /// <summary>
    /// Computes the Longest Common Subsequence of two paragraph lists.
    /// </summary>
    private static List<(int Index1, int Index2)> ComputeLCS(List<ParagraphInfo> list1, List<ParagraphInfo> list2)
    {
        int m = list1.Count;
        int n = list2.Count;

        // Build the LCS table
        var dp = new int[m + 1, n + 1];

        for (int i = 1; i <= m; i++)
        {
            for (int j = 1; j <= n; j++)
            {
                if (list1[i - 1].Hash == list2[j - 1].Hash)
                {
                    dp[i, j] = dp[i - 1, j - 1] + 1;
                }
                else
                {
                    dp[i, j] = Math.Max(dp[i - 1, j], dp[i, j - 1]);
                }
            }
        }

        // Backtrack to find the LCS
        var lcs = new List<(int, int)>();
        int x = m, y = n;
        while (x > 0 && y > 0)
        {
            if (list1[x - 1].Hash == list2[y - 1].Hash)
            {
                lcs.Add((x - 1, y - 1));
                x--;
                y--;
            }
            else if (dp[x - 1, y] > dp[x, y - 1])
            {
                x--;
            }
            else
            {
                y--;
            }
        }

        lcs.Reverse();
        return lcs;
    }

    /// <summary>
    /// Builds the result document with revision marks.
    /// </summary>
    private static XDocument BuildResultDocument(
        XDocument originalDoc,
        List<ParagraphInfo> paras1,
        List<ParagraphInfo> paras2,
        List<(int Index1, int Index2)> lcs,
        WasiComparerSettings settings,
        out int insertions,
        out int deletions)
    {
        insertions = 0;
        deletions = 0;

        var result = new XDocument(originalDoc);
        var body = result.Root?.Element(W + "body");
        if (body == null) return result;

        // Save sectPr (section properties) - it must be last in the body
        var sectPr = body.Element(W + "sectPr");
        sectPr?.Remove();

        // Remove existing paragraphs and other content
        body.Elements(W + "p").Remove();
        body.Elements(W + "tbl").Remove(); // tables too

        // Track positions in both documents
        int pos1 = 0;
        int pos2 = 0;
        int lcsIndex = 0;
        int revisionId = 0; // Counter for unique revision IDs

        var author = settings.AuthorForRevisions;
        var date = settings.DateTimeForRevisions.ToString("yyyy-MM-ddTHH:mm:ssZ");

        while (pos1 < paras1.Count || pos2 < paras2.Count)
        {
            if (lcsIndex < lcs.Count && pos1 == lcs[lcsIndex].Index1 && pos2 == lcs[lcsIndex].Index2)
            {
                // This paragraph is in both - add it unchanged
                body.Add(new XElement(paras1[pos1].Element));
                pos1++;
                pos2++;
                lcsIndex++;
            }
            else if (lcsIndex < lcs.Count && pos1 < lcs[lcsIndex].Index1 && (pos2 >= paras2.Count || pos2 >= lcs[lcsIndex].Index2))
            {
                // Paragraph in doc1 but not in doc2 - mark as deleted
                var deleted = CreateDeletedParagraph(paras1[pos1].Element, author, date, ref revisionId);
                body.Add(deleted);
                pos1++;
                deletions++;
            }
            else if (lcsIndex < lcs.Count && pos2 < lcs[lcsIndex].Index2 && (pos1 >= paras1.Count || pos1 >= lcs[lcsIndex].Index1))
            {
                // Paragraph in doc2 but not in doc1 - mark as inserted
                var inserted = CreateInsertedParagraph(paras2[pos2].Element, author, date, ref revisionId);
                body.Add(inserted);
                pos2++;
                insertions++;
            }
            else if (pos1 < paras1.Count && (lcsIndex >= lcs.Count || pos1 < lcs[lcsIndex].Index1))
            {
                // Remaining in doc1 - mark as deleted
                var deleted = CreateDeletedParagraph(paras1[pos1].Element, author, date, ref revisionId);
                body.Add(deleted);
                pos1++;
                deletions++;
            }
            else if (pos2 < paras2.Count && (lcsIndex >= lcs.Count || pos2 < lcs[lcsIndex].Index2))
            {
                // Remaining in doc2 - mark as inserted
                var inserted = CreateInsertedParagraph(paras2[pos2].Element, author, date, ref revisionId);
                body.Add(inserted);
                pos2++;
                insertions++;
            }
            else
            {
                // Safety break to avoid infinite loop
                break;
            }
        }

        // Re-add sectPr at the end of body (required by Word)
        if (sectPr != null)
        {
            body.Add(sectPr);
        }

        return result;
    }

    /// <summary>
    /// Creates a paragraph marked as deleted.
    /// </summary>
    private static XElement CreateDeletedParagraph(XElement para, string author, string date, ref int revisionId)
    {
        var result = new XElement(para);

        // Wrap all runs in w:del elements
        foreach (var run in result.Elements(W + "r").ToList())
        {
            var del = new XElement(W + "del",
                new XAttribute(W + "id", revisionId++.ToString()),
                new XAttribute(W + "author", author),
                new XAttribute(W + "date", date),
                run);
            run.ReplaceWith(del);
        }

        // Also wrap the paragraph properties reference if exists
        var pPr = result.Element(W + "pPr");
        if (pPr != null)
        {
            // Add revision property for deleted paragraph mark
            var rPr = pPr.Element(W + "rPr") ?? new XElement(W + "rPr");
            if (rPr.Parent == null) pPr.Add(rPr);
            rPr.Add(new XElement(W + "del",
                new XAttribute(W + "id", revisionId++.ToString()),
                new XAttribute(W + "author", author),
                new XAttribute(W + "date", date)));
        }

        return result;
    }

    /// <summary>
    /// Creates a paragraph marked as inserted.
    /// </summary>
    private static XElement CreateInsertedParagraph(XElement para, string author, string date, ref int revisionId)
    {
        var result = new XElement(para);

        // Wrap all runs in w:ins elements
        foreach (var run in result.Elements(W + "r").ToList())
        {
            var ins = new XElement(W + "ins",
                new XAttribute(W + "id", revisionId++.ToString()),
                new XAttribute(W + "author", author),
                new XAttribute(W + "date", date),
                run);
            run.ReplaceWith(ins);
        }

        // Also mark the paragraph itself as inserted
        var pPr = result.Element(W + "pPr");
        if (pPr != null)
        {
            var rPr = pPr.Element(W + "rPr") ?? new XElement(W + "rPr");
            if (rPr.Parent == null) pPr.Add(rPr);
            rPr.Add(new XElement(W + "ins",
                new XAttribute(W + "id", revisionId++.ToString()),
                new XAttribute(W + "author", author),
                new XAttribute(W + "date", date)));
        }

        return result;
    }

    /// <summary>
    /// Creates the result document by copying the source and updating the main document part.
    /// </summary>
    private static byte[] CreateResultDocument(WasiWordDocument sourceDoc, XDocument resultXml)
    {
        // Get source document bytes
        var sourceBytes = sourceDoc.ToByteArray();

        // Open source package to read all parts
        using var sourceStream = new MemoryStream(sourceBytes);
        using var sourcePackage = SharpCompressPackage.Open(sourceStream, FileAccess.Read);

        // Create new output package
        using var outputStream = new MemoryStream();
        using var outputPackage = SharpCompressPackage.Create(outputStream);

        // Find the main document part URI
        var mainDocRel = sourcePackage.Relationships
            .FirstOrDefault(r => r.RelationshipType == WasiRelationshipTypes.OfficeDocument);

        Uri? mainDocUri = null;
        if (mainDocRel != null)
        {
            mainDocUri = mainDocRel.TargetUri;
            if (!mainDocUri.OriginalString.StartsWith("/"))
                mainDocUri = new Uri("/" + mainDocUri.OriginalString, UriKind.Relative);
        }

        // Copy all parts from source to output
        foreach (var sourcePart in sourcePackage.GetParts())
        {
            var newPart = outputPackage.CreatePart(
                sourcePart.Uri,
                sourcePart.ContentType,
                System.IO.Packaging.CompressionOption.Normal);

            using var targetStream = newPart.GetStream(FileMode.Create, FileAccess.Write);

            // For the main document part, write the modified XML
            if (mainDocUri != null && sourcePart.Uri.OriginalString == mainDocUri.OriginalString)
            {
                resultXml.Save(targetStream);
            }
            else
            {
                // Copy the original content
                using var partSourceStream = sourcePart.GetStream(FileMode.Open, FileAccess.Read);
                partSourceStream.CopyTo(targetStream);
            }

            // Copy part relationships
            foreach (var rel in sourcePart.Relationships)
            {
                newPart.Relationships.Create(rel.TargetUri, rel.TargetMode, rel.RelationshipType, rel.Id);
            }
        }

        // Copy package-level relationships
        foreach (var rel in sourcePackage.Relationships)
        {
            outputPackage.Relationships.Create(rel.TargetUri, rel.TargetMode, rel.RelationshipType, rel.Id);
        }

        outputPackage.Save();

        return outputStream.ToArray();
    }

    /// <summary>
    /// Internal class to hold paragraph information.
    /// </summary>
    private class ParagraphInfo
    {
        public XElement Element { get; set; } = null!;
        public string Text { get; set; } = string.Empty;
        public string Hash { get; set; } = string.Empty;
    }
}
