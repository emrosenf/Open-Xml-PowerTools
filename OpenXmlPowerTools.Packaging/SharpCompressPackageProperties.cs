// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools.Packaging;

/// <summary>
/// A pure managed implementation of IPackageProperties.
/// </summary>
public class SharpCompressPackageProperties : IPackageProperties
{
    private static readonly XNamespace CpNs = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    private static readonly XNamespace DcNs = "http://purl.org/dc/elements/1.1/";
    private static readonly XNamespace DcTermsNs = "http://purl.org/dc/terms/";

    private readonly SharpCompressPackage _package;

    public SharpCompressPackageProperties(SharpCompressPackage package)
    {
        _package = package;
    }

    public string? Title { get; set; }
    public string? Subject { get; set; }
    public string? Creator { get; set; }
    public string? Keywords { get; set; }
    public string? Description { get; set; }
    public string? LastModifiedBy { get; set; }
    public string? Revision { get; set; }
    public DateTime? LastPrinted { get; set; }
    public DateTime? Created { get; set; }
    public DateTime? Modified { get; set; }
    public string? Category { get; set; }
    public string? Identifier { get; set; }
    public string? ContentType { get; set; }
    public string? Language { get; set; }
    public string? Version { get; set; }
    public string? ContentStatus { get; set; }

    internal bool HasValues =>
        !string.IsNullOrEmpty(Title) ||
        !string.IsNullOrEmpty(Subject) ||
        !string.IsNullOrEmpty(Creator) ||
        !string.IsNullOrEmpty(Keywords) ||
        !string.IsNullOrEmpty(Description) ||
        !string.IsNullOrEmpty(LastModifiedBy) ||
        Created.HasValue ||
        Modified.HasValue;

    internal void LoadFromXDocument(XDocument doc)
    {
        var root = doc.Root;
        if (root == null) return;

        Title = GetElementValue(root, DcNs + "title");
        Subject = GetElementValue(root, DcNs + "subject");
        Creator = GetElementValue(root, DcNs + "creator");
        Keywords = GetElementValue(root, CpNs + "keywords");
        Description = GetElementValue(root, DcNs + "description");
        LastModifiedBy = GetElementValue(root, CpNs + "lastModifiedBy");
        Revision = GetElementValue(root, CpNs + "revision");
        Category = GetElementValue(root, CpNs + "category");
        ContentStatus = GetElementValue(root, CpNs + "contentStatus");
        Identifier = GetElementValue(root, DcNs + "identifier");
        Language = GetElementValue(root, DcNs + "language");
        Version = GetElementValue(root, CpNs + "version");

        var createdStr = GetElementValue(root, DcTermsNs + "created");
        if (!string.IsNullOrEmpty(createdStr) && DateTime.TryParse(createdStr, out var created))
        {
            Created = created;
        }

        var modifiedStr = GetElementValue(root, DcTermsNs + "modified");
        if (!string.IsNullOrEmpty(modifiedStr) && DateTime.TryParse(modifiedStr, out var modified))
        {
            Modified = modified;
        }

        var lastPrintedStr = GetElementValue(root, CpNs + "lastPrinted");
        if (!string.IsNullOrEmpty(lastPrintedStr) && DateTime.TryParse(lastPrintedStr, out var lastPrinted))
        {
            LastPrinted = lastPrinted;
        }
    }

    internal XDocument ToXDocument()
    {
        var doc = new XDocument(
            new XElement(CpNs + "coreProperties",
                new XAttribute(XNamespace.Xmlns + "cp", CpNs),
                new XAttribute(XNamespace.Xmlns + "dc", DcNs),
                new XAttribute(XNamespace.Xmlns + "dcterms", DcTermsNs)
            )
        );

        var root = doc.Root!;

        AddElementIfNotEmpty(root, DcNs + "title", Title);
        AddElementIfNotEmpty(root, DcNs + "subject", Subject);
        AddElementIfNotEmpty(root, DcNs + "creator", Creator);
        AddElementIfNotEmpty(root, CpNs + "keywords", Keywords);
        AddElementIfNotEmpty(root, DcNs + "description", Description);
        AddElementIfNotEmpty(root, CpNs + "lastModifiedBy", LastModifiedBy);
        AddElementIfNotEmpty(root, CpNs + "revision", Revision);
        AddElementIfNotEmpty(root, CpNs + "category", Category);
        AddElementIfNotEmpty(root, CpNs + "contentStatus", ContentStatus);
        AddElementIfNotEmpty(root, DcNs + "identifier", Identifier);
        AddElementIfNotEmpty(root, DcNs + "language", Language);
        AddElementIfNotEmpty(root, CpNs + "version", Version);

        if (Created.HasValue)
        {
            var xsiNs = XNamespace.Get("http://www.w3.org/2001/XMLSchema-instance");
            root.Add(new XElement(DcTermsNs + "created",
                new XAttribute(xsiNs + "type", "dcterms:W3CDTF"),
                Created.Value.ToString("yyyy-MM-ddTHH:mm:ssZ")));
        }

        if (Modified.HasValue)
        {
            var xsiNs = XNamespace.Get("http://www.w3.org/2001/XMLSchema-instance");
            root.Add(new XElement(DcTermsNs + "modified",
                new XAttribute(xsiNs + "type", "dcterms:W3CDTF"),
                Modified.Value.ToString("yyyy-MM-ddTHH:mm:ssZ")));
        }

        if (LastPrinted.HasValue)
        {
            root.Add(new XElement(CpNs + "lastPrinted",
                LastPrinted.Value.ToString("yyyy-MM-ddTHH:mm:ssZ")));
        }

        return doc;
    }

    private static string? GetElementValue(XElement parent, XName name)
    {
        return parent.Element(name)?.Value;
    }

    private static void AddElementIfNotEmpty(XElement parent, XName name, string? value)
    {
        if (!string.IsNullOrEmpty(value))
        {
            parent.Add(new XElement(name, value));
        }
    }
}
