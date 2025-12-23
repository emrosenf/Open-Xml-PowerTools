// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools.Packaging;

/// <summary>
/// Relationship type constants for Office Open XML documents.
/// </summary>
public static class WasiRelationshipTypes
{
    public const string OfficeDocument = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    public const string Styles = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    public const string Numbering = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
    public const string FontTable = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable";
    public const string Theme = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
    public const string Footnotes = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
    public const string Endnotes = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
    public const string Header = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
    public const string Footer = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
    public const string Comments = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
    public const string Image = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    public const string Chart = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
    public const string Settings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
    public const string WebSettings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings";
    public const string Hyperlink = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
}

/// <summary>
/// A part within a WASI-compatible Office document.
/// Provides XDocument-based access to XML parts.
/// Implements IDocumentPart for compatibility with comparison algorithms.
/// </summary>
public class WasiPart : IDocumentPart
{
    private readonly SharpCompressPackage _package;
    private readonly IPackagePart _part;
    private XDocument? _xDocument;

    internal WasiPart(SharpCompressPackage package, IPackagePart part)
    {
        _package = package;
        _part = part;
    }

    public Uri Uri => _part.Uri;
    public string ContentType => _part.ContentType;

    /// <summary>
    /// Gets the XML document for this part.
    /// </summary>
    public XDocument GetXDocument()
    {
        if (_xDocument != null) return _xDocument;

        using var stream = _part.GetStream(FileMode.Open, FileAccess.Read);
        if (stream.Length == 0)
        {
            _xDocument = new XDocument();
            _xDocument.Declaration = new XDeclaration("1.0", "UTF-8", "yes");
        }
        else
        {
            using var reader = XmlReader.Create(stream);
            _xDocument = XDocument.Load(reader);
        }

        return _xDocument;
    }

    /// <summary>
    /// Saves the current XDocument back to the part.
    /// </summary>
    public void PutXDocument()
    {
        if (_xDocument == null) return;

        using var stream = _part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = XmlWriter.Create(stream);
        _xDocument.Save(writer);
    }

    /// <summary>
    /// Sets a new XDocument for this part.
    /// </summary>
    public void PutXDocument(XDocument document)
    {
        _xDocument = document;
        PutXDocument();
    }

    /// <summary>
    /// Gets the root element of the XML document.
    /// </summary>
    public XElement? RootElement => GetXDocument().Root;

    /// <summary>
    /// Gets related parts by relationship type.
    /// </summary>
    public IEnumerable<WasiPart> GetRelatedParts(string relationshipType)
    {
        foreach (var rel in _part.Relationships.Where(r => r.RelationshipType == relationshipType))
        {
            var targetUri = ResolveUri(_part.Uri, rel.TargetUri);
            if (_package.PartExists(targetUri))
            {
                var part = _package.GetPart(targetUri);
                yield return new WasiPart(_package, part);
            }
        }
    }

    /// <summary>
    /// Gets related parts by relationship type (IDocumentPart interface).
    /// </summary>
    IEnumerable<IDocumentPart> IDocumentPart.GetRelatedParts(string relationshipType)
    {
        return GetRelatedParts(relationshipType);
    }

    /// <summary>
    /// Gets a single related part by relationship type.
    /// </summary>
    public WasiPart? GetRelatedPart(string relationshipType)
    {
        return GetRelatedParts(relationshipType).FirstOrDefault();
    }

    /// <summary>
    /// Gets a related part by relationship ID.
    /// </summary>
    public IDocumentPart? GetPartById(string relationshipId)
    {
        var rel = _part.Relationships.FirstOrDefault(r => r.Id == relationshipId);
        if (rel == null) return null;

        var targetUri = ResolveUri(_part.Uri, rel.TargetUri);
        if (!_package.PartExists(targetUri)) return null;

        var part = _package.GetPart(targetUri);
        return new WasiPart(_package, part);
    }

    /// <summary>
    /// Gets the raw bytes of this part.
    /// </summary>
    public byte[] GetBytes()
    {
        using var stream = _part.GetStream(FileMode.Open, FileAccess.Read);
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        return ms.ToArray();
    }

    private static Uri ResolveUri(Uri sourceUri, Uri targetUri)
    {
        if (targetUri.IsAbsoluteUri)
            return targetUri;

        var sourcePath = sourceUri.OriginalString;
        var sourceDir = sourcePath.Substring(0, sourcePath.LastIndexOf('/') + 1);
        var targetPath = targetUri.OriginalString;

        // Handle relative paths
        while (targetPath.StartsWith("../"))
        {
            targetPath = targetPath.Substring(3);
            if (sourceDir.Length > 1)
            {
                sourceDir = sourceDir.TrimEnd('/');
                var lastSlash = sourceDir.LastIndexOf('/');
                sourceDir = lastSlash >= 0 ? sourceDir.Substring(0, lastSlash + 1) : "/";
            }
        }

        var resolved = sourceDir + targetPath;
        if (!resolved.StartsWith("/"))
            resolved = "/" + resolved;

        return new Uri(resolved, UriKind.Relative);
    }
}

/// <summary>
/// A WASI-compatible Word document wrapper.
/// Uses SharpCompress for ZIP operations instead of System.IO.Compression.
/// Implements IWordDocument for compatibility with comparison algorithms.
/// </summary>
public class WasiWordDocument : IWordDocument
{
    private readonly SharpCompressPackage _package;
    private readonly bool _isEditable;
    private WasiPart? _mainDocumentPart;
    private bool _disposed;

    private WasiWordDocument(SharpCompressPackage package, bool isEditable)
    {
        _package = package;
        _isEditable = isEditable;
    }

    /// <summary>
    /// Opens a Word document from a byte array.
    /// </summary>
    public static WasiWordDocument Open(byte[] data, bool isEditable = false)
    {
        var package = SharpCompressPackage.Open(data, isEditable ? FileAccess.ReadWrite : FileAccess.Read);
        return new WasiWordDocument(package, isEditable);
    }

    /// <summary>
    /// Opens a Word document from a stream.
    /// </summary>
    public static WasiWordDocument Open(Stream stream, bool isEditable = false)
    {
        var package = SharpCompressPackage.Open(stream, isEditable ? FileAccess.ReadWrite : FileAccess.Read);
        return new WasiWordDocument(package, isEditable);
    }

    /// <summary>
    /// Gets the main document part.
    /// </summary>
    public WasiPart MainDocumentPart
    {
        get
        {
            if (_mainDocumentPart != null) return _mainDocumentPart;

            // Find the main document part via relationships
            var rel = _package.Relationships
                .FirstOrDefault(r => r.RelationshipType == WasiRelationshipTypes.OfficeDocument);

            if (rel == null)
                throw new InvalidOperationException("No main document part found");

            var uri = rel.TargetUri;
            if (!uri.OriginalString.StartsWith("/"))
                uri = new Uri("/" + uri.OriginalString, UriKind.Relative);

            var part = _package.GetPart(uri);
            _mainDocumentPart = new WasiPart(_package, part);
            return _mainDocumentPart;
        }
    }

    // Explicit interface implementation for IWordDocument
    IDocumentPart IWordDocument.MainDocumentPart => MainDocumentPart;

    /// <summary>
    /// Gets the styles part if it exists.
    /// </summary>
    public WasiPart? StyleDefinitionsPart =>
        MainDocumentPart.GetRelatedPart(WasiRelationshipTypes.Styles);

    IDocumentPart? IWordDocument.StyleDefinitionsPart => StyleDefinitionsPart;

    /// <summary>
    /// Gets the numbering definitions part if it exists.
    /// </summary>
    public WasiPart? NumberingDefinitionsPart =>
        MainDocumentPart.GetRelatedPart(WasiRelationshipTypes.Numbering);

    IDocumentPart? IWordDocument.NumberingDefinitionsPart => NumberingDefinitionsPart;

    /// <summary>
    /// Gets the font table part if it exists.
    /// </summary>
    public WasiPart? FontTablePart =>
        MainDocumentPart.GetRelatedPart(WasiRelationshipTypes.FontTable);

    /// <summary>
    /// Gets the theme part if it exists.
    /// </summary>
    public WasiPart? ThemePart =>
        MainDocumentPart.GetRelatedPart(WasiRelationshipTypes.Theme);

    /// <summary>
    /// Gets the footnotes part if it exists.
    /// </summary>
    public WasiPart? FootnotesPart =>
        MainDocumentPart.GetRelatedPart(WasiRelationshipTypes.Footnotes);

    IDocumentPart? IWordDocument.FootnotesPart => FootnotesPart;

    /// <summary>
    /// Gets the endnotes part if it exists.
    /// </summary>
    public WasiPart? EndnotesPart =>
        MainDocumentPart.GetRelatedPart(WasiRelationshipTypes.Endnotes);

    IDocumentPart? IWordDocument.EndnotesPart => EndnotesPart;

    /// <summary>
    /// Gets the comments part if it exists.
    /// </summary>
    public WasiPart? CommentsPart =>
        MainDocumentPart.GetRelatedPart(WasiRelationshipTypes.Comments);

    IDocumentPart? IWordDocument.CommentsPart => CommentsPart;

    /// <summary>
    /// Gets the settings part if it exists.
    /// </summary>
    public WasiPart? SettingsPart =>
        MainDocumentPart.GetRelatedPart(WasiRelationshipTypes.Settings);

    /// <summary>
    /// Gets all header parts.
    /// </summary>
    public IEnumerable<WasiPart> HeaderParts =>
        MainDocumentPart.GetRelatedParts(WasiRelationshipTypes.Header);

    IEnumerable<IDocumentPart> IWordDocument.HeaderParts => HeaderParts;

    /// <summary>
    /// Gets all footer parts.
    /// </summary>
    public IEnumerable<WasiPart> FooterParts =>
        MainDocumentPart.GetRelatedParts(WasiRelationshipTypes.Footer);

    IEnumerable<IDocumentPart> IWordDocument.FooterParts => FooterParts;

    /// <summary>
    /// Gets a part by URI.
    /// </summary>
    public WasiPart? GetPart(Uri uri)
    {
        if (!_package.PartExists(uri)) return null;
        var part = _package.GetPart(uri);
        return new WasiPart(_package, part);
    }

    /// <summary>
    /// Gets a part by URI string.
    /// </summary>
    public WasiPart? GetPart(string uri)
    {
        return GetPart(new Uri(uri, UriKind.Relative));
    }

    /// <summary>
    /// Gets all parts in the package.
    /// </summary>
    public IEnumerable<WasiPart> GetAllParts()
    {
        foreach (var part in _package.GetParts())
        {
            yield return new WasiPart(_package, part);
        }
    }

    IEnumerable<IDocumentPart> IWordDocument.GetAllParts() => GetAllParts();

    IDocumentPart? IWordDocument.GetPart(string uri) => GetPart(uri);

    /// <summary>
    /// Saves the document.
    /// </summary>
    public void Save()
    {
        if (!_isEditable)
            throw new InvalidOperationException("Document is read-only");

        _package.Save();
    }

    /// <summary>
    /// Saves the document to a new stream.
    /// </summary>
    public void SaveAs(Stream stream)
    {
        // Create a new package and copy all parts
        using var newPackage = SharpCompressPackage.Create(stream);

        foreach (var part in _package.GetParts())
        {
            var newPart = newPackage.CreatePart(
                part.Uri,
                part.ContentType,
                System.IO.Packaging.CompressionOption.Normal);

            using var sourceStream = part.GetStream(FileMode.Open, FileAccess.Read);
            using var targetStream = newPart.GetStream(FileMode.Create, FileAccess.Write);
            sourceStream.CopyTo(targetStream);

            // Copy part relationships
            foreach (var rel in part.Relationships)
            {
                newPart.Relationships.Create(rel.TargetUri, rel.TargetMode, rel.RelationshipType, rel.Id);
            }
        }

        // Copy package relationships
        foreach (var rel in _package.Relationships)
        {
            newPackage.Relationships.Create(rel.TargetUri, rel.TargetMode, rel.RelationshipType, rel.Id);
        }

        newPackage.Save();
    }

    /// <summary>
    /// Gets the document as a byte array.
    /// </summary>
    public byte[] ToByteArray()
    {
        using var ms = new MemoryStream();
        SaveAs(ms);
        return ms.ToArray();
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _package.Dispose();
    }
}

/// <summary>
/// A WASI-compatible PowerPoint document wrapper.
/// </summary>
public class WasiPresentationDocument : IDisposable
{
    private readonly SharpCompressPackage _package;
    private readonly bool _isEditable;
    private WasiPart? _presentationPart;
    private bool _disposed;

    private WasiPresentationDocument(SharpCompressPackage package, bool isEditable)
    {
        _package = package;
        _isEditable = isEditable;
    }

    public static WasiPresentationDocument Open(byte[] data, bool isEditable = false)
    {
        var package = SharpCompressPackage.Open(data, isEditable ? FileAccess.ReadWrite : FileAccess.Read);
        return new WasiPresentationDocument(package, isEditable);
    }

    public static WasiPresentationDocument Open(Stream stream, bool isEditable = false)
    {
        var package = SharpCompressPackage.Open(stream, isEditable ? FileAccess.ReadWrite : FileAccess.Read);
        return new WasiPresentationDocument(package, isEditable);
    }

    public WasiPart PresentationPart
    {
        get
        {
            if (_presentationPart != null) return _presentationPart;

            var rel = _package.Relationships
                .FirstOrDefault(r => r.RelationshipType == WasiRelationshipTypes.OfficeDocument);

            if (rel == null)
                throw new InvalidOperationException("No presentation part found");

            var uri = rel.TargetUri;
            if (!uri.OriginalString.StartsWith("/"))
                uri = new Uri("/" + uri.OriginalString, UriKind.Relative);

            var part = _package.GetPart(uri);
            _presentationPart = new WasiPart(_package, part);
            return _presentationPart;
        }
    }

    public WasiPart? GetPart(Uri uri)
    {
        if (!_package.PartExists(uri)) return null;
        var part = _package.GetPart(uri);
        return new WasiPart(_package, part);
    }

    public IEnumerable<WasiPart> GetAllParts()
    {
        foreach (var part in _package.GetParts())
        {
            yield return new WasiPart(_package, part);
        }
    }

    public void Save() => _package.Save();

    public byte[] ToByteArray()
    {
        using var ms = new MemoryStream();
        using var newPackage = SharpCompressPackage.Create(ms);

        foreach (var part in _package.GetParts())
        {
            var newPart = newPackage.CreatePart(part.Uri, part.ContentType, System.IO.Packaging.CompressionOption.Normal);
            using var sourceStream = part.GetStream(FileMode.Open, FileAccess.Read);
            using var targetStream = newPart.GetStream(FileMode.Create, FileAccess.Write);
            sourceStream.CopyTo(targetStream);
        }

        foreach (var rel in _package.Relationships)
        {
            newPackage.Relationships.Create(rel.TargetUri, rel.TargetMode, rel.RelationshipType, rel.Id);
        }

        newPackage.Save();
        return ms.ToArray();
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _package.Dispose();
    }
}

/// <summary>
/// A WASI-compatible Excel document wrapper.
/// </summary>
public class WasiSpreadsheetDocument : IDisposable
{
    private readonly SharpCompressPackage _package;
    private readonly bool _isEditable;
    private WasiPart? _workbookPart;
    private bool _disposed;

    private WasiSpreadsheetDocument(SharpCompressPackage package, bool isEditable)
    {
        _package = package;
        _isEditable = isEditable;
    }

    public static WasiSpreadsheetDocument Open(byte[] data, bool isEditable = false)
    {
        var package = SharpCompressPackage.Open(data, isEditable ? FileAccess.ReadWrite : FileAccess.Read);
        return new WasiSpreadsheetDocument(package, isEditable);
    }

    public static WasiSpreadsheetDocument Open(Stream stream, bool isEditable = false)
    {
        var package = SharpCompressPackage.Open(stream, isEditable ? FileAccess.ReadWrite : FileAccess.Read);
        return new WasiSpreadsheetDocument(package, isEditable);
    }

    public WasiPart WorkbookPart
    {
        get
        {
            if (_workbookPart != null) return _workbookPart;

            var rel = _package.Relationships
                .FirstOrDefault(r => r.RelationshipType == WasiRelationshipTypes.OfficeDocument);

            if (rel == null)
                throw new InvalidOperationException("No workbook part found");

            var uri = rel.TargetUri;
            if (!uri.OriginalString.StartsWith("/"))
                uri = new Uri("/" + uri.OriginalString, UriKind.Relative);

            var part = _package.GetPart(uri);
            _workbookPart = new WasiPart(_package, part);
            return _workbookPart;
        }
    }

    public WasiPart? GetPart(Uri uri)
    {
        if (!_package.PartExists(uri)) return null;
        var part = _package.GetPart(uri);
        return new WasiPart(_package, part);
    }

    public IEnumerable<WasiPart> GetAllParts()
    {
        foreach (var part in _package.GetParts())
        {
            yield return new WasiPart(_package, part);
        }
    }

    public void Save() => _package.Save();

    public byte[] ToByteArray()
    {
        using var ms = new MemoryStream();
        using var newPackage = SharpCompressPackage.Create(ms);

        foreach (var part in _package.GetParts())
        {
            var newPart = newPackage.CreatePart(part.Uri, part.ContentType, System.IO.Packaging.CompressionOption.Normal);
            using var sourceStream = part.GetStream(FileMode.Open, FileAccess.Read);
            using var targetStream = newPart.GetStream(FileMode.Create, FileAccess.Write);
            sourceStream.CopyTo(targetStream);
        }

        foreach (var rel in _package.Relationships)
        {
            newPackage.Relationships.Create(rel.TargetUri, rel.TargetMode, rel.RelationshipType, rel.Id);
        }

        newPackage.Save();
        return ms.ToArray();
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _package.Dispose();
    }
}
