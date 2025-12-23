// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO.Packaging;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using SharpCompress.Archives;
using SharpCompress.Archives.Zip;
using SharpCompress.Common;
using SharpCompress.Writers;

// Alias to disambiguate from System.IO.Compression.ZipArchive
using SCZipArchive = SharpCompress.Archives.Zip.ZipArchive;
using SCZipArchiveEntry = SharpCompress.Archives.Zip.ZipArchiveEntry;

namespace OpenXmlPowerTools.Packaging;

/// <summary>
/// A pure managed implementation of IPackage using SharpCompress.
/// This implementation does not rely on System.IO.Packaging native code,
/// making it suitable for WASI and other environments without native ZIP support.
/// </summary>
public class SharpCompressPackage : IPackage, IDisposable
{
    private readonly SCZipArchive _archive;
    private readonly Stream _stream;
    private readonly bool _ownsStream;
    private readonly FileAccess _access;
    private readonly Dictionary<Uri, SharpCompressPart> _parts = new();
    private readonly SharpCompressRelationshipCollection _relationships;
    private readonly SharpCompressPackageProperties _properties;
    private bool _disposed;

    private SharpCompressPackage(SCZipArchive archive, Stream stream, bool ownsStream, FileAccess access)
    {
        _archive = archive;
        _stream = stream;
        _ownsStream = ownsStream;
        _access = access;
        _relationships = new SharpCompressRelationshipCollection(this, null);
        _properties = new SharpCompressPackageProperties(this);

        // Load existing parts from archive
        LoadParts();
    }

    /// <summary>
    /// Opens a package from a stream.
    /// </summary>
    public static SharpCompressPackage Open(Stream stream, FileAccess access = FileAccess.ReadWrite)
    {
        if (stream.Length == 0)
        {
            // New empty package
            var archive = SCZipArchive.Create();
            return new SharpCompressPackage(archive, stream, false, access);
        }

        // Open existing package
        var existingArchive = SCZipArchive.Open(stream);
        return new SharpCompressPackage(existingArchive, stream, false, access);
    }

    /// <summary>
    /// Opens a package from a byte array.
    /// </summary>
    public static SharpCompressPackage Open(byte[] data, FileAccess access = FileAccess.ReadWrite)
    {
        var stream = new MemoryStream(data);
        var archive = SCZipArchive.Open(stream);
        return new SharpCompressPackage(archive, stream, true, access);
    }

    /// <summary>
    /// Creates a new empty package.
    /// </summary>
    public static SharpCompressPackage Create(Stream stream)
    {
        var archive = SCZipArchive.Create();
        return new SharpCompressPackage(archive, stream, false, FileAccess.ReadWrite);
    }

    public FileAccess FileOpenAccess => _access;

    public IPackageProperties PackageProperties => _properties;

    public IRelationshipCollection Relationships => _relationships;

    internal SCZipArchive Archive => _archive;

    public IEnumerable<IPackagePart> GetParts() => _parts.Values;

    public IPackagePart GetPart(Uri uri)
    {
        var normalizedUri = NormalizeUri(uri);
        if (_parts.TryGetValue(normalizedUri, out var part))
        {
            return part;
        }
        throw new InvalidOperationException($"Part not found: {uri}");
    }

    public bool PartExists(Uri uri)
    {
        var normalizedUri = NormalizeUri(uri);
        return _parts.ContainsKey(normalizedUri);
    }

    public IPackagePart CreatePart(Uri uri, string contentType, CompressionOption compressionOption)
    {
        var normalizedUri = NormalizeUri(uri);

        if (_parts.ContainsKey(normalizedUri))
        {
            throw new InvalidOperationException($"Part already exists: {uri}");
        }

        var part = new SharpCompressPart(this, normalizedUri, contentType, compressionOption);
        _parts[normalizedUri] = part;

        return part;
    }

    public void DeletePart(Uri uri)
    {
        var normalizedUri = NormalizeUri(uri);

        _parts.Remove(normalizedUri);
    }

    public void Save()
    {
        if (_access == FileAccess.Read)
        {
            throw new InvalidOperationException("Package is read-only");
        }

        // Write all parts to the archive
        using var outputStream = new MemoryStream();
        using (var writer = WriterFactory.Open(outputStream, ArchiveType.Zip, new WriterOptions(CompressionType.Deflate)))
        {
            // Write [Content_Types].xml
            WriteContentTypes(writer);

            // Write package relationships
            WriteRelationships(writer, "/_rels/.rels", _relationships);

            // Write core properties if they exist
            if (_properties.HasValues)
            {
                WritePackageProperties(writer);
            }

            // Write all parts
            foreach (var part in _parts.Values)
            {
                var entryPath = part.Uri.OriginalString.TrimStart('/');

                // Write part content
                using var partStream = part.GetDataStream();
                if (partStream != null && partStream.Length > 0)
                {
                    partStream.Position = 0;
                    writer.Write(entryPath, partStream);
                }

                // Write part relationships if any
                if (part.Relationships.Any())
                {
                    var relsPath = GetRelationshipPartPath(part.Uri);
                    WriteRelationships(writer, relsPath, part.Relationships);
                }
            }
        }

        // Copy to target stream
        _stream.Position = 0;
        _stream.SetLength(0);
        outputStream.Position = 0;
        outputStream.CopyTo(_stream);
        _stream.Flush();
    }

    private void LoadParts()
    {
        var contentTypes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var defaultTypes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // First, read [Content_Types].xml
        var contentTypesEntry = _archive.Entries.FirstOrDefault(e =>
            e.Key != null && e.Key.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase));

        if (contentTypesEntry != null)
        {
            using var stream = contentTypesEntry.OpenEntryStream();
            var doc = XDocument.Load(stream);
            var ns = doc.Root?.Name.Namespace ?? XNamespace.None;

            foreach (var element in doc.Root?.Elements() ?? Enumerable.Empty<XElement>())
            {
                if (element.Name.LocalName == "Override")
                {
                    var partName = element.Attribute("PartName")?.Value;
                    var type = element.Attribute("ContentType")?.Value;
                    if (partName != null && type != null)
                    {
                        contentTypes[partName] = type;
                    }
                }
                else if (element.Name.LocalName == "Default")
                {
                    var extension = element.Attribute("Extension")?.Value;
                    var type = element.Attribute("ContentType")?.Value;
                    if (extension != null && type != null)
                    {
                        defaultTypes[extension] = type;
                    }
                }
            }
        }

        // Load package-level relationships
        var packageRelsEntry = _archive.Entries.FirstOrDefault(e =>
            e.Key != null && e.Key.Equals("_rels/.rels", StringComparison.OrdinalIgnoreCase));

        if (packageRelsEntry != null)
        {
            _relationships.LoadFromEntry(packageRelsEntry);
        }

        // Load parts (skip special files)
        foreach (var entry in _archive.Entries)
        {
            var path = entry.Key;

            // Skip entries without keys, directories, and special files
            if (path == null || entry.IsDirectory ||
                path.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase) ||
                path.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var uri = new Uri("/" + path, UriKind.Relative);
            var normalizedUri = NormalizeUri(uri);

            // Determine content type
            string? contentType = null;
            var uriString = "/" + path;

            if (contentTypes.TryGetValue(uriString, out var overrideType))
            {
                contentType = overrideType;
            }
            else
            {
                var extension = Path.GetExtension(path).TrimStart('.');
                if (defaultTypes.TryGetValue(extension, out var defaultType))
                {
                    contentType = defaultType;
                }
            }

            contentType ??= "application/octet-stream";

            var part = new SharpCompressPart(this, normalizedUri, contentType, CompressionOption.Normal);
            part.LoadFromEntry(entry);
            _parts[normalizedUri] = part;

            // Load part relationships
            var relsPath = GetRelationshipPartPath(normalizedUri).TrimStart('/');
            var partRelsEntry = _archive.Entries.FirstOrDefault(e =>
                e.Key != null && e.Key.Equals(relsPath, StringComparison.OrdinalIgnoreCase));

            if (partRelsEntry != null)
            {
                ((SharpCompressRelationshipCollection)part.Relationships).LoadFromEntry(partRelsEntry);
            }
        }
    }

    private void WriteContentTypes(IWriter writer)
    {
        var ns = XNamespace.Get("http://schemas.openxmlformats.org/package/2006/content-types");
        var doc = new XDocument(
            new XElement(ns + "Types",
                new XElement(ns + "Default",
                    new XAttribute("Extension", "rels"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")),
                new XElement(ns + "Default",
                    new XAttribute("Extension", "xml"),
                    new XAttribute("ContentType", "application/xml"))
            )
        );

        foreach (var part in _parts.Values)
        {
            doc.Root!.Add(new XElement(ns + "Override",
                new XAttribute("PartName", part.Uri.OriginalString),
                new XAttribute("ContentType", part.ContentType)));
        }

        using var ms = new MemoryStream();
        doc.Save(ms);
        ms.Position = 0;
        writer.Write("[Content_Types].xml", ms);
    }

    private void WriteRelationships(IWriter writer, string path, IRelationshipCollection relationships)
    {
        if (!relationships.Any())
        {
            return;
        }

        var ns = XNamespace.Get("http://schemas.openxmlformats.org/package/2006/relationships");
        var doc = new XDocument(new XElement(ns + "Relationships"));

        foreach (var rel in relationships)
        {
            var element = new XElement(ns + "Relationship",
                new XAttribute("Id", rel.Id),
                new XAttribute("Type", rel.RelationshipType),
                new XAttribute("Target", rel.TargetUri.OriginalString));

            if (rel.TargetMode == TargetMode.External)
            {
                element.Add(new XAttribute("TargetMode", "External"));
            }

            doc.Root!.Add(element);
        }

        using var ms = new MemoryStream();
        doc.Save(ms);
        ms.Position = 0;
        writer.Write(path.TrimStart('/'), ms);
    }

    private void WritePackageProperties(IWriter writer)
    {
        const string corePropsPath = "docProps/core.xml";

        // Ensure relationship exists
        if (!_relationships.Any(r => r.RelationshipType ==
            "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"))
        {
            // Would need to add relationship
        }

        var doc = _properties.ToXDocument();
        using var ms = new MemoryStream();
        doc.Save(ms);
        ms.Position = 0;
        writer.Write(corePropsPath, ms);
    }

    private static string GetRelationshipPartPath(Uri partUri)
    {
        var path = partUri.OriginalString;
        var dir = Path.GetDirectoryName(path)?.Replace('\\', '/') ?? "";
        var name = Path.GetFileName(path);

        if (string.IsNullOrEmpty(dir) || dir == "/")
        {
            return $"/_rels/{name}.rels";
        }

        return $"{dir}/_rels/{name}.rels";
    }

    private static Uri NormalizeUri(Uri uri)
    {
        var path = uri.OriginalString;
        if (!path.StartsWith("/"))
        {
            path = "/" + path;
        }
        return new Uri(path, UriKind.Relative);
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        _archive.Dispose();

        if (_ownsStream)
        {
            _stream.Dispose();
        }
    }
}
