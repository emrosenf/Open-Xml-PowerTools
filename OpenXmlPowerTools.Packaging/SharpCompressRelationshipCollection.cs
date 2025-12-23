// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections;
using System.IO.Packaging;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

// Alias to disambiguate from System.IO.Compression
using SCZipArchiveEntry = SharpCompress.Archives.Zip.ZipArchiveEntry;

namespace OpenXmlPowerTools.Packaging;

/// <summary>
/// A pure managed implementation of IRelationshipCollection.
/// </summary>
public class SharpCompressRelationshipCollection : IRelationshipCollection
{
    private static readonly XNamespace RelsNs = "http://schemas.openxmlformats.org/package/2006/relationships";

    private readonly SharpCompressPackage _package;
    private readonly SharpCompressPart? _sourcePart;
    private readonly List<SharpCompressRelationship> _relationships = new();
    private int _nextId = 1;

    internal SharpCompressRelationshipCollection(SharpCompressPackage package, SharpCompressPart? sourcePart)
    {
        _package = package;
        _sourcePart = sourcePart;
    }

    public IPackageRelationship Create(Uri targetUri, TargetMode targetMode, string relationshipType)
    {
        var id = GenerateId();
        return Create(targetUri, targetMode, relationshipType, id);
    }

    public IPackageRelationship Create(Uri targetUri, TargetMode targetMode, string relationshipType, string? id)
    {
        id ??= GenerateId();

        if (_relationships.Any(r => r.Id == id))
        {
            throw new InvalidOperationException($"Relationship with id '{id}' already exists");
        }

        var relationship = new SharpCompressRelationship(id, relationshipType, targetUri, targetMode);
        _relationships.Add(relationship);
        return relationship;
    }

    public void Remove(string id)
    {
        var rel = _relationships.FirstOrDefault(r => r.Id == id);
        if (rel != null)
        {
            _relationships.Remove(rel);
        }
    }

    public bool Contains(string id)
    {
        return _relationships.Any(r => r.Id == id);
    }

    public IPackageRelationship this[string id] =>
        _relationships.FirstOrDefault(r => r.Id == id)
        ?? throw new InvalidOperationException($"Relationship '{id}' not found");

    public int Count => _relationships.Count;

    public IEnumerator<IPackageRelationship> GetEnumerator() => _relationships.Cast<IPackageRelationship>().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    internal void LoadFromEntry(SCZipArchiveEntry entry)
    {
        using var stream = entry.OpenEntryStream();
        var doc = XDocument.Load(stream);

        foreach (var element in doc.Root?.Elements(RelsNs + "Relationship") ?? Enumerable.Empty<XElement>())
        {
            var id = element.Attribute("Id")?.Value ?? GenerateId();
            var type = element.Attribute("Type")?.Value ?? "";
            var target = element.Attribute("Target")?.Value ?? "";
            var targetModeAttr = element.Attribute("TargetMode")?.Value;

            var targetMode = targetModeAttr?.Equals("External", StringComparison.OrdinalIgnoreCase) == true
                ? TargetMode.External
                : TargetMode.Internal;

            Uri targetUri;
            if (targetMode == TargetMode.External || target.Contains("://"))
            {
                targetUri = new Uri(target, UriKind.Absolute);
            }
            else
            {
                targetUri = new Uri(target, UriKind.Relative);
            }

            var relationship = new SharpCompressRelationship(id, type, targetUri, targetMode);
            _relationships.Add(relationship);

            // Track highest ID for generation
            if (id.StartsWith("rId", StringComparison.OrdinalIgnoreCase) &&
                int.TryParse(id.Substring(3), out var numericId))
            {
                _nextId = Math.Max(_nextId, numericId + 1);
            }
        }
    }

    private string GenerateId()
    {
        return $"rId{_nextId++}";
    }
}

/// <summary>
/// A simple relationship implementation.
/// </summary>
public class SharpCompressRelationship : IPackageRelationship
{
    public SharpCompressRelationship(string id, string relationshipType, Uri targetUri, TargetMode targetMode, Uri? sourceUri = null)
    {
        Id = id;
        RelationshipType = relationshipType;
        TargetUri = targetUri;
        TargetMode = targetMode;
        SourceUri = sourceUri ?? new Uri("/", UriKind.Relative);
    }

    public string Id { get; }
    public string RelationshipType { get; }
    public Uri TargetUri { get; }
    public TargetMode TargetMode { get; }
    public Uri SourceUri { get; }
}
