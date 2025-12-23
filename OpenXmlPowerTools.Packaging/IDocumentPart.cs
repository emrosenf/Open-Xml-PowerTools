// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Xml.Linq;

namespace OpenXmlPowerTools.Packaging;

/// <summary>
/// Abstraction for document parts that works with both OpenXML SDK and WASI implementations.
/// </summary>
public interface IDocumentPart
{
    /// <summary>
    /// Gets the URI of this part within the package.
    /// </summary>
    Uri Uri { get; }

    /// <summary>
    /// Gets the content type of this part.
    /// </summary>
    string ContentType { get; }

    /// <summary>
    /// Gets the XML document for this part.
    /// </summary>
    XDocument GetXDocument();

    /// <summary>
    /// Saves the current XDocument back to the part.
    /// </summary>
    void PutXDocument();

    /// <summary>
    /// Saves a new XDocument to the part.
    /// </summary>
    void PutXDocument(XDocument xDocument);

    /// <summary>
    /// Gets the root element of the XDocument.
    /// </summary>
    XElement? RootElement { get; }

    /// <summary>
    /// Gets a related part by relationship ID.
    /// </summary>
    IDocumentPart? GetPartById(string relationshipId);

    /// <summary>
    /// Gets all parts related by a specific relationship type.
    /// </summary>
    IEnumerable<IDocumentPart> GetRelatedParts(string relationshipType);

    /// <summary>
    /// Gets the raw bytes of this part.
    /// </summary>
    byte[] GetBytes();
}

/// <summary>
/// Abstraction for Word documents that works with both OpenXML SDK and WASI implementations.
/// </summary>
public interface IWordDocument : IDisposable
{
    /// <summary>
    /// Gets the main document part.
    /// </summary>
    IDocumentPart MainDocumentPart { get; }

    /// <summary>
    /// Gets the styles part, if present.
    /// </summary>
    IDocumentPart? StyleDefinitionsPart { get; }

    /// <summary>
    /// Gets the numbering definitions part, if present.
    /// </summary>
    IDocumentPart? NumberingDefinitionsPart { get; }

    /// <summary>
    /// Gets the footnotes part, if present.
    /// </summary>
    IDocumentPart? FootnotesPart { get; }

    /// <summary>
    /// Gets the endnotes part, if present.
    /// </summary>
    IDocumentPart? EndnotesPart { get; }

    /// <summary>
    /// Gets the comments part, if present.
    /// </summary>
    IDocumentPart? CommentsPart { get; }

    /// <summary>
    /// Gets all header parts.
    /// </summary>
    IEnumerable<IDocumentPart> HeaderParts { get; }

    /// <summary>
    /// Gets all footer parts.
    /// </summary>
    IEnumerable<IDocumentPart> FooterParts { get; }

    /// <summary>
    /// Gets all parts in the document.
    /// </summary>
    IEnumerable<IDocumentPart> GetAllParts();

    /// <summary>
    /// Gets a part by its URI.
    /// </summary>
    IDocumentPart? GetPart(string uri);

    /// <summary>
    /// Converts the document to a byte array.
    /// </summary>
    byte[] ToByteArray();
}
