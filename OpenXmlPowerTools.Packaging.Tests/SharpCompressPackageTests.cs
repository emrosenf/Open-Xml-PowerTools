// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using Xunit;

namespace OpenXmlPowerTools.Packaging.Tests;

public class SharpCompressPackageTests
{
    private static readonly string TestFilesDir = Path.Combine(
        AppDomain.CurrentDomain.BaseDirectory,
        "..", "..", "..", "..", "TestFiles", "WC");

    [Fact]
    public void CanOpenAndReadDocxParts()
    {
        // Arrange
        var docxPath = Path.Combine(TestFilesDir, "WC001-Digits.docx");
        var docxBytes = File.ReadAllBytes(docxPath);

        // Act
        using var package = SharpCompressPackage.Open(docxBytes);

        // Assert
        var parts = package.GetParts().ToList();
        Assert.NotEmpty(parts);

        // Should have main document part
        var documentPart = parts.FirstOrDefault(p =>
            p.Uri.OriginalString.Contains("document.xml"));
        Assert.NotNull(documentPart);
        Assert.Equal(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            documentPart.ContentType);
    }

    [Fact]
    public void CanReadDocumentContent()
    {
        // Arrange
        var docxPath = Path.Combine(TestFilesDir, "WC001-Digits.docx");
        var docxBytes = File.ReadAllBytes(docxPath);

        // Act
        using var package = SharpCompressPackage.Open(docxBytes);
        var documentPart = package.GetParts()
            .First(p => p.Uri.OriginalString.Contains("document.xml"));

        using var stream = documentPart.GetStream(FileMode.Open, FileAccess.Read);
        using var reader = new StreamReader(stream);
        var content = reader.ReadToEnd();

        // Assert
        Assert.Contains("w:document", content);
    }

    [Fact]
    public void CanReadPackageRelationships()
    {
        // Arrange
        var docxPath = Path.Combine(TestFilesDir, "WC001-Digits.docx");
        var docxBytes = File.ReadAllBytes(docxPath);

        // Act
        using var package = SharpCompressPackage.Open(docxBytes);
        var relationships = package.Relationships.ToList();

        // Assert
        Assert.NotEmpty(relationships);

        // Should have document relationship
        var docRel = relationships.FirstOrDefault(r =>
            r.RelationshipType.Contains("officeDocument"));
        Assert.NotNull(docRel);
    }

    [Fact]
    public void CanReadPartRelationships()
    {
        // Arrange
        var docxPath = Path.Combine(TestFilesDir, "WC001-Digits.docx");
        var docxBytes = File.ReadAllBytes(docxPath);

        // Act
        using var package = SharpCompressPackage.Open(docxBytes);
        var documentPart = package.GetParts()
            .First(p => p.Uri.OriginalString.Contains("document.xml"));
        var relationships = documentPart.Relationships.ToList();

        // Assert - document should have relationships to styles, settings, etc.
        Assert.NotEmpty(relationships);
    }

    [Fact]
    public void CanAccessPackageProperties()
    {
        // Arrange
        var docxPath = Path.Combine(TestFilesDir, "WC001-Digits.docx");
        var docxBytes = File.ReadAllBytes(docxPath);

        // Act
        using var package = SharpCompressPackage.Open(docxBytes);
        var properties = package.PackageProperties;

        // Assert - properties object should exist (values may or may not be set)
        Assert.NotNull(properties);
    }

    [Fact]
    public void CanCreateNewPackage()
    {
        // Arrange
        using var stream = new MemoryStream();

        // Act
        using var package = SharpCompressPackage.Create(stream);

        // Assert
        Assert.Empty(package.GetParts());
        Assert.Equal(FileAccess.ReadWrite, package.FileOpenAccess);
    }

    [Fact]
    public void CanCreatePartInNewPackage()
    {
        // Arrange
        using var stream = new MemoryStream();
        using var package = SharpCompressPackage.Create(stream);

        // Act
        var uri = new Uri("/word/document.xml", UriKind.Relative);
        var part = package.CreatePart(uri,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            System.IO.Packaging.CompressionOption.Normal);

        // Write some content
        using (var partStream = part.GetStream(FileMode.Create, FileAccess.Write))
        using (var writer = new StreamWriter(partStream))
        {
            writer.Write("<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>");
        }

        // Assert
        Assert.Single(package.GetParts());
        Assert.True(package.PartExists(uri));
    }

    [Fact]
    public void CanSaveAndReloadPackage()
    {
        // Arrange
        using var stream = new MemoryStream();
        var uri = new Uri("/word/document.xml", UriKind.Relative);
        var content = "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body/></w:document>";

        // Create and save
        using (var package = SharpCompressPackage.Create(stream))
        {
            var part = package.CreatePart(uri,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                System.IO.Packaging.CompressionOption.Normal);

            using (var partStream = part.GetStream(FileMode.Create, FileAccess.Write))
            using (var writer = new StreamWriter(partStream))
            {
                writer.Write(content);
            }

            package.Save();
        }

        // Reload
        stream.Position = 0;
        using var reloaded = SharpCompressPackage.Open(stream);

        // Assert
        Assert.Single(reloaded.GetParts());
        var reloadedPart = reloaded.GetPart(uri);
        using var readStream = reloadedPart.GetStream(FileMode.Open, FileAccess.Read);
        using var reader = new StreamReader(readStream);
        var readContent = reader.ReadToEnd();

        Assert.Contains("w:document", readContent);
        Assert.Contains("w:body", readContent);
    }
}
