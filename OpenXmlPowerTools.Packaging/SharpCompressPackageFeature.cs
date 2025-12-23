// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools.Packaging;

/// <summary>
/// Extension methods for using SharpCompress packages with OpenXML SDK.
/// Note: IPackageFeature is internal to the SDK, so we cannot implement it directly.
/// Instead, users must work with SharpCompressPackage at a lower level.
/// </summary>
public static class SharpCompressExtensions
{
    /// <summary>
    /// Opens a package from a byte array using SharpCompress for ZIP operations.
    /// This is WASI-compatible as it doesn't use native compression.
    /// </summary>
    /// <param name="data">The document data as a byte array</param>
    /// <param name="isEditable">Whether to open for editing</param>
    /// <returns>A SharpCompressPackage that can be used to access document parts</returns>
    public static SharpCompressPackage OpenPackage(byte[] data, bool isEditable = false)
    {
        return SharpCompressPackage.Open(data, isEditable ? FileAccess.ReadWrite : FileAccess.Read);
    }

    /// <summary>
    /// Opens a package from a stream using SharpCompress for ZIP operations.
    /// This is WASI-compatible as it doesn't use native compression.
    /// </summary>
    /// <param name="stream">The document stream</param>
    /// <param name="isEditable">Whether to open for editing</param>
    /// <returns>A SharpCompressPackage that can be used to access document parts</returns>
    public static SharpCompressPackage OpenPackage(Stream stream, bool isEditable = false)
    {
        return SharpCompressPackage.Open(stream, isEditable ? FileAccess.ReadWrite : FileAccess.Read);
    }

    /// <summary>
    /// Creates a new empty package using SharpCompress for ZIP operations.
    /// </summary>
    /// <param name="stream">The stream to write the package to</param>
    /// <returns>A SharpCompressPackage for the new document</returns>
    public static SharpCompressPackage CreatePackage(Stream stream)
    {
        return SharpCompressPackage.Create(stream);
    }
}
