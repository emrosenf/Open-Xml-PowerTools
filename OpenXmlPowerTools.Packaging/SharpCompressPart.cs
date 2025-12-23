// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO.Packaging;
using DocumentFormat.OpenXml.Packaging;

// Alias to disambiguate from System.IO.Compression
using SCZipArchiveEntry = SharpCompress.Archives.Zip.ZipArchiveEntry;

namespace OpenXmlPowerTools.Packaging;

/// <summary>
/// A pure managed implementation of IPackagePart using SharpCompress.
/// </summary>
public class SharpCompressPart : IPackagePart
{
    private readonly SharpCompressPackage _package;
    private readonly Uri _uri;
    private readonly string _contentType;
    private readonly CompressionOption _compressionOption;
    private readonly SharpCompressRelationshipCollection _relationships;
    private MemoryStream? _data;

    internal SharpCompressPart(
        SharpCompressPackage package,
        Uri uri,
        string contentType,
        CompressionOption compressionOption)
    {
        _package = package;
        _uri = uri;
        _contentType = contentType;
        _compressionOption = compressionOption;
        _relationships = new SharpCompressRelationshipCollection(package, this);
    }

    public IPackage Package => _package;

    public Uri Uri => _uri;

    public string ContentType => _contentType;

    public IRelationshipCollection Relationships => _relationships;

    internal CompressionOption CompressionOption => _compressionOption;

    public Stream GetStream(FileMode mode, FileAccess access)
    {
        if (_data == null)
        {
            _data = new MemoryStream();
        }

        if (mode == FileMode.Create || mode == FileMode.Truncate)
        {
            _data.SetLength(0);
        }

        // Return a wrapper stream that allows reading/writing to our buffer
        return new PartStream(_data, access);
    }

    internal void LoadFromEntry(SCZipArchiveEntry entry)
    {
        _data = new MemoryStream();
        using var entryStream = entry.OpenEntryStream();
        entryStream.CopyTo(_data);
        _data.Position = 0;
    }

    internal Stream? GetDataStream()
    {
        return _data;
    }

    /// <summary>
    /// A stream wrapper that controls access to the underlying data.
    /// </summary>
    private class PartStream : Stream
    {
        private readonly MemoryStream _inner;
        private readonly FileAccess _access;
        private long _position;

        public PartStream(MemoryStream inner, FileAccess access)
        {
            _inner = inner;
            _access = access;
            _position = 0;
        }

        public override bool CanRead => _access.HasFlag(FileAccess.Read);
        public override bool CanSeek => true;
        public override bool CanWrite => _access.HasFlag(FileAccess.Write);
        public override long Length => _inner.Length;

        public override long Position
        {
            get => _position;
            set => _position = value;
        }

        public override void Flush()
        {
            _inner.Flush();
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            if (!CanRead)
                throw new NotSupportedException("Stream does not support reading");

            _inner.Position = _position;
            var bytesRead = _inner.Read(buffer, offset, count);
            _position = _inner.Position;
            return bytesRead;
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            switch (origin)
            {
                case SeekOrigin.Begin:
                    _position = offset;
                    break;
                case SeekOrigin.Current:
                    _position += offset;
                    break;
                case SeekOrigin.End:
                    _position = _inner.Length + offset;
                    break;
            }
            return _position;
        }

        public override void SetLength(long value)
        {
            if (!CanWrite)
                throw new NotSupportedException("Stream does not support writing");

            _inner.SetLength(value);
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            if (!CanWrite)
                throw new NotSupportedException("Stream does not support writing");

            _inner.Position = _position;
            _inner.Write(buffer, offset, count);
            _position = _inner.Position;
        }
    }
}
