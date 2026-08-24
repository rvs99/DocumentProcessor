using System.IO.Compression;

namespace DocumentProcessor.Core.Security;

/// <summary>
/// Bounds applied to an untrusted document before it is parsed.
/// <para>
/// A .docx is a ZIP, and every OOXML entry point expands it into an in-memory DOM. Without limits a
/// 10 KB upload whose <c>word/document.xml</c> inflates to several gigabytes is fully materialised
/// before anything notices — a one-request memory-exhaustion attack in a shared host.
/// </para>
/// The defaults are sized for contract documents, which are text-heavy but rarely enormous. Raise
/// them deliberately if a tenant legitimately handles larger files.
/// </summary>
public sealed record DocumentLimits
{
    /// <summary>Maximum size of the file as supplied, on disk or in memory.</summary>
    public long MaxCompressedBytes { get; init; } = 64L * 1024 * 1024;

    /// <summary>Maximum total size of all entries once decompressed.</summary>
    public long MaxUncompressedBytes { get; init; } = 512L * 1024 * 1024;

    /// <summary>
    /// Maximum decompressed-to-compressed ratio. Ordinary Office documents land well under 20:1;
    /// a zip bomb is engineered for ratios in the thousands, so this catches the shape of the
    /// attack rather than only its absolute size.
    /// </summary>
    public int MaxCompressionRatio { get; init; } = 200;

    /// <summary>Maximum number of entries in the package. A normal .docx has tens.</summary>
    public int MaxEntryCount { get; init; } = 5_000;

    /// <summary>The defaults described above.</summary>
    public static DocumentLimits Default { get; } = new();

    /// <summary>
    /// Every check disabled. Only appropriate for documents this system produced itself; never for
    /// anything that arrived from outside.
    /// </summary>
    public static DocumentLimits Unbounded { get; } = new()
    {
        MaxCompressedBytes = long.MaxValue,
        MaxUncompressedBytes = long.MaxValue,
        MaxCompressionRatio = int.MaxValue,
        MaxEntryCount = int.MaxValue,
    };
}

/// <summary>
/// Validates that a document package is within <see cref="DocumentLimits"/> before any parser sees
/// it. Inspects only the ZIP central directory — entry count and declared sizes — so it is cheap
/// and, critically, never decompresses anything to find out how big it would be.
/// </summary>
public static class DocumentPackageGuard
{
    /// <summary>Validates an in-memory package.</summary>
    /// <exception cref="DocumentTooComplexException">A limit is exceeded.</exception>
    /// <exception cref="CorruptDocumentException">The bytes are not a readable ZIP package.</exception>
    public static void Validate(byte[] documentBytes, DocumentLimits? limits = null)
    {
        ArgumentNullException.ThrowIfNull(documentBytes);
        limits ??= DocumentLimits.Default;

        GuardCompressedSize(documentBytes.LongLength, limits);

        using var stream = new MemoryStream(documentBytes, writable: false);
        ValidateArchive(stream, documentBytes.LongLength, limits);
    }

    /// <summary>Validates a package on disk without reading its contents into memory.</summary>
    /// <exception cref="DocumentTooComplexException">A limit is exceeded.</exception>
    /// <exception cref="CorruptDocumentException">The file is not a readable ZIP package.</exception>
    public static void ValidateFile(string documentPath, DocumentLimits? limits = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(documentPath);
        limits ??= DocumentLimits.Default;

        var info = new FileInfo(documentPath);
        if (!info.Exists)
            throw new FileNotFoundException("Document not found.", documentPath);

        GuardCompressedSize(info.Length, limits);

        using var stream = File.OpenRead(documentPath);
        ValidateArchive(stream, info.Length, limits);
    }

    private static void GuardCompressedSize(long compressedBytes, DocumentLimits limits)
    {
        if (compressedBytes > limits.MaxCompressedBytes)
        {
            throw new DocumentTooComplexException(
                $"Document is {compressedBytes:N0} bytes; the limit is {limits.MaxCompressedBytes:N0}.");
        }
    }

    private static void ValidateArchive(Stream stream, long compressedBytes, DocumentLimits limits)
    {
        ZipArchive archive;
        try
        {
            archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        }
        catch (InvalidDataException ex)
        {
            throw new CorruptDocumentException("Document is not a readable Office package (invalid ZIP).", ex);
        }

        using (archive)
        {
            if (archive.Entries.Count > limits.MaxEntryCount)
            {
                throw new DocumentTooComplexException(
                    $"Document contains {archive.Entries.Count:N0} parts; the limit is {limits.MaxEntryCount:N0}.");
            }

            // Length is the entry's declared uncompressed size, read from the central directory —
            // no decompression happens here. A lying header would be caught by the parser later;
            // the point of this check is to reject the obvious case cheaply.
            long uncompressedTotal = 0;
            foreach (var entry in archive.Entries)
            {
                uncompressedTotal += entry.Length;
                if (uncompressedTotal > limits.MaxUncompressedBytes)
                {
                    throw new DocumentTooComplexException(
                        $"Document expands to more than {limits.MaxUncompressedBytes:N0} bytes; refusing to parse it.");
                }
            }

            if (compressedBytes > 0 && limits.MaxCompressionRatio != int.MaxValue)
            {
                var ratio = (double)uncompressedTotal / compressedBytes;
                if (ratio > limits.MaxCompressionRatio)
                {
                    throw new DocumentTooComplexException(
                        $"Document expands {ratio:F0}:1, beyond the {limits.MaxCompressionRatio}:1 limit — this is the " +
                        "signature of a decompression bomb.");
                }
            }
        }
    }
}
