using System.IO.Compression;
using DocumentProcessor.Core;
using DocumentProcessor.Core.Samples;
using DocumentProcessor.Core.Security;
using DocumentProcessor.Core.Sessions;

namespace DocumentProcessor.Tests.Security;

/// <summary>
/// Bounds on untrusted packages. A .docx is a ZIP and every parser expands it into a DOM, so
/// without these a small crafted upload can exhaust a shared host's memory.
/// </summary>
public class DocumentLimitsTests : IDisposable
{
    private readonly List<string> _cleanup = [];

    private string NewRealDocx()
    {
        var path = TestFiles.NewTempPath(".docx");
        _cleanup.Add(path);
        SampleDocumentFactory.CreateBasicDocument(path, "Contract", ["Body text."]);
        return path;
    }

    /// <summary>A tiny archive whose single entry expands enormously — the zip-bomb shape.</summary>
    private static byte[] BuildHighRatioArchive(int uncompressedBytes)
    {
        using var buffer = new MemoryStream();
        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true))
        {
            var entry = archive.CreateEntry("word/document.xml", CompressionLevel.SmallestSize);
            using var entryStream = entry.Open();
            // Highly repetitive content compresses to almost nothing.
            var chunk = new byte[64 * 1024];
            var written = 0;
            while (written < uncompressedBytes)
            {
                var size = Math.Min(chunk.Length, uncompressedBytes - written);
                entryStream.Write(chunk, 0, size);
                written += size;
            }
        }

        return buffer.ToArray();
    }

    [Fact]
    public void A_normal_document_passes_the_default_limits()
    {
        var bytes = File.ReadAllBytes(NewRealDocx());

        DocumentPackageGuard.Validate(bytes);   // must not throw
        DocumentPackageGuard.ValidateFile(NewRealDocx());
    }

    [Fact]
    public void An_oversized_file_is_rejected_before_parsing()
    {
        var bytes = File.ReadAllBytes(NewRealDocx());
        var limits = new DocumentLimits { MaxCompressedBytes = 100 };

        var ex = Assert.Throws<DocumentTooComplexException>(() => DocumentPackageGuard.Validate(bytes, limits));
        Assert.Contains("limit", ex.Message);
    }

    [Fact]
    public void A_high_compression_ratio_payload_is_rejected()
    {
        // ~8 MB of zeros compresses to a few KB — a ratio far beyond anything a real document has.
        var bomb = BuildHighRatioArchive(8 * 1024 * 1024);

        var ex = Assert.Throws<DocumentTooComplexException>(() => DocumentPackageGuard.Validate(bomb));
        Assert.Contains("bomb", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void An_excessive_uncompressed_size_is_rejected()
    {
        var bomb = BuildHighRatioArchive(8 * 1024 * 1024);
        var limits = new DocumentLimits { MaxUncompressedBytes = 1024, MaxCompressionRatio = int.MaxValue };

        Assert.Throws<DocumentTooComplexException>(() => DocumentPackageGuard.Validate(bomb, limits));
    }

    [Fact]
    public void Too_many_entries_is_rejected()
    {
        using var buffer = new MemoryStream();
        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true))
        {
            for (var i = 0; i < 50; i++)
                archive.CreateEntry($"part{i}.xml");
        }

        var limits = new DocumentLimits { MaxEntryCount = 10 };

        Assert.Throws<DocumentTooComplexException>(() => DocumentPackageGuard.Validate(buffer.ToArray(), limits));
    }

    [Fact]
    public void Non_zip_bytes_are_reported_as_corrupt_rather_than_too_complex()
    {
        Assert.Throws<CorruptDocumentException>(() => DocumentPackageGuard.Validate([1, 2, 3, 4, 5, 6, 7, 8]));
    }

    [Fact]
    public void Unbounded_limits_accept_what_the_defaults_reject()
    {
        var bomb = BuildHighRatioArchive(8 * 1024 * 1024);

        Assert.Throws<DocumentTooComplexException>(() => DocumentPackageGuard.Validate(bomb));
        DocumentPackageGuard.Validate(bomb, DocumentLimits.Unbounded);   // must not throw
    }

    [Fact]
    public void DocumentSession_applies_the_limits_by_default()
    {
        var bomb = BuildHighRatioArchive(8 * 1024 * 1024);

        // The bound is enforced at the entry point untrusted uploads actually come through, rather
        // than depending on the caller remembering to validate first.
        Assert.Throws<DocumentTooComplexException>(() => DocumentSession.Open(bomb));
    }

    [Fact]
    public void DocumentSession_still_opens_a_normal_document()
    {
        using var session = DocumentSession.Open(File.ReadAllBytes(NewRealDocx()));

        Assert.NotNull(session.Document.MainDocumentPart);
    }

    public void Dispose()
    {
        foreach (var path in _cleanup)
            if (File.Exists(path)) File.Delete(path);
    }
}
