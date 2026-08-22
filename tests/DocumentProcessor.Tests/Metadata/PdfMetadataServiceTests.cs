using DocumentProcessor.Core.Metadata;
using PdfSharp.Pdf.IO;

namespace DocumentProcessor.Tests.Metadata;

public class PdfMetadataServiceTests : IDisposable
{
    private readonly string _pdfPath;
    private readonly string _outputPath = TestFiles.NewTempPath(".pdf");
    private readonly PdfMetadataService _sut = new();

    public PdfMetadataServiceTests()
    {
        _pdfPath = TestFiles.NewSimplePdf(1, "Contract");
    }

    [Fact]
    public void SetXmpMetadata_embeds_a_readable_xmp_packet_in_the_output()
    {
        _sut.SetXmpMetadata(_pdfPath, _outputPath, new XmpMetadata(
            Title: "Master Services Agreement",
            Author: "Acme Legal",
            Keywords: ["contract", "legal"]));

        var rawText = System.Text.Encoding.Latin1.GetString(File.ReadAllBytes(_outputPath));
        Assert.Contains("xmpmeta", rawText);
        Assert.Contains("Master Services Agreement", rawText);
        Assert.Contains("Acme Legal", rawText);
        Assert.Contains("W5M0MpCehiHzreSzNTczkc9d", rawText);
    }

    [Fact]
    public void SetXmpMetadata_also_sets_the_classic_info_dictionary_for_older_readers()
    {
        _sut.SetXmpMetadata(_pdfPath, _outputPath, new XmpMetadata(Title: "Master Services Agreement", Author: "Acme Legal"));

        using var doc = PdfReader.Open(_outputPath, PdfDocumentOpenMode.Import);
        Assert.Equal("Master Services Agreement", doc.Info.Title);
        Assert.Equal("Acme Legal", doc.Info.Author);
    }

    [Fact]
    public void SetXmpMetadata_omits_fields_that_were_not_supplied()
    {
        _sut.SetXmpMetadata(_pdfPath, _outputPath, new XmpMetadata(Title: "Only A Title"));

        // Scoped to *this* packet specifically: PDFsharp writes its own separate baseline XMP
        // metadata (with different formatting) into every saved PDF regardless, which does use a
        // "dc:creator" tag of its own for an unrelated purpose — a whole-file substring search
        // would false-positive on that, so isolate the packet this service actually wrote by its
        // unique xpacket id before asserting what it left out.
        var rawText = System.Text.Encoding.Latin1.GetString(File.ReadAllBytes(_outputPath));
        var packetStart = rawText.IndexOf("<x:xmpmeta", StringComparison.Ordinal);
        var packetEnd = rawText.IndexOf("</x:xmpmeta>", StringComparison.Ordinal) + "</x:xmpmeta>".Length;
        var packet = rawText[packetStart..packetEnd];

        Assert.Contains("Only A Title", packet);
        Assert.DoesNotContain("dc:creator", packet);
    }

    public void Dispose()
    {
        File.Delete(_pdfPath);
        if (File.Exists(_outputPath))
            File.Delete(_outputPath);
    }
}
