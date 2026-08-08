using DocumentProcessor.Core.Watermarking;
using UglyToad.PdfPig;

namespace DocumentProcessor.Tests.Watermarking;

public class PdfWatermarkServiceTests : IDisposable
{
    private readonly string _pdfPath = TestFiles.NewTestPdf("Watermark Test", ["Some body content."]);
    private readonly string _outputPath = TestFiles.NewTempPath(".pdf");
    private readonly PdfWatermarkService _sut = new();

    [Fact]
    public void AddTextWatermark_produces_a_valid_pdf_containing_the_watermark_text()
    {
        _sut.AddTextWatermark(_pdfPath, _outputPath, "CONFIDENTIAL");

        Assert.True(File.Exists(_outputPath));
        using var doc = PdfDocument.Open(_outputPath);
        var text = doc.GetPage(1).Text;

        Assert.Contains("CONFIDENTIAL", text);
        Assert.Contains("Some body content", text);
    }

    public void Dispose()
    {
        File.Delete(_pdfPath);
        if (File.Exists(_outputPath))
            File.Delete(_outputPath);
    }
}
