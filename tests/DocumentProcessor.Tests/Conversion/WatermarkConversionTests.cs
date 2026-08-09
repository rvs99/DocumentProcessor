using DocumentProcessor.Core.Conversion;
using DocumentProcessor.Core.Samples;
using DocumentProcessor.Core.Watermarking;
using UglyToad.PdfPig;

namespace DocumentProcessor.Tests.Conversion;

/// <summary>
/// Documents a known LibreOffice limitation, found while investigating why a docx watermark
/// wasn't showing up after PDF conversion: converting a .docx watermarked via
/// <see cref="DocxWatermarkService"/> to PDF through <see cref="WordToPdfConverter"/> (LibreOffice
/// headless) does not carry the watermark through. LibreOffice silently drops the watermark's VML
/// "_x0000_t136" WordArt shapetype-based shape during conversion instead of erroring — confirmed
/// independent of the watermark's text content (tested with "DRAFT" and with a longer string
/// containing an em dash), removable/locked mode, and rotation.
///
/// This is scoped specifically to the docx-watermark -> LibreOffice-conversion path: the watermark
/// renders correctly in real MS Word (see DocxWatermarkServiceTests), and PdfWatermarkService's PDF
/// watermarks render correctly too, since that service draws directly onto an already-converted PDF
/// and never touches LibreOffice's VML import.
///
/// If LibreOffice ever adds full support for this VML shapetype, the assertion below will start
/// failing — that's the point: treat a failure here as "go update this test to assert the watermark
/// IS present now," not as a regression in our own code.
/// </summary>
public class WatermarkConversionTests : IDisposable
{
    private readonly string _docxPath = TestFiles.NewTempPath(".docx");
    private readonly string _pdfPath = TestFiles.NewTempPath(".pdf");

    public WatermarkConversionTests()
    {
        SampleDocumentFactory.CreateBasicDocument(_docxPath, "Watermark Conversion Test", ["Body text that should survive conversion."]);
        new DocxWatermarkService().AddTextWatermark(_docxPath, "DRAFT");
    }

    [Fact]
    public async Task Converting_a_watermarked_docx_via_LibreOffice_currently_drops_the_watermark_but_keeps_body_text()
    {
        await new WordToPdfConverter(TestFiles.ConversionOptions()).ConvertAsync(_docxPath, _pdfPath);

        using var doc = PdfDocument.Open(_pdfPath);
        var text = doc.GetPage(1).Text;

        Assert.Contains("Body text that should survive conversion", text);

        // KNOWN LIMITATION (see class remarks) — not desired behavior. Confirms the watermark shape
        // itself, not just its text styling, is what LibreOffice is dropping.
        Assert.DoesNotContain("DRAFT", text);
    }

    public void Dispose()
    {
        File.Delete(_docxPath);
        if (File.Exists(_pdfPath))
            File.Delete(_pdfPath);
    }
}
