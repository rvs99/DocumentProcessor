using DocumentProcessor.Core.Conversion;
using DocumentProcessor.Core.Samples;
using DocumentProcessor.Core.Watermarking;
using UglyToad.PdfPig;

namespace DocumentProcessor.Tests.Conversion;

/// <summary>
/// Regression test for a LibreOffice VML limitation found while investigating why a docx watermark
/// wasn't showing up after PDF conversion: LibreOffice's VML importer does not support
/// <c>v:textpath</c> (Word's WordArt curve-fit text mechanism) at all — confirmed via a systematic
/// diagnostic sweep independent of the shape's size, position, rotation, z-index, and whether it's
/// in the header or body. <c>v:textbox</c> (a normal text box holding regular WordprocessingML
/// paragraph content) is supported by both Word and LibreOffice, which is why
/// <see cref="DocxWatermarkService"/> builds the watermark shape that way instead of using Word's
/// own textpath-based watermark generator output.
///
/// Note what this test does NOT claim: the watermark surviving conversion here is real
/// WordprocessingML text, which LibreOffice's PDF export renders as genuine selectable/extractable
/// PDF text — fine for a docx you're only ever going to open in Word, but not what you want for a
/// PDF a reader could select and delete. For a non-selectable PDF watermark, see
/// <see cref="WatermarkPipelineTests"/>, which strips the watermark before conversion and reapplies
/// it via <see cref="Watermarking.PdfWatermarkService"/> (rasterized) afterwards instead.
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
    public async Task Converting_a_watermarked_docx_via_LibreOffice_preserves_both_the_watermark_and_body_text()
    {
        await new WordToPdfConverter(TestFiles.ConversionOptions()).ConvertAsync(_docxPath, _pdfPath);

        using var doc = PdfDocument.Open(_pdfPath);
        var text = doc.GetPage(1).Text;

        Assert.Contains("Body text that should survive conversion", text);
        Assert.Contains("DRAFT", text);
    }

    public void Dispose()
    {
        File.Delete(_docxPath);
        if (File.Exists(_pdfPath))
            File.Delete(_pdfPath);
    }
}
