using DocumentProcessor.Core.Comparison;
using DocumentProcessor.Core.Conversion;
using DocumentProcessor.Core.Samples;
using DocumentProcessor.Core.Watermarking;
using UglyToad.PdfPig;

namespace DocumentProcessor.Tests.Conversion;

/// <summary>
/// Verifies the recommended pipeline for producing a non-selectable watermarked PDF from a
/// watermarked docx: strip the watermark before conversion, convert the clean docx via LibreOffice,
/// then apply <see cref="PdfWatermarkService"/> (which rasterizes) to the result — rather than
/// relying on the docx watermark surviving LibreOffice's conversion, which (per
/// <see cref="WatermarkConversionTests"/>) produces real, selectable PDF text, not an image.
/// </summary>
public class WatermarkPipelineTests : IDisposable
{
    private readonly string _docxPath = TestFiles.NewTempPath(".docx");
    private readonly string _cleanPdfPath = TestFiles.NewTempPath(".pdf");
    private readonly string _watermarkedPdfPath = TestFiles.NewTempPath(".pdf");

    public WatermarkPipelineTests()
    {
        SampleDocumentFactory.CreateBasicDocument(_docxPath, "Pipeline Test", ["Body text that should remain selectable."]);
    }

    [Fact]
    public async Task Strip_convert_reapply_produces_a_visible_but_non_selectable_pdf_watermark()
    {
        var docxWatermarkService = new DocxWatermarkService();
        docxWatermarkService.AddTextWatermark(_docxPath, "DRAFT");

        var removed = docxWatermarkService.RemoveWatermark(_docxPath);
        Assert.True(removed, "Expected a watermark to be present and removed before conversion.");

        await new WordToPdfConverter(TestFiles.ConversionOptions()).ConvertAsync(_docxPath, _cleanPdfPath);
        new PdfWatermarkService().AddTextWatermark(_cleanPdfPath, _watermarkedPdfPath, "DRAFT");

        using var doc = PdfDocument.Open(_watermarkedPdfPath);
        var text = doc.GetPage(1).Text;

        // Real body text should still be extractable/selectable...
        Assert.Contains("Body text that should remain selectable", text);
        // ...but the watermark should not be, since PdfWatermarkService rasterizes it.
        Assert.DoesNotContain("DRAFT", text);
    }

    [Fact]
    public async Task Strip_convert_reapply_watermark_renders_behind_the_real_text()
    {
        var docxWatermarkService = new DocxWatermarkService();
        docxWatermarkService.AddTextWatermark(_docxPath, "DRAFT");
        docxWatermarkService.RemoveWatermark(_docxPath);

        await new WordToPdfConverter(TestFiles.ConversionOptions()).ConvertAsync(_docxPath, _cleanPdfPath);
        new PdfWatermarkService().AddTextWatermark(_cleanPdfPath, _watermarkedPdfPath, "DRAFT");

        // The clean (pre-watermark) and final (post-watermark) PDFs should visibly differ...
        var visualDiff = new PdfComparisonService().CompareVisual(_cleanPdfPath, _watermarkedPdfPath, differenceThresholdPercent: 0.0001);
        Assert.False(visualDiff.PagesMatch);
        Assert.True(visualDiff.PerPageDifferencePercent.Single() > 0);
    }

    public void Dispose()
    {
        File.Delete(_docxPath);
        if (File.Exists(_cleanPdfPath))
            File.Delete(_cleanPdfPath);
        if (File.Exists(_watermarkedPdfPath))
            File.Delete(_watermarkedPdfPath);
    }
}
