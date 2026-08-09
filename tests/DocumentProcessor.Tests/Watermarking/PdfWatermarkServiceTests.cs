using System.Text;
using DocumentProcessor.Core.Comparison;
using DocumentProcessor.Core.Watermarking;
using PdfSharp.Pdf.IO;
using UglyToad.PdfPig;

namespace DocumentProcessor.Tests.Watermarking;

public class PdfWatermarkServiceTests : IDisposable
{
    private readonly string _pdfPath = TestFiles.NewTestPdf("Watermark Test", ["Some body content."]);
    private readonly string _outputPath = TestFiles.NewTempPath(".pdf");
    private readonly PdfWatermarkService _sut = new();

    [Fact]
    public void AddTextWatermark_leaves_the_real_body_text_selectable_and_unchanged()
    {
        _sut.AddTextWatermark(_pdfPath, _outputPath, "CONFIDENTIAL");

        Assert.True(File.Exists(_outputPath));
        using var doc = PdfDocument.Open(_outputPath);
        Assert.Contains("Some body content", doc.GetPage(1).Text);
    }

    [Fact]
    public void AddTextWatermark_is_not_extractable_as_text_because_it_is_rasterized_not_drawn_as_pdf_text()
    {
        // This is the actual fix under test: watermark text used to be drawn with XGraphics.DrawString,
        // which is real selectable/extractable PDF text — dragging a selection over it (or copy-pasting
        // page text) picked up the watermark string along with the real content. Rasterizing it to an
        // image means it can never appear in extracted text, so it can't be selected either.
        _sut.AddTextWatermark(_pdfPath, _outputPath, "CONFIDENTIAL");

        using var doc = PdfDocument.Open(_outputPath);
        Assert.DoesNotContain("CONFIDENTIAL", doc.GetPage(1).Text);
    }

    [Fact]
    public void AddTextWatermark_visibly_changes_the_rendered_page()
    {
        _sut.AddTextWatermark(_pdfPath, _outputPath, "CONFIDENTIAL");

        var visualDiff = new PdfComparisonService().CompareVisual(_pdfPath, _outputPath, differenceThresholdPercent: 0.0001);
        Assert.False(visualDiff.PagesMatch);
        Assert.True(visualDiff.PerPageDifferencePercent.Single() > 0);
    }

    [Fact]
    public void AddTextWatermark_paints_the_watermark_before_the_page_text_so_text_renders_on_top()
    {
        // XGraphics.FromPdfPage defaults to Append (new drawing on top of existing content), which
        // would paint the watermark over the real text instead of behind it. The fix passes
        // XGraphicsPdfPageOptions.Prepend explicitly. Verify by reading the page's content stream(s)
        // directly: PDF content streams execute — and therefore paint — in array order, so the
        // watermark's image-drawing operator ("Do") must appear before the text-drawing operator
        // ("BT", begin text) across the concatenated stream content.
        _sut.AddTextWatermark(_pdfPath, _outputPath, "CONFIDENTIAL");

        using var doc = PdfReader.Open(_outputPath, PdfDocumentOpenMode.Import);
        var page = doc.Pages[0];
        var combined = new StringBuilder();
        for (var i = 0; i < page.Contents.Elements.Count; i++)
            combined.Append(Encoding.Latin1.GetString(page.Contents.Elements.GetDictionary(i)!.Stream!.UnfilteredValue!));

        var content = combined.ToString();
        var imageDrawIndex = content.IndexOf(" Do", StringComparison.Ordinal);
        var textBeginIndex = content.IndexOf("BT", StringComparison.Ordinal);

        Assert.True(imageDrawIndex >= 0, "Expected an image-drawing ('Do') operator in the page content.");
        Assert.True(textBeginIndex >= 0, "Expected a text-drawing ('BT') operator in the page content.");
        Assert.True(imageDrawIndex < textBeginIndex, "Watermark image must be painted before (behind) the page text.");
    }

    public void Dispose()
    {
        File.Delete(_pdfPath);
        if (File.Exists(_outputPath))
            File.Delete(_outputPath);
    }
}
