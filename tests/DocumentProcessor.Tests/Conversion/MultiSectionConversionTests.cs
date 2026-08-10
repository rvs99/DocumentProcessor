using DocumentProcessor.Core.Conversion;
using DocumentProcessor.Core.Layout;
using DocumentProcessor.Core.Samples;
using UglyToad.PdfPig;

namespace DocumentProcessor.Tests.Conversion;

/// <summary>
/// The real test for Phase 4 (multi-section documents): does a document with two independently
/// laid-out sections actually come out of LibreOffice with two different page geometries, not just
/// look right in the docx XML? A section boundary with no visible effect after conversion would be
/// a much worse failure mode than an exception, since nothing would flag it.
/// </summary>
public class MultiSectionConversionTests : IDisposable
{
    private readonly string _docxPath = TestFiles.NewTempPath(".docx");
    private readonly string _pdfPath = TestFiles.NewTempPath(".pdf");

    public MultiSectionConversionTests()
    {
        SampleDocumentFactory.CreateBasicDocument(_docxPath, "Multi-Section Conversion Test",
            ["Exhibit body text.", "Appendix body text."]);
    }

    [Fact]
    public async Task A_landscape_exhibit_followed_by_a_portrait_appendix_converts_with_two_distinct_page_geometries()
    {
        var layout = new PageLayoutService();

        // Split before "Appendix body text." (index 2: Title=0, "Exhibit..."=1, "Appendix..."=2).
        layout.InsertSectionBreak(_docxPath, beforeParagraphIndex: 2);
        layout.SetPageSize(_docxPath, PageSize.Letter(PageOrientation.Landscape), sectionIndex: 0);
        layout.SetPageSize(_docxPath, PageSize.Letter(PageOrientation.Portrait), sectionIndex: 1);

        await new WordToPdfConverter(TestFiles.ConversionOptions()).ConvertAsync(_docxPath, _pdfPath);

        using var doc = PdfDocument.Open(_pdfPath);
        Assert.True(doc.NumberOfPages >= 2, $"Expected at least 2 pages (one per section), got {doc.NumberOfPages}.");

        var firstPage = doc.GetPage(1);
        var lastPage = doc.GetPage(doc.NumberOfPages);

        Assert.True(firstPage.Width > firstPage.Height,
            $"Expected the exhibit's first page to be landscape, got {firstPage.Width}x{firstPage.Height}pt.");
        Assert.True(lastPage.Height > lastPage.Width,
            $"Expected the appendix's last page to be portrait, got {lastPage.Width}x{lastPage.Height}pt.");

        Assert.Contains("Exhibit body text", firstPage.Text);
        Assert.Contains("Appendix body text", lastPage.Text);
    }

    public void Dispose()
    {
        File.Delete(_docxPath);
        if (File.Exists(_pdfPath))
            File.Delete(_pdfPath);
    }
}
