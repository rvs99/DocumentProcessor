using DocumentProcessor.Core.Conversion;
using DocumentProcessor.Core.Layout;
using DocumentProcessor.Core.Samples;
using UglyToad.PdfPig;

namespace DocumentProcessor.Tests.Conversion;

/// <summary>
/// Page setup only has value if it survives the docx -> PDF path, since that's the artifact most
/// callers actually distribute. This verifies landscape orientation and custom margins produced by
/// <see cref="PageLayoutService"/> come through a real LibreOffice conversion correctly, not just
/// that the docx XML looks right.
/// </summary>
public class LayoutConversionTests : IDisposable
{
    private readonly string _docxPath = TestFiles.NewTempPath(".docx");
    private readonly string _pdfPath = TestFiles.NewTempPath(".pdf");

    public LayoutConversionTests()
    {
        SampleDocumentFactory.CreateBasicDocument(_docxPath, "Layout Conversion Test", ["Body text."]);
    }

    [Fact]
    public async Task Landscape_Letter_page_size_survives_conversion_to_PDF()
    {
        new PageLayoutService().SetPageSize(_docxPath, PageSize.Letter(PageOrientation.Landscape));

        await new WordToPdfConverter(TestFiles.ConversionOptions()).ConvertAsync(_docxPath, _pdfPath);

        using var doc = PdfDocument.Open(_pdfPath);
        var page = doc.GetPage(1);

        // Letter is 8.5x11in = 612x792pt; landscape swaps that to 792x612.
        Assert.True(page.Width > page.Height, $"Expected landscape (width > height), got {page.Width}x{page.Height}pt.");
        Assert.Equal(792, page.Width, tolerance: 2);
        Assert.Equal(612, page.Height, tolerance: 2);
    }

    [Fact]
    public async Task Portrait_A4_page_size_survives_conversion_to_PDF()
    {
        new PageLayoutService().SetPageSize(_docxPath, PageSize.A4());

        await new WordToPdfConverter(TestFiles.ConversionOptions()).ConvertAsync(_docxPath, _pdfPath);

        using var doc = PdfDocument.Open(_pdfPath);
        var page = doc.GetPage(1);

        // A4 is 210x297mm = ~595x842pt.
        Assert.True(page.Height > page.Width, $"Expected portrait (height > width), got {page.Width}x{page.Height}pt.");
        Assert.Equal(595, page.Width, tolerance: 2);
        Assert.Equal(842, page.Height, tolerance: 2);
    }

    public void Dispose()
    {
        File.Delete(_docxPath);
        if (File.Exists(_pdfPath))
            File.Delete(_pdfPath);
    }
}
