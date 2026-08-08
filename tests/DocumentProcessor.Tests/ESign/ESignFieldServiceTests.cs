using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.ESign;
using DocumentProcessor.Core.Samples;
using UglyToad.PdfPig;

namespace DocumentProcessor.Tests.ESign;

public class ESignFieldServiceTests : IDisposable
{
    private readonly string _docxPath = TestFiles.NewTempPath(".docx");
    private readonly string _pdfPath = TestFiles.NewTestPdf("ESign Test", ["Please review and sign below."]);
    private readonly string _pdfOutputPath = TestFiles.NewTempPath(".pdf");
    private readonly ESignFieldService _sut = new();

    public ESignFieldServiceTests()
    {
        SampleDocumentFactory.CreateBasicDocument(_docxPath, "Agreement", ["Signature required below."]);
    }

    [Fact]
    public void InjectDocxAnchor_adds_a_tagged_content_control_containing_the_anchor_text()
    {
        _sut.InjectDocxAnchor(_docxPath, "/sig1/", tag: "ESignatureField");

        using var doc = WordprocessingDocument.Open(_docxPath, isEditable: false);
        var sdt = doc.MainDocumentPart!.Document!.Descendants<SdtElement>().Single();

        Assert.Equal("ESignatureField", sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value);
        Assert.Contains("/sig1/", sdt.InnerText);
    }

    [Fact]
    public void InjectPdfAnchor_stamps_visible_anchor_text_that_can_be_extracted()
    {
        _sut.InjectPdfAnchor(_pdfPath, _pdfOutputPath, "/sig1/", pageIndex: 0, x: 100, y: 700, invisible: false);

        using var doc = PdfDocument.Open(_pdfOutputPath);
        Assert.Contains("/sig1/", doc.GetPage(1).Text);
    }

    [Fact]
    public void InjectPdfAnchor_rejects_an_out_of_range_page_index()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            _sut.InjectPdfAnchor(_pdfPath, _pdfOutputPath, "/sig1/", pageIndex: 99, x: 0, y: 0));
    }

    public void Dispose()
    {
        File.Delete(_docxPath);
        File.Delete(_pdfPath);
        if (File.Exists(_pdfOutputPath))
            File.Delete(_pdfOutputPath);
    }
}
