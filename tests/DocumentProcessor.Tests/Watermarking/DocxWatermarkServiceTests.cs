using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.Samples;
using DocumentProcessor.Core.Watermarking;

namespace DocumentProcessor.Tests.Watermarking;

public class DocxWatermarkServiceTests : IDisposable
{
    private readonly string _path = TestFiles.NewTempPath(".docx");
    private readonly DocxWatermarkService _sut = new();

    public DocxWatermarkServiceTests()
    {
        SampleDocumentFactory.CreateBasicDocument(_path, "Confidential Report", ["Body text."]);
    }

    [Fact]
    public void AddTextWatermark_creates_a_header_containing_the_watermark_text()
    {
        _sut.AddTextWatermark(_path, "DRAFT");

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var headerPart = doc.MainDocumentPart!.HeaderParts.Single();
        // The watermark shapetype boilerplate also contains a (contentless) TextPath, so find the
        // one that actually carries the watermark's text rather than assuming there's only one.
        var textPath = headerPart.Header!.Descendants<DocumentFormat.OpenXml.Vml.TextPath>().Single(t => t.String is not null);

        Assert.Equal("DRAFT", textPath.String?.Value);
    }

    [Fact]
    public void AddTextWatermark_wires_a_default_header_reference_into_the_section()
    {
        _sut.AddTextWatermark(_path, "DRAFT");

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var sectPr = doc.MainDocumentPart!.Document!.Body!.Elements<SectionProperties>().Single();
        var headerRef = sectPr.Elements<HeaderReference>().Single();

        Assert.Equal(HeaderFooterValues.Default, headerRef.Type!.Value);
    }

    [Fact]
    public void AddTextWatermark_called_twice_replaces_rather_than_duplicates_the_header_reference()
    {
        _sut.AddTextWatermark(_path, "DRAFT");
        _sut.AddTextWatermark(_path, "CONFIDENTIAL");

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var sectPr = doc.MainDocumentPart!.Document!.Body!.Elements<SectionProperties>().Single();

        Assert.Single(sectPr.Elements<HeaderReference>());
        var headerPart = (HeaderPart)doc.MainDocumentPart!.GetPartById(sectPr.Elements<HeaderReference>().Single().Id!);
        var textPath = headerPart.Header!.Descendants<DocumentFormat.OpenXml.Vml.TextPath>().Single(t => t.String is not null);
        Assert.Equal("CONFIDENTIAL", textPath.String?.Value);
    }

    public void Dispose() => File.Delete(_path);
}
