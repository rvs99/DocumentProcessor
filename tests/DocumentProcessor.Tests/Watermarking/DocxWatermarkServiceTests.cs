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
        var textBox = headerPart.Header!.Descendants<DocumentFormat.OpenXml.Vml.TextBox>().Single();

        Assert.Equal("DRAFT", textBox.InnerText);
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
        var textBox = headerPart.Header!.Descendants<DocumentFormat.OpenXml.Vml.TextBox>().Single();
        Assert.Equal("CONFIDENTIAL", textBox.InnerText);
    }

    [Fact]
    public void AddTextWatermark_defaults_to_removable_matching_words_own_naming_convention()
    {
        // Word's Design -> Watermark UI (both "Remove Watermark" and the predefined gallery's
        // replace-existing behavior) identifies a watermark purely by this id prefix on the shape,
        // not by its appearance or position. Without it, "Remove Watermark" silently does nothing,
        // and picking a predefined watermark adds a second shape instead of replacing this one.
        _sut.AddTextWatermark(_path, "DRAFT");

        var shape = GetWatermarkShape(_path);
        Assert.StartsWith("PowerPlusWaterMarkObject", shape.Id?.Value);
    }

    [Fact]
    public void AddTextWatermark_with_removable_false_uses_an_id_word_does_not_recognize_as_a_watermark()
    {
        _sut.AddTextWatermark(_path, "CONFIDENTIAL", removable: false);

        var shape = GetWatermarkShape(_path);
        Assert.DoesNotContain("PowerPlusWaterMarkObject", shape.Id?.Value);
    }

    [Fact]
    public void RemoveWatermark_removes_a_removable_watermark_and_reports_it_removed_something()
    {
        _sut.AddTextWatermark(_path, "DRAFT", removable: true);

        var removed = _sut.RemoveWatermark(_path);

        Assert.True(removed);
        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var headerPart = doc.MainDocumentPart!.HeaderParts.Single();
        Assert.Empty(headerPart.Header!.Descendants<DocumentFormat.OpenXml.Vml.Shape>());
    }

    [Fact]
    public void RemoveWatermark_also_removes_a_locked_watermark()
    {
        // "Irrespective of type" — RemoveWatermark shouldn't just handle the Word-recognized
        // (removable) case; a locked watermark's whole point is that Word's own UI can't clear it,
        // not that our own code can't.
        _sut.AddTextWatermark(_path, "CONFIDENTIAL", removable: false);

        var removed = _sut.RemoveWatermark(_path);

        Assert.True(removed);
        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var headerPart = doc.MainDocumentPart!.HeaderParts.Single();
        Assert.Empty(headerPart.Header!.Descendants<DocumentFormat.OpenXml.Vml.Shape>());
    }

    [Fact]
    public void RemoveWatermark_on_a_document_with_no_watermark_reports_nothing_removed_and_does_not_throw()
    {
        var removed = _sut.RemoveWatermark(_path);

        Assert.False(removed);
    }

    private static DocumentFormat.OpenXml.Vml.Shape GetWatermarkShape(string path)
    {
        using var doc = WordprocessingDocument.Open(path, isEditable: false);
        var headerPart = doc.MainDocumentPart!.HeaderParts.Single();
        return headerPart.Header!.Descendants<DocumentFormat.OpenXml.Vml.Shape>().Single();
    }

    public void Dispose() => File.Delete(_path);
}
