using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.Layout;
using DocumentProcessor.Core.Samples;

namespace DocumentProcessor.Tests.Layout;

public class HeaderFooterServiceTests : IDisposable
{
    private readonly string _path = TestFiles.NewTempPath(".docx");
    private readonly HeaderFooterService _sut = new();

    public HeaderFooterServiceTests()
    {
        SampleDocumentFactory.CreateBasicDocument(_path, "Header/Footer Test", ["Body text."]);
    }

    [Fact]
    public void SetHeaderText_wires_a_default_header_reference_containing_the_text()
    {
        _sut.SetHeaderText(_path, "Acme Corporation");

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var sectPr = doc.MainDocumentPart!.Document!.Body!.Elements<SectionProperties>().Single();
        var headerRef = sectPr.Elements<HeaderReference>().Single(r => r.Type!.Value == HeaderFooterValues.Default);
        var headerPart = (HeaderPart)doc.MainDocumentPart!.GetPartById(headerRef.Id!);

        Assert.Contains("Acme Corporation", headerPart.Header!.InnerText);
    }

    [Fact]
    public void SetFooterText_wires_a_default_footer_reference_containing_the_text()
    {
        _sut.SetFooterText(_path, "Confidential");

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var sectPr = doc.MainDocumentPart!.Document!.Body!.Elements<SectionProperties>().Single();
        var footerRef = sectPr.Elements<FooterReference>().Single(r => r.Type!.Value == HeaderFooterValues.Default);
        var footerPart = (FooterPart)doc.MainDocumentPart!.GetPartById(footerRef.Id!);

        Assert.Contains("Confidential", footerPart.Footer!.InnerText);
    }

    [Fact]
    public void Setting_a_footer_does_not_disturb_an_existing_header_and_vice_versa()
    {
        _sut.SetHeaderText(_path, "Header text");
        _sut.SetFooterText(_path, "Footer text");

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var sectPr = doc.MainDocumentPart!.Document!.Body!.Elements<SectionProperties>().Single();

        var headerPart = (HeaderPart)doc.MainDocumentPart!.GetPartById(sectPr.Elements<HeaderReference>().Single().Id!);
        var footerPart = (FooterPart)doc.MainDocumentPart!.GetPartById(sectPr.Elements<FooterReference>().Single().Id!);

        Assert.Contains("Header text", headerPart.Header!.InnerText);
        Assert.Contains("Footer text", footerPart.Footer!.InnerText);
    }

    [Fact]
    public void SetHeaderText_called_twice_for_the_same_type_replaces_rather_than_duplicates()
    {
        _sut.SetHeaderText(_path, "First");
        _sut.SetHeaderText(_path, "Second");

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var sectPr = doc.MainDocumentPart!.Document!.Body!.Elements<SectionProperties>().Single();

        Assert.Single(sectPr.Elements<HeaderReference>());
        var headerPart = (HeaderPart)doc.MainDocumentPart!.GetPartById(sectPr.Elements<HeaderReference>().Single().Id!);
        Assert.Contains("Second", headerPart.Header!.InnerText);
        Assert.DoesNotContain("First", headerPart.Header!.InnerText);
    }

    [Fact]
    public void SetHeaderText_with_First_type_also_turns_on_titlePg_so_word_actually_shows_it()
    {
        _sut.SetHeaderText(_path, "First page only", HeaderFooterValues.First);

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var sectPr = doc.MainDocumentPart!.Document!.Body!.Elements<SectionProperties>().Single();

        Assert.True(sectPr.Elements<TitlePage>().Any());
        Assert.Single(sectPr.Elements<HeaderReference>(), r => r.Type!.Value == HeaderFooterValues.First);
    }

    [Fact]
    public void SetHeaderText_with_Even_type_also_turns_on_evenAndOddHeaders_setting()
    {
        _sut.SetHeaderText(_path, "Even pages", HeaderFooterValues.Even);

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var settings = doc.MainDocumentPart!.DocumentSettingsPart!.Settings!;

        Assert.True(settings.Elements<EvenAndOddHeaders>().Any());
    }

    [Fact]
    public void RemoveHeader_removes_the_reference_and_reports_it_removed_something()
    {
        _sut.SetHeaderText(_path, "Some header");

        var removed = _sut.RemoveHeader(_path);

        Assert.True(removed);
        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var sectPr = doc.MainDocumentPart!.Document!.Body!.Elements<SectionProperties>().Single();
        Assert.Empty(sectPr.Elements<HeaderReference>());
    }

    [Fact]
    public void RemoveFooter_on_a_document_with_no_footer_reports_nothing_removed_and_does_not_throw()
    {
        var removed = _sut.RemoveFooter(_path);

        Assert.False(removed);
    }

    [Fact]
    public void SetHeaderContent_lays_out_text_and_fields_in_order()
    {
        _sut.SetHeaderContent(_path,
        [
            new TextPart("Acme Corp — Page "),
            new FieldPart(HeaderFooterFieldType.PageNumber),
            new TextPart(" of "),
            new FieldPart(HeaderFooterFieldType.TotalPages)
        ]);

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var sectPr = doc.MainDocumentPart!.Document!.Body!.Elements<SectionProperties>().Single();
        var headerPart = (HeaderPart)doc.MainDocumentPart!.GetPartById(sectPr.Elements<HeaderReference>().Single().Id!);

        Assert.Contains("Acme Corp", headerPart.Header!.InnerText);
        var fields = headerPart.Header!.Descendants<SimpleField>().ToList();
        Assert.Equal(2, fields.Count);
        Assert.Equal("PAGE", fields[0].Instruction!.Value!.Trim());
        Assert.Equal("NUMPAGES", fields[1].Instruction!.Value!.Trim());
    }

    [Fact]
    public void SetFooterContent_resolves_tokens_from_the_supplied_data()
    {
        _sut.SetFooterContent(_path, [new TextPart("Prepared for {{Client.Name}}")],
            new Dictionary<string, object?> { ["Client"] = new Dictionary<string, object?> { ["Name"] = "Acme Corp" } });

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var sectPr = doc.MainDocumentPart!.Document!.Body!.Elements<SectionProperties>().Single();
        var footerPart = (FooterPart)doc.MainDocumentPart!.GetPartById(sectPr.Elements<FooterReference>().Single().Id!);

        Assert.Contains("Prepared for Acme Corp", footerPart.Footer!.InnerText);
    }

    [Fact]
    public void SetHeaderContent_with_logo_embeds_an_image_part_referenced_by_a_drawing()
    {
        // A minimal valid 1x1 PNG, just enough to be a real image file the ImagePart accepts.
        var pngBytes = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

        _sut.SetHeaderContent(_path, [new LogoPart(pngBytes, 100, 40)]);

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var sectPr = doc.MainDocumentPart!.Document!.Body!.Elements<SectionProperties>().Single();
        var headerPart = (HeaderPart)doc.MainDocumentPart!.GetPartById(sectPr.Elements<HeaderReference>().Single().Id!);

        Assert.Single(headerPart.ImageParts);
        Assert.Single(headerPart.Header!.Descendants<Drawing>());
    }

    public void Dispose() => File.Delete(_path);
}
