using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.ContentControls;
using DocumentProcessor.Core.Samples;

namespace DocumentProcessor.Tests.ContentControls;

public class ContentControlTypeAwareTests : IDisposable
{
    private readonly string _path = TestFiles.NewTempPath(".docx");
    private readonly ContentControlService _sut = new();

    private static Paragraph DateControl(string tag, string placeholder = "[Date]") => new(
        new SdtRun(
            new SdtProperties(new Tag { Val = tag }, new SdtContentDate(new DateFormat { Val = "yyyy-MM-dd" })),
            new SdtContentRun(new Run(new Text(placeholder)))));

    private static Paragraph DropDownControl(string tag, params (string Display, string Value)[] items)
    {
        var dropDown = new SdtContentDropDownList();
        foreach (var (display, value) in items)
            dropDown.AppendChild(new ListItem { DisplayText = display, Value = value });

        return new Paragraph(
            new SdtRun(
                new SdtProperties(new Tag { Val = tag }, dropDown),
                new SdtContentRun(new Run(new Text("Choose an item.")))));
    }

    private static SdtBlock RichTextControl(string tag, string placeholder = "placeholder") => new(
        new SdtProperties(new Tag { Val = tag }, new SdtContentRichText()),
        new SdtContentBlock(new Paragraph(new Run(new Text(placeholder)))));

    [Fact]
    public void ReplaceByTag_on_date_picker_sets_display_text_and_full_date_metadata()
    {
        CreateFixture(DateControl("EffectiveDate"));

        _sut.ReplaceByTag(_path, "EffectiveDate", "2026-08-22");

        var control = _sut.ListContentControls(_path).Single(c => c.Tag == "EffectiveDate");
        Assert.Equal("2026-08-22", control.Text);

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var sdt = doc.MainDocumentPart!.Document!.Body!.Descendants<SdtRun>().Single();
        var fullDate = sdt.SdtProperties!.GetFirstChild<SdtContentDate>()!.FullDate!.Value;
        Assert.Equal(new DateTime(2026, 8, 22), fullDate);
    }

    [Fact]
    public void SetContentDateByTag_formats_using_the_supplied_format()
    {
        CreateFixture(DateControl("EffectiveDate"));

        _sut.SetContentDateByTag(_path, "EffectiveDate", new DateTime(2026, 8, 22), "MMMM d, yyyy");

        var control = _sut.ListContentControls(_path).Single(c => c.Tag == "EffectiveDate");
        Assert.Equal("August 22, 2026", control.Text);
    }

    [Fact]
    public void SetContentDropDownSelectionByTag_updates_last_value_and_display_text()
    {
        CreateFixture(DropDownControl("State", ("California", "CA"), ("New York", "NY")));

        _sut.SetContentDropDownSelectionByTag(_path, "State", "NY");

        var control = _sut.ListContentControls(_path).Single(c => c.Tag == "State");
        Assert.Equal("New York", control.Text);

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var sdt = doc.MainDocumentPart!.Document!.Body!.Descendants<SdtRun>().Single();
        var lastValue = sdt.SdtProperties!.GetFirstChild<SdtContentDropDownList>()!.LastValue!.Value;
        Assert.Equal("NY", lastValue);
    }

    [Fact]
    public void SetContentDropDownSelectionByTag_rejects_a_value_not_in_the_list()
    {
        CreateFixture(DropDownControl("State", ("California", "CA"), ("New York", "NY")));

        Assert.Throws<ArgumentException>(() => _sut.SetContentDropDownSelectionByTag(_path, "State", "TX"));
    }

    [Fact]
    public void ReplaceByTag_on_dropdown_best_effort_sets_last_value_without_validation()
    {
        CreateFixture(DropDownControl("State", ("California", "CA")));

        _sut.ReplaceByTag(_path, "State", "Texas");

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var sdt = doc.MainDocumentPart!.Document!.Body!.Descendants<SdtRun>().Single();
        Assert.Equal("Texas", sdt.SdtProperties!.GetFirstChild<SdtContentDropDownList>()!.LastValue!.Value);
    }

    [Fact]
    public void SetContentRichTextByTag_injects_multi_paragraph_html_into_a_richtext_control()
    {
        CreateFixture(RichTextControl("Description"));

        _sut.SetContentRichTextByTag(_path, "Description", "<p>First.</p><p>Second <b>bold</b>.</p>");

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var sdtBlock = doc.MainDocumentPart!.Document!.Body!.Descendants<SdtBlock>().Single();
        var paragraphs = sdtBlock.SdtContentBlock!.Elements<Paragraph>().ToList();
        Assert.Equal(2, paragraphs.Count);
        Assert.Equal("First.", paragraphs[0].InnerText);
        Assert.Equal("Second bold.", paragraphs[1].InnerText);
    }

    [Theory]
    [InlineData(ContentControlLockMode.SdtLocked)]
    [InlineData(ContentControlLockMode.ContentLocked)]
    [InlineData(ContentControlLockMode.SdtContentLocked)]
    public void SetLock_then_ListContentControls_reports_the_lock_mode(ContentControlLockMode mode)
    {
        CreateFixture(DateControl("EffectiveDate"));

        _sut.SetLock(_path, "EffectiveDate", mode);

        var control = _sut.ListContentControls(_path).Single(c => c.Tag == "EffectiveDate");
        Assert.Equal(mode, control.LockMode);
    }

    [Fact]
    public void SetLock_Unlocked_removes_a_previously_set_lock()
    {
        CreateFixture(DateControl("EffectiveDate"));
        _sut.SetLock(_path, "EffectiveDate", ContentControlLockMode.SdtContentLocked);

        _sut.SetLock(_path, "EffectiveDate", ContentControlLockMode.Unlocked);

        var control = _sut.ListContentControls(_path).Single(c => c.Tag == "EffectiveDate");
        Assert.Null(control.LockMode);
    }

    private void CreateFixture(params OpenXmlElement[] elements)
    {
        using var doc = WordprocessingDocument.Create(_path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        var body = mainPart.Document.Body!;
        foreach (var element in elements)
            body.AppendChild(element);
        body.AppendChild(new SectionProperties());
        SampleDocumentFactory.AddMinimalStyles(mainPart);
        mainPart.Document.Save();
    }

    public void Dispose() => File.Delete(_path);
}
