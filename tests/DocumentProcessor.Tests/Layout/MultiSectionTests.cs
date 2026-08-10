using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.Layout;
using DocumentProcessor.Core.Samples;
using OoxmlPageSize = DocumentFormat.OpenXml.Wordprocessing.PageSize;
using LayoutPageSize = DocumentProcessor.Core.Layout.PageSize;

namespace DocumentProcessor.Tests.Layout;

/// <summary>Covers the Phase 4 addition: splitting a document into multiple independently-laid-out sections.</summary>
public class MultiSectionTests : IDisposable
{
    private readonly string _path = TestFiles.NewTempPath(".docx");
    private readonly PageLayoutService _sut = new();

    public MultiSectionTests()
    {
        // 4 top-level paragraphs: Title, "One.", "Two.", "Three." (indices 0-3).
        SampleDocumentFactory.CreateBasicDocument(_path, "Multi-Section Test", ["One.", "Two.", "Three."]);
    }

    [Fact]
    public void InsertSectionBreak_produces_two_sections_where_there_was_one()
    {
        _sut.InsertSectionBreak(_path, beforeParagraphIndex: 2);

        Assert.Equal(2, GetAllSectionProperties().Count);
    }

    [Fact]
    public void InsertSectionBreak_attaches_the_new_sectPr_to_the_paragraph_immediately_before_the_split()
    {
        _sut.InsertSectionBreak(_path, beforeParagraphIndex: 2);

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var paragraphs = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().ToList();

        // paragraphs[1] is "One." (index 1: Title=0, "One."=1) -> the paragraph right before index 2 ("Two.").
        Assert.NotNull(paragraphs[1].ParagraphProperties?.SectionProperties);
        Assert.Equal("One.", paragraphs[1].InnerText);
    }

    [Fact]
    public void InsertSectionBreak_at_index_zero_inserts_an_anchor_paragraph_rather_than_failing()
    {
        _sut.InsertSectionBreak(_path, beforeParagraphIndex: 0);

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var firstParagraph = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();

        Assert.NotNull(firstParagraph.ParagraphProperties?.SectionProperties);
        Assert.Equal("", firstParagraph.InnerText);
    }

    [Fact]
    public void InsertSectionBreak_rejects_an_out_of_range_index()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => _sut.InsertSectionBreak(_path, beforeParagraphIndex: 999));
    }

    [Fact]
    public void The_new_section_starts_as_a_copy_of_the_documents_existing_page_setup()
    {
        _sut.SetPageSize(_path, LayoutPageSize.Legal());
        _sut.InsertSectionBreak(_path, beforeParagraphIndex: 2);

        var sections = GetAllSectionProperties();
        var firstSectionSize = sections[0].Elements<OoxmlPageSize>().Single();
        var secondSectionSize = sections[1].Elements<OoxmlPageSize>().Single();

        Assert.Equal(firstSectionSize.Width!.Value, secondSectionSize.Width!.Value);
        Assert.Equal(firstSectionSize.Height!.Value, secondSectionSize.Height!.Value);
    }

    [Fact]
    public void SetPageSize_with_a_sectionIndex_after_a_split_only_changes_that_section()
    {
        _sut.InsertSectionBreak(_path, beforeParagraphIndex: 2);

        _sut.SetPageSize(_path, LayoutPageSize.Letter(PageOrientation.Landscape), sectionIndex: 0);
        _sut.SetPageSize(_path, LayoutPageSize.A4(), sectionIndex: 1);

        var sections = GetAllSectionProperties();
        var first = sections[0].Elements<OoxmlPageSize>().Single();
        var second = sections[1].Elements<OoxmlPageSize>().Single();

        Assert.Equal(PageOrientationValues.Landscape, first.Orient!.Value);
        Assert.Equal(11906u, second.Width!.Value); // A4 width, unaffected by section 0's change
    }

    [Fact]
    public void InsertSectionBreak_called_twice_produces_three_sections()
    {
        // Neither call is at index 0, so no anchor paragraph gets inserted and the top-level
        // paragraph count/indices stay stable across both calls.
        _sut.InsertSectionBreak(_path, beforeParagraphIndex: 1);
        _sut.InsertSectionBreak(_path, beforeParagraphIndex: 3);

        Assert.Equal(3, GetAllSectionProperties().Count);
    }

    private List<SectionProperties> GetAllSectionProperties()
    {
        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        return doc.MainDocumentPart!.Document!.Body!.Descendants<SectionProperties>().ToList();
    }

    public void Dispose() => File.Delete(_path);
}
