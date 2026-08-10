using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.Layout;
using DocumentProcessor.Core.Samples;
using OoxmlPageSize = DocumentFormat.OpenXml.Wordprocessing.PageSize;
using LayoutPageSize = DocumentProcessor.Core.Layout.PageSize;

namespace DocumentProcessor.Tests.Layout;

public class PageLayoutServiceTests : IDisposable
{
    private readonly string _path = TestFiles.NewTempPath(".docx");
    private readonly PageLayoutService _sut = new();

    public PageLayoutServiceTests()
    {
        SampleDocumentFactory.CreateBasicDocument(_path, "Layout Test", ["Paragraph one.", "Paragraph two."]);
    }

    [Fact]
    public void SetPageSize_writes_the_requested_dimensions()
    {
        _sut.SetPageSize(_path, new LayoutPageSize(11906, 16838));

        var pgSz = GetSectionProperties().Elements<OoxmlPageSize>().Single();
        Assert.Equal(11906u, pgSz.Width!.Value);
        Assert.Equal(16838u, pgSz.Height!.Value);
    }

    [Fact]
    public void SetPageSize_landscape_swaps_width_and_height_and_sets_the_orient_flag()
    {
        _sut.SetPageSize(_path, LayoutPageSize.Letter(PageOrientation.Landscape));

        var pgSz = GetSectionProperties().Elements<OoxmlPageSize>().Single();
        Assert.Equal(15840u, pgSz.Width!.Value);
        Assert.Equal(12240u, pgSz.Height!.Value);
        Assert.Equal(PageOrientationValues.Landscape, pgSz.Orient!.Value);
    }

    [Fact]
    public void SetPageSize_called_twice_replaces_rather_than_duplicates_the_element()
    {
        _sut.SetPageSize(_path, LayoutPageSize.Letter());
        _sut.SetPageSize(_path, LayoutPageSize.A4());

        Assert.Single(GetSectionProperties().Elements<OoxmlPageSize>());
    }

    [Fact]
    public void SetMargins_writes_the_requested_values()
    {
        _sut.SetMargins(_path, PageMargins.FromInches(top: 1.5, bottom: 1.5, left: 0.5, right: 0.5));

        var margin = GetSectionProperties().Elements<PageMargin>().Single();
        Assert.Equal(2160, margin.Top!.Value);
        Assert.Equal(2160, margin.Bottom!.Value);
        Assert.Equal(720u, margin.Left!.Value);
        Assert.Equal(720u, margin.Right!.Value);
    }

    [Fact]
    public void SetColumns_writes_the_requested_count_and_spacing()
    {
        _sut.SetColumns(_path, columnCount: 3, spacingTwips: 360);

        var cols = GetSectionProperties().Elements<Columns>().Single();
        Assert.Equal(3, cols.ColumnCount!.Value);
        Assert.Equal("360", cols.Space!.Value);
    }

    [Fact]
    public void SetColumns_rejects_a_count_below_one()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => _sut.SetColumns(_path, columnCount: 0));
    }

    [Fact]
    public void InsertPageBreak_adds_a_break_in_a_new_paragraph_before_the_target_index()
    {
        _sut.InsertPageBreak(_path, beforeParagraphIndex: 1);

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var paragraphs = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().ToList();

        // original paragraphs were [Title, "Paragraph one.", "Paragraph two."] -> break goes before index 1
        Assert.Equal(4, paragraphs.Count);
        Assert.NotNull(paragraphs[1].Descendants<Break>().SingleOrDefault(b => b.Type?.Value == BreakValues.Page));
    }

    [Fact]
    public void InsertPageBreak_rejects_an_out_of_range_index()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => _sut.InsertPageBreak(_path, beforeParagraphIndex: 999));
    }

    [Fact]
    public void SetDefaultParagraphSpacing_updates_both_docDefaults_and_the_Normal_style()
    {
        _sut.SetDefaultParagraphSpacing(_path, afterTwips: 240, lineTwips: 480, LineSpacingRuleValues.Auto);

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var styles = doc.MainDocumentPart!.StyleDefinitionsPart!.Styles!;

        var defaultSpacing = styles.Elements<DocDefaults>().Single()
            .ParagraphPropertiesDefault!.ParagraphPropertiesBaseStyle!.Elements<SpacingBetweenLines>().Single();
        Assert.Equal("240", defaultSpacing.After!.Value);
        Assert.Equal("480", defaultSpacing.Line!.Value);

        var normalStyle = styles.Elements<Style>().Single(s => s.StyleId!.Value == "Normal");
        var normalSpacing = normalStyle.StyleParagraphProperties!.Elements<SpacingBetweenLines>().Single();
        Assert.Equal("240", normalSpacing.After!.Value);
    }

    [Fact]
    public void SetDefaultParagraphSpacing_called_twice_replaces_rather_than_duplicates()
    {
        _sut.SetDefaultParagraphSpacing(_path, afterTwips: 100, lineTwips: 200, LineSpacingRuleValues.Auto);
        _sut.SetDefaultParagraphSpacing(_path, afterTwips: 300, lineTwips: 400, LineSpacingRuleValues.Auto);

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var styles = doc.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        var normalStyle = styles.Elements<Style>().Single(s => s.StyleId!.Value == "Normal");

        Assert.Single(normalStyle.StyleParagraphProperties!.Elements<SpacingBetweenLines>());
        Assert.Equal("300", normalStyle.StyleParagraphProperties!.Elements<SpacingBetweenLines>().Single().After!.Value);
    }

    private SectionProperties GetSectionProperties()
    {
        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        return doc.MainDocumentPart!.Document!.Body!.Elements<SectionProperties>().Single();
    }

    public void Dispose() => File.Delete(_path);
}
