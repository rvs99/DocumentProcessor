using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.Samples;
using DocumentProcessor.Core.Transplant;

namespace DocumentProcessor.Tests.Transplant;

public class ClauseTransplantServiceTests : IDisposable
{
    private readonly string _sourcePath = TestFiles.NewTempPath(".docx");
    private readonly string _targetPath = TestFiles.NewTempPath(".docx");
    private readonly string _outputPath = TestFiles.NewTempPath(".docx");
    private readonly ClauseTransplantService _sut = new();

    public ClauseTransplantServiceTests()
    {
        SampleDocumentFactory.CreateBasicDocument(_sourcePath, "Master Clause Library",
            ["Governing Law: This agreement is governed by the laws of Delaware.", "Some other unrelated clause."]);
        SampleDocumentFactory.CreateBasicDocument(_targetPath, "New Contract",
            ["Parties: Acme Corp and Widget LLC.", "Term: Twelve months."]);
    }

    [Fact]
    public void ListParagraphs_returns_index_and_text_for_every_top_level_paragraph()
    {
        var paragraphs = _sut.ListParagraphs(_sourcePath);

        // Title paragraph + 2 body paragraphs from SampleDocumentFactory.
        Assert.Equal(3, paragraphs.Count);
        Assert.Contains(paragraphs, p => p.Text.Contains("Governing Law"));
    }

    [Fact]
    public void TransplantParagraphs_inserts_the_source_clause_verbatim_at_the_target_position()
    {
        var governingLawIndex = _sut.ListParagraphs(_sourcePath).Single(p => p.Text.Contains("Governing Law")).Index;

        _sut.TransplantParagraphs(
            sourcePath: _sourcePath, sourceStartIndex: governingLawIndex, paragraphCount: 1,
            targetPath: _targetPath, insertBeforeParagraphIndex: 2,
            outputPath: _outputPath);

        var merged = _sut.ListParagraphs(_outputPath).Select(p => p.Text).ToList();

        Assert.Contains(merged, t => t.Contains("Parties: Acme Corp"));
        Assert.Contains(merged, t => t.Contains("Term: Twelve months"));
        Assert.Contains(merged, t => t.Contains("Governing Law: This agreement is governed by the laws of Delaware."));
        // The clause lands between the target's two paragraphs, not appended at the very end.
        Assert.True(merged.FindIndex(t => t.Contains("Governing Law")) > merged.FindIndex(t => t.Contains("Parties")));
    }

    [Fact]
    public void TransplantParagraphs_leaves_the_source_document_untouched()
    {
        var before = _sut.ListParagraphs(_sourcePath).Select(p => p.Text).ToList();

        _sut.TransplantParagraphs(_sourcePath, 0, 1, _targetPath, 0, _outputPath);

        var after = _sut.ListParagraphs(_sourcePath).Select(p => p.Text).ToList();
        Assert.Equal(before, after);
    }

    [Fact]
    public void TransplantParagraphs_rejects_an_out_of_range_insertion_point()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            _sut.TransplantParagraphs(_sourcePath, 0, 1, _targetPath, 999, _outputPath));
    }

    [Fact]
    public void RemoveParagraphs_removes_the_given_range_and_keeps_the_rest()
    {
        _sut.RemoveParagraphs(_targetPath, startIndex: 1, count: 1, _outputPath);

        var texts = _sut.ListParagraphs(_outputPath).Select(p => p.Text).ToList();
        Assert.DoesNotContain(texts, t => t.Contains("Parties: Acme Corp"));
        Assert.Contains(texts, t => t.Contains("Term: Twelve months"));
    }

    [Fact]
    public void ReplaceParagraphs_swaps_the_target_range_for_the_source_range()
    {
        var governingLawIndex = _sut.ListParagraphs(_sourcePath).Single(p => p.Text.Contains("Governing Law")).Index;

        _sut.ReplaceParagraphs(
            sourcePath: _sourcePath, sourceStartIndex: governingLawIndex, sourceParagraphCount: 1,
            targetPath: _targetPath, replacedStartIndex: 1, replacedCount: 1,
            outputPath: _outputPath);

        var texts = _sut.ListParagraphs(_outputPath).Select(p => p.Text).ToList();
        Assert.DoesNotContain(texts, t => t.Contains("Parties: Acme Corp"));
        Assert.Contains(texts, t => t.Contains("Governing Law"));
        Assert.Contains(texts, t => t.Contains("Term: Twelve months"));
    }

    [Fact]
    public void RemoveParagraphsWithCrossReferenceCleanup_reports_a_reference_that_becomes_dangling()
    {
        var path = TestFiles.NewTempPath(".docx");
        try
        {
            SampleDocumentFactory.CreateDocumentFromParagraphs(path,
            [
                new Paragraph(new BookmarkStart { Id = "1", Name = "_Ref100" }, new Run(new Text("Target clause.")), new BookmarkEnd { Id = "1" }),
                new Paragraph(new SimpleField(new Run(new Text("Target clause"))) { Instruction = " REF _Ref100 \\h " })
            ]);

            var warnings = _sut.RemoveParagraphsWithCrossReferenceCleanup(path, startIndex: 0, count: 1, path);

            Assert.Single(warnings);
            Assert.Equal("_Ref100", warnings[0].BookmarkName);
            // The removal itself still happened despite the warning.
            Assert.Single(_sut.ListParagraphs(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveParagraphsWithCrossReferenceCleanup_reports_nothing_when_no_reference_is_affected()
    {
        var warnings = _sut.RemoveParagraphsWithCrossReferenceCleanup(_targetPath, startIndex: 1, count: 1, _outputPath);

        Assert.Empty(warnings);
    }

    public void Dispose()
    {
        File.Delete(_sourcePath);
        File.Delete(_targetPath);
        if (File.Exists(_outputPath))
            File.Delete(_outputPath);
    }
}
