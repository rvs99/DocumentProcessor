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

    public void Dispose()
    {
        File.Delete(_sourcePath);
        File.Delete(_targetPath);
        if (File.Exists(_outputPath))
            File.Delete(_outputPath);
    }
}
