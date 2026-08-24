using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.Extraction;
using DocumentProcessor.Core.Layout;
using DocumentProcessor.Core.Samples;
using DocumentProcessor.Core.Sessions;
using DocumentProcessor.Core.Tables;

namespace DocumentProcessor.Tests.Extraction;

public class TextExtractionServiceTests : IDisposable
{
    private readonly List<string> _cleanup = [];
    private readonly TextExtractionService _service = new();

    private string NewPath()
    {
        var path = TestFiles.NewTempPath(".docx");
        _cleanup.Add(path);
        return path;
    }

    private string NewContract()
    {
        var path = NewPath();
        SampleDocumentFactory.CreateBasicDocument(path, "Services Agreement",
        [
            "This Agreement is entered into by the parties.",
            "Client shall pay $150,000 annually.",
        ]);
        return path;
    }

    /// <summary>A document with real Heading1/Heading2 structure, which the basic factory doesn't produce.</summary>
    private string NewStructuredContract()
    {
        var path = NewPath();
        SampleDocumentFactory.CreateDocumentFromParagraphs(path,
        [
            Styled("Liability", "Heading1"),
            new Paragraph(new Run(new Text("Liability is capped at fees paid."))),
            Styled("Exclusions", "Heading2"),
            new Paragraph(new Run(new Text("Gross negligence is excluded from the cap."))),
        ]);
        return path;
    }

    private static Paragraph Styled(string text, string styleId) => new(
        new ParagraphProperties(new ParagraphStyleId { Val = styleId }),
        new Run(new Text(text)));

    [Fact]
    public void Body_text_is_extracted_in_document_order()
    {
        var text = _service.ExtractText(NewContract());

        Assert.Equal(
            """
            Services Agreement
            This Agreement is entered into by the parties.
            Client shall pay $150,000 annually.
            """.ReplaceLineEndings("\n"),
            text);
    }

    [Fact]
    public void Empty_paragraphs_are_skipped_by_default_and_kept_on_request()
    {
        var path = NewPath();
        SampleDocumentFactory.CreateDocumentFromParagraphs(path,
        [
            new Paragraph(new Run(new Text("First."))),
            new Paragraph(),
            new Paragraph(new Run(new Text("Second."))),
        ]);

        Assert.Equal(2, _service.ExtractBlocks(path).Count);
        Assert.Equal(3, _service.ExtractBlocks(path, new TextExtractionOptions { SkipEmpty = false }).Count);
    }

    [Fact]
    public void Headings_are_identified_and_carried_onto_the_blocks_beneath_them()
    {
        var blocks = _service.ExtractBlocks(NewStructuredContract());

        var heading = blocks.Single(b => b.Text == "Liability");
        Assert.Equal(1, heading.HeadingLevel);
        Assert.Equal("Heading1", heading.StyleId);
        Assert.Null(heading.Heading);        // a heading isn't filed under itself

        var body = blocks.Single(b => b.Text.StartsWith("Liability is capped"));
        Assert.Null(body.HeadingLevel);
        Assert.Equal("Liability", body.Heading);

        // The nearest preceding heading wins, whatever its level.
        Assert.Equal("Exclusions", blocks.Single(b => b.Text.StartsWith("Gross negligence")).Heading);
        Assert.Equal(2, blocks.Single(b => b.Text == "Exclusions").HeadingLevel);
    }

    [Fact]
    public void Table_text_is_included_in_place_and_flagged()
    {
        var path = NewContract();
        new TableGenerationService().AppendTable(path,
            new TableSpec(["Item", "Amount"], [["Implementation", "$150,000"]]));

        var blocks = _service.ExtractBlocks(path);

        Assert.Contains(blocks, b => b.Text == "Implementation" && b.IsTableContent);
        Assert.DoesNotContain(blocks.Where(b => !b.IsTableContent), b => b.Text == "$150,000");

        // Excluding tables drops exactly those blocks and nothing else.
        var withoutTables = _service.ExtractBlocks(path, new TextExtractionOptions { IncludeTables = false });
        Assert.DoesNotContain(withoutTables, b => b.IsTableContent);
        Assert.Equal(blocks.Count(b => !b.IsTableContent), withoutTables.Count);
    }

    [Fact]
    public void Deleted_tracked_text_is_excluded_by_default()
    {
        var path = NewPath();
        SampleDocumentFactory.CreateDocumentWithTrackedChanges(path, author: "Counsel");

        // InnerText would concatenate both the old and the new wording — text the document never
        // actually reads as. This is the case naive extraction gets wrong.
        var text = _service.ExtractText(path);
        Assert.Contains("twenty-four", text);
        Assert.DoesNotContain("twelve", text);

        Assert.Contains("twelve", _service.ExtractText(path, new TextExtractionOptions { IncludeDeletedText = true }));
    }

    [Fact]
    public void Header_and_footer_text_is_excluded_by_default()
    {
        var path = NewContract();
        new HeaderFooterService().SetHeaderText(path, "CONFIDENTIAL — Acme Corporation");

        Assert.DoesNotContain("CONFIDENTIAL", _service.ExtractText(path));
        Assert.Contains("CONFIDENTIAL", _service.ExtractText(path,
            new TextExtractionOptions { IncludeHeadersAndFooters = true }));
    }

    [Fact]
    public void Blocks_are_indexed_contiguously_from_zero()
    {
        var blocks = _service.ExtractBlocks(NewStructuredContract());
        Assert.Equal(Enumerable.Range(0, blocks.Count), blocks.Select(b => b.Index));
    }

    [Fact]
    public void Text_can_be_extracted_from_an_open_session_after_edits()
    {
        using var session = DocumentSession.Open(File.ReadAllBytes(NewContract()));

        session.Tables.AppendTable(new TableSpec(["Item"], [["Implementation"]]));

        // Reading through the session sees the in-memory edit, without a save/reopen cycle.
        var text = session.Text.Extract();
        Assert.Contains("Services Agreement", text);
        Assert.Contains("Implementation", text);
    }

    public void Dispose()
    {
        foreach (var path in _cleanup)
            if (File.Exists(path)) File.Delete(path);
    }
}
