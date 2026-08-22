using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.Redlining;
using DocumentProcessor.Core.Samples;
using DocumentProcessor.Tests.Diagnostics;
using Microsoft.Extensions.Logging;

namespace DocumentProcessor.Tests.Redlining;

public class DocumentComparisonServiceTests : IDisposable
{
    private readonly string _originalPath = TestFiles.NewTempPath(".docx");
    private readonly string _revisedPath = TestFiles.NewTempPath(".docx");
    private readonly string _outputPath = TestFiles.NewTempPath(".docx");
    private readonly DocumentComparisonService _sut = new();

    public DocumentComparisonServiceTests()
    {
        SampleDocumentFactory.CreateBasicDocument(_originalPath, "Agreement",
            ["The term of this agreement is twelve months.", "This clause is unchanged."]);
        SampleDocumentFactory.CreateBasicDocument(_revisedPath, "Agreement",
            ["The term of this agreement is twenty-four months.", "This clause is unchanged.", "This clause is new."]);
    }

    [Fact]
    public void Compare_writes_a_redlined_document_containing_tracked_changes()
    {
        _sut.Compare(_originalPath, _revisedPath, _outputPath);

        Assert.True(File.Exists(_outputPath));
        using var doc = WordprocessingDocument.Open(_outputPath, isEditable: false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        Assert.NotEmpty(body.Descendants<InsertedRun>());
    }

    [Fact]
    public void Compare_summary_reports_the_insertion_of_a_new_clause()
    {
        var summary = _sut.Compare(_originalPath, _revisedPath, _outputPath);

        Assert.True(summary.InsertedCount > 0);
        Assert.Contains(summary.InsertedText, t => t.Contains("new", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Compare_logs_the_insertion_and_deletion_counts()
    {
        var logger = new CapturingLogger<DocumentComparisonService>();
        var sut = new DocumentComparisonService(logger);

        sut.Compare(_originalPath, _revisedPath, _outputPath);

        Assert.Contains(logger.Entries, e => e.Level == LogLevel.Information && e.Message.Contains("insertion"));
    }

    [Fact]
    public void Compare_of_identical_documents_reports_no_changes()
    {
        var summary = _sut.Compare(_originalPath, _originalPath, _outputPath);

        Assert.Equal(0, summary.InsertedCount);
        Assert.Equal(0, summary.DeletedCount);
    }

    [Fact]
    public void CompareDetailed_reports_a_nonzero_percent_changed()
    {
        var summary = _sut.CompareDetailed(_originalPath, _revisedPath, _outputPath);

        Assert.True(summary.InsertedCount > 0);
        Assert.True(summary.PercentChanged > 0);
    }

    [Fact]
    public void CompareDetailed_counts_a_same_text_insert_delete_pair_as_a_format_change()
    {
        var original = TestFiles.NewTempPath(".docx");
        var revised = TestFiles.NewTempPath(".docx");
        var output = TestFiles.NewTempPath(".docx");
        try
        {
            // Same words, only the formatting differs (bold) — WmlComparer represents this as a
            // delete-old-run + insert-new-run pair of identical text, which is exactly the signal
            // FormatChangeCount looks for.
            SampleDocumentFactory.CreateDocumentFromParagraphs(original, [new Paragraph(new Run(new Text("Confidential")))]);
            SampleDocumentFactory.CreateDocumentFromParagraphs(revised,
                [new Paragraph(new Run(new RunProperties(new Bold()), new Text("Confidential")))]);

            var summary = _sut.CompareDetailed(original, revised, output);

            // WmlComparer's own revision stream is empty here (it diffs text content only, not
            // formatting) — FormatChangeCount comes from a separate position-paired pass instead.
            Assert.Equal(0, summary.InsertedCount);
            Assert.Equal(0, summary.DeletedCount);
            Assert.Equal(1, summary.FormatChangeCount);
        }
        finally
        {
            File.Delete(original);
            File.Delete(revised);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public void CompareDetailed_identifies_the_heading_a_change_falls_under()
    {
        var original = TestFiles.NewTempPath(".docx");
        var revised = TestFiles.NewTempPath(".docx");
        var output = TestFiles.NewTempPath(".docx");
        try
        {
            var headingParagraph = new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
                new Run(new Text("Payment Terms")));

            SampleDocumentFactory.CreateDocumentFromParagraphs(original,
                [headingParagraph, new Paragraph(new Run(new Text("Net 30 days.")))]);
            SampleDocumentFactory.CreateDocumentFromParagraphs(revised,
                [(Paragraph)headingParagraph.CloneNode(true), new Paragraph(new Run(new Text("Net 60 days.")))]);

            var summary = _sut.CompareDetailed(original, revised, output);

            Assert.Contains("Payment Terms", summary.AffectedHeadings);
        }
        finally
        {
            File.Delete(original);
            File.Delete(revised);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    public void Dispose()
    {
        File.Delete(_originalPath);
        File.Delete(_revisedPath);
        if (File.Exists(_outputPath))
            File.Delete(_outputPath);
    }
}
