using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.Redlining;
using DocumentProcessor.Core.Samples;

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
    public void Compare_of_identical_documents_reports_no_changes()
    {
        var summary = _sut.Compare(_originalPath, _originalPath, _outputPath);

        Assert.Equal(0, summary.InsertedCount);
        Assert.Equal(0, summary.DeletedCount);
    }

    public void Dispose()
    {
        File.Delete(_originalPath);
        File.Delete(_revisedPath);
        if (File.Exists(_outputPath))
            File.Delete(_outputPath);
    }
}
