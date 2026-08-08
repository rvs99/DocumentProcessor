using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.Samples;
using DocumentProcessor.Core.TrackChanges;

namespace DocumentProcessor.Tests.TrackChanges;

public class TrackChangesServiceTests : IDisposable
{
    private readonly string _path = TestFiles.NewTempPath(".docx");
    private readonly TrackChangesService _sut = new();

    public TrackChangesServiceTests()
    {
        SampleDocumentFactory.CreateDocumentWithTrackedChanges(_path);
    }

    [Fact]
    public void HasTrackedChanges_is_true_before_resolving_and_false_after()
    {
        Assert.True(_sut.HasTrackedChanges(_path));

        _sut.AcceptAll(_path);

        Assert.False(_sut.HasTrackedChanges(_path));
    }

    [Fact]
    public void AcceptAll_keeps_inserted_text_and_discards_deleted_text()
    {
        _sut.AcceptAll(_path);

        var text = GetBodyText(_path);
        Assert.Contains("twenty-four", text);
        Assert.DoesNotContain("twelve", text);
        Assert.Equal("The term is twenty-four months.", text);
    }

    [Fact]
    public void RejectAll_discards_inserted_text_and_restores_deleted_text()
    {
        _sut.RejectAll(_path);

        var text = GetBodyText(_path);
        Assert.Contains("twelve", text);
        Assert.DoesNotContain("twenty-four", text);
        Assert.Equal("The term is twelve months.", text);
    }

    [Fact]
    public void RejectAll_leaves_no_ins_or_del_markers_behind()
    {
        _sut.RejectAll(_path);

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        Assert.Empty(body.Descendants<InsertedRun>());
        Assert.Empty(body.Descendants<DeletedRun>());
        Assert.False(_sut.HasTrackedChanges(_path));
    }

    private static string GetBodyText(string path)
    {
        using var doc = WordprocessingDocument.Open(path, isEditable: false);
        return doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First().InnerText;
    }

    public void Dispose() => File.Delete(_path);
}
