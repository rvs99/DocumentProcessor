using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.Samples;
using DocumentProcessor.Core.TrackChanges;

namespace DocumentProcessor.Tests.TrackChanges;

public class TrackChangesSelectiveTests : IDisposable
{
    private readonly string _path = TestFiles.NewTempPath(".docx");
    private readonly TrackChangesService _sut = new();
    private static readonly DateTimeValue Date = new(new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc));

    public TrackChangesSelectiveTests()
    {
        var paragraph1 = new Paragraph(
            new Run(new Text("Alpha ") { Space = SpaceProcessingModeValues.Preserve }),
            new InsertedRun(new Run(new Text("new-alpha"))) { Id = "1", Author = "Alice", Date = Date },
            new DeletedRun(new Run(new DeletedText("old-alpha"))) { Id = "2", Author = "Alice", Date = Date });
        var paragraph2 = new Paragraph(
            new Run(new Text("Beta ") { Space = SpaceProcessingModeValues.Preserve }),
            new InsertedRun(new Run(new Text("new-beta"))) { Id = "3", Author = "Bob", Date = Date },
            new DeletedRun(new Run(new DeletedText("old-beta"))) { Id = "4", Author = "Bob", Date = Date });

        SampleDocumentFactory.CreateDocumentFromParagraphs(_path, [paragraph1, paragraph2]);
    }

    [Fact]
    public void GetTrackedChanges_reports_author_date_kind_id_and_paragraph_index_for_every_change()
    {
        var changes = _sut.GetTrackedChanges(_path);

        Assert.Equal(4, changes.Count);
        Assert.Contains(changes, c => c.Author == "Alice" && c.Kind == TrackedChangeKind.Insertion && c.Text == "new-alpha" && c.ParagraphIndex == 0 && c.ChangeId == "1");
        Assert.Contains(changes, c => c.Author == "Alice" && c.Kind == TrackedChangeKind.Deletion && c.Text == "old-alpha" && c.ParagraphIndex == 0 && c.ChangeId == "2");
        Assert.Contains(changes, c => c.Author == "Bob" && c.Kind == TrackedChangeKind.Insertion && c.Text == "new-beta" && c.ParagraphIndex == 1 && c.ChangeId == "3");
        Assert.All(changes, c => Assert.Equal(Date.Value, c.Date));
    }

    [Fact]
    public void AcceptByAuthor_resolves_only_that_authors_changes()
    {
        _sut.AcceptByAuthor(_path, "Alice");

        var text = GetParagraphTexts(_path);
        Assert.Equal("Alpha new-alpha", text[0]); // Alice's change accepted: insertion kept, deletion gone
        Assert.Contains("new-beta", text[1]);
        Assert.Contains("old-beta", text[1]); // Bob's change untouched, still pending
    }

    [Fact]
    public void RejectByAuthor_resolves_only_that_authors_changes()
    {
        _sut.RejectByAuthor(_path, "Bob");

        var text = GetParagraphTexts(_path);
        Assert.Equal("Beta old-beta", text[1]); // Bob's change rejected: deletion restored, insertion gone
        Assert.Contains("new-alpha", text[0]);
        Assert.Contains("old-alpha", text[0]); // Alice's change untouched, still pending
    }

    [Fact]
    public void AcceptById_resolves_only_the_matching_change_id()
    {
        _sut.AcceptById(_path, "1"); // Alice's insertion only

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        Assert.DoesNotContain(body.Descendants<InsertedRun>(), r => r.Id?.Value == "1");
        Assert.Contains(body.Descendants<DeletedRun>(), r => r.Id?.Value == "2"); // Alice's deletion still pending
        Assert.Contains(body.Descendants<InsertedRun>(), r => r.Id?.Value == "3"); // Bob's changes untouched
    }

    [Fact]
    public void RejectById_resolves_only_the_matching_change_id()
    {
        _sut.RejectById(_path, "4"); // Bob's deletion only

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        Assert.DoesNotContain(body.Descendants<DeletedRun>(), r => r.Id?.Value == "4");
        Assert.Contains(body.Descendants<InsertedRun>(), r => r.Id?.Value == "3"); // Bob's insertion still pending
        Assert.Contains(body.Descendants<InsertedRun>(), r => r.Id?.Value == "1"); // Alice's changes untouched
    }

    private static List<string> GetParagraphTexts(string path)
    {
        using var doc = WordprocessingDocument.Open(path, isEditable: false);
        return doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().Select(p => p.InnerText).ToList();
    }

    public void Dispose() => File.Delete(_path);
}
