using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.Comments;
using DocumentProcessor.Core.Samples;
using DocumentProcessor.Core.Sessions;

namespace DocumentProcessor.Tests.Comments;

public class DocumentCommentServiceTests : IDisposable
{
    private readonly List<string> _cleanup = [];
    private readonly DocumentCommentService _service = new();

    private string NewContract()
    {
        var path = TestFiles.NewTempPath(".docx");
        _cleanup.Add(path);
        SampleDocumentFactory.CreateBasicDocument(path, "Services Agreement",
        [
            "1. Term. This Agreement runs for twelve months.",
            "2. Fees. Client shall pay $150,000 annually.",
            "3. Liability. Liability is capped at fees paid.",
        ]);
        return path;
    }

    [Fact]
    public void A_document_with_no_comments_reports_none()
    {
        Assert.Empty(_service.GetComments(NewContract()));
    }

    [Fact]
    public void A_comment_round_trips_with_its_author_and_anchor()
    {
        var path = NewContract();

        var id = _service.AddComment(path, paragraphIndex: 3, "Jordan Ellis", "JE",
            "Cap is too low — push for 2x fees.");

        var comment = Assert.Single(_service.GetComments(path));
        Assert.Equal(id, comment.Id);
        Assert.Equal("Jordan Ellis", comment.Author);
        Assert.Equal("JE", comment.Initials);
        Assert.Equal("Cap is too low — push for 2x fees.", comment.Text);
        Assert.Contains("Liability is capped", comment.AnchorText);
        Assert.Null(comment.ParentId);
        Assert.False(comment.IsResolved);
        Assert.NotNull(comment.Date);
    }

    [Fact]
    public void The_anchor_markers_Word_needs_are_all_written()
    {
        var path = NewContract();
        var id = _service.AddComment(path, 1, "Jordan Ellis", "JE", "Confirm the term.");

        using var doc = WordprocessingDocument.Open(path, isEditable: false);
        var body = doc.MainDocumentPart!.Document!.Body!;

        // A comment missing any one of these renders wrong in Word: no highlight, or no bubble.
        Assert.Contains(body.Descendants<CommentRangeStart>(), e => e.Id?.Value == id);
        Assert.Contains(body.Descendants<CommentRangeEnd>(), e => e.Id?.Value == id);
        Assert.Contains(body.Descendants<CommentReference>(), e => e.Id?.Value == id);
    }

    [Fact]
    public void A_reply_is_threaded_under_its_parent()
    {
        var path = NewContract();
        var parentId = _service.AddComment(path, 2, "Jordan Ellis", "JE", "Fees look high.");

        var replyId = _service.ReplyToComment(path, parentId, "Sam Okafor", "SO", "Approved by finance.");

        var comments = _service.GetComments(path);
        Assert.Equal(2, comments.Count);

        var reply = comments.Single(c => c.Id == replyId);
        Assert.Equal(parentId, reply.ParentId);
        Assert.Equal("Sam Okafor", reply.Author);
        Assert.Null(comments.Single(c => c.Id == parentId).ParentId);
    }

    [Fact]
    public void Replying_to_an_unknown_comment_is_rejected()
    {
        var path = NewContract();
        _service.AddComment(path, 1, "Jordan Ellis", "JE", "Anything.");

        Assert.Throws<ArgumentException>(() =>
            _service.ReplyToComment(path, "999", "Sam Okafor", "SO", "Reply to nothing."));
    }

    [Fact]
    public void Resolving_a_comment_survives_the_round_trip_and_can_be_undone()
    {
        var path = NewContract();
        var id = _service.AddComment(path, 1, "Jordan Ellis", "JE", "Check the term.");

        _service.ResolveComment(path, id);
        Assert.True(_service.GetComments(path).Single().IsResolved);

        _service.ResolveComment(path, id, resolved: false);
        Assert.False(_service.GetComments(path).Single().IsResolved);
    }

    [Fact]
    public void Resolving_a_reply_keeps_it_threaded()
    {
        var path = NewContract();
        var parentId = _service.AddComment(path, 2, "Jordan Ellis", "JE", "Fees look high.");
        var replyId = _service.ReplyToComment(path, parentId, "Sam Okafor", "SO", "Approved.");

        // Resolution and parentage share one commentsEx entry, so setting one must not drop the other.
        _service.ResolveComment(path, replyId);

        var reply = _service.GetComments(path).Single(c => c.Id == replyId);
        Assert.True(reply.IsResolved);
        Assert.Equal(parentId, reply.ParentId);
    }

    [Fact]
    public void Deleting_a_comment_removes_it_and_every_anchor()
    {
        var path = NewContract();
        var keep = _service.AddComment(path, 1, "Jordan Ellis", "JE", "Keep this one.");
        var drop = _service.AddComment(path, 2, "Jordan Ellis", "JE", "Delete this one.");

        Assert.True(_service.DeleteComment(path, drop));

        var remaining = Assert.Single(_service.GetComments(path));
        Assert.Equal(keep, remaining.Id);

        using var doc = WordprocessingDocument.Open(path, isEditable: false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        Assert.DoesNotContain(body.Descendants<CommentRangeStart>(), e => e.Id?.Value == drop);
        Assert.DoesNotContain(body.Descendants<CommentRangeEnd>(), e => e.Id?.Value == drop);
        Assert.DoesNotContain(body.Descendants<CommentReference>(), e => e.Id?.Value == drop);
        // A dangling commentsEx entry makes Word repair the file on open.
        Assert.DoesNotContain("w15:paraIdParent", doc.MainDocumentPart.WordprocessingCommentsExPart?.CommentsEx?.InnerXml ?? "");
    }

    [Fact]
    public void Deleting_a_comment_that_does_not_exist_reports_false()
    {
        Assert.False(_service.DeleteComment(NewContract(), "42"));
    }

    [Fact]
    public void Commenting_outside_the_document_is_rejected()
    {
        var path = NewContract();
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            _service.AddComment(path, 99, "Jordan Ellis", "JE", "Nowhere."));
    }

    [Fact]
    public void A_whole_review_thread_can_be_built_in_one_session()
    {
        using var session = DocumentSession.Open(File.ReadAllBytes(NewContract()));

        var id = session.Comments.Add(3, "Jordan Ellis", "JE", "Cap is too low.");
        session.Comments.Reply(id, "Sam Okafor", "SO", "Agreed, countering at 2x.");
        session.Comments.Resolve(id);

        var bytes = session.Save();

        var comments = ReadComments(bytes);
        Assert.Equal(2, comments.Count);
        Assert.True(comments.Single(c => c.Id == id).IsResolved);
        Assert.Equal(id, comments.Single(c => c.Id != id).ParentId);
    }

    private static IReadOnlyList<DocumentComment> ReadComments(byte[] docxBytes)
    {
        var path = TestFiles.NewTempPath(".docx");
        File.WriteAllBytes(path, docxBytes);
        try
        {
            return new DocumentCommentService().GetComments(path);
        }
        finally
        {
            File.Delete(path);
        }
    }

    public void Dispose()
    {
        foreach (var path in _cleanup)
            if (File.Exists(path)) File.Delete(path);
    }
}
