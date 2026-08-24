using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using W15 = DocumentFormat.OpenXml.Office2013.Word;

namespace DocumentProcessor.Core.Comments;

/// <summary>One Word comment, as reviewers see it in the margin.</summary>
/// <param name="Id">The comment's <c>w:id</c>, stable within a document and used to reply, resolve or delete it.</param>
/// <param name="Author">Display name of the reviewer.</param>
/// <param name="Initials">Initials Word shows on the comment bubble.</param>
/// <param name="Date">When the comment was written, if recorded.</param>
/// <param name="Text">The comment's own text.</param>
/// <param name="AnchorText">
/// The document text the comment is attached to. Empty when the comment is anchored at a single
/// point rather than over a range, which is what Word produces when a reviewer comments without
/// selecting anything.
/// </param>
/// <param name="ParentId">The comment this one replies to, or null if it starts a thread.</param>
/// <param name="IsResolved">Whether the thread has been marked resolved in Word.</param>
public sealed record DocumentComment(
    string Id,
    string? Author,
    string? Initials,
    DateTime? Date,
    string Text,
    string AnchorText,
    string? ParentId,
    bool IsResolved);

/// <summary>
/// Reads and writes Word comments — the margin conversation reviewers actually negotiate in.
/// <para>
/// A contract round-trip that only handles tracked changes sees the edits but is blind to the
/// discussion attached to them, which in practice is where the reasoning lives ("we can't accept
/// this indemnity cap", "legal approved 45 days"). Threading and resolution are stored separately
/// from the comments themselves, in the <c>commentsEx</c> part keyed by the comment paragraph's
/// <c>w14:paraId</c>, so both parts are maintained together here.
/// </para>
/// </summary>
public sealed class DocumentCommentService : IDocumentCommentService
{
    /// <summary>Lists every comment in the document, including replies and resolution state.</summary>
    public IReadOnlyList<DocumentComment> GetComments(string docxPath)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: false);
        return GetCommentsCore(doc);
    }

    /// <summary>Runs against an already-open package, so a <see cref="Sessions.DocumentSession"/>
    /// pipeline pays one open/save for the whole sequence instead of one per call.</summary>
    internal static IReadOnlyList<DocumentComment> GetCommentsCore(WordprocessingDocument doc)
    {
        var mainPart = doc.MainDocumentPart ?? throw CorruptDocumentException.MissingBody();
        var commentsPart = mainPart.WordprocessingCommentsPart;
        if (commentsPart?.Comments is null)
            return [];

        var anchors = BuildAnchorIndex(mainPart);
        var threading = BuildThreadingIndex(mainPart);

        var result = new List<DocumentComment>();
        foreach (var comment in commentsPart.Comments.Elements<Comment>())
        {
            var id = comment.Id?.Value;
            if (id is null)
                continue;

            var paraId = comment.Elements<Paragraph>().LastOrDefault()?.ParagraphId?.Value;
            var (parentId, done) = paraId is not null && threading.TryGetValue(paraId, out var t) ? t : (null, false);

            result.Add(new DocumentComment(
                id,
                comment.Author?.Value,
                comment.Initials?.Value,
                comment.Date?.Value,
                comment.InnerText,
                anchors.TryGetValue(id, out var anchor) ? anchor : string.Empty,
                parentId,
                done));
        }

        return result;
    }

    /// <summary>
    /// Attaches a comment to paragraph <paramref name="paragraphIndex"/> (0-based), spanning the
    /// whole paragraph.
    /// </summary>
    /// <returns>The new comment's id.</returns>
    public string AddComment(string docxPath, int paragraphIndex, string author, string initials, string text)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var id = AddCommentCore(doc, paragraphIndex, author, initials, text);
        doc.MainDocumentPart!.Document.Save();
        return id;
    }

    /// <summary>Runs against an already-open package.</summary>
    internal static string AddCommentCore(WordprocessingDocument doc, int paragraphIndex, string author, string initials, string text)
    {
        var mainPart = doc.MainDocumentPart ?? throw CorruptDocumentException.MissingBody();
        var body = mainPart.Document?.Body ?? throw CorruptDocumentException.MissingBody();

        var paragraphs = body.Elements<Paragraph>().ToList();
        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(paragraphIndex),
                $"Document has {paragraphs.Count} paragraphs; valid indices are 0..{paragraphs.Count - 1}.");
        }

        var commentsPart = mainPart.WordprocessingCommentsPart ?? mainPart.AddNewPart<WordprocessingCommentsPart>();
        commentsPart.Comments ??= new DocumentFormat.OpenXml.Wordprocessing.Comments();

        var id = NextCommentId(commentsPart.Comments);
        var paraId = NewParagraphId();

        var commentParagraph = new Paragraph(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }))
        {
            ParagraphId = paraId,
        };

        commentsPart.Comments.AppendChild(new Comment(commentParagraph)
        {
            Id = id,
            Author = author,
            Initials = initials,
            Date = DateTime.UtcNow,
        });
        commentsPart.Comments.Save();

        // Word needs all three markers: a range to highlight, and a reference run that carries the
        // bubble. A comment with only a reference shows up but highlights nothing.
        var target = paragraphs[paragraphIndex];
        target.InsertAt(new CommentRangeStart { Id = id }, 0);
        target.AppendChild(new CommentRangeEnd { Id = id });
        target.AppendChild(new Run(new CommentReference { Id = id }));

        SetThreading(mainPart, paraId, parentParaId: null, done: false);
        return id;
    }

    /// <summary>
    /// Adds a reply to an existing comment, forming a thread the way Word's own reply button does.
    /// </summary>
    /// <returns>The reply's comment id.</returns>
    public string ReplyToComment(string docxPath, string parentCommentId, string author, string initials, string text)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var id = ReplyToCommentCore(doc, parentCommentId, author, initials, text);
        doc.MainDocumentPart!.Document.Save();
        return id;
    }

    /// <summary>Runs against an already-open package.</summary>
    internal static string ReplyToCommentCore(WordprocessingDocument doc, string parentCommentId, string author, string initials, string text)
    {
        var mainPart = doc.MainDocumentPart ?? throw CorruptDocumentException.MissingBody();
        var commentsPart = mainPart.WordprocessingCommentsPart
            ?? throw new ArgumentException("Document has no comments.", nameof(parentCommentId));

        var parent = commentsPart.Comments?.Elements<Comment>().FirstOrDefault(c => c.Id?.Value == parentCommentId)
            ?? throw new ArgumentException($"No comment with id '{parentCommentId}'.", nameof(parentCommentId));

        var parentParaId = parent.Elements<Paragraph>().LastOrDefault()?.ParagraphId?.Value;

        var id = NextCommentId(commentsPart.Comments!);
        var paraId = NewParagraphId();

        commentsPart.Comments!.AppendChild(new Comment(
            new Paragraph(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve })) { ParagraphId = paraId })
        {
            Id = id,
            Author = author,
            Initials = initials,
            Date = DateTime.UtcNow,
        });
        commentsPart.Comments.Save();

        // A reply anchors to the same range as its parent, so Word shows them together.
        var body = mainPart.Document?.Body ?? throw CorruptDocumentException.MissingBody();
        var parentEnd = body.Descendants<CommentRangeEnd>().FirstOrDefault(e => e.Id?.Value == parentCommentId);
        if (parentEnd?.Parent is { } anchorParagraph)
        {
            anchorParagraph.InsertBefore(new CommentRangeStart { Id = id }, parentEnd);
            anchorParagraph.InsertAfter(new CommentRangeEnd { Id = id }, parentEnd);
            anchorParagraph.AppendChild(new Run(new CommentReference { Id = id }));
        }

        SetThreading(mainPart, paraId, parentParaId, done: false);
        return id;
    }

    /// <summary>Marks a comment thread resolved, as Word's "Resolve" does.</summary>
    public void ResolveComment(string docxPath, string commentId, bool resolved = true)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        ResolveCommentCore(doc, commentId, resolved);
        doc.MainDocumentPart!.Document.Save();
    }

    /// <summary>Runs against an already-open package.</summary>
    internal static void ResolveCommentCore(WordprocessingDocument doc, string commentId, bool resolved = true)
    {
        var mainPart = doc.MainDocumentPart ?? throw CorruptDocumentException.MissingBody();
        var comment = mainPart.WordprocessingCommentsPart?.Comments?.Elements<Comment>()
            .FirstOrDefault(c => c.Id?.Value == commentId)
            ?? throw new ArgumentException($"No comment with id '{commentId}'.", nameof(commentId));

        var paraId = comment.Elements<Paragraph>().LastOrDefault()?.ParagraphId?.Value
            ?? throw new CorruptDocumentException($"Comment '{commentId}' has no paragraph id, so its thread state cannot be set.");

        var existingParent = FindCommentEx(mainPart, paraId)?.ParaIdParent?.Value;
        SetThreading(mainPart, paraId, existingParent, resolved);
    }

    /// <summary>Removes a comment and its anchors. Replies to it are left in place.</summary>
    /// <returns>True if a comment with that id was found and removed.</returns>
    public bool DeleteComment(string docxPath, string commentId)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var removed = DeleteCommentCore(doc, commentId);
        if (removed)
            doc.MainDocumentPart!.Document.Save();
        return removed;
    }

    /// <summary>Runs against an already-open package.</summary>
    internal static bool DeleteCommentCore(WordprocessingDocument doc, string commentId)
    {
        var mainPart = doc.MainDocumentPart ?? throw CorruptDocumentException.MissingBody();
        var comment = mainPart.WordprocessingCommentsPart?.Comments?.Elements<Comment>()
            .FirstOrDefault(c => c.Id?.Value == commentId);
        if (comment is null)
            return false;

        var paraId = comment.Elements<Paragraph>().LastOrDefault()?.ParagraphId?.Value;
        comment.Remove();
        mainPart.WordprocessingCommentsPart!.Comments!.Save();

        var body = mainPart.Document?.Body;
        if (body is not null)
        {
            foreach (var start in body.Descendants<CommentRangeStart>().Where(e => e.Id?.Value == commentId).ToList())
                start.Remove();
            foreach (var end in body.Descendants<CommentRangeEnd>().Where(e => e.Id?.Value == commentId).ToList())
                end.Remove();

            // The reference lives inside its own run, which is left empty and pointless otherwise.
            foreach (var reference in body.Descendants<CommentReference>().Where(e => e.Id?.Value == commentId).ToList())
            {
                var run = reference.Ancestors<Run>().FirstOrDefault();
                reference.Remove();
                if (run is not null && !run.HasChildren)
                    run.Remove();
            }
        }

        if (paraId is not null && FindCommentEx(mainPart, paraId) is { } commentEx)
        {
            commentEx.Remove();
            mainPart.WordprocessingCommentsExPart!.CommentsEx!.Save();
        }

        return true;
    }

    /// <summary>Maps comment id to the document text its range covers.</summary>
    private static Dictionary<string, string> BuildAnchorIndex(MainDocumentPart mainPart)
    {
        var anchors = new Dictionary<string, string>();
        var body = mainPart.Document?.Body;
        if (body is null)
            return anchors;

        foreach (var start in body.Descendants<CommentRangeStart>())
        {
            if (start.Id?.Value is not { } id)
                continue;

            // Walk forward from the range start to its matching end, collecting run text. This is
            // the text Word highlights for the comment.
            var text = new System.Text.StringBuilder();
            foreach (var element in start.ElementsAfter())
            {
                if (element is CommentRangeEnd end && end.Id?.Value == id)
                    break;
                if (element is Run run)
                    text.Append(run.InnerText);
            }

            anchors[id] = text.ToString();
        }

        return anchors;
    }

    /// <summary>Maps a comment paragraph's w14:paraId to its parent comment id and resolved flag.</summary>
    private static Dictionary<string, (string? ParentId, bool Done)> BuildThreadingIndex(MainDocumentPart mainPart)
    {
        var index = new Dictionary<string, (string?, bool)>();
        var commentsEx = mainPart.WordprocessingCommentsExPart?.CommentsEx;
        if (commentsEx is null)
            return index;

        // commentsEx refers to parents by paraId; callers think in comment ids, so translate.
        var commentIdByParaId = new Dictionary<string, string>();
        foreach (var comment in mainPart.WordprocessingCommentsPart?.Comments?.Elements<Comment>() ?? [])
        {
            var pid = comment.Elements<Paragraph>().LastOrDefault()?.ParagraphId?.Value;
            if (pid is not null && comment.Id?.Value is { } cid)
                commentIdByParaId[pid] = cid;
        }

        foreach (var ex in commentsEx.Elements<W15.CommentEx>())
        {
            if (ex.ParaId?.Value is not { } paraId)
                continue;

            var parentParaId = ex.ParaIdParent?.Value;
            var parentCommentId = parentParaId is not null && commentIdByParaId.TryGetValue(parentParaId, out var pc) ? pc : null;
            index[paraId] = (parentCommentId, ex.Done?.Value ?? false);
        }

        return index;
    }

    private static W15.CommentEx? FindCommentEx(MainDocumentPart mainPart, string paraId) =>
        mainPart.WordprocessingCommentsExPart?.CommentsEx?.Elements<W15.CommentEx>()
            .FirstOrDefault(e => e.ParaId?.Value == paraId);

    private static void SetThreading(MainDocumentPart mainPart, string paraId, string? parentParaId, bool done)
    {
        var exPart = mainPart.WordprocessingCommentsExPart ?? mainPart.AddNewPart<WordprocessingCommentsExPart>();
        exPart.CommentsEx ??= new W15.CommentsEx();

        var existing = exPart.CommentsEx.Elements<W15.CommentEx>().FirstOrDefault(e => e.ParaId?.Value == paraId);
        existing?.Remove();

        var commentEx = new W15.CommentEx { ParaId = paraId, Done = done };
        if (parentParaId is not null)
            commentEx.ParaIdParent = parentParaId;

        exPart.CommentsEx.AppendChild(commentEx);
        exPart.CommentsEx.Save();
    }

    private static string NextCommentId(DocumentFormat.OpenXml.Wordprocessing.Comments comments)
    {
        var max = comments.Elements<Comment>()
            .Select(c => int.TryParse(c.Id?.Value, out var n) ? n : 0)
            .DefaultIfEmpty(0)
            .Max();
        return (max + 1).ToString();
    }

    /// <summary>
    /// w14:paraId is a 8-digit hex value that must be unique in the document and, per the schema,
    /// must not be 00000000 or exceed 7FFFFFFF.
    /// </summary>
    private static string NewParagraphId()
    {
        var value = (uint)Random.Shared.Next(1, int.MaxValue);
        return value.ToString("X8");
    }
}
