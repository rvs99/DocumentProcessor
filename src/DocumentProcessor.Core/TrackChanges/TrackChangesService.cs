using Clippit.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.Diagnostics;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;

namespace DocumentProcessor.Core.TrackChanges;

public enum TrackedChangeKind { Insertion, Deletion }

/// <summary>One tracked-change run, as recorded by Word itself — the same author/date/id already
/// present in the OOXML (<c>w:author</c>/<c>w:date</c>/<c>w:id</c> on <c>w:ins</c>/<c>w:del</c>),
/// just surfaced as data instead of requiring callers to walk the XML themselves.</summary>
public sealed record TrackedChange(string? ChangeId, string? Author, DateTime? Date, TrackedChangeKind Kind, int ParagraphIndex, string Text);

/// <summary>
/// Accepts or rejects Word's tracked changes (w:ins/w:del run-level revisions) programmatically —
/// e.g. to finalize a document after negotiation, or to roll back a batch of proposed edits.
/// Accept-all reuses Clippit's RevisionAccepter; everything else (reject, and every author/id-scoped
/// variant) is hand-rolled since no free library ships selective accept/reject.
/// Scope: covers run-level insertions/deletions, which is the vast majority of real-world tracked
/// changes. Paragraph-mark and formatting-change revisions (pPrChange/rPrChange) are left as-is.
/// </summary>
public sealed class TrackChangesService(ILogger<TrackChangesService>? logger = null)
{
    private readonly ILogger<TrackChangesService> _logger = logger ?? NullLogger<TrackChangesService>.Instance;

    /// <summary>Accepts every tracked change: inserted text is kept, deleted text is discarded.</summary>
    public void AcceptAll(string docxPath)
    {
        using var activity = DocumentProcessorDiagnostics.ActivitySource.StartActivity("TrackChangesService.AcceptAll");
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        RevisionAccepter.AcceptRevisions(doc);
        _logger.LogInformation("Accepted all tracked changes in {DocxPath}", docxPath);
    }

    /// <summary>Accepts only the tracked changes made by <paramref name="author"/> (an exact match
    /// against <c>w:author</c>), leaving every other author's changes still pending.</summary>
    public void AcceptByAuthor(string docxPath, string author) =>
        AcceptWhere(docxPath, (_, a) => a == author);

    /// <summary>Accepts only the single tracked change with the given <c>w:id</c>.</summary>
    public void AcceptById(string docxPath, string changeId) =>
        AcceptWhere(docxPath, (id, _) => id == changeId);

    /// <summary>Rejects every tracked change: inserted text is discarded, deleted text is restored.</summary>
    public void RejectAll(string docxPath) => RejectWhere(docxPath, (_, _) => true);

    /// <summary>Rejects only the tracked changes made by <paramref name="author"/>, leaving every
    /// other author's changes still pending.</summary>
    public void RejectByAuthor(string docxPath, string author) =>
        RejectWhere(docxPath, (_, a) => a == author);

    /// <summary>Rejects only the single tracked change with the given <c>w:id</c>.</summary>
    public void RejectById(string docxPath, string changeId) =>
        RejectWhere(docxPath, (id, _) => id == changeId);

    /// <summary>Whether the document contains any unresolved tracked changes.</summary>
    public bool HasTrackedChanges(string docxPath)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: false);
        return RevisionAccepter.HasTrackedRevisions(doc);
    }

    /// <summary>Lists every unresolved tracked change with its author, timestamp, id, kind, and the
    /// (0-based) paragraph index it lives in — for building a review UI or audit trail without the
    /// caller having to walk <c>w:ins</c>/<c>w:del</c> elements directly.</summary>
    public IReadOnlyList<TrackedChange> GetTrackedChanges(string docxPath)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: false);
        var body = doc.MainDocumentPart?.Document?.Body ?? throw new CorruptDocumentException("Document has no main part/body.");
        var paragraphs = body.Elements<Paragraph>().ToList();

        var changes = new List<TrackedChange>();
        for (var i = 0; i < paragraphs.Count; i++)
        {
            foreach (var insertedRun in paragraphs[i].Descendants<InsertedRun>())
            {
                var text = string.Concat(insertedRun.Descendants<Text>().Select(t => t.Text));
                changes.Add(new TrackedChange(insertedRun.Id?.Value, insertedRun.Author?.Value, insertedRun.Date?.Value, TrackedChangeKind.Insertion, i, text));
            }

            foreach (var deletedRun in paragraphs[i].Descendants<DeletedRun>())
            {
                var text = string.Concat(deletedRun.Descendants<DeletedText>().Select(t => t.Text));
                changes.Add(new TrackedChange(deletedRun.Id?.Value, deletedRun.Author?.Value, deletedRun.Date?.Value, TrackedChangeKind.Deletion, i, text));
            }
        }

        return changes;
    }

    private void AcceptWhere(string docxPath, Func<string?, string?, bool> matches)
    {
        using var activity = DocumentProcessorDiagnostics.ActivitySource.StartActivity("TrackChangesService.AcceptWhere");
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var document = doc.MainDocumentPart?.Document ?? throw new CorruptDocumentException("Document has no main part/body.");
        var resolvedCount = 0;

        // Accepting an insertion keeps its content — unwrap the w:ins wrapper in place.
        foreach (var insertedRun in document.Descendants<InsertedRun>().ToList())
        {
            if (matches(insertedRun.Id?.Value, insertedRun.Author?.Value))
            {
                UnwrapInPlace(insertedRun);
                resolvedCount++;
            }
        }

        // Accepting a deletion confirms it — the deleted content goes away entirely.
        foreach (var deletedRun in document.Descendants<DeletedRun>().ToList())
        {
            if (matches(deletedRun.Id?.Value, deletedRun.Author?.Value))
            {
                deletedRun.Remove();
                resolvedCount++;
            }
        }

        document.Save();
        _logger.LogInformation("Accepted {Count} tracked change(s) in {DocxPath}", resolvedCount, docxPath);
    }

    private void RejectWhere(string docxPath, Func<string?, string?, bool> matches)
    {
        using var activity = DocumentProcessorDiagnostics.ActivitySource.StartActivity("TrackChangesService.RejectWhere");
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var document = doc.MainDocumentPart?.Document ?? throw new CorruptDocumentException("Document has no main part/body.");
        var resolvedCount = 0;

        // Rejecting an insertion undoes it — the inserted content goes away entirely.
        foreach (var insertedRun in document.Descendants<InsertedRun>().ToList())
        {
            if (matches(insertedRun.Id?.Value, insertedRun.Author?.Value))
            {
                insertedRun.Remove();
                resolvedCount++;
            }
        }

        // Rejecting a deletion restores it — unwrap the w:del wrapper, turning w:delText back into w:t.
        foreach (var deletedRun in document.Descendants<DeletedRun>().ToList())
        {
            if (!matches(deletedRun.Id?.Value, deletedRun.Author?.Value))
                continue;

            RestoreDeletedText(deletedRun);
            UnwrapInPlace(deletedRun);
            resolvedCount++;
        }

        _logger.LogInformation("Rejected {Count} tracked change(s) in {DocxPath}", resolvedCount, docxPath);

        document.Save();
    }

    private static void RestoreDeletedText(DeletedRun deletedRun)
    {
        foreach (var run in deletedRun.Elements<Run>())
        {
            foreach (var deletedText in run.Elements<DeletedText>().ToList())
            {
                var restored = new Text(deletedText.Text) { Space = deletedText.Space };
                deletedText.InsertBeforeSelf(restored);
                deletedText.Remove();
            }
        }
    }

    /// <summary>Replaces a wrapper element (e.g. w:del) with its own children, then removes the now-empty wrapper.</summary>
    private static void UnwrapInPlace(DocumentFormat.OpenXml.OpenXmlElement wrapper)
    {
        foreach (var child in wrapper.ChildElements.ToList())
        {
            child.Remove();
            wrapper.InsertBeforeSelf(child);
        }

        wrapper.Remove();
    }
}
