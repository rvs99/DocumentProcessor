using Clippit.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentProcessor.Core.TrackChanges;

/// <summary>
/// Accepts or rejects Word's tracked changes (w:ins/w:del run-level revisions) programmatically —
/// e.g. to finalize a document after negotiation, or to roll back a batch of proposed edits.
/// Accept reuses Clippit's RevisionAccepter; reject is hand-rolled since no free library ships one
/// (the reverse operation is the symmetric case: discard insertions, restore deletions).
/// Scope: covers run-level insertions/deletions, which is the vast majority of real-world tracked
/// changes. Paragraph-mark and formatting-change revisions (pPrChange/rPrChange) are left as-is.
/// </summary>
public sealed class TrackChangesService
{
    /// <summary>Accepts every tracked change: inserted text is kept, deleted text is discarded.</summary>
    public void AcceptAll(string docxPath)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        RevisionAccepter.AcceptRevisions(doc);
    }

    /// <summary>Rejects every tracked change: inserted text is discarded, deleted text is restored.</summary>
    public void RejectAll(string docxPath)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var document = doc.MainDocumentPart?.Document ?? throw new InvalidOperationException("Document has no main part/body.");

        foreach (var insertedRun in document.Descendants<InsertedRun>().ToList())
            insertedRun.Remove();

        foreach (var deletedRun in document.Descendants<DeletedRun>().ToList())
        {
            RestoreDeletedText(deletedRun);
            UnwrapInPlace(deletedRun);
        }

        document.Save();
    }

    /// <summary>Whether the document contains any unresolved tracked changes.</summary>
    public bool HasTrackedChanges(string docxPath)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: false);
        return RevisionAccepter.HasTrackedRevisions(doc);
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
