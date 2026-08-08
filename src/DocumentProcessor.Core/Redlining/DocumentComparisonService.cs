using Clippit;
using Clippit.Word;

namespace DocumentProcessor.Core.Redlining;

public sealed record ChangeSummary(int InsertedCount, int DeletedCount, IReadOnlyList<string> InsertedText, IReadOnlyList<string> DeletedText);

/// <summary>
/// Compares two .docx files and produces a redlined document — insertions/deletions expressed as
/// standard Word tracked changes (w:ins/w:del), openable and reviewable in Word itself — via
/// Clippit's WmlComparer, a maintained fork of Open-Xml-PowerTools' comparison engine.
/// </summary>
public sealed class DocumentComparisonService
{
    /// <summary>
    /// Compares <paramref name="originalPath"/> against <paramref name="revisedPath"/>, writes the
    /// redlined result (revised content with tracked changes showing the diff) to
    /// <paramref name="outputRedlinedPath"/>, and returns a structured summary of what changed.
    /// </summary>
    public ChangeSummary Compare(string originalPath, string revisedPath, string outputRedlinedPath, string authorForRevisions = "Document Comparison")
    {
        var settings = new WmlComparerSettings { AuthorForRevisions = authorForRevisions };

        var redlined = WmlComparer.Compare(new WmlDocument(originalPath), new WmlDocument(revisedPath), settings);
        redlined.SaveAs(outputRedlinedPath);

        var revisions = WmlComparer.GetRevisions(redlined, settings);
        var inserted = revisions
            .Where(r => r.RevisionType == WmlComparer.WmlComparerRevisionType.Inserted)
            .Select(r => r.Text)
            .ToList();
        var deleted = revisions
            .Where(r => r.RevisionType == WmlComparer.WmlComparerRevisionType.Deleted)
            .Select(r => r.Text)
            .ToList();

        return new ChangeSummary(inserted.Count, deleted.Count, inserted, deleted);
    }
}
