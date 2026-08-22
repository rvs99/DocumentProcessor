using System.Xml.Linq;
using Clippit;
using Clippit.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentProcessor.Core.Redlining;

public sealed record ChangeSummary(int InsertedCount, int DeletedCount, IReadOnlyList<string> InsertedText, IReadOnlyList<string> DeletedText);

/// <summary>
/// A richer comparison result than <see cref="ChangeSummary"/>: alongside insert/delete counts, it
/// estimates how much of the document actually changed and which headings the changes fall under —
/// useful for a "here's what changed" review summary rather than a raw diff dump.
/// </summary>
public sealed record ComparisonSummary(
    int InsertedCount,
    int DeletedCount,
    /// <summary>Paragraphs at the same position in both documents with identical text but different
    /// run formatting (e.g. a bold/italic toggle with no wording change). Confirmed empirically that
    /// WmlComparer's own revision stream contains *zero* entries for a pure formatting edit — it
    /// diffs text content only — so this is computed as a separate, position-paired pass rather than
    /// derived from <see cref="WmlComparer.GetRevisions"/>. Being position-based, a paragraph
    /// inserted/removed earlier in the document can misalign this pairing for everything after it;
    /// it's accurate for the common case (formatting-only edits with no other structural change).</summary>
    int FormatChangeCount,
    /// <summary>(Words inserted + words deleted) / words in the original document, as a percentage.</summary>
    double PercentChanged,
    /// <summary>Text of every Heading-styled paragraph that a change falls under or immediately follows.</summary>
    IReadOnlyList<string> AffectedHeadings,
    IReadOnlyList<string> InsertedText,
    IReadOnlyList<string> DeletedText);

/// <summary>
/// Compares two .docx files and produces a redlined document — insertions/deletions expressed as
/// standard Word tracked changes (w:ins/w:del), openable and reviewable in Word itself — via
/// Clippit's WmlComparer, a maintained fork of Open-Xml-PowerTools' comparison engine.
/// </summary>
public sealed class DocumentComparisonService
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>
    /// Compares <paramref name="originalPath"/> against <paramref name="revisedPath"/>, writes the
    /// redlined result (revised content with tracked changes showing the diff) to
    /// <paramref name="outputRedlinedPath"/>, and returns a structured summary of what changed.
    /// </summary>
    public ChangeSummary Compare(string originalPath, string revisedPath, string outputRedlinedPath, string authorForRevisions = "Document Comparison")
    {
        var (inserted, deleted) = CompareCore(originalPath, revisedPath, outputRedlinedPath, authorForRevisions);
        return new ChangeSummary(inserted.Count, deleted.Count, inserted.Select(r => r.Text).ToList(), deleted.Select(r => r.Text).ToList());
    }

    /// <summary>Same comparison as <see cref="Compare"/>, but returns the richer <see cref="ComparisonSummary"/>
    /// (format-change estimate, percent changed, affected headings) instead of the bare counts.</summary>
    public ComparisonSummary CompareDetailed(string originalPath, string revisedPath, string outputRedlinedPath, string authorForRevisions = "Document Comparison")
    {
        var (inserted, deleted) = CompareCore(originalPath, revisedPath, outputRedlinedPath, authorForRevisions);

        return new ComparisonSummary(
            inserted.Count,
            deleted.Count,
            CountFormatOnlyChanges(originalPath, revisedPath),
            ComputePercentChanged(originalPath, inserted, deleted),
            FindAffectedHeadings(inserted.Concat(deleted)),
            inserted.Select(r => r.Text).ToList(),
            deleted.Select(r => r.Text).ToList());
    }

    private static (List<WmlComparer.WmlComparerRevision> Inserted, List<WmlComparer.WmlComparerRevision> Deleted) CompareCore(
        string originalPath, string revisedPath, string outputRedlinedPath, string authorForRevisions)
    {
        var settings = new WmlComparerSettings { AuthorForRevisions = authorForRevisions };

        var redlined = WmlComparer.Compare(new WmlDocument(originalPath), new WmlDocument(revisedPath), settings);
        redlined.SaveAs(outputRedlinedPath);

        var revisions = WmlComparer.GetRevisions(redlined, settings);
        var inserted = revisions.Where(r => r.RevisionType == WmlComparer.WmlComparerRevisionType.Inserted).ToList();
        var deleted = revisions.Where(r => r.RevisionType == WmlComparer.WmlComparerRevisionType.Deleted).ToList();
        return (inserted, deleted);
    }

    /// <summary>Pairs up paragraphs by position and flags ones with identical text but different
    /// run formatting — see the caveats on <see cref="ComparisonSummary.FormatChangeCount"/>.</summary>
    private static int CountFormatOnlyChanges(string originalPath, string revisedPath)
    {
        using var originalDoc = WordprocessingDocument.Open(originalPath, isEditable: false);
        using var revisedDoc = WordprocessingDocument.Open(revisedPath, isEditable: false);

        var originalParagraphs = originalDoc.MainDocumentPart?.Document?.Body?.Elements<Paragraph>().ToList() ?? [];
        var revisedParagraphs = revisedDoc.MainDocumentPart?.Document?.Body?.Elements<Paragraph>().ToList() ?? [];

        var count = 0;
        for (var i = 0; i < Math.Min(originalParagraphs.Count, revisedParagraphs.Count); i++)
        {
            if (originalParagraphs[i].InnerText != revisedParagraphs[i].InnerText)
                continue; // a real content change — WmlComparer's insert/delete counts already cover this

            var originalFormatting = string.Concat(originalParagraphs[i].Descendants<RunProperties>().Select(p => p.OuterXml));
            var revisedFormatting = string.Concat(revisedParagraphs[i].Descendants<RunProperties>().Select(p => p.OuterXml));
            if (originalFormatting != revisedFormatting)
                count++;
        }

        return count;
    }

    private static double ComputePercentChanged(string originalPath, List<WmlComparer.WmlComparerRevision> inserted, List<WmlComparer.WmlComparerRevision> deleted)
    {
        using var doc = WordprocessingDocument.Open(originalPath, isEditable: false);
        var originalWordCount = CountWords(doc.MainDocumentPart?.Document?.Body?.InnerText ?? "");
        if (originalWordCount == 0)
            return 0;

        var changedWordCount = inserted.Sum(r => CountWords(r.Text)) + deleted.Sum(r => CountWords(r.Text));
        return Math.Round(100.0 * changedWordCount / originalWordCount, 1);
    }

    private static int CountWords(string text) =>
        text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length;

    private static IReadOnlyList<string> FindAffectedHeadings(IEnumerable<WmlComparer.WmlComparerRevision> revisions)
    {
        var headings = new List<string>();

        foreach (var revision in revisions)
        {
            var element = revision.RevisionXElement ?? revision.ContentXElement;
            var paragraph = element?.AncestorsAndSelf(W + "p").FirstOrDefault();

            // Walk backward through preceding sibling paragraphs until a Heading-styled one is found —
            // a change doesn't have to sit *inside* the heading paragraph itself to fall "under" it.
            while (paragraph is not null)
            {
                var styleId = paragraph.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val")?.Value;
                if (styleId is not null && styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase))
                {
                    var headingText = string.Concat(paragraph.Descendants(W + "t").Select(t => t.Value));
                    if (!headings.Contains(headingText))
                        headings.Add(headingText);
                    break;
                }

                paragraph = paragraph.ElementsBeforeSelf(W + "p").LastOrDefault();
            }
        }

        return headings;
    }
}
