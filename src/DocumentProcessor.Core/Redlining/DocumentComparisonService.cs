using System.Xml.Linq;
using Clippit;
using Clippit.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.Diagnostics;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;

namespace DocumentProcessor.Core.Redlining;

public sealed record ChangeSummary(int InsertedCount, int DeletedCount, IReadOnlyList<string> InsertedText, IReadOnlyList<string> DeletedText);

/// <summary>
/// A richer comparison result than <see cref="ChangeSummary"/>: alongside insert/delete counts, it
/// estimates how much of the document actually changed and which headings the changes fall under —
/// useful for a "here's what changed" review summary rather than a raw diff dump.
/// </summary>
/// <param name="InsertedCount">Number of inserted revisions.</param>
/// <param name="DeletedCount">Number of deleted revisions.</param>
/// <param name="FormatChangeCount">
/// Paragraphs at the same position in both documents with identical text but different run
/// formatting (e.g. a bold/italic toggle with no wording change). Confirmed empirically that
/// WmlComparer's own revision stream contains <em>zero</em> entries for a pure formatting edit — it
/// diffs text content only — so this is computed as a separate, position-paired pass rather than
/// derived from the comparer's revisions. Being position-based, a paragraph inserted or removed
/// earlier in the document can misalign the pairing for everything after it; it is accurate for the
/// common case of formatting-only edits with no other structural change.
/// </param>
/// <param name="PercentChanged">(Words inserted + words deleted) / words in the original document, as a percentage.</param>
/// <param name="AffectedHeadings">Text of every Heading-styled paragraph that a change falls under.</param>
/// <param name="InsertedText">The text of each insertion.</param>
/// <param name="DeletedText">The text of each deletion.</param>
public sealed record ComparisonSummary(
    int InsertedCount,
    int DeletedCount,
    int FormatChangeCount,
    double PercentChanged,
    IReadOnlyList<string> AffectedHeadings,
    IReadOnlyList<string> InsertedText,
    IReadOnlyList<string> DeletedText);

/// <summary>
/// Compares two .docx files and produces a redlined document — insertions/deletions expressed as
/// standard Word tracked changes (w:ins/w:del), openable and reviewable in Word itself — via
/// Clippit's WmlComparer, a maintained fork of Open-Xml-PowerTools' comparison engine.
/// </summary>
public sealed class DocumentComparisonService(ILogger<DocumentComparisonService>? logger = null)
{
    private readonly ILogger<DocumentComparisonService> _logger = logger ?? NullLogger<DocumentComparisonService>.Instance;
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

    private (List<WmlComparer.WmlComparerRevision> Inserted, List<WmlComparer.WmlComparerRevision> Deleted) CompareCore(
        string originalPath, string revisedPath, string outputRedlinedPath, string authorForRevisions)
    {
        using var activity = DocumentProcessorDiagnostics.ActivitySource.StartActivity("DocumentComparisonService.Compare");
        _logger.LogDebug("Comparing {OriginalPath} against {RevisedPath} -> {OutputPath}", originalPath, revisedPath, outputRedlinedPath);

        var settings = new WmlComparerSettings { AuthorForRevisions = authorForRevisions };

        var redlined = WmlComparer.Compare(new WmlDocument(originalPath), new WmlDocument(revisedPath), settings);
        redlined.SaveAs(outputRedlinedPath);

        var revisions = WmlComparer.GetRevisions(redlined, settings);
        var inserted = revisions.Where(r => r.RevisionType == WmlComparer.WmlComparerRevisionType.Inserted).ToList();
        var deleted = revisions.Where(r => r.RevisionType == WmlComparer.WmlComparerRevisionType.Deleted).ToList();
        _logger.LogInformation("Comparison of {OriginalPath} vs {RevisedPath} found {InsertedCount} insertion(s), {DeletedCount} deletion(s)",
            originalPath, revisedPath, inserted.Count, deleted.Count);
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

            if (!SameRunFormatting(originalParagraphs[i], revisedParagraphs[i]))
                count++;
        }

        return count;
    }

    /// <summary>
    /// Compares two paragraphs' run formatting by walking both <c>w:rPr</c> sequences in step.
    /// Previously this concatenated every <c>OuterXml</c> into two strings and compared those,
    /// which re-serialised each subtree to a string on every access and always built both strings
    /// in full — even when they differed at the first character. On a 2,000-paragraph contract with
    /// ~20 runs per paragraph that was tens of thousands of XML serialisations per comparison.
    /// </summary>
    private static bool SameRunFormatting(Paragraph original, Paragraph revised)
    {
        using var left = original.Descendants<RunProperties>().GetEnumerator();
        using var right = revised.Descendants<RunProperties>().GetEnumerator();

        while (true)
        {
            var leftHasMore = left.MoveNext();
            var rightHasMore = right.MoveNext();

            if (leftHasMore != rightHasMore)
                return false;
            if (!leftHasMore)
                return true;

            // Serialising a single rPr is cheap and bounded; the previous cost was concatenating
            // every one of them into a whole-paragraph string before any comparison happened.
            if (!string.Equals(left.Current.OuterXml, right.Current.OuterXml, StringComparison.Ordinal))
                return false;
        }
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

    /// <summary>
    /// Maps each revision to the heading it falls under, in document order.
    /// <para>
    /// Built as one forward pass that records, for every paragraph, the most recent heading seen so
    /// far — then each revision is an O(1) lookup. The obvious implementation walks backward from
    /// each revision instead, but <c>ElementsBeforeSelf()</c> on an <see cref="XElement"/>
    /// enumerates forward from the parent's first child, and <c>LastOrDefault()</c> forces that
    /// whole prefix to be enumerated, so stepping back a single paragraph costs O(position). Across
    /// R revisions in a P-paragraph document that is O(R·P²) — on a 200-page contract with a
    /// heavily negotiated redline, billions of traversal steps and a multi-minute hang.
    /// </para>
    /// </summary>
    private static IReadOnlyList<string> FindAffectedHeadings(IEnumerable<WmlComparer.WmlComparerRevision> revisions)
    {
        var revisionList = revisions.ToList();
        if (revisionList.Count == 0)
            return [];

        // Every revision in one comparison belongs to the same redlined document, so the heading
        // index only has to be built once, from whichever revision can point at the body.
        var body = revisionList
            .Select(r => (r.RevisionXElement ?? r.ContentXElement)?.AncestorsAndSelf(W + "body").FirstOrDefault())
            .FirstOrDefault(b => b is not null);

        if (body is null)
            return [];

        var headingByParagraph = new Dictionary<XElement, string>();
        string? currentHeading = null;
        foreach (var paragraph in body.Elements(W + "p"))
        {
            var styleId = paragraph.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val")?.Value;
            if (styleId is not null && styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase))
                currentHeading = string.Concat(paragraph.Descendants(W + "t").Select(t => t.Value));

            if (currentHeading is not null)
                headingByParagraph[paragraph] = currentHeading;
        }

        // Ordered set: callers get headings in document order, but membership testing stays O(1)
        // rather than the linear List.Contains scan this previously did per revision.
        var headings = new List<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);

        foreach (var revision in revisionList)
        {
            var element = revision.RevisionXElement ?? revision.ContentXElement;
            var paragraph = element?.AncestorsAndSelf(W + "p").FirstOrDefault();

            if (paragraph is not null && headingByParagraph.TryGetValue(paragraph, out var heading) && seen.Add(heading))
                headings.Add(heading);
        }

        return headings;
    }
}
