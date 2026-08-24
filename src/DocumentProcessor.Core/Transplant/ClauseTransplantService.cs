using Clippit.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.Diagnostics;
using DocumentProcessor.Core.DocumentAssembly;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;

namespace DocumentProcessor.Core.Transplant;

public sealed record ParagraphInfo(int Index, string Text);

/// <summary>
/// Copies a paragraph (or contiguous range of paragraphs — a clause/section) from one .docx file
/// into another, preserving its original formatting, styles, and numbering exactly, via Clippit's
/// DocumentBuilder. DocumentBuilder remaps style/numbering IDs between the source and target parts
/// so the transplanted content doesn't collide with or inherit the target's styles — the piece that
/// is otherwise unsolved hand-rolled work in comparable free tooling.
/// </summary>
public sealed class ClauseTransplantService(ILogger<ClauseTransplantService>? logger = null) : IClauseTransplantService
{
    private readonly ILogger<ClauseTransplantService> _logger = logger ?? NullLogger<ClauseTransplantService>.Instance;

    /// <summary>Lists every top-level paragraph in the document with its index, for locating a clause to transplant.</summary>
    public IReadOnlyList<ParagraphInfo> ListParagraphs(string docxPath)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: false);
        var body = doc.MainDocumentPart?.Document?.Body ?? throw new CorruptDocumentException("Document has no main part/body.");

        return body.Elements<Paragraph>()
            .Select((p, i) => new ParagraphInfo(i, p.InnerText))
            .ToList();
    }

    /// <summary>
    /// Copies paragraphs [<paramref name="sourceStartIndex"/>, <paramref name="sourceStartIndex"/> +
    /// <paramref name="paragraphCount"/>) from <paramref name="sourcePath"/> into
    /// <paramref name="targetPath"/>, inserted before target paragraph
    /// <paramref name="insertBeforeParagraphIndex"/>, writing the merged result to
    /// <paramref name="outputPath"/>. The source document is left untouched.
    /// </summary>
    public void TransplantParagraphs(
        string sourcePath, int sourceStartIndex, int paragraphCount,
        string targetPath, int insertBeforeParagraphIndex,
        string outputPath)
    {
        if (paragraphCount <= 0)
            throw new ArgumentOutOfRangeException(nameof(paragraphCount), "Must copy at least one paragraph.");

        var targetParagraphCount = ListParagraphs(targetPath).Count;
        if (insertBeforeParagraphIndex < 0 || insertBeforeParagraphIndex > targetParagraphCount)
        {
            throw new ArgumentOutOfRangeException(nameof(insertBeforeParagraphIndex),
                $"Target document has {targetParagraphCount} paragraphs; valid insertion points are 0..{targetParagraphCount}.");
        }

        var target = new WmlDocument(targetPath);
        var source = new WmlDocument(sourcePath);

        var sources = new List<ISource>();
        if (insertBeforeParagraphIndex > 0)
            sources.Add(new Source(target, 0, insertBeforeParagraphIndex, keepSections: false));

        sources.Add(new Source(source, sourceStartIndex, paragraphCount, keepSections: false));

        var remaining = targetParagraphCount - insertBeforeParagraphIndex;
        if (remaining > 0)
            sources.Add(new Source(target, insertBeforeParagraphIndex, remaining, keepSections: true));

        var merged = DocumentBuilder.BuildDocument(sources);
        merged.SaveAs(outputPath);
        _logger.LogInformation("Transplanted {Count} paragraph(s) from {SourcePath} into {TargetPath} -> {OutputPath}",
            paragraphCount, sourcePath, targetPath, outputPath);
    }

    /// <summary>
    /// Removes paragraphs [<paramref name="startIndex"/>, <paramref name="startIndex"/> +
    /// <paramref name="count"/>) from <paramref name="docxPath"/>, writing the result to
    /// <paramref name="outputPath"/> (which may be the same path, to edit in place).
    /// </summary>
    public void RemoveParagraphs(string docxPath, int startIndex, int count, string outputPath)
    {
        if (count <= 0)
            throw new ArgumentOutOfRangeException(nameof(count), "Must remove at least one paragraph.");

        var totalParagraphs = ListParagraphs(docxPath).Count;
        if (startIndex < 0 || startIndex + count > totalParagraphs)
        {
            throw new ArgumentOutOfRangeException(nameof(startIndex),
                $"Document has {totalParagraphs} paragraphs; can't remove [{startIndex}, {startIndex + count}).");
        }

        var target = new WmlDocument(docxPath);
        var sources = new List<ISource>();

        if (startIndex > 0)
            sources.Add(new Source(target, 0, startIndex, keepSections: false));

        var afterStart = startIndex + count;
        var afterCount = totalParagraphs - afterStart;
        if (afterCount > 0)
            sources.Add(new Source(target, afterStart, afterCount, keepSections: true));

        var merged = sources.Count > 0
            ? DocumentBuilder.BuildDocument(sources)
            : DocumentBuilder.BuildDocument([new Source(target, 0, 0, keepSections: true)]);
        merged.SaveAs(outputPath);
    }

    /// <summary>
    /// Removes paragraphs [<paramref name="replacedStartIndex"/>, <paramref name="replacedStartIndex"/> +
    /// <paramref name="replacedCount"/>) from <paramref name="targetPath"/> and inserts paragraphs
    /// [<paramref name="sourceStartIndex"/>, <paramref name="sourceStartIndex"/> +
    /// <paramref name="sourceParagraphCount"/>) from <paramref name="sourcePath"/> in their place —
    /// a clause swap in one step, rather than a separate remove then transplant.
    /// </summary>
    public void ReplaceParagraphs(
        string sourcePath, int sourceStartIndex, int sourceParagraphCount,
        string targetPath, int replacedStartIndex, int replacedCount,
        string outputPath)
    {
        if (sourceParagraphCount <= 0)
            throw new ArgumentOutOfRangeException(nameof(sourceParagraphCount), "Must copy at least one paragraph.");
        if (replacedCount <= 0)
            throw new ArgumentOutOfRangeException(nameof(replacedCount), "Must replace at least one paragraph.");

        var targetParagraphCount = ListParagraphs(targetPath).Count;
        if (replacedStartIndex < 0 || replacedStartIndex + replacedCount > targetParagraphCount)
        {
            throw new ArgumentOutOfRangeException(nameof(replacedStartIndex),
                $"Target document has {targetParagraphCount} paragraphs; can't replace [{replacedStartIndex}, {replacedStartIndex + replacedCount}).");
        }

        var target = new WmlDocument(targetPath);
        var source = new WmlDocument(sourcePath);
        var sources = new List<ISource>();

        if (replacedStartIndex > 0)
            sources.Add(new Source(target, 0, replacedStartIndex, keepSections: false));

        sources.Add(new Source(source, sourceStartIndex, sourceParagraphCount, keepSections: false));

        var afterStart = replacedStartIndex + replacedCount;
        var afterCount = targetParagraphCount - afterStart;
        if (afterCount > 0)
            sources.Add(new Source(target, afterStart, afterCount, keepSections: true));

        var merged = DocumentBuilder.BuildDocument(sources);
        merged.SaveAs(outputPath);
    }

    /// <summary>
    /// Removes paragraphs [<paramref name="startIndex"/>, <paramref name="startIndex"/> +
    /// <paramref name="count"/>), same as <see cref="RemoveParagraphs"/>, but first checks whether
    /// any bookmark defined inside that range is still referenced by a REF/PAGEREF field elsewhere
    /// in the document. Matches how Word itself behaves when you delete bookmarked content — the
    /// reference isn't rewritten or deleted automatically (there's no single correct replacement
    /// text to substitute), it's left in place to show "Error! Reference source not found." on the
    /// next field update. This just surfaces that outcome up front as a warning list, rather than
    /// leaving the caller to discover it only when someone opens the document in Word.
    /// </summary>
    /// <returns>Every reference that pointed into the removed range and is now dangling.</returns>
    public IReadOnlyList<DanglingReference> RemoveParagraphsWithCrossReferenceCleanup(string docxPath, int startIndex, int count, string outputPath)
    {
        var removedBookmarks = GetBookmarkNamesInRange(docxPath, startIndex, count);
        var referencesOutsideRange = GetReferencesOutsideRange(docxPath, startIndex, count);
        var newlyDangling = referencesOutsideRange.Where(r => removedBookmarks.Contains(r.BookmarkName)).ToList();

        RemoveParagraphs(docxPath, startIndex, count, outputPath);

        if (newlyDangling.Count > 0)
        {
            _logger.LogWarning(
                "Removing paragraphs [{Start}, {End}) from {DocxPath} left {Count} reference(s) dangling: {References}",
                startIndex, startIndex + count, docxPath, newlyDangling.Count, newlyDangling.Select(r => r.BookmarkName));
        }

        return newlyDangling;
    }

    private static HashSet<string> GetBookmarkNamesInRange(string docxPath, int startIndex, int count)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: false);
        var body = doc.MainDocumentPart?.Document?.Body ?? throw new CorruptDocumentException("Document has no main part/body.");
        var paragraphs = body.Elements<Paragraph>().ToList();
        var rangeEnd = Math.Min(startIndex + count, paragraphs.Count);

        var names = new HashSet<string>();
        for (var i = startIndex; i < rangeEnd; i++)
        {
            foreach (var bookmark in paragraphs[i].Descendants<BookmarkStart>())
            {
                if (bookmark.Name?.Value is { } name)
                    names.Add(name);
            }
        }

        return names;
    }

    private static IReadOnlyList<DanglingReference> GetReferencesOutsideRange(string docxPath, int startIndex, int count)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: false);
        var body = doc.MainDocumentPart?.Document?.Body ?? throw new CorruptDocumentException("Document has no main part/body.");
        var paragraphs = body.Elements<Paragraph>().ToList();
        var rangeEnd = Math.Min(startIndex + count, paragraphs.Count);

        var references = new List<DanglingReference>();
        for (var i = 0; i < paragraphs.Count; i++)
        {
            if (i >= startIndex && i < rangeEnd)
                continue; // inside the range being removed — not "elsewhere"

            references.AddRange(CrossReferenceValidator.FindReferences(paragraphs[i]));
        }

        return references;
    }

    /// <summary>
    /// Rewrites the numbering id of the first heading paragraph in the just-inserted range
    /// [<paramref name="insertedStartIndex"/>, <paramref name="insertedStartIndex"/> +
    /// <paramref name="insertedCount"/>) so it continues the numbering sequence of the nearest
    /// preceding paragraph sharing the same heading style — e.g. a transplanted "Section 7" clause
    /// picks up the target document's own section-numbering counter instead of arriving as an
    /// independent (and visibly wrong) "Section 1". No-ops if the inserted range has no heading
    /// paragraph, or no earlier paragraph shares its style.
    /// </summary>
    public void ContinueHeadingNumbering(string docxPath, int insertedStartIndex, int insertedCount)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var body = doc.MainDocumentPart?.Document?.Body ?? throw new CorruptDocumentException("Document has no main part/body.");
        var paragraphs = body.Elements<Paragraph>().ToList();

        var rangeEnd = Math.Min(insertedStartIndex + insertedCount, paragraphs.Count);
        for (var i = insertedStartIndex; i < rangeEnd; i++)
        {
            var candidate = paragraphs[i];
            var styleId = candidate.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            var numPr = candidate.ParagraphProperties?.NumberingProperties;
            if (styleId is null || numPr?.NumberingId is null)
                continue;

            var precedent = paragraphs.Take(insertedStartIndex).LastOrDefault(p =>
                p.ParagraphProperties?.ParagraphStyleId?.Val?.Value == styleId &&
                p.ParagraphProperties?.NumberingProperties?.NumberingId is not null);

            if (precedent is null)
                continue;

            var precedentNumId = precedent.ParagraphProperties!.NumberingProperties!.NumberingId!.Val!.Value;
            numPr.NumberingId!.Val = precedentNumId;
            break;
        }

        doc.MainDocumentPart!.Document.Save();
    }
}
