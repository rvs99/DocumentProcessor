using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentProcessor.Core.Templating;

/// <summary>
/// A .docx whose clauses are individually addressable by id, for <c>{{clause:id}}</c> markers. A
/// clause is delimited by a bookmark named <c>clause_&lt;id&gt;</c> wrapping the paragraph(s) that
/// make up that clause — the same bookmark convention Word's own "Insert Bookmark" feature produces,
/// so a legal team can mark up a clause library docx in Word itself with no special tooling.
/// </summary>
public sealed record ClauseLibrary(string DocxPath)
{
    /// <summary>Returns the (0-based start paragraph index, paragraph count) spanned by the clause
    /// named <paramref name="clauseId"/>, or <see langword="null"/> if no bookmark named
    /// <c>clause_&lt;clauseId&gt;</c> exists in this library.</summary>
    public (int StartIndex, int Count)? FindClause(string clauseId)
    {
        using var doc = WordprocessingDocument.Open(DocxPath, isEditable: false);
        var body = doc.MainDocumentPart?.Document?.Body
            ?? throw new CorruptDocumentException("Clause library document has no body.");

        var bookmarkName = $"clause_{clauseId}";
        var start = body.Descendants<BookmarkStart>().FirstOrDefault(b => b.Name?.Value == bookmarkName);
        if (start is null)
            return null;

        var end = body.Descendants<BookmarkEnd>().FirstOrDefault(b => b.Id?.Value == start.Id?.Value);
        if (end is null)
            return null;

        var paragraphs = body.Elements<Paragraph>().ToList();
        var startParagraph = start.Ancestors<Paragraph>().FirstOrDefault();
        var endParagraph = end.Ancestors<Paragraph>().FirstOrDefault();
        if (startParagraph is null || endParagraph is null)
            return null;

        var startIndex = paragraphs.IndexOf(startParagraph);
        var endIndex = paragraphs.IndexOf(endParagraph);
        if (startIndex < 0 || endIndex < 0 || endIndex < startIndex)
            return null;

        return (startIndex, endIndex - startIndex + 1);
    }

    /// <summary>Lists every clause id defined in this library, for discovery/validation.</summary>
    public IReadOnlyList<string> ListClauseIds()
    {
        using var doc = WordprocessingDocument.Open(DocxPath, isEditable: false);
        var body = doc.MainDocumentPart?.Document?.Body
            ?? throw new CorruptDocumentException("Clause library document has no body.");

        const string prefix = "clause_";
        return body.Descendants<BookmarkStart>()
            .Select(b => b.Name?.Value)
            .Where(name => name is not null && name.StartsWith(prefix, StringComparison.Ordinal))
            .Select(name => name![prefix.Length..])
            .ToList();
    }
}
