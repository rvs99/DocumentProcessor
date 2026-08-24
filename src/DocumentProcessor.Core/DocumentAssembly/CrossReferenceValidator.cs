using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentProcessor.Core.DocumentAssembly;

/// <summary>A <c>REF</c>/<c>PAGEREF</c> field whose target bookmark doesn't exist in the document —
/// what Word itself would show as "Error! Reference source not found." on the next field update.</summary>
public sealed record DanglingReference(string BookmarkName, string FieldType);

/// <summary>
/// Checks whether every <c>REF</c>/<c>PAGEREF</c> cross-reference field in a document still points
/// at a bookmark that exists — covering both field forms Word produces (<c>w:fldSimple</c>, the
/// common case from Insert &gt; Cross-reference, and the <c>w:fldChar</c>/<c>w:instrText</c> triplet
/// Word converts to after certain edits).
/// </summary>
public sealed class CrossReferenceValidator : ICrossReferenceValidator
{
    private static readonly Regex ReferencePattern = new(@"\b(REF|PAGEREF)\s+(\S+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public IReadOnlyList<DanglingReference> Validate(string docxPath)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: false);
        return ValidateCore(doc);
    }

    /// <summary>Runs against an already-open package, so a <see cref="Sessions.DocumentSession"/>
    /// pipeline pays one open/save for the whole sequence instead of one per call.</summary>
    internal static IReadOnlyList<DanglingReference> ValidateCore(WordprocessingDocument doc)
    {        var body = doc.MainDocumentPart?.Document?.Body ?? throw new CorruptDocumentException("Document has no main part/body.");

        var bookmarkNames = new HashSet<string>(
            body.Descendants<BookmarkStart>().Select(b => b.Name?.Value).Where(n => n is not null)!);

        return FindReferences(body)
            .Where(r => !bookmarkNames.Contains(r.BookmarkName))
            .ToList();
}

    /// <summary>Every REF/PAGEREF reference within <paramref name="scope"/>, valid or not — used by
    /// callers (e.g. clause removal) that need to compare before/after a structural edit rather than
    /// just the current state.</summary>
    internal static IReadOnlyList<DanglingReference> FindReferences(DocumentFormat.OpenXml.OpenXmlElement scope)
    {
        var references = new List<DanglingReference>();

        foreach (var simpleField in scope.Descendants<SimpleField>())
            AddIfReference(references, simpleField.Instruction?.Value);

        foreach (var fieldCode in scope.Descendants<FieldCode>())
            AddIfReference(references, fieldCode.Text);

        return references;
    }

    private static void AddIfReference(List<DanglingReference> references, string? instruction)
    {
        if (instruction is null)
            return;

        var match = ReferencePattern.Match(instruction);
        if (match.Success)
            references.Add(new DanglingReference(match.Groups[2].Value, match.Groups[1].Value.ToUpperInvariant()));
    }
}
