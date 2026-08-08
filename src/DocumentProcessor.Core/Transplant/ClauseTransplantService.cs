using Clippit.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentProcessor.Core.Transplant;

public sealed record ParagraphInfo(int Index, string Text);

/// <summary>
/// Copies a paragraph (or contiguous range of paragraphs — a clause/section) from one .docx file
/// into another, preserving its original formatting, styles, and numbering exactly, via Clippit's
/// DocumentBuilder. DocumentBuilder remaps style/numbering IDs between the source and target parts
/// so the transplanted content doesn't collide with or inherit the target's styles — the piece that
/// is otherwise unsolved hand-rolled work in comparable free tooling.
/// </summary>
public sealed class ClauseTransplantService
{
    /// <summary>Lists every top-level paragraph in the document with its index, for locating a clause to transplant.</summary>
    public IReadOnlyList<ParagraphInfo> ListParagraphs(string docxPath)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: false);
        var body = doc.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Document has no main part/body.");

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
    }
}
