using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentProcessor.Core.Extraction;

/// <summary>One extracted block of document text, with enough structure to be useful downstream.</summary>
/// <param name="Index">0-based position among extracted blocks, in document order.</param>
/// <param name="Text">The block's text, with runs joined.</param>
/// <param name="StyleId">The paragraph style, when it has one — this is how headings are identified.</param>
/// <param name="HeadingLevel">1-9 for a Heading-styled paragraph, otherwise null.</param>
/// <param name="Heading">The nearest preceding heading's text, so a block carries the clause it sits under.</param>
/// <param name="IsTableContent">Whether the block came from inside a table cell.</param>
public sealed record TextBlock(
    int Index,
    string Text,
    string? StyleId,
    int? HeadingLevel,
    string? Heading,
    bool IsTableContent);

/// <summary>What to include when extracting.</summary>
public sealed record TextExtractionOptions
{
    /// <summary>Include text inside tables. On by default: in contracts, tables carry pricing,
    /// schedules and payment terms, which are exactly the terms people search for.</summary>
    public bool IncludeTables { get; init; } = true;

    /// <summary>Include header and footer text. Off by default — headers usually repeat boilerplate
    /// on every page, which pollutes a search index and skews relevance.</summary>
    public bool IncludeHeadersAndFooters { get; init; }

    /// <summary>
    /// Include text inside tracked deletions. Off by default: deleted text is not part of the
    /// document as it currently reads, so indexing it produces search hits for wording that was
    /// removed.
    /// </summary>
    public bool IncludeDeletedText { get; init; }

    /// <summary>Drop blocks that are empty or whitespace-only.</summary>
    public bool SkipEmpty { get; init; } = true;

    public static TextExtractionOptions Default { get; } = new();
}

/// <summary>
/// Extracts plain and structured text from a .docx.
/// <para>
/// A contract system needs document text as a first-class output — for full-text search, for
/// clause classification, and for handing passages to a model for review. Previously the only text
/// extraction in the library was a private helper inside the PDF comparison service, so every
/// consumer would have had to hand-roll this against the OOXML tree, and would likely have got the
/// tracked-changes case wrong: <c>InnerText</c> happily returns deleted text, so naive extraction
/// indexes wording that no longer appears in the document.
/// </para>
/// </summary>
public sealed class TextExtractionService : ITextExtractionService
{
    /// <summary>Extracts the document's text as a single string, one block per line.</summary>
    public string ExtractText(string docxPath, TextExtractionOptions? options = null)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: false);
        return ExtractTextCore(doc, options);
    }

    /// <summary>Runs against an already-open package, so a <see cref="Sessions.DocumentSession"/>
    /// pipeline pays one open/save for the whole sequence instead of one per call.</summary>
    internal static string ExtractTextCore(WordprocessingDocument doc, TextExtractionOptions? options = null) =>
        string.Join('\n', ExtractBlocksCore(doc, options).Select(b => b.Text));

    /// <summary>
    /// Extracts text as structured blocks, each tagged with its style and the heading it falls
    /// under — enough for clause-level indexing or for feeding a specific section to a model
    /// without sending the whole contract.
    /// </summary>
    public IReadOnlyList<TextBlock> ExtractBlocks(string docxPath, TextExtractionOptions? options = null)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: false);
        return ExtractBlocksCore(doc, options);
    }

    /// <summary>Runs against an already-open package.</summary>
    internal static IReadOnlyList<TextBlock> ExtractBlocksCore(WordprocessingDocument doc, TextExtractionOptions? options = null)
    {
        options ??= TextExtractionOptions.Default;

        var mainPart = doc.MainDocumentPart ?? throw CorruptDocumentException.MissingBody();
        var body = mainPart.Document?.Body ?? throw CorruptDocumentException.MissingBody();

        var blocks = new List<TextBlock>();
        string? currentHeading = null;

        foreach (var paragraph in EnumerateParagraphs(body, options.IncludeTables))
        {
            var text = ParagraphText(paragraph, options.IncludeDeletedText);
            var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            var headingLevel = HeadingLevelOf(styleId);

            if (headingLevel is not null)
                currentHeading = text;

            if (options.SkipEmpty && string.IsNullOrWhiteSpace(text))
                continue;

            blocks.Add(new TextBlock(
                blocks.Count,
                text,
                styleId,
                headingLevel,
                headingLevel is null ? currentHeading : null,
                IsTableContent: paragraph.Ancestors<Table>().Any()));
        }

        if (options.IncludeHeadersAndFooters)
        {
            foreach (var part in mainPart.HeaderParts.Cast<OpenXmlPart>().Concat(mainPart.FooterParts))
            {
                foreach (var paragraph in part.RootElement?.Descendants<Paragraph>() ?? [])
                {
                    var text = ParagraphText(paragraph, options.IncludeDeletedText);
                    if (options.SkipEmpty && string.IsNullOrWhiteSpace(text))
                        continue;

                    blocks.Add(new TextBlock(blocks.Count, text, null, null, null, IsTableContent: false));
                }
            }
        }

        return blocks;
    }

    /// <summary>
    /// Body paragraphs in document order, descending into tables only when asked. Table paragraphs
    /// are yielded in place rather than appended afterwards, so extracted text preserves the order
    /// a reader would encounter it in.
    /// </summary>
    private static IEnumerable<Paragraph> EnumerateParagraphs(Body body, bool includeTables)
    {
        foreach (var element in body.Elements())
        {
            switch (element)
            {
                case Paragraph paragraph:
                    yield return paragraph;
                    break;

                case Table table when includeTables:
                    foreach (var paragraph in table.Descendants<Paragraph>())
                        yield return paragraph;
                    break;
            }
        }
    }

    /// <summary>
    /// Joins a paragraph's run text. Deliberately not <c>InnerText</c>: that includes
    /// <c>w:delText</c>, so a document with unresolved tracked changes would extract both the old
    /// and new wording concatenated together — text that never appears in the document as read.
    /// </summary>
    private static string ParagraphText(Paragraph paragraph, bool includeDeletedText)
    {
        var builder = new StringBuilder();

        foreach (var element in paragraph.Descendants())
        {
            switch (element)
            {
                case Text text when includeDeletedText || !IsInsideDeletion(text):
                    builder.Append(text.Text);
                    break;

                case DeletedText deleted when includeDeletedText:
                    builder.Append(deleted.Text);
                    break;

                case TabChar:
                    builder.Append('\t');
                    break;

                case Break:
                    builder.Append(' ');
                    break;
            }
        }

        return builder.ToString().Trim();
    }

    private static bool IsInsideDeletion(OpenXmlElement element) =>
        element.Ancestors<DeletedRun>().Any();

    /// <summary>Maps Heading1..Heading9 to 1..9. Word's built-in style ids are unlocalised, so a
    /// simple prefix check is reliable across language versions.</summary>
    private static int? HeadingLevelOf(string? styleId)
    {
        if (styleId is null || !styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase))
            return null;

        var suffix = styleId["Heading".Length..];
        return int.TryParse(suffix, out var level) && level is >= 1 and <= 9 ? level : null;
    }
}
