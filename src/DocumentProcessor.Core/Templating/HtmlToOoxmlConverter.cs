using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentProcessor.Core.Templating;

/// <summary>
/// An HTML fragment was rejected before processing because it was too large or too deeply nested.
/// Distinct from a malformed-content error: the fragment may be perfectly well-formed, it is simply
/// beyond the bounds this converter will walk without risking the host process.
/// </summary>
public sealed class HtmlTooComplexException(string message) : InvalidOperationException(message);

public sealed record HtmlConversionOptions
{
    /// <summary>Render <c>&lt;a href&gt;</c> as underlined/colored text with the URL appended in
    /// parentheses. Real <c>w:hyperlink</c> relationships aren't created — v1 scope is visual fidelity
    /// for review, not click-through links.</summary>
    public bool RenderHyperlinkUrls { get; init; } = true;
}

/// <summary>
/// Converts a sanitized HTML fragment (as would come from a rich-text editor) into OOXML paragraphs —
/// for injecting rich-text values into a content control or template token, without trusting the
/// fragment's markup as-is. Sanitization is allow-list based: only a fixed set of formatting tags
/// survive (p/div/span/h1-6/b/strong/i/em/u/ul/ol/li/a/br), every attribute is dropped except
/// <c>href</c> on <c>&lt;a&gt;</c> (and only when it isn't a <c>javascript:</c> URL), and
/// <c>&lt;script&gt;</c>/<c>&lt;style&gt;</c>/<c>&lt;iframe&gt;</c>/<c>&lt;object&gt;</c>/
/// <c>&lt;embed&gt;</c>/<c>&lt;form&gt;</c>/<c>&lt;input&gt;</c>/<c>&lt;button&gt;</c>/<c>&lt;svg&gt;</c>
/// are removed along with their entire subtree — event-handler attributes (<c>onclick</c> etc.) are
/// stripped as a side effect of the attribute allow-list, not detected by name.
/// </summary>
public static class HtmlToOoxmlConverter
{
    private static readonly HashSet<string> KnownTags = new(StringComparer.OrdinalIgnoreCase)
    {
        "p", "div", "span", "br", "b", "strong", "i", "em", "u",
        "ul", "ol", "li", "a", "h1", "h2", "h3", "h4", "h5", "h6"
    };

    private static readonly HashSet<string> DropWithSubtree = new(StringComparer.OrdinalIgnoreCase)
    {
        "script", "style", "iframe", "object", "embed", "form", "input", "button", "svg", "textarea", "select"
    };

    /// <summary>
    /// Maximum element nesting this converter will walk. Both the sanitizer and the OOXML builder
    /// recurse once per level, so without a cap a deeply nested fragment exhausts the thread stack —
    /// and <see cref="StackOverflowException"/> cannot be caught in .NET, so the whole process dies.
    /// In a multi-tenant host that is one tenant's malformed rich-text field terminating every other
    /// tenant's in-flight request. 100 is far beyond any realistic authored document (Word itself
    /// caps nesting well below this) while staying far below the ~1 MB default stack.
    /// </summary>
    public const int MaxNestingDepth = 100;

    /// <summary>Maximum input length accepted, in characters. Rich-text field values are prose, not
    /// documents; anything larger is a payload rather than content.</summary>
    public const int MaxInputLength = 1024 * 1024;

    /// <summary>Sanitizes an HTML fragment down to the allow-listed tags/attributes and returns it as
    /// an HTML string, without converting to OOXML. Used for item 33's standalone sanitization need.</summary>
    /// <exception cref="HtmlTooComplexException">The fragment is longer than
    /// <see cref="MaxInputLength"/> or nested deeper than <see cref="MaxNestingDepth"/>.</exception>
    public static string Sanitize(string html)
    {
        GuardInputLength(html);
        var document = ParseFragment(html);
        SanitizeSubtree(document.Body!, depth: 0);
        return document.Body!.InnerHtml;
    }

    private static void GuardInputLength(string html)
    {
        if (html.Length > MaxInputLength)
            throw new HtmlTooComplexException($"HTML fragment is {html.Length:N0} characters; the limit is {MaxInputLength:N0}.");
    }

    private static void GuardDepth(int depth)
    {
        if (depth > MaxNestingDepth)
            throw new HtmlTooComplexException($"HTML nesting exceeds the maximum depth of {MaxNestingDepth}.");
    }

    /// <summary>
    /// Converts <paramref name="html"/> into a sequence of <see cref="Paragraph"/> elements ready to
    /// insert into a document body or content-control block. <paramref name="mainPart"/> is needed
    /// only when the fragment contains list markup (<c>ul</c>/<c>ol</c>) — list paragraphs reference a
    /// numbering definition that must live in that document's <see cref="NumberingDefinitionsPart"/>.
    /// </summary>
    public static IReadOnlyList<Paragraph> ConvertFragment(MainDocumentPart mainPart, string html, HtmlConversionOptions? options = null)
    {
        options ??= new HtmlConversionOptions();
        GuardInputLength(html);
        var document = ParseFragment(html);
        SanitizeSubtree(document.Body!, depth: 0);

        var state = new ConversionState(mainPart, options);
        var paragraphs = new List<Paragraph>();
        foreach (var child in document.Body!.ChildNodes)
            WalkBlock(child, state, paragraphs, listLevel: null, depth: 0);

        if (paragraphs.Count == 0)
            paragraphs.Add(new Paragraph());

        return paragraphs;
    }

    private static IHtmlDocument ParseFragment(string html)
    {
        var parser = new HtmlParser();
        return parser.ParseDocument($"<!doctype html><html><body>{html}</body></html>");
    }

    private static void SanitizeSubtree(INode node, int depth)
    {
        GuardDepth(depth);

        foreach (var child in node.ChildNodes.ToList())
        {
            if (child is not IElement element)
                continue; // text nodes and comments pass through untouched (comments are inert)

            if (DropWithSubtree.Contains(element.TagName))
            {
                element.Remove();
                continue;
            }

            if (!KnownTags.Contains(element.TagName))
            {
                // Unknown/disallowed tag: unwrap it (keep children, drop the wrapper) rather than
                // dropping content the caller probably still wants shown.
                SanitizeSubtree(element, depth + 1);
                foreach (var grandchild in element.ChildNodes.ToList())
                    element.Parent!.InsertBefore(grandchild, element);
                element.Remove();
                continue;
            }

            foreach (var attrName in element.Attributes.Select(a => a.Name).ToList())
            {
                if (string.Equals(element.TagName, "a", StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(attrName, "href", StringComparison.OrdinalIgnoreCase))
                {
                    var href = element.GetAttribute("href") ?? "";
                    if (IsSafeUrl(href))
                        continue;
                }

                element.RemoveAttribute(attrName);
            }

            SanitizeSubtree(element, depth + 1);
        }
    }

    private static bool IsSafeUrl(string url)
    {
        var normalized = new string(url.Where(c => !char.IsControl(c) && !char.IsWhiteSpace(c)).ToArray()).TrimStart();
        return !normalized.StartsWith("javascript:", StringComparison.OrdinalIgnoreCase) &&
               !normalized.StartsWith("data:", StringComparison.OrdinalIgnoreCase) &&
               !normalized.StartsWith("vbscript:", StringComparison.OrdinalIgnoreCase);
    }

    private sealed class ConversionState(MainDocumentPart mainPart, HtmlConversionOptions options)
    {
        public MainDocumentPart MainPart { get; } = mainPart;
        public HtmlConversionOptions Options { get; } = options;
        public int? BulletNumId;
        public int? DecimalNumId;
    }

    private sealed record ListContext(int NumId, int Level);

    private static void WalkBlock(INode node, ConversionState state, List<Paragraph> output, ListContext? listLevel, int depth)
    {
        GuardDepth(depth);

        if (node is IText textNode)
        {
            if (!string.IsNullOrWhiteSpace(textNode.Text))
                output.Add(new Paragraph(BuildRuns(textNode.Text, RunFormat.None)));
            return;
        }

        if (node is not IElement element)
            return;

        switch (element.TagName.ToLowerInvariant())
        {
            case "p":
            case "div":
                output.Add(BuildParagraph(element, state, listLevel, depth));
                break;

            case "h1": case "h2": case "h3": case "h4": case "h5": case "h6":
                var level = element.TagName[1] - '0';
                output.Add(BuildParagraph(element, state, listLevel, depth, headingStyleId: $"Heading{level}"));
                break;

            case "ul":
            case "ol":
                var numId = element.TagName.Equals("ul", StringComparison.OrdinalIgnoreCase)
                    ? EnsureBulletNumbering(state)
                    : EnsureDecimalNumbering(state);
                var nextLevel = new ListContext(numId, (listLevel?.Level ?? -1) + 1);
                foreach (var child in element.Children.Where(c => c.TagName.Equals("li", StringComparison.OrdinalIgnoreCase)))
                    output.Add(BuildParagraph(child, state, nextLevel, depth + 1));
                break;

            default:
                // Bare inline content at block level (e.g. top-level "Hello <b>World</b>" with no <p>).
                output.Add(BuildParagraph(element, state, listLevel, depth));
                break;
        }
    }

    private static Paragraph BuildParagraph(IElement element, ConversionState state, ListContext? listLevel, int depth, string? headingStyleId = null)
    {
        var properties = new ParagraphProperties();
        if (headingStyleId is not null)
            properties.AppendChild(new ParagraphStyleId { Val = headingStyleId });
        if (listLevel is not null)
        {
            properties.AppendChild(new NumberingProperties(
                new NumberingLevelReference { Val = listLevel.Level },
                new NumberingId { Val = listLevel.NumId }));
        }

        var paragraph = new Paragraph();
        if (properties.HasChildren)
            paragraph.AppendChild(properties);

        foreach (var run in BuildInlineRuns(element, RunFormat.None, state, depth + 1))
            paragraph.AppendChild(run);

        return paragraph;
    }

    [Flags]
    private enum RunFormat { None = 0, Bold = 1, Italic = 2, Underline = 4 }

    private static IEnumerable<Run> BuildInlineRuns(INode node, RunFormat format, ConversionState state, int depth)
    {
        GuardDepth(depth);

        foreach (var child in node.ChildNodes)
        {
            if (child is IText text)
            {
                if (text.Text.Length > 0)
                    foreach (var run in BuildRuns(text.Text, format))
                        yield return run;
                continue;
            }

            if (child is not IElement element)
                continue;

            switch (element.TagName.ToLowerInvariant())
            {
                case "br":
                    yield return new Run(new Break());
                    break;
                case "b": case "strong":
                    foreach (var run in BuildInlineRuns(element, format | RunFormat.Bold, state, depth + 1)) yield return run;
                    break;
                case "i": case "em":
                    foreach (var run in BuildInlineRuns(element, format | RunFormat.Italic, state, depth + 1)) yield return run;
                    break;
                case "u":
                    foreach (var run in BuildInlineRuns(element, format | RunFormat.Underline, state, depth + 1)) yield return run;
                    break;
                case "a":
                    var href = element.GetAttribute("href");
                    foreach (var run in BuildInlineRuns(element, format | RunFormat.Underline, state, depth + 1)) yield return run;
                    if (state.Options.RenderHyperlinkUrls && !string.IsNullOrEmpty(href))
                        foreach (var run in BuildRuns($" ({href})", format))
                            yield return run;
                    break;
                default:
                    foreach (var run in BuildInlineRuns(element, format, state, depth + 1)) yield return run;
                    break;
            }
        }
    }

    private static IEnumerable<Run> BuildRuns(string text, RunFormat format)
    {
        var props = BuildRunProperties(format);
        yield return props is not null
            ? new Run(props, new Text(text) { Space = SpaceProcessingModeValues.Preserve })
            : new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
    }

    private static RunProperties? BuildRunProperties(RunFormat format)
    {
        if (format == RunFormat.None)
            return null;

        var props = new RunProperties();
        if (format.HasFlag(RunFormat.Bold)) props.AppendChild(new Bold());
        if (format.HasFlag(RunFormat.Italic)) props.AppendChild(new Italic());
        if (format.HasFlag(RunFormat.Underline)) props.AppendChild(new Underline { Val = UnderlineValues.Single });
        return props;
    }

    private static int EnsureBulletNumbering(ConversionState state)
    {
        state.BulletNumId ??= NumberingDefinitions.EnsureListDefinition(state.MainPart, NumberFormatValues.Bullet, "");
        return state.BulletNumId.Value;
    }

    private static int EnsureDecimalNumbering(ConversionState state)
    {
        state.DecimalNumId ??= NumberingDefinitions.EnsureListDefinition(state.MainPart, NumberFormatValues.Decimal, "%1.");
        return state.DecimalNumId.Value;
    }
}

/// <summary>Allocates a minimal single-level abstract/concrete numbering definition pair in a
/// document's <see cref="NumberingDefinitionsPart"/>, creating the part if it doesn't exist yet.</summary>
internal static class NumberingDefinitions
{
    public static int EnsureListDefinition(MainDocumentPart mainPart, NumberFormatValues format, string levelText)
    {
        var numberingPart = mainPart.NumberingDefinitionsPart ?? mainPart.AddNewPart<NumberingDefinitionsPart>();
        numberingPart.Numbering ??= new Numbering();
        var numbering = numberingPart.Numbering;

        var nextAbstractId = numbering.Elements<AbstractNum>().Select(a => a.AbstractNumberId?.Value ?? 0).DefaultIfEmpty(-1).Max() + 1;
        var nextNumId = numbering.Elements<NumberingInstance>().Select(n => n.NumberID?.Value ?? 0).DefaultIfEmpty(0).Max() + 1;

        var abstractNum = new AbstractNum(
            new Level(
                new StartNumberingValue { Val = 1 },
                new NumberingFormat { Val = format },
                new LevelText { Val = levelText },
                new LevelJustification { Val = LevelJustificationValues.Left },
                new PreviousParagraphProperties(new Indentation { Left = "720", Hanging = "360" }))
            { LevelIndex = 0 })
        { AbstractNumberId = nextAbstractId };

        numbering.AppendChild(abstractNum);
        numbering.AppendChild(new NumberingInstance(new AbstractNumId { Val = nextAbstractId }) { NumberID = nextNumId });
        numberingPart.Numbering.Save();

        return nextNumId;
    }
}
