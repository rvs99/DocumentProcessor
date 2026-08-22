using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.Transplant;

namespace DocumentProcessor.Core.Templating;

/// <summary>
/// Fills a .docx template's <c>{{token}}</c> markers from a data dictionary — scalar substitution,
/// rich-text (HTML) injection, <c>{{if:...}}</c> conditional sections, <c>{{repeat:...}}</c>
/// repeating sections, and <c>{{clause:id}}</c> clause-library injection — leaving the template file
/// itself untouched and writing the filled result to a new file.
///
/// Token scanning is run-merged: Word routinely splits a paragraph's text across several
/// <c>w:r</c>/<c>w:t</c> runs (spellcheck/grammar/revision boundaries), so a token like
/// <c>{{ClientName}}</c> can straddle 2-3 runs in a real document. See <see cref="RunTextScanner"/>
/// for how matches are found and spliced back without losing surrounding run formatting.
///
/// Block markers (<c>{{if:...}}</c>/<c>{{else}}</c>/<c>{{/if}}</c>,
/// <c>{{repeat:...}}</c>/<c>{{/repeat}}</c>, <c>{{clause:id}}</c>) must each be the entire text of
/// their own paragraph — the same convention mail-merge tools like docxtemplater use, and necessary
/// because these operate on whole paragraph ranges, not inline text spans.
/// </summary>
public sealed class TemplateEngine
{
    private static readonly Regex InlineTokenPattern = new(@"\{\{(html:)?([A-Za-z0-9_.]+)\}\}", RegexOptions.Compiled);
    private static readonly Regex IfPattern = new(@"^\{\{if:(.+)\}\}$", RegexOptions.Compiled);
    private static readonly Regex ElsePattern = new(@"^\{\{else\}\}$", RegexOptions.Compiled);
    private static readonly Regex EndIfPattern = new(@"^\{\{/if\}\}$", RegexOptions.Compiled);
    private static readonly Regex RepeatPattern = new(@"^\{\{repeat:(.+)\}\}$", RegexOptions.Compiled);
    private static readonly Regex EndRepeatPattern = new(@"^\{\{/repeat\}\}$", RegexOptions.Compiled);
    private static readonly Regex ClausePattern = new(@"^\{\{clause:(.+)\}\}$", RegexOptions.Compiled);

    /// <summary>
    /// Fills <paramref name="templatePath"/> using <paramref name="data"/> and writes the result to
    /// <paramref name="outputPath"/>. Collection values for <c>{{repeat:...}}</c> must be
    /// <see cref="IEnumerable{T}"/> of <c>IReadOnlyDictionary&lt;string, object?&gt;</c> — one
    /// dictionary per row/item.
    /// </summary>
    public void Fill(
        string templatePath,
        string outputPath,
        IReadOnlyDictionary<string, object?> data,
        MissingTokenPolicy missingTokenPolicy = MissingTokenPolicy.Error,
        ClauseLibrary? clauseLibrary = null)
    {
        File.Copy(templatePath, outputPath, overwrite: true);

        using (var doc = WordprocessingDocument.Open(outputPath, isEditable: true))
        {
            var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Document has no main part.");
            var document = mainPart.Document ?? throw new InvalidOperationException("Document has no body.");
            var body = document.Body ?? throw new InvalidOperationException("Document has no body.");

            var paragraphs = body.Elements<Paragraph>().ToList();
            var index = 0;
            var nodes = ParseSequence(paragraphs, ref index);
            if (index != paragraphs.Count)
                throw new FormatException("Unbalanced {{if}}/{{repeat}} markers in template.");

            ExpandNodes(nodes, new TemplateContext(data), mainPart, missingTokenPolicy);

            document.Save();
        }

        if (clauseLibrary is not null)
            ResolveClauseMarkers(outputPath, clauseLibrary, data, missingTokenPolicy);
    }

    // ---- Block structure parsing -------------------------------------------------------------

    private abstract record Node;
    private sealed record LiteralNode(List<Paragraph> Paragraphs) : Node;
    private sealed record IfNode(TemplateCondition Condition, List<Node> Then, List<Node> Else, Paragraph IfMarker, Paragraph? ElseMarker, Paragraph EndMarker) : Node;
    private sealed record RepeatNode(string CollectionPath, List<Node> Body, Paragraph StartMarker, Paragraph EndMarker) : Node;
    private sealed record ClauseNode(string ClauseId, Paragraph Marker) : Node;

    private static List<Node> ParseSequence(List<Paragraph> paragraphs, ref int index, bool stopAtElseOrEnd = false)
    {
        var nodes = new List<Node>();
        List<Paragraph>? literalRun = null;

        while (index < paragraphs.Count)
        {
            var text = paragraphs[index].InnerText.Trim();

            if (stopAtElseOrEnd && (ElsePattern.IsMatch(text) || EndIfPattern.IsMatch(text) || EndRepeatPattern.IsMatch(text)))
                break;

            var ifMatch = IfPattern.Match(text);
            if (ifMatch.Success)
            {
                FlushLiteral(nodes, ref literalRun);
                var ifMarker = paragraphs[index];
                index++;
                var thenBranch = ParseSequence(paragraphs, ref index, stopAtElseOrEnd: true);

                Paragraph? elseMarker = null;
                var elseBranch = new List<Node>();
                if (index < paragraphs.Count && ElsePattern.IsMatch(paragraphs[index].InnerText.Trim()))
                {
                    elseMarker = paragraphs[index];
                    index++;
                    elseBranch = ParseSequence(paragraphs, ref index, stopAtElseOrEnd: true);
                }

                if (index >= paragraphs.Count || !EndIfPattern.IsMatch(paragraphs[index].InnerText.Trim()))
                    throw new FormatException($"Missing {{{{/if}}}} for {{{{if:{ifMatch.Groups[1].Value}}}}}.");
                var endMarker = paragraphs[index];
                index++;

                nodes.Add(new IfNode(TemplateCondition.Parse(ifMatch.Groups[1].Value), thenBranch, elseBranch, ifMarker, elseMarker, endMarker));
                continue;
            }

            var repeatMatch = RepeatPattern.Match(text);
            if (repeatMatch.Success)
            {
                FlushLiteral(nodes, ref literalRun);
                var startMarker = paragraphs[index];
                index++;
                var bodyNodes = ParseSequence(paragraphs, ref index, stopAtElseOrEnd: true);

                if (index >= paragraphs.Count || !EndRepeatPattern.IsMatch(paragraphs[index].InnerText.Trim()))
                    throw new FormatException($"Missing {{{{/repeat}}}} for {{{{repeat:{repeatMatch.Groups[1].Value}}}}}.");
                var endMarker = paragraphs[index];
                index++;

                nodes.Add(new RepeatNode(repeatMatch.Groups[1].Value, bodyNodes, startMarker, endMarker));
                continue;
            }

            var clauseMatch = ClausePattern.Match(text);
            if (clauseMatch.Success)
            {
                FlushLiteral(nodes, ref literalRun);
                nodes.Add(new ClauseNode(clauseMatch.Groups[1].Value, paragraphs[index]));
                index++;
                continue;
            }

            (literalRun ??= []).Add(paragraphs[index]);
            index++;
        }

        FlushLiteral(nodes, ref literalRun);
        return nodes;
    }

    private static void FlushLiteral(List<Node> nodes, ref List<Paragraph>? literalRun)
    {
        if (literalRun is { Count: > 0 })
            nodes.Add(new LiteralNode(literalRun));
        literalRun = null;
    }

    private static List<Paragraph> CollectParagraphsInOrder(IEnumerable<Node> nodes)
    {
        var result = new List<Paragraph>();
        foreach (var node in nodes)
        {
            switch (node)
            {
                case LiteralNode literal:
                    result.AddRange(literal.Paragraphs);
                    break;
                case IfNode ifNode:
                    result.Add(ifNode.IfMarker);
                    result.AddRange(CollectParagraphsInOrder(ifNode.Then));
                    if (ifNode.ElseMarker is not null) result.Add(ifNode.ElseMarker);
                    result.AddRange(CollectParagraphsInOrder(ifNode.Else));
                    result.Add(ifNode.EndMarker);
                    break;
                case RepeatNode repeatNode:
                    result.Add(repeatNode.StartMarker);
                    result.AddRange(CollectParagraphsInOrder(repeatNode.Body));
                    result.Add(repeatNode.EndMarker);
                    break;
                case ClauseNode clauseNode:
                    result.Add(clauseNode.Marker);
                    break;
            }
        }
        return result;
    }

    // ---- Expansion (direct in-place OOXML mutation) ------------------------------------------

    private static void ExpandNodes(IReadOnlyList<Node> nodes, TemplateContext context, MainDocumentPart mainPart, MissingTokenPolicy policy)
    {
        foreach (var node in nodes)
        {
            switch (node)
            {
                case LiteralNode literal:
                    foreach (var paragraph in literal.Paragraphs)
                        SubstituteInline(paragraph, context, mainPart, policy);
                    break;

                case IfNode ifNode:
                    context.TryResolve(ifNode.Condition.FieldPath, out var actual);
                    var chosen = ifNode.Condition.Evaluate(actual) ? ifNode.Then : ifNode.Else;
                    var discarded = ReferenceEquals(chosen, ifNode.Then) ? ifNode.Else : ifNode.Then;

                    ifNode.IfMarker.Remove();
                    ifNode.ElseMarker?.Remove();
                    ifNode.EndMarker.Remove();
                    foreach (var paragraph in CollectParagraphsInOrder(discarded))
                        paragraph.Remove();

                    ExpandNodes(chosen, context, mainPart, policy);
                    break;

                case RepeatNode repeatNode:
                    ExpandRepeat(repeatNode, context, mainPart, policy);
                    break;

                case ClauseNode:
                    // Resolved in a second pass (ResolveClauseMarkers) after this document is saved —
                    // clause content lives in a separate library docx and needs Clippit's
                    // cross-document style/numbering remapping, which requires file paths.
                    break;
            }
        }
    }

    private static void ExpandRepeat(RepeatNode repeatNode, TemplateContext context, MainDocumentPart mainPart, MissingTokenPolicy policy)
    {
        var found = context.TryResolve(repeatNode.CollectionPath, out var rawCollection);
        if (!found && policy == MissingTokenPolicy.Error)
            throw new MissingTemplateTokenException(repeatNode.CollectionPath);

        var items = rawCollection as IEnumerable<IReadOnlyDictionary<string, object?>> ?? [];
        var templateParagraphs = CollectParagraphsInOrder(repeatNode.Body);

        foreach (var item in items)
        {
            var clones = templateParagraphs.Select(p => (Paragraph)p.CloneNode(true)).ToList();
            foreach (var clone in clones)
                repeatNode.StartMarker.InsertBeforeSelf(clone);

            var cloneIndex = 0;
            var childNodes = ParseSequence(clones, ref cloneIndex);
            ExpandNodes(childNodes, context.Push(item), mainPart, policy);
        }

        foreach (var paragraph in templateParagraphs)
            paragraph.Remove();
        repeatNode.StartMarker.Remove();
        repeatNode.EndMarker.Remove();
    }

    // ---- Inline scalar / HTML substitution ------------------------------------------------------

    /// <summary>Applies inline <c>{{token}}</c>/<c>{{html:token}}</c> substitution to a single
    /// paragraph. Exposed beyond <see cref="Fill"/> so other services needing the same run-merged
    /// substitution primitive (e.g. <see cref="Tables.TableGenerationService"/>'s prototype-row
    /// population) don't have to reimplement it.</summary>
    internal static void SubstituteInline(Paragraph paragraph, TemplateContext context, MainDocumentPart mainPart, MissingTokenPolicy policy)
    {
        var trimmed = paragraph.InnerText.Trim();
        var wholeParagraphHtmlMatch = Regex.Match(trimmed, @"^\{\{html:([A-Za-z0-9_.]+)\}\}$");
        if (wholeParagraphHtmlMatch.Success)
        {
            SubstituteWholeParagraphHtml(paragraph, wholeParagraphHtmlMatch.Groups[1].Value, context, mainPart, policy);
            return;
        }

        // Find every token position in one pass, then splice back-to-front: a match's *position*
        // in the merged text never shifts for an earlier (leftward) match once a later one is
        // replaced (only text at/after the edit point moves). Re-running the *token regex* after
        // each splice would instead re-find a token whose replacement text is itself still
        // "{{token}}" — the Highlight policy's whole point — and loop forever, so positions are
        // captured once, up front, from the untouched text.
        //
        // The run/span *map* is a different story: two tokens routinely share one run (e.g. a
        // paragraph typed in a single pass has no mid-run split at all), so splicing the first
        // token removes and replaces that run — any span still pointing at it is now parentless.
        // Re-scanning the map (cheap: just walks the current runs) before every splice keeps it
        // valid without re-tokenizing.
        var matches = InlineTokenPattern.Matches(RunTextScanner.Scan(paragraph).MergedText)
            .OrderByDescending(m => m.Index).ToList();

        foreach (var match in matches)
        {
            var spans = RunTextScanner.Scan(paragraph).Spans;
            var isHtml = match.Groups[1].Success;
            var fieldPath = match.Groups[2].Value;
            var replacementText = ResolveScalarReplacement(fieldPath, context, policy, out var highlight);

            if (isHtml)
            {
                var htmlValue = context.TryResolve(fieldPath, out var raw) ? raw as string : null;
                replacementText = htmlValue ?? replacementText;
                var converted = HtmlToOoxmlConverter.ConvertFragment(mainPart, replacementText ?? "");
                if (converted.Count > 1)
                {
                    throw new InvalidOperationException(
                        $"{{{{html:{fieldPath}}}}} produced multi-paragraph content but was used inline mid-paragraph. " +
                        "Put multi-block HTML on its own paragraph containing only that token.");
                }

                var inlineRuns = converted.Count == 1 ? converted[0].Elements<Run>().Select(r => (Run)r.CloneNode(true)).ToList() : [];
                var anchor = RunTextScanner.ReplaceRange(paragraph, spans, match.Index, match.Length, "");
                foreach (var run in inlineRuns)
                    anchor.InsertBeforeSelf(run);
                anchor.Remove();
            }
            else
            {
                var newRun = RunTextScanner.ReplaceRange(paragraph, spans, match.Index, match.Length, replacementText ?? "");
                if (highlight)
                    ApplyHighlight(newRun);
            }
        }
    }

    private static void SubstituteWholeParagraphHtml(Paragraph paragraph, string fieldPath, TemplateContext context, MainDocumentPart mainPart, MissingTokenPolicy policy)
    {
        var found = context.TryResolve(fieldPath, out var raw);
        if (!found || raw is not string html)
        {
            switch (policy)
            {
                case MissingTokenPolicy.Error:
                    throw new MissingTemplateTokenException(fieldPath);
                case MissingTokenPolicy.Redact:
                    paragraph.RemoveAllChildren<Run>();
                    return;
                case MissingTokenPolicy.Highlight:
                    paragraph.RemoveAllChildren<Run>();
                    var run = new Run(new Text($"{{{{html:{fieldPath}}}}}") { Space = SpaceProcessingModeValues.Preserve });
                    ApplyHighlight(run);
                    paragraph.AppendChild(run);
                    return;
            }
        }

        var converted = HtmlToOoxmlConverter.ConvertFragment(mainPart, raw as string ?? "");
        foreach (var newParagraph in converted)
            paragraph.InsertBeforeSelf(newParagraph);
        paragraph.Remove();
    }

    private static string? ResolveScalarReplacement(string fieldPath, TemplateContext context, MissingTokenPolicy policy, out bool highlight)
    {
        highlight = false;
        if (context.TryResolve(fieldPath, out var value) && value is not null)
            return TemplateValueFormatter.ToComparableString(value);

        return policy switch
        {
            MissingTokenPolicy.Error => throw new MissingTemplateTokenException(fieldPath),
            MissingTokenPolicy.Redact => "",
            MissingTokenPolicy.Highlight => Highlighted($"{{{{{fieldPath}}}}}", out highlight),
            _ => throw new NotSupportedException(policy.ToString())
        };
    }

    private static string Highlighted(string text, out bool highlight)
    {
        highlight = true;
        return text;
    }

    private static void ApplyHighlight(Run run)
    {
        var props = run.RunProperties ??= new RunProperties();
        if (!props.Elements<Highlight>().Any())
            props.AppendChild(new Highlight { Val = HighlightColorValues.Yellow });
    }

    // ---- Clause marker resolution (second pass, after the main document is saved) ------------

    private static void ResolveClauseMarkers(string outputPath, ClauseLibrary clauseLibrary, IReadOnlyDictionary<string, object?> data, MissingTokenPolicy policy)
    {
        var transplant = new ClauseTransplantService();
        var context = new TemplateContext(data);

        while (true)
        {
            var paragraphs = transplant.ListParagraphs(outputPath);
            var markerIndex = -1;
            string? clauseId = null;
            foreach (var p in paragraphs)
            {
                var match = ClausePattern.Match(p.Text.Trim());
                if (match.Success)
                {
                    markerIndex = p.Index;
                    clauseId = match.Groups[1].Value;
                    break;
                }
            }

            if (markerIndex < 0)
                return;

            var range = clauseLibrary.FindClause(clauseId!);
            if (range is null)
            {
                switch (policy)
                {
                    case MissingTokenPolicy.Error:
                        throw new InvalidOperationException($"Clause library has no clause with id '{clauseId}'.");
                    case MissingTokenPolicy.Redact:
                        transplant.RemoveParagraphs(outputPath, markerIndex, 1, outputPath);
                        continue;
                    case MissingTokenPolicy.Highlight:
                        ReplaceParagraphTextHighlighted(outputPath, markerIndex, $"[Missing clause: {clauseId}]");
                        continue;
                }
            }

            var (start, count) = range!.Value;
            transplant.TransplantParagraphs(clauseLibrary.DocxPath, start, count, outputPath, markerIndex, outputPath);
            transplant.ContinueHeadingNumbering(outputPath, markerIndex, count);
            SubstituteInlineInRange(outputPath, markerIndex, count, context, policy);
            transplant.RemoveParagraphs(outputPath, markerIndex + count, 1, outputPath);
        }
    }

    private static void SubstituteInlineInRange(string docxPath, int startIndex, int count, TemplateContext context, MissingTokenPolicy policy)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Document has no main part.");
        var body = mainPart.Document?.Body ?? throw new InvalidOperationException("Document has no body.");
        var paragraphs = body.Elements<Paragraph>().Skip(startIndex).Take(count).ToList();

        foreach (var paragraph in paragraphs)
            SubstituteInline(paragraph, context, mainPart, policy);

        mainPart.Document!.Save();
    }

    private static void ReplaceParagraphTextHighlighted(string docxPath, int paragraphIndex, string text)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var body = doc.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Document has no body.");
        var paragraph = body.Elements<Paragraph>().ElementAt(paragraphIndex);
        paragraph.RemoveAllChildren<Run>();
        var run = new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        ApplyHighlight(run);
        paragraph.AppendChild(run);
        doc.MainDocumentPart!.Document.Save();
    }
}
