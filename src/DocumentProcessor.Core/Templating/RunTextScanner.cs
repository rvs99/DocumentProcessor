using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentProcessor.Core.Templating;

/// <summary>
/// Merges a paragraph's text across however many <c>w:r</c> runs Word split it into (Word fragments
/// runs at spellcheck/grammar/revision boundaries, so a single <c>{{token}}</c> routinely spans 2-4
/// runs in real documents) into one string for regex scanning, then splices matched ranges back into
/// the run tree — splitting the run(s) at the match boundary and reusing the formatting of whichever
/// run the match started in, the same "preserve the first run's formatting" convention
/// <see cref="ContentControls.ContentControlService"/> already uses for content controls.
/// </summary>
internal static class RunTextScanner
{
    internal readonly record struct RunSpan(Run Run, int Start, int Length);

    /// <summary>Concatenates every <c>w:t</c> in <paramref name="paragraph"/>'s direct runs into one
    /// string, alongside a map of which run covers which range of that string. Runs with no text
    /// (e.g. a bare tab/break run) are omitted from the map — they can never contain a match and are
    /// left completely untouched by <see cref="ReplaceRange"/>.</summary>
    public static (string MergedText, List<RunSpan> Spans) Scan(Paragraph paragraph)
    {
        var text = new System.Text.StringBuilder();
        var spans = new List<RunSpan>();

        foreach (var run in paragraph.Elements<Run>())
        {
            var runText = string.Concat(run.Elements<Text>().Select(t => t.Text));
            if (runText.Length == 0)
                continue;

            spans.Add(new RunSpan(run, text.Length, runText.Length));
            text.Append(runText);
        }

        return (text.ToString(), spans);
    }

    /// <summary>
    /// Replaces merged-text range [<paramref name="start"/>, <paramref name="start"/> +
    /// <paramref name="length"/>) with <paramref name="replacementText"/>. Any run fully inside the
    /// range is removed; a run only partially covered is split so its untouched prefix/suffix survive
    /// as their own runs with the original formatting. Returns the newly inserted run so callers
    /// (highlighting, HTML injection) can decorate it further.
    /// </summary>
    public static Run ReplaceRange(Paragraph paragraph, IReadOnlyList<RunSpan> spans, int start, int length, string replacementText)
    {
        var end = start + length;
        RunProperties? templateProps = null;
        Run? newRun = null;
        Run? lastTouched = null;

        foreach (var span in spans)
        {
            var spanEnd = span.Start + span.Length;
            if (spanEnd <= start || span.Start >= end)
                continue;

            templateProps ??= span.Run.RunProperties?.CloneNode(true) as RunProperties;

            var runText = string.Concat(span.Run.Elements<Text>().Select(t => t.Text));
            var localStart = Math.Max(0, start - span.Start);
            var localEnd = Math.Min(span.Length, end - span.Start);
            var prefix = runText[..localStart];
            var suffix = runText[localEnd..];

            if (prefix.Length > 0)
                span.Run.InsertBeforeSelf(CloneRunWithText(span.Run, prefix));

            if (newRun is null)
            {
                newRun = new Run(new Text(replacementText) { Space = SpaceProcessingModeValues.Preserve });
                if (templateProps is not null)
                    newRun.PrependChild((RunProperties)templateProps.CloneNode(true));
                span.Run.InsertBeforeSelf(newRun);
            }

            if (suffix.Length > 0)
                span.Run.InsertBeforeSelf(CloneRunWithText(span.Run, suffix));

            lastTouched = span.Run;
            span.Run.Remove();
        }

        if (newRun is null)
        {
            // Zero-width match with nothing covering it (shouldn't happen for a real regex match,
            // but keep the paragraph valid rather than throwing).
            newRun = new Run(new Text(replacementText) { Space = SpaceProcessingModeValues.Preserve });
            paragraph.AppendChild(newRun);
        }

        _ = lastTouched;
        return newRun;
    }

    private static Run CloneRunWithText(Run original, string text)
    {
        var clone = (Run)original.CloneNode(true);
        foreach (var t in clone.Elements<Text>().ToList())
            t.Remove();
        clone.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        return clone;
    }
}
