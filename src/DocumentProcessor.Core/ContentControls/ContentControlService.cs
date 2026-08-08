using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentProcessor.Core.ContentControls;

public sealed record ContentControlInfo(string? Tag, string? Alias, string Text);

/// <summary>
/// Reads and replaces the values of Word content controls (structured document tags, w:sdt)
/// identified by their tag — the standard mechanism for template-driven contract assembly.
/// </summary>
public sealed class ContentControlService
{
    /// <summary>
    /// Replaces the text content of every content control matching <paramref name="tag"/>
    /// with <paramref name="newValue"/>, preserving the run formatting of the first existing run.
    /// </summary>
    /// <returns>The number of content controls updated.</returns>
    public int ReplaceByTag(string docxPath, string tag, string newValue)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var document = RequireDocument(doc);

        var updated = 0;
        foreach (var sdt in document.Descendants<SdtElement>())
        {
            var sdtTag = sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;
            if (!string.Equals(sdtTag, tag, StringComparison.Ordinal))
                continue;

            SetContentText(sdt, newValue);
            updated++;
        }

        if (updated > 0)
            document.Save();

        return updated;
    }

    /// <summary>
    /// Replaces content-control values in bulk from a tag→value map. Useful for template assembly
    /// where a single pass should populate every field in a contract in one document open/save.
    /// </summary>
    public IReadOnlyDictionary<string, int> ReplaceMany(string docxPath, IReadOnlyDictionary<string, string> tagToValue)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var document = RequireDocument(doc);

        var counts = tagToValue.Keys.ToDictionary(t => t, _ => 0);
        foreach (var sdt in document.Descendants<SdtElement>())
        {
            var sdtTag = sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;
            if (sdtTag is null || !tagToValue.TryGetValue(sdtTag, out var newValue))
                continue;

            SetContentText(sdt, newValue);
            counts[sdtTag]++;
        }

        document.Save();
        return counts;
    }

    /// <summary>Lists every content control in the document with its tag, alias, and current text.</summary>
    public IReadOnlyList<ContentControlInfo> ListContentControls(string docxPath)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: false);
        var document = RequireDocument(doc);

        return document.Descendants<SdtElement>()
            .Select(sdt => new ContentControlInfo(
                sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value,
                sdt.SdtProperties?.GetFirstChild<SdtAlias>()?.Val?.Value,
                string.Concat(sdt.Descendants<Text>().Select(t => t.Text))))
            .ToList();
    }

    private static Document RequireDocument(WordprocessingDocument doc) =>
        doc.MainDocumentPart?.Document ?? throw new InvalidOperationException("Document has no main part/body.");

    private static void SetContentText(SdtElement sdt, string newValue)
    {
        // The content control's editable region is whichever SdtContent* child it has
        // (SdtContentRun for inline controls, SdtContentBlock for block-level ones).
        var contentContainer = sdt switch
        {
            SdtRun run => (OpenXmlCompositeElement?)run.SdtContentRun,
            SdtBlock block => block.SdtContentBlock,
            SdtCell cell => cell.SdtContentCell,
            _ => null
        };
        if (contentContainer is null)
            return;

        var existingRuns = contentContainer.Descendants<Run>().ToList();
        var templateRunProps = existingRuns.FirstOrDefault()?.RunProperties?.CloneNode(true) as RunProperties;

        contentContainer.RemoveAllChildren();
        var newRun = new Run(new Text(newValue) { Space = SpaceProcessingModeValues.Preserve });
        if (templateRunProps is not null)
            newRun.PrependChild(templateRunProps);

        contentContainer.AppendChild(newRun);
    }
}
