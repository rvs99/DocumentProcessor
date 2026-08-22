using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.Templating;
using OoxmlLock = DocumentFormat.OpenXml.Wordprocessing.Lock;

namespace DocumentProcessor.Core.ContentControls;

public sealed record ContentControlInfo(string? Tag, string? Alias, string Text, ContentControlLockMode? LockMode = null);

/// <summary>Mirrors OOXML's <c>ST_Lock</c> values for a content control's <c>w:lock</c> element.</summary>
public enum ContentControlLockMode
{
    /// <summary>The control can be deleted and its content edited (the default when no <c>w:lock</c> is present).</summary>
    Unlocked,
    /// <summary>The control itself can't be deleted, but its content can still be edited.</summary>
    SdtLocked,
    /// <summary>The control can be deleted, but its content can't be edited.</summary>
    ContentLocked,
    /// <summary>Neither the control nor its content can be edited or deleted.</summary>
    SdtContentLocked
}

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

    /// <summary>Lists every content control in the document with its tag, alias, current text, and lock state.</summary>
    public IReadOnlyList<ContentControlInfo> ListContentControls(string docxPath)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: false);
        var document = RequireDocument(doc);

        return document.Descendants<SdtElement>()
            .Select(sdt => new ContentControlInfo(
                sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value,
                sdt.SdtProperties?.GetFirstChild<SdtAlias>()?.Val?.Value,
                string.Concat(sdt.Descendants<Text>().Select(t => t.Text)),
                ReadLockMode(sdt)))
            .ToList();
    }

    /// <summary>
    /// Injects a sanitized HTML fragment as the content of every Rich Text content control matching
    /// <paramref name="tag"/> — the one SDT type whose content area allows multiple paragraphs and
    /// run-level formatting, so it's the only type a real HTML fragment (as opposed to a single
    /// plain-text value) can be mapped onto meaningfully.
    /// </summary>
    /// <returns>The number of content controls updated.</returns>
    public int SetContentRichTextByTag(string docxPath, string tag, string html)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var document = RequireDocument(doc);
        var mainPart = doc.MainDocumentPart!;

        var updated = 0;
        foreach (var sdt in FindByTag(document, tag))
        {
            var contentContainer = GetContentContainer(sdt);
            if (contentContainer is null)
                continue;

            var paragraphs = HtmlToOoxmlConverter.ConvertFragment(mainPart, html);
            contentContainer.RemoveAllChildren();
            if (contentContainer is SdtContentRun or SdtContentCell)
            {
                // Inline/cell containers hold runs directly, not paragraph wrappers — flatten.
                foreach (var run in paragraphs.SelectMany(p => p.Elements<Run>()))
                    contentContainer.AppendChild((Run)run.CloneNode(true));
            }
            else
            {
                foreach (var paragraph in paragraphs)
                    contentContainer.AppendChild((Paragraph)paragraph.CloneNode(true));
            }
            updated++;
        }

        if (updated > 0)
            document.Save();

        return updated;
    }

    /// <summary>
    /// Sets a Date Picker content control's value from a real <see cref="DateTime"/> — populating
    /// both the visible display text (formatted per <paramref name="displayFormat"/>, default
    /// <c>yyyy-MM-dd</c>) and the <c>w:date/@w:fullDate</c> metadata Word's date picker UI reads back.
    /// Applies to every control matching <paramref name="tag"/>, regardless of whether it's actually
    /// a Date Picker (falls back to plain-text substitution otherwise).
    /// </summary>
    /// <returns>The number of content controls updated.</returns>
    public int SetContentDateByTag(string docxPath, string tag, DateTime value, string? displayFormat = null)
    {
        var format = displayFormat ?? "yyyy-MM-dd";
        return ReplaceByTag(docxPath, tag, value.ToString(format));
    }

    /// <summary>
    /// Sets a Drop-Down List content control's selection to <paramref name="value"/>, validated
    /// against the control's own defined list items (<c>w:listItem/@w:value</c>) — unlike
    /// <see cref="ReplaceByTag"/>'s best-effort handling, this throws if <paramref name="value"/>
    /// isn't one of the control's valid options, since silently accepting an invalid selection would
    /// leave the control in a state Word's own dropdown UI could never produce.
    /// </summary>
    /// <returns>The number of content controls updated.</returns>
    public int SetContentDropDownSelectionByTag(string docxPath, string tag, string value)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var document = RequireDocument(doc);

        var updated = 0;
        foreach (var sdt in FindByTag(document, tag))
        {
            var dropDown = sdt.SdtProperties?.GetFirstChild<SdtContentDropDownList>()
                ?? throw new InvalidOperationException($"Content control '{tag}' is not a Drop-Down List control.");

            var listItems = dropDown.Elements<ListItem>().ToList();
            var match = listItems.FirstOrDefault(i => i.Value?.Value == value)
                ?? throw new ArgumentException(
                    $"'{value}' is not a valid option for drop-down '{tag}'. Valid values: " +
                    string.Join(", ", listItems.Select(i => i.Value?.Value)), nameof(value));

            // Set LastValue and the visible text independently — routing the display text back
            // through SetContentText would hit its own dropdown-aware branch and stomp LastValue
            // with the display text instead of leaving the just-set option value in place.
            dropDown.LastValue = value;
            if (GetContentContainer(sdt) is { } contentContainer)
                SetVisibleRunText(contentContainer, match.DisplayText?.Value ?? value);
            updated++;
        }

        if (updated > 0)
            document.Save();

        return updated;
    }

    /// <summary>Adds or removes a <c>w:lock</c> on every content control matching <paramref name="tag"/>
    /// — e.g. to lock fields against further edits once a template fill is complete.</summary>
    /// <returns>The number of content controls updated.</returns>
    public int SetLock(string docxPath, string tag, ContentControlLockMode mode)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var document = RequireDocument(doc);

        var updated = 0;
        foreach (var sdt in FindByTag(document, tag))
        {
            var properties = sdt.SdtProperties ??= new SdtProperties();
            properties.GetFirstChild<OoxmlLock>()?.Remove();
            if (mode != ContentControlLockMode.Unlocked)
                properties.AppendChild(new OoxmlLock { Val = ToLockingValue(mode) });
            updated++;
        }

        if (updated > 0)
            document.Save();

        return updated;
    }

    private static IEnumerable<SdtElement> FindByTag(Document document, string tag) =>
        document.Descendants<SdtElement>().Where(sdt => sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value == tag);

    private static Document RequireDocument(WordprocessingDocument doc) =>
        doc.MainDocumentPart?.Document ?? throw new InvalidOperationException("Document has no main part/body.");

    private static ContentControlLockMode? ReadLockMode(SdtElement sdt)
    {
        var val = sdt.SdtProperties?.GetFirstChild<OoxmlLock>()?.Val?.Value;
        if (val is null) return null;
        if (val == LockingValues.SdtLocked) return ContentControlLockMode.SdtLocked;
        if (val == LockingValues.ContentLocked) return ContentControlLockMode.ContentLocked;
        if (val == LockingValues.SdtContentLocked) return ContentControlLockMode.SdtContentLocked;
        return ContentControlLockMode.Unlocked;
    }

    private static LockingValues ToLockingValue(ContentControlLockMode mode) => mode switch
    {
        ContentControlLockMode.SdtLocked => LockingValues.SdtLocked,
        ContentControlLockMode.ContentLocked => LockingValues.ContentLocked,
        ContentControlLockMode.SdtContentLocked => LockingValues.SdtContentLocked,
        _ => LockingValues.Unlocked
    };

    private static OpenXmlCompositeElement? GetContentContainer(SdtElement sdt) => sdt switch
    {
        SdtRun run => run.SdtContentRun,
        SdtBlock block => block.SdtContentBlock,
        SdtCell cell => cell.SdtContentCell,
        _ => null
    };

    private static void SetContentText(SdtElement sdt, string newValue)
    {
        var contentContainer = GetContentContainer(sdt);
        if (contentContainer is null)
            return;

        // Type-specific metadata: a Date Picker's visible text is separate from the w:fullDate
        // value Word's own picker UI reads, and a Drop-Down/Combo Box's is separate from its
        // w:lastValue — update whichever metadata this control actually carries, in addition to
        // the visible run text every SDT type shares.
        var properties = sdt.SdtProperties;
        if (properties?.GetFirstChild<SdtContentDate>() is { } date)
        {
            if (DateTime.TryParse(newValue, out var parsed))
                date.FullDate = new DateTimeValue(parsed);
        }
        else if (properties?.GetFirstChild<SdtContentDropDownList>() is { } dropDown)
        {
            dropDown.LastValue = newValue;
        }
        else if (properties?.GetFirstChild<SdtContentComboBox>() is { } comboBox)
        {
            comboBox.LastValue = newValue;
        }

        SetVisibleRunText(contentContainer, newValue);
    }

    private static void SetVisibleRunText(OpenXmlCompositeElement contentContainer, string text)
    {
        var existingRuns = contentContainer.Descendants<Run>().ToList();
        var templateRunProps = existingRuns.FirstOrDefault()?.RunProperties?.CloneNode(true) as RunProperties;

        contentContainer.RemoveAllChildren();
        var newRun = new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        if (templateRunProps is not null)
            newRun.PrependChild(templateRunProps);

        contentContainer.AppendChild(newRun);
    }
}
