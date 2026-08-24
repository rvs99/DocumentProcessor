using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.DocumentAssembly;
using DocumentProcessor.Core.FontEmbedding;
using DocumentProcessor.Core.Layout;
using DocumentProcessor.Core.Redlining;
using DocumentProcessor.Core.TrackChanges;
using DocumentProcessor.Core.Watermarking;
using PageSize = DocumentProcessor.Core.Layout.PageSize;

namespace DocumentProcessor.Core.Sessions;

/// <summary>Tracked-change operations bound to an open <see cref="DocumentSession"/>.</summary>
public sealed class TrackChangesOperations(DocumentSession session)
{
    /// <inheritdoc cref="TrackChangesService.AcceptAll(string)"/>
    public void AcceptAll() => TrackChangesService.AcceptAllCore(session.Document);

    /// <inheritdoc cref="TrackChangesService.AcceptByAuthor(string, string)"/>
    public int AcceptByAuthor(string author) =>
        TrackChangesService.AcceptWhereCore(session.Document, (_, a) => a == author);

    /// <inheritdoc cref="TrackChangesService.AcceptById(string, string)"/>
    public int AcceptById(string changeId) =>
        TrackChangesService.AcceptWhereCore(session.Document, (id, _) => id == changeId);

    /// <inheritdoc cref="TrackChangesService.RejectAll(string)"/>
    public int RejectAll() =>
        TrackChangesService.RejectWhereCore(session.Document, (_, _) => true);

    /// <inheritdoc cref="TrackChangesService.RejectByAuthor(string, string)"/>
    public int RejectByAuthor(string author) =>
        TrackChangesService.RejectWhereCore(session.Document, (_, a) => a == author);

    /// <inheritdoc cref="TrackChangesService.RejectById(string, string)"/>
    public int RejectById(string changeId) =>
        TrackChangesService.RejectWhereCore(session.Document, (id, _) => id == changeId);

    /// <inheritdoc cref="TrackChangesService.HasTrackedChanges(string)"/>
    public bool HasTrackedChanges() => TrackChangesService.HasTrackedChangesCore(session.Document);

    /// <inheritdoc cref="TrackChangesService.GetTrackedChanges(string)"/>
    public IReadOnlyList<TrackedChange> GetTrackedChanges() =>
        TrackChangesService.GetTrackedChangesCore(session.Document);
}

/// <summary>Page-setup operations bound to an open <see cref="DocumentSession"/>.</summary>
public sealed class PageLayoutOperations(DocumentSession session)
{
    /// <inheritdoc cref="PageLayoutService.SetPageSize(string, PageSize, int?)"/>
    public void SetPageSize(PageSize size, int? sectionIndex = null) =>
        PageLayoutService.SetPageSizeCore(session.Document, size, sectionIndex);

    /// <inheritdoc cref="PageLayoutService.SetMargins(string, PageMargins, int?)"/>
    public void SetMargins(PageMargins margins, int? sectionIndex = null) =>
        PageLayoutService.SetMarginsCore(session.Document, margins, sectionIndex);

    /// <inheritdoc cref="PageLayoutService.SetColumns(string, int, int, int?)"/>
    public void SetColumns(int columnCount, int spacingTwips = 720, int? sectionIndex = null) =>
        PageLayoutService.SetColumnsCore(session.Document, columnCount, spacingTwips, sectionIndex);

    /// <inheritdoc cref="PageLayoutService.InsertPageBreak(string, int)"/>
    public void InsertPageBreak(int beforeParagraphIndex) =>
        PageLayoutService.InsertPageBreakCore(session.Document, beforeParagraphIndex);

    /// <inheritdoc cref="PageLayoutService.InsertSectionBreak(string, int, SectionMarkValues?)"/>
    public void InsertSectionBreak(int beforeParagraphIndex, SectionMarkValues? breakType = null) =>
        PageLayoutService.InsertSectionBreakCore(session.Document, beforeParagraphIndex, breakType);

    /// <inheritdoc cref="PageLayoutService.SetDefaultParagraphSpacing(string, int, int, LineSpacingRuleValues)"/>
    public void SetDefaultParagraphSpacing(int afterTwips, int lineTwips, LineSpacingRuleValues lineRule) =>
        PageLayoutService.SetDefaultParagraphSpacingCore(session.Document, afterTwips, lineTwips, lineRule);
}

/// <summary>Watermark operations bound to an open <see cref="DocumentSession"/>.</summary>
public sealed class WatermarkOperations(DocumentSession session)
{
    /// <inheritdoc cref="DocxWatermarkService.AddTextWatermark(string, string, string, int, string, bool, WatermarkPosition, double, double, double)"/>
    public void AddText(
        string text,
        string fontFamily = "Calibri",
        int rotationDegrees = -45,
        string colorHex = "C0C0C0",
        bool removable = true,
        WatermarkPosition position = WatermarkPosition.Center,
        double widthPt = 415,
        double heightPt = 207.5,
        double fontSizePt = 72) =>
        DocxWatermarkService.AddTextWatermarkCore(
            session.Document, text, fontFamily, rotationDegrees, colorHex, removable, position, widthPt, heightPt, fontSizePt);

    /// <inheritdoc cref="DocxWatermarkService.RemoveWatermark(string)"/>
    public bool Remove() => DocxWatermarkService.RemoveWatermarkCore(session.Document);
}

/// <summary>Editing-restriction operations bound to an open <see cref="DocumentSession"/>.</summary>
public sealed class DocumentProtectionOperations(DocumentSession session)
{
    /// <inheritdoc cref="DocumentProtectionService.SetDocumentProtection(string, EditRestriction, string?)"/>
    public void Restrict(EditRestriction restriction, string? password = null) =>
        DocumentProtectionService.SetDocumentProtectionCore(session.Document, restriction, password);

    /// <inheritdoc cref="DocumentProtectionService.RemoveDocumentProtection(string)"/>
    public void Remove() => DocumentProtectionService.RemoveDocumentProtectionCore(session.Document);

    /// <inheritdoc cref="DocumentProtectionService.AllowEditingInRange(string, int, int, EditorGroup)"/>
    public void AllowEditingInRange(int startParagraphIndex, int endParagraphIndex, EditorGroup editorGroup = EditorGroup.Everyone) =>
        DocumentProtectionService.AllowEditingInRangeCore(session.Document, startParagraphIndex, endParagraphIndex, editorGroup);
}

/// <summary>Font-embedding operations bound to an open <see cref="DocumentSession"/>.</summary>
public sealed class FontOperations(DocumentSession session)
{
    /// <inheritdoc cref="FontEmbeddingService.EmbedFontFamily(string, string, FontFamilyFiles)"/>
    public void EmbedFamily(string fontFamilyName, FontFamilyFiles files) =>
        FontEmbeddingService.EmbedFontFamilyCore(session.Document, fontFamilyName, files);

    /// <inheritdoc cref="FontEmbeddingService.ApplyFontToAllRuns(string, string)"/>
    public void ApplyToAllRuns(string fontFamilyName) =>
        FontEmbeddingService.ApplyFontToAllRunsCore(session.Document, fontFamilyName);

    /// <summary>
    /// Embeds a family and applies it in one step. These are documented as needing to be used
    /// together — embedding alone changes nothing visible — and via the session they now cost one
    /// package cycle rather than two.
    /// </summary>
    public void EmbedAndApply(string fontFamilyName, FontFamilyFiles files)
    {
        FontEmbeddingService.EmbedFontFamilyCore(session.Document, fontFamilyName, files);
        FontEmbeddingService.ApplyFontToAllRunsCore(session.Document, fontFamilyName);
    }

    /// <inheritdoc cref="FontEmbeddingService.ListEmbeddedFonts(string)"/>
    public IReadOnlyList<string> ListEmbedded() =>
        FontEmbeddingService.ListEmbeddedFontsCore(session.Document);
}

/// <summary>Field-update operations bound to an open <see cref="DocumentSession"/>.</summary>
public sealed class FieldOperations(DocumentSession session)
{
    /// <inheritdoc cref="FieldUpdateService.MarkAllFieldsDirty(string)"/>
    public void MarkAllDirty() => FieldUpdateService.MarkAllFieldsDirtyCore(session.Document);

    /// <inheritdoc cref="FieldUpdateService.SetUpdateFieldsOnOpen(string, bool)"/>
    public void SetUpdateOnOpen(bool updateOnOpen = true) =>
        FieldUpdateService.SetUpdateFieldsOnOpenCore(session.Document, updateOnOpen);

    /// <inheritdoc cref="CrossReferenceValidator.Validate(string)"/>
    public IReadOnlyList<DanglingReference> ValidateCrossReferences() =>
        CrossReferenceValidator.ValidateCore(session.Document);
}
