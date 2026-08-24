using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentProcessor.Core.DocumentAssembly;

/// <summary>
/// Handles the "make the TOC/numbering catch up after I edited the document" problem. Two distinct
/// mechanisms, because Word computes displayed numbers/page references at render time, not as stored
/// text — this library can correct the underlying structure but can't pre-render Word's own layout:
/// <list type="bullet">
/// <item>Numbered-list paragraphs (<c>w:numId</c>) don't need explicit renumbering at all — Word
/// computes each list item's displayed number from its position among paragraphs sharing that numId,
/// live, every time it renders. Inserting/removing such a paragraph "renumbers" everything after it
/// automatically. There's nothing to fix here unless numbers were manually typed as literal text.</item>
/// <item>TOC and REF/PAGEREF fields cache their last-computed display text and only recompute it
/// when Word updates fields (F9, print, or "update fields on open"). <see cref="MarkAllFieldsDirty"/>
/// and <see cref="SetUpdateFieldsOnOpen"/> force that recompute to happen automatically the next time
/// the document is opened in Word, so a caller doesn't have to instruct the recipient to press F9.</item>
/// </list>
/// </summary>
public sealed class FieldUpdateService
{
    /// <summary>Marks every field in the document dirty (<c>w:dirty="true"</c> on simple fields and
    /// on the "begin" fldChar of complex fields) so Word recomputes their displayed text — TOC
    /// entries, REF/PAGEREF results, page counts — the next time fields are updated.</summary>
    public void MarkAllFieldsDirty(string docxPath)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var body = doc.MainDocumentPart?.Document?.Body ?? throw new CorruptDocumentException("Document has no main part/body.");

        foreach (var simpleField in body.Descendants<SimpleField>())
            simpleField.Dirty = true;

        foreach (var fieldChar in body.Descendants<FieldChar>().Where(f => f.FieldCharType?.Value == FieldCharValues.Begin))
            fieldChar.Dirty = true;

        doc.MainDocumentPart!.Document!.Save();
    }

    /// <summary>Sets (or clears) <c>w:updateFields</c> in document settings, so Word prompts to
    /// (or automatically does, per the user's Word options) update every field as soon as the
    /// document opens — the standard way a generated document ensures its TOC/cross-references are
    /// current without the recipient needing to know to press F9.</summary>
    public void SetUpdateFieldsOnOpen(string docxPath, bool updateOnOpen = true)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var mainPart = doc.MainDocumentPart ?? throw new CorruptDocumentException("Document has no main part.");
        var settingsPart = mainPart.DocumentSettingsPart ?? mainPart.AddNewPart<DocumentSettingsPart>();
        settingsPart.Settings ??= new Settings();

        settingsPart.Settings.RemoveAllChildren<UpdateFieldsOnOpen>();
        if (updateOnOpen)
            settingsPart.Settings.PrependChild(new UpdateFieldsOnOpen { Val = true });

        settingsPart.Settings.Save();
    }
}
