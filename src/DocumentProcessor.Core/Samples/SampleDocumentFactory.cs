using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentProcessor.Core.Samples;

/// <summary>
/// Generates throwaway .docx fixtures used by the demo app and tests.
/// Not part of the production surface — real callers bring their own documents.
/// </summary>
public static class SampleDocumentFactory
{
    public static void CreateBasicDocument(string path, string title, IEnumerable<string> paragraphs)
    {
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        var body = mainPart.Document.Body!;

        body.AppendChild(new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Title" }),
            new Run(new Text(title))));

        foreach (var text in paragraphs)
        {
            body.AppendChild(new Paragraph(new Run(new Text(text))));
        }

        body.AppendChild(CreateDefaultSectionProperties());
        AddMinimalStyles(mainPart);
        mainPart.Document.Save();
    }

    public static void CreateDocumentWithContentControls(string path, IReadOnlyDictionary<string, string> tagToPlaceholder)
    {
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        var body = mainPart.Document.Body!;

        body.AppendChild(new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Title" }),
            new Run(new Text("Sample Content-Controlled Agreement"))));

        foreach (var (tag, placeholder) in tagToPlaceholder)
        {
            var sdt = new SdtRun(
                new SdtProperties(
                    new Tag { Val = tag },
                    new SdtAlias { Val = tag }),
                new SdtContentRun(new Run(new Text(placeholder))));

            body.AppendChild(new Paragraph(
                new Run(new Text($"{tag}: ")),
                sdt));
        }

        body.AppendChild(CreateDefaultSectionProperties());
        AddMinimalStyles(mainPart);
        mainPart.Document.Save();
    }

    /// <summary>Appends plain paragraphs to an existing document, before its final section properties.</summary>
    public static void AppendParagraphs(string path, IEnumerable<string> paragraphs)
    {
        using var doc = WordprocessingDocument.Open(path, isEditable: true);
        var document = doc.MainDocumentPart?.Document ?? throw new InvalidOperationException("Document has no main part/body.");
        var body = document.Body ?? throw new InvalidOperationException("Document has no body.");
        var sectPr = body.Elements<SectionProperties>().FirstOrDefault();

        foreach (var text in paragraphs)
            body.InsertBefore(new Paragraph(new Run(new Text(text))), sectPr);

        document.Save();
    }

    public static void CreateDocumentWithTrackedChanges(string path, string author = "Reviewer")
    {
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        var body = mainPart.Document.Body!;

        var dateTime = new DateTimeValue(new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc));
        body.AppendChild(new Paragraph(
            new Run(new Text("The term is ") { Space = SpaceProcessingModeValues.Preserve }),
            new InsertedRun(new Run(new Text("twenty-four") { Space = SpaceProcessingModeValues.Preserve }))
                { Id = "1", Author = author, Date = dateTime },
            new DeletedRun(new Run(new DeletedText("twelve") { Space = SpaceProcessingModeValues.Preserve }))
                { Id = "2", Author = author, Date = dateTime },
            new Run(new Text(" months.") { Space = SpaceProcessingModeValues.Preserve })));

        body.AppendChild(CreateDefaultSectionProperties());
        AddMinimalStyles(mainPart);
        mainPart.Document.Save();
    }

    public static void AddMinimalStyles(MainDocumentPart mainPart)
    {
        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        var styles = new Styles(
            // Real Word documents always carry these Normal.dotm baseline values (Calibri 11pt,
            // 8pt space-after, 1.08 line spacing) — they aren't an implicit zero, they're Word's
            // own default template. Leaving them unspecified means every renderer (Word, LibreOffice,
            // any other OOXML consumer) falls back to its own built-in default instead, and those
            // defaults don't agree — confirmed to cause substantial pagination drift in testing.
            // Setting them explicitly removes that ambiguity for whichever engine renders this.
            new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" },
                    new FontSize { Val = "22" })),
                new ParagraphPropertiesDefault(new ParagraphPropertiesBaseStyle(
                    new SpacingBetweenLines { After = "160", Line = "259", LineRule = LineSpacingRuleValues.Auto }))),
            new Style(
                new StyleName { Val = "Title" },
                new BasedOn { Val = "Normal" },
                new NextParagraphStyle { Val = "Normal" },
                new RunProperties(new Bold(), new FontSize { Val = "32" }))
            {
                Type = StyleValues.Paragraph,
                StyleId = "Title"
            },
            new Style(
                new StyleName { Val = "Normal" },
                new ParagraphProperties(new SpacingBetweenLines { After = "160", Line = "259", LineRule = LineSpacingRuleValues.Auto }),
                new RunProperties(new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" }, new FontSize { Val = "22" }))
            {
                Type = StyleValues.Paragraph,
                StyleId = "Normal",
                Default = true
            });
        stylesPart.Styles = styles;
        stylesPart.Styles.Save();
    }

    /// <summary>
    /// Explicit 1-inch margins on every side. LibreOffice's implicit default already matches this
    /// when w:pgMar is omitted, but leaving it explicit removes that as a variable when comparing
    /// against other renderers (e.g. real Word, or alternative conversion engines).
    /// </summary>
    private static SectionProperties CreateDefaultSectionProperties() =>
        new(new PageMargin { Top = 1440, Bottom = 1440, Left = 1440, Right = 1440, Header = 720, Footer = 720, Gutter = 0 });
}
