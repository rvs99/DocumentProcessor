using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.Layout;

namespace DocumentProcessor.Core.Samples;

/// <summary>
/// Generates throwaway .docx fixtures used by the demo app and tests.
/// Not part of the production surface — real callers bring their own documents.
/// </summary>
public static class SampleDocumentFactory
{
    public static void CreateBasicDocument(string path, string title, IEnumerable<string> paragraphs)
    {
        using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
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

            body.AppendChild(new SectionProperties());
            AddMinimalStyles(mainPart);
            mainPart.Document.Save();
        }

        ApplyPageDefaults(path);
    }

    public static void CreateDocumentWithContentControls(string path, IReadOnlyDictionary<string, string> tagToPlaceholder)
    {
        using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
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

            body.AppendChild(new SectionProperties());
            AddMinimalStyles(mainPart);
            mainPart.Document.Save();
        }

        ApplyPageDefaults(path);
    }

    /// <summary>Adds a minimal single-level numbering definition (one <c>w:abstractNum</c> +
    /// <c>w:num</c> pair) for <paramref name="numId"/> to an existing document — any paragraph
    /// referencing a <c>w:numId</c> via <see cref="NumberingProperties"/> needs a matching
    /// definition in the document's <see cref="NumberingDefinitionsPart"/>, or consumers that
    /// validate document structure (e.g. Clippit's <c>DocumentBuilder</c>) reject it outright.</summary>
    public static void AddNumberingDefinition(string path, int numId, NumberFormatValues? format = null)
    {
        using var doc = WordprocessingDocument.Open(path, isEditable: true);
        var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Document has no main part.");
        var numberingPart = mainPart.NumberingDefinitionsPart ?? mainPart.AddNewPart<NumberingDefinitionsPart>();
        numberingPart.Numbering ??= new Numbering();

        numberingPart.Numbering.AppendChild(new AbstractNum(
            new Level(
                new StartNumberingValue { Val = 1 },
                new NumberingFormat { Val = format ?? NumberFormatValues.Decimal },
                new LevelText { Val = "%1." })
            { LevelIndex = 0 })
        { AbstractNumberId = numId });
        numberingPart.Numbering.AppendChild(new NumberingInstance(new AbstractNumId { Val = numId }) { NumberID = numId });
        numberingPart.Numbering.Save();
    }

    /// <summary>Builds a document from caller-supplied paragraphs verbatim — for tests that need
    /// precise control over run boundaries/formatting (e.g. simulating Word's habit of splitting a
    /// single piece of text across several runs at spellcheck/revision boundaries).</summary>
    public static void CreateDocumentFromParagraphs(string path, IEnumerable<Paragraph> paragraphs)
    {
        using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body!;

            foreach (var paragraph in paragraphs)
                body.AppendChild(paragraph);

            body.AppendChild(new SectionProperties());
            AddMinimalStyles(mainPart);
            mainPart.Document.Save();
        }

        ApplyPageDefaults(path);
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
        using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
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

            body.AppendChild(new SectionProperties());
            AddMinimalStyles(mainPart);
            mainPart.Document.Save();
        }

        ApplyPageDefaults(path);
    }

    public static void AddMinimalStyles(MainDocumentPart mainPart)
    {
        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        var styles = new Styles(
            // Calibri 11pt is Word's own Normal.dotm baseline font — not an implicit zero. Leaving
            // it unspecified means every renderer (Word, LibreOffice, any other OOXML consumer)
            // falls back to its own built-in default instead, and those defaults don't agree.
            // Setting it explicitly removes that ambiguity for whichever engine renders this.
            // (Paragraph/line spacing is applied separately, after the initial save, via
            // PageLayoutService.SetDefaultParagraphSpacing — see ApplyPageDefaults below — rather
            // than duplicated here, so there's exactly one place that logic lives.)
            new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" },
                    new FontSize { Val = "22" }))),
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
    /// Applies the same page setup every real Word document effectively carries — 1-inch margins
    /// and the Normal.dotm baseline paragraph/line spacing (8pt space-after, 1.08x line spacing) —
    /// via the production <see cref="PageLayoutService"/> API, rather than duplicating that logic
    /// as sample-only inline construction. Missing paragraph/line spacing in particular was
    /// confirmed to cause substantial pagination drift against real Word during testing.
    /// </summary>
    private static void ApplyPageDefaults(string path)
    {
        var layout = new PageLayoutService();
        layout.SetMargins(path, PageMargins.FromInches(top: 1, bottom: 1, left: 1, right: 1));
        layout.SetDefaultParagraphSpacing(path, afterTwips: 160, lineTwips: 259, LineSpacingRuleValues.Auto);
    }
}
