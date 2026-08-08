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

        body.AppendChild(new SectionProperties());
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

        body.AppendChild(new SectionProperties());
        AddMinimalStyles(mainPart);
        mainPart.Document.Save();
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

        body.AppendChild(new SectionProperties());
        AddMinimalStyles(mainPart);
        mainPart.Document.Save();
    }

    public static void AddMinimalStyles(MainDocumentPart mainPart)
    {
        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        var styles = new Styles(
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
                new RunProperties(new FontSize { Val = "22" }))
            {
                Type = StyleValues.Paragraph,
                StyleId = "Normal",
                Default = true
            });
        stylesPart.Styles = styles;
        stylesPart.Styles.Save();
    }
}
