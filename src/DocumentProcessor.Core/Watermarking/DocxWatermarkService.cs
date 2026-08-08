using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;
using Picture = DocumentFormat.OpenXml.Wordprocessing.Picture;

namespace DocumentProcessor.Core.Watermarking;

/// <summary>
/// Adds a diagonal text watermark (e.g. "DRAFT", "CONFIDENTIAL") to a .docx file, using the same
/// VML shape-in-header mechanism Word itself generates for Insert → Watermark. Applied via the
/// document header so it repeats on every page, matching Word's own behavior.
/// </summary>
public sealed class DocxWatermarkService
{
    /// <summary>
    /// Adds (or replaces) a text watermark on every section of the document. Safe to call more than
    /// once — a prior watermark header added by this method is replaced rather than duplicated.
    /// </summary>
    public void AddTextWatermark(
        string docxPath,
        string text,
        string fontFamily = "Calibri",
        int rotationDegrees = -45,
        string colorHex = "C0C0C0")
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Document has no main part.");
        var document = mainPart.Document ?? throw new InvalidOperationException("Document has no body.");
        var body = document.Body ?? throw new InvalidOperationException("Document has no body.");

        var headerPart = mainPart.AddNewPart<HeaderPart>();
        headerPart.Header = BuildWatermarkHeader(text, fontFamily, rotationDegrees, colorHex);
        headerPart.Header.Save();
        var headerRelId = mainPart.GetIdOfPart(headerPart);

        var sectionProperties = body.Elements<SectionProperties>().ToList();
        if (sectionProperties.Count == 0)
        {
            sectionProperties.Add(new SectionProperties());
            body.AppendChild(sectionProperties[0]);
        }

        foreach (var sectPr in sectionProperties)
        {
            foreach (var existing in sectPr.Elements<HeaderReference>()
                .Where(h => h.Type is null || h.Type == HeaderFooterValues.Default).ToList())
            {
                existing.Remove();
            }

            sectPr.PrependChild(new HeaderReference { Type = HeaderFooterValues.Default, Id = headerRelId });
        }

        document.Save();
    }

    private static Header BuildWatermarkHeader(string text, string fontFamily, int rotationDegrees, string colorHex)
    {
        var shape = new Shape(
            new TextPath
            {
                Style = $"font-family:'{fontFamily}';font-size:1pt",
                String = text,
                FitShape = true
            },
            new Fill { Opacity = "0.5" })
        {
            Id = "DocProcWatermark",
            Style = "position:absolute;left:0;top:0;width:415pt;height:207.5pt;" +
                    $"rotation:{rotationDegrees};z-index:-251654144;" +
                    "mso-position-horizontal:center;mso-position-horizontal-relative:margin;" +
                    "mso-position-vertical:center;mso-position-vertical-relative:margin",
            Filled = true,
            FillColor = $"#{colorHex}",
            Stroked = false
        };

        var run = new Run(new Picture(shape));
        var paragraph = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Header" }), run);
        return new Header(paragraph);
    }
}
