using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.PdfFonts;
using PdfSharp.Drawing;
using PdfSharp.Pdf.IO;

namespace DocumentProcessor.Core.ESign;

/// <summary>
/// Injects e-signature placeholder fields using the "anchor text" convention that DocuSign, Adobe
/// Sign, and most e-signature platforms use for auto-placement: a distinctive text string (e.g.
/// "/sig1/") is embedded in the document at the desired location, and the e-signature platform's
/// own upload/tagging step detects that string and places a real signature field there. This is
/// the standard integration point for both formats — there is no free library that can pre-populate
/// a genuinely provider-agnostic native PDF signature widget or docx signature line, since accepting
/// the signature is always done by the e-sign provider once the document is uploaded to them.
/// For docx, the anchor is additionally wrapped in a tagged content control so platforms that scan
/// content controls (rather than plain text) can also detect it.
/// </summary>
public sealed class ESignFieldService
{
    /// <summary>Appends a paragraph containing a tagged anchor for an e-signature field to the end of the document.</summary>
    public void InjectDocxAnchor(string docxPath, string anchorText, string tag = "ESignatureField")
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Document has no main part.");
        var document = mainPart.Document ?? throw new InvalidOperationException("Document has no body.");
        var body = document.Body ?? throw new InvalidOperationException("Document has no body.");

        var sdt = new SdtRun(
            new SdtProperties(new Tag { Val = tag }, new SdtAlias { Val = tag }),
            new SdtContentRun(new Run(new Text(anchorText))));

        var sectPr = body.Elements<SectionProperties>().FirstOrDefault();
        body.InsertBefore(new Paragraph(sdt), sectPr);

        document.Save();
    }

    /// <summary>Stamps an anchor for an e-signature field onto a PDF page at the given position.</summary>
    public void InjectPdfAnchor(
        string pdfPath, string outputPath, string anchorText,
        int pageIndex, double x, double y,
        bool invisible = false)
    {
        PdfFontResolver.EnsureRegistered();
        using var document = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Modify);
        if (pageIndex < 0 || pageIndex >= document.PageCount)
            throw new ArgumentOutOfRangeException(nameof(pageIndex), $"Document has {document.PageCount} page(s).");

        var page = document.Pages[pageIndex];
        using var gfx = XGraphics.FromPdfPage(page);

        var font = new XFont("Arial", 8);
        // Some e-sign platforms scan for invisible (white-on-white) anchor text so it doesn't
        // clutter the visible page; others expect visible tags reviewers can see before sending.
        var brush = invisible ? new XSolidBrush(XColor.FromArgb(255, 255, 255)) : new XSolidBrush(XColor.FromArgb(0, 0, 0));
        gfx.DrawString(anchorText, font, brush, new XPoint(x, y));

        document.Save(outputPath);
    }
}
