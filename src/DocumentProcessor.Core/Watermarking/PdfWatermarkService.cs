using DocumentProcessor.Core.PdfFonts;
using PdfSharp.Drawing;
using PdfSharp.Pdf.IO;

namespace DocumentProcessor.Core.Watermarking;

/// <summary>
/// Stamps a large, semi-transparent diagonal text watermark onto every page of a PDF.
/// </summary>
public sealed class PdfWatermarkService
{
    public void AddTextWatermark(
        string pdfPath,
        string outputPath,
        string text,
        string fontFamily = "Arial",
        double rotationDegrees = -45,
        byte grayLevel = 192,
        byte alpha = 100)
    {
        PdfFontResolver.EnsureRegistered();
        using var document = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Modify);

        var font = new XFont(fontFamily, 72, XFontStyleEx.Bold);
        var brush = new XSolidBrush(XColor.FromArgb(alpha, grayLevel, grayLevel, grayLevel));
        var format = new XStringFormat { Alignment = XStringAlignment.Center, LineAlignment = XLineAlignment.Center };

        foreach (var page in document.Pages)
        {
            using var gfx = XGraphics.FromPdfPage(page);
            var center = new XPoint(page.Width.Point / 2, page.Height.Point / 2);

            var state = gfx.Save();
            gfx.RotateAtTransform(rotationDegrees, center);
            gfx.DrawString(text, font, brush, center, format);
            gfx.Restore(state);
        }

        document.Save(outputPath);
    }
}
