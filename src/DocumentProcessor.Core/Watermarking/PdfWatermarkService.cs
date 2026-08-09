using DocumentProcessor.Core.PdfFonts;
using PdfSharp.Drawing;
using PdfSharp.Pdf.IO;
using SkiaSharp;

namespace DocumentProcessor.Core.Watermarking;

/// <summary>
/// Stamps a large, semi-transparent diagonal text watermark onto every page of a PDF.
/// </summary>
/// <remarks>
/// The watermark is rasterized to an image and drawn as a picture rather than drawn as PDF text.
/// Text drawn directly onto a PDF page (e.g. via XGraphics.DrawString) becomes ordinary selectable/
/// copyable page content — indistinguishable to a PDF viewer from the document's real text, so
/// dragging a selection across the page picks up watermark characters along with it. Rasterizing
/// avoids that entirely: pixels can't be text-selected. This is the standard approach production
/// watermarking tools use. (Contrast with ESignFieldService's PDF anchors, which must stay real
/// text on purpose — e-signature platforms detect them by scanning the page's text layer.)
/// </remarks>
public sealed class PdfWatermarkService
{
    private const float SupersampleScale = 2f;

    public void AddTextWatermark(
        string pdfPath,
        string outputPath,
        string text,
        string fontFamily = "Arial",
        double rotationDegrees = -45,
        byte grayLevel = 192,
        byte alpha = 100)
    {
        using var document = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Modify);
        var fontBytes = PdfFontResolver.Instance.GetFontBytes(fontFamily);

        foreach (var page in document.Pages)
        {
            var widthPt = page.Width.Point;
            var heightPt = page.Height.Point;

            using var watermarkPng = RenderWatermarkPng(text, fontBytes, rotationDegrees, grayLevel, alpha, widthPt, heightPt);
            using var image = XImage.FromStream(watermarkPng);

            // Prepend, not the default Append: FromPdfPage's default draws new content on top of
            // the page's existing content stream, which would paint the watermark over the real
            // text instead of behind it. Prepend inserts our drawing before the existing content
            // so the document's actual text renders on top of the watermark, as a watermark should.
            using var gfx = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Prepend);
            gfx.DrawImage(image, 0, 0, widthPt, heightPt);
        }

        document.Save(outputPath);
    }

    private static MemoryStream RenderWatermarkPng(
        string text, byte[] fontBytes, double rotationDegrees, byte grayLevel, byte alpha, double widthPt, double heightPt)
    {
        var widthPx = (int)(widthPt * SupersampleScale);
        var heightPx = (int)(heightPt * SupersampleScale);

        using var bitmap = new SKBitmap(widthPx, heightPx, SKColorType.Rgba8888, SKAlphaType.Premul);
        using var canvas = new SKCanvas(bitmap);
        canvas.Clear(SKColors.Transparent);

        using var typefaceData = SKData.CreateCopy(fontBytes);
        using var typeface = SKTypeface.FromData(typefaceData, 0) ?? SKTypeface.Default;
        using var font = new SKFont(typeface, 72 * SupersampleScale);
        using var paint = new SKPaint
        {
            Color = new SKColor(grayLevel, grayLevel, grayLevel, alpha),
            IsAntialias = true
        };

        canvas.Translate(widthPx / 2f, heightPx / 2f);
        canvas.RotateDegrees((float)rotationDegrees);
        canvas.DrawText(text, 0, 0, SKTextAlign.Center, font, paint);
        canvas.Flush();

        using var image = SKImage.FromBitmap(bitmap);
        using var data = image.Encode(SKEncodedImageFormat.Png, quality: 100);
        var stream = new MemoryStream();
        data.SaveTo(stream);
        stream.Position = 0;
        return stream;
    }
}
