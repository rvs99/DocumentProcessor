using DocumentProcessor.Core.Conversion;
using DocumentProcessor.Core.PdfFonts;
using DocumentProcessor.Core.Samples;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace DocumentProcessor.Tests;

/// <summary>Creates and cleans up temp file paths for test fixtures.</summary>
public static class TestFiles
{
    public static string NewTempPath(string extension) =>
        Path.Combine(Path.GetTempPath(), $"docproc-test-{Guid.NewGuid():N}{extension}");

    /// <summary>
    /// Conversion options for tests that exercise LibreOffice. Defaults to a native soffice binary
    /// (what CI/Linux containers use); set DOCPROC_LIBREOFFICE_WSL_DISTRO=Ubuntu to route through
    /// WSL instead, for local Windows development without installing the LibreOffice desktop app.
    /// </summary>
    public static WordToPdfConversionOptions ConversionOptions() =>
        Environment.GetEnvironmentVariable("DOCPROC_LIBREOFFICE_WSL_DISTRO") is { } distro
            ? new WordToPdfConversionOptions { UseWslDistro = distro }
            : new WordToPdfConversionOptions();

    /// <summary>Builds a simple .docx and converts it to PDF, for tests exercising PDF-only services.</summary>
    public static string NewTestPdf(string title, IEnumerable<string> paragraphs)
    {
        var docxPath = NewTempPath(".docx");
        var pdfPath = NewTempPath(".pdf");
        try
        {
            SampleDocumentFactory.CreateBasicDocument(docxPath, title, paragraphs);
            new WordToPdfConverter(ConversionOptions()).Convert(docxPath, pdfPath);
            return pdfPath;
        }
        finally
        {
            File.Delete(docxPath);
        }
    }

    /// <summary>
    /// Builds a PDF with <paramref name="pageCount"/> pages directly via PDFsharp — no LibreOffice
    /// involved — for tests of PDF-only services that just need identifiable pages, not real docx
    /// conversion fidelity. Each page's text is "{labelPrefix} {1-based page number}".
    /// </summary>
    public static string NewSimplePdf(int pageCount, string labelPrefix)
    {
        var path = NewTempPath(".pdf");
        PdfFontResolver.EnsureRegistered();
        using var document = new PdfDocument();
        var font = new XFont("Arial", 20);

        for (var i = 1; i <= pageCount; i++)
        {
            var page = document.AddPage();
            using var gfx = XGraphics.FromPdfPage(page);
            gfx.DrawString($"{labelPrefix} {i}", font, XBrushes.Black, new XPoint(50, 50));
        }

        document.Save(path);
        return path;
    }
}
