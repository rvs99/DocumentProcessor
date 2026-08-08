using DocumentProcessor.Core.Conversion;
using DocumentProcessor.Core.Samples;

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
}
