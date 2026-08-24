using DocumentProcessor.Core.Conversion;
using DocumentProcessor.Core.Diagnostics;
using DocumentProcessor.Core.TrackChanges;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;

namespace DocumentProcessor.Core.Redlining;

/// <summary>The four standard deliverables of a redline review cycle.</summary>
public sealed record RedlineExportPaths(string CleanDocx, string RedlinedDocx, string CleanPdf, string RedlinedPdf);

/// <summary>
/// Produces all four standard redline deliverables in one call — Redlined DOCX (tracked changes,
/// for the counterparty to review in Word), Clean DOCX (changes accepted, for signature), and a PDF
/// of each — rather than requiring the caller to hand-wire <see cref="DocumentComparisonService"/>,
/// <see cref="TrackChangesService"/>, and <see cref="WordToPdfConverter"/> themselves. Converting a
/// still-redlined document to PDF works the same way as any other docx→PDF conversion: the
/// w:ins/w:del markup renders as Word itself would show it (underline/strikethrough), since that's
/// standard OOXML any compliant renderer honors — this just wires up the previously-unexercised path.
/// </summary>
public sealed class RedlineExportService(ILogger<RedlineExportService>? logger = null) : IRedlineExportService
{
    private readonly ILogger<RedlineExportService> _logger = logger ?? NullLogger<RedlineExportService>.Instance;

    public RedlineExportPaths ExportAllVariants(
        string originalPath,
        string revisedPath,
        string outputDirectory,
        string authorForRevisions = "Document Comparison",
        WordToPdfConversionOptions? conversionOptions = null)
    {
        using var activity = DocumentProcessorDiagnostics.ActivitySource.StartActivity("RedlineExportService.ExportAllVariants");
        _logger.LogInformation("Exporting all four redline variants for {OriginalPath} vs {RevisedPath} into {OutputDirectory}",
            originalPath, revisedPath, outputDirectory);

        Directory.CreateDirectory(outputDirectory);

        var redlinedDocx = Path.Combine(outputDirectory, "redlined.docx");
        var cleanDocx = Path.Combine(outputDirectory, "clean.docx");
        var redlinedPdf = Path.Combine(outputDirectory, "redlined.pdf");
        var cleanPdf = Path.Combine(outputDirectory, "clean.pdf");

        new DocumentComparisonService().Compare(originalPath, revisedPath, redlinedDocx, authorForRevisions);
        _logger.LogDebug("Redlined DOCX written to {RedlinedDocx}", redlinedDocx);

        File.Copy(redlinedDocx, cleanDocx, overwrite: true);
        new TrackChangesService().AcceptAll(cleanDocx);
        _logger.LogDebug("Clean DOCX written to {CleanDocx}", cleanDocx);

        var converter = new WordToPdfConverter(conversionOptions);
        converter.Convert(redlinedDocx, redlinedPdf);
        _logger.LogDebug("Redlined PDF written to {RedlinedPdf}", redlinedPdf);
        converter.Convert(cleanDocx, cleanPdf);
        _logger.LogDebug("Clean PDF written to {CleanPdf}", cleanPdf);

        _logger.LogInformation("Finished exporting all four redline variants into {OutputDirectory}", outputDirectory);
        return new RedlineExportPaths(cleanDocx, redlinedDocx, cleanPdf, redlinedPdf);
    }
}
