using DiffPlex.Chunkers;
using DiffPlex.DiffBuilder;
using DiffPlex.DiffBuilder.Model;
using PDFtoImage;
using UglyToad.PdfPig;

namespace DocumentProcessor.Core.Comparison;

public sealed record PdfTextDiffResult(bool HasDifferences, IReadOnlyList<string> InsertedWords, IReadOnlyList<string> DeletedWords);

public sealed record PdfVisualDiffResult(bool PagesMatch, int ComparedPageCount, bool PageCountMatches, IReadOnlyList<double> PerPageDifferencePercent);

/// <summary>
/// Compares two PDFs two ways: a text-content diff (what changed, word-by-word — via PdfPig text
/// extraction and DiffPlex), and a visual/pixel diff per page (did the rendered layout change at
/// all, even for content that isn't plain text — via PDFtoImage/PDFium rasterization and a
/// byte-level pixel comparison). Useful together: text diff catches wording changes, visual diff
/// catches formatting/layout/image changes that don't show up in extracted text.
/// </summary>
public sealed class PdfComparisonService : IPdfComparisonService
{
    public PdfTextDiffResult CompareText(string pdfPathA, string pdfPathB)
    {
        var textA = ExtractText(pdfPathA);
        var textB = ExtractText(pdfPathB);

        var diff = InlineDiffBuilder.Diff(textA, textB, false, false, new WordChunker());

        var inserted = diff.Lines.Where(l => l.Type == ChangeType.Inserted).Select(l => l.Text).ToList();
        var deleted = diff.Lines.Where(l => l.Type == ChangeType.Deleted).Select(l => l.Text).ToList();

        return new PdfTextDiffResult(diff.HasDifferences, inserted, deleted);
    }

    /// <summary>
    /// Renders each PDF page to a bitmap and compares them byte-for-byte, flagging pages whose
    /// differing-pixel percentage exceeds <paramref name="differenceThresholdPercent"/>.
    /// </summary>
    public PdfVisualDiffResult CompareVisual(string pdfPathA, string pdfPathB, double differenceThresholdPercent = 0.5, int dpi = 150)
    {
        var options = new RenderOptions(Dpi: dpi);
        var bytesA = File.ReadAllBytes(pdfPathA);
        var bytesB = File.ReadAllBytes(pdfPathB);

        // Page counts come from PdfPig rather than by rendering, so nothing is rasterized just to
        // find out how many pages there are.
        int pageCountA, pageCountB;
        using (var docA = PdfDocument.Open(bytesA)) pageCountA = docA.NumberOfPages;
        using (var docB = PdfDocument.Open(bytesB)) pageCountB = docB.NumberOfPages;

        var pageCountMatches = pageCountA == pageCountB;
        var comparedPageCount = Math.Min(pageCountA, pageCountB);
        var perPageDiffs = new List<double>(comparedPageCount);

        // Render one page from each document, compare, dispose, advance. Materializing both page
        // sets up front (the previous approach) held every page of both documents live at once —
        // at 150 DPI a Letter page is ~8.4 MB of unmanaged Skia memory, so a 200-page pair was
        // ~3.4 GB resident, invisible to GC pressure heuristics and reliably OOM-killed. Peak here
        // is two pages regardless of document length.
        for (var i = 0; i < comparedPageCount; i++)
        {
            // PDFium's platform list is scoped to the platforms it ships native binaries for (which
            // covers Windows/Linux/macOS — everything this service targets); the analyzer just flags
            // that the list excludes platforms like wasm/tvOS that are irrelevant here.
#pragma warning disable CA1416
            using var pageA = PDFtoImage.Conversion.ToImage(bytesA, page: i, password: null, options: options);
            using var pageB = PDFtoImage.Conversion.ToImage(bytesB, page: i, password: null, options: options);
#pragma warning restore CA1416
            perPageDiffs.Add(ComputePixelDifferencePercent(pageA, pageB));
        }

        var pagesMatch = pageCountMatches && perPageDiffs.All(d => d <= differenceThresholdPercent);
        return new PdfVisualDiffResult(pagesMatch, comparedPageCount, pageCountMatches, perPageDiffs);
    }

    private static string ExtractText(string pdfPath)
    {
        using var document = PdfDocument.Open(pdfPath);
        return string.Join('\n', Enumerable.Range(1, document.NumberOfPages).Select(i => document.GetPage(i).Text));
    }

    private static double ComputePixelDifferencePercent(SkiaSharp.SKBitmap a, SkiaSharp.SKBitmap b)
    {
        if (a.Width != b.Width || a.Height != b.Height)
            return 100.0;

        // GetPixelSpan reads the existing buffer in place. SKBitmap.Bytes (the previous approach)
        // allocates and copies the whole pixel buffer per page — two ~8.4 MB large-object-heap
        // allocations per page pair, purely to read them once.
        var spanA = a.GetPixelSpan();
        var spanB = b.GetPixelSpan();
        if (spanA.Length != spanB.Length)
            return 100.0;

        long differingBytes = 0;
        for (var i = 0; i < spanA.Length; i++)
        {
            if (spanA[i] != spanB[i])
                differingBytes++;
        }

        return spanA.Length == 0 ? 0.0 : 100.0 * differingBytes / spanA.Length;
    }
}
