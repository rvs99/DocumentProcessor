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
public sealed class PdfComparisonService
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
        // PDFium's platform list is scoped to the platforms it ships native binaries for (which
        // covers Windows/Linux/macOS — everything this service targets); the analyzer just flags
        // that the list excludes platforms like wasm/tvOS that are irrelevant here.
#pragma warning disable CA1416
        var imagesA = PDFtoImage.Conversion.ToImages(File.ReadAllBytes(pdfPathA), password: null, options).ToList();
        var imagesB = PDFtoImage.Conversion.ToImages(File.ReadAllBytes(pdfPathB), password: null, options).ToList();
#pragma warning restore CA1416

        try
        {
            var pageCountMatches = imagesA.Count == imagesB.Count;
            var comparedPageCount = Math.Min(imagesA.Count, imagesB.Count);

            var perPageDiffs = new List<double>();
            for (var i = 0; i < comparedPageCount; i++)
                perPageDiffs.Add(ComputePixelDifferencePercent(imagesA[i], imagesB[i]));

            var pagesMatch = pageCountMatches && perPageDiffs.All(d => d <= differenceThresholdPercent);
            return new PdfVisualDiffResult(pagesMatch, comparedPageCount, pageCountMatches, perPageDiffs);
        }
        finally
        {
            foreach (var image in imagesA) image.Dispose();
            foreach (var image in imagesB) image.Dispose();
        }
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

        var bytesA = a.Bytes;
        var bytesB = b.Bytes;
        long differingBytes = 0;
        for (var i = 0; i < bytesA.Length; i++)
        {
            if (bytesA[i] != bytesB[i])
                differingBytes++;
        }

        return 100.0 * differingBytes / bytesA.Length;
    }
}
