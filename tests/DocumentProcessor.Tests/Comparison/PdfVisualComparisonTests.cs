using DocumentProcessor.Core.Comparison;

namespace DocumentProcessor.Tests.Comparison;

/// <summary>
/// Exercises the page-streaming visual comparison against PDFs built directly with PDFsharp, so it
/// runs without a LibreOffice install. The pre-existing <see cref="PdfComparisonServiceTests"/>
/// fixtures go through docx→PDF conversion, which means they are skipped wherever LibreOffice is
/// absent — leaving the comparison path itself unverified in exactly the environments where it
/// matters most.
/// </summary>
public class PdfVisualComparisonTests : IDisposable
{
    private readonly List<string> _cleanup = [];
    private readonly PdfComparisonService _sut = new();

    private string NewPdf(int pageCount, string label)
    {
        var path = TestFiles.NewSimplePdf(pageCount, label);
        _cleanup.Add(path);
        return path;
    }

    [Fact]
    public void Identical_documents_report_matching_pages_and_zero_difference()
    {
        var a = NewPdf(3, "Contract");
        var b = NewPdf(3, "Contract");

        var result = _sut.CompareVisual(a, b);

        Assert.True(result.PageCountMatches);
        Assert.True(result.PagesMatch);
        Assert.Equal(3, result.ComparedPageCount);
        Assert.All(result.PerPageDifferencePercent, d => Assert.Equal(0.0, d, precision: 6));
    }

    [Fact]
    public void Differing_content_is_reported_per_page()
    {
        var a = NewPdf(2, "Original");
        var b = NewPdf(2, "Amended");

        var result = _sut.CompareVisual(a, b);

        Assert.True(result.PageCountMatches);
        Assert.Equal(2, result.PerPageDifferencePercent.Count);
        Assert.All(result.PerPageDifferencePercent, d => Assert.True(d > 0, $"expected a visible difference, got {d}%"));
    }

    [Fact]
    public void The_threshold_decides_whether_a_small_change_counts_as_matching()
    {
        var a = NewPdf(1, "Original");
        var b = NewPdf(1, "Amended");

        // A short text edit on an otherwise blank page changes only a tiny fraction of the pixels,
        // so it falls under the 0.5% default and the pages are reported as matching. That default
        // is a similarity tolerance, not a change detector — callers who need "any visible edit at
        // all" have to say so, which is worth pinning explicitly since the distinction is easy to
        // get wrong when wiring this into a review workflow.
        Assert.True(_sut.CompareVisual(a, b).PagesMatch);
        Assert.False(_sut.CompareVisual(a, b, differenceThresholdPercent: 0.0).PagesMatch);
    }

    [Fact]
    public void A_page_count_mismatch_is_reported_and_only_the_common_pages_are_compared()
    {
        var a = NewPdf(2, "Contract");
        var b = NewPdf(5, "Contract");

        var result = _sut.CompareVisual(a, b);

        Assert.False(result.PageCountMatches);
        Assert.False(result.PagesMatch);
        Assert.Equal(2, result.ComparedPageCount);
        Assert.Equal(2, result.PerPageDifferencePercent.Count);
    }

    [Fact]
    public void Comparing_a_long_document_does_not_hold_every_page_in_memory()
    {
        // 40 pages at 150 DPI would be ~670 MB if both page sets were materialized up front, as
        // the previous implementation did. Streaming keeps the peak at two pages, so the working
        // set should barely move. The threshold is deliberately loose — this is a guard against
        // reintroducing whole-document materialization, not a precise memory assertion.
        var a = NewPdf(40, "Long");
        var b = NewPdf(40, "Long");

        var before = GC.GetTotalMemory(forceFullCollection: true);
        var result = _sut.CompareVisual(a, b);
        var after = GC.GetTotalMemory(forceFullCollection: true);

        Assert.Equal(40, result.ComparedPageCount);
        var growthMb = (after - before) / (1024.0 * 1024.0);
        Assert.True(growthMb < 200, $"managed heap grew {growthMb:F1} MB comparing 40 pages — expected streaming behaviour.");
    }

    public void Dispose()
    {
        foreach (var path in _cleanup)
            File.Delete(path);
    }
}
