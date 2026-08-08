using DocumentProcessor.Core.Comparison;

namespace DocumentProcessor.Tests.Comparison;

public class PdfComparisonServiceTests : IDisposable
{
    private readonly string _pdfA = TestFiles.NewTestPdf("Contract", ["The fee is one thousand dollars.", "This clause is unchanged."]);
    private readonly string _pdfB = TestFiles.NewTestPdf("Contract", ["The fee is two thousand dollars.", "This clause is unchanged."]);
    private readonly PdfComparisonService _sut = new();

    [Fact]
    public void CompareText_detects_the_changed_wording()
    {
        var result = _sut.CompareText(_pdfA, _pdfB);

        Assert.True(result.HasDifferences);
        Assert.Contains(result.DeletedWords, w => w.Contains("one"));
        Assert.Contains(result.InsertedWords, w => w.Contains("two"));
    }

    [Fact]
    public void CompareText_of_identical_pdfs_reports_no_differences()
    {
        var result = _sut.CompareText(_pdfA, _pdfA);

        Assert.False(result.HasDifferences);
        Assert.Empty(result.InsertedWords);
        Assert.Empty(result.DeletedWords);
    }

    [Fact]
    public void CompareVisual_flags_pages_that_differ_and_matches_identical_ones()
    {
        // A one-word change only moves a small fraction of a full page's pixels, so use a
        // threshold tight enough to catch it rather than the default "meaningfully different" cutoff.
        var differing = _sut.CompareVisual(_pdfA, _pdfB, differenceThresholdPercent: 0.0001);
        Assert.False(differing.PagesMatch);
        Assert.True(differing.PageCountMatches);
        Assert.True(differing.PerPageDifferencePercent.Single() > 0);

        var identical = _sut.CompareVisual(_pdfA, _pdfA);
        Assert.True(identical.PagesMatch);
        Assert.Equal(0, identical.PerPageDifferencePercent.Single());
    }

    public void Dispose()
    {
        File.Delete(_pdfA);
        File.Delete(_pdfB);
    }
}
