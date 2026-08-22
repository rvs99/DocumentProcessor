using DocumentProcessor.Core.DocumentAssembly;
using UglyToad.PdfPig;

namespace DocumentProcessor.Tests.DocumentAssembly;

public class PdfAssemblyServiceTests : IDisposable
{
    private readonly string _outputPath = TestFiles.NewTempPath(".pdf");
    private readonly PdfAssemblyService _sut = new();
    private readonly List<string> _cleanup = [];

    private string NewPdf(int pageCount, string label)
    {
        var path = TestFiles.NewSimplePdf(pageCount, label);
        _cleanup.Add(path);
        return path;
    }

    private static List<string> PageTexts(string path)
    {
        using var doc = PdfDocument.Open(path);
        return Enumerable.Range(1, doc.NumberOfPages).Select(i => doc.GetPage(i).Text).ToList();
    }

    [Fact]
    public void MergePdfs_concatenates_pages_in_order()
    {
        var a = NewPdf(2, "A");
        var b = NewPdf(3, "B");

        _sut.MergePdfs([a, b], _outputPath);

        var texts = PageTexts(_outputPath);
        Assert.Equal(5, texts.Count);
        Assert.Contains("A 1", texts[0]);
        Assert.Contains("A 2", texts[1]);
        Assert.Contains("B 1", texts[2]);
        Assert.Contains("B 3", texts[4]);
    }

    [Fact]
    public void MergePdfs_requires_at_least_one_document()
    {
        Assert.Throws<ArgumentException>(() => _sut.MergePdfs([], _outputPath));
    }

    [Fact]
    public void ExtractPages_returns_only_the_requested_inclusive_range()
    {
        var source = NewPdf(5, "Page");

        _sut.ExtractPages(source, startPageIndex: 1, endPageIndex: 3, _outputPath);

        var texts = PageTexts(_outputPath);
        Assert.Equal(3, texts.Count);
        Assert.Contains("Page 2", texts[0]);
        Assert.Contains("Page 3", texts[1]);
        Assert.Contains("Page 4", texts[2]);
    }

    [Fact]
    public void ExtractPages_rejects_an_out_of_range_end_index()
    {
        var source = NewPdf(3, "Page");

        Assert.Throws<ArgumentOutOfRangeException>(() => _sut.ExtractPages(source, 0, 10, _outputPath));
    }

    [Fact]
    public void AppendWithContinuedPageNumbers_stamps_a_continuous_sequence_across_both_documents()
    {
        var main = NewPdf(2, "Contract");
        var exhibit = NewPdf(3, "Exhibit");

        _sut.AppendWithContinuedPageNumbers(main, exhibit, _outputPath);

        var texts = PageTexts(_outputPath);
        Assert.Equal(5, texts.Count);
        // Page numbers 1-5 stamped, continuing into the exhibit rather than restarting at 1.
        for (var i = 0; i < 5; i++)
            Assert.Contains((i + 1).ToString(), texts[i]);
        Assert.Contains("Exhibit 1", texts[2]);
    }

    public void Dispose()
    {
        foreach (var path in _cleanup)
            File.Delete(path);
        if (File.Exists(_outputPath))
            File.Delete(_outputPath);
    }
}
