using System.Diagnostics;
using DocumentProcessor.Core.Conversion;
using DocumentProcessor.Core.Samples;
using UglyToad.PdfPig;

namespace DocumentProcessor.Tests.Conversion;

/// <summary>
/// Conversions are the expensive part of this library — LibreOffice spends roughly 450 ms starting
/// before it looks at the first document — so several documents now share one invocation. These
/// cover the correctness that sharing puts at risk: results landing in the wrong caller's file, one
/// bad document taking down the batch, and a pooled profile outliving a killed process.
/// </summary>
public class BatchConversionTests : IDisposable
{
    private readonly List<string> _cleanup = [];

    private string NewDocx(string fileName, string body)
    {
        var dir = Path.Combine(Path.GetTempPath(), $"docproc-batch-{Guid.NewGuid():N}");
        Directory.CreateDirectory(dir);
        _cleanup.Add(dir);

        var path = Path.Combine(dir, fileName);
        SampleDocumentFactory.CreateBasicDocument(path, body, [body]);
        return path;
    }

    private string NewOutputPath()
    {
        var path = TestFiles.NewTempPath(".pdf");
        _cleanup.Add(path);
        return path;
    }

    /// <summary>
    /// A file LibreOffice genuinely refuses. It has to start with a ZIP local-file-header signature:
    /// LibreOffice cheerfully converts arbitrary bytes as a plain-text document, so only something
    /// that claims to be a package and then isn't one actually fails to load.
    /// </summary>
    private string NewUnreadableDocx(string fileName)
    {
        var dir = Path.Combine(Path.GetTempPath(), $"docproc-batch-{Guid.NewGuid():N}");
        Directory.CreateDirectory(dir);
        _cleanup.Add(dir);

        var path = Path.Combine(dir, fileName);
        File.WriteAllBytes(path, [0x50, 0x4B, 0x03, 0x04, .. "truncated package"u8]);
        return path;
    }

    private static string TextOf(string pdfPath)
    {
        using var doc = PdfDocument.Open(pdfPath);
        return string.Join(" ", doc.GetPages().Select(p => p.Text));
    }

    private static WordToPdfConversionOptions Options(Action<WordToPdfConversionOptions>? configure = null)
    {
        var options = TestFiles.ConversionOptions();
        configure?.Invoke(options);
        return options;
    }

    [Fact]
    public async Task A_batch_converts_every_document()
    {
        var requests = Enumerable.Range(0, 5)
            .Select(i => new ConversionRequest(NewDocx($"doc{i}.docx", $"Document number {i}"), NewOutputPath()))
            .ToList();

        var results = await new WordToPdfConverter(Options()).ConvertBatchAsync(requests);

        Assert.All(results, r => Assert.True(r.Succeeded, r.Error?.Message));
        for (var i = 0; i < requests.Count; i++)
        {
            Assert.True(File.Exists(requests[i].OutputPdfPath));
            Assert.Contains($"Document number {i}", TextOf(requests[i].OutputPdfPath));
        }
    }

    [Fact]
    public async Task Documents_sharing_a_file_name_do_not_overwrite_each_other()
    {
        // LibreOffice names its output after the input's base name, so a batch containing two files
        // both called contract.docx — the single likeliest collision in a system where customers
        // name their own uploads — would otherwise produce one PDF and hand it to both callers.
        var first = new ConversionRequest(NewDocx("contract.docx", "First party agreement"), NewOutputPath());
        var second = new ConversionRequest(NewDocx("contract.docx", "Second party agreement"), NewOutputPath());

        var results = await new WordToPdfConverter(Options()).ConvertBatchAsync([first, second]);

        Assert.All(results, r => Assert.True(r.Succeeded, r.Error?.Message));
        Assert.Contains("First party agreement", TextOf(first.OutputPdfPath));
        Assert.Contains("Second party agreement", TextOf(second.OutputPdfPath));
    }

    [Fact]
    public async Task One_unreadable_document_does_not_deny_the_others_their_results()
    {
        var good = new ConversionRequest(NewDocx("good.docx", "A readable contract"), NewOutputPath());

        var bad = new ConversionRequest(NewUnreadableDocx("corrupt.docx"), NewOutputPath());

        var results = await new WordToPdfConverter(Options()).ConvertBatchAsync([good, bad]);

        Assert.True(results.Single(r => r.Request == good).Succeeded);
        Assert.Contains("A readable contract", TextOf(good.OutputPdfPath));

        var failure = results.Single(r => r.Request == bad);
        Assert.False(failure.Succeeded);
        Assert.NotNull(failure.Error);
    }

    [Fact]
    public async Task A_batch_larger_than_the_batch_size_is_chunked_and_still_complete()
    {
        var requests = Enumerable.Range(0, 7)
            .Select(i => new ConversionRequest(NewDocx($"chunked{i}.docx", $"Chunk {i}"), NewOutputPath()))
            .ToList();

        var results = await new WordToPdfConverter(Options(o => o.MaxBatchSize = 3)).ConvertBatchAsync(requests);

        Assert.Equal(7, results.Count);
        Assert.All(results, r => Assert.True(r.Succeeded, r.Error?.Message));
        for (var i = 0; i < requests.Count; i++)
            Assert.Contains($"Chunk {i}", TextOf(requests[i].OutputPdfPath));
    }

    [Fact]
    public async Task An_empty_batch_does_no_work()
    {
        Assert.Empty(await new WordToPdfConverter(Options()).ConvertBatchAsync([]));
    }

    [Fact]
    public async Task A_missing_input_is_rejected_before_any_conversion_starts()
    {
        var present = new ConversionRequest(NewDocx("present.docx", "Here"), NewOutputPath());
        var absent = new ConversionRequest(TestFiles.NewTempPath(".docx"), NewOutputPath());

        await Assert.ThrowsAsync<FileNotFoundException>(() =>
            new WordToPdfConverter(Options()).ConvertBatchAsync([present, absent]));

        Assert.False(File.Exists(present.OutputPdfPath));
    }

    [Fact]
    public async Task Concurrent_conversions_coalesce_without_crossing_results()
    {
        // The real risk of coalescing: eight independent callers, one shared invocation, and each
        // one must still receive its own document rather than a neighbour's.
        var converter = new WordToPdfConverter(Options());
        var pairs = Enumerable.Range(0, 8)
            .Select(i => (Input: NewDocx($"concurrent{i}.docx", $"Tenant {i} agreement"), Output: NewOutputPath()))
            .ToList();

        await Task.WhenAll(pairs.Select(p => converter.ConvertAsync(p.Input, p.Output)));

        for (var i = 0; i < pairs.Count; i++)
        {
            var text = TextOf(pairs[i].Output);
            Assert.Contains($"Tenant {i} agreement", text);

            // Not merely "contains mine" — no other tenant's document may have landed here.
            foreach (var other in Enumerable.Range(0, pairs.Count).Where(j => j != i))
                Assert.DoesNotContain($"Tenant {other} agreement", text);
        }
    }

    [Fact]
    public async Task Coalescing_can_be_turned_off_without_changing_the_result()
    {
        var converter = new WordToPdfConverter(Options(o => o.EnableBatching = false));
        var pairs = Enumerable.Range(0, 3)
            .Select(i => (Input: NewDocx($"unbatched{i}.docx", $"Solo {i}"), Output: NewOutputPath()))
            .ToList();

        await Task.WhenAll(pairs.Select(p => converter.ConvertAsync(p.Input, p.Output)));

        for (var i = 0; i < pairs.Count; i++)
            Assert.Contains($"Solo {i}", TextOf(pairs[i].Output));
    }

    [Fact]
    public async Task A_single_conversion_still_reports_its_own_failure()
    {
        // Batching must not swallow a solo caller's error into a per-item result they never see.
        // Note LibreOffice exits 0 even here, so the failure is detected by the absent output, not
        // by the exit code.
        await Assert.ThrowsAnyAsync<Exception>(() =>
            new WordToPdfConverter(Options()).ConvertAsync(NewUnreadableDocx("solo-corrupt.docx"), NewOutputPath()));
    }

    [Fact]
    public async Task Cancelling_a_solo_conversion_still_stops_it()
    {
        using var cts = new CancellationTokenSource();
        await cts.CancelAsync();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            new WordToPdfConverter(Options()).ConvertAsync(NewDocx("cancelled.docx", "Never converted"), NewOutputPath(), cts.Token));
    }

    [Fact]
    public async Task Batching_converts_eight_documents_faster_than_converting_them_one_by_one()
    {
        var oneByOne = Enumerable.Range(0, 8)
            .Select(i => new ConversionRequest(NewDocx($"slow{i}.docx", $"Sequential {i}"), NewOutputPath()))
            .ToList();
        var together = Enumerable.Range(0, 8)
            .Select(i => new ConversionRequest(NewDocx($"fast{i}.docx", $"Batched {i}"), NewOutputPath()))
            .ToList();

        var unbatched = new WordToPdfConverter(Options(o => o.EnableBatching = false));
        var sequentialTimer = Stopwatch.StartNew();
        foreach (var request in oneByOne)
            await unbatched.ConvertAsync(request.DocxPath, request.OutputPdfPath);
        sequentialTimer.Stop();

        var batchTimer = Stopwatch.StartNew();
        await new WordToPdfConverter(Options()).ConvertBatchAsync(together);
        batchTimer.Stop();

        Assert.True(batchTimer.Elapsed < sequentialTimer.Elapsed,
            $"batch {batchTimer.ElapsedMilliseconds} ms vs one-at-a-time {sequentialTimer.ElapsedMilliseconds} ms for eight documents.");
    }

    public void Dispose()
    {
        foreach (var path in _cleanup)
        {
            try
            {
                if (Directory.Exists(path)) Directory.Delete(path, recursive: true);
                else if (File.Exists(path)) File.Delete(path);
            }
            catch { /* temp files */ }
        }
    }
}
