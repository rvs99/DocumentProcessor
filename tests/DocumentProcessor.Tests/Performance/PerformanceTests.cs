using System.Diagnostics;
using DocumentProcessor.Core.Redlining;
using DocumentProcessor.Core.Samples;
using DocumentProcessor.Core.Templating;

namespace DocumentProcessor.Tests.Performance;

/// <summary>
/// Regression-guard benchmarks for the capabilities built in this pass (item 9's 10,000-row table
/// benchmark lives with TableGenerationServiceTests instead, since it's more natural alongside that
/// service's other tests). These assert generous ceilings, not tight SLAs — the goal is catching an
/// accidental O(n²) regression, not chasing a specific number. The 30s/512MB/60s floors named in the
/// original ask were specifically about docx→PDF conversion (LibreOffice), which is out of scope
/// here; these targets are sized for the in-scope, pure-managed-code paths instead.
/// </summary>
public class PerformanceTests
{
    [Fact]
    public void TemplateEngine_fills_a_1000_paragraph_document_with_scalar_tokens_quickly()
    {
        var templatePath = TestFiles.NewTempPath(".docx");
        var outputPath = TestFiles.NewTempPath(".docx");
        try
        {
            var paragraphs = Enumerable.Range(0, 1000)
                .Select(i => $"Clause {i}: the party {{{{PartyName}}}} agrees to term {{{{TermValue}}}} on item {i}.");
            SampleDocumentFactory.CreateBasicDocument(templatePath, "Large Template", paragraphs);

            var data = new Dictionary<string, object?> { ["PartyName"] = "Acme Corp", ["TermValue"] = "30 days" };
            var engine = new TemplateEngine();

            var stopwatch = Stopwatch.StartNew();
            engine.Fill(templatePath, outputPath, data);
            stopwatch.Stop();

            Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(15),
                $"Filling 1000 paragraphs (2000 tokens) took {stopwatch.Elapsed}, expected under 15s.");
        }
        finally
        {
            File.Delete(templatePath);
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }

    [Fact]
    public void DocumentComparison_diffs_a_500_paragraph_document_quickly()
    {
        var originalPath = TestFiles.NewTempPath(".docx");
        var revisedPath = TestFiles.NewTempPath(".docx");
        var outputPath = TestFiles.NewTempPath(".docx");
        try
        {
            var original = Enumerable.Range(0, 500).Select(i => $"Paragraph {i}: the original wording for clause {i}.");
            var revised = Enumerable.Range(0, 500).Select(i => i % 3 == 0
                ? $"Paragraph {i}: the REVISED wording for clause {i}, now with more detail."
                : $"Paragraph {i}: the original wording for clause {i}.");

            SampleDocumentFactory.CreateBasicDocument(originalPath, "Large Doc", original);
            SampleDocumentFactory.CreateBasicDocument(revisedPath, "Large Doc", revised);

            var service = new DocumentComparisonService();
            var stopwatch = Stopwatch.StartNew();
            var summary = service.Compare(originalPath, revisedPath, outputPath);
            stopwatch.Stop();

            Assert.True(summary.InsertedCount > 0);
            Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(30),
                $"Comparing 500 paragraphs (~167 changed) took {stopwatch.Elapsed}, expected under 30s.");
        }
        finally
        {
            File.Delete(originalPath);
            File.Delete(revisedPath);
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }

    [Fact]
    public void TemplateEngine_fill_of_a_large_document_does_not_balloon_memory()
    {
        var templatePath = TestFiles.NewTempPath(".docx");
        var outputPath = TestFiles.NewTempPath(".docx");
        try
        {
            var paragraphs = Enumerable.Range(0, 2000).Select(i => $"Item {{{{Index}}}} — {{{{Name}}}} — line {i}.");
            SampleDocumentFactory.CreateBasicDocument(templatePath, "Memory Test", paragraphs);
            var data = new Dictionary<string, object?> { ["Index"] = "1", ["Name"] = "Widget" };

            var process = Process.GetCurrentProcess();
            process.Refresh();
            var before = process.WorkingSet64;

            new TemplateEngine().Fill(templatePath, outputPath, data);

            process.Refresh();
            var after = process.WorkingSet64;
            var deltaMb = (after - before) / (1024.0 * 1024.0);

            Assert.True(deltaMb < 512,
                $"Filling 2000 paragraphs increased working set by {deltaMb:F1}MB, expected under 512MB.");
        }
        finally
        {
            File.Delete(templatePath);
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }
}
