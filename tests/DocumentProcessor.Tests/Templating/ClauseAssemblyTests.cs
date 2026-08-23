using System.Diagnostics;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.Samples;
using DocumentProcessor.Core.Templating;
using DocumentProcessor.Core.Transplant;

namespace DocumentProcessor.Tests.Templating;

/// <summary>
/// Multi-marker clause assembly. The single-marker path was already covered, but the cost and the
/// ordering risk both only appear with several markers — resolution used to transplant one clause
/// per pass over a document that grew each time, which was quadratic as well as slow.
/// </summary>
public class ClauseAssemblyTests : IDisposable
{
    private readonly string _libraryPath = TestFiles.NewTempPath(".docx");
    private readonly string _templatePath = TestFiles.NewTempPath(".docx");
    private readonly string _outputPath = TestFiles.NewTempPath(".docx");
    private readonly TemplateEngine _sut = new();

    private static Paragraph Literal(string text) => new(new Run(new Text(text)));

    private void CreateLibrary(int clauseCount)
    {
        var paragraphs = new List<Paragraph>();
        for (var i = 0; i < clauseCount; i++)
        {
            paragraphs.Add(new Paragraph(
                new BookmarkStart { Id = i.ToString(), Name = $"clause_c{i}" },
                new Run(new Text($"Clause C{i} body text for {{{{Client}}}}.")),
                new BookmarkEnd { Id = i.ToString() }));
        }

        SampleDocumentFactory.CreateDocumentFromParagraphs(_libraryPath, paragraphs);
    }

    private IReadOnlyList<string> OutputParagraphs() =>
        new ClauseTransplantService().ListParagraphs(_outputPath).Select(p => p.Text).ToList();

    [Fact]
    public void Every_marker_is_resolved_and_document_order_is_preserved()
    {
        CreateLibrary(5);
        SampleDocumentFactory.CreateDocumentFromParagraphs(_templatePath,
        [
            Literal("Preamble."),
            Literal("{{clause:c0}}"),
            Literal("Between one and two."),
            Literal("{{clause:c1}}"),
            Literal("{{clause:c2}}"),
            Literal("Tail.")
        ]);

        _sut.Fill(_templatePath, _outputPath, new Dictionary<string, object?> { ["Client"] = "Acme" },
            MissingTokenPolicy.Error, new ClauseLibrary(_libraryPath));

        var texts = OutputParagraphs();
        Assert.Equal(
        [
            "Preamble.",
            "Clause C0 body text for Acme.",
            "Between one and two.",
            "Clause C1 body text for Acme.",
            "Clause C2 body text for Acme.",
            "Tail."
        ], texts);
    }

    [Fact]
    public void Tokens_inside_injected_clauses_are_substituted()
    {
        CreateLibrary(2);
        SampleDocumentFactory.CreateDocumentFromParagraphs(_templatePath, [Literal("{{clause:c0}}"), Literal("{{clause:c1}}")]);

        _sut.Fill(_templatePath, _outputPath, new Dictionary<string, object?> { ["Client"] = "Globex" },
            MissingTokenPolicy.Error, new ClauseLibrary(_libraryPath));

        Assert.All(OutputParagraphs(), t => Assert.Contains("Globex", t));
        Assert.DoesNotContain(OutputParagraphs(), t => t.Contains("{{"));
    }

    [Fact]
    public void A_missing_clause_among_present_ones_is_redacted_without_disturbing_the_rest()
    {
        CreateLibrary(2);
        SampleDocumentFactory.CreateDocumentFromParagraphs(_templatePath,
            [Literal("{{clause:c0}}"), Literal("{{clause:nope}}"), Literal("{{clause:c1}}")]);

        _sut.Fill(_templatePath, _outputPath, new Dictionary<string, object?> { ["Client"] = "Acme" },
            MissingTokenPolicy.Redact, new ClauseLibrary(_libraryPath));

        var texts = OutputParagraphs();
        Assert.Equal(["Clause C0 body text for Acme.", "Clause C1 body text for Acme."], texts);
    }

    [Fact]
    public void A_missing_clause_is_highlighted_in_place_under_the_Highlight_policy()
    {
        CreateLibrary(1);
        SampleDocumentFactory.CreateDocumentFromParagraphs(_templatePath,
            [Literal("{{clause:c0}}"), Literal("{{clause:absent}}")]);

        var result = _sut.Fill(_templatePath, _outputPath, new Dictionary<string, object?> { ["Client"] = "Acme" },
            MissingTokenPolicy.Highlight, new ClauseLibrary(_libraryPath));

        var texts = OutputParagraphs();
        Assert.Contains(texts, t => t.Contains("Clause C0"));
        Assert.Contains(texts, t => t.Contains("[Missing clause: absent]"));
        Assert.Contains(result.Warnings, w => w.Contains("absent"));
    }

    [Fact]
    public void Resolving_many_markers_stays_linear_rather_than_quadratic()
    {
        // 30 markers previously meant ~360 full document parses and ~120 re-serialisations, on a
        // document growing with every pass. A generous ceiling: the point is to catch a return to
        // per-marker assembly, not to pin a machine-specific number.
        const int ClauseCount = 30;
        CreateLibrary(ClauseCount);
        SampleDocumentFactory.CreateDocumentFromParagraphs(_templatePath,
            Enumerable.Range(0, ClauseCount).Select(i => Literal($"{{{{clause:c{i}}}}}")).ToList());

        var stopwatch = Stopwatch.StartNew();
        _sut.Fill(_templatePath, _outputPath, new Dictionary<string, object?> { ["Client"] = "Acme" },
            MissingTokenPolicy.Error, new ClauseLibrary(_libraryPath));
        stopwatch.Stop();

        var texts = OutputParagraphs();
        Assert.Equal(ClauseCount, texts.Count(t => t.StartsWith("Clause C")));
        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(20),
            $"resolving {ClauseCount} clause markers took {stopwatch.Elapsed} — expected single-pass assembly.");
    }

    public void Dispose()
    {
        foreach (var path in new[] { _libraryPath, _templatePath, _outputPath })
            if (File.Exists(path)) File.Delete(path);
    }
}
