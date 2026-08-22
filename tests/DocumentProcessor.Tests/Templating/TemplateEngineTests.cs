using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.Samples;
using DocumentProcessor.Core.Templating;

namespace DocumentProcessor.Tests.Templating;

public class TemplateEngineTests : IDisposable
{
    private readonly string _templatePath = TestFiles.NewTempPath(".docx");
    private readonly string _outputPath = TestFiles.NewTempPath(".docx");
    private readonly TemplateEngine _sut = new();

    private static Paragraph Literal(string text) => new(new Run(new Text(text)));

    private static IReadOnlyList<Paragraph> ReadParagraphs(string path)
    {
        using var doc = WordprocessingDocument.Open(path, isEditable: false);
        return doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().Select(p => (Paragraph)p.CloneNode(true)).ToList();
    }

    [Fact]
    public void Fill_substitutes_a_token_split_across_three_runs()
    {
        // Word routinely fragments a single logical piece of text across several runs at
        // spellcheck/revision boundaries — simulate that directly rather than relying on Word
        // to have done it for us.
        var paragraph = new Paragraph(
            new Run(new Text("Dear {{Cli")),
            new Run(new Text("entNa")),
            new Run(new Text("me}}, welcome.")));
        SampleDocumentFactory.CreateDocumentFromParagraphs(_templatePath, [paragraph]);

        _sut.Fill(_templatePath, _outputPath, new Dictionary<string, object?> { ["ClientName"] = "Acme Corp" });

        var result = ReadParagraphs(_outputPath);
        Assert.Single(result);
        Assert.Equal("Dear Acme Corp, welcome.", result[0].InnerText);
    }

    [Fact]
    public void Fill_resolves_a_nested_dotted_field_path()
    {
        SampleDocumentFactory.CreateDocumentFromParagraphs(_templatePath, [Literal("Client: {{Company.Name}}")]);

        var data = new Dictionary<string, object?>
        {
            ["Company"] = new Dictionary<string, object?> { ["Name"] = "Acme Corp" }
        };
        _sut.Fill(_templatePath, _outputPath, data);

        Assert.Equal("Client: Acme Corp", ReadParagraphs(_outputPath)[0].InnerText);
    }

    [Fact]
    public void Fill_leaves_the_template_file_untouched()
    {
        SampleDocumentFactory.CreateDocumentFromParagraphs(_templatePath, [Literal("Hello {{Name}}")]);

        _sut.Fill(_templatePath, _outputPath, new Dictionary<string, object?> { ["Name"] = "World" });

        Assert.Equal("Hello {{Name}}", ReadParagraphs(_templatePath)[0].InnerText);
        Assert.Equal("Hello World", ReadParagraphs(_outputPath)[0].InnerText);
    }

    [Fact]
    public void Fill_with_Error_policy_throws_on_a_missing_field()
    {
        SampleDocumentFactory.CreateDocumentFromParagraphs(_templatePath, [Literal("Hello {{Missing}}")]);

        var ex = Assert.Throws<MissingTemplateTokenException>(() =>
            _sut.Fill(_templatePath, _outputPath, new Dictionary<string, object?>()));
        Assert.Equal("Missing", ex.FieldPath);
    }

    [Fact]
    public void Fill_with_Redact_policy_blanks_a_missing_field()
    {
        SampleDocumentFactory.CreateDocumentFromParagraphs(_templatePath, [Literal("Hello {{Missing}}!")]);

        _sut.Fill(_templatePath, _outputPath, new Dictionary<string, object?>(), MissingTokenPolicy.Redact);

        Assert.Equal("Hello !", ReadParagraphs(_outputPath)[0].InnerText);
    }

    [Fact]
    public void Fill_with_Highlight_policy_keeps_the_token_visible_and_highlighted()
    {
        SampleDocumentFactory.CreateDocumentFromParagraphs(_templatePath, [Literal("Hello {{Missing}}!")]);

        _sut.Fill(_templatePath, _outputPath, new Dictionary<string, object?>(), MissingTokenPolicy.Highlight);

        var result = ReadParagraphs(_outputPath)[0];
        Assert.Equal("Hello {{Missing}}!", result.InnerText);
        var highlightedRun = result.Elements<Run>().First(r => r.InnerText == "{{Missing}}");
        Assert.NotNull(highlightedRun.RunProperties?.GetFirstChild<Highlight>());
    }

    [Fact]
    public void Fill_evaluates_if_else_and_keeps_only_the_chosen_branch()
    {
        SampleDocumentFactory.CreateDocumentFromParagraphs(_templatePath,
        [
            Literal("{{if:Amount > 1000}}"),
            Literal("Large order."),
            Literal("{{else}}"),
            Literal("Small order."),
            Literal("{{/if}}")
        ]);

        _sut.Fill(_templatePath, _outputPath, new Dictionary<string, object?> { ["Amount"] = "5000" });
        var result = ReadParagraphs(_outputPath);
        Assert.Single(result);
        Assert.Equal("Large order.", result[0].InnerText);

        _sut.Fill(_templatePath, _outputPath, new Dictionary<string, object?> { ["Amount"] = "10" });
        result = ReadParagraphs(_outputPath);
        Assert.Single(result);
        Assert.Equal("Small order.", result[0].InnerText);
    }

    [Fact]
    public void Fill_supports_three_levels_of_nested_conditionals()
    {
        SampleDocumentFactory.CreateDocumentFromParagraphs(_templatePath,
        [
            Literal("{{if:A == \"yes\"}}"),
                Literal("A-yes"),
                Literal("{{if:B == \"yes\"}}"),
                    Literal("B-yes"),
                    Literal("{{if:C == \"yes\"}}"),
                        Literal("C-yes"),
                    Literal("{{else}}"),
                        Literal("C-no"),
                    Literal("{{/if}}"),
                Literal("{{else}}"),
                    Literal("B-no"),
                Literal("{{/if}}"),
            Literal("{{else}}"),
                Literal("A-no"),
            Literal("{{/if}}")
        ]);

        var data = new Dictionary<string, object?> { ["A"] = "yes", ["B"] = "yes", ["C"] = "no" };
        _sut.Fill(_templatePath, _outputPath, data);

        var texts = ReadParagraphs(_outputPath).Select(p => p.InnerText).ToList();
        Assert.Equal(["A-yes", "B-yes", "C-no"], texts);
    }

    [Fact]
    public void Fill_expands_a_repeat_block_once_per_item_with_outer_scope_fallback()
    {
        SampleDocumentFactory.CreateDocumentFromParagraphs(_templatePath,
        [
            Literal("Parties to {{Deal}}:"),
            Literal("{{repeat:Parties}}"),
            Literal("- {{Name}}, a {{Type}} entity"),
            Literal("{{/repeat}}"),
            Literal("End of parties.")
        ]);

        var data = new Dictionary<string, object?>
        {
            ["Deal"] = "Project Falcon",
            ["Parties"] = new List<IReadOnlyDictionary<string, object?>>
            {
                new Dictionary<string, object?> { ["Name"] = "Acme Corp", ["Type"] = "Corporation" },
                new Dictionary<string, object?> { ["Name"] = "Beta LLC", ["Type"] = "LLC" }
            }
        };

        _sut.Fill(_templatePath, _outputPath, data);

        var texts = ReadParagraphs(_outputPath).Select(p => p.InnerText).ToList();
        Assert.Equal(
        [
            "Parties to Project Falcon:",
            "- Acme Corp, a Corporation entity",
            "- Beta LLC, a LLC entity",
            "End of parties."
        ], texts);
    }

    [Fact]
    public void Fill_injects_a_clause_by_id_and_continues_heading_numbering()
    {
        var libraryPath = TestFiles.NewTempPath(".docx");
        try
        {
            SampleDocumentFactory.CreateDocumentFromParagraphs(libraryPath,
            [
                new Paragraph(
                    new BookmarkStart { Id = "1", Name = "clause_governinglaw" },
                    new ParagraphProperties(
                        new ParagraphStyleId { Val = "Heading2" },
                        new NumberingProperties(new NumberingLevelReference { Val = 0 }, new NumberingId { Val = 99 })),
                    new Run(new Text("Governing Law. This agreement is governed by Delaware law.")),
                    new BookmarkEnd { Id = "1" })
            ]);
            SampleDocumentFactory.AddNumberingDefinition(libraryPath, numId: 99);

            SampleDocumentFactory.CreateDocumentFromParagraphs(_templatePath,
            [
                new Paragraph(
                    new ParagraphProperties(
                        new ParagraphStyleId { Val = "Heading2" },
                        new NumberingProperties(new NumberingLevelReference { Val = 0 }, new NumberingId { Val = 3 })),
                    new Run(new Text("Section 1. Recitals"))),
                Literal("Some recitals text."),
                Literal("{{clause:governinglaw}}"),
                Literal("End of agreement.")
            ]);
            SampleDocumentFactory.AddNumberingDefinition(_templatePath, numId: 3);

            _sut.Fill(_templatePath, _outputPath, new Dictionary<string, object?>(), MissingTokenPolicy.Error, new ClauseLibrary(libraryPath));

            var result = ReadParagraphs(_outputPath);
            var texts = result.Select(p => p.InnerText).ToList();
            Assert.Contains("Governing Law. This agreement is governed by Delaware law.", texts);
            Assert.DoesNotContain(texts, t => t.Contains("{{clause:"));

            // Clippit's DocumentBuilder rebuilds a fresh, unified numbering part for the merged
            // output, so the *specific* id (3, 99, ...) each side started with doesn't survive —
            // what "continuation" means, and what's worth asserting, is that the transplanted
            // clause ends up sharing the same numbering id as its neighboring heading, so Word
            // renders it as part of the same counted sequence rather than starting over at 1.
            var recitalsParagraph = result.First(p => p.InnerText.StartsWith("Section 1"));
            var recitalsNumId = recitalsParagraph.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value;
            var clauseParagraph = result.First(p => p.InnerText.StartsWith("Governing Law"));
            var numId = clauseParagraph.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value;
            Assert.NotNull(recitalsNumId);
            Assert.Equal(recitalsNumId, numId);
        }
        finally
        {
            File.Delete(libraryPath);
        }
    }

    [Fact]
    public void Fill_with_missing_clause_and_Redact_policy_removes_the_marker()
    {
        var libraryPath = TestFiles.NewTempPath(".docx");
        try
        {
            SampleDocumentFactory.CreateDocumentFromParagraphs(libraryPath, [Literal("unrelated")]);
            SampleDocumentFactory.CreateDocumentFromParagraphs(_templatePath,
                [Literal("Before."), Literal("{{clause:doesnotexist}}"), Literal("After.")]);

            _sut.Fill(_templatePath, _outputPath, new Dictionary<string, object?>(), MissingTokenPolicy.Redact, new ClauseLibrary(libraryPath));

            var texts = ReadParagraphs(_outputPath).Select(p => p.InnerText).ToList();
            Assert.Equal(["Before.", "After."], texts);
        }
        finally
        {
            File.Delete(libraryPath);
        }
    }

    [Fact]
    public void Fill_injects_html_as_a_whole_paragraph_replacement()
    {
        SampleDocumentFactory.CreateDocumentFromParagraphs(_templatePath,
            [Literal("Description:"), Literal("{{html:Description}}")]);

        var html = "<p>This is <b>bold</b> and <i>italic</i>.</p><ul><li>One</li><li>Two</li></ul>";
        _sut.Fill(_templatePath, _outputPath, new Dictionary<string, object?> { ["Description"] = html });

        var result = ReadParagraphs(_outputPath);
        // "Description:" + 1 paragraph for <p> + 2 <li> paragraphs = 4 total.
        Assert.Equal(4, result.Count);
        Assert.Contains(result, p => p.InnerText == "This is bold and italic.");
        Assert.Contains(result, p => p.InnerText == "One");
        Assert.Contains(result, p => p.InnerText == "Two");

        var boldRun = result.SelectMany(p => p.Elements<Run>()).First(r => r.InnerText == "bold");
        Assert.NotNull(boldRun.RunProperties?.GetFirstChild<Bold>());
    }

    [Fact]
    public void Fill_sanitizes_script_tags_and_event_handlers_out_of_injected_html()
    {
        SampleDocumentFactory.CreateDocumentFromParagraphs(_templatePath, [Literal("{{html:Body}}")]);

        var maliciousHtml = "<p onclick=\"alert(1)\">Hello <script>alert('xss')</script>World</p>" +
                             "<a href=\"javascript:alert(1)\">bad link</a>";
        _sut.Fill(_templatePath, _outputPath, new Dictionary<string, object?> { ["Body"] = maliciousHtml });

        var result = ReadParagraphs(_outputPath);
        var allText = string.Join(" ", result.Select(p => p.InnerText));
        Assert.DoesNotContain("alert", allText);
        Assert.DoesNotContain("script", allText, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Hello World", allText);
        Assert.Contains("bad link", allText);
    }

    public void Dispose()
    {
        File.Delete(_templatePath);
        File.Delete(_outputPath);
    }
}
