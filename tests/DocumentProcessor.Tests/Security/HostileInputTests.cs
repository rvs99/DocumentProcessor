using DocumentProcessor.Core.Templating;

namespace DocumentProcessor.Tests.Security;

/// <summary>
/// Guards against tenant-supplied content that is well-formed but unbounded. These are not
/// correctness tests for the conversion itself — they exist because the failure mode without them
/// is an uncatchable <c>StackOverflowException</c> that terminates the host process, taking every
/// other tenant's in-flight request with it.
/// </summary>
public class HostileInputTests
{
    private static string NestedDivs(int depth) =>
        string.Concat(Enumerable.Repeat("<div>", depth)) + "payload" + string.Concat(Enumerable.Repeat("</div>", depth));

    [Fact]
    public void Sanitize_rejects_html_nested_past_the_depth_limit_instead_of_overflowing_the_stack()
    {
        var hostile = NestedDivs(HtmlToOoxmlConverter.MaxNestingDepth + 50);

        var ex = Assert.Throws<HtmlTooComplexException>(() => HtmlToOoxmlConverter.Sanitize(hostile));
        Assert.Contains("nesting", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Sanitize_still_accepts_realistically_nested_html()
    {
        // Well within the cap — a real rich-text field from an editor nests a handful of levels.
        var reasonable = "<div><p>Terms include <b>priority <i>24/7</i></b> support.</p></div>";

        var result = HtmlToOoxmlConverter.Sanitize(reasonable);

        Assert.Contains("priority", result);
        Assert.Contains("24/7", result);
    }

    [Fact]
    public void Sanitize_rejects_an_oversized_fragment()
    {
        var oversized = new string('a', HtmlToOoxmlConverter.MaxInputLength + 1);

        var ex = Assert.Throws<HtmlTooComplexException>(() => HtmlToOoxmlConverter.Sanitize(oversized));
        Assert.Contains("characters", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Template_fill_surfaces_the_limit_rather_than_killing_the_process()
    {
        var templatePath = TestFiles.NewTempPath(".docx");
        var outputPath = TestFiles.NewTempPath(".docx");
        try
        {
            Core.Samples.SampleDocumentFactory.CreateDocumentFromParagraphs(templatePath,
            [
                new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                    new DocumentFormat.OpenXml.Wordprocessing.Run(
                        new DocumentFormat.OpenXml.Wordprocessing.Text("{{html:Body}}")))
            ]);

            var data = new Dictionary<string, object?> { ["Body"] = NestedDivs(HtmlToOoxmlConverter.MaxNestingDepth + 50) };

            // The point is that this is a catchable exception the caller can map to a 400, not a
            // process kill: reaching this assertion at all proves the host survived.
            Assert.Throws<HtmlTooComplexException>(() => new TemplateEngine().Fill(templatePath, outputPath, data));
        }
        finally
        {
            File.Delete(templatePath);
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }
}
