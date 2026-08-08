using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.FontEmbedding;
using DocumentProcessor.Core.Samples;

namespace DocumentProcessor.Tests.Fonts;

public class FontEmbeddingServiceTests : IDisposable
{
    // SIL Open Font License 1.1 (see Assets/OFL.txt) — safe to redistribute as a test fixture.
    private static readonly string TestFontPath = Path.Combine(AppContext.BaseDirectory, "Assets", "RobotoMono-Regular.ttf");

    private readonly string _path = TestFiles.NewTempPath(".docx");
    private readonly FontEmbeddingService _sut = new();

    public FontEmbeddingServiceTests()
    {
        SampleDocumentFactory.CreateBasicDocument(_path, "Font Test", ["Body text that should switch fonts."]);
    }

    [Fact]
    public void EmbedFontFamily_writes_a_font_table_part_containing_the_family()
    {
        _sut.EmbedFontFamily(_path, "Roboto Mono", new FontFamilyFiles(TestFontPath));

        var embedded = _sut.ListEmbeddedFonts(_path);
        Assert.Contains("Roboto Mono", embedded);

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var fontTablePart = doc.MainDocumentPart!.FontTablePart;
        Assert.NotNull(fontTablePart);
        Assert.Single(fontTablePart!.FontParts);

        var settings = doc.MainDocumentPart!.DocumentSettingsPart!.Settings!;
        Assert.NotNull(settings.Elements<EmbedTrueTypeFonts>().FirstOrDefault());
    }

    [Fact]
    public void EmbedFontFamily_re_embedding_the_same_family_replaces_rather_than_duplicates()
    {
        _sut.EmbedFontFamily(_path, "Roboto Mono", new FontFamilyFiles(TestFontPath));
        _sut.EmbedFontFamily(_path, "Roboto Mono", new FontFamilyFiles(TestFontPath));

        var embedded = _sut.ListEmbeddedFonts(_path);
        Assert.Single(embedded, f => f == "Roboto Mono");
    }

    [Fact]
    public void ApplyFontToAllRuns_sets_run_fonts_on_every_run_in_the_document()
    {
        _sut.EmbedFontFamily(_path, "Roboto Mono", new FontFamilyFiles(TestFontPath));
        _sut.ApplyFontToAllRuns(_path, "Roboto Mono");

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var runs = doc.MainDocumentPart!.Document!.Descendants<Run>().ToList();

        Assert.NotEmpty(runs);
        Assert.All(runs, r => Assert.Equal("Roboto Mono", r.RunProperties?.RunFonts?.Ascii?.Value));
    }

    [Fact]
    public void EmbedFontFamily_rejects_unsupported_font_file_extensions()
    {
        var otfPath = TestFiles.NewTempPath(".otf");
        File.WriteAllBytes(otfPath, [0, 1, 2, 3]);
        try
        {
            Assert.Throws<NotSupportedException>(() =>
                _sut.EmbedFontFamily(_path, "Fake", new FontFamilyFiles(otfPath)));
        }
        finally
        {
            File.Delete(otfPath);
        }
    }

    public void Dispose() => File.Delete(_path);
}
