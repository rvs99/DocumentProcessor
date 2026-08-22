using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.Layout;
using DocumentProcessor.Core.Samples;

namespace DocumentProcessor.Tests.Layout;

public class BrandingServiceTests : IDisposable
{
    private readonly string _path = TestFiles.NewTempPath(".docx");
    private readonly BrandingService _sut = new();
    private static readonly byte[] PngBytes = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

    public BrandingServiceTests()
    {
        SampleDocumentFactory.CreateBasicDocument(_path, "Branding Test", ["Body text."]);
    }

    [Fact]
    public void ApplyBranding_with_a_logo_adds_a_header_image()
    {
        _sut.ApplyBranding(_path, new TenantBrandingSpec(LogoBytes: PngBytes));

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var headerParts = doc.MainDocumentPart!.HeaderParts.ToList();
        Assert.Single(headerParts);
        Assert.Single(headerParts[0].ImageParts);
    }

    [Fact]
    public void ApplyBranding_with_an_accent_color_sets_it_on_the_default_heading_styles()
    {
        _sut.ApplyBranding(_path, new TenantBrandingSpec(AccentColorHex: "#2A6DB0"));

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var styles = doc.MainDocumentPart!.StyleDefinitionsPart!.Styles!;

        foreach (var styleId in new[] { "Heading1", "Heading2", "Heading3" })
        {
            var style = styles.Elements<Style>().Single(s => s.StyleId?.Value == styleId);
            var color = style.StyleRunProperties!.GetFirstChild<Color>()!;
            Assert.Equal("2A6DB0", color.Val!.Value);
        }
    }

    [Fact]
    public void ApplyBranding_with_custom_heading_style_ids_only_touches_those()
    {
        _sut.ApplyBranding(_path, new TenantBrandingSpec(AccentColorHex: "FF0000", HeadingStyleIds: ["Title"]));

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var styles = doc.MainDocumentPart!.StyleDefinitionsPart!.Styles!;

        var titleStyle = styles.Elements<Style>().Single(s => s.StyleId?.Value == "Title");
        Assert.Equal("FF0000", titleStyle.StyleRunProperties!.GetFirstChild<Color>()!.Val!.Value);

        // Heading1 already exists in the base fixture (unrelated to branding) — it just shouldn't
        // have been colored, since it wasn't in the requested HeadingStyleIds list.
        var heading1Style = styles.Elements<Style>().Single(s => s.StyleId?.Value == "Heading1");
        Assert.Null(heading1Style.StyleRunProperties?.GetFirstChild<Color>());
    }

    public void Dispose() => File.Delete(_path);
}
