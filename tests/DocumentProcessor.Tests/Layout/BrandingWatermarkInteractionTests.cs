using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.Layout;
using DocumentProcessor.Core.Samples;
using DocumentProcessor.Core.Watermarking;

namespace DocumentProcessor.Tests.Layout;

/// <summary>
/// Branding and watermarking both write the section's default header, and both are normal steps in
/// a contract pipeline (brand the document, then stamp it DRAFT). Each service was built and tested
/// in isolation, so the interaction between them went uncovered — and watermarking used to replace
/// the header reference outright, silently dropping the tenant's logo. These tests hold the
/// composition behaviour in place.
/// </summary>
public class BrandingWatermarkInteractionTests : IDisposable
{
    private readonly string _path = TestFiles.NewTempPath(".docx");
    private static readonly byte[] PngBytes = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

    public BrandingWatermarkInteractionTests()
    {
        SampleDocumentFactory.CreateBasicDocument(_path, "Contract", ["Body text."]);
    }

    private int HeaderPartCount()
    {
        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        return doc.MainDocumentPart!.HeaderParts.Count();
    }

    /// <summary>Resolves the header the document will actually render — the part referenced by the
    /// section's default HeaderReference, not merely any header part left in the package.</summary>
    private HeaderPart? RenderedHeader(WordprocessingDocument doc)
    {
        var mainPart = doc.MainDocumentPart!;
        var sectPr = mainPart.Document!.Body!.Elements<SectionProperties>().Single();
        var reference = sectPr.Elements<HeaderReference>()
            .FirstOrDefault(r => r.Type is null || r.Type == HeaderFooterValues.Default);

        return reference?.Id?.Value is { } relId ? (HeaderPart)mainPart.GetPartById(relId) : null;
    }

    private (bool HasLogo, bool HasWatermark) RenderedHeaderContents()
    {
        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var header = RenderedHeader(doc);
        if (header is null)
            return (false, false);

        return (header.ImageParts.Any(), header.Header!.Descendants<Shape>().Any());
    }

    [Fact]
    public void Branding_alone_puts_the_logo_in_the_rendered_header()
    {
        new BrandingService().ApplyBranding(_path, new TenantBrandingSpec(LogoBytes: PngBytes));

        Assert.True(RenderedHeaderContents().HasLogo);
    }

    [Fact]
    public void Watermarking_after_branding_keeps_both_the_logo_and_the_watermark()
    {
        new BrandingService().ApplyBranding(_path, new TenantBrandingSpec(LogoBytes: PngBytes));
        new DocxWatermarkService().AddTextWatermark(_path, "DRAFT");

        var (hasLogo, hasWatermark) = RenderedHeaderContents();
        Assert.True(hasLogo, "the tenant's branding logo must survive watermarking");
        Assert.True(hasWatermark, "the watermark shape must be present");
    }

    [Fact]
    public void Watermarking_after_branding_reuses_the_existing_header_part()
    {
        new BrandingService().ApplyBranding(_path, new TenantBrandingSpec(LogoBytes: PngBytes));
        var afterBranding = HeaderPartCount();

        new DocxWatermarkService().AddTextWatermark(_path, "DRAFT");

        // Composing into the existing part rather than superseding it means no orphan is left
        // behind to be re-inflated and re-compressed by every later save.
        Assert.Equal(1, afterBranding);
        Assert.Equal(1, HeaderPartCount());
    }

    [Fact]
    public void Applying_a_watermark_twice_replaces_it_rather_than_stacking()
    {
        var service = new DocxWatermarkService();
        service.AddTextWatermark(_path, "DRAFT");
        service.AddTextWatermark(_path, "CONFIDENTIAL");

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var header = RenderedHeader(doc)!;
        Assert.Single(header.Header!.Descendants<Shape>());
        Assert.Contains("CONFIDENTIAL", header.Header.InnerText);
        Assert.DoesNotContain("DRAFT", header.Header.InnerText);
    }

    [Fact]
    public void Removing_the_watermark_leaves_the_branding_logo_intact()
    {
        new BrandingService().ApplyBranding(_path, new TenantBrandingSpec(LogoBytes: PngBytes));
        var watermark = new DocxWatermarkService();
        watermark.AddTextWatermark(_path, "DRAFT");

        Assert.True(watermark.RemoveWatermark(_path));

        var (hasLogo, hasWatermark) = RenderedHeaderContents();
        Assert.True(hasLogo, "removing the watermark must not take the logo with it");
        Assert.False(hasWatermark);
    }

    public void Dispose() => File.Delete(_path);
}
