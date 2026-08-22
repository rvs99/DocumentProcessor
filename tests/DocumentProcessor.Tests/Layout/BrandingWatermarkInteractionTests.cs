using DocumentFormat.OpenXml.Packaging;
using DocumentProcessor.Core.Layout;
using DocumentProcessor.Core.Samples;
using DocumentProcessor.Core.Watermarking;

namespace DocumentProcessor.Tests.Layout;

/// <summary>
/// Branding and watermarking both write into the document's *default* header, and both are normal
/// steps in a contract pipeline (brand the document, then stamp it DRAFT). These tests pin down
/// what happens when they're combined — the services were built and tested in isolation, so the
/// interaction between them was never covered.
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

    /// <summary>The header the document will actually render: the part referenced by the section's
    /// default HeaderReference, not merely any header part still sitting in the package.</summary>
    private bool ReferencedHeaderContainsLogo()
    {
        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var mainPart = doc.MainDocumentPart!;
        var sectPr = mainPart.Document!.Body!
            .Elements<DocumentFormat.OpenXml.Wordprocessing.SectionProperties>().Single();
        var reference = sectPr.Elements<DocumentFormat.OpenXml.Wordprocessing.HeaderReference>()
            .FirstOrDefault(r => r.Type is null || r.Type == DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default);

        if (reference?.Id?.Value is not { } relId)
            return false;

        return ((HeaderPart)mainPart.GetPartById(relId)).ImageParts.Any();
    }

    [Fact]
    public void Branding_alone_puts_the_logo_in_the_referenced_header()
    {
        new BrandingService().ApplyBranding(_path, new TenantBrandingSpec(LogoBytes: PngBytes));

        Assert.True(ReferencedHeaderContainsLogo());
    }

    [Fact]
    public void Watermarking_after_branding_discards_the_tenant_logo()
    {
        new BrandingService().ApplyBranding(_path, new TenantBrandingSpec(LogoBytes: PngBytes));
        new DocxWatermarkService().AddTextWatermark(_path, "DRAFT");

        // The rendered header is now the watermark-only header: AddTextWatermark removes the
        // existing default HeaderReference and prepends its own, so the branding logo is no
        // longer reachable from the document even though its part is still in the package.
        Assert.False(ReferencedHeaderContainsLogo());
    }

    [Fact]
    public void Watermarking_after_branding_leaves_an_orphaned_header_part_in_the_package()
    {
        new BrandingService().ApplyBranding(_path, new TenantBrandingSpec(LogoBytes: PngBytes));
        var afterBranding = HeaderPartCount();

        new DocxWatermarkService().AddTextWatermark(_path, "DRAFT");
        var afterWatermark = HeaderPartCount();

        // The superseded branding header part is never deleted — it stays zipped into the package
        // (along with its logo image) and is re-inflated and re-compressed by every later save.
        Assert.Equal(1, afterBranding);
        Assert.Equal(2, afterWatermark);
    }

    public void Dispose() => File.Delete(_path);
}
