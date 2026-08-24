using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentProcessor.Core.Layout;

/// <summary>
/// A tenant's visual identity: a header logo plus the accent color applied to heading styles —
/// beyond <see cref="FontEmbedding.FontEmbeddingService"/>'s font-only customization, this is what most
/// callers actually mean by "branding" a generated document.
/// </summary>
public sealed record TenantBrandingSpec(
    byte[]? LogoBytes = null,
    double LogoWidthPt = 100,
    double LogoHeightPt = 40,
    string? AccentColorHex = null,
    IReadOnlyList<string>? HeadingStyleIds = null)
{
    public IReadOnlyList<string> ResolvedHeadingStyleIds => HeadingStyleIds ?? ["Heading1", "Heading2", "Heading3"];
}

/// <summary>Applies a <see cref="TenantBrandingSpec"/> to a document: a header logo (via
/// <see cref="HeaderFooterService"/>) and/or a run-color override on the target heading styles in
/// the document's style definitions part — a plain OOXML <c>w:color</c> change, not a full theme
/// swap, so it applies immediately without depending on which Office theme the recipient has.</summary>
public sealed class BrandingService : IBrandingService
{
    private readonly HeaderFooterService _headerFooterService = new();

    public void ApplyBranding(string docxPath, TenantBrandingSpec branding)
    {
        if (branding.LogoBytes is not null)
        {
            _headerFooterService.SetHeaderContent(docxPath, [new LogoPart(branding.LogoBytes, branding.LogoWidthPt, branding.LogoHeightPt)]);
        }

        if (branding.AccentColorHex is not null)
            ApplyHeadingColor(docxPath, branding.AccentColorHex, branding.ResolvedHeadingStyleIds);
    }

    private static void ApplyHeadingColor(string docxPath, string colorHex, IReadOnlyList<string> headingStyleIds)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var mainPart = doc.MainDocumentPart ?? throw new CorruptDocumentException("Document has no main part.");
        var stylesPart = mainPart.StyleDefinitionsPart ?? mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles ??= new Styles();

        foreach (var styleId in headingStyleIds)
        {
            var style = stylesPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == styleId);
            if (style is null)
            {
                style = new Style(new StyleName { Val = styleId }, new BasedOn { Val = "Normal" })
                {
                    Type = StyleValues.Paragraph,
                    StyleId = styleId
                };
                stylesPart.Styles.AppendChild(style);
            }

            // A style's run properties are StyleRunProperties, not the plain RunProperties a Run or
            // an inline w:rPr override uses — same "w:rPr" tag name, different backing CLR type, so
            // using the wrong one here would silently create a second, schema-invalid <w:rPr>
            // instead of finding and updating the existing one.
            var styleRunProperties = style.StyleRunProperties;
            if (styleRunProperties is null)
            {
                styleRunProperties = new StyleRunProperties();
                style.AppendChild(styleRunProperties);
            }

            styleRunProperties.RemoveAllChildren<Color>();
            styleRunProperties.AppendChild(new Color { Val = colorHex.TrimStart('#') });
        }

        stylesPart.Styles.Save();
    }
}
