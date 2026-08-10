using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentProcessor.Core.Layout;

/// <summary>
/// General-purpose header/footer text content — distinct from
/// <see cref="Watermarking.DocxWatermarkService"/>'s header injection, which is single-purpose (it
/// only ever hosts the watermark shape). Applies to every section in the document.
/// </summary>
public sealed class HeaderFooterService
{
    /// <param name="type">
    /// <see cref="HeaderFooterValues.Default"/> (every page), <see cref="HeaderFooterValues.First"/>
    /// (first page only — also turns on <c>w:titlePg</c> on every section, since Word ignores a
    /// first-page header/footer without it), or <see cref="HeaderFooterValues.Even"/> (even pages —
    /// also turns on <c>w:evenAndOddHeaders</c> in document settings, for the same reason).
    /// </param>
    public void SetHeaderText(string docxPath, string text, HeaderFooterValues? type = null)
    {
        var resolvedType = type ?? HeaderFooterValues.Default;
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Document has no main part.");
        var document = mainPart.Document ?? throw new InvalidOperationException("Document has no body.");
        var body = document.Body ?? throw new InvalidOperationException("Document has no body.");

        var headerPart = mainPart.AddNewPart<HeaderPart>();
        headerPart.Header = new Header(new Paragraph(new Run(new Text(text))));
        headerPart.Header.Save();
        var relId = mainPart.GetIdOfPart(headerPart);

        foreach (var sectPr in ResolveSections(body))
        {
            foreach (var existing in sectPr.Elements<HeaderReference>().Where(r => r.Type?.Value == resolvedType).ToList())
                existing.Remove();
            sectPr.PrependChild(new HeaderReference { Type = resolvedType, Id = relId });
            ApplyTypeSideEffects(mainPart, sectPr, resolvedType);
        }

        document.Save();
    }

    /// <param name="type">Same semantics as <see cref="SetHeaderText"/>'s <c>type</c> parameter.</param>
    public void SetFooterText(string docxPath, string text, HeaderFooterValues? type = null)
    {
        var resolvedType = type ?? HeaderFooterValues.Default;
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Document has no main part.");
        var document = mainPart.Document ?? throw new InvalidOperationException("Document has no body.");
        var body = document.Body ?? throw new InvalidOperationException("Document has no body.");

        var footerPart = mainPart.AddNewPart<FooterPart>();
        footerPart.Footer = new Footer(new Paragraph(new Run(new Text(text))));
        footerPart.Footer.Save();
        var relId = mainPart.GetIdOfPart(footerPart);

        foreach (var sectPr in ResolveSections(body))
        {
            foreach (var existing in sectPr.Elements<FooterReference>().Where(r => r.Type?.Value == resolvedType).ToList())
                existing.Remove();
            sectPr.PrependChild(new FooterReference { Type = resolvedType, Id = relId });
            ApplyTypeSideEffects(mainPart, sectPr, resolvedType);
        }

        document.Save();
    }

    public bool RemoveHeader(string docxPath, HeaderFooterValues? type = null)
    {
        var resolvedType = type ?? HeaderFooterValues.Default;
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Document has no main part.");
        var document = mainPart.Document ?? throw new InvalidOperationException("Document has no body.");
        var body = document.Body ?? throw new InvalidOperationException("Document has no body.");

        var removedAny = false;
        foreach (var sectPr in body.Descendants<SectionProperties>())
        {
            foreach (var reference in sectPr.Elements<HeaderReference>().Where(r => r.Type?.Value == resolvedType).ToList())
            {
                reference.Remove();
                removedAny = true;
            }
        }

        if (removedAny)
            document.Save();

        return removedAny;
    }

    public bool RemoveFooter(string docxPath, HeaderFooterValues? type = null)
    {
        var resolvedType = type ?? HeaderFooterValues.Default;
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Document has no main part.");
        var document = mainPart.Document ?? throw new InvalidOperationException("Document has no body.");
        var body = document.Body ?? throw new InvalidOperationException("Document has no body.");

        var removedAny = false;
        foreach (var sectPr in body.Descendants<SectionProperties>())
        {
            foreach (var reference in sectPr.Elements<FooterReference>().Where(r => r.Type?.Value == resolvedType).ToList())
            {
                reference.Remove();
                removedAny = true;
            }
        }

        if (removedAny)
            document.Save();

        return removedAny;
    }

    private static IReadOnlyList<SectionProperties> ResolveSections(Body body)
    {
        var sections = body.Descendants<SectionProperties>().ToList();
        if (sections.Count > 0)
            return sections;

        var sectPr = new SectionProperties();
        body.AppendChild(sectPr);
        return [sectPr];
    }

    private static void ApplyTypeSideEffects(MainDocumentPart mainPart, SectionProperties sectPr, HeaderFooterValues type)
    {
        if (type == HeaderFooterValues.First && !sectPr.Elements<TitlePage>().Any())
        {
            sectPr.AppendChild(new TitlePage());
        }
        else if (type == HeaderFooterValues.Even)
        {
            var settingsPart = mainPart.DocumentSettingsPart ?? mainPart.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings ??= new Settings();
            if (!settingsPart.Settings.Elements<EvenAndOddHeaders>().Any())
                settingsPart.Settings.PrependChild(new EvenAndOddHeaders());
            settingsPart.Settings.Save();
        }
    }
}
