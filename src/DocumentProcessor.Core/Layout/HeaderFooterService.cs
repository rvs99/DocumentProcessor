using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.Templating;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

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
        var mainPart = doc.MainDocumentPart ?? throw new CorruptDocumentException("Document has no main part.");
        var document = mainPart.Document ?? throw new CorruptDocumentException("Document has no body.");
        var body = document.Body ?? throw new CorruptDocumentException("Document has no body.");

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
        var mainPart = doc.MainDocumentPart ?? throw new CorruptDocumentException("Document has no main part.");
        var document = mainPart.Document ?? throw new CorruptDocumentException("Document has no body.");
        var body = document.Body ?? throw new CorruptDocumentException("Document has no body.");

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

    /// <summary>
    /// Sets a header built from mixed content — text (optionally <c>{{token}}</c>-templated when
    /// <paramref name="data"/> is supplied), live PAGE/NUMPAGES/DATE fields, and an inline logo —
    /// laid out left-to-right in one paragraph, going beyond <see cref="SetHeaderText"/>'s
    /// plain-text-only content.
    /// </summary>
    public void SetHeaderContent(string docxPath, IReadOnlyList<HeaderFooterPart> parts, IReadOnlyDictionary<string, object?>? data = null, HeaderFooterValues? type = null) =>
        SetContent(docxPath, parts, data, type, isHeader: true);

    /// <summary>Same content model as <see cref="SetHeaderContent"/>, for the footer.</summary>
    public void SetFooterContent(string docxPath, IReadOnlyList<HeaderFooterPart> parts, IReadOnlyDictionary<string, object?>? data = null, HeaderFooterValues? type = null) =>
        SetContent(docxPath, parts, data, type, isHeader: false);

    private void SetContent(string docxPath, IReadOnlyList<HeaderFooterPart> parts, IReadOnlyDictionary<string, object?>? data, HeaderFooterValues? type, bool isHeader)
    {
        var resolvedType = type ?? HeaderFooterValues.Default;
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var mainPart = doc.MainDocumentPart ?? throw new CorruptDocumentException("Document has no main part.");
        var document = mainPart.Document ?? throw new CorruptDocumentException("Document has no body.");
        var body = document.Body ?? throw new CorruptDocumentException("Document has no body.");

        var paragraph = new Paragraph();
        string relId;

        if (isHeader)
        {
            var headerPart = mainPart.AddNewPart<HeaderPart>();
            foreach (var part in parts)
                paragraph.AppendChild(BuildPartElement(headerPart, part, data));
            headerPart.Header = new Header(paragraph);
            headerPart.Header.Save();
            relId = mainPart.GetIdOfPart(headerPart);
        }
        else
        {
            var footerPart = mainPart.AddNewPart<FooterPart>();
            foreach (var part in parts)
                paragraph.AppendChild(BuildPartElement(footerPart, part, data));
            footerPart.Footer = new Footer(paragraph);
            footerPart.Footer.Save();
            relId = mainPart.GetIdOfPart(footerPart);
        }

        foreach (var sectPr in ResolveSections(body))
        {
            if (isHeader)
            {
                foreach (var existing in sectPr.Elements<HeaderReference>().Where(r => r.Type?.Value == resolvedType).ToList())
                    existing.Remove();
                sectPr.PrependChild(new HeaderReference { Type = resolvedType, Id = relId });
            }
            else
            {
                foreach (var existing in sectPr.Elements<FooterReference>().Where(r => r.Type?.Value == resolvedType).ToList())
                    existing.Remove();
                sectPr.PrependChild(new FooterReference { Type = resolvedType, Id = relId });
            }

            ApplyTypeSideEffects(mainPart, sectPr, resolvedType);
        }

        document.Save();
    }

    private static OpenXmlElement BuildPartElement<TPart>(TPart owningPart, HeaderFooterPart part, IReadOnlyDictionary<string, object?>? data)
        where TPart : OpenXmlPart, ISupportedRelationship<ImagePart> => part switch
    {
        TextPart textPart => new Run(new Text(ResolveTokens(textPart.Text, data)) { Space = SpaceProcessingModeValues.Preserve }),
        FieldPart fieldPart => BuildFieldElement(fieldPart.FieldType),
        LogoPart logoPart => BuildImageRun(owningPart, logoPart),
        _ => throw new NotSupportedException($"Unknown header/footer part type '{part.GetType().Name}'.")
    };

    private static string ResolveTokens(string text, IReadOnlyDictionary<string, object?>? data)
    {
        if (data is null)
            return text;

        var context = new TemplateContext(data);
        return Regex.Replace(text, @"\{\{([A-Za-z0-9_.]+)\}\}", m =>
            context.TryResolve(m.Groups[1].Value, out var value) && value is not null
                ? TemplateValueFormatter.ToComparableString(value)
                : m.Value);
    }

    private static SimpleField BuildFieldElement(HeaderFooterFieldType type)
    {
        var (instruction, cachedText) = type switch
        {
            HeaderFooterFieldType.PageNumber => ("PAGE", "1"),
            HeaderFooterFieldType.TotalPages => ("NUMPAGES", "1"),
            HeaderFooterFieldType.Date => ("DATE", DateTime.Now.ToString("d")),
            _ => throw new ArgumentOutOfRangeException(nameof(type))
        };

        return new SimpleField(new Run(new Text(cachedText))) { Instruction = instruction };
    }

    private static Run BuildImageRun<TPart>(TPart owningPart, LogoPart logo)
        where TPart : OpenXmlPart, ISupportedRelationship<ImagePart>
    {
        var partType = IsPng(logo.ImageBytes) ? ImagePartType.Png : ImagePartType.Jpeg;
        var imagePart = owningPart.AddImagePart(partType);
        using (var stream = new MemoryStream(logo.ImageBytes))
            imagePart.FeedData(stream);
        var relId = owningPart.GetIdOfPart(imagePart);

        var widthEmu = (long)(logo.WidthPt * 12700);
        var heightEmu = (long)(logo.HeightPt * 12700);

        var shapeProperties = new PIC.ShapeProperties(
            new A.Transform2D(
                new A.Offset { X = 0L, Y = 0L },
                new A.Extents { Cx = widthEmu, Cy = heightEmu }),
            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle });

        var blipFill = new PIC.BlipFill(
            new A.Blip { Embed = relId },
            new A.Stretch(new A.FillRectangle()));

        var nonVisualPictureProperties = new PIC.NonVisualPictureProperties(
            new PIC.NonVisualDrawingProperties { Id = 0U, Name = "Logo" },
            new PIC.NonVisualPictureDrawingProperties());

        var picture = new PIC.Picture(nonVisualPictureProperties, blipFill, shapeProperties);

        var graphicData = new A.GraphicData(picture) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };
        var graphic = new A.Graphic(graphicData);

        var inline = new DW.Inline(
            new DW.Extent { Cx = widthEmu, Cy = heightEmu },
            new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
            new DW.DocProperties { Id = 1U, Name = "Logo" },
            new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
            graphic)
        {
            DistanceFromTop = 0U,
            DistanceFromBottom = 0U,
            DistanceFromLeft = 0U,
            DistanceFromRight = 0U
        };

        return new Run(new Drawing(inline));
    }

    private static bool IsPng(byte[] bytes) =>
        bytes.Length >= 8 && bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47;

    public bool RemoveHeader(string docxPath, HeaderFooterValues? type = null)
    {
        var resolvedType = type ?? HeaderFooterValues.Default;
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var mainPart = doc.MainDocumentPart ?? throw new CorruptDocumentException("Document has no main part.");
        var document = mainPart.Document ?? throw new CorruptDocumentException("Document has no body.");
        var body = document.Body ?? throw new CorruptDocumentException("Document has no body.");

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
        var mainPart = doc.MainDocumentPart ?? throw new CorruptDocumentException("Document has no main part.");
        var document = mainPart.Document ?? throw new CorruptDocumentException("Document has no body.");
        var body = document.Body ?? throw new CorruptDocumentException("Document has no body.");

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
