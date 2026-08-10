using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentProcessor.Core.Layout;

public enum PageOrientation { Portrait, Landscape }

public sealed record PageMargins(
    int TopTwips, int BottomTwips, int LeftTwips, int RightTwips,
    int HeaderTwips = 720, int FooterTwips = 720, int GutterTwips = 0)
{
    /// <summary>Builds margins from inch values (1 inch = 1440 twips).</summary>
    public static PageMargins FromInches(
        double top, double bottom, double left, double right,
        double header = 0.5, double footer = 0.5, double gutter = 0) => new(
        (int)(top * 1440), (int)(bottom * 1440), (int)(left * 1440), (int)(right * 1440),
        (int)(header * 1440), (int)(footer * 1440), (int)(gutter * 1440));
}

/// <summary>
/// A page's physical dimensions in twips (1/1440 inch), portrait-oriented by convention — the
/// <see cref="Orientation"/> factory parameter swaps <see cref="WidthTwips"/>/<see cref="HeightTwips"/>
/// for you rather than requiring the caller to pre-swap them, since OOXML represents landscape as
/// "swapped dimensions plus an orientation flag," not as a flag alone.
/// </summary>
public sealed record PageSize(int WidthTwips, int HeightTwips, PageOrientation Orientation = PageOrientation.Portrait)
{
    public static PageSize Letter(PageOrientation orientation = PageOrientation.Portrait) =>
        Build(12240, 15840, orientation);

    public static PageSize A4(PageOrientation orientation = PageOrientation.Portrait) =>
        Build(11906, 16838, orientation);

    public static PageSize Legal(PageOrientation orientation = PageOrientation.Portrait) =>
        Build(12240, 20160, orientation);

    private static PageSize Build(int portraitWidth, int portraitHeight, PageOrientation orientation) =>
        orientation == PageOrientation.Landscape
            ? new PageSize(portraitHeight, portraitWidth, orientation)
            : new PageSize(portraitWidth, portraitHeight, orientation);
}

/// <summary>
/// Sets physical page geometry (size, orientation, margins, columns) and the two content-flow
/// primitives that don't belong to any other service (page breaks, default paragraph/line spacing).
/// </summary>
/// <remarks>
/// Every section-properties method accepts an optional <c>sectionIndex</c> — every document produced
/// by this library today has exactly one section, so it defaults to "apply to all sections," but the
/// parameter exists so a future multi-section document (independently laid-out parts within one
/// docx) doesn't require reworking this API, only calling it once per section.
/// </remarks>
public sealed class PageLayoutService
{
    public void SetPageSize(string docxPath, PageSize size, int? sectionIndex = null)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var document = GetDocument(doc);

        foreach (var sectPr in ResolveSections(document, sectionIndex))
        {
            sectPr.Elements<DocumentFormat.OpenXml.Wordprocessing.PageSize>().FirstOrDefault()?.Remove();
            sectPr.PrependChild(new DocumentFormat.OpenXml.Wordprocessing.PageSize
            {
                Width = (uint)size.WidthTwips,
                Height = (uint)size.HeightTwips,
                Orient = size.Orientation == PageOrientation.Landscape ? PageOrientationValues.Landscape : PageOrientationValues.Portrait
            });
        }

        document.Save();
    }

    public void SetMargins(string docxPath, PageMargins margins, int? sectionIndex = null)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var document = GetDocument(doc);

        foreach (var sectPr in ResolveSections(document, sectionIndex))
        {
            sectPr.Elements<PageMargin>().FirstOrDefault()?.Remove();
            sectPr.AppendChild(new PageMargin
            {
                Top = margins.TopTwips,
                Bottom = margins.BottomTwips,
                Left = (uint)margins.LeftTwips,
                Right = (uint)margins.RightTwips,
                Header = (uint)margins.HeaderTwips,
                Footer = (uint)margins.FooterTwips,
                Gutter = (uint)margins.GutterTwips
            });
        }

        document.Save();
    }

    public void SetColumns(string docxPath, int columnCount, int spacingTwips = 720, int? sectionIndex = null)
    {
        if (columnCount < 1)
            throw new ArgumentOutOfRangeException(nameof(columnCount), "Must have at least 1 column.");

        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var document = GetDocument(doc);

        foreach (var sectPr in ResolveSections(document, sectionIndex))
        {
            sectPr.Elements<Columns>().FirstOrDefault()?.Remove();
            sectPr.AppendChild(new Columns
            {
                ColumnCount = (Int16Value)(short)columnCount,
                Space = spacingTwips.ToString(),
                EqualWidth = true
            });
        }

        document.Save();
    }

    /// <summary>
    /// Inserts a hard page break in a new, empty paragraph before the top-level paragraph at
    /// <paramref name="beforeParagraphIndex"/> (0-based, same indexing convention as
    /// <see cref="Transplant.ClauseTransplantService.ListParagraphs"/>).
    /// </summary>
    public void InsertPageBreak(string docxPath, int beforeParagraphIndex)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var document = GetDocument(doc);
        var body = document.Body ?? throw new InvalidOperationException("Document has no body.");

        var paragraphs = body.Elements<Paragraph>().ToList();
        if (beforeParagraphIndex < 0 || beforeParagraphIndex > paragraphs.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(beforeParagraphIndex),
                $"Document has {paragraphs.Count} paragraphs; valid insertion points are 0..{paragraphs.Count}.");
        }

        var breakParagraph = new Paragraph(new Run(new Break { Type = BreakValues.Page }));

        if (beforeParagraphIndex < paragraphs.Count)
            paragraphs[beforeParagraphIndex].InsertBeforeSelf(breakParagraph);
        else
            body.InsertBefore(breakParagraph, body.Elements<SectionProperties>().FirstOrDefault());

        document.Save();
    }

    /// <summary>
    /// Sets the document's default paragraph spacing (space-after, line spacing) on both
    /// <c>w:docDefaults</c> and the <c>Normal</c> style, so it applies to every paragraph that
    /// doesn't explicitly override it — the same mechanism (and the same Word Normal.dotm baseline
    /// values, by default) that <c>Samples.SampleDocumentFactory</c> uses to keep generated fixtures
    /// paginating the way a real Word document would.
    /// </summary>
    public void SetDefaultParagraphSpacing(string docxPath, int afterTwips, int lineTwips, LineSpacingRuleValues lineRule)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Document has no main part.");

        var stylesPart = mainPart.StyleDefinitionsPart ?? mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles ??= new Styles();

        var spacing = new SpacingBetweenLines { After = afterTwips.ToString(), Line = lineTwips.ToString(), LineRule = lineRule };

        var docDefaults = stylesPart.Styles.Elements<DocDefaults>().FirstOrDefault();
        if (docDefaults is null)
        {
            docDefaults = new DocDefaults();
            stylesPart.Styles.PrependChild(docDefaults);
        }

        var pPrDefault = docDefaults.ParagraphPropertiesDefault ??= new ParagraphPropertiesDefault();
        pPrDefault.ParagraphPropertiesBaseStyle ??= new ParagraphPropertiesBaseStyle();
        pPrDefault.ParagraphPropertiesBaseStyle.Elements<SpacingBetweenLines>().FirstOrDefault()?.Remove();
        pPrDefault.ParagraphPropertiesBaseStyle.AppendChild((SpacingBetweenLines)spacing.CloneNode(true));

        var normalStyle = stylesPart.Styles.Elements<Style>()
            .FirstOrDefault(s => string.Equals(s.StyleId?.Value, "Normal", StringComparison.Ordinal));
        if (normalStyle is null)
        {
            normalStyle = new Style(new StyleName { Val = "Normal" }) { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true };
            stylesPart.Styles.AppendChild(normalStyle);
        }

        normalStyle.StyleParagraphProperties ??= new StyleParagraphProperties();
        normalStyle.StyleParagraphProperties.Elements<SpacingBetweenLines>().FirstOrDefault()?.Remove();
        normalStyle.StyleParagraphProperties.AppendChild((SpacingBetweenLines)spacing.CloneNode(true));

        stylesPart.Styles.Save();
    }

    private static Document GetDocument(WordprocessingDocument doc) =>
        doc.MainDocumentPart?.Document ?? throw new InvalidOperationException("Document has no main part/body.");

    private static IReadOnlyList<SectionProperties> ResolveSections(Document document, int? sectionIndex)
    {
        var body = document.Body ?? throw new InvalidOperationException("Document has no body.");
        var allSections = body.Descendants<SectionProperties>().ToList();

        if (allSections.Count == 0)
        {
            var sectPr = new SectionProperties();
            body.AppendChild(sectPr);
            allSections.Add(sectPr);
        }

        if (sectionIndex is null)
            return allSections;

        if (sectionIndex < 0 || sectionIndex >= allSections.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(sectionIndex),
                $"Document has {allSections.Count} section(s); valid indexes are 0..{allSections.Count - 1}.");
        }

        return [allSections[sectionIndex.Value]];
    }
}
