using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;
using Picture = DocumentFormat.OpenXml.Wordprocessing.Picture;

namespace DocumentProcessor.Core.Watermarking;

/// <summary>
/// Adds a diagonal text watermark (e.g. "DRAFT", "CONFIDENTIAL") to a .docx file, using a VML
/// shape in the document header so it repeats on every page — the same mechanism Word itself uses
/// for Insert → Watermark.
/// </summary>
/// <remarks>
/// Uses a VML text box (<c>v:textbox</c> holding a normal WordprocessingML paragraph) rather than
/// Word's own WordArt "text path" mechanism (<c>v:textpath</c> with <c>fitshape</c>, which is what
/// Word's built-in watermark generator produces). Both render correctly in real Word, but
/// LibreOffice's VML importer — used by <see cref="Conversion.WordToPdfConverter"/> — does not
/// support <c>v:textpath</c> at all: it silently drops the shape during docx→PDF conversion,
/// confirmed independent of the shape's size, position, rotation, or z-index.
///
/// If a document with a docx watermark needs to become a non-selectable-watermarked PDF, don't
/// rely on converting the watermarked docx directly — even though <c>v:textbox</c> does survive
/// LibreOffice's conversion (unlike <c>v:textpath</c>), it survives as real selectable PDF text,
/// since that's what it fundamentally is. Use <see cref="RemoveWatermark"/> to strip the watermark
/// before conversion, convert the clean docx, then apply
/// <see cref="Watermarking.PdfWatermarkService"/> to the resulting PDF — that service rasterizes
/// the watermark, which is what actually makes it non-selectable.
/// </remarks>
public sealed class DocxWatermarkService
{
    /// <summary>
    /// Word's Design → Watermark UI (both "Remove Watermark" and the predefined gallery's
    /// replace-existing behavior) doesn't identify a watermark by its appearance or position — it
    /// looks for a header shape whose id matches exactly this prefix followed by digits, which is
    /// what Word's own watermark generator names its shape. A shape with any other id is invisible
    /// to that command entirely: "Remove Watermark" won't find it, and picking a predefined
    /// watermark won't replace it (Word just adds its own shape alongside, producing two).
    /// </summary>
    private const string RemovableShapeIdPrefix = "PowerPlusWaterMarkObject";

    /// <summary>The shape id used for a non-removable (<c>removable: false</c>) watermark.</summary>
    private const string LockedShapeId = "DocProcWatermark";

    /// <summary>
    /// Adds (or replaces) a text watermark on every section of the document. Safe to call more than
    /// once — a prior watermark header added by this method is replaced rather than duplicated.
    /// </summary>
    /// <param name="removable">
    /// When true (the default, matching what Word's own Insert → Watermark produces), the watermark
    /// is user-removable via Word's Design → Watermark → Remove Watermark command, and picking a
    /// predefined watermark from the gallery cleanly replaces it. When false, the watermark uses a
    /// different shape id on purpose so Word's Watermark UI doesn't recognize or manage it — it can
    /// still be deleted by someone editing the document's XML directly, but not with one click from
    /// the ribbon. Use this for content that shouldn't be casually removable (e.g. a legal/compliance
    /// disclaimer) as opposed to a draft-status marker end users are expected to clear themselves.
    /// </param>
    /// <param name="position">Where the watermark sits on the page. Defaults to dead-center, matching Word's own watermark placement.</param>
    /// <param name="widthPt">Width of the shape's bounding box, in points. The text auto-fits to this box (<c>mso-fit-shape-to-text</c>).</param>
    /// <param name="heightPt">Height of the shape's bounding box, in points.</param>
    /// <param name="fontSizePt">Font size of the watermark text, in points.</param>
    public void AddTextWatermark(
        string docxPath,
        string text,
        string fontFamily = "Calibri",
        int rotationDegrees = -45,
        string colorHex = "C0C0C0",
        bool removable = true,
        WatermarkPosition position = WatermarkPosition.Center,
        double widthPt = 415,
        double heightPt = 207.5,
        double fontSizePt = 72)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Document has no main part.");
        var document = mainPart.Document ?? throw new InvalidOperationException("Document has no body.");
        var body = document.Body ?? throw new InvalidOperationException("Document has no body.");

        var shapeId = removable ? $"{RemovableShapeIdPrefix}{Random.Shared.Next(10_000_000, 99_999_999)}" : LockedShapeId;

        var sectionProperties = body.Elements<SectionProperties>().ToList();
        if (sectionProperties.Count == 0)
        {
            sectionProperties.Add(new SectionProperties());
            body.AppendChild(sectionProperties[0]);
        }

        // Compose into whatever default header the section already has, rather than replacing the
        // reference with a watermark-only header. Replacing it silently discarded any existing
        // header content — most visibly a tenant's branding logo, since branding and watermarking
        // are both ordinary steps in the same pipeline and both target the default header.
        var updatedParts = new HashSet<HeaderPart>();
        foreach (var sectPr in sectionProperties)
        {
            var headerPart = ResolveOrCreateDefaultHeader(mainPart, sectPr);

            // Several sections can share one header part; only stamp it once.
            if (!updatedParts.Add(headerPart))
                continue;

            var header = headerPart.Header ??= new Header();

            // Drop any prior watermark so re-applying replaces rather than stacks. Identified by
            // shape id, so unrelated header content is left alone.
            foreach (var run in header.Descendants<Shape>().Where(IsWatermarkShape)
                         .Select(shape => shape.Ancestors<Run>().FirstOrDefault())
                         .Where(run => run is not null).Distinct().ToList())
            {
                run!.Remove();
            }

            header.AppendChild(BuildWatermarkParagraph(text, fontFamily, rotationDegrees, colorHex, shapeId, position, widthPt, heightPt, fontSizePt));
            header.Save();
        }

        document.Save();
    }

    /// <summary>Returns the header part the section's default <c>w:headerReference</c> points at,
    /// creating and wiring one only when the section has none.</summary>
    private static HeaderPart ResolveOrCreateDefaultHeader(MainDocumentPart mainPart, SectionProperties sectPr)
    {
        var existingReference = sectPr.Elements<HeaderReference>()
            .FirstOrDefault(h => h.Type is null || h.Type == HeaderFooterValues.Default);

        if (existingReference?.Id?.Value is { } relId)
        {
            // A reference can dangle if the part was removed by other tooling; fall through to
            // creating a fresh one rather than throwing.
            try
            {
                if (mainPart.GetPartById(relId) is HeaderPart existingPart)
                    return existingPart;
            }
            catch (ArgumentOutOfRangeException)
            {
                existingReference.Remove();
            }
        }

        var headerPart = mainPart.AddNewPart<HeaderPart>();
        headerPart.Header = new Header();
        headerPart.Header.Save();

        foreach (var stale in sectPr.Elements<HeaderReference>()
                     .Where(h => h.Type is null || h.Type == HeaderFooterValues.Default).ToList())
        {
            stale.Remove();
        }

        sectPr.PrependChild(new HeaderReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(headerPart) });
        return headerPart;
    }

    /// <summary>
    /// Removes any watermark previously added by <see cref="AddTextWatermark"/> (removable or
    /// locked), and any native watermark added via Word's own Insert → Watermark (which uses the
    /// same <c>PowerPlusWaterMarkObject*</c> shape id convention this class matches). Identifies
    /// watermark shapes by id rather than assuming a header contains nothing else, so unrelated
    /// header content (e.g. a page number or logo sharing the same header) is left intact.
    /// </summary>
    /// <returns>True if a watermark shape was found and removed from at least one header.</returns>
    public bool RemoveWatermark(string docxPath)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Document has no main part.");

        var removedAny = false;
        foreach (var headerPart in mainPart.HeaderParts)
        {
            var header = headerPart.Header ?? throw new InvalidOperationException("Header part has no content.");

            var watermarkRuns = header.Descendants<Shape>()
                .Where(IsWatermarkShape)
                .Select(shape => shape.Ancestors<Run>().FirstOrDefault())
                .Where(run => run is not null)
                .Distinct()
                .ToList();

            if (watermarkRuns.Count == 0)
                continue;

            foreach (var run in watermarkRuns)
                run!.Remove();

            header.Save();
            removedAny = true;
        }

        return removedAny;
    }

    private static bool IsWatermarkShape(Shape shape) =>
        shape.Id?.Value is { } id && (id.StartsWith(RemovableShapeIdPrefix, StringComparison.Ordinal) || id == LockedShapeId);

    private static Paragraph BuildWatermarkParagraph(
        string text, string fontFamily, int rotationDegrees, string colorHex, string shapeId,
        WatermarkPosition position, double widthPt, double heightPt, double fontSizePt)
    {
        var textRun = new Run(
            new RunProperties(
                new RunFonts { Ascii = fontFamily, HighAnsi = fontFamily, ComplexScript = fontFamily },
                new Color { Val = colorHex },
                new FontSize { Val = ((int)(fontSizePt * 2)).ToString() }), // FontSize is in half-points
            new Text(text));

        var textParagraph = new Paragraph(
            new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
            textRun);

        var textBox = new TextBox(new TextBoxContent(textParagraph)) { Style = "mso-fit-shape-to-text:t" };

        var (horizontal, vertical) = PositionKeywords(position);
        var shape = new Shape(textBox)
        {
            Id = shapeId,
            Style = $"position:absolute;left:0;top:0;width:{widthPt}pt;height:{heightPt}pt;" +
                    $"rotation:{rotationDegrees};z-index:-251654144;" +
                    $"mso-position-horizontal:{horizontal};mso-position-horizontal-relative:margin;" +
                    $"mso-position-vertical:{vertical};mso-position-vertical-relative:margin",
            AllowInCell = false,
            Filled = false,
            Stroked = false
        };

        // The run needs w:noProof — Word's convention for machine-generated graphical content.
        var run = new Run(new RunProperties(new NoProof()), new Picture(shape));
        return new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Header" }), run);
    }

    /// <summary>Maps a <see cref="WatermarkPosition"/> to VML's own <c>mso-position-horizontal</c>/<c>-vertical</c> keywords.</summary>
    private static (string Horizontal, string Vertical) PositionKeywords(WatermarkPosition position) => position switch
    {
        WatermarkPosition.TopLeft => ("left", "top"),
        WatermarkPosition.TopCenter => ("center", "top"),
        WatermarkPosition.TopRight => ("right", "top"),
        WatermarkPosition.MiddleLeft => ("left", "center"),
        WatermarkPosition.MiddleRight => ("right", "center"),
        WatermarkPosition.BottomLeft => ("left", "bottom"),
        WatermarkPosition.BottomCenter => ("center", "bottom"),
        WatermarkPosition.BottomRight => ("right", "bottom"),
        _ => ("center", "center")
    };
}
