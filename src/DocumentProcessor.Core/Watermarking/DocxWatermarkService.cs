using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.Wordprocessing;
using Picture = DocumentFormat.OpenXml.Wordprocessing.Picture;
using VmlLock = DocumentFormat.OpenXml.Vml.Office.Lock;
using VmlPath = DocumentFormat.OpenXml.Vml.Path;

namespace DocumentProcessor.Core.Watermarking;

/// <summary>
/// Adds a diagonal text watermark (e.g. "DRAFT", "CONFIDENTIAL") to a .docx file, using the same
/// VML shape-in-header mechanism Word itself generates for Insert → Watermark — including Word's
/// own "_x0000_t136" WordArt shapetype, which is what lets the shape's fitshape behavior scale the
/// text up to fill the watermark's bounding box. Without it, Word renders the text at its literal
/// nominal size (1pt) instead of stretching it, which looks like the watermark is simply missing.
/// Applied via the document header so it repeats on every page, matching Word's own behavior.
/// </summary>
public sealed class DocxWatermarkService
{
    private const string WatermarkShapeTypeId = "_x0000_t136";

    /// <summary>
    /// Word's Design → Watermark UI (both "Remove Watermark" and the predefined gallery's
    /// replace-existing behavior) doesn't identify a watermark by its appearance or position — it
    /// looks for a header shape whose id matches exactly this prefix followed by digits, which is
    /// what Word's own watermark generator names its shape. A shape with any other id is invisible
    /// to that command entirely: "Remove Watermark" won't find it, and picking a predefined
    /// watermark won't replace it (Word just adds its own shape alongside, producing two).
    /// </summary>
    private const string RemovableShapeIdPrefix = "PowerPlusWaterMarkObject";

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
    public void AddTextWatermark(
        string docxPath,
        string text,
        string fontFamily = "Calibri",
        int rotationDegrees = -45,
        string colorHex = "C0C0C0",
        bool removable = true)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Document has no main part.");
        var document = mainPart.Document ?? throw new InvalidOperationException("Document has no body.");
        var body = document.Body ?? throw new InvalidOperationException("Document has no body.");

        var shapeId = removable ? $"{RemovableShapeIdPrefix}{Random.Shared.Next(10_000_000, 99_999_999)}" : "DocProcWatermark";

        var headerPart = mainPart.AddNewPart<HeaderPart>();
        headerPart.Header = BuildWatermarkHeader(text, fontFamily, rotationDegrees, colorHex, shapeId);
        headerPart.Header.Save();
        var headerRelId = mainPart.GetIdOfPart(headerPart);

        var sectionProperties = body.Elements<SectionProperties>().ToList();
        if (sectionProperties.Count == 0)
        {
            sectionProperties.Add(new SectionProperties());
            body.AppendChild(sectionProperties[0]);
        }

        foreach (var sectPr in sectionProperties)
        {
            foreach (var existing in sectPr.Elements<HeaderReference>()
                .Where(h => h.Type is null || h.Type == HeaderFooterValues.Default).ToList())
            {
                existing.Remove();
            }

            sectPr.PrependChild(new HeaderReference { Type = HeaderFooterValues.Default, Id = headerRelId });
        }

        document.Save();
    }

    private static Header BuildWatermarkHeader(string text, string fontFamily, int rotationDegrees, string colorHex, string shapeId)
    {
        var shape = new Shape(
            new TextPath
            {
                Style = $"font-family:'{fontFamily}';font-size:1pt",
                String = text,
                FitShape = true
            },
            new Fill { Opacity = "0.5" })
        {
            Id = shapeId,
            Type = $"#{WatermarkShapeTypeId}",
            Style = "position:absolute;left:0;top:0;width:415pt;height:207.5pt;" +
                    $"rotation:{rotationDegrees};z-index:-251654144;" +
                    "mso-position-horizontal:center;mso-position-horizontal-relative:margin;" +
                    "mso-position-vertical:center;mso-position-vertical-relative:margin",
            AllowInCell = false,
            Filled = true,
            FillColor = $"#{colorHex}",
            Stroked = false
        };

        // The run needs w:noProof (Word's convention for machine-generated graphical content) and
        // the picture needs both the shapetype definition and the shape that references it.
        var run = new Run(new RunProperties(new NoProof()), new Picture(BuildWatermarkShapetype(), shape));
        var paragraph = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Header" }), run);
        return new Header(paragraph);
    }

    /// <summary>
    /// Word's built-in "Text Plain" WordArt shapetype — stable, unchanged boilerplate present in
    /// every native Word-generated watermark since Office 2003. Defines the outline path that
    /// TextPath.FitShape stretches the watermark text into.
    /// </summary>
    private static Shapetype BuildWatermarkShapetype() =>
        new(
            new Formulas(
                new Formula { Equation = "sum #0 0 10800" },
                new Formula { Equation = "prod #0 2 1" },
                new Formula { Equation = "sum 21600 0 @1" },
                new Formula { Equation = "sum 0 0 @2" },
                new Formula { Equation = "sum 21600 0 @3" },
                new Formula { Equation = "if @0 @3 0" },
                new Formula { Equation = "if @0 21600 @1" },
                new Formula { Equation = "if @0 0 @2" },
                new Formula { Equation = "if @0 @4 21600" },
                new Formula { Equation = "mid @5 @6" },
                new Formula { Equation = "mid @8 @5" },
                new Formula { Equation = "mid @7 @8" },
                new Formula { Equation = "mid @6 @7" },
                new Formula { Equation = "sum @6 0 @5" }),
            new VmlPath
            {
                AllowTextPath = true,
                ConnectionPointType = ConnectValues.Custom,
                ConnectionPoints = "@9,0;@10,10800;@11,21600;@12,10800",
                ConnectAngles = "270,180,90,0"
            },
            new TextPath { On = true, FitShape = true },
            new ShapeHandles(new ShapeHandle { Position = "#0,bottomRight", XRange = "0,21600" }),
            new VmlLock { Extension = ExtensionHandlingBehaviorValues.Edit, TextLock = true, ShapeType = true })
        {
            Id = WatermarkShapeTypeId,
            CoordinateSize = "1600,21600",
            Adjustment = "10800",
            EdgePath = "m@7,0l@8,0m@5,21600l@6,21600e"
        };
}
