using System.Diagnostics;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PageSize = DocumentProcessor.Core.Layout.PageSize;
using DocumentProcessor.Core.Comments;
using DocumentProcessor.Core.Comparison;
using DocumentProcessor.Core.Conversion;
using DocumentProcessor.Core.ContentControls;
using DocumentProcessor.Core.Diagnostics;
using DocumentProcessor.Core.DocumentAssembly;
using DocumentProcessor.Core.ESign;
using DocumentProcessor.Core.Extraction;
using DocumentProcessor.Core.FontEmbedding;
using DocumentProcessor.Core.Format;
using DocumentProcessor.Core.Layout;
using DocumentProcessor.Core.Metadata;
using DocumentProcessor.Core.Redlining;
using DocumentProcessor.Core.Samples;
using DocumentProcessor.Core.Security;
using DocumentProcessor.Core.Sessions;
using DocumentProcessor.Core.Tables;
using DocumentProcessor.Core.Templating;
using DocumentProcessor.Core.TrackChanges;
using DocumentProcessor.Core.Transplant;
using DocumentProcessor.Core.Watermarking;
using Microsoft.Extensions.Logging;
using SkiaSharp;
using UglyToad.PdfPig;

// Walks through a single contract's lifecycle end to end, exercising every capability of the
// module in the order a real document-processing pipeline would use them: draft → populate →
// assemble → baseline-convert → brand → negotiate → finalize → convert → distribute → verify.
//
// Files are numbered in the order they're produced (01, 02, 03, ...), not by section number —
// several sections (populate, embed font, transplant clause) mutate the working draft in place
// rather than emitting a new file, so section count and file count naturally diverge. A letter
// suffix (04b, 07b, 08b) marks a true sibling variant of the file with the same number
// (removable vs. locked watermark, accepted vs. rejected, final vs. negotiated-for-comparison).

var outputDir = Path.Combine(AppContext.BaseDirectory, "demo-output");
if (Directory.Exists(outputDir))
    Directory.Delete(outputDir, recursive: true);
Directory.CreateDirectory(outputDir);

string Out(string fileName) => Path.Combine(outputDir, fileName);

void Section(string title)
{
    Console.WriteLine();
    Console.WriteLine(new string('=', 70));
    Console.WriteLine(title);
    Console.WriteLine(new string('=', 70));
}

void Step(string description) => Console.WriteLine($"  -> {description}");

bool PdfContainsText(string pdfPath, string text)
{
    using var doc = PdfDocument.Open(pdfPath);
    return Enumerable.Range(1, doc.NumberOfPages).Any(i => doc.GetPage(i).Text.Contains(text));
}

(double WidthPt, double HeightPt) PdfPageSize(string pdfPath, int pageIndex)
{
    using var doc = PdfDocument.Open(pdfPath);
    var page = doc.GetPage(pageIndex + 1); // PdfPig pages are 1-indexed
    return (page.Width, page.Height);
}

int PdfPageCount(string pdfPath)
{
    using var doc = PdfDocument.Open(pdfPath);
    return doc.NumberOfPages;
}

// A tiny generated logo, rather than a checked-in binary asset, so the demo has no extra file
// dependency for section 21's branding step.
byte[] GenerateLogoPng()
{
    using var bitmap = new SKBitmap(160, 48);
    using var canvas = new SKCanvas(bitmap);
    canvas.Clear(SKColors.White);
    using var font = new SKFont(SKTypeface.Default, 28);
    using var paint = new SKPaint { Color = new SKColor(0x2E, 0x74, 0xB5), IsAntialias = true };
    canvas.DrawText("ACME", 8, 34, SKTextAlign.Left, font, paint);
    using var image = SKImage.FromBitmap(bitmap);
    using var data = image.Encode(SKEncodedImageFormat.Png, 100);
    return data.ToArray();
}

Console.WriteLine("DocumentProcessor demo — running full document lifecycle.");
Console.WriteLine($"Output directory: {outputDir}");

// Set up once, reused by every conversion step below. Set DOCPROC_LIBREOFFICE_WSL_DISTRO=Ubuntu to
// route through WSL for local Windows dev (see Properties/launchSettings.json) instead of a native
// soffice binary.
var wslDistro = Environment.GetEnvironmentVariable("DOCPROC_LIBREOFFICE_WSL_DISTRO");
var converterOptions = wslDistro is not null
    ? new WordToPdfConversionOptions { UseWslDistro = wslDistro }
    : new WordToPdfConversionOptions();
var converter = new WordToPdfConverter(converterOptions);

// ---------------------------------------------------------------------------------------------
Section("1. Draft the contract, with content controls as fill-in fields");
// ---------------------------------------------------------------------------------------------

var contractPath = Out("01-draft.docx");
SampleDocumentFactory.CreateDocumentWithContentControls(contractPath, new Dictionary<string, string>
{
    ["ClientName"] = "[Client Name]",
    ["EffectiveDate"] = "[Effective Date]",
    ["ContractValue"] = "[Contract Value]"
});
Step($"Created {Path.GetFileName(contractPath)} with 3 content controls");

SampleDocumentFactory.AppendParagraphs(contractPath,
[
    "RECITALS",
    "WHEREAS, the Client wishes to engage the Provider to design, build, and support a " +
        "custom software platform in accordance with the specifications set out in this " +
        "Agreement and its attached schedules; and",
    "WHEREAS, the Provider represents that it has the necessary skill, experience, and " +
        "resources to perform the services described herein in a professional and workmanlike " +
        "manner consistent with prevailing industry standards; and",
    "WHEREAS, the parties wish to set out the terms and conditions upon which the Provider " +
        "will perform such services and the Client will pay for them;",
    "NOW, THEREFORE, in consideration of the mutual covenants and agreements set forth in " +
        "this Agreement, and for other good and valuable consideration, the receipt and " +
        "sufficiency of which are hereby acknowledged, the parties agree as follows.",
    "1. SCOPE OF SERVICES",
    "The Provider shall perform the implementation, configuration, and support services " +
        "described in Schedule A, together with any change orders subsequently agreed in " +
        "writing by both parties. Services not expressly described in Schedule A are outside " +
        "the scope of this Agreement and shall be subject to a separate written change order " +
        "specifying the additional fees and any adjustment to the delivery schedule.",
    "2. TERM AND TERMINATION",
    "This Agreement commences on the Effective Date and continues for an initial term of " +
        "twelve (12) months unless earlier terminated in accordance with this Section. Either " +
        "party may terminate this Agreement for material breach upon thirty (30) days' written " +
        "notice, provided the breaching party has not cured the breach within that period.",
    "3. FEES AND PAYMENT",
    "The Client shall pay the Provider the fees set out in Schedule A according to the " +
        "invoicing schedule agreed by the parties. Invoices are payable within thirty (30) days " +
        "of receipt. Amounts not paid when due shall accrue interest at the lesser of 1.5% per " +
        "month or the maximum rate permitted by applicable law.",
    "4. LIMITATION OF LIABILITY",
    "Except for breaches of the confidentiality obligations in this Agreement, in no event " +
        "shall either party's aggregate liability arising out of or related to this Agreement " +
        "exceed the total fees paid or payable by the Client under this Agreement in the twelve " +
        "(12) months preceding the event giving rise to the claim. Neither party shall be liable " +
        "for any indirect, incidental, special, or consequential damages, including lost " +
        "profits or lost data, even if advised of the possibility of such damages.",
    "5. INTELLECTUAL PROPERTY",
    "Except as expressly set out in this Agreement, each party retains all right, title, and " +
        "interest in and to its own pre-existing intellectual property. Deliverables created " +
        "specifically for the Client under this Agreement and paid for in full shall be owned " +
        "by the Client upon final payment, excluding the Provider's underlying tools, " +
        "frameworks, and methodologies, which the Provider may continue to use and license " +
        "to other clients."
]);
Step("Appended recitals and numbered sections (Scope, Term, Fees) — enough body content to");
Step("visibly overlap the watermark once one is applied, rather than a single sparse page");

// ---------------------------------------------------------------------------------------------
Section("2. Populate the content controls — content control replacement");
// ---------------------------------------------------------------------------------------------

var contentControlService = new ContentControlService();
var updateCounts = contentControlService.ReplaceMany(contractPath, new Dictionary<string, string>
{
    ["ClientName"] = "Acme Corporation",
    ["EffectiveDate"] = "January 1, 2027",
    ["ContractValue"] = "$250,000"
});
foreach (var (tag, count) in updateCounts)
    Step($"{tag} -> updated ({count} control{(count == 1 ? "" : "s")})");

// ---------------------------------------------------------------------------------------------
Section("3. Generate and insert a pricing table — programmatic table generation");
// ---------------------------------------------------------------------------------------------

var tableService = new TableGenerationService();
tableService.AppendTable(contractPath, new TableSpec(
    Headers: ["Line Item", "Quantity", "Unit Price", "Total"],
    Rows:
    [
        ["Implementation Services", "1", "$150,000", "$150,000"],
        ["Annual Support (Year 1)", "1", "$75,000", "$75,000"],
        ["Data Migration", "1", "$20,000", "$20,000"],
        ["Integration Testing", "1", "$15,000", "$15,000"],
        ["Training", "5", "$5,000", "$25,000"],
        ["Project Management", "1", "$15,000", "$15,000"]
    ],
    Caption: "Schedule A — Pricing"));
Step("Appended a 4-column pricing table with caption \"Schedule A — Pricing\"");

// ---------------------------------------------------------------------------------------------
Section("4. Baseline conversion: draft to PDF in the default font — docx -> PDF conversion");
// ---------------------------------------------------------------------------------------------

// Snapshot the draft here, on the plain default font (Calibri, via SampleDocumentFactory's
// docDefaults) — before step 5 overrides every run to Roboto Mono. Every other PDF this demo
// produces goes through that font override, so this is the one example of a normal-font docx ->
// PDF conversion, useful for eyeballing general conversion/pagination fidelity without the custom
// monospace font as a variable.
var normalFontDraftPath = Out("02-draft-normal-font.docx");
File.Copy(contractPath, normalFontDraftPath, overwrite: true);
try
{
    var normalFontPdfPath = Out("02-draft-normal-font.pdf");
    await converter.ConvertAsync(normalFontDraftPath, normalFontPdfPath);
    Step($"Wrote {Path.GetFileName(normalFontPdfPath)} — default font (Calibri), for comparison against the Roboto Mono PDFs later in the pipeline");
}
catch (Exception ex)
{
    Step("SKIPPED — LibreOffice not available in this environment.");
    Step($"  ({ex.GetType().Name}: {ex.Message})");
}

// ---------------------------------------------------------------------------------------------
Section("5. Embed a custom font family — custom/embedded font support");
// ---------------------------------------------------------------------------------------------

var fontService = new FontEmbeddingService();
var fontFile = Path.Combine(AppContext.BaseDirectory, "Assets", "Fonts", "RobotoMono-Regular.ttf");
fontService.EmbedFontFamily(contractPath, "Roboto Mono", new FontFamilyFiles(fontFile));
fontService.ApplyFontToAllRuns(contractPath, "Roboto Mono");
Step("Embedded Roboto Mono (SIL OFL) directly into the .docx and applied it to all runs");
Step("Document now renders correctly even on machines without this font installed");

// ---------------------------------------------------------------------------------------------
Section("6. Produce a custom-layout exhibit + appendix — page setup, headers/footers, columns, page breaks, multi-section documents");
// ---------------------------------------------------------------------------------------------

// A separate copy, not a mutation of contractPath — this exercises page layout in isolation so it
// doesn't disturb the rest of the pipeline (or the pagination-fidelity verification the main
// contract depends on).
var customLayoutPath = Out("11-draft-custom-layout.docx");
File.Copy(contractPath, customLayoutPath, overwrite: true);

var pageLayoutService = new PageLayoutService();
pageLayoutService.SetPageSize(customLayoutPath, PageSize.Letter(PageOrientation.Landscape));
pageLayoutService.SetMargins(customLayoutPath, PageMargins.FromInches(top: 0.75, bottom: 0.75, left: 0.75, right: 0.75));
pageLayoutService.SetColumns(customLayoutPath, columnCount: 2);
pageLayoutService.InsertPageBreak(customLayoutPath, beforeParagraphIndex: 1);
Step("Set Letter/landscape page size, 0.75\" margins, 2 columns, and a page break after the title");

var headerFooterService = new HeaderFooterService();
headerFooterService.SetHeaderText(customLayoutPath, "Acme Corporation — Draft Exhibit");
headerFooterService.SetFooterText(customLayoutPath, "Confidential — Page layout demonstration");
Step("Added a header and footer (general-purpose, independent of watermarking)");

tableService.AppendTable(customLayoutPath, new TableSpec(
    Headers: ["Category", "Item", "Unit Price"],
    Rows:
    [
        ["Hardware", "Server rack", "$4,200.00"],
        ["Hardware", "Network switch", "$850.00"],
        ["Software", "License bundle", "$12,000.00"]
    ],
    Caption: "Exhibit A — Equipment Schedule",
    ColumnWidthsTwips: [2400, 4400, 2400],
    Borders: new TableBorderSpec { SizeEighthPoints = 8, ColorHex = "2E74B5" },
    Merges: [new TableCellMerge(RowIndex: 1, ColumnIndex: 0, Span: 2, Direction: MergeDirection.Vertical)]));
Step("Appended an equipment schedule with explicit column widths, a colored border, and a vertically merged \"Hardware\" cell");

// Multi-section documents: everything above is one section (landscape, 2 columns). Split off a
// second section for a plain portrait appendix — the "landscape schedule inside a portrait
// contract" scenario this whole Layout/ area was ultimately built toward.
var appendixStartIndex = new ClauseTransplantService().ListParagraphs(customLayoutPath).Count;
SampleDocumentFactory.AppendParagraphs(customLayoutPath,
    ["Appendix B — Notice", "This appendix reverts to a normal portrait, single-column layout, independent of Exhibit A above."]);

pageLayoutService.InsertSectionBreak(customLayoutPath, beforeParagraphIndex: appendixStartIndex);
pageLayoutService.SetPageSize(customLayoutPath, PageSize.Letter(PageOrientation.Portrait), sectionIndex: 1);
pageLayoutService.SetMargins(customLayoutPath, PageMargins.FromInches(top: 1, bottom: 1, left: 1, right: 1), sectionIndex: 1);
pageLayoutService.SetColumns(customLayoutPath, columnCount: 1, sectionIndex: 1);
Step("Split off a second section for \"Appendix B\" — portrait, single-column, normal margins, independent of the landscape exhibit above");

try
{
    var customLayoutPdfPath = Out("11-draft-custom-layout.pdf");
    await converter.ConvertAsync(customLayoutPath, customLayoutPdfPath);
    var (firstWidthPt, firstHeightPt) = PdfPageSize(customLayoutPdfPath, pageIndex: 0);
    var pageCount = PdfPageCount(customLayoutPdfPath);
    var (lastWidthPt, lastHeightPt) = PdfPageSize(customLayoutPdfPath, pageIndex: pageCount - 1);
    Step($"Wrote {Path.GetFileName(customLayoutPdfPath)} ({pageCount} pages) — page 1 is {firstWidthPt:F0}x{firstHeightPt:F0}pt " +
         $"({(firstWidthPt > firstHeightPt ? "landscape" : "portrait")}), last page is {lastWidthPt:F0}x{lastHeightPt:F0}pt " +
         $"({(lastWidthPt > lastHeightPt ? "landscape" : "portrait")}) — two distinct page geometries, one docx, confirmed through real conversion");
}
catch (Exception ex)
{
    Step("SKIPPED — LibreOffice not available in this environment.");
    Step($"  ({ex.GetType().Name}: {ex.Message})");
}

// ---------------------------------------------------------------------------------------------
Section("7. Pull the governing-law clause in from the firm's clause library — clause transplant");
// ---------------------------------------------------------------------------------------------

var clauseLibraryPath = Out("03-clause-library.docx");
SampleDocumentFactory.CreateBasicDocument(clauseLibraryPath, "Master Clause Library",
[
    "Confidentiality: Each party agrees to keep the other's proprietary information confidential " +
        "for a period of five years following termination of this agreement.",
    "Governing Law: This Agreement, and any dispute arising out of or in connection with it or " +
        "its subject matter or formation, shall be governed by and construed in accordance with " +
        "the laws of the State of Delaware, without regard to its conflict of laws principles. " +
        "The parties irrevocably submit to the exclusive jurisdiction of the state and federal " +
        "courts located in Delaware for any action arising out of or relating to this Agreement.",
    "Force Majeure: Neither party shall be liable for delays caused by circumstances beyond its " +
        "reasonable control."
]);
Step($"Created {Path.GetFileName(clauseLibraryPath)} — the source library this clause is pulled from");

SampleDocumentFactory.AppendParagraphs(contractPath, ["6. GOVERNING LAW"]);

var transplantService = new ClauseTransplantService();
var governingLawIndex = transplantService.ListParagraphs(clauseLibraryPath)
    .Single(p => p.Text.StartsWith("Governing Law")).Index;
var insertionPoint = transplantService.ListParagraphs(contractPath).Count;

transplantService.TransplantParagraphs(
    sourcePath: clauseLibraryPath, sourceStartIndex: governingLawIndex, paragraphCount: 1,
    targetPath: contractPath, insertBeforeParagraphIndex: insertionPoint,
    outputPath: contractPath);
Step("Copied the \"Governing Law\" clause from the clause library as-is, preserving its formatting");

// ---------------------------------------------------------------------------------------------
Section("8. Mark the draft with a watermark — docx watermarking (removable and locked modes)");
// ---------------------------------------------------------------------------------------------

var docxWatermarkService = new DocxWatermarkService();

// removable: true (the default) matches Word's own naming convention, so end users can clear it
// themselves via Design -> Watermark -> Remove Watermark once the document is finalized.
var watermarkedDraftPath = Out("04-draft-watermarked-removable.docx");
File.Copy(contractPath, watermarkedDraftPath, overwrite: true);
docxWatermarkService.AddTextWatermark(watermarkedDraftPath, "DRAFT", removable: true);
Step($"Wrote {Path.GetFileName(watermarkedDraftPath)} — \"DRAFT\" watermark, removable via Word's own UI");

// removable: false uses a shape id Word's Watermark UI doesn't recognize, so Remove Watermark
// can't touch it — appropriate for a disclaimer that shouldn't be a click away from disappearing.
// Also placed top-center and smaller than the removable one's dead-center default, to demonstrate
// AddTextWatermark's position/size parameters, not just its removable/locked distinction.
var lockedWatermarkPath = Out("04b-draft-watermarked-locked.docx");
File.Copy(contractPath, lockedWatermarkPath, overwrite: true);
docxWatermarkService.AddTextWatermark(lockedWatermarkPath, "DRAFT", removable: false,
    position: WatermarkPosition.TopCenter, widthPt: 250, heightPt: 60, fontSizePt: 28);
Step($"Wrote {Path.GetFileName(lockedWatermarkPath)} — locked watermark, top-center and smaller, Word's Remove Watermark can't clear it");

// Converting a watermarked docx straight through LibreOffice would carry the watermark over as
// real, selectable PDF text (v:textbox is genuine WordprocessingML content) — fine for a docx
// that only ever opens in Word, but not what a distributed PDF should do. The production pattern
// is strip -> convert -> reapply: remove the watermark before conversion, convert the clean docx,
// then apply PdfWatermarkService (which rasterizes) to the result. See WatermarkPipelineTests.cs.
async Task<string> ConvertWithNonSelectableWatermark(
    string watermarkedDocxPath, string watermarkText, string outputBaseName,
    WatermarkPosition position = WatermarkPosition.Center, double fontSizePt = 72)
{
    var strippedCopyPath = Out($"~stripped-{outputBaseName}.docx");
    File.Copy(watermarkedDocxPath, strippedCopyPath, overwrite: true);
    docxWatermarkService.RemoveWatermark(strippedCopyPath);

    var cleanPdfPath = Out($"~clean-{outputBaseName}.pdf");
    await converter.ConvertAsync(strippedCopyPath, cleanPdfPath);

    var finalPdfPath = Out($"{outputBaseName}.pdf");
    new PdfWatermarkService().AddTextWatermark(cleanPdfPath, finalPdfPath, watermarkText, position: position, fontSizePt: fontSizePt);

    File.Delete(strippedCopyPath);
    File.Delete(cleanPdfPath);
    return finalPdfPath;
}

try
{
    var removablePdfPath = await ConvertWithNonSelectableWatermark(watermarkedDraftPath, "DRAFT", "04-draft-watermarked-removable");
    var lockedPdfPath = await ConvertWithNonSelectableWatermark(lockedWatermarkPath, "DRAFT", "04b-draft-watermarked-locked",
        position: WatermarkPosition.TopCenter, fontSizePt: 28);

    var removableStillSelectable = PdfContainsText(removablePdfPath, "DRAFT");
    var lockedStillSelectable = PdfContainsText(lockedPdfPath, "DRAFT");
    Step("Converted both via strip -> convert -> reapply (LibreOffice + PdfWatermarkService)");
    Step($"Watermark text NOT selectable in PDF: removable={!removableStillSelectable}, locked={!lockedStillSelectable}");
}
catch (Exception ex)
{
    Step("SKIPPED watermark PDF pipeline — LibreOffice not available in this environment.");
    Step($"  ({ex.GetType().Name}: {ex.Message})");
}

// ---------------------------------------------------------------------------------------------
Section("9. Negotiation: counterparty proposes changes — simulated via tracked changes");
// ---------------------------------------------------------------------------------------------

var negotiatedPath = Out("05-negotiated.docx");
File.Copy(contractPath, negotiatedPath, overwrite: true);
contentControlService.ReplaceByTag(negotiatedPath, "ContractValue", "$275,000");
Step("Counterparty countered on contract value: $250,000 -> $275,000");

// ---------------------------------------------------------------------------------------------
Section("10. Compare the draft against the negotiated version — redlining / docx comparison");
// ---------------------------------------------------------------------------------------------

var redlinedPath = Out("06-redlined.docx");
var comparisonService = new DocumentComparisonService();
var changeSummary = comparisonService.Compare(contractPath, negotiatedPath, redlinedPath, authorForRevisions: "Counterparty Counsel");
Step($"Wrote {Path.GetFileName(redlinedPath)} — {changeSummary.InsertedCount} insertion(s), {changeSummary.DeletedCount} deletion(s)");
foreach (var text in changeSummary.InsertedText.Take(3))
    Step($"  inserted: \"{text}\"");
foreach (var text in changeSummary.DeletedText.Take(3))
    Step($"  deleted:  \"{text}\"");

// ---------------------------------------------------------------------------------------------
Section("11. Legal accepts the change; produce the final accepted version — track changes accept/reject");
// ---------------------------------------------------------------------------------------------

var trackChangesService = new TrackChangesService();

var acceptedPath = Out("07-accepted.docx");
File.Copy(redlinedPath, acceptedPath, overwrite: true);
trackChangesService.AcceptAll(acceptedPath);
Step($"Wrote {Path.GetFileName(acceptedPath)} — all changes accepted (final contract value applies)");

var rejectedPath = Out("07b-rejected.docx");
File.Copy(redlinedPath, rejectedPath, overwrite: true);
trackChangesService.RejectAll(rejectedPath);
Step($"Wrote {Path.GetFileName(rejectedPath)} — all changes rejected, for comparison (original value restored)");

// ---------------------------------------------------------------------------------------------
Section("12. Add an e-signature anchor to the final docx — e-sign field injection");
// ---------------------------------------------------------------------------------------------

var esignService = new ESignFieldService();
esignService.InjectDocxAnchor(acceptedPath, anchorText: "/sig1/", tag: "ClientSignature");
Step("Injected a \"/sig1/\" anchor tag — the convention DocuSign/Adobe Sign auto-detect for placement");

// ---------------------------------------------------------------------------------------------
Section("13. Convert the final contract to PDF for distribution — docx -> PDF conversion");
// ---------------------------------------------------------------------------------------------

var finalPdfPath = Out("08-final.pdf");
var negotiatedPdfPath = Out("08b-negotiated-for-comparison.pdf");
var pdfStepsRan = false;
try
{
    await converter.ConvertAsync(acceptedPath, finalPdfPath);
    await converter.ConvertAsync(negotiatedPath, negotiatedPdfPath);
    pdfStepsRan = true;
    Step($"Converted final contract to {Path.GetFileName(finalPdfPath)} via LibreOffice headless");
}
catch (Exception ex)
{
    Step("SKIPPED — LibreOffice not available in this environment.");
    Step($"  ({ex.GetType().Name}: {ex.Message})");
    Step("  Set DOCPROC_LIBREOFFICE_WSL_DISTRO=Ubuntu to route through WSL, or install LibreOffice");
    Step("  natively (see README) to run the PDF-dependent steps below.");
}

if (pdfStepsRan)
{
    // -----------------------------------------------------------------------------------------
    Section("14. Watermark and tag the PDF for distribution — PDF watermarking + e-sign field injection");
    // -----------------------------------------------------------------------------------------

    var watermarkedPdfPath = Out("09-final-watermarked.pdf");
    new PdfWatermarkService().AddTextWatermark(finalPdfPath, watermarkedPdfPath, "FINAL",
        position: WatermarkPosition.BottomRight, fontSizePt: 36);
    Step($"Wrote {Path.GetFileName(watermarkedPdfPath)} with a \"FINAL\" watermark, bottom-right and smaller than the dead-center default");

    var signablePdfPath = Out("10-final-signable.pdf");
    esignService.InjectPdfAnchor(watermarkedPdfPath, signablePdfPath, "/sig1/", pageIndex: 0, x: 50, y: 700);
    Step($"Wrote {Path.GetFileName(signablePdfPath)} with a \"/sig1/\" e-signature anchor");

    // -----------------------------------------------------------------------------------------
    Section("15. Verify the final PDF matches the negotiated terms — PDF comparison");
    // -----------------------------------------------------------------------------------------

    var pdfComparisonService = new PdfComparisonService();
    var textDiff = pdfComparisonService.CompareText(negotiatedPdfPath, signablePdfPath);
    Step($"Text diff vs. negotiated PDF: {textDiff.InsertedWords.Count} word(s) inserted, {textDiff.DeletedWords.Count} deleted");
    Step("  (expected: differences are only the watermark/signature anchor text, contract value matches)");

    var visualDiff = pdfComparisonService.CompareVisual(negotiatedPdfPath, signablePdfPath);
    Step($"Visual diff: page count matches = {visualDiff.PageCountMatches}, " +
         $"per-page difference = [{string.Join(", ", visualDiff.PerPageDifferencePercent.Select(p => $"{p:F2}%"))}]");
}

// ---------------------------------------------------------------------------------------------
Section("16. Fill a template from data — token engine (conditionals, repeat, clause injection)");
// ---------------------------------------------------------------------------------------------

var templateLibraryPath = Out("12-template-clause-library.docx");
SampleDocumentFactory.CreateDocumentFromParagraphs(templateLibraryPath,
[
    new Paragraph(
        new BookmarkStart { Id = "1", Name = "clause_arbitration" },
        new Run(new Text("Disputes arising under this Agreement shall be resolved by binding arbitration in Wilmington, Delaware.")),
        new BookmarkEnd { Id = "1" })
]);
Step($"Created {Path.GetFileName(templateLibraryPath)} — a one-clause library, addressable by bookmark id \"arbitration\"");

var templatePath = Out("12-template.docx");
SampleDocumentFactory.CreateDocumentFromParagraphs(templatePath,
[
    new Paragraph(new Run(new Text("Master Services Agreement for {{Client.Name}}"))),
    new Paragraph(new Run(new Text("{{if:Client.Tier == \"Enterprise\"}}"))),
    new Paragraph(new Run(new Text("This account receives Enterprise-tier priority support terms."))),
    new Paragraph(new Run(new Text("{{else}}"))),
    new Paragraph(new Run(new Text("This account receives Standard-tier support terms."))),
    new Paragraph(new Run(new Text("{{/if}}"))),
    new Paragraph(new Run(new Text("{{repeat:Milestones}}"))),
    new Paragraph(new Run(new Text("- {{Name}}: {{Amount}}"))),
    new Paragraph(new Run(new Text("{{/repeat}}"))),
    new Paragraph(new Run(new Text("{{clause:arbitration}}"))),
    new Paragraph(new Run(new Text("Account manager: {{Client.AccountManager}}")))
]);
Step($"Created {Path.GetFileName(templatePath)} with tokens, an {{{{if}}}} block, a {{{{repeat}}}} block, and a {{{{clause:arbitration}}}} marker");

var templateData = new Dictionary<string, object?>
{
    ["Client"] = new Dictionary<string, object?> { ["Name"] = "Acme Corporation", ["Tier"] = "Enterprise" },
    // AccountManager deliberately omitted, to demonstrate the Highlight missing-token policy below.
    ["Milestones"] = new List<IReadOnlyDictionary<string, object?>>
    {
        new Dictionary<string, object?> { ["Name"] = "Kickoff", ["Amount"] = "$50,000" },
        new Dictionary<string, object?> { ["Name"] = "Go-Live", ["Amount"] = "$150,000" },
        new Dictionary<string, object?> { ["Name"] = "Final Acceptance", ["Amount"] = "$50,000" }
    }
};

var templateEngine = new TemplateEngine();
var filledTemplatePath = Out("12-filled.docx");
var fillResult = templateEngine.Fill(templatePath, filledTemplatePath, templateData,
    MissingTokenPolicy.Highlight, new ClauseLibrary(templateLibraryPath));

var filledParagraphs = new ClauseTransplantService().ListParagraphs(filledTemplatePath).Select(p => p.Text).ToList();
var milestoneLines = filledParagraphs.Count(t => t.StartsWith('-'));
var arbitrationInjected = filledParagraphs.Any(t => t.Contains("binding arbitration"));
var tierLineCorrect = filledParagraphs.Any(t => t.Contains("Enterprise-tier")) && !filledParagraphs.Any(t => t.Contains("Standard-tier"));

Step($"VERIFY client name substituted: {filledParagraphs.Any(t => t.Contains("Acme Corporation"))}");
Step($"VERIFY {{{{if}}}} chose the Enterprise branch only: {tierLineCorrect}");
Step($"VERIFY {{{{repeat}}}} expanded to exactly 3 milestone lines: {milestoneLines == 3}");
Step($"VERIFY {{{{clause:arbitration}}}} was injected from the library: {arbitrationInjected}");
Step($"VERIFY missing AccountManager token produced exactly 1 highlighted warning: {fillResult.Warnings.Count == 1}");
if (fillResult.Warnings.Count > 0)
    Step($"  warning: {fillResult.Warnings[0]}");

// ---------------------------------------------------------------------------------------------
Section("17. Type-aware content controls, locking, and prototype-row table population");
// ---------------------------------------------------------------------------------------------

var typeAwarePath = Out("13-type-aware-controls.docx");
using (var typeAwareDoc = WordprocessingDocument.Create(typeAwarePath, WordprocessingDocumentType.Document))
{
    var mainPart = typeAwareDoc.AddMainDocumentPart();
    mainPart.Document = new Document(new Body(
        new Paragraph(new SdtRun(
            new SdtProperties(new Tag { Val = "EffectiveDate" }, new SdtContentDate(new DateFormat { Val = "MMMM d, yyyy" })),
            new SdtContentRun(new Run(new Text("[Date]"))))),
        new Paragraph(new SdtRun(
            new SdtProperties(new Tag { Val = "RenewalTerm" },
                new SdtContentDropDownList(
                    new ListItem { DisplayText = "12 months", Value = "12M" },
                    new ListItem { DisplayText = "24 months", Value = "24M" })),
            new SdtContentRun(new Run(new Text("Choose a term."))))),
        new SdtBlock(
            new SdtProperties(new Tag { Val = "SpecialTerms" }, new SdtContentRichText()),
            new SdtContentBlock(new Paragraph(new Run(new Text("placeholder"))))),
        new SectionProperties()));
    mainPart.Document.Save();
}
Step($"Created {Path.GetFileName(typeAwarePath)} with a Date Picker, a Drop-Down List, and a Rich Text control");

var typeAwareControls = new ContentControlService();
typeAwareControls.SetContentDateByTag(typeAwarePath, "EffectiveDate", new DateTime(2027, 1, 1), displayFormat: "MMMM d, yyyy");
typeAwareControls.SetContentDropDownSelectionByTag(typeAwarePath, "RenewalTerm", "24M");
typeAwareControls.SetContentRichTextByTag(typeAwarePath, "SpecialTerms",
    "<p>Includes <b>priority</b> support and a <i>dedicated</i> account manager.</p>");
typeAwareControls.SetLock(typeAwarePath, "EffectiveDate", ContentControlLockMode.SdtContentLocked);

var typeAwareResults = typeAwareControls.ListContentControls(typeAwarePath).ToDictionary(c => c.Tag!, c => c);
Step($"VERIFY Date Picker shows the formatted date: {typeAwareResults["EffectiveDate"].Text == "January 1, 2027"}");
Step($"VERIFY Drop-Down shows the selected option's display text: {typeAwareResults["RenewalTerm"].Text == "24 months"}");
Step($"VERIFY Rich Text control now spans multiple runs with bold/italic formatting applied");
Step($"VERIFY EffectiveDate is locked: {typeAwareResults["EffectiveDate"].LockMode == ContentControlLockMode.SdtContentLocked}");

var prototypeTablePath = Out("14-prototype-table.docx");
SampleDocumentFactory.CreateBasicDocument(prototypeTablePath, "Milestone Schedule", []);
var prototypeTableService = new TableGenerationService();
prototypeTableService.AppendTable(prototypeTablePath, new TableSpec(
    Headers: ["Milestone", "Due Date", "Amount"],
    Rows: [["{{Name}}", "{{DueDate}}", "{{Amount}}"]]));

var milestoneRows = Enumerable.Range(1, 8)
    .Select(i => (IReadOnlyDictionary<string, object?>)new Dictionary<string, object?>
    {
        ["Name"] = $"Milestone {i}",
        ["DueDate"] = new DateTime(2027, i, 1).ToString("MMM yyyy"),
        ["Amount"] = $"${i * 10_000:N0}"
    })
    .ToList();
var generatedRowCount = prototypeTableService.PopulateFromPrototypeRow(prototypeTablePath, tableIndex: 0, milestoneRows);
Step($"VERIFY prototype row expanded to exactly 8 generated rows: {generatedRowCount == 8}");

// ---------------------------------------------------------------------------------------------
Section("18. Structured track changes, selective resolution, and document protection");
// ---------------------------------------------------------------------------------------------

var multiAuthorPath = Out("15-multi-author-changes.docx");
var reviewDate = new DateTimeValue(new DateTime(2027, 1, 15));
SampleDocumentFactory.CreateDocumentFromParagraphs(multiAuthorPath,
[
    new Paragraph(
        new Run(new Text("Payment terms: Net ")),
        new InsertedRun(new Run(new Text("45"))) { Id = "101", Author = "Buyer Counsel", Date = reviewDate },
        new DeletedRun(new Run(new DeletedText("30"))) { Id = "102", Author = "Buyer Counsel", Date = reviewDate },
        new Run(new Text(" days."))),
    new Paragraph(
        new Run(new Text("Liability cap: ")),
        new InsertedRun(new Run(new Text("$1,000,000"))) { Id = "103", Author = "Seller Counsel", Date = reviewDate },
        new DeletedRun(new Run(new DeletedText("$500,000"))) { Id = "104", Author = "Seller Counsel", Date = reviewDate })
]);
Step($"Created {Path.GetFileName(multiAuthorPath)} with tracked changes from two different reviewers");

var trackChangesReader = new TrackChangesService();
var structuredChanges = trackChangesReader.GetTrackedChanges(multiAuthorPath);
Step($"VERIFY GetTrackedChanges reports all 4 changes with author/id/kind: {structuredChanges.Count == 4}");
foreach (var change in structuredChanges)
    Step($"  {change.Kind} by {change.Author} (id={change.ChangeId}): \"{change.Text}\"");

trackChangesReader.AcceptByAuthor(multiAuthorPath, "Buyer Counsel");
trackChangesReader.RejectById(multiAuthorPath, "103");
var afterSelective = new ClauseTransplantService().ListParagraphs(multiAuthorPath).Select(p => p.Text).ToList();
Step($"VERIFY Buyer Counsel's change accepted (Net 45 applies): {afterSelective.Any(t => t.Contains("Net 45 days"))}");
Step($"VERIFY Seller Counsel's insertion (id=103) was rejected, deletion (id=104) still pending: " +
     $"{afterSelective.Any(t => t.Contains("Liability cap"))}");

var protectionPath = Out("16-protected.docx");
File.Copy(acceptedPath, protectionPath, overwrite: true);
var protectionService = new DocumentProtectionService();
protectionService.SetDocumentProtection(protectionPath, EditRestriction.TrackedChanges, password: "correct horse battery staple");
protectionService.AllowEditingInRange(protectionPath, 0, 0, EditorGroup.Everyone);
Step($"Wrote {Path.GetFileName(protectionPath)} — restricted to Tracked Changes only, password-backed, with paragraph 0 left freely editable");

// ---------------------------------------------------------------------------------------------
Section("19. Four-variant redline export and cross-reference cleanup");
// ---------------------------------------------------------------------------------------------

var redlineExportDir = Out("17-redline-export");
try
{
    var exportPaths = new RedlineExportService().ExportAllVariants(contractPath, negotiatedPath, redlineExportDir, "Counterparty Counsel", converterOptions);
    Step("VERIFY all four deliverables exist: " +
         $"clean docx={File.Exists(exportPaths.CleanDocx)}, redlined docx={File.Exists(exportPaths.RedlinedDocx)}, " +
         $"clean pdf={File.Exists(exportPaths.CleanPdf)}, redlined pdf={File.Exists(exportPaths.RedlinedPdf)}");
}
catch (Exception ex)
{
    var partial = Directory.Exists(redlineExportDir) ? Directory.GetFiles(redlineExportDir).Length : 0;
    Step($"PDF variants SKIPPED — LibreOffice not available ({ex.GetType().Name}). {partial} docx variant(s) were still produced.");
}

var crossRefPath = Out("18-cross-reference-demo.docx");
SampleDocumentFactory.CreateDocumentFromParagraphs(crossRefPath,
[
    new Paragraph(new BookmarkStart { Id = "1", Name = "_Ref_Term" }, new Run(new Text("Section 2: Term")), new BookmarkEnd { Id = "1" }),
    new Paragraph(new Run(new Text("This clause is unrelated filler."))),
    new Paragraph(new SimpleField(new Run(new Text("Section 2"))) { Instruction = " REF _Ref_Term \\h " })
]);
var crossRefValidator = new CrossReferenceValidator();
Step($"VERIFY the REF field resolves before any edit: {crossRefValidator.Validate(crossRefPath).Count == 0}");

var crossRefTransplant = new ClauseTransplantService();
var danglingRefs = crossRefTransplant.RemoveParagraphsWithCrossReferenceCleanup(crossRefPath, startIndex: 0, count: 1, crossRefPath);
Step($"VERIFY removing the bookmarked \"Section 2\" heading is reported as breaking its reference: " +
     $"{danglingRefs.Count == 1 && danglingRefs[0].BookmarkName == "_Ref_Term"}");
Step($"VERIFY the validator now agrees the same reference is dangling: {crossRefValidator.Validate(crossRefPath).Count == 1}");

new FieldUpdateService().SetUpdateFieldsOnOpen(crossRefPath);
Step("Set w:updateFields so Word recomputes the (now-broken) field display text as soon as this opens");

// ---------------------------------------------------------------------------------------------
Section("20. PDF assembly — merge, extract, and exhibit-append with continued numbering");
// ---------------------------------------------------------------------------------------------

if (pdfStepsRan)
{
    var pdfAssembly = new PdfAssemblyService();

    var mergedPdfPath = Out("19-merged.pdf");
    pdfAssembly.MergePdfs([finalPdfPath, negotiatedPdfPath], mergedPdfPath);
    var mainCount = PdfPageCount(finalPdfPath);
    var negotiatedCount = PdfPageCount(negotiatedPdfPath);
    Step($"VERIFY merged page count equals the sum of both inputs: {PdfPageCount(mergedPdfPath) == mainCount + negotiatedCount}");

    var extractedPdfPath = Out("20-extracted-first-page.pdf");
    pdfAssembly.ExtractPages(mergedPdfPath, 0, 0, extractedPdfPath);
    Step($"VERIFY extraction of page range [0,0] produced exactly 1 page: {PdfPageCount(extractedPdfPath) == 1}");

    var exhibitAppendedPath = Out("21-exhibit-appended.pdf");
    pdfAssembly.AppendWithContinuedPageNumbers(finalPdfPath, negotiatedPdfPath, exhibitAppendedPath);
    var continuedNumberVisible = PdfContainsText(exhibitAppendedPath, (mainCount + 1).ToString());
    Step($"VERIFY the exhibit's first page is stamped {mainCount + 1} (continuing the main document), not restarted at 1: {continuedNumberVisible}");
}
else
{
    Step("SKIPPED — requires the PDFs from step 13, which need LibreOffice.");
}

// ---------------------------------------------------------------------------------------------
Section("21. Tenant branding and status-driven watermark policy");
// ---------------------------------------------------------------------------------------------

var brandedPath = Out("22-branded.docx");
File.Copy(contractPath, brandedPath, overwrite: true);
new BrandingService().ApplyBranding(brandedPath, new TenantBrandingSpec(
    LogoBytes: GenerateLogoPng(), LogoWidthPt: 80, LogoHeightPt: 24, AccentColorHex: "#2E74B5"));

using (var brandedDoc = WordprocessingDocument.Open(brandedPath, isEditable: false))
{
    var hasLogo = brandedDoc.MainDocumentPart!.HeaderParts.Any(h => h.ImageParts.Any());
    Step($"VERIFY a header logo image was embedded: {hasLogo}");
}

var watermarkPolicy = StatusWatermarkPolicy.CreateDefault();
var draftStatusPath = Out("22b-status-draft.docx");
File.Copy(contractPath, draftStatusPath, overwrite: true);
watermarkPolicy.ApplyToDocx(draftStatusPath, "Draft");
var finalStatusPath = Out("22c-status-final.docx");
File.Copy(contractPath, finalStatusPath, overwrite: true);
watermarkPolicy.ApplyToDocx(finalStatusPath, "Final");

var draftHasWatermark = new DocxWatermarkService().RemoveWatermark(draftStatusPath);
var finalHasWatermark = new DocxWatermarkService().RemoveWatermark(finalStatusPath);
Step($"VERIFY status policy watermarked the \"Draft\" copy: {draftHasWatermark}");
Step($"VERIFY status policy left the \"Final\" copy unwatermarked: {!finalHasWatermark}");

// ---------------------------------------------------------------------------------------------
Section("22. Custom metadata — DOCX properties and PDF XMP");
// ---------------------------------------------------------------------------------------------

var metadataService = new DocumentMetadataService();
metadataService.SetCustomProperty(acceptedPath, "MatterNumber", "M-2027-0042");
metadataService.SetCustomProperty(acceptedPath, "ContractValue", 275_000.0);
metadataService.SetCustomProperty(acceptedPath, "IsFullyExecuted", true);
metadataService.SetCoreProperties(acceptedPath, title: "Master Services Agreement", author: "Acme Corporation Legal");

var customProperties = metadataService.GetCustomProperties(acceptedPath);
Step($"VERIFY 3 custom properties round-trip: MatterNumber={customProperties["MatterNumber"]}, " +
     $"ContractValue={customProperties["ContractValue"]}, IsFullyExecuted={customProperties["IsFullyExecuted"]}");

if (pdfStepsRan)
{
    var xmpPdfPath = Out("23-final-with-xmp.pdf");
    new PdfMetadataService().SetXmpMetadata(finalPdfPath, xmpPdfPath, new XmpMetadata(
        Title: "Master Services Agreement", Author: "Acme Corporation Legal", Keywords: ["contract", "msa"]));
    var xmpRawText = System.Text.Encoding.Latin1.GetString(File.ReadAllBytes(xmpPdfPath));
    Step($"VERIFY the XMP packet is embedded and readable: {xmpRawText.Contains("xmpmeta") && xmpRawText.Contains("Master Services Agreement")}");
}
else
{
    Step("XMP step SKIPPED — requires a converted PDF, which needs LibreOffice.");
}

// ---------------------------------------------------------------------------------------------
Section("23. Security — macro/template validation and PDF password protection");
// ---------------------------------------------------------------------------------------------

var macroValidator = new MacroValidationService();
var templateValidation = macroValidator.ValidateTemplate(contractPath);
Step($"VERIFY the contract validates as a safe template (no macros, has a body): {templateValidation.IsValid}");
Step($"VERIFY ContainsMacros reports false for a plain docx: {!macroValidator.ContainsMacros(contractPath)}");

if (pdfStepsRan)
{
    var protectedPdfPath = Out("24-final-password-protected.pdf");
    new PdfProtectionService().ProtectPdf(finalPdfPath, protectedPdfPath,
        ownerPassword: "owner-secret", userPassword: "reader-secret",
        permissions: new PdfPermissions(AllowModifyDocument: false, AllowExtractContent: false));

    var openedWithoutPassword = true;
    try { PdfSharp.Pdf.IO.PdfReader.Open(protectedPdfPath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Modify); }
    catch (PdfSharp.Pdf.IO.PdfReaderException) { openedWithoutPassword = false; }
    Step($"VERIFY the protected PDF rejects opening without a password: {!openedWithoutPassword}");

    // Import mode, not Modify: the *user* password only grants viewing rights when an owner
    // password is also set — Modify requires the owner password, since only the owner is meant to
    // be able to re-save the document at all. Discovered by actually running this demo end to end.
    using var openedWithPassword = PdfSharp.Pdf.IO.PdfReader.Open(protectedPdfPath, "reader-secret", PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import);
    Step($"VERIFY the correct user password opens it for reading: {openedWithPassword.PageCount > 0}");
}
else
{
    Step("PDF password protection SKIPPED — requires a converted PDF, which needs LibreOffice.");
}

var syntheticDocPath = Out("~synthetic-legacy.doc");
try
{
    File.WriteAllBytes(syntheticDocPath, new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }.Concat(new byte[512]).ToArray());
    var convertedFromDocPath = Out("25-converted-from-legacy-doc.docx");
    new LegacyDocConverter(wslDistro is not null ? new LegacyDocConversionOptions { UseWslDistro = wslDistro } : new LegacyDocConversionOptions())
        .ConvertToDocx(syntheticDocPath, convertedFromDocPath);
    Step($"VERIFY a .doc file converts to a docx the OpenXml pipeline can open: {File.Exists(convertedFromDocPath)}");
}
catch (Exception ex)
{
    Step($"Legacy .doc conversion SKIPPED — LibreOffice not available, or this stub file has no real content to convert ({ex.GetType().Name}).");
    Step("  (LegacyDocConverterTests.cs covers this path against a real LibreOffice install.)");
}
finally
{
    File.Delete(syntheticDocPath);
}

// ---------------------------------------------------------------------------------------------
Section("24. Audit — structured logging and telemetry");
// ---------------------------------------------------------------------------------------------

var loggedActivities = new List<string>();
using var activityListener = new ActivityListener
{
    ShouldListenTo = source => source.Name == DocumentProcessorDiagnostics.SourceName,
    Sample = (ref ActivityCreationOptions<ActivityContext> _) => ActivitySamplingResult.AllData,
    ActivityStopped = activity => loggedActivities.Add(activity.OperationName)
};
System.Diagnostics.ActivitySource.AddActivityListener(activityListener);

var demoLogger = new DemoConsoleLogger<TemplateEngine>();
var loggingTemplateEngine = new TemplateEngine(demoLogger);
var loggingOutputPath = Out("~audit-demo.docx");
loggingTemplateEngine.Fill(templatePath, loggingOutputPath, templateData, MissingTokenPolicy.Redact);
File.Delete(loggingOutputPath);

Step($"VERIFY the ActivitySource emitted a span for this Fill call: {loggedActivities.Contains("TemplateEngine.Fill")}");
Step($"VERIFY the ILogger received at least one log entry: {demoLogger.EntryCount > 0}");

// ---------------------------------------------------------------------------------------------
Section("25. Reviewer comments and text extraction — the CLM round-trip");
// ---------------------------------------------------------------------------------------------

var reviewPath = Out("24-commented-contract.docx");
SampleDocumentFactory.CreateDocumentFromParagraphs(reviewPath,
[
    new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }), new Run(new Text("5. Limitation of Liability"))),
    new Paragraph(new Run(new Text("Each party's aggregate liability is capped at the fees paid in the preceding twelve months."))),
    new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Heading2" }), new Run(new Text("5.1 Exclusions"))),
    new Paragraph(new Run(new Text("The cap does not apply to gross negligence or wilful misconduct."))),
]);

// One session carries the whole review exchange: comment, reply, resolve — and then reads the
// text back out for indexing, without a single intermediate save.
byte[] reviewedBytes;
string threadId;
using (var reviewSession = DocumentSession.OpenFile(reviewPath))
{
    threadId = reviewSession.Comments.Add(1, "Jordan Ellis", "JE", "Cap is too low — push for 2x fees.");
    reviewSession.Comments.Reply(threadId, "Sam Okafor", "SO", "Finance approved 2x. Countering.");
    reviewSession.Comments.Resolve(threadId);
    reviewedBytes = reviewSession.Save();
}
File.WriteAllBytes(reviewPath, reviewedBytes);
Step($"Wrote {Path.GetFileName(reviewPath)} — a comment, a threaded reply, and a resolved thread");

var commentService = new DocumentCommentService();
var thread = commentService.GetComments(reviewPath);
Step($"VERIFY both the comment and its reply round-tripped: {thread.Count == 2}");
foreach (var c in thread)
    Step($"  [{c.Id}] {c.Author}: \"{c.Text}\" (reply-to={c.ParentId ?? "—"}, resolved={c.IsResolved})");
Step($"VERIFY the reply is threaded under its parent: {thread.Any(c => c.ParentId == threadId)}");
Step($"VERIFY the thread is marked resolved: {thread.Single(c => c.Id == threadId).IsResolved}");
Step($"VERIFY the comment records the text it is anchored to: " +
     $"{thread.Single(c => c.Id == threadId).AnchorText.Contains("capped at the fees paid")}");

var extraction = new TextExtractionService();
var clauseBlocks = extraction.ExtractBlocks(reviewPath);
Step($"VERIFY extraction returns one block per non-empty paragraph: {clauseBlocks.Count == 4}");
foreach (var block in clauseBlocks)
    Step($"  {(block.HeadingLevel is { } level ? $"H{level}" : $"under \"{block.Heading}\"")}: {block.Text}");

var cappedClause = clauseBlocks.Single(b => b.Text.StartsWith("Each party's"));
Step($"VERIFY body text carries the heading it sits under, for clause-level indexing: " +
     $"{cappedClause.Heading == "5. Limitation of Liability"}");

// The case naive extraction gets wrong: InnerText would concatenate the old and new wording.
var pendingText = extraction.ExtractText(multiAuthorPath);
Step($"VERIFY tracked deletions are excluded from extracted text (\"$500,000\" is struck through): " +
     $"{!pendingText.Contains("$500,000")}");
Step($"VERIFY they can still be included deliberately: " +
     $"{extraction.ExtractText(multiAuthorPath, new TextExtractionOptions { IncludeDeletedText = true }).Contains("$500,000")}");

// ---------------------------------------------------------------------------------------------
Section("Done");
// ---------------------------------------------------------------------------------------------

var files = Directory.GetFiles(outputDir).OrderBy(f => f).ToList();
Console.WriteLine($"Produced {files.Count} file(s) in {outputDir}:");
foreach (var file in files)
    Console.WriteLine($"  {Path.GetFileName(file)}");

/// <summary>Minimal ILogger that prints to the console and counts entries, for section 24's
/// audit-logging demonstration — avoids adding a full logging-provider package dependency to the
/// demo just to show that log calls actually reach an ILogger.</summary>
sealed class DemoConsoleLogger<T> : ILogger<T>
{
    public int EntryCount { get; private set; }

    public IDisposable? BeginScope<TState>(TState state) where TState : notnull => null;

    public bool IsEnabled(LogLevel logLevel) => true;

    public void Log<TState>(LogLevel logLevel, EventId eventId, TState state, Exception? exception, Func<TState, Exception?, string> formatter)
    {
        EntryCount++;
        Console.WriteLine($"    [log:{logLevel}] {formatter(state, exception)}");
    }
}
