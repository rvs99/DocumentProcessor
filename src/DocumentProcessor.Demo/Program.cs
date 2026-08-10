using DocumentFormat.OpenXml.Packaging;
using DocumentProcessor.Core.Comparison;
using DocumentProcessor.Core.Conversion;
using DocumentProcessor.Core.ContentControls;
using DocumentProcessor.Core.ESign;
using DocumentProcessor.Core.FontEmbedding;
using DocumentProcessor.Core.Redlining;
using DocumentProcessor.Core.Samples;
using DocumentProcessor.Core.Tables;
using DocumentProcessor.Core.TrackChanges;
using DocumentProcessor.Core.Transplant;
using DocumentProcessor.Core.Watermarking;
using UglyToad.PdfPig;

// Walks through a single contract's lifecycle end to end, exercising every capability of the
// module in the order a real document-processing pipeline would use them: draft → populate →
// assemble → brand → negotiate → finalize → convert → distribute → verify.

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

Console.WriteLine("DocumentProcessor demo — running full document lifecycle.");
Console.WriteLine($"Output directory: {outputDir}");

// Set up once, used both for the LibreOffice-rendering check in step 6 and the real conversion in
// step 11. Set DOCPROC_LIBREOFFICE_WSL_DISTRO=Ubuntu to route through WSL for local Windows dev
// (see Properties/launchSettings.json) instead of a native soffice binary.
var wslDistro = Environment.GetEnvironmentVariable("DOCPROC_LIBREOFFICE_WSL_DISTRO");
var converterOptions = wslDistro is not null
    ? new WordToPdfConversionOptions { UseWslDistro = wslDistro }
    : new WordToPdfConversionOptions();
var converter = new WordToPdfConverter(converterOptions);

// ---------------------------------------------------------------------------------------------
Section("1. Draft the contract, with content controls as fill-in fields");
// ---------------------------------------------------------------------------------------------

var contractPath = Out("01-contract-draft.docx");
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

// Snapshot the draft here, on the plain default font (Calibri, via SampleDocumentFactory's
// docDefaults), and convert it to PDF — every other PDF this demo produces goes through the
// Roboto Mono override applied in the next step, so this is the one example of a normal-font
// docx -> PDF conversion, useful for eyeballing general conversion/pagination fidelity without
// the custom monospace font as a variable.
var normalFontDraftPath = Out("01b-contract-draft-normal-font.docx");
File.Copy(contractPath, normalFontDraftPath, overwrite: true);
try
{
    var normalFontPdfPath = Out("01b-contract-draft-normal-font.pdf");
    await converter.ConvertAsync(normalFontDraftPath, normalFontPdfPath);
    Step($"Wrote {Path.GetFileName(normalFontPdfPath)} — same content, default font (Calibri), for comparison against the Roboto Mono version below");
}
catch (Exception ex)
{
    Step("SKIPPED normal-font PDF conversion — LibreOffice not available in this environment.");
    Step($"  ({ex.GetType().Name}: {ex.Message})");
}

// ---------------------------------------------------------------------------------------------
Section("4. Embed a custom font family — custom/embedded font support");
// ---------------------------------------------------------------------------------------------

var fontService = new FontEmbeddingService();
var fontFile = Path.Combine(AppContext.BaseDirectory, "Assets", "Fonts", "RobotoMono-Regular.ttf");
fontService.EmbedFontFamily(contractPath, "Roboto Mono", new FontFamilyFiles(fontFile));
fontService.ApplyFontToAllRuns(contractPath, "Roboto Mono");
Step("Embedded Roboto Mono (SIL OFL) directly into the .docx and applied it to all runs");
Step("Document now renders correctly even on machines without this font installed");

// ---------------------------------------------------------------------------------------------
Section("5. Pull the governing-law clause in from the firm's clause library — clause transplant");
// ---------------------------------------------------------------------------------------------

var clauseLibraryPath = Out("clause-library.docx");
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
Section("6. Mark the draft with a watermark — docx watermarking (removable and locked modes)");
// ---------------------------------------------------------------------------------------------

var docxWatermarkService = new DocxWatermarkService();

// removable: true (the default) matches Word's own naming convention, so end users can clear it
// themselves via Design -> Watermark -> Remove Watermark once the document is finalized.
var watermarkedDraftPath = Out("02-contract-draft-watermarked-removable.docx");
File.Copy(contractPath, watermarkedDraftPath, overwrite: true);
docxWatermarkService.AddTextWatermark(watermarkedDraftPath, "DRAFT", removable: true);
Step($"Wrote {Path.GetFileName(watermarkedDraftPath)} — \"DRAFT\" watermark, removable via Word's own UI");

// removable: false uses a shape id Word's Watermark UI doesn't recognize, so Remove Watermark
// can't touch it — appropriate for a disclaimer that shouldn't be a click away from disappearing.
var lockedWatermarkPath = Out("02b-contract-draft-watermarked-locked.docx");
File.Copy(contractPath, lockedWatermarkPath, overwrite: true);
docxWatermarkService.AddTextWatermark(lockedWatermarkPath, "DRAFT", removable: false);
Step($"Wrote {Path.GetFileName(lockedWatermarkPath)} — locked watermark, Word's Remove Watermark can't clear it");

// Converting a watermarked docx straight through LibreOffice would carry the watermark over as
// real, selectable PDF text (v:textbox is genuine WordprocessingML content) — fine for a docx
// that only ever opens in Word, but not what a distributed PDF should do. The production pattern
// is strip -> convert -> reapply: remove the watermark before conversion, convert the clean docx,
// then apply PdfWatermarkService (which rasterizes) to the result. See WatermarkPipelineTests.cs.
async Task<string> ConvertWithNonSelectableWatermark(string watermarkedDocxPath, string watermarkText, string outputBaseName)
{
    var strippedCopyPath = Out($"~stripped-{outputBaseName}.docx");
    File.Copy(watermarkedDocxPath, strippedCopyPath, overwrite: true);
    docxWatermarkService.RemoveWatermark(strippedCopyPath);

    var cleanPdfPath = Out($"~clean-{outputBaseName}.pdf");
    await converter.ConvertAsync(strippedCopyPath, cleanPdfPath);

    var finalPdfPath = Out($"{outputBaseName}.pdf");
    new PdfWatermarkService().AddTextWatermark(cleanPdfPath, finalPdfPath, watermarkText);

    File.Delete(strippedCopyPath);
    File.Delete(cleanPdfPath);
    return finalPdfPath;
}

try
{
    var removablePdfPath = await ConvertWithNonSelectableWatermark(watermarkedDraftPath, "DRAFT", "02-contract-draft-watermarked-removable");
    var lockedPdfPath = await ConvertWithNonSelectableWatermark(lockedWatermarkPath, "DRAFT", "02b-contract-draft-watermarked-locked");

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
Section("7. Negotiation: counterparty proposes changes — simulated via tracked changes");
// ---------------------------------------------------------------------------------------------

var negotiatedPath = Out("03-contract-negotiated.docx");
File.Copy(contractPath, negotiatedPath, overwrite: true);
contentControlService.ReplaceByTag(negotiatedPath, "ContractValue", "$275,000");
Step("Counterparty countered on contract value: $250,000 -> $275,000");

// ---------------------------------------------------------------------------------------------
Section("8. Compare the draft against the negotiated version — redlining / docx comparison");
// ---------------------------------------------------------------------------------------------

var redlinedPath = Out("04-contract-redlined.docx");
var comparisonService = new DocumentComparisonService();
var changeSummary = comparisonService.Compare(contractPath, negotiatedPath, redlinedPath, authorForRevisions: "Counterparty Counsel");
Step($"Wrote {Path.GetFileName(redlinedPath)} — {changeSummary.InsertedCount} insertion(s), {changeSummary.DeletedCount} deletion(s)");
foreach (var text in changeSummary.InsertedText.Take(3))
    Step($"  inserted: \"{text}\"");
foreach (var text in changeSummary.DeletedText.Take(3))
    Step($"  deleted:  \"{text}\"");

// ---------------------------------------------------------------------------------------------
Section("9. Legal accepts the change; produce the final accepted version — track changes accept/reject");
// ---------------------------------------------------------------------------------------------

var trackChangesService = new TrackChangesService();

var acceptedPath = Out("05-contract-accepted.docx");
File.Copy(redlinedPath, acceptedPath, overwrite: true);
trackChangesService.AcceptAll(acceptedPath);
Step($"Wrote {Path.GetFileName(acceptedPath)} — all changes accepted (final contract value applies)");

var rejectedPath = Out("06-contract-rejected.docx");
File.Copy(redlinedPath, rejectedPath, overwrite: true);
trackChangesService.RejectAll(rejectedPath);
Step($"Wrote {Path.GetFileName(rejectedPath)} — all changes rejected, for comparison (original value restored)");

// ---------------------------------------------------------------------------------------------
Section("10. Add an e-signature anchor to the final docx — e-sign field injection");
// ---------------------------------------------------------------------------------------------

var esignService = new ESignFieldService();
esignService.InjectDocxAnchor(acceptedPath, anchorText: "/sig1/", tag: "ClientSignature");
Step("Injected a \"/sig1/\" anchor tag — the convention DocuSign/Adobe Sign auto-detect for placement");

// ---------------------------------------------------------------------------------------------
Section("11. Convert the final contract to PDF for distribution — docx -> PDF conversion");
// ---------------------------------------------------------------------------------------------

var finalPdfPath = Out("07-contract-final.pdf");
var negotiatedPdfPath = Out("08-contract-negotiated-for-comparison.pdf");
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
    Section("12. Watermark and tag the PDF for distribution — PDF watermarking + e-sign field injection");
    // -----------------------------------------------------------------------------------------

    var watermarkedPdfPath = Out("09-contract-final-watermarked.pdf");
    new PdfWatermarkService().AddTextWatermark(finalPdfPath, watermarkedPdfPath, "FINAL");
    Step($"Wrote {Path.GetFileName(watermarkedPdfPath)} with a \"FINAL\" watermark");

    var signablePdfPath = Out("10-contract-final-signable.pdf");
    esignService.InjectPdfAnchor(watermarkedPdfPath, signablePdfPath, "/sig1/", pageIndex: 0, x: 50, y: 700);
    Step($"Wrote {Path.GetFileName(signablePdfPath)} with a \"/sig1/\" e-signature anchor");

    // -----------------------------------------------------------------------------------------
    Section("13. Verify the final PDF matches the negotiated terms — PDF comparison");
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
Section("Done");
// ---------------------------------------------------------------------------------------------

var files = Directory.GetFiles(outputDir).OrderBy(f => f).ToList();
Console.WriteLine($"Produced {files.Count} file(s) in {outputDir}:");
foreach (var file in files)
    Console.WriteLine($"  {Path.GetFileName(file)}");
