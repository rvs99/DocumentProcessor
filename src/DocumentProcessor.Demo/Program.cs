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

Console.WriteLine("DocumentProcessor demo — running full document lifecycle.");
Console.WriteLine($"Output directory: {outputDir}");

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
        ["Training", "5", "$5,000", "$25,000"]
    ],
    Caption: "Schedule A — Pricing"));
Step("Appended a 4-column pricing table with caption \"Schedule A — Pricing\"");

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
    "Governing Law: This agreement is governed by and construed in accordance with the laws of " +
        "the State of Delaware, without regard to its conflict of laws principles.",
    "Force Majeure: Neither party shall be liable for delays caused by circumstances beyond its " +
        "reasonable control."
]);

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
Section("6. Mark the draft with a watermark — docx watermarking");
// ---------------------------------------------------------------------------------------------

var watermarkedDraftPath = Out("02-contract-draft-watermarked.docx");
File.Copy(contractPath, watermarkedDraftPath, overwrite: true);
new DocxWatermarkService().AddTextWatermark(watermarkedDraftPath, "DRAFT");
Step($"Wrote {Path.GetFileName(watermarkedDraftPath)} with a diagonal \"DRAFT\" watermark on every page");

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

var wslDistro = Environment.GetEnvironmentVariable("DOCPROC_LIBREOFFICE_WSL_DISTRO");
var converterOptions = wslDistro is not null
    ? new WordToPdfConversionOptions { UseWslDistro = wslDistro }
    : new WordToPdfConversionOptions();
var converter = new WordToPdfConverter(converterOptions);

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
