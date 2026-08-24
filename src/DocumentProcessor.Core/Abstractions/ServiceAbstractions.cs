// Interfaces for the library's services, so a consuming application can register them in a DI
// container and fake them in its own unit tests. Without these, every service was a sealed class
// with no abstraction: Moq and NSubstitute could not stand in for any of them, which meant a
// consumer testing their own contract-assembly logic had to hit the real filesystem, and anything
// touching conversion forced a real LibreOffice install into their CI just to exercise their
// business rules.
//
// Each interface mirrors its service exactly. Generated from the compiled assembly (including
// nullable annotations) rather than transcribed, then reviewed; the compiler enforces that they
// stay in step, since each service declares it implements its own interface.

using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.Comments;
using DocumentProcessor.Core.Comparison;
using DocumentProcessor.Core.Extraction;
using DocumentProcessor.Core.ContentControls;
using DocumentProcessor.Core.Conversion;
using DocumentProcessor.Core.DocumentAssembly;
using DocumentProcessor.Core.FontEmbedding;
using DocumentProcessor.Core.Layout;
using DocumentProcessor.Core.Metadata;
using DocumentProcessor.Core.Redlining;
using DocumentProcessor.Core.Security;
using DocumentProcessor.Core.Tables;
using DocumentProcessor.Core.Templating;
using DocumentProcessor.Core.TrackChanges;
using DocumentProcessor.Core.Transplant;
using DocumentProcessor.Core.Watermarking;
using PageSize = DocumentProcessor.Core.Layout.PageSize;

namespace DocumentProcessor.Core.Comparison
{
    public interface IPdfComparisonService
    {
        PdfTextDiffResult CompareText(string pdfPathA, string pdfPathB);
        PdfVisualDiffResult CompareVisual(string pdfPathA, string pdfPathB, double differenceThresholdPercent = 0.5, int dpi = 150);
    }
}

namespace DocumentProcessor.Core.ContentControls
{
    public interface IContentControlService
    {
        IReadOnlyList<ContentControlInfo> ListContentControls(string docxPath);
        int ReplaceByTag(string docxPath, string tag, string newValue);
        (byte[] Document, int UpdatedCount) ReplaceByTag(byte[] docxBytes, string tag, string newValue);
        IReadOnlyDictionary<string, int> ReplaceMany(string docxPath, IReadOnlyDictionary<string, string> tagToValue);
        int SetContentDateByTag(string docxPath, string tag, DateTime value, string? displayFormat = null);
        int SetContentDropDownSelectionByTag(string docxPath, string tag, string value);
        int SetContentRichTextByTag(string docxPath, string tag, string html);
        int SetLock(string docxPath, string tag, ContentControlLockMode mode);
    }
}

namespace DocumentProcessor.Core.Conversion
{
    public interface IWordToPdfConverter
    {
        void Convert(string docxPath, string outputPdfPath);
        Task ConvertAsync(string docxPath, string outputPdfPath, CancellationToken cancellationToken = default);
        Task<IReadOnlyList<ConversionResult>> ConvertBatchAsync(IReadOnlyList<ConversionRequest> requests, CancellationToken cancellationToken = default);
    }
}

namespace DocumentProcessor.Core.DocumentAssembly
{
    public interface IPdfAssemblyService
    {
        void AppendWithContinuedPageNumbers(string mainPdfPath, string exhibitPdfPath, string outputPath, int startingPageNumber = 1, Func<int, string>? formatPageNumber = null, double marginBottomPt = 24);
        void ExtractPages(string pdfPath, int startPageIndex, int endPageIndex, string outputPath);
        void MergePdfs(IReadOnlyList<string> pdfPaths, string outputPath, CancellationToken cancellationToken = default);
    }

    public interface ICrossReferenceValidator
    {
        IReadOnlyList<DanglingReference> Validate(string docxPath);
    }

    public interface IFieldUpdateService
    {
        void MarkAllFieldsDirty(string docxPath);
        void SetUpdateFieldsOnOpen(string docxPath, bool updateOnOpen = true);
    }
}

namespace DocumentProcessor.Core.ESign
{
    public interface IESignFieldService
    {
        void InjectDocxAnchor(string docxPath, string anchorText, string tag = "ESignatureField");
        void InjectPdfAnchor(string pdfPath, string outputPath, string anchorText, int pageIndex, double x, double y, bool invisible = false);
    }
}

namespace DocumentProcessor.Core.FontEmbedding
{
    public interface IFontEmbeddingService
    {
        void ApplyFontToAllRuns(string docxPath, string fontFamilyName);
        void EmbedFontFamily(string docxPath, string fontFamilyName, FontFamilyFiles files);
        IReadOnlyList<string> ListEmbeddedFonts(string docxPath);
    }
}

namespace DocumentProcessor.Core.Format
{
    public interface ILegacyDocConverter
    {
        void ConvertToDocx(string docPath, string outputDocxPath);
        Task ConvertToDocxAsync(string docPath, string outputDocxPath, CancellationToken cancellationToken = default);
    }
}

namespace DocumentProcessor.Core.Layout
{
    public interface IHeaderFooterService
    {
        bool RemoveFooter(string docxPath, HeaderFooterValues? type = null);
        bool RemoveHeader(string docxPath, HeaderFooterValues? type = null);
        void SetFooterContent(string docxPath, IReadOnlyList<HeaderFooterPart> parts, IReadOnlyDictionary<string, object?>? data = null, HeaderFooterValues? type = null);
        void SetFooterText(string docxPath, string text, HeaderFooterValues? type = null);
        void SetHeaderContent(string docxPath, IReadOnlyList<HeaderFooterPart> parts, IReadOnlyDictionary<string, object?>? data = null, HeaderFooterValues? type = null);
        void SetHeaderText(string docxPath, string text, HeaderFooterValues? type = null);
    }

    public interface IPageLayoutService
    {
        void InsertPageBreak(string docxPath, int beforeParagraphIndex);
        void InsertSectionBreak(string docxPath, int beforeParagraphIndex, SectionMarkValues? breakType = null);
        void SetColumns(string docxPath, int columnCount, int spacingTwips = 720, int? sectionIndex = null);
        void SetDefaultParagraphSpacing(string docxPath, int afterTwips, int lineTwips, LineSpacingRuleValues lineRule);
        void SetMargins(string docxPath, PageMargins margins, int? sectionIndex = null);
        void SetPageSize(string docxPath, PageSize size, int? sectionIndex = null);
    }

    public interface IBrandingService
    {
        void ApplyBranding(string docxPath, TenantBrandingSpec branding);
    }
}

namespace DocumentProcessor.Core.Metadata
{
    public interface IDocumentMetadataService
    {
        IReadOnlyDictionary<string, string> GetCustomProperties(string docxPath);
        bool RemoveCustomProperty(string docxPath, string name);
        void SetCoreProperties(string docxPath, string? title = null, string? author = null, string? subject = null, string? keywords = null);
        void SetCustomProperties(string docxPath, IReadOnlyDictionary<string, object?> properties);
        void SetCustomProperty(string docxPath, string name, string value);
        void SetCustomProperty(string docxPath, string name, bool value);
        void SetCustomProperty(string docxPath, string name, int value);
        void SetCustomProperty(string docxPath, string name, double value);
        void SetCustomProperty(string docxPath, string name, DateTime value);
    }

    public interface IPdfMetadataService
    {
        void SetXmpMetadata(string pdfPath, string outputPath, XmpMetadata metadata);
    }
}

namespace DocumentProcessor.Core.Redlining
{
    public interface IDocumentComparisonService
    {
        ChangeSummary Compare(string originalPath, string revisedPath, string outputRedlinedPath, string authorForRevisions = "Document Comparison");
        ComparisonSummary CompareDetailed(string originalPath, string revisedPath, string outputRedlinedPath, string authorForRevisions = "Document Comparison");
    }

    public interface IDocumentProtectionService
    {
        void AllowEditingInRange(string docxPath, int startParagraphIndex, int endParagraphIndex, EditorGroup editorGroup = EditorGroup.Everyone);
        void RemoveDocumentProtection(string docxPath);
        void SetDocumentProtection(string docxPath, EditRestriction restriction, string? password = null);
    }

    public interface IRedlineExportService
    {
        RedlineExportPaths ExportAllVariants(string originalPath, string revisedPath, string outputDirectory, string authorForRevisions = "Document Comparison", WordToPdfConversionOptions? conversionOptions = null);
    }
}

namespace DocumentProcessor.Core.Security
{
    public interface IMacroValidationService
    {
        bool ContainsMacros(string docxPath);
        void StripMacros(string docxPath, string outputPath);
        TemplateValidationResult ValidateTemplate(string docxPath);
    }

    public interface IPdfProtectionService
    {
        void ProtectPdf(string pdfPath, string outputPath, string? ownerPassword = null, string? userPassword = null, PdfPermissions? permissions = null);
    }
}

namespace DocumentProcessor.Core.Tables
{
    public interface ITableGenerationService
    {
        void AppendTable(string docxPath, TableSpec spec);
        int PopulateFromPrototypeRow(string docxPath, int tableIndex, IReadOnlyList<IReadOnlyDictionary<string, object?>> rows, MissingTokenPolicy missingTokenPolicy = MissingTokenPolicy.Error, CancellationToken cancellationToken = default);
        void ReplaceTable(string docxPath, int tableIndex, TableSpec spec);
    }
}

namespace DocumentProcessor.Core.Templating
{
    public interface ITemplateEngine
    {
        TemplateFillResult Fill(string templatePath, string outputPath, IReadOnlyDictionary<string, object?> data, MissingTokenPolicy missingTokenPolicy = MissingTokenPolicy.Error, ClauseLibrary? clauseLibrary = null, CancellationToken cancellationToken = default);
    }
}

namespace DocumentProcessor.Core.TrackChanges
{
    public interface ITrackChangesService
    {
        void AcceptAll(string docxPath);
        void AcceptByAuthor(string docxPath, string author);
        void AcceptById(string docxPath, string changeId);
        IReadOnlyList<TrackedChange> GetTrackedChanges(string docxPath);
        bool HasTrackedChanges(string docxPath);
        void RejectAll(string docxPath);
        void RejectByAuthor(string docxPath, string author);
        void RejectById(string docxPath, string changeId);
    }
}

namespace DocumentProcessor.Core.Transplant
{
    public interface IClauseTransplantService
    {
        void ContinueHeadingNumbering(string docxPath, int insertedStartIndex, int insertedCount);
        IReadOnlyList<ParagraphInfo> ListParagraphs(string docxPath);
        void RemoveParagraphs(string docxPath, int startIndex, int count, string outputPath);
        IReadOnlyList<DanglingReference> RemoveParagraphsWithCrossReferenceCleanup(string docxPath, int startIndex, int count, string outputPath);
        void ReplaceParagraphs(string sourcePath, int sourceStartIndex, int sourceParagraphCount, string targetPath, int replacedStartIndex, int replacedCount, string outputPath);
        void TransplantParagraphs(string sourcePath, int sourceStartIndex, int paragraphCount, string targetPath, int insertBeforeParagraphIndex, string outputPath);
    }
}

namespace DocumentProcessor.Core.Watermarking
{
    public interface IDocxWatermarkService
    {
        void AddTextWatermark(string docxPath, string text, string fontFamily = "Calibri", int rotationDegrees = -45, string colorHex = "C0C0C0", bool removable = true, WatermarkPosition position = WatermarkPosition.Center, double widthPt = 415, double heightPt = 207.5, double fontSizePt = 72);
        bool RemoveWatermark(string docxPath);
    }

    public interface IPdfWatermarkService
    {
        void AddTextWatermark(string pdfPath, string outputPath, string text, string fontFamily = "Arial", double rotationDegrees = -45, byte grayLevel = 192, byte alpha = 100, WatermarkPosition position = WatermarkPosition.Center, double fontSizePt = 72);
    }
}

namespace DocumentProcessor.Core.Comments
{
    public interface IDocumentCommentService
    {
        IReadOnlyList<DocumentComment> GetComments(string docxPath);
        string AddComment(string docxPath, int paragraphIndex, string author, string initials, string text);
        string ReplyToComment(string docxPath, string parentCommentId, string author, string initials, string text);
        void ResolveComment(string docxPath, string commentId, bool resolved = true);
        bool DeleteComment(string docxPath, string commentId);
    }
}

namespace DocumentProcessor.Core.Extraction
{
    public interface ITextExtractionService
    {
        string ExtractText(string docxPath, TextExtractionOptions? options = null);
        IReadOnlyList<TextBlock> ExtractBlocks(string docxPath, TextExtractionOptions? options = null);
    }
}
