using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace DocumentProcessor.Core.Security;

public sealed record TemplateValidationResult(bool IsValid, IReadOnlyList<string> Issues);

/// <summary>
/// Detects and strips VBA macros from a .docx, and does a basic structural sanity check before a
/// document is used as a fill-in template — e.g. reject a macro-enabled file uploaded as a contract
/// template, since executing arbitrary VBA was never part of that workflow's threat model.
/// </summary>
public sealed class MacroValidationService
{
    public bool ContainsMacros(string docxPath)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: false);
        return ContainsMacros(doc);
    }

    /// <summary>Removes the VBA project (if any) and, for a macro-enabled document/template,
    /// switches its content type to the plain equivalent — the same effect as Word's own "Save As"
    /// to the non-macro-enabled format, minus the UI prompt.</summary>
    public void StripMacros(string docxPath, string outputPath)
    {
        File.Copy(docxPath, outputPath, overwrite: true);

        using var doc = WordprocessingDocument.Open(outputPath, isEditable: true);
        var mainPart = doc.MainDocumentPart;
        if (mainPart?.VbaProjectPart is { } vbaProjectPart)
            mainPart.DeletePart(vbaProjectPart);

        if (doc.DocumentType == WordprocessingDocumentType.MacroEnabledDocument)
            doc.ChangeDocumentType(WordprocessingDocumentType.Document);
        else if (doc.DocumentType == WordprocessingDocumentType.MacroEnabledTemplate)
            doc.ChangeDocumentType(WordprocessingDocumentType.Template);
    }

    /// <summary>
    /// A minimal "is this safe/sane to use as a template" check: the file must open as a valid OOXML
    /// package with a main document body, and must not carry macros. Doesn't validate template-
    /// specific content (e.g. token syntax) — pair with <see cref="Templating.TemplateEngine"/>'s own
    /// error reporting for that.
    /// </summary>
    public TemplateValidationResult ValidateTemplate(string docxPath)
    {
        if (!File.Exists(docxPath))
            return new TemplateValidationResult(false, ["File does not exist."]);

        try
        {
            using var doc = WordprocessingDocument.Open(docxPath, isEditable: false);
            var issues = new List<string>();

            if (doc.MainDocumentPart?.Document?.Body is null)
                issues.Add("Document has no main body content.");
            if (ContainsMacros(doc))
                issues.Add("Document contains VBA macros, which are not permitted in a content template.");

            return new TemplateValidationResult(issues.Count == 0, issues);
        }
        catch (Exception ex)
        {
            // Deliberately broad: this method's entire contract is "tell me why this file isn't
            // usable" rather than throw, so any failure to open it — corrupt zip, wrong format,
            // missing required parts — is itself the validation result, not an exceptional case.
            return new TemplateValidationResult(false, [$"Not a valid .docx package: {ex.Message}"]);
        }
    }

    private static bool ContainsMacros(WordprocessingDocument doc) =>
        doc.MainDocumentPart?.VbaProjectPart is not null ||
        doc.DocumentType is WordprocessingDocumentType.MacroEnabledDocument or WordprocessingDocumentType.MacroEnabledTemplate;
}
