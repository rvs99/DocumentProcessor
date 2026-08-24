using PdfSharp.Pdf.IO;

namespace DocumentProcessor.Core.Security;

/// <summary>Permission flags for a password-protected PDF — defaults match a reasonable "shared for
/// review, not for editing" baseline: printable and annotatable, but not modifiable/extractable/
/// reassemblable.</summary>
public sealed record PdfPermissions(
    bool AllowPrinting = true,
    bool AllowFullQualityPrint = true,
    bool AllowModifyDocument = false,
    bool AllowExtractContent = false,
    bool AllowAnnotations = true,
    bool AllowFormFilling = true,
    bool AllowAssembleDocument = false);

/// <summary>
/// Password-protects a PDF via PDFsharp's own <c>PdfSecuritySettings</c> — capability that already
/// existed in the dependency this library uses for watermarking/e-sign, just never wired up. An
/// owner password restricts permissions (printing/editing/etc.) while the document still opens
/// without a prompt; a user password requires that password just to open the document at all.
/// </summary>
public sealed class PdfProtectionService : IPdfProtectionService
{
    public void ProtectPdf(string pdfPath, string outputPath, string? ownerPassword = null, string? userPassword = null, PdfPermissions? permissions = null)
    {
        if (ownerPassword is null && userPassword is null)
            throw new ArgumentException("Supply at least one of ownerPassword or userPassword.", nameof(ownerPassword));

        using var document = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Modify);
        var security = document.SecuritySettings;

        if (ownerPassword is not null)
            security.OwnerPassword = ownerPassword;
        if (userPassword is not null)
            security.UserPassword = userPassword;

        var perms = permissions ?? new PdfPermissions();
        security.PermitPrint = perms.AllowPrinting;
        security.PermitFullQualityPrint = perms.AllowFullQualityPrint;
        security.PermitModifyDocument = perms.AllowModifyDocument;
        security.PermitExtractContent = perms.AllowExtractContent;
        security.PermitAnnotations = perms.AllowAnnotations;
        security.PermitFormsFill = perms.AllowFormFilling;
        security.PermitAssembleDocument = perms.AllowAssembleDocument;

        document.Save(outputPath);
    }
}
