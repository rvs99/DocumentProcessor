using DocumentProcessor.Core.Security;
using PdfSharp.Pdf.IO;

namespace DocumentProcessor.Tests.Security;

public class PdfProtectionServiceTests : IDisposable
{
    private readonly string _pdfPath;
    private readonly string _outputPath = TestFiles.NewTempPath(".pdf");
    private readonly PdfProtectionService _sut = new();

    public PdfProtectionServiceTests()
    {
        _pdfPath = TestFiles.NewSimplePdf(1, "Contract");
    }

    [Fact]
    public void ProtectPdf_with_a_user_password_rejects_opening_without_it()
    {
        _sut.ProtectPdf(_pdfPath, _outputPath, userPassword: "secret123");

        Assert.Throws<PdfSharp.Pdf.IO.PdfReaderException>(() =>
            PdfReader.Open(_outputPath, PdfDocumentOpenMode.Modify));
    }

    [Fact]
    public void ProtectPdf_with_a_user_password_opens_successfully_with_the_correct_password()
    {
        _sut.ProtectPdf(_pdfPath, _outputPath, userPassword: "secret123");

        using var doc = PdfReader.Open(_outputPath, "secret123", PdfDocumentOpenMode.Modify);
        Assert.Equal(1, doc.PageCount);
    }

    [Fact]
    public void ProtectPdf_without_any_password_throws()
    {
        Assert.Throws<ArgumentException>(() => _sut.ProtectPdf(_pdfPath, _outputPath));
    }

    [Fact]
    public void ProtectPdf_applies_the_requested_permission_flags()
    {
        _sut.ProtectPdf(_pdfPath, _outputPath, ownerPassword: "owner-secret",
            permissions: new PdfPermissions(AllowPrinting: false, AllowModifyDocument: false, AllowExtractContent: false));

        using var doc = PdfReader.Open(_outputPath, "owner-secret", PdfDocumentOpenMode.Modify);
        Assert.False(doc.SecuritySettings.PermitPrint);
        Assert.False(doc.SecuritySettings.PermitModifyDocument);
        Assert.False(doc.SecuritySettings.PermitExtractContent);
    }

    [Fact]
    public void ProtectPdf_defaults_still_allow_printing_and_annotations()
    {
        _sut.ProtectPdf(_pdfPath, _outputPath, ownerPassword: "owner-secret");

        using var doc = PdfReader.Open(_outputPath, "owner-secret", PdfDocumentOpenMode.Modify);
        Assert.True(doc.SecuritySettings.PermitPrint);
        Assert.True(doc.SecuritySettings.PermitAnnotations);
    }

    public void Dispose()
    {
        File.Delete(_pdfPath);
        if (File.Exists(_outputPath))
            File.Delete(_outputPath);
    }
}
