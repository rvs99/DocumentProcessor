using DocumentFormat.OpenXml.Packaging;
using DocumentProcessor.Core.Format;

namespace DocumentProcessor.Tests.Format;

public class LegacyDocConverterTests : IDisposable
{
    private readonly string _docPath = TestFiles.NewTempPath(".doc");
    private readonly string _docxPath = TestFiles.NewTempPath(".docx");
    private readonly LegacyDocConverter _sut = new(ConversionOptions());

    public LegacyDocConverterTests()
    {
        // The OLE Compound File Binary Format signature every legacy .doc file starts with — not a
        // complete, renderable document (this repo has no Word install to generate a real one from),
        // but a real .doc file for soffice to attempt the conversion against.
        var oleSignature = new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
        File.WriteAllBytes(_docPath, oleSignature.Concat(new byte[512]).ToArray());
    }

    private static LegacyDocConversionOptions ConversionOptions() =>
        Environment.GetEnvironmentVariable("DOCPROC_LIBREOFFICE_WSL_DISTRO") is { } distro
            ? new LegacyDocConversionOptions { UseWslDistro = distro }
            : new LegacyDocConversionOptions();

    [Fact]
    public void ConvertToDocx_missing_input_throws_FileNotFoundException()
    {
        Assert.Throws<FileNotFoundException>(() =>
            _sut.ConvertToDocx(TestFiles.NewTempPath(".doc"), _docxPath));
    }

    [Fact]
    public void ConvertToDocx_produces_a_docx_the_OpenXml_pipeline_can_open()
    {
        _sut.ConvertToDocx(_docPath, _docxPath);

        Assert.True(File.Exists(_docxPath));
        using var doc = WordprocessingDocument.Open(_docxPath, isEditable: false);
        Assert.NotNull(doc.MainDocumentPart?.Document?.Body);
    }

    public void Dispose()
    {
        File.Delete(_docPath);
        if (File.Exists(_docxPath))
            File.Delete(_docxPath);
    }
}
