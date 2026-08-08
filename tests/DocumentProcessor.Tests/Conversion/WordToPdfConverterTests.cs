using DocumentProcessor.Core.Conversion;
using DocumentProcessor.Core.Samples;

namespace DocumentProcessor.Tests.Conversion;

/// <summary>
/// Exercises the real LibreOffice conversion path. Uses a native soffice binary by default (what
/// CI/Linux containers use); set DOCPROC_LIBREOFFICE_WSL_DISTRO=Ubuntu to route through WSL for
/// local Windows development instead — see TestFiles.ConversionOptions.
/// </summary>
public class WordToPdfConverterTests : IDisposable
{
    private readonly string _docxPath = TestFiles.NewTempPath(".docx");
    private readonly string _pdfPath = TestFiles.NewTempPath(".pdf");
    private readonly WordToPdfConverter _sut = new(TestFiles.ConversionOptions());

    public WordToPdfConverterTests()
    {
        SampleDocumentFactory.CreateBasicDocument(_docxPath, "Conversion Test",
            ["This paragraph should survive the round trip to PDF.", "So should this second one."]);
    }

    [Fact]
    public async Task ConvertAsync_produces_a_non_empty_pdf_file()
    {
        await _sut.ConvertAsync(_docxPath, _pdfPath);

        Assert.True(File.Exists(_pdfPath));
        var bytes = await File.ReadAllBytesAsync(_pdfPath);
        Assert.True(bytes.Length > 500, "Produced PDF is suspiciously small.");
        Assert.Equal("%PDF"u8.ToArray(), bytes[..4]);
    }

    [Fact]
    public async Task ConvertAsync_throws_for_a_missing_input_file()
    {
        await Assert.ThrowsAsync<FileNotFoundException>(() =>
            _sut.ConvertAsync(TestFiles.NewTempPath(".docx"), _pdfPath));
    }

    public void Dispose()
    {
        File.Delete(_docxPath);
        if (File.Exists(_pdfPath))
            File.Delete(_pdfPath);
    }
}
