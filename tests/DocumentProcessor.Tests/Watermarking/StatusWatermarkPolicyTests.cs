using DocumentProcessor.Core.Samples;
using DocumentProcessor.Core.Watermarking;

namespace DocumentProcessor.Tests.Watermarking;

public class StatusWatermarkPolicyTests : IDisposable
{
    private readonly string _docxPath = TestFiles.NewTempPath(".docx");
    private readonly List<string> _cleanup = [];
    private readonly StatusWatermarkPolicy _sut = StatusWatermarkPolicy.CreateDefault();

    public StatusWatermarkPolicyTests()
    {
        SampleDocumentFactory.CreateBasicDocument(_docxPath, "Policy Test", ["Body text."]);
    }

    [Fact]
    public void Resolve_returns_the_mapped_config_for_a_known_status()
    {
        var config = _sut.Resolve("Draft");

        Assert.NotNull(config);
        Assert.Equal("DRAFT", config!.Text);
    }

    [Fact]
    public void Resolve_returns_null_for_Final_which_is_explicitly_mapped_to_no_watermark()
    {
        Assert.Null(_sut.Resolve("Final"));
    }

    [Fact]
    public void Resolve_returns_null_for_an_unmapped_status()
    {
        Assert.Null(_sut.Resolve("SomethingNotConfigured"));
    }

    [Fact]
    public void ApplyToDocx_for_Draft_adds_a_removable_watermark()
    {
        _sut.ApplyToDocx(_docxPath, "Draft");

        var docxWatermark = new DocxWatermarkService();
        Assert.True(docxWatermark.RemoveWatermark(_docxPath)); // only removable if it was actually added
    }

    [Fact]
    public void ApplyToDocx_for_Final_removes_any_existing_watermark()
    {
        _sut.ApplyToDocx(_docxPath, "Draft"); // add one first

        _sut.ApplyToDocx(_docxPath, "Final");

        var docxWatermark = new DocxWatermarkService();
        Assert.False(docxWatermark.RemoveWatermark(_docxPath)); // nothing left to remove
    }

    [Fact]
    public void ApplyToPdf_for_Confidential_produces_a_watermarked_output_file()
    {
        var pdfPath = TestFiles.NewSimplePdf(1, "Contract");
        var outputPath = TestFiles.NewTempPath(".pdf");
        _cleanup.Add(pdfPath);
        _cleanup.Add(outputPath);

        _sut.ApplyToPdf(pdfPath, outputPath, "Confidential");

        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void ApplyToPdf_for_Final_copies_the_input_through_unwatermarked()
    {
        var pdfPath = TestFiles.NewSimplePdf(1, "Contract");
        var outputPath = TestFiles.NewTempPath(".pdf");
        _cleanup.Add(pdfPath);
        _cleanup.Add(outputPath);

        _sut.ApplyToPdf(pdfPath, outputPath, "Final");

        Assert.Equal(File.ReadAllBytes(pdfPath), File.ReadAllBytes(outputPath));
    }

    public void Dispose()
    {
        File.Delete(_docxPath);
        foreach (var path in _cleanup)
            File.Delete(path);
    }
}
