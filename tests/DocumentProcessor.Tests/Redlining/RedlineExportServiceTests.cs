using DocumentProcessor.Core.Redlining;
using DocumentProcessor.Core.Samples;

namespace DocumentProcessor.Tests.Redlining;

public class RedlineExportServiceTests : IDisposable
{
    private readonly string _originalPath = TestFiles.NewTempPath(".docx");
    private readonly string _revisedPath = TestFiles.NewTempPath(".docx");
    private readonly string _outputDir = Path.Combine(Path.GetTempPath(), $"docproc-redline-export-{Guid.NewGuid():N}");
    private readonly RedlineExportService _sut = new();

    public RedlineExportServiceTests()
    {
        SampleDocumentFactory.CreateBasicDocument(_originalPath, "Agreement", ["The term is twelve months."]);
        SampleDocumentFactory.CreateBasicDocument(_revisedPath, "Agreement", ["The term is twenty-four months."]);
    }

    [Fact]
    public void ExportAllVariants_produces_all_four_deliverables()
    {
        var result = _sut.ExportAllVariants(_originalPath, _revisedPath, _outputDir, conversionOptions: TestFiles.ConversionOptions());

        Assert.True(File.Exists(result.RedlinedDocx));
        Assert.True(File.Exists(result.CleanDocx));
        Assert.True(File.Exists(result.RedlinedPdf));
        Assert.True(File.Exists(result.CleanPdf));
    }

    [Fact]
    public void ExportAllVariants_clean_docx_has_no_tracked_changes_but_redlined_does()
    {
        var result = _sut.ExportAllVariants(_originalPath, _revisedPath, _outputDir, conversionOptions: TestFiles.ConversionOptions());

        Assert.True(new Core.TrackChanges.TrackChangesService().HasTrackedChanges(result.RedlinedDocx));
        Assert.False(new Core.TrackChanges.TrackChangesService().HasTrackedChanges(result.CleanDocx));
    }

    public void Dispose()
    {
        File.Delete(_originalPath);
        File.Delete(_revisedPath);
        if (Directory.Exists(_outputDir))
            Directory.Delete(_outputDir, recursive: true);
    }
}
