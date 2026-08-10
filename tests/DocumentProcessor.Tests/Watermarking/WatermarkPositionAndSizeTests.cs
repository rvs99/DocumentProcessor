using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.Comparison;
using DocumentProcessor.Core.Samples;
using DocumentProcessor.Core.Watermarking;

namespace DocumentProcessor.Tests.Watermarking;

public class DocxWatermarkPositionAndSizeTests : IDisposable
{
    private readonly string _path = TestFiles.NewTempPath(".docx");
    private readonly DocxWatermarkService _sut = new();

    public DocxWatermarkPositionAndSizeTests()
    {
        SampleDocumentFactory.CreateBasicDocument(_path, "Watermark Position Test", ["Body text."]);
    }

    [Theory]
    [InlineData(WatermarkPosition.Center, "center", "center")]
    [InlineData(WatermarkPosition.TopLeft, "left", "top")]
    [InlineData(WatermarkPosition.TopCenter, "center", "top")]
    [InlineData(WatermarkPosition.TopRight, "right", "top")]
    [InlineData(WatermarkPosition.MiddleLeft, "left", "center")]
    [InlineData(WatermarkPosition.MiddleRight, "right", "center")]
    [InlineData(WatermarkPosition.BottomLeft, "left", "bottom")]
    [InlineData(WatermarkPosition.BottomCenter, "center", "bottom")]
    [InlineData(WatermarkPosition.BottomRight, "right", "bottom")]
    public void AddTextWatermark_writes_the_VML_alignment_keywords_matching_the_requested_position(
        WatermarkPosition position, string expectedHorizontal, string expectedVertical)
    {
        _sut.AddTextWatermark(_path, "DRAFT", position: position);

        var style = GetWatermarkShape().Style!.Value!;
        Assert.Contains($"mso-position-horizontal:{expectedHorizontal}", style);
        Assert.Contains($"mso-position-vertical:{expectedVertical}", style);
    }

    [Fact]
    public void AddTextWatermark_applies_custom_box_size()
    {
        _sut.AddTextWatermark(_path, "DRAFT", widthPt: 300, heightPt: 90);

        var style = GetWatermarkShape().Style!.Value!;
        Assert.Contains("width:300pt", style);
        Assert.Contains("height:90pt", style);
    }

    [Fact]
    public void AddTextWatermark_applies_custom_font_size_in_half_points()
    {
        _sut.AddTextWatermark(_path, "DRAFT", fontSizePt: 40);

        var fontSize = GetWatermarkShape().Descendants<FontSize>().Single();
        Assert.Equal("80", fontSize.Val!.Value); // 40pt * 2 = 80 half-points
    }

    private Shape GetWatermarkShape()
    {
        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        return doc.MainDocumentPart!.HeaderParts.Single().Header!.Descendants<Shape>().Single();
    }

    public void Dispose() => File.Delete(_path);
}

public class PdfWatermarkPositionAndSizeTests : IDisposable
{
    private readonly string _sourcePdfPath;
    private readonly List<string> _outputPaths = [];
    private readonly PdfWatermarkService _sut = new();

    public PdfWatermarkPositionAndSizeTests()
    {
        _sourcePdfPath = TestFiles.NewTestPdf("Watermark Position Test", ["Body text that should not move."]);
    }

    [Theory]
    [InlineData(WatermarkPosition.Center)]
    [InlineData(WatermarkPosition.TopLeft)]
    [InlineData(WatermarkPosition.TopCenter)]
    [InlineData(WatermarkPosition.TopRight)]
    [InlineData(WatermarkPosition.MiddleLeft)]
    [InlineData(WatermarkPosition.MiddleRight)]
    [InlineData(WatermarkPosition.BottomLeft)]
    [InlineData(WatermarkPosition.BottomCenter)]
    [InlineData(WatermarkPosition.BottomRight)]
    public void AddTextWatermark_runs_without_error_for_every_position(WatermarkPosition position)
    {
        var outputPath = NewOutputPath();

        _sut.AddTextWatermark(_sourcePdfPath, outputPath, "DRAFT", position: position);

        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void AddTextWatermark_center_vs_top_left_produce_visibly_different_output()
    {
        var centerPath = NewOutputPath();
        var topLeftPath = NewOutputPath();

        _sut.AddTextWatermark(_sourcePdfPath, centerPath, "DRAFT", position: WatermarkPosition.Center);
        _sut.AddTextWatermark(_sourcePdfPath, topLeftPath, "DRAFT", position: WatermarkPosition.TopLeft);

        var diff = new PdfComparisonService().CompareVisual(centerPath, topLeftPath);

        Assert.True(diff.PageCountMatches);
        Assert.True(diff.PerPageDifferencePercent.Single() > 0.5,
            $"Expected a visible difference between center and top-left placement, got {diff.PerPageDifferencePercent.Single():F2}%.");
    }

    [Fact]
    public void AddTextWatermark_applies_a_custom_font_size()
    {
        var smallPath = NewOutputPath();
        var largePath = NewOutputPath();

        _sut.AddTextWatermark(_sourcePdfPath, smallPath, "DRAFT", fontSizePt: 20);
        _sut.AddTextWatermark(_sourcePdfPath, largePath, "DRAFT", fontSizePt: 100);

        var diff = new PdfComparisonService().CompareVisual(smallPath, largePath);

        Assert.True(diff.PerPageDifferencePercent.Single() > 0.5,
            $"Expected a visible difference between a 20pt and a 100pt watermark, got {diff.PerPageDifferencePercent.Single():F2}%.");
    }

    private string NewOutputPath()
    {
        var path = TestFiles.NewTempPath(".pdf");
        _outputPaths.Add(path);
        return path;
    }

    public void Dispose()
    {
        File.Delete(_sourcePdfPath);
        foreach (var path in _outputPaths)
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }
}
