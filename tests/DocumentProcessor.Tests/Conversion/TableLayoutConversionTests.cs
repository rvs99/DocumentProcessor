using DocumentProcessor.Core.Conversion;
using DocumentProcessor.Core.Samples;
using DocumentProcessor.Core.Tables;
using UglyToad.PdfPig;

namespace DocumentProcessor.Tests.Conversion;

/// <summary>
/// Verifies a table using the Phase 2 layout options (merged cells, explicit column widths) survives
/// a real LibreOffice conversion with its text intact — merged-cell XML in particular is easy to get
/// subtly wrong in a way that still opens fine in the OpenXml SDK but confuses a renderer.
/// </summary>
public class TableLayoutConversionTests : IDisposable
{
    private readonly string _docxPath = TestFiles.NewTempPath(".docx");
    private readonly string _pdfPath = TestFiles.NewTempPath(".pdf");

    public TableLayoutConversionTests()
    {
        SampleDocumentFactory.CreateBasicDocument(_docxPath, "Table Layout Conversion Test", ["Intro."]);
    }

    [Fact]
    public async Task A_table_with_merged_cells_and_explicit_column_widths_converts_and_keeps_its_text()
    {
        var spec = new TableSpec(
            Headers: ["Category", "Item", "Price"],
            Rows:
            [
                ["Fruit", "Apple", "$1.00"],
                ["Fruit", "Banana", "$0.50"],
                ["Vegetable", "Carrot", "$0.75"]
            ],
            ColumnWidthsTwips: [2000, 3000, 2000],
            Borders: new TableBorderSpec { SizeEighthPoints = 8, ColorHex = "2E74B5" },
            Merges:
            [
                new TableCellMerge(RowIndex: 1, ColumnIndex: 0, Span: 2, Direction: MergeDirection.Vertical)
            ]);

        new TableGenerationService().AppendTable(_docxPath, spec);
        await new WordToPdfConverter(TestFiles.ConversionOptions()).ConvertAsync(_docxPath, _pdfPath);

        using var doc = PdfDocument.Open(_pdfPath);
        var text = doc.GetPage(1).Text;

        Assert.Contains("Fruit", text);
        Assert.Contains("Apple", text);
        Assert.Contains("Banana", text);
        Assert.Contains("Vegetable", text);
        Assert.Contains("Carrot", text);
    }

    public void Dispose()
    {
        File.Delete(_docxPath);
        if (File.Exists(_pdfPath))
            File.Delete(_pdfPath);
    }
}
