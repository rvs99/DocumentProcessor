using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.Samples;
using DocumentProcessor.Core.Tables;

namespace DocumentProcessor.Tests.Tables;

public class TableGenerationServiceTests : IDisposable
{
    private readonly string _path = TestFiles.NewTempPath(".docx");
    private readonly TableGenerationService _sut = new();

    public TableGenerationServiceTests()
    {
        SampleDocumentFactory.CreateBasicDocument(_path, "Invoice", ["Some intro paragraph."]);
    }

    [Fact]
    public void AppendTable_inserts_a_table_with_the_expected_dimensions_and_cell_text()
    {
        var spec = new TableSpec(
            Headers: ["Item", "Qty", "Price"],
            Rows:
            [
                ["Widget", "3", "$10.00"],
                ["Gadget", "1", "$25.00"]
            ],
            Caption: "Line Items");

        _sut.AppendTable(_path, spec);

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var table = doc.MainDocumentPart!.Document!.Body!.Elements<Table>().Single();
        var rows = table.Elements<TableRow>().ToList();

        Assert.Equal(3, rows.Count); // header + 2 data rows
        Assert.Equal(["Item", "Qty", "Price"], CellText(rows[0]));
        Assert.Equal(["Widget", "3", "$10.00"], CellText(rows[1]));
        Assert.Equal(["Gadget", "1", "$25.00"], CellText(rows[2]));
    }

    [Fact]
    public void AppendTable_rejects_a_row_whose_column_count_does_not_match_the_headers()
    {
        var spec = new TableSpec(
            Headers: ["Item", "Qty"],
            Rows: [["Widget", "3", "extra"]]);

        Assert.Throws<ArgumentException>(() => _sut.AppendTable(_path, spec));
    }

    [Fact]
    public void ReplaceTable_swaps_an_existing_table_for_a_new_one()
    {
        _sut.AppendTable(_path, new TableSpec(["A"], [["1"]]));
        _sut.AppendTable(_path, new TableSpec(["B"], [["2"]]));

        _sut.ReplaceTable(_path, 0, new TableSpec(["Replaced"], [["yes"]]));

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var tables = doc.MainDocumentPart!.Document!.Body!.Elements<Table>().ToList();

        Assert.Equal(2, tables.Count);
        Assert.Equal(["Replaced"], CellText(tables[0].Elements<TableRow>().First()));
        Assert.Equal(["B"], CellText(tables[1].Elements<TableRow>().First()));
    }

    private static List<string> CellText(TableRow row) =>
        row.Elements<TableCell>().Select(c => c.InnerText).ToList();

    public void Dispose() => File.Delete(_path);
}
