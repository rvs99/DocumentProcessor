using System.Diagnostics;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.Samples;
using DocumentProcessor.Core.Tables;
using DocumentProcessor.Core.Templating;

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

    [Fact]
    public void PopulateFromPrototypeRow_clones_the_last_row_once_per_item_and_removes_the_prototype()
    {
        _sut.AppendTable(_path, new TableSpec(
            Headers: ["Item", "Qty"],
            Rows: [["{{Name}}", "{{Quantity}}"]]));

        var rows = new List<IReadOnlyDictionary<string, object?>>
        {
            new Dictionary<string, object?> { ["Name"] = "Widget", ["Quantity"] = "3" },
            new Dictionary<string, object?> { ["Name"] = "Gadget", ["Quantity"] = "1" }
        };

        var generated = _sut.PopulateFromPrototypeRow(_path, tableIndex: 0, rows);

        Assert.Equal(2, generated);
        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var table = doc.MainDocumentPart!.Document!.Body!.Elements<Table>().Single();
        var dataRows = table.Elements<TableRow>().Skip(1).ToList(); // skip header

        Assert.Equal(2, dataRows.Count); // prototype row itself is gone, replaced by exactly 2 clones
        Assert.Equal(["Widget", "3"], CellText(dataRows[0]));
        Assert.Equal(["Gadget", "1"], CellText(dataRows[1]));
    }

    [Fact]
    public void PopulateFromPrototypeRow_with_Error_policy_throws_on_a_missing_field()
    {
        _sut.AppendTable(_path, new TableSpec(["Item"], [["{{Missing}}"]]));

        Assert.Throws<MissingTemplateTokenException>(() => _sut.PopulateFromPrototypeRow(
            _path, 0, [new Dictionary<string, object?>()]));
    }

    [Fact]
    public void PopulateFromPrototypeRow_scales_to_ten_thousand_rows_in_a_reasonable_time()
    {
        _sut.AppendTable(_path, new TableSpec(
            Headers: ["Index", "Name"],
            Rows: [["{{Index}}", "{{Name}}"]]));

        var rows = Enumerable.Range(0, 10_000)
            .Select(i => (IReadOnlyDictionary<string, object?>)new Dictionary<string, object?>
            {
                ["Index"] = i.ToString(),
                ["Name"] = $"Row {i}"
            })
            .ToList();

        var stopwatch = Stopwatch.StartNew();
        var generated = _sut.PopulateFromPrototypeRow(_path, 0, rows);
        stopwatch.Stop();

        Assert.Equal(10_000, generated);
        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(30),
            $"Populating 10,000 rows took {stopwatch.Elapsed}, expected under 30s.");

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var table = doc.MainDocumentPart!.Document!.Body!.Elements<Table>().Single();
        var dataRows = table.Elements<TableRow>().Skip(1).ToList();
        Assert.Equal(10_000, dataRows.Count);
        Assert.Equal(["0", "Row 0"], CellText(dataRows[0]));
        Assert.Equal(["9999", "Row 9999"], CellText(dataRows[9999]));
    }

    private static List<string> CellText(TableRow row) =>
        row.Elements<TableCell>().Select(c => c.InnerText).ToList();

    public void Dispose() => File.Delete(_path);
}
