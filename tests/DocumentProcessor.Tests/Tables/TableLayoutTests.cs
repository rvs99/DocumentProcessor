using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.Samples;
using DocumentProcessor.Core.Tables;

namespace DocumentProcessor.Tests.Tables;

/// <summary>Covers the Phase 2 additions to <see cref="TableSpec"/>: column widths, borders, table style, and merged cells.</summary>
public class TableLayoutTests : IDisposable
{
    private readonly string _path = TestFiles.NewTempPath(".docx");
    private readonly TableGenerationService _sut = new();

    public TableLayoutTests()
    {
        SampleDocumentFactory.CreateBasicDocument(_path, "Table Layout Test", ["Intro."]);
    }

    [Fact]
    public void ColumnWidthsTwips_sets_an_explicit_width_on_the_grid_and_every_cell()
    {
        var spec = new TableSpec(["A", "B"], [["1", "2"]], ColumnWidthsTwips: [3000, 1500]);

        _sut.AppendTable(_path, spec);

        var table = GetTable();
        var gridColumns = table.Elements<TableGrid>().Single().Elements<GridColumn>().ToList();
        Assert.Equal("3000", gridColumns[0].Width!.Value);
        Assert.Equal("1500", gridColumns[1].Width!.Value);

        var firstRowCells = table.Elements<TableRow>().First().Elements<TableCell>().ToList();
        var widths = firstRowCells.Select(c => c.TableCellProperties!.TableCellWidth!.Width!.Value).ToList();
        Assert.Equal(["3000", "1500"], widths);
    }

    [Fact]
    public void ColumnWidthsTwips_with_wrong_count_throws()
    {
        var spec = new TableSpec(["A", "B"], [["1", "2"]], ColumnWidthsTwips: [3000]);

        Assert.Throws<ArgumentException>(() => _sut.AppendTable(_path, spec));
    }

    [Fact]
    public void Borders_applies_the_requested_style_size_and_color_to_all_six_positions()
    {
        var spec = new TableSpec(["A"], [["1"]],
            Borders: new TableBorderSpec { Style = BorderValues.Double, SizeEighthPoints = 12, ColorHex = "FF0000" });

        _sut.AppendTable(_path, spec);

        var borders = GetTable().Elements<TableProperties>().Single().Elements<TableBorders>().Single();
        Assert.Equal(BorderValues.Double, borders.TopBorder!.Val!.Value);
        Assert.Equal(12u, borders.TopBorder!.Size!.Value);
        Assert.Equal("FF0000", borders.TopBorder!.Color!.Value);
        Assert.Equal(BorderValues.Double, borders.InsideVerticalBorder!.Val!.Value);
    }

    [Fact]
    public void TableStyleId_sets_the_table_style_reference()
    {
        var spec = new TableSpec(["A"], [["1"]], TableStyleId: "GridTable1");

        _sut.AppendTable(_path, spec);

        var tableStyle = GetTable().Elements<TableProperties>().Single().Elements<TableStyle>().Single();
        Assert.Equal("GridTable1", tableStyle.Val!.Value);
    }

    [Fact]
    public void Horizontal_merge_gives_the_anchor_cell_a_GridSpan_and_omits_the_covered_cells()
    {
        // Row 0 (header): merge columns 0-1 into one cell.
        var spec = new TableSpec(["Merged", "B", "C"], [["x", "y", "z"]],
            Merges: [new TableCellMerge(RowIndex: 0, ColumnIndex: 0, Span: 2, Direction: MergeDirection.Horizontal)]);

        _sut.AppendTable(_path, spec);

        var headerRow = GetTable().Elements<TableRow>().First();
        var cells = headerRow.Elements<TableCell>().ToList();

        Assert.Equal(2, cells.Count); // 3 columns, but 2 are merged into 1 cell -> 2 physical cells
        Assert.Equal(2, cells[0].TableCellProperties!.GridSpan!.Val!.Value);
        Assert.Equal("Merged", cells[0].InnerText);
        Assert.Equal("C", cells[1].InnerText);
    }

    [Fact]
    public void Vertical_merge_marks_the_anchor_as_restart_and_the_continuation_row_as_continue()
    {
        // Data rows 1-2 (i.e. spec.Rows[0] and spec.Rows[1]): merge column 0 vertically.
        var spec = new TableSpec(["Category", "Item"], [["Fruit", "Apple"], ["Fruit", "Banana"]],
            Merges: [new TableCellMerge(RowIndex: 1, ColumnIndex: 0, Span: 2, Direction: MergeDirection.Vertical)]);

        _sut.AppendTable(_path, spec);

        var rows = GetTable().Elements<TableRow>().ToList();
        var anchorCell = rows[1].Elements<TableCell>().First();
        var continuationCell = rows[2].Elements<TableCell>().First();

        Assert.Equal(MergedCellValues.Restart, anchorCell.TableCellProperties!.VerticalMerge!.Val!.Value);
        Assert.Equal("Fruit", anchorCell.InnerText);

        Assert.NotNull(continuationCell.TableCellProperties!.VerticalMerge);
        Assert.Null(continuationCell.TableCellProperties!.VerticalMerge!.Val);
        Assert.Equal("", continuationCell.InnerText);
    }

    [Fact]
    public void Overlapping_merges_throw()
    {
        var spec = new TableSpec(["A", "B", "C"], [["1", "2", "3"]],
            Merges:
            [
                new TableCellMerge(0, 0, 2, MergeDirection.Horizontal),
                new TableCellMerge(0, 1, 2, MergeDirection.Horizontal)
            ]);

        Assert.Throws<ArgumentException>(() => _sut.AppendTable(_path, spec));
    }

    [Fact]
    public void Merge_span_below_two_throws()
    {
        var spec = new TableSpec(["A", "B"], [["1", "2"]],
            Merges: [new TableCellMerge(0, 0, 1, MergeDirection.Horizontal)]);

        Assert.Throws<ArgumentException>(() => _sut.AppendTable(_path, spec));
    }

    [Fact]
    public void Merge_outside_table_bounds_throws()
    {
        var spec = new TableSpec(["A", "B"], [["1", "2"]],
            Merges: [new TableCellMerge(0, 1, 2, MergeDirection.Horizontal)]); // column 1 + span 2 = 3, only 2 columns

        Assert.Throws<ArgumentOutOfRangeException>(() => _sut.AppendTable(_path, spec));
    }

    private Table GetTable()
    {
        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        return doc.MainDocumentPart!.Document!.Body!.Elements<Table>().Single();
    }

    public void Dispose() => File.Delete(_path);
}
