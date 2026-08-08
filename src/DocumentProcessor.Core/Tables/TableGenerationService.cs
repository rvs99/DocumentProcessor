using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentProcessor.Core.Tables;

public sealed record TableSpec(IReadOnlyList<string> Headers, IReadOnlyList<IReadOnlyList<string>> Rows, string? Caption = null);

/// <summary>
/// Builds Word tables programmatically (e.g. from a data grid, report, or clause schedule)
/// and inserts them into an existing .docx, or appends new tables to one.
/// </summary>
public sealed class TableGenerationService
{
    /// <summary>Appends a table built from <paramref name="spec"/> to the end of the document body.</summary>
    public void AppendTable(string docxPath, TableSpec spec)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var document = doc.MainDocumentPart?.Document ?? throw new InvalidOperationException("Document has no main part/body.");
        var body = document.Body ?? throw new InvalidOperationException("Document has no body.");
        var sectPr = body.Elements<SectionProperties>().FirstOrDefault();

        if (spec.Caption is not null)
            body.InsertBefore(new Paragraph(new Run(new RunProperties(new Bold()), new Text(spec.Caption))), sectPr);

        var table = BuildTable(spec);
        body.InsertBefore(table, sectPr);
        // Word requires a paragraph after a table that isn't followed by another block element.
        body.InsertBefore(new Paragraph(), sectPr);

        document.Save();
    }

    /// <summary>Replaces the Nth table (0-indexed, document order) with a freshly generated one.</summary>
    public void ReplaceTable(string docxPath, int tableIndex, TableSpec spec)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var document = doc.MainDocumentPart?.Document ?? throw new InvalidOperationException("Document has no main part/body.");
        var body = document.Body ?? throw new InvalidOperationException("Document has no body.");

        var existing = body.Elements<Table>().ElementAtOrDefault(tableIndex)
            ?? throw new ArgumentOutOfRangeException(nameof(tableIndex), $"Document has no table at index {tableIndex}.");

        var replacement = BuildTable(spec);
        body.ReplaceChild(replacement, existing);
        document.Save();
    }

    private Table BuildTable(TableSpec spec)
    {
        var columnCount = spec.Headers.Count;
        var table = new Table();

        table.AppendChild(new TableProperties(
            new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" },
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4 },
                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                new LeftBorder { Val = BorderValues.Single, Size = 4 },
                new RightBorder { Val = BorderValues.Single, Size = 4 },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 })));

        var grid = new TableGrid();
        for (var i = 0; i < columnCount; i++)
            grid.AppendChild(new GridColumn());
        table.AppendChild(grid);

        table.AppendChild(BuildRow(spec.Headers, isHeader: true));
        foreach (var row in spec.Rows)
        {
            if (row.Count != columnCount)
                throw new ArgumentException($"Row has {row.Count} cells but table has {columnCount} header columns.");
            table.AppendChild(BuildRow(row, isHeader: false));
        }

        return table;
    }

    private static TableRow BuildRow(IReadOnlyList<string> cellValues, bool isHeader)
    {
        var row = new TableRow();
        if (isHeader)
            row.AppendChild(new TableRowProperties(new TableHeader()));

        foreach (var value in cellValues)
        {
            var runProps = isHeader ? new RunProperties(new Bold()) : null;
            var run = runProps is not null ? new Run(runProps, new Text(value)) : new Run(new Text(value));
            var cell = new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Auto }),
                new Paragraph(run));
            row.AppendChild(cell);
        }

        return row;
    }
}
