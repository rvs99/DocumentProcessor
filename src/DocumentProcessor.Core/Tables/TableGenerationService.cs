using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.Templating;

namespace DocumentProcessor.Core.Tables;

/// <summary>Uniform border style applied to all six border positions (outer edges + inside lines).</summary>
public sealed record TableBorderSpec
{
    public BorderValues Style { get; init; } = BorderValues.Single;
    public uint SizeEighthPoints { get; init; } = 4;
    public string? ColorHex { get; init; }
}

public enum MergeDirection { Horizontal, Vertical }

/// <summary>
/// Merges <paramref name="Span"/> cells starting at (<paramref name="RowIndex"/>,
/// <paramref name="ColumnIndex"/>) — row 0 is the header row, row 1+ are data rows — in
/// <paramref name="Direction"/>. Merges may not overlap each other; only one merge may claim a
/// given cell.
/// </summary>
public sealed record TableCellMerge(int RowIndex, int ColumnIndex, int Span, MergeDirection Direction);

/// <summary>Describes a table to generate: its content plus optional layout and styling.</summary>
/// <param name="Headers">Header-row cell text; the column count is taken from this.</param>
/// <param name="Rows">Data rows, each with one entry per header column.</param>
/// <param name="Caption">Optional bold caption paragraph inserted above the table.</param>
/// <param name="ColumnWidthsTwips">Explicit column widths in twips. Must have one entry per header column when set; null means auto-width.</param>
/// <param name="Borders">Border style for all six border positions. Null keeps the default single 0.5pt line with no explicit colour.</param>
/// <param name="TableStyleId">References a table style already defined in the target document's styles part. Not validated — the caller is responsible for the style existing.</param>
/// <param name="Merges">Cell merges to apply. Merges may not overlap.</param>
public sealed record TableSpec(
    IReadOnlyList<string> Headers,
    IReadOnlyList<IReadOnlyList<string>> Rows,
    string? Caption = null,
    IReadOnlyList<int>? ColumnWidthsTwips = null,
    TableBorderSpec? Borders = null,
    string? TableStyleId = null,
    IReadOnlyList<TableCellMerge>? Merges = null);

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
        AppendTableCore(doc, spec);
    }

    internal void AppendTableCore(WordprocessingDocument doc, TableSpec spec)
    {
        var document = doc.MainDocumentPart?.Document ?? throw new CorruptDocumentException("Document has no main part/body.");
        var body = document.Body ?? throw new CorruptDocumentException("Document has no body.");
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
        ReplaceTableCore(doc, tableIndex, spec);
    }

    internal void ReplaceTableCore(WordprocessingDocument doc, int tableIndex, TableSpec spec)
    {
        var document = doc.MainDocumentPart?.Document ?? throw new CorruptDocumentException("Document has no main part/body.");
        var body = document.Body ?? throw new CorruptDocumentException("Document has no body.");

        var existing = body.Elements<Table>().ElementAtOrDefault(tableIndex)
            ?? throw new ArgumentOutOfRangeException(nameof(tableIndex), $"Document has no table at index {tableIndex}.");

        var replacement = BuildTable(spec);
        body.ReplaceChild(replacement, existing);
        document.Save();
    }

    /// <summary>
    /// Populates a table by cloning its last row as a <c>{{token}}</c>-templated prototype — once
    /// per entry in <paramref name="rows"/> — rather than requiring the caller to pre-build every
    /// row's cell text themselves as <see cref="AppendTable"/>/<see cref="ReplaceTable"/> do. Each
    /// cell in the prototype row may contain any number of <c>{{field}}</c> tokens, resolved against
    /// that row's own dictionary via the same run-merged scanner <see cref="TemplateEngine"/> uses,
    /// so a token split across several runs (a real risk in a prototype row typed directly in Word)
    /// still resolves correctly. The prototype row itself is removed after the last clone is inserted.
    /// </summary>
    /// <returns>The number of rows generated (equal to <paramref name="rows"/>.Count).</returns>
    public int PopulateFromPrototypeRow(
        string docxPath, int tableIndex,
        IReadOnlyList<IReadOnlyDictionary<string, object?>> rows,
        MissingTokenPolicy missingTokenPolicy = MissingTokenPolicy.Error,
        CancellationToken cancellationToken = default)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        return PopulateFromPrototypeRowCore(doc, tableIndex, rows, missingTokenPolicy, cancellationToken);
    }

    internal int PopulateFromPrototypeRowCore(
        WordprocessingDocument doc, int tableIndex,
        IReadOnlyList<IReadOnlyDictionary<string, object?>> rows,
        MissingTokenPolicy missingTokenPolicy,
        CancellationToken cancellationToken)
    {
        var mainPart = doc.MainDocumentPart ?? throw new CorruptDocumentException("Document has no main part.");
        var document = mainPart.Document ?? throw new CorruptDocumentException("Document has no body.");
        var body = document.Body ?? throw new CorruptDocumentException("Document has no body.");

        var table = body.Elements<Table>().ElementAtOrDefault(tableIndex)
            ?? throw new ArgumentOutOfRangeException(nameof(tableIndex), $"Document has no table at index {tableIndex}.");

        var prototypeRow = table.Elements<TableRow>().LastOrDefault()
            ?? throw new InvalidOperationException("Table has no rows to use as a prototype.");

        // Checked every 100 rows rather than every row: ThrowIfCancellationRequested's overhead is
        // trivial per-call, but at 10,000+ rows even a trivial per-iteration check adds up, and
        // cancellation latency of "up to 100 rows late" is irrelevant at this operation's timescale.
        var processed = 0;
        foreach (var item in rows)
        {
            if (++processed % 100 == 0)
                cancellationToken.ThrowIfCancellationRequested();

            var clone = (TableRow)prototypeRow.CloneNode(true);
            var context = new TemplateContext(item);
            foreach (var paragraph in clone.Descendants<Paragraph>().ToList())
                TemplateEngine.SubstituteInline(paragraph, context, mainPart, missingTokenPolicy);

            prototypeRow.InsertBeforeSelf(clone);
        }

        prototypeRow.Remove();
        document.Save();
        return rows.Count;
    }

    private Table BuildTable(TableSpec spec)
    {
        var columnCount = spec.Headers.Count;
        var rowCount = spec.Rows.Count + 1; // +1 for the header row

        if (spec.ColumnWidthsTwips is not null && spec.ColumnWidthsTwips.Count != columnCount)
        {
            throw new ArgumentException(
                $"ColumnWidthsTwips has {spec.ColumnWidthsTwips.Count} entries but the table has {columnCount} columns.",
                nameof(spec));
        }

        var roles = ResolveMergeRoles(spec.Merges, rowCount, columnCount);

        var table = new Table();
        table.AppendChild(BuildTableProperties(spec));
        table.AppendChild(BuildTableGrid(columnCount, spec.ColumnWidthsTwips));

        for (var row = 0; row < rowCount; row++)
        {
            var cellValues = row == 0 ? spec.Headers : spec.Rows[row - 1];
            if (row > 0 && cellValues.Count != columnCount)
                throw new ArgumentException($"Row has {cellValues.Count} cells but table has {columnCount} header columns.");

            table.AppendChild(BuildRow(cellValues, isHeader: row == 0, row, columnCount, roles, spec.ColumnWidthsTwips));
        }

        return table;
    }

    private static TableProperties BuildTableProperties(TableSpec spec)
    {
        var borders = spec.Borders ?? new TableBorderSpec();
        var props = new TableProperties(
            spec.ColumnWidthsTwips is { } widths
                ? new TableWidth { Type = TableWidthUnitValues.Dxa, Width = widths.Sum().ToString() }
                : new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" },
            new TableBorders(
                BuildBorder<TopBorder>(borders), BuildBorder<BottomBorder>(borders),
                BuildBorder<LeftBorder>(borders), BuildBorder<RightBorder>(borders),
                BuildBorder<InsideHorizontalBorder>(borders), BuildBorder<InsideVerticalBorder>(borders)));

        if (spec.ColumnWidthsTwips is not null)
            props.AppendChild(new TableLayout { Type = TableLayoutValues.Fixed });

        if (spec.TableStyleId is not null)
            props.AppendChild(new TableStyle { Val = spec.TableStyleId });

        return props;
    }

    private static TBorder BuildBorder<TBorder>(TableBorderSpec spec) where TBorder : BorderType, new()
    {
        var border = new TBorder { Val = spec.Style, Size = spec.SizeEighthPoints };
        if (spec.ColorHex is not null)
            border.Color = spec.ColorHex;
        return border;
    }

    private static TableGrid BuildTableGrid(int columnCount, IReadOnlyList<int>? columnWidthsTwips)
    {
        var grid = new TableGrid();
        for (var i = 0; i < columnCount; i++)
        {
            grid.AppendChild(columnWidthsTwips is not null
                ? new GridColumn { Width = columnWidthsTwips[i].ToString() }
                : new GridColumn());
        }
        return grid;
    }

    private static TableRow BuildRow(
        IReadOnlyList<string> cellValues, bool isHeader, int rowIndex, int columnCount,
        CellRole[,] roles, IReadOnlyList<int>? columnWidthsTwips)
    {
        var row = new TableRow();
        if (isHeader)
            row.AppendChild(new TableRowProperties(new TableHeader()));

        for (var col = 0; col < columnCount; col++)
        {
            var role = roles[rowIndex, col];
            if (role.Kind == MergeRoleKind.HorizontalContinuation)
                continue; // absorbed into the horizontal-merge anchor's GridSpan; no cell of its own

            var isContinuationCell = role.Kind == MergeRoleKind.VerticalContinuation;
            var text = isContinuationCell ? "" : cellValues[col];

            var cellProps = new TableCellProperties(columnWidthsTwips is not null
                ? new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = columnWidthsTwips[col].ToString() }
                : new TableCellWidth { Type = TableWidthUnitValues.Auto });

            if (role.Kind == MergeRoleKind.HorizontalAnchor)
                cellProps.AppendChild(new GridSpan { Val = role.Span });
            else if (role.Kind == MergeRoleKind.VerticalAnchor)
                cellProps.AppendChild(new VerticalMerge { Val = MergedCellValues.Restart });
            else if (isContinuationCell)
                cellProps.AppendChild(new VerticalMerge());

            var runProps = isHeader ? new RunProperties(new Bold()) : null;
            var run = runProps is not null ? new Run(runProps, new Text(text)) : new Run(new Text(text));
            row.AppendChild(new TableCell(cellProps, new Paragraph(run)));
        }

        return row;
    }

    private enum MergeRoleKind { None, HorizontalAnchor, HorizontalContinuation, VerticalAnchor, VerticalContinuation }

    private readonly record struct CellRole(MergeRoleKind Kind, int Span);

    private static CellRole[,] ResolveMergeRoles(IReadOnlyList<TableCellMerge>? merges, int rowCount, int columnCount)
    {
        var roles = new CellRole[rowCount, columnCount];
        if (merges is null)
            return roles;

        var claimed = new bool[rowCount, columnCount];

        foreach (var merge in merges)
        {
            if (merge.Span < 2)
                throw new ArgumentException($"Merge span must be at least 2 (got {merge.Span}).", nameof(merges));

            var rowSpan = merge.Direction == MergeDirection.Vertical ? merge.Span : 1;
            var colSpan = merge.Direction == MergeDirection.Horizontal ? merge.Span : 1;

            if (merge.RowIndex < 0 || merge.ColumnIndex < 0 ||
                merge.RowIndex + rowSpan > rowCount || merge.ColumnIndex + colSpan > columnCount)
            {
                throw new ArgumentOutOfRangeException(nameof(merges),
                    $"Merge at row {merge.RowIndex}, column {merge.ColumnIndex} with span {merge.Span} " +
                    $"({merge.Direction}) falls outside the {rowCount}x{columnCount} table.");
            }

            for (var r = merge.RowIndex; r < merge.RowIndex + rowSpan; r++)
            {
                for (var c = merge.ColumnIndex; c < merge.ColumnIndex + colSpan; c++)
                {
                    if (claimed[r, c])
                    {
                        throw new ArgumentException(
                            $"Cell (row {r}, column {c}) is claimed by more than one merge — overlapping merges are not supported.",
                            nameof(merges));
                    }
                    claimed[r, c] = true;
                }
            }

            var anchorKind = merge.Direction == MergeDirection.Horizontal ? MergeRoleKind.HorizontalAnchor : MergeRoleKind.VerticalAnchor;
            roles[merge.RowIndex, merge.ColumnIndex] = new CellRole(anchorKind, merge.Span);

            if (merge.Direction == MergeDirection.Horizontal)
            {
                for (var c = merge.ColumnIndex + 1; c < merge.ColumnIndex + colSpan; c++)
                    roles[merge.RowIndex, c] = new CellRole(MergeRoleKind.HorizontalContinuation, 0);
            }
            else
            {
                for (var r = merge.RowIndex + 1; r < merge.RowIndex + rowSpan; r++)
                    roles[r, merge.ColumnIndex] = new CellRole(MergeRoleKind.VerticalContinuation, 0);
            }
        }

        return roles;
    }
}

/// <summary>
/// Table operations bound to an open <see cref="Sessions.DocumentSession"/>.
/// </summary>
public sealed class TableOperations
{
    private readonly Sessions.DocumentSession _session;
    private readonly TableGenerationService _service = new();

    internal TableOperations(Sessions.DocumentSession session) => _session = session;

    /// <inheritdoc cref="TableGenerationService.AppendTable(string, TableSpec)"/>
    public void AppendTable(TableSpec spec) => _service.AppendTableCore(_session.Document, spec);

    /// <inheritdoc cref="TableGenerationService.ReplaceTable(string, int, TableSpec)"/>
    public void ReplaceTable(int tableIndex, TableSpec spec) => _service.ReplaceTableCore(_session.Document, tableIndex, spec);

    /// <inheritdoc cref="TableGenerationService.PopulateFromPrototypeRow(string, int, IReadOnlyList{IReadOnlyDictionary{string, object}}, MissingTokenPolicy, CancellationToken)"/>
    public int PopulateFromPrototypeRow(
        int tableIndex,
        IReadOnlyList<IReadOnlyDictionary<string, object?>> rows,
        MissingTokenPolicy missingTokenPolicy = MissingTokenPolicy.Error,
        CancellationToken cancellationToken = default) =>
        _service.PopulateFromPrototypeRowCore(_session.Document, tableIndex, rows, missingTokenPolicy, cancellationToken);
}
