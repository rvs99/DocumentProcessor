namespace DocumentProcessor.Core.Format;

/// <summary>Configuration for the LibreOffice headless process this converter shells out to —
/// deliberately a separate type from <c>Conversion.WordToPdfConversionOptions</c> (same shape, but
/// that one's name and namespace are specific to the docx→PDF path) rather than reusing it.</summary>
public sealed class LegacyDocConversionOptions
{
    public string ExecutablePath { get; init; } = OperatingSystem.IsWindows()
        ? @"C:\Program Files\LibreOffice\program\soffice.exe"
        : "/usr/bin/soffice";

    public string? UseWslDistro { get; init; }

    public TimeSpan Timeout { get; init; } = TimeSpan.FromSeconds(120);

    /// <summary>How long to wait for a free conversion slot before giving up. Shares the same
    /// process-wide cap as docx-to-PDF conversion, so a mixed .doc/.docx workload cannot spawn
    /// twice the intended number of LibreOffice processes.</summary>
    public TimeSpan QueueTimeout { get; init; } = TimeSpan.FromSeconds(60);
}

/// <summary>
/// Reads legacy binary .doc (Word 97-2003, OLE Compound File Binary Format) files by converting them
/// to .docx first, then handing off to the normal OpenXml-based pipeline unmodified. DocumentFormat.
/// OpenXml only understands the OOXML ZIP package format — it cannot parse the binary format at all
/// — and no free .NET library parses it either (the few that exist are abandoned/incomplete). Rather
/// than adding a new, large, high-risk binary-format parser dependency, this reuses the LibreOffice
/// subprocess this codebase already depends on for docx→PDF conversion: <c>soffice --convert-to
/// docx</c> handles the legacy format conversion, and everything downstream is exactly the same
/// OpenXml pipeline .docx input already goes through.
/// </summary>
public sealed class LegacyDocConverter
{
    private readonly LegacyDocConversionOptions _options;

    public LegacyDocConverter(LegacyDocConversionOptions? options = null) => _options = options ?? new();

    public void ConvertToDocx(string docPath, string outputDocxPath) =>
        ConvertToDocxAsync(docPath, outputDocxPath).GetAwaiter().GetResult();

    public async Task ConvertToDocxAsync(string docPath, string outputDocxPath, CancellationToken cancellationToken = default)
    {
        if (!File.Exists(docPath))
            throw new FileNotFoundException("Input .doc file not found.", docPath);

        await Conversion.LibreOfficeRunner.ConvertAsync(
            new Conversion.LibreOfficeSettings(_options.ExecutablePath, _options.UseWslDistro, _options.Timeout, _options.QueueTimeout),
            docPath,
            targetExtension: "docx",
            finalOutputPath: outputDocxPath,
            cancellationToken).ConfigureAwait(false);
    }
}
