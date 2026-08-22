namespace DocumentProcessor.Core.Watermarking;

/// <summary>Watermark appearance for one document status. The docx- and PDF-specific fields
/// (<see cref="ColorHex"/>/<see cref="Removable"/> vs <see cref="GrayLevel"/>/<see cref="Alpha"/>)
/// exist together so one config maps a status to both output formats consistently.</summary>
public sealed record WatermarkConfig(
    string Text,
    string FontFamily = "Calibri",
    int RotationDegrees = -45,
    WatermarkPosition Position = WatermarkPosition.Center,
    double FontSizePt = 72,
    string ColorHex = "C0C0C0",
    bool Removable = true,
    double WidthPt = 415,
    double HeightPt = 207.5,
    byte GrayLevel = 192,
    byte Alpha = 100);

/// <summary>
/// Maps a document's status (e.g. "Draft", "Confidential", "Final") to the watermark it should
/// carry — the policy layer the primitive watermark services don't have on their own. A status
/// mapped to <see langword="null"/> (or simply not mapped) means "no watermark for this status,"
/// e.g. a finalized document.
/// </summary>
public sealed class StatusWatermarkPolicy
{
    private readonly Dictionary<string, WatermarkConfig?> _map = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Maps <paramref name="status"/> to a watermark configuration. Fluent, for chaining multiple mappings.</summary>
    public StatusWatermarkPolicy Map(string status, WatermarkConfig config)
    {
        _map[status] = config;
        return this;
    }

    /// <summary>Explicitly maps <paramref name="status"/> to "no watermark" — distinct from simply
    /// not calling <see cref="Map"/>, so a policy can document that a status was considered and
    /// deliberately left unwatermarked.</summary>
    public StatusWatermarkPolicy MapNone(string status)
    {
        _map[status] = null;
        return this;
    }

    public WatermarkConfig? Resolve(string status) => _map.TryGetValue(status, out var config) ? config : null;

    /// <summary>A reasonable starting policy: Draft gets a removable gray "DRAFT" watermark,
    /// Confidential gets a non-removable red one, Final gets none.</summary>
    public static StatusWatermarkPolicy CreateDefault() => new StatusWatermarkPolicy()
        .Map("Draft", new WatermarkConfig("DRAFT"))
        .Map("Confidential", new WatermarkConfig("CONFIDENTIAL", ColorHex: "C00000", Removable: false, GrayLevel: 200, Alpha: 140))
        .MapNone("Final");

    /// <summary>Applies (or clears) the watermark for <paramref name="status"/> on a .docx file.
    /// A status resolving to no config removes any existing policy-managed watermark instead of
    /// leaving a stale one from a prior status.</summary>
    public void ApplyToDocx(string docxPath, string status, DocxWatermarkService? service = null)
    {
        service ??= new DocxWatermarkService();
        var config = Resolve(status);
        if (config is null)
        {
            service.RemoveWatermark(docxPath);
            return;
        }

        service.AddTextWatermark(
            docxPath, config.Text, config.FontFamily, config.RotationDegrees, config.ColorHex,
            config.Removable, config.Position, config.WidthPt, config.HeightPt, config.FontSizePt);
    }

    /// <summary>Applies the watermark for <paramref name="status"/> to a PDF, writing to
    /// <paramref name="outputPath"/>. A status resolving to no config copies the input through
    /// unwatermarked.</summary>
    public void ApplyToPdf(string pdfPath, string outputPath, string status, PdfWatermarkService? service = null)
    {
        var config = Resolve(status);
        if (config is null)
        {
            File.Copy(pdfPath, outputPath, overwrite: true);
            return;
        }

        (service ?? new PdfWatermarkService()).AddTextWatermark(
            pdfPath, outputPath, config.Text, config.FontFamily, config.RotationDegrees,
            config.GrayLevel, config.Alpha, config.Position, config.FontSizePt);
    }
}
