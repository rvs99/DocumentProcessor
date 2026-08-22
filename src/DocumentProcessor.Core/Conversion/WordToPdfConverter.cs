using System.Diagnostics;

namespace DocumentProcessor.Core.Conversion;

/// <summary>
/// Configuration for the LibreOffice headless conversion process.
/// </summary>
public sealed class WordToPdfConversionOptions
{
    /// <summary>
    /// Path to the soffice executable. In a Linux container this is typically "/usr/bin/soffice"
    /// (installed via `apt-get install libreoffice-writer`); on a native Windows install it's
    /// "C:\Program Files\LibreOffice\program\soffice.exe". Ignored when <see cref="UseWslDistro"/> is set.
    /// </summary>
    public string ExecutablePath { get; init; } = OperatingSystem.IsWindows()
        ? @"C:\Program Files\LibreOffice\program\soffice.exe"
        : "/usr/bin/soffice";

    /// <summary>
    /// Windows-development convenience: when set, conversions are routed through
    /// `wsl.exe -d {distro} -- soffice ...` instead of a native executable, so LibreOffice doesn't
    /// need to be installed as a Windows desktop app. Not used in production container deployments,
    /// where <see cref="ExecutablePath"/> points straight at the Linux soffice binary.
    /// </summary>
    public string? UseWslDistro { get; init; }

    public TimeSpan Timeout { get; init; } = TimeSpan.FromSeconds(120);

    /// <summary>
    /// How long to wait for a free conversion slot before giving up. Conversions are capped
    /// process-wide (see <c>DOCPROC_MAX_CONCURRENT_CONVERSIONS</c>) so a burst of requests can't
    /// spawn an unbounded number of LibreOffice processes. Failing fast here is deliberate: under
    /// saturation it is better to shed load than to queue every request until it times out anyway.
    /// </summary>
    public TimeSpan QueueTimeout { get; init; } = TimeSpan.FromSeconds(60);
}

/// <summary>
/// Converts .docx files to PDF via LibreOffice headless — the only free, high-fidelity docx→PDF
/// path (no free .NET-native renderer matches its layout/font accuracy). Runs LibreOffice as an
/// external process with an isolated user profile per call, so concurrent conversions on the same
/// host don't contend for a shared LibreOffice profile lock.
/// </summary>
public sealed class WordToPdfConverter
{
    private readonly WordToPdfConversionOptions _options;

    public WordToPdfConverter(WordToPdfConversionOptions? options = null) => _options = options ?? new();

    public void Convert(string docxPath, string outputPdfPath) =>
        ConvertAsync(docxPath, outputPdfPath).GetAwaiter().GetResult();

    public async Task ConvertAsync(string docxPath, string outputPdfPath, CancellationToken cancellationToken = default)
    {
        if (!File.Exists(docxPath))
            throw new FileNotFoundException("Input .docx file not found.", docxPath);

        await LibreOfficeRunner.ConvertAsync(
            new LibreOfficeSettings(_options.ExecutablePath, _options.UseWslDistro, _options.Timeout, _options.QueueTimeout),
            docxPath,
            targetExtension: "pdf",
            finalOutputPath: outputPdfPath,
            cancellationToken).ConfigureAwait(false);
    }
}
