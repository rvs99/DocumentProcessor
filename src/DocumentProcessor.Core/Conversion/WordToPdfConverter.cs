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
    public string ExecutablePath { get; set; } = OperatingSystem.IsWindows()
        ? @"C:\Program Files\LibreOffice\program\soffice.exe"
        : "/usr/bin/soffice";

    /// <summary>
    /// Windows-development convenience: when set, conversions are routed through
    /// `wsl.exe -d {distro} -- soffice ...` instead of a native executable, so LibreOffice doesn't
    /// need to be installed as a Windows desktop app. Not used in production container deployments,
    /// where <see cref="ExecutablePath"/> points straight at the Linux soffice binary.
    /// </summary>
    public string? UseWslDistro { get; set; }

    /// <summary>
    /// Budget for converting one document. A batch of N documents is allowed N times this, since
    /// most of a conversion is per-document work.
    /// </summary>
    public TimeSpan Timeout { get; set; } = TimeSpan.FromSeconds(120);

    /// <summary>
    /// How long to wait for a free conversion slot before giving up. Conversions are capped
    /// process-wide (see <c>DOCPROC_MAX_CONCURRENT_CONVERSIONS</c>) so a burst of requests can't
    /// spawn an unbounded number of LibreOffice processes. Failing fast here is deliberate: under
    /// saturation it is better to shed load than to queue every request until it times out anyway.
    /// </summary>
    public TimeSpan QueueTimeout { get; set; } = TimeSpan.FromSeconds(60);

    /// <summary>
    /// Whether concurrent conversions waiting on a slot are coalesced into shared LibreOffice
    /// invocations. On by default, and free when the host is idle: nothing is ever held back
    /// waiting for a companion request, so batching only happens among documents that were queued
    /// anyway. Turn it off to have every conversion run in its own process — useful when isolating
    /// a document suspected of destabilising LibreOffice.
    /// </summary>
    public bool EnableBatching { get; set; } = true;

    /// <summary>
    /// The most documents that may share one LibreOffice invocation. Larger batches amortise
    /// startup further but widen the blast radius: a single wedged document takes its whole batch
    /// down with it.
    /// </summary>
    public int MaxBatchSize { get; set; } = 8;

    /// <summary>
    /// Whether LibreOffice user profiles are kept warm and reused across conversions rather than
    /// rebuilt each time. Building one costs roughly 80–100 ms. Profiles are still leased
    /// exclusively — two LibreOffice processes sharing a profile contend for its lock and hang —
    /// and any profile whose process had to be killed is destroyed rather than reused.
    /// </summary>
    public bool ReuseProfiles { get; set; } = true;

    /// <summary>How many conversions a pooled profile serves before being rebuilt, which bounds
    /// the state it accumulates and how long a subtly corrupted profile can survive.</summary>
    public int MaxProfileReuses { get; set; } = 50;

    internal LibreOfficeSettings ToSettings() => new(
        ExecutablePath, UseWslDistro, Timeout, QueueTimeout,
        ReuseProfiles, MaxProfileReuses, MaxBatchSize, EnableBatching);
}

/// <summary>One document to convert, and where its PDF belongs.</summary>
public sealed record ConversionRequest(string DocxPath, string OutputPdfPath);

/// <summary>What became of one document in a batch conversion.</summary>
/// <param name="Request">The document this result belongs to.</param>
/// <param name="Error">Why it failed, or null if it succeeded.</param>
public sealed record ConversionResult(ConversionRequest Request, Exception? Error)
{
    public bool Succeeded => Error is null;
}

/// <summary>
/// Converts .docx files to PDF via LibreOffice headless — the only free, high-fidelity docx→PDF
/// path (no free .NET-native renderer matches its layout/font accuracy).
/// <para>
/// Conversions run as external processes, capped process-wide so a burst of traffic cannot spawn an
/// unbounded number of ~300 MB LibreOffice instances. Two things keep that subprocess from
/// dominating the request: warm user profiles are pooled rather than rebuilt per call, and
/// conversions that end up queued behind a busy host are coalesced into shared invocations, so
/// LibreOffice's ~450 ms startup is paid once for the group instead of once per document.
/// </para>
/// </summary>
public sealed class WordToPdfConverter : IWordToPdfConverter
{
    private readonly WordToPdfConversionOptions _options;

    public WordToPdfConverter(WordToPdfConversionOptions? options = null) => _options = options ?? new();

    public void Convert(string docxPath, string outputPdfPath) =>
        ConvertAsync(docxPath, outputPdfPath).GetAwaiter().GetResult();

    public async Task ConvertAsync(string docxPath, string outputPdfPath, CancellationToken cancellationToken = default)
    {
        if (!File.Exists(docxPath))
            throw new FileNotFoundException("Input .docx file not found.", docxPath);

        var settings = _options.ToSettings();
        var item = new ConversionItem(docxPath, outputPdfPath);

        if (!settings.EnableBatching)
        {
            await LibreOfficeRunner.ConvertAsync(settings, docxPath, "pdf", outputPdfPath, cancellationToken).ConfigureAwait(false);
            return;
        }

        await ConversionQueue.EnqueueAsync(settings, item, "pdf", cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Converts several documents together, which is substantially cheaper than converting them one
    /// at a time: LibreOffice's startup is paid once for the whole set. Measured on eight small
    /// documents, 715 ms against 3,262 ms for eight separate conversions.
    /// <para>
    /// One document failing does not deny the others their results — inspect
    /// <see cref="ConversionResult.Succeeded"/> per request. Faults that apply to the whole
    /// invocation, such as LibreOffice not being installed, still throw.
    /// </para>
    /// </summary>
    public async Task<IReadOnlyList<ConversionResult>> ConvertBatchAsync(
        IReadOnlyList<ConversionRequest> requests, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(requests);
        if (requests.Count == 0)
            return [];

        foreach (var request in requests)
        {
            if (!File.Exists(request.DocxPath))
                throw new FileNotFoundException("Input .docx file not found.", request.DocxPath);
        }

        var settings = _options.ToSettings();

        // Chunked so an explicit batch of 500 documents still respects the blast-radius limit that
        // MaxBatchSize sets. Chunks run concurrently, but each takes a conversion slot first, so the
        // process-wide cap — not the size of the caller's request — decides how many LibreOffice
        // instances exist at once.
        var chunks = await Task.WhenAll(
            requests.Chunk(Math.Max(1, settings.MaxBatchSize)).Select(async chunk =>
            {
                using var slot = await LibreOfficeGate.EnterAsync(settings.QueueTimeout, cancellationToken).ConfigureAwait(false);

                var outcomes = await LibreOfficeRunner.ConvertBatchAsync(
                    settings,
                    [.. chunk.Select(r => new ConversionItem(r.DocxPath, r.OutputPdfPath))],
                    "pdf",
                    cancellationToken).ConfigureAwait(false);

                return chunk.Select((request, i) => new ConversionResult(request, outcomes[i].Error)).ToList();
            })).ConfigureAwait(false);

        // Results come back in request order regardless of which chunk finished first.
        return [.. chunks.SelectMany(c => c)];
    }
}
