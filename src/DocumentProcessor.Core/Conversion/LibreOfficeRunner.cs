using System.Diagnostics;
using System.Text;

namespace DocumentProcessor.Core.Conversion;

/// <summary>
/// A unique working directory that deletes itself, used to give one LibreOffice invocation its own
/// output directory. That isolation matters: sharing an output directory lets two conversions of
/// files with the same base name (<c>contract.docx</c> — i.e. whatever the customer happened to
/// name their upload) race on the same produced filename and hand one caller the other's document.
/// </summary>
internal sealed class TempScratchDirectory : IDisposable
{
    public string Path { get; }

    public TempScratchDirectory()
    {
        Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"docproc-lo-{Guid.NewGuid():N}");
        Directory.CreateDirectory(Path);
    }

    public void Dispose()
    {
        // Best-effort: a failure to clean up must never mask the caller's real result or error.
        try { Directory.Delete(Path, recursive: true); }
        catch { /* the OS temp reaper is the backstop */ }
    }
}

/// <summary>
/// Caps how many LibreOffice processes this process will run at once.
/// <para>
/// Process-wide static state is the correct scope here, unlike tenant data: the constraint being
/// modelled is "how many ~300 MB subprocesses fit on this machine", which is a property of the
/// host, not of the caller. Without it, N concurrent requests spawn N LibreOffice processes and
/// the container is OOM-killed.
/// </para>
/// Tunable via <c>DOCPROC_MAX_CONCURRENT_CONVERSIONS</c>; defaults to half the CPU count, which
/// leaves headroom for the managed work (OOXML parsing) happening alongside it.
/// </summary>
internal static class LibreOfficeGate
{
    private static readonly Lazy<SemaphoreSlim> Semaphore = new(() => new SemaphoreSlim(ResolveLimit()));

    public static int Limit => ResolveLimit();

    private static int ResolveLimit()
    {
        var configured = Environment.GetEnvironmentVariable("DOCPROC_MAX_CONCURRENT_CONVERSIONS");
        if (int.TryParse(configured, out var parsed) && parsed > 0)
            return parsed;

        return Math.Max(1, Environment.ProcessorCount / 2);
    }

    public static async Task<IDisposable?> TryEnterAsync(TimeSpan queueTimeout, CancellationToken cancellationToken)
    {
        var semaphore = Semaphore.Value;
        return await semaphore.WaitAsync(queueTimeout, cancellationToken).ConfigureAwait(false)
            ? new Release(semaphore)
            : null;
    }

    public static async Task<IDisposable> EnterAsync(TimeSpan queueTimeout, CancellationToken cancellationToken) =>
        await TryEnterAsync(queueTimeout, cancellationToken).ConfigureAwait(false)
        ?? throw SaturatedException(queueTimeout);

    public static TimeoutException SaturatedException(TimeSpan queueTimeout) =>
        new($"Timed out after {queueTimeout} waiting for a conversion slot ({Limit} concurrent conversions allowed). " +
            "The host is saturated — shed load or raise DOCPROC_MAX_CONCURRENT_CONVERSIONS.");

    private sealed class Release(SemaphoreSlim semaphore) : IDisposable
    {
        private int _disposed;
        public void Dispose()
        {
            if (Interlocked.Exchange(ref _disposed, 1) == 0)
                semaphore.Release();
        }
    }
}

/// <summary>How a LibreOffice invocation should be located and bounded.</summary>
/// <param name="ExecutablePath">The soffice binary, when not routing through WSL.</param>
/// <param name="UseWslDistro">WSL distro to route through, for Windows development.</param>
/// <param name="Timeout">Budget for converting one document. A batch of N gets N times this.</param>
/// <param name="QueueTimeout">How long a request waits for a conversion slot before shedding.</param>
/// <param name="ReuseProfiles">Whether warm user profiles are pooled across invocations.</param>
/// <param name="MaxProfileReuses">How many invocations a pooled profile serves before it is rebuilt.</param>
/// <param name="MaxBatchSize">How many queued documents may share one LibreOffice invocation.</param>
/// <param name="EnableBatching">Whether concurrent requests are coalesced into shared invocations.</param>
internal sealed record LibreOfficeSettings(
    string ExecutablePath,
    string? UseWslDistro,
    TimeSpan Timeout,
    TimeSpan QueueTimeout,
    bool ReuseProfiles = true,
    int MaxProfileReuses = 50,
    int MaxBatchSize = 8,
    bool EnableBatching = true);

/// <summary>One document to convert, and where its result belongs.</summary>
internal readonly record struct ConversionItem(string InputPath, string FinalOutputPath);

/// <summary>What became of one document in a batch. <see cref="Error"/> is null on success.</summary>
internal readonly record struct ConversionOutcome(ConversionItem Item, Exception? Error);

/// <summary>
/// Runs <c>soffice --convert-to</c> invocations. Shared by every converter so the process-lifetime
/// rules — concurrency cap, guaranteed child termination, scratch-directory cleanup, bounded output
/// capture — are implemented and fixed in exactly one place.
/// <para>
/// One invocation can convert many documents. That matters more than it sounds: LibreOffice spends
/// roughly 450 ms starting up before it looks at the first document, and nothing after the first
/// costs that again. Measured on eight small documents, one invocation took 715 ms against 3,262 ms
/// for eight sequential invocations, and still beat running eight of them in parallel (1,000 ms)
/// while occupying one process instead of eight.
/// </para>
/// </summary>
internal static class LibreOfficeRunner
{
    /// <summary>Cap on captured stdout/stderr. A looping or chatty LibreOffice (which a hostile
    /// document can provoke) would otherwise grow an unbounded string per request, and that string
    /// goes into exception messages and therefore into logs.</summary>
    private const int MaxCapturedChars = 16 * 1024;

    /// <summary>Converts a single document, throwing on failure.</summary>
    public static async Task ConvertAsync(
        LibreOfficeSettings settings,
        string inputPath,
        string targetExtension,
        string finalOutputPath,
        CancellationToken cancellationToken)
    {
        using var gate = await LibreOfficeGate.EnterAsync(settings.QueueTimeout, cancellationToken).ConfigureAwait(false);

        var outcomes = await ConvertBatchAsync(
            settings, [new ConversionItem(inputPath, finalOutputPath)], targetExtension, cancellationToken).ConfigureAwait(false);

        if (outcomes[0].Error is { } error)
            throw error;
    }

    /// <summary>
    /// Converts several documents in one LibreOffice invocation.
    /// <para>
    /// The caller is responsible for holding a <see cref="LibreOfficeGate"/> slot: a batch occupies
    /// exactly one LibreOffice process regardless of its size, and the gate counts processes.
    /// </para>
    /// <para>
    /// Failures are reported per document rather than thrown, because one unreadable file in a
    /// batch must not deny the other callers their results. Faults that apply to the whole
    /// invocation — LibreOffice missing, the process timing out or being cancelled — still throw.
    /// </para>
    /// </summary>
    public static async Task<IReadOnlyList<ConversionOutcome>> ConvertBatchAsync(
        LibreOfficeSettings settings,
        IReadOnlyList<ConversionItem> items,
        string targetExtension,
        CancellationToken cancellationToken)
    {
        if (items.Count == 0)
            return [];

        foreach (var item in items)
        {
            var finalOutputDir = Path.GetDirectoryName(Path.GetFullPath(item.FinalOutputPath))
                ?? throw new ArgumentException("Output path has no directory component.", nameof(items));
            Directory.CreateDirectory(finalOutputDir);
        }

        using var scratch = new TempScratchDirectory();
        var profile = LibreOfficeProfilePool.Lease(settings);
        var profileIsClean = false;
        try
        {
            var convertOutDir = Path.Combine(scratch.Path, "out");
            Directory.CreateDirectory(convertOutDir);

            var staged = StageInputs(items, scratch, targetExtension);
            var (fileName, args) = BuildInvocation(
                settings, staged.Select(s => s.EffectiveInputPath).ToList(), convertOutDir, profile, targetExtension);

            var (exitCode, stdOut, stdErr) = await RunProcessAsync(settings, fileName, args, items.Count, cancellationToken)
                .ConfigureAwait(false);

            // A clean exit means LibreOffice released its profile lock, so the profile is reusable.
            // Reaching here at all means the process was not killed.
            profileIsClean = true;

            var outcomes = CollectResults(staged, convertOutDir, targetExtension, exitCode, stdOut, stdErr);

            // A non-zero exit does not condemn the batch. LibreOffice reports failure for the
            // invocation when any one document fails to load, having already converted the rest —
            // so the produced files are the authority on what succeeded, and the exit code only
            // decides how a total failure is reported.
            if (exitCode != 0 && outcomes.All(o => o.Error is not null))
            {
                throw new InvalidOperationException(
                    $"LibreOffice conversion failed (exit code {exitCode}).\nstdout: {stdOut}\nstderr: {stdErr}");
            }

            return outcomes;
        }
        finally
        {
            if (profileIsClean)
                LibreOfficeProfilePool.Return(profile, settings);
            else
                LibreOfficeProfilePool.Discard(profile);
        }
    }

    /// <summary>One document's place in a batch: where LibreOffice reads it, and where the result goes.</summary>
    private readonly record struct StagedItem(ConversionItem Item, string EffectiveInputPath, string ProducedName);

    /// <summary>
    /// LibreOffice names each output after its input's base name, so two documents in one batch
    /// that are both called <c>contract.docx</c> would produce one file and silently hand two
    /// callers the same PDF. Colliding inputs are copied into the scratch directory under distinct
    /// names; non-colliding ones — the overwhelmingly common case, and every single-document
    /// conversion — are read where they lie, so batching costs no extra I/O for them.
    /// </summary>
    private static List<StagedItem> StageInputs(IReadOnlyList<ConversionItem> items, TempScratchDirectory scratch, string targetExtension)
    {
        var staged = new List<StagedItem>(items.Count);
        var takenNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        string? stageDir = null;

        for (var i = 0; i < items.Count; i++)
        {
            var item = items[i];
            var baseName = Path.GetFileNameWithoutExtension(item.InputPath);

            if (takenNames.Add(baseName))
            {
                staged.Add(new StagedItem(item, item.InputPath, baseName + "." + targetExtension));
                continue;
            }

            stageDir ??= Directory.CreateDirectory(Path.Combine(scratch.Path, "in")).FullName;
            var uniqueName = $"{baseName}-{i}";
            var stagedPath = Path.Combine(stageDir, uniqueName + Path.GetExtension(item.InputPath));
            File.Copy(item.InputPath, stagedPath, overwrite: true);
            takenNames.Add(uniqueName);
            staged.Add(new StagedItem(item, stagedPath, uniqueName + "." + targetExtension));
        }

        return staged;
    }

    private static IReadOnlyList<ConversionOutcome> CollectResults(
        List<StagedItem> staged, string convertOutDir, string targetExtension, int exitCode, string stdOut, string stdErr)
    {
        var outcomes = new List<ConversionOutcome>(staged.Count);

        foreach (var entry in staged)
        {
            var producedPath = Path.Combine(convertOutDir, entry.ProducedName);
            if (!File.Exists(producedPath))
            {
                outcomes.Add(new ConversionOutcome(entry.Item, new InvalidOperationException(
                    $"LibreOffice produced no {targetExtension} for '{entry.Item.InputPath}' (exit code {exitCode}).\n" +
                    $"stdout: {stdOut}\nstderr: {stdErr}")));
                continue;
            }

            try
            {
                File.Move(producedPath, entry.Item.FinalOutputPath, overwrite: true);
                outcomes.Add(new ConversionOutcome(entry.Item, null));
            }
            catch (Exception ex)
            {
                outcomes.Add(new ConversionOutcome(entry.Item, ex));
            }
        }

        return outcomes;
    }

    private static async Task<(int ExitCode, string StdOut, string StdErr)> RunProcessAsync(
        LibreOfficeSettings settings, string fileName, List<string> args, int documentCount, CancellationToken cancellationToken)
    {
        var psi = new ProcessStartInfo
        {
            FileName = fileName,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
        };
        foreach (var arg in args)
            psi.ArgumentList.Add(arg);

        Process? started;
        try
        {
            started = Process.Start(psi);
        }
        catch (Exception ex)
        {
            // Almost always "LibreOffice isn't installed here" or a wrong ExecutablePath. That is a
            // deployment fault, not a transient one: it will fail identically on every retry, so it
            // gets its own non-retryable type rather than surfacing as a raw Win32Exception.
            throw new ConversionUnavailableException(
                $"Could not start the document converter '{fileName}'. Check that LibreOffice is installed and that " +
                "ExecutablePath (or UseWslDistro) points at it.", ex);
        }

        using var process = started
            ?? throw new ConversionUnavailableException($"Could not start the document converter '{fileName}'.");

        // Read both pipes concurrently with the wait below — draining them only after the process
        // exits would deadlock once a pipe buffer fills.
        var stdOutTask = ReadCappedAsync(process.StandardOutput, CancellationToken.None);
        var stdErrTask = ReadCappedAsync(process.StandardError, CancellationToken.None);

        // Timeout is the budget for one document, so a batch is allowed proportionally longer.
        var budget = documentCount == 1 ? settings.Timeout : settings.Timeout * documentCount;

        using var timeoutCts = new CancellationTokenSource(budget);
        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, timeoutCts.Token);
        try
        {
            await process.WaitForExitAsync(linkedCts.Token).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            // Kill on BOTH cancellation paths. Previously this ran only for the internal timeout,
            // so a caller cancelling (in ASP.NET Core, any client disconnect) left soffice running
            // forever — Process.Dispose releases the handle but never terminates the child.
            TryKill(process);

            if (cancellationToken.IsCancellationRequested)
                throw;

            throw new TimeoutException(
                $"LibreOffice conversion did not complete within {budget} for {documentCount} document(s).");
        }
        finally
        {
            // Backstop for any other escape between Process.Start and here.
            TryKill(process);
        }

        return (process.ExitCode, await stdOutTask.ConfigureAwait(false), await stdErrTask.ConfigureAwait(false));
    }

    private static (string fileName, List<string> args) BuildInvocation(
        LibreOfficeSettings settings, IReadOnlyList<string> inputPaths, string convertOutDir, PooledProfile profile, string targetExtension)
    {
        var args = new List<string>();
        string fileName;
        string effectiveOutDir;
        List<string> effectiveInputPaths;

        if (settings.UseWslDistro is { } distro)
        {
            fileName = "wsl.exe";
            args.AddRange(["-d", distro, "--", "soffice"]);
            effectiveInputPaths = inputPaths.Select(ToWslPath).ToList();
            effectiveOutDir = ToWslPath(convertOutDir);
        }
        else
        {
            fileName = settings.ExecutablePath;
            effectiveInputPaths = [.. inputPaths];
            effectiveOutDir = convertOutDir;
        }

        args.AddRange([
            "--headless",
            "--norestore",
            $"-env:UserInstallation=file:///{profile.PathForLibreOffice.Replace('\\', '/').TrimStart('/')}",
            "--convert-to", targetExtension,
            "--outdir", effectiveOutDir
        ]);
        args.AddRange(effectiveInputPaths);

        return (fileName, args);
    }

    private static async Task<string> ReadCappedAsync(StreamReader reader, CancellationToken cancellationToken)
    {
        var buffer = new char[8192];
        var builder = new StringBuilder();
        int read;

        // Keep draining after the cap is reached — stopping early would fill the pipe buffer and
        // block the child — but stop growing the string.
        while ((read = await reader.ReadAsync(buffer.AsMemory(), cancellationToken).ConfigureAwait(false)) > 0)
        {
            if (builder.Length >= MaxCapturedChars)
                continue;

            builder.Append(buffer, 0, Math.Min(read, MaxCapturedChars - builder.Length));
            if (builder.Length >= MaxCapturedChars)
                builder.Append("… (truncated)");
        }

        return builder.ToString();
    }

    private static string ToWslPath(string windowsPath)
    {
        var full = Path.GetFullPath(windowsPath);
        var drive = char.ToLowerInvariant(full[0]);
        var rest = full[2..].Replace('\\', '/');
        return $"/mnt/{drive}{rest}";
    }

    private static void TryKill(Process process)
    {
        try
        {
            if (!process.HasExited)
                process.Kill(entireProcessTree: true);
        }
        catch { /* already exited, or exited between the check and the kill */ }
    }
}
