using System.Diagnostics;
using System.Text;

namespace DocumentProcessor.Core.Conversion;

/// <summary>
/// A unique working directory that deletes itself, used to give one LibreOffice invocation its own
/// user profile and its own output directory. Both matter: sharing a profile makes concurrent
/// conversions hang, and sharing an output directory lets two conversions of files with the same
/// base name (<c>contract.docx</c> — i.e. whatever the customer happened to name their upload)
/// race on the same produced filename and hand one caller the other's document.
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

    public static async Task<IDisposable> EnterAsync(TimeSpan queueTimeout, CancellationToken cancellationToken)
    {
        var semaphore = Semaphore.Value;
        if (!await semaphore.WaitAsync(queueTimeout, cancellationToken).ConfigureAwait(false))
        {
            throw new TimeoutException(
                $"Timed out after {queueTimeout} waiting for a conversion slot ({Limit} concurrent conversions allowed). " +
                "The host is saturated — shed load or raise DOCPROC_MAX_CONCURRENT_CONVERSIONS.");
        }

        return new Release(semaphore);
    }

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
internal sealed record LibreOfficeSettings(string ExecutablePath, string? UseWslDistro, TimeSpan Timeout, TimeSpan QueueTimeout);

/// <summary>
/// The private user profile directory one LibreOffice invocation runs against, and its cleanup.
/// <para>
/// Where the profile lives matters a great deal on the WSL path. LibreOffice populates a profile
/// with hundreds of small files, and doing that across the <c>/mnt/c</c> 9p mount takes roughly a
/// minute per conversion — measured, not assumed. So under WSL the profile stays on the distro's
/// own filesystem, which means the host cannot delete it with <see cref="Directory.Delete(string, bool)"/> and
/// has to shell back in to remove it. That extra spawn costs milliseconds and buys back ~59 seconds.
/// </para>
/// </summary>
internal sealed class LibreOfficeProfile : IDisposable
{
    private readonly string? _wslDistro;
    private readonly string _cleanupPath;

    private LibreOfficeProfile(string pathForLibreOffice, string cleanupPath, string? wslDistro)
    {
        PathForLibreOffice = pathForLibreOffice;
        _cleanupPath = cleanupPath;
        _wslDistro = wslDistro;
    }

    /// <summary>The profile path as LibreOffice itself must see it.</summary>
    public string PathForLibreOffice { get; }

    public static LibreOfficeProfile Create(LibreOfficeSettings settings, TempScratchDirectory scratch)
    {
        if (settings.UseWslDistro is { } distro)
        {
            var linuxPath = $"/tmp/docproc-lo-profile-{Guid.NewGuid():N}";
            return new LibreOfficeProfile(linuxPath, linuxPath, distro);
        }

        // Native: the profile can live inside the scratch directory, which is removed wholesale.
        var profileDir = Path.Combine(scratch.Path, "profile");
        Directory.CreateDirectory(profileDir);
        return new LibreOfficeProfile(profileDir, profileDir, wslDistro: null);
    }

    public void Dispose()
    {
        if (_wslDistro is null)
            return; // the scratch directory's own cleanup covers it

        try
        {
            using var process = System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = "wsl.exe",
                ArgumentList = { "-d", _wslDistro, "--", "rm", "-rf", _cleanupPath },
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
            });

            // Bounded: cleanup must never be able to hang or fail the caller's operation.
            process?.WaitForExit(milliseconds: 15_000);
        }
        catch
        {
            // Best-effort. WSL profiles live in the distro's /tmp, which is cleared on restart.
        }
    }
}

/// <summary>
/// Runs one <c>soffice --convert-to</c> invocation to completion. Shared by every converter so the
/// process-lifetime rules — concurrency cap, guaranteed child termination, scratch-directory
/// cleanup, bounded output capture — are implemented and fixed in exactly one place.
/// </summary>
internal static class LibreOfficeRunner
{
    /// <summary>Cap on captured stdout/stderr. A looping or chatty LibreOffice (which a hostile
    /// document can provoke) would otherwise grow an unbounded string per request, and that string
    /// goes into exception messages and therefore into logs.</summary>
    private const int MaxCapturedChars = 16 * 1024;

    public static async Task ConvertAsync(
        LibreOfficeSettings settings,
        string inputPath,
        string targetExtension,
        string finalOutputPath,
        CancellationToken cancellationToken)
    {
        var finalOutputDir = Path.GetDirectoryName(Path.GetFullPath(finalOutputPath))
            ?? throw new ArgumentException("Output path has no directory component.", nameof(finalOutputPath));
        Directory.CreateDirectory(finalOutputDir);

        using var gate = await LibreOfficeGate.EnterAsync(settings.QueueTimeout, cancellationToken).ConfigureAwait(false);
        using var scratch = new TempScratchDirectory();
        using var profile = LibreOfficeProfile.Create(settings, scratch);

        var convertOutDir = Path.Combine(scratch.Path, "out");
        Directory.CreateDirectory(convertOutDir);

        var (fileName, args) = BuildInvocation(settings, inputPath, convertOutDir, profile, targetExtension);

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

        using var process = Process.Start(psi)
            ?? throw new InvalidOperationException($"Failed to start conversion process '{fileName}'.");

        // Read both pipes concurrently with the wait below — draining them only after the process
        // exits would deadlock once a pipe buffer fills.
        var stdOutTask = ReadCappedAsync(process.StandardOutput, CancellationToken.None);
        var stdErrTask = ReadCappedAsync(process.StandardError, CancellationToken.None);

        using var timeoutCts = new CancellationTokenSource(settings.Timeout);
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
                $"LibreOffice conversion of '{inputPath}' did not complete within {settings.Timeout}.");
        }
        finally
        {
            // Backstop for any other escape between Process.Start and here.
            TryKill(process);
        }

        var stdOut = await stdOutTask.ConfigureAwait(false);
        var stdErr = await stdErrTask.ConfigureAwait(false);

        if (process.ExitCode != 0)
        {
            throw new InvalidOperationException(
                $"LibreOffice conversion failed (exit code {process.ExitCode}).\nstdout: {stdOut}\nstderr: {stdErr}");
        }

        // soffice names its output after the input's base name, inside --outdir. Because that
        // directory is unique per invocation, no other conversion can collide with it.
        var producedPath = Path.Combine(convertOutDir, Path.GetFileNameWithoutExtension(inputPath) + "." + targetExtension);
        if (!File.Exists(producedPath))
        {
            throw new InvalidOperationException(
                $"LibreOffice reported success but no {targetExtension} was produced at '{producedPath}'.\nstdout: {stdOut}\nstderr: {stdErr}");
        }

        File.Move(producedPath, finalOutputPath, overwrite: true);
    }

    private static (string fileName, List<string> args) BuildInvocation(
        LibreOfficeSettings settings, string inputPath, string convertOutDir, LibreOfficeProfile profile, string targetExtension)
    {
        var args = new List<string>();
        string fileName;
        string effectiveInputPath;
        string effectiveOutDir;

        if (settings.UseWslDistro is { } distro)
        {
            fileName = "wsl.exe";
            args.AddRange(["-d", distro, "--", "soffice"]);
            effectiveInputPath = ToWslPath(inputPath);
            effectiveOutDir = ToWslPath(convertOutDir);
        }
        else
        {
            fileName = settings.ExecutablePath;
            effectiveInputPath = inputPath;
            effectiveOutDir = convertOutDir;
        }

        args.AddRange([
            "--headless",
            "--norestore",
            $"-env:UserInstallation=file:///{profile.PathForLibreOffice.Replace('\\', '/').TrimStart('/')}",
            "--convert-to", targetExtension,
            "--outdir", effectiveOutDir,
            effectiveInputPath
        ]);

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
