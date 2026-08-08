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

        var outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPdfPath))
            ?? throw new ArgumentException("Output path has no directory component.", nameof(outputPdfPath));
        Directory.CreateDirectory(outputDir);

        // Isolating each conversion in its own LibreOffice user profile is what makes concurrent
        // server-side conversions safe — sharing the default profile causes conversions to hang
        // or fail under concurrency.
        var profileArg = _options.UseWslDistro is not null
            ? $"file:///tmp/docproc-lo-profile-{Guid.NewGuid():N}"
            : $"file:///{Path.Combine(Path.GetTempPath(), $"docproc-lo-profile-{Guid.NewGuid():N}").Replace('\\', '/')}";

        var (fileName, args) = BuildInvocation(docxPath, outputDir, profileArg);

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

        var stdOutTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
        var stdErrTask = process.StandardError.ReadToEndAsync(cancellationToken);

        using var timeoutCts = new CancellationTokenSource(_options.Timeout);
        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, timeoutCts.Token);
        try
        {
            await process.WaitForExitAsync(linkedCts.Token);
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            TryKill(process);
            throw new TimeoutException(
                $"LibreOffice conversion of '{docxPath}' did not complete within {_options.Timeout}.");
        }

        var stdOut = await stdOutTask;
        var stdErr = await stdErrTask;

        if (process.ExitCode != 0)
        {
            throw new InvalidOperationException(
                $"LibreOffice conversion failed (exit code {process.ExitCode}).\nstdout: {stdOut}\nstderr: {stdErr}");
        }

        // soffice names the output after the input's base name, inside --outdir.
        var producedPath = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(docxPath) + ".pdf");
        if (!File.Exists(producedPath))
        {
            throw new InvalidOperationException(
                $"LibreOffice reported success but no PDF was produced at '{producedPath}'.\nstdout: {stdOut}\nstderr: {stdErr}");
        }

        if (!string.Equals(Path.GetFullPath(producedPath), Path.GetFullPath(outputPdfPath), StringComparison.OrdinalIgnoreCase))
            File.Move(producedPath, outputPdfPath, overwrite: true);
    }

    private (string fileName, List<string> args) BuildInvocation(string docxPath, string outputDir, string profileArg)
    {
        var args = new List<string>();
        string fileName;
        string effectiveDocxPath;
        string effectiveOutDir;

        if (_options.UseWslDistro is { } distro)
        {
            fileName = "wsl.exe";
            args.AddRange(["-d", distro, "--", "soffice"]);
            effectiveDocxPath = ToWslPath(docxPath);
            effectiveOutDir = ToWslPath(outputDir);
        }
        else
        {
            fileName = _options.ExecutablePath;
            effectiveDocxPath = docxPath;
            effectiveOutDir = outputDir;
        }

        args.AddRange([
            "--headless",
            "--norestore",
            $"-env:UserInstallation={profileArg}",
            "--convert-to", "pdf",
            "--outdir", effectiveOutDir,
            effectiveDocxPath
        ]);

        return (fileName, args);
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
        try { process.Kill(entireProcessTree: true); }
        catch { /* best-effort: process may have already exited */ }
    }
}
