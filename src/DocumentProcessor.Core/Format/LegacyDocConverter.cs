using System.Diagnostics;

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

        var outputDir = Path.GetDirectoryName(Path.GetFullPath(outputDocxPath))
            ?? throw new ArgumentException("Output path has no directory component.", nameof(outputDocxPath));
        Directory.CreateDirectory(outputDir);

        var profileArg = _options.UseWslDistro is not null
            ? $"file:///tmp/docproc-lo-profile-{Guid.NewGuid():N}"
            : $"file:///{Path.Combine(Path.GetTempPath(), $"docproc-lo-profile-{Guid.NewGuid():N}").Replace('\\', '/')}";

        var (fileName, args) = BuildInvocation(docPath, outputDir, profileArg);

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
            throw new TimeoutException($"LibreOffice conversion of '{docPath}' did not complete within {_options.Timeout}.");
        }

        var stdOut = await stdOutTask;
        var stdErr = await stdErrTask;

        if (process.ExitCode != 0)
        {
            throw new InvalidOperationException(
                $"LibreOffice conversion failed (exit code {process.ExitCode}).\nstdout: {stdOut}\nstderr: {stdErr}");
        }

        var producedPath = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(docPath) + ".docx");
        if (!File.Exists(producedPath))
        {
            throw new InvalidOperationException(
                $"LibreOffice reported success but no .docx was produced at '{producedPath}'.\nstdout: {stdOut}\nstderr: {stdErr}");
        }

        if (!string.Equals(Path.GetFullPath(producedPath), Path.GetFullPath(outputDocxPath), StringComparison.OrdinalIgnoreCase))
            File.Move(producedPath, outputDocxPath, overwrite: true);
    }

    private (string fileName, List<string> args) BuildInvocation(string docPath, string outputDir, string profileArg)
    {
        var args = new List<string>();
        string fileName;
        string effectiveDocPath;
        string effectiveOutDir;

        if (_options.UseWslDistro is { } distro)
        {
            fileName = "wsl.exe";
            args.AddRange(["-d", distro, "--", "soffice"]);
            effectiveDocPath = ToWslPath(docPath);
            effectiveOutDir = ToWslPath(outputDir);
        }
        else
        {
            fileName = _options.ExecutablePath;
            effectiveDocPath = docPath;
            effectiveOutDir = outputDir;
        }

        args.AddRange([
            "--headless",
            "--norestore",
            $"-env:UserInstallation={profileArg}",
            "--convert-to", "docx",
            "--outdir", effectiveOutDir,
            effectiveDocPath
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
