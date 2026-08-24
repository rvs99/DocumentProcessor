using System.Collections.Concurrent;
using System.Diagnostics;

namespace DocumentProcessor.Core.Conversion;

/// <summary>
/// One LibreOffice user profile directory, owned by <see cref="LibreOfficeProfilePool"/> rather
/// than by a single conversion.
/// <para>
/// Where the profile lives matters a great deal on the WSL path. LibreOffice populates a profile
/// with hundreds of small files, and doing that across the <c>/mnt/c</c> 9p mount takes roughly a
/// minute per conversion — measured, not assumed. So under WSL the profile stays on the distro's
/// own filesystem, which means the host cannot delete it with <see cref="Directory.Delete(string, bool)"/>
/// and has to shell back in to remove it. Pooling makes that spawn rare rather than per-conversion.
/// </para>
/// </summary>
internal sealed class PooledProfile
{
    private readonly string _cleanupPath;
    private readonly string? _wslDistro;

    private PooledProfile(string pathForLibreOffice, string cleanupPath, string? wslDistro)
    {
        PathForLibreOffice = pathForLibreOffice;
        _cleanupPath = cleanupPath;
        _wslDistro = wslDistro;
    }

    /// <summary>The profile path as LibreOffice itself must see it.</summary>
    public string PathForLibreOffice { get; }

    /// <summary>How many invocations have run against this profile. Drives recycling.</summary>
    public int UseCount { get; private set; }

    public void RecordUse() => UseCount++;

    public static PooledProfile Create(LibreOfficeSettings settings)
    {
        if (settings.UseWslDistro is not null)
        {
            // Created lazily by LibreOffice itself, inside the distro — nothing to mkdir from here.
            var linuxPath = $"/tmp/docproc-lo-profile-{Guid.NewGuid():N}";
            return new PooledProfile(linuxPath, linuxPath, settings.UseWslDistro);
        }

        var profileDir = Path.Combine(Path.GetTempPath(), $"docproc-lo-profile-{Guid.NewGuid():N}");
        Directory.CreateDirectory(profileDir);
        return new PooledProfile(profileDir, profileDir, wslDistro: null);
    }

    /// <summary>Removes the profile from disk. Best-effort: cleanup must never mask a real result.</summary>
    public void Delete()
    {
        try
        {
            if (_wslDistro is null)
            {
                if (Directory.Exists(_cleanupPath))
                    Directory.Delete(_cleanupPath, recursive: true);
                return;
            }

            using var process = Process.Start(new ProcessStartInfo
            {
                FileName = "wsl.exe",
                ArgumentList = { "-d", _wslDistro, "--", "rm", "-rf", _cleanupPath },
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
            });

            // Bounded: cleanup must never be able to hang the caller's operation.
            process?.WaitForExit(milliseconds: 15_000);
        }
        catch
        {
            // The OS temp reaper is the backstop; WSL profiles live in the distro's /tmp, which is
            // cleared on restart.
        }
    }
}

/// <summary>
/// Keeps warm LibreOffice user profiles alive across conversions instead of building a fresh one
/// per invocation.
/// <para>
/// Creating a profile costs roughly 80–100 ms of a ~550 ms conversion (measured on a Linux
/// LibreOffice 7 install) because LibreOffice writes several hundred small configuration files into
/// it. That work is identical every time, so paying it once per profile rather than once per
/// document removes it from the request path.
/// </para>
/// <para>
/// A leased profile is held exclusively for the duration of an invocation. That is not an
/// optimisation but a correctness requirement: two LibreOffice processes sharing one profile
/// contend for its lock file and hang, which is exactly why the original code built a throwaway
/// profile per call.
/// </para>
/// </summary>
internal static class LibreOfficeProfilePool
{
    /// <summary>
    /// Profiles are only interchangeable between invocations that run the same LibreOffice.
    /// A profile written by the WSL install would be handed to a native soffice of a different
    /// version, which migrates (or rejects) it on load.
    /// </summary>
    private static readonly ConcurrentDictionary<string, ConcurrentBag<PooledProfile>> Idle = new();

    static LibreOfficeProfilePool()
    {
        // Best-effort tidy-up on orderly shutdown. Profiles live under the temp directory precisely
        // so that a hard kill still leaves them to the OS reaper rather than leaking forever.
        AppDomain.CurrentDomain.ProcessExit += (_, _) => DrainAll();
    }

    private static string KeyFor(LibreOfficeSettings settings) =>
        $"{settings.UseWslDistro ?? "native"}{settings.ExecutablePath}";

    /// <summary>
    /// Takes a warm profile, or builds one if none is idle. The caller must hand it back with
    /// exactly one of <see cref="Return"/> or <see cref="Discard"/>.
    /// </summary>
    public static PooledProfile Lease(LibreOfficeSettings settings)
    {
        if (settings.ReuseProfiles && Idle.TryGetValue(KeyFor(settings), out var bag) && bag.TryTake(out var profile))
            return profile;

        return PooledProfile.Create(settings);
    }

    /// <summary>Returns a profile after a clean run, so the next conversion can reuse it.</summary>
    public static void Return(PooledProfile profile, LibreOfficeSettings settings)
    {
        profile.RecordUse();

        // A profile accumulates state — recent-document lists, caches — for as long as it lives.
        // Recycling bounds that drift, and bounds how long a subtly corrupted profile can keep
        // failing conversions before it is replaced.
        if (!settings.ReuseProfiles || profile.UseCount >= settings.MaxProfileReuses)
        {
            profile.Delete();
            return;
        }

        var bag = Idle.GetOrAdd(KeyFor(settings), _ => []);

        // No point holding more idle profiles than can ever be leased at once.
        if (bag.Count >= LibreOfficeGate.Limit)
        {
            profile.Delete();
            return;
        }

        bag.Add(profile);
    }

    /// <summary>
    /// Destroys a profile instead of pooling it. Used whenever the invocation did not exit
    /// cleanly: LibreOffice removes its lock file on a normal exit but not when killed, so a
    /// timed-out or cancelled conversion leaves behind a profile that would wedge every subsequent
    /// conversion that leased it.
    /// </summary>
    public static void Discard(PooledProfile profile) => profile.Delete();

    /// <summary>Empties the pool. Exposed for tests and for process shutdown.</summary>
    public static void DrainAll()
    {
        foreach (var bag in Idle.Values)
            while (bag.TryTake(out var profile))
                profile.Delete();
    }

    /// <summary>Idle profile count for a given configuration. Test seam.</summary>
    internal static int IdleCount(LibreOfficeSettings settings) =>
        Idle.TryGetValue(KeyFor(settings), out var bag) ? bag.Count : 0;
}
