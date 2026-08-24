using DocumentProcessor.Core;
using DocumentProcessor.Core.Conversion;

namespace DocumentProcessor.Tests.Conversion;

/// <summary>
/// Building a LibreOffice user profile costs roughly 80–100 ms of a ~550 ms conversion, so profiles
/// are pooled. The mechanics are tested against synthetic settings rather than a real conversion:
/// the pool is process-wide, so a test that asserted on the shared idle count would race every
/// other conversion running beside it. Giving each test its own executable path gives it its own
/// pool key, which makes these deterministic. The end-to-end saving is measured by the demo rather
/// than asserted here.
/// </summary>
public class LibreOfficeProfilePoolTests
{
    private static LibreOfficeSettings Settings(bool reuse = true, int maxReuses = 50) => new(
        ExecutablePath: $"/nonexistent/soffice-{Guid.NewGuid():N}",
        UseWslDistro: null,
        Timeout: TimeSpan.FromSeconds(30),
        QueueTimeout: TimeSpan.FromSeconds(30),
        ReuseProfiles: reuse,
        MaxProfileReuses: maxReuses);

    [Fact]
    public void A_returned_profile_is_handed_back_to_the_next_lease()
    {
        var settings = Settings();

        var first = LibreOfficeProfilePool.Lease(settings);
        LibreOfficeProfilePool.Return(first, settings);

        Assert.Equal(1, LibreOfficeProfilePool.IdleCount(settings));
        Assert.Same(first, LibreOfficeProfilePool.Lease(settings));
        Assert.Equal(0, LibreOfficeProfilePool.IdleCount(settings));
    }

    [Fact]
    public void Two_simultaneous_leases_never_share_a_profile()
    {
        // Not an optimisation but a correctness rule: two LibreOffice processes pointed at one
        // profile contend for its lock file and hang.
        var settings = Settings();

        var first = LibreOfficeProfilePool.Lease(settings);
        var second = LibreOfficeProfilePool.Lease(settings);

        Assert.NotSame(first, second);
        Assert.NotEqual(first.PathForLibreOffice, second.PathForLibreOffice);

        LibreOfficeProfilePool.Discard(first);
        LibreOfficeProfilePool.Discard(second);
    }

    [Fact]
    public void A_discarded_profile_is_erased_rather_than_pooled()
    {
        var settings = Settings();
        var profile = LibreOfficeProfilePool.Lease(settings);
        Assert.True(Directory.Exists(profile.PathForLibreOffice));

        LibreOfficeProfilePool.Discard(profile);

        // This is the path a killed LibreOffice takes: it leaves its lock file behind, so the
        // profile would wedge whichever conversion leased it next.
        Assert.False(Directory.Exists(profile.PathForLibreOffice));
        Assert.Equal(0, LibreOfficeProfilePool.IdleCount(settings));
    }

    [Fact]
    public void Pooling_can_be_turned_off_entirely()
    {
        var settings = Settings(reuse: false);

        var profile = LibreOfficeProfilePool.Lease(settings);
        LibreOfficeProfilePool.Return(profile, settings);

        Assert.Equal(0, LibreOfficeProfilePool.IdleCount(settings));
        Assert.False(Directory.Exists(profile.PathForLibreOffice));
    }

    [Fact]
    public void A_profile_is_rebuilt_once_it_has_served_its_quota()
    {
        var settings = Settings(maxReuses: 3);
        var profile = LibreOfficeProfilePool.Lease(settings);
        var originalPath = profile.PathForLibreOffice;

        for (var i = 0; i < 2; i++)
        {
            LibreOfficeProfilePool.Return(profile, settings);
            profile = LibreOfficeProfilePool.Lease(settings);
            Assert.Equal(originalPath, profile.PathForLibreOffice);
        }

        // Third return hits the quota: the profile is retired rather than pooled, bounding both the
        // state it accumulates and how long a subtly corrupted one can keep failing conversions.
        LibreOfficeProfilePool.Return(profile, settings);
        Assert.Equal(0, LibreOfficeProfilePool.IdleCount(settings));
        Assert.False(Directory.Exists(originalPath));

        Assert.NotEqual(originalPath, LibreOfficeProfilePool.Lease(settings).PathForLibreOffice);
    }

    [Fact]
    public void Draining_the_pool_erases_what_it_was_holding()
    {
        var settings = Settings();
        var profile = LibreOfficeProfilePool.Lease(settings);
        LibreOfficeProfilePool.Return(profile, settings);

        LibreOfficeProfilePool.DrainAll();

        Assert.Equal(0, LibreOfficeProfilePool.IdleCount(settings));
        Assert.False(Directory.Exists(profile.PathForLibreOffice));
    }

    [Fact]
    public async Task A_run_that_never_started_LibreOffice_does_not_pool_its_profile()
    {
        // The runner has to hand every profile back exactly once, to Return or to Discard. This is
        // the Discard side, and the important one: a profile whose process was killed keeps its
        // lock file and would wedge whichever conversion leased it next.
        var settings = Settings();
        var docxPath = TestFiles.NewTempPath(".docx");
        DocumentProcessor.Core.Samples.SampleDocumentFactory.CreateBasicDocument(docxPath, "Never converted", ["Body."]);

        try
        {
            await Assert.ThrowsAsync<ConversionUnavailableException>(() => LibreOfficeRunner.ConvertBatchAsync(
                settings, [new ConversionItem(docxPath, TestFiles.NewTempPath(".pdf"))], "pdf", CancellationToken.None));

            Assert.Equal(0, LibreOfficeProfilePool.IdleCount(settings));
        }
        finally
        {
            File.Delete(docxPath);
        }
    }
}
