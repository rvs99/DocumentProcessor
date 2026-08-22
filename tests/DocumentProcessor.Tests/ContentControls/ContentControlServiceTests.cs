using DocumentProcessor.Core.ContentControls;
using DocumentProcessor.Core.Samples;

namespace DocumentProcessor.Tests.ContentControls;

public class ContentControlServiceTests : IDisposable
{
    private readonly string _path = TestFiles.NewTempPath(".docx");
    private readonly ContentControlService _sut = new();

    public ContentControlServiceTests()
    {
        SampleDocumentFactory.CreateDocumentWithContentControls(_path, new Dictionary<string, string>
        {
            ["ClientName"] = "[Client Name]",
            ["EffectiveDate"] = "[Date]"
        });
    }

    [Fact]
    public void ReplaceByTag_updates_matching_control_and_leaves_others_untouched()
    {
        var updated = _sut.ReplaceByTag(_path, "ClientName", "Acme Corp");

        Assert.Equal(1, updated);
        var controls = _sut.ListContentControls(_path).ToDictionary(c => c.Tag!, c => c.Text);
        Assert.Equal("Acme Corp", controls["ClientName"]);
        Assert.Equal("[Date]", controls["EffectiveDate"]);
    }

    [Fact]
    public void ReplaceByTag_with_unknown_tag_updates_nothing()
    {
        var updated = _sut.ReplaceByTag(_path, "DoesNotExist", "value");

        Assert.Equal(0, updated);
    }

    [Fact]
    public void ReplaceMany_populates_every_matching_control_in_one_pass()
    {
        var counts = _sut.ReplaceMany(_path, new Dictionary<string, string>
        {
            ["ClientName"] = "Acme Corp",
            ["EffectiveDate"] = "2026-08-09"
        });

        Assert.Equal(1, counts["ClientName"]);
        Assert.Equal(1, counts["EffectiveDate"]);

        var controls = _sut.ListContentControls(_path).ToDictionary(c => c.Tag!, c => c.Text);
        Assert.Equal("Acme Corp", controls["ClientName"]);
        Assert.Equal("2026-08-09", controls["EffectiveDate"]);
    }

    [Fact]
    public void ReplaceByTag_byte_array_overload_updates_the_control_with_no_filesystem_output()
    {
        var originalBytes = File.ReadAllBytes(_path);

        var (updatedBytes, updatedCount) = _sut.ReplaceByTag(originalBytes, "ClientName", "Acme Corp");

        Assert.Equal(1, updatedCount);
        Assert.NotEqual(originalBytes, updatedBytes);

        var roundTripPath = TestFiles.NewTempPath(".docx");
        try
        {
            File.WriteAllBytes(roundTripPath, updatedBytes);
            var controls = _sut.ListContentControls(roundTripPath).ToDictionary(c => c.Tag!, c => c.Text);
            Assert.Equal("Acme Corp", controls["ClientName"]);
            Assert.Equal("[Date]", controls["EffectiveDate"]);
        }
        finally
        {
            File.Delete(roundTripPath);
        }
    }

    [Fact]
    public void ReplaceByTag_byte_array_overload_does_not_modify_the_original_file_on_disk()
    {
        var originalBytes = File.ReadAllBytes(_path);

        _sut.ReplaceByTag(originalBytes, "ClientName", "Acme Corp");

        Assert.Equal(originalBytes, File.ReadAllBytes(_path));
    }

    [Fact]
    public void ListContentControls_returns_tag_alias_and_current_text_for_every_control()
    {
        var controls = _sut.ListContentControls(_path);

        Assert.Equal(2, controls.Count);
        Assert.Contains(controls, c => c.Tag == "ClientName" && c.Alias == "ClientName" && c.Text == "[Client Name]");
    }

    public void Dispose() => File.Delete(_path);
}
