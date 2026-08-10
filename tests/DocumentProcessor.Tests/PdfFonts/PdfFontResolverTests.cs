using System.Text;
using DocumentProcessor.Core.PdfFonts;

namespace DocumentProcessor.Tests.PdfFonts;

/// <summary>
/// <see cref="PdfFontResolver.Instance"/> is a process-wide singleton, so every test uses a unique
/// family name (a fresh GUID) to stay isolated from other tests registering fonts concurrently.
/// </summary>
public class PdfFontResolverTests : IDisposable
{
    private readonly List<string> _tempFiles = [];
    private readonly PdfFontResolver _sut = PdfFontResolver.Instance;

    [Fact]
    public void GetFontBytes_with_only_regular_registered_returns_regular_for_every_style_request()
    {
        var family = UniqueFamily();
        var regular = MarkerFile("REGULAR");
        _sut.RegisterFontFamily(family, new PdfFontFamilyFiles(regular));

        Assert.Equal("REGULAR", ReadMarker(_sut.GetFontBytes(family)));
        Assert.Equal("REGULAR", ReadMarker(_sut.GetFontBytes(family, isBold: true)));
        Assert.Equal("REGULAR", ReadMarker(_sut.GetFontBytes(family, isItalic: true)));
        Assert.Equal("REGULAR", ReadMarker(_sut.GetFontBytes(family, isBold: true, isItalic: true)));
    }

    [Fact]
    public void GetFontBytes_returns_the_exact_registered_face_for_each_style()
    {
        var family = UniqueFamily();
        _sut.RegisterFontFamily(family, new PdfFontFamilyFiles(
            Regular: MarkerFile("REGULAR"),
            Bold: MarkerFile("BOLD"),
            Italic: MarkerFile("ITALIC"),
            BoldItalic: MarkerFile("BOLDITALIC")));

        Assert.Equal("REGULAR", ReadMarker(_sut.GetFontBytes(family)));
        Assert.Equal("BOLD", ReadMarker(_sut.GetFontBytes(family, isBold: true)));
        Assert.Equal("ITALIC", ReadMarker(_sut.GetFontBytes(family, isItalic: true)));
        Assert.Equal("BOLDITALIC", ReadMarker(_sut.GetFontBytes(family, isBold: true, isItalic: true)));
    }

    [Fact]
    public void GetFontBytes_degrades_from_bold_italic_to_bold_when_only_bold_is_registered()
    {
        var family = UniqueFamily();
        _sut.RegisterFontFamily(family, new PdfFontFamilyFiles(Regular: MarkerFile("REGULAR"), Bold: MarkerFile("BOLD")));

        // Asked for bold+italic, but only bold exists -> falls back to bold rather than regular or the bundled default.
        Assert.Equal("BOLD", ReadMarker(_sut.GetFontBytes(family, isBold: true, isItalic: true)));
    }

    [Fact]
    public void ResolveTypeface_for_an_unregistered_family_falls_back_to_the_default_face()
    {
        var info = _sut.ResolveTypeface(UniqueFamily(), isBold: true, isItalic: true);

        Assert.Equal("DocProcDefault", info.FaceName);
    }

    [Fact]
    public void ResolveTypeface_picks_the_composite_face_key_matching_the_requested_style()
    {
        var family = UniqueFamily();
        _sut.RegisterFontFamily(family, new PdfFontFamilyFiles(Regular: MarkerFile("REGULAR"), Italic: MarkerFile("ITALIC")));

        var italicInfo = _sut.ResolveTypeface(family, isBold: false, isItalic: true);
        var regularInfo = _sut.ResolveTypeface(family, isBold: false, isItalic: false);

        // The face name ResolveTypeface returns is later passed straight back into GetFont — verify it round-trips.
        Assert.Equal("ITALIC", ReadMarker(_sut.GetFont(italicInfo.FaceName)));
        Assert.Equal("REGULAR", ReadMarker(_sut.GetFont(regularInfo.FaceName)));
    }

    private static string UniqueFamily() => $"TestFamily-{Guid.NewGuid():N}";

    private string MarkerFile(string marker)
    {
        var path = TestFiles.NewTempPath(".ttf");
        File.WriteAllBytes(path, Encoding.UTF8.GetBytes(marker));
        _tempFiles.Add(path);
        return path;
    }

    private static string ReadMarker(byte[] bytes) => Encoding.UTF8.GetString(bytes);

    public void Dispose()
    {
        foreach (var file in _tempFiles)
            File.Delete(file);
    }
}
