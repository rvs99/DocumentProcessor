using System.Collections.Concurrent;
using System.Reflection;
using PdfSharp.Fonts;

namespace DocumentProcessor.Core.PdfFonts;

/// <summary>
/// PDFsharp on .NET (as opposed to .NET Framework/GDI+) has no built-in font source — it requires
/// an explicit <see cref="IFontResolver"/> to supply font bytes, since there's no guarantee of
/// Windows GDI or any particular OS font store being present (notably not in a Linux container,
/// the deployment target this module is built for). This resolver ships a bundled, permissively
/// licensed default font so PDF text drawing (watermarks, e-sign anchors) works out of the box
/// with zero environment setup, and lets callers register their own custom fonts for
/// production use, mirroring the docx side's font-embedding support.
/// </summary>
public sealed class PdfFontResolver : IFontResolver
{
    public static readonly PdfFontResolver Instance = new();

    private const string DefaultFamilyName = "DocProcDefault";
    private const string DefaultFontResourceName = "DocumentProcessor.Core.Assets.Fonts.RobotoMono-Regular.ttf";

    private readonly ConcurrentDictionary<string, byte[]> _customFonts = new(StringComparer.OrdinalIgnoreCase);
    private readonly Lazy<byte[]> _defaultFontBytes = new(LoadEmbeddedDefaultFont);

    private PdfFontResolver() { }

    /// <summary>Installs this resolver as PDFsharp's global font source, if one isn't already set.</summary>
    public static void EnsureRegistered()
    {
        GlobalFontSettings.FontResolver ??= Instance;
    }

    /// <summary>
    /// Registers a custom TrueType/OpenType font family for use in PDF drawing calls (e.g.
    /// <c>new XFont("Brand Sans", 12)</c> after calling <c>RegisterFont("Brand Sans", bytes)</c>).
    /// </summary>
    public void RegisterFont(string familyName, byte[] fontBytes) => _customFonts[familyName] = fontBytes;

    public void RegisterFont(string familyName, string fontFilePath) => RegisterFont(familyName, File.ReadAllBytes(fontFilePath));

    public byte[] GetFont(string faceName) =>
        _customFonts.TryGetValue(faceName, out var bytes) ? bytes : _defaultFontBytes.Value;

    public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
    {
        var faceName = _customFonts.ContainsKey(familyName) ? familyName : DefaultFamilyName;
        return new FontResolverInfo(faceName);
    }

    private static byte[] LoadEmbeddedDefaultFont()
    {
        using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(DefaultFontResourceName)
            ?? throw new InvalidOperationException($"Embedded default font resource '{DefaultFontResourceName}' not found.");
        using var memoryStream = new MemoryStream();
        stream.CopyTo(memoryStream);
        return memoryStream.ToArray();
    }
}
