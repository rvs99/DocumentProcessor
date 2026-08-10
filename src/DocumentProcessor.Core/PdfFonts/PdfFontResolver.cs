using System.Collections.Concurrent;
using System.Reflection;
using PdfSharp.Fonts;

namespace DocumentProcessor.Core.PdfFonts;

/// <summary>
/// The font files for one PDF-drawing family. Mirrors <c>Fonts.FontFamilyFiles</c> on the docx side
/// — regular is required, bold/italic/bold-italic are optional. Files are read as-is; PDFsharp/
/// SkiaSharp accept TTF and OTF here (unlike the docx embedding path, PDF font drawing has no
/// Word-compatibility constraint to satisfy).
/// </summary>
public sealed record PdfFontFamilyFiles(string Regular, string? Bold = null, string? Italic = null, string? BoldItalic = null);

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

    /// <summary>
    /// Registers up to 4 style variants for a family in one call, so bold/italic PDF text (e.g. a
    /// bold watermark, or an italic e-sign label) uses an actual bold/italic font file instead of
    /// PDFsharp/SkiaSharp's synthetic (faux) bold/italic, which just skews or thickens the regular
    /// glyphs. Unset variants fall back to whichever of bold/italic/regular is the closest available
    /// match — see <see cref="ResolveTypeface"/>.
    /// </summary>
    public void RegisterFontFamily(string familyName, PdfFontFamilyFiles files)
    {
        RegisterFont(familyName, files.Regular);
        if (files.Bold is not null)
            _customFonts[FaceKey(familyName, bold: true, italic: false)] = File.ReadAllBytes(files.Bold);
        if (files.Italic is not null)
            _customFonts[FaceKey(familyName, bold: false, italic: true)] = File.ReadAllBytes(files.Italic);
        if (files.BoldItalic is not null)
            _customFonts[FaceKey(familyName, bold: true, italic: true)] = File.ReadAllBytes(files.BoldItalic);
    }

    public byte[] GetFont(string faceName) =>
        _customFonts.TryGetValue(faceName, out var bytes) ? bytes : _defaultFontBytes.Value;

    /// <summary>
    /// Resolves font bytes for direct use outside PDFsharp (e.g. SkiaSharp text rasterization),
    /// with the same registered-custom-font-or-bundled-default fallback as <see cref="GetFont"/>,
    /// plus the same bold/italic degradation as <see cref="ResolveTypeface"/>.
    /// </summary>
    public byte[] GetFontBytes(string? familyName, bool isBold = false, bool isItalic = false)
    {
        if (familyName is null)
            return _defaultFontBytes.Value;

        if (TryGetBestFace(familyName, isBold, isItalic, out var bytes))
            return bytes;

        return _defaultFontBytes.Value;
    }

    public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
    {
        if (!_customFonts.ContainsKey(familyName))
            return new FontResolverInfo(DefaultFamilyName);

        if (isBold && isItalic && _customFonts.ContainsKey(FaceKey(familyName, bold: true, italic: true)))
            return new FontResolverInfo(FaceKey(familyName, bold: true, italic: true));
        if (isBold && _customFonts.ContainsKey(FaceKey(familyName, bold: true, italic: false)))
            return new FontResolverInfo(FaceKey(familyName, bold: true, italic: false));
        if (isItalic && _customFonts.ContainsKey(FaceKey(familyName, bold: false, italic: true)))
            return new FontResolverInfo(FaceKey(familyName, bold: false, italic: true));

        return new FontResolverInfo(familyName);
    }

    /// <summary>
    /// Degrades gracefully through exact match -> bold-italic -> bold -> italic -> regular, so
    /// asking for a style that wasn't registered still returns the closest available face instead
    /// of silently falling back to the bundled default font.
    /// </summary>
    private bool TryGetBestFace(string familyName, bool isBold, bool isItalic, out byte[] bytes)
    {
        if (isBold && isItalic && _customFonts.TryGetValue(FaceKey(familyName, bold: true, italic: true), out bytes!))
            return true;
        if (isBold && _customFonts.TryGetValue(FaceKey(familyName, bold: true, italic: false), out bytes!))
            return true;
        if (isItalic && _customFonts.TryGetValue(FaceKey(familyName, bold: false, italic: true), out bytes!))
            return true;

        return _customFonts.TryGetValue(familyName, out bytes!);
    }

    /// <summary>Composite dictionary key for a non-regular style variant — regular itself is keyed by the bare family name.</summary>
    private static string FaceKey(string familyName, bool bold, bool italic) =>
        (bold, italic) switch
        {
            (true, true) => $"{familyName}#BoldItalic",
            (true, false) => $"{familyName}#Bold",
            (false, true) => $"{familyName}#Italic",
            _ => familyName
        };

    private static byte[] LoadEmbeddedDefaultFont()
    {
        using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(DefaultFontResourceName)
            ?? throw new InvalidOperationException($"Embedded default font resource '{DefaultFontResourceName}' not found.");
        using var memoryStream = new MemoryStream();
        stream.CopyTo(memoryStream);
        return memoryStream.ToArray();
    }
}
