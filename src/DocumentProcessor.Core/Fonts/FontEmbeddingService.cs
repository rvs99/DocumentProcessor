using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentProcessor.Core.FontEmbedding;

/// <summary>
/// The TrueType/OpenType-TT files for one font family. Only .ttf/.ttc are supported for direct
/// (non-obfuscated) embedding — the format Word itself uses for "Embed fonts in the file".
/// </summary>
public sealed record FontFamilyFiles(string Regular, string? Bold = null, string? Italic = null, string? BoldItalic = null);

/// <summary>
/// Embeds custom/non-standard font files directly into a .docx so the document renders correctly
/// on machines that don't have the font installed — required for production output whose branding
/// or legal templates depend on a specific typeface, and a prerequisite for high-fidelity docx→PDF
/// conversion when the conversion host doesn't have the font installed system-wide.
/// </summary>
public sealed class FontEmbeddingService
{
    /// <summary>
    /// Embeds the given font family's files into the document and marks the document as using
    /// embedded fonts. Call <see cref="ApplyFontToAllRuns"/> (or set run fonts yourself) to actually
    /// use the family in the document body — embedding alone only makes the font available.
    /// </summary>
    public void EmbedFontFamily(string docxPath, string fontFamilyName, FontFamilyFiles files)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var mainPart = doc.MainDocumentPart ?? throw new CorruptDocumentException("Document has no main part.");

        var fontTablePart = mainPart.FontTablePart ?? mainPart.AddNewPart<FontTablePart>();
        fontTablePart.Fonts ??= new Fonts();

        // Re-embedding the same family replaces the prior entry rather than duplicating it.
        fontTablePart.Fonts.Elements<Font>()
            .FirstOrDefault(f => string.Equals(f.Name?.Value, fontFamilyName, StringComparison.Ordinal))
            ?.Remove();

        var font = new Font { Name = fontFamilyName };
        font.AppendChild(new EmbedRegularFont { Id = EmbedFontFile(fontTablePart, files.Regular) });
        if (files.Bold is not null)
            font.AppendChild(new EmbedBoldFont { Id = EmbedFontFile(fontTablePart, files.Bold) });
        if (files.Italic is not null)
            font.AppendChild(new EmbedItalicFont { Id = EmbedFontFile(fontTablePart, files.Italic) });
        if (files.BoldItalic is not null)
            font.AppendChild(new EmbedBoldItalicFont { Id = EmbedFontFile(fontTablePart, files.BoldItalic) });

        fontTablePart.Fonts.AppendChild(font);
        fontTablePart.Fonts.Save();

        var settingsPart = mainPart.DocumentSettingsPart ?? mainPart.AddNewPart<DocumentSettingsPart>();
        settingsPart.Settings ??= new Settings();
        if (!settingsPart.Settings.Elements<EmbedTrueTypeFonts>().Any())
            settingsPart.Settings.PrependChild(new EmbedTrueTypeFonts { Val = true });
        settingsPart.Settings.Save();
    }

    /// <summary>Sets every run in the document to use <paramref name="fontFamilyName"/> for all script types.</summary>
    public void ApplyFontToAllRuns(string docxPath, string fontFamilyName)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var document = doc.MainDocumentPart?.Document ?? throw new CorruptDocumentException("Document has no main part/body.");

        foreach (var run in document.Descendants<Run>())
        {
            run.RunProperties ??= new RunProperties();
            run.RunProperties.RunFonts = new RunFonts
            {
                Ascii = fontFamilyName,
                HighAnsi = fontFamilyName,
                ComplexScript = fontFamilyName,
                EastAsia = fontFamilyName
            };
        }

        document.Save();
    }

    /// <summary>Lists the font families currently embedded in the document.</summary>
    public IReadOnlyList<string> ListEmbeddedFonts(string docxPath)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: false);
        var fontTablePart = doc.MainDocumentPart?.FontTablePart;
        if (fontTablePart?.Fonts is null)
            return [];

        return fontTablePart.Fonts.Elements<Font>()
            .Select(f => f.Name?.Value)
            .Where(name => name is not null)
            .Select(name => name!)
            .ToList();
    }

    private static string EmbedFontFile(FontTablePart fontTablePart, string fontFilePath)
    {
        var extension = Path.GetExtension(fontFilePath).ToLowerInvariant();
        var partType = extension switch
        {
            ".ttf" or ".ttc" => FontPartType.FontTtf,
            _ => throw new NotSupportedException(
                $"Unsupported font file extension '{extension}'. Only .ttf/.ttc are supported for " +
                "direct embedding (e.g. .otf must be converted to .ttf first).")
        };

        var fontPart = fontTablePart.AddFontPart(partType);
        using (var fileStream = File.OpenRead(fontFilePath))
        {
            fontPart.FeedData(fileStream);
        }

        return fontTablePart.GetIdOfPart(fontPart);
    }
}
