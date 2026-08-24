using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using CustomProps = DocumentFormat.OpenXml.CustomProperties;
using VT = DocumentFormat.OpenXml.VariantTypes;

namespace DocumentProcessor.Core.Metadata;

/// <summary>
/// Reads and writes .docx document properties: the standard core set (title, author, subject —
/// what Word's own File &gt; Info panel calls "Properties") and arbitrary named custom properties
/// (what that panel's "Advanced Properties &gt; Custom" tab manages) — e.g. a matter number, contract
/// value, or approval status a downstream system needs without parsing document content.
/// </summary>
public sealed class DocumentMetadataService : IDocumentMetadataService
{
    // Fixed per the OOXML/COM custom-properties convention (every custom document property uses
    // this exact format id — it identifies the "Summary Information"-style property set, not this
    // specific property) and property ids below 2 are reserved.
    private const string CustomPropertyFormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
    private const int FirstCustomPropertyId = 2;

    public void SetCustomProperty(string docxPath, string name, string value) =>
        SetCustomPropertyValue(docxPath, name, new VT.VTLPWSTR { Text = value });

    public void SetCustomProperty(string docxPath, string name, bool value) =>
        SetCustomPropertyValue(docxPath, name, new VT.VTBool { Text = value ? "true" : "false" });

    public void SetCustomProperty(string docxPath, string name, int value) =>
        SetCustomPropertyValue(docxPath, name, new VT.VTInt32 { Text = value.ToString(CultureInfo.InvariantCulture) });

    public void SetCustomProperty(string docxPath, string name, double value) =>
        SetCustomPropertyValue(docxPath, name, new VT.VTDouble { Text = value.ToString(CultureInfo.InvariantCulture) });

    /// <summary>Word stores custom date properties as an ISO-8601 UTC filetime string.</summary>
    public void SetCustomProperty(string docxPath, string name, DateTime value) =>
        SetCustomPropertyValue(docxPath, name, new VT.VTFileTime { Text = value.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture) });

    /// <summary>Removes a custom property. No-ops (returns false) if it wasn't set.</summary>
    public bool RemoveCustomProperty(string docxPath, string name)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var properties = doc.CustomFilePropertiesPart?.Properties;
        var existing = properties?.Elements<CustomProps.CustomDocumentProperty>().FirstOrDefault(p => p.Name?.Value == name);
        if (existing is null)
            return false;

        existing.Remove();
        properties!.Save();
        return true;
    }

    /// <summary>Lists every custom property's name and raw text value. Callers who know a property's
    /// type can re-parse the text (e.g. <see cref="int.Parse(string)"/>) — the value is returned as
    /// text rather than <see langword="object"/> since the OOXML variant-type tag doesn't always
    /// round-trip losslessly to a single .NET type worth guessing at generically.</summary>
    public IReadOnlyDictionary<string, string> GetCustomProperties(string docxPath)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: false);
        var properties = doc.CustomFilePropertiesPart?.Properties;
        if (properties is null)
            return new Dictionary<string, string>();

        return properties.Elements<CustomProps.CustomDocumentProperty>()
            .Where(p => p.Name?.Value is not null)
            .ToDictionary(p => p.Name!.Value!, p => p.InnerText);
    }

    /// <summary>Sets one or more standard core properties (title/author/subject/keywords) — the same
    /// values Word's File &gt; Info panel shows outside the "Custom" tab.</summary>
    public void SetCoreProperties(string docxPath, string? title = null, string? author = null, string? subject = null, string? keywords = null)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
#pragma warning disable OOXML0001 // PackageProperties is marked experimental in the SDK but is the only supported way to set core properties.
        var properties = doc.PackageProperties;
        if (title is not null) properties.Title = title;
        if (author is not null) properties.Creator = author;
        if (subject is not null) properties.Subject = subject;
        if (keywords is not null) properties.Keywords = keywords;
#pragma warning restore OOXML0001
    }

    /// <summary>
    /// Sets many custom properties in one document open/save. The single-property overloads each
    /// pay a full unzip/parse/rezip of the whole package to write well under a kilobyte into
    /// <c>docProps/custom.xml</c> — five properties meant five complete cycles of a 200-page
    /// contract. Prefer this whenever setting more than one.
    /// </summary>
    public void SetCustomProperties(string docxPath, IReadOnlyDictionary<string, object?> properties)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        SetCustomPropertiesCore(doc, properties);
    }

    internal static void SetCustomPropertiesCore(WordprocessingDocument doc, IReadOnlyDictionary<string, object?> properties)
    {
        foreach (var (name, value) in properties)
            SetCustomPropertyValueCore(doc, name, BuildVariant(value));
    }

    /// <summary>Maps a CLR value onto the OOXML variant type Word expects for a custom property.</summary>
    private static DocumentFormat.OpenXml.OpenXmlElement BuildVariant(object? value) => value switch
    {
        null => new VT.VTLPWSTR { Text = string.Empty },
        bool b => new VT.VTBool { Text = b ? "true" : "false" },
        int i => new VT.VTInt32 { Text = i.ToString(CultureInfo.InvariantCulture) },
        long l => new VT.VTInt32 { Text = l.ToString(CultureInfo.InvariantCulture) },
        double d => new VT.VTDouble { Text = d.ToString(CultureInfo.InvariantCulture) },
        decimal m => new VT.VTDouble { Text = m.ToString(CultureInfo.InvariantCulture) },
        DateTime dt => new VT.VTFileTime { Text = dt.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture) },
        _ => new VT.VTLPWSTR { Text = value.ToString() ?? string.Empty }
    };

    internal static IReadOnlyDictionary<string, string> GetCustomPropertiesCore(WordprocessingDocument doc)
    {
        var properties = doc.CustomFilePropertiesPart?.Properties;
        if (properties is null)
            return new Dictionary<string, string>();

        return properties.Elements<CustomProps.CustomDocumentProperty>()
            .Where(p => p.Name?.Value is not null)
            .ToDictionary(p => p.Name!.Value!, p => p.InnerText);
    }

    internal static bool RemoveCustomPropertyCore(WordprocessingDocument doc, string name)
    {
        var properties = doc.CustomFilePropertiesPart?.Properties;
        var existing = properties?.Elements<CustomProps.CustomDocumentProperty>().FirstOrDefault(p => p.Name?.Value == name);
        if (existing is null)
            return false;

        existing.Remove();
        properties!.Save();
        return true;
    }

    private void SetCustomPropertyValue(string docxPath, string name, DocumentFormat.OpenXml.OpenXmlElement valueElement)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        SetCustomPropertyValueCore(doc, name, valueElement);
    }

    internal static void SetCustomPropertyValueCore(WordprocessingDocument doc, string name, DocumentFormat.OpenXml.OpenXmlElement valueElement)
    {
        var customPart = doc.CustomFilePropertiesPart ?? doc.AddNewPart<CustomFilePropertiesPart>();
        customPart.Properties ??= new CustomProps.Properties();

        customPart.Properties.Elements<CustomProps.CustomDocumentProperty>()
            .FirstOrDefault(p => p.Name?.Value == name)?.Remove();

        var nextId = customPart.Properties.Elements<CustomProps.CustomDocumentProperty>()
            .Select(p => p.PropertyId?.Value ?? FirstCustomPropertyId - 1)
            .DefaultIfEmpty(FirstCustomPropertyId - 1)
            .Max() + 1;

        var property = new CustomProps.CustomDocumentProperty
        {
            FormatId = CustomPropertyFormatId,
            PropertyId = nextId,
            Name = name
        };
        property.AppendChild(valueElement);

        customPart.Properties.AppendChild(property);
        customPart.Properties.Save();
    }
}

/// <summary>
/// Document-property operations bound to an open <see cref="Sessions.DocumentSession"/>.
/// </summary>
public sealed class DocumentMetadataOperations
{
    private readonly Sessions.DocumentSession _session;

    internal DocumentMetadataOperations(Sessions.DocumentSession session) => _session = session;

    /// <inheritdoc cref="DocumentMetadataService.SetCustomProperties(string, IReadOnlyDictionary{string, object})"/>
    public void SetCustomProperties(IReadOnlyDictionary<string, object?> properties) =>
        DocumentMetadataService.SetCustomPropertiesCore(_session.Document, properties);

    /// <inheritdoc cref="DocumentMetadataService.GetCustomProperties(string)"/>
    public IReadOnlyDictionary<string, string> GetCustomProperties() =>
        DocumentMetadataService.GetCustomPropertiesCore(_session.Document);

    /// <inheritdoc cref="DocumentMetadataService.RemoveCustomProperty(string, string)"/>
    public bool RemoveCustomProperty(string name) =>
        DocumentMetadataService.RemoveCustomPropertyCore(_session.Document, name);

    /// <inheritdoc cref="DocumentMetadataService.SetCoreProperties(string, string?, string?, string?, string?)"/>
    public void SetCoreProperties(string? title = null, string? author = null, string? subject = null, string? keywords = null)
    {
#pragma warning disable OOXML0001 // PackageProperties is marked experimental but is the only supported route.
        var properties = _session.Document.PackageProperties;
        if (title is not null) properties.Title = title;
        if (author is not null) properties.Creator = author;
        if (subject is not null) properties.Subject = subject;
        if (keywords is not null) properties.Keywords = keywords;
#pragma warning restore OOXML0001
    }
}
