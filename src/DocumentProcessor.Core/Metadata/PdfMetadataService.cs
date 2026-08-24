using System.Text;
using System.Xml.Linq;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace DocumentProcessor.Core.Metadata;

/// <summary>Metadata to embed as an XMP packet in a PDF. Fields left null are omitted from the
/// packet entirely, not written as empty.</summary>
public sealed record XmpMetadata(
    string? Title = null,
    string? Author = null,
    string? Subject = null,
    IReadOnlyList<string>? Keywords = null,
    DateTime? CreateDate = null,
    DateTime? ModifyDate = null);

/// <summary>
/// Embeds an XMP metadata packet in a PDF's document catalog — the metadata form PDF/A and most
/// document-management systems read, as opposed to the older Info dictionary (Title/Author/Subject/
/// Keywords) most PDF libraries default to. Written via PDFsharp's low-level object model
/// (<see cref="PdfDocument.Internals"/>): PDFsharp's public high-level API has no XMP support, but
/// it does let you construct and attach an arbitrary indirect PDF object — a stream dictionary with
/// <c>/Type /Metadata /Subtype /XML</c> referenced from the catalog's <c>/Metadata</c> entry is
/// exactly what Acrobat/Word/veraPDF write themselves, just built by hand here instead.
/// </summary>
public sealed class PdfMetadataService : IPdfMetadataService
{
    public void SetXmpMetadata(string pdfPath, string outputPath, XmpMetadata metadata)
    {
        using var document = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Modify);

        var metadataDictionary = new PdfDictionary(document);
        metadataDictionary.Elements["/Type"] = new PdfName("/Metadata");
        metadataDictionary.Elements["/Subtype"] = new PdfName("/XML");
        metadataDictionary.CreateStream(Encoding.UTF8.GetBytes(BuildXmpPacket(metadata)));

        document.Internals.AddObject(metadataDictionary);
        document.Internals.Catalog.Elements["/Metadata"] = metadataDictionary.Reference;

        // The classic Info dictionary is still what many viewers/libraries (including this one's
        // own PdfPig-based text extraction path) read instead of or alongside XMP — keep both in
        // sync rather than requiring the caller to set each separately.
        if (metadata.Title is not null) document.Info.Title = metadata.Title;
        if (metadata.Author is not null) document.Info.Author = metadata.Author;
        if (metadata.Subject is not null) document.Info.Subject = metadata.Subject;
        if (metadata.Keywords is { Count: > 0 }) document.Info.Keywords = string.Join(", ", metadata.Keywords);

        document.Save(outputPath);
    }

    private static string BuildXmpPacket(XmpMetadata metadata)
    {
        XNamespace x = "adobe:ns:meta/";
        XNamespace rdf = "http://www.w3.org/1999/02/22-rdf-syntax-ns#";
        XNamespace dc = "http://purl.org/dc/elements/1.1/";
        XNamespace xmp = "http://ns.adobe.com/xap/1.0/";

        var description = new XElement(rdf + "Description",
            new XAttribute(XNamespace.Xmlns + "dc", dc),
            new XAttribute(XNamespace.Xmlns + "xmp", xmp),
            new XAttribute(rdf + "about", ""));

        if (metadata.Title is not null)
            description.Add(new XElement(dc + "title", new XElement(rdf + "Alt", LangAlt(metadata.Title))));
        if (metadata.Author is not null)
            description.Add(new XElement(dc + "creator", new XElement(rdf + "Seq", new XElement(rdf + "li", metadata.Author))));
        if (metadata.Subject is not null)
            description.Add(new XElement(dc + "description", new XElement(rdf + "Alt", LangAlt(metadata.Subject))));
        if (metadata.Keywords is { Count: > 0 })
            description.Add(new XElement(dc + "subject", new XElement(rdf + "Bag", metadata.Keywords.Select(k => new XElement(rdf + "li", k)))));
        if (metadata.CreateDate is { } created)
            description.Add(new XElement(xmp + "CreateDate", created.ToString("yyyy-MM-ddTHH:mm:ssK")));
        if (metadata.ModifyDate is { } modified)
            description.Add(new XElement(xmp + "ModifyDate", modified.ToString("yyyy-MM-ddTHH:mm:ssK")));

        var xmpMeta = new XElement(x + "xmpmeta",
            new XAttribute(XNamespace.Xmlns + "x", x),
            new XElement(rdf + "RDF", new XAttribute(XNamespace.Xmlns + "rdf", rdf), description));

        var body = new XDocument(xmpMeta).ToString(SaveOptions.DisableFormatting);
        return $"<?xpacket begin=\"﻿\" id=\"W5M0MpCehiHzreSzNTczkc9d\"?>\n{body}\n<?xpacket end=\"w\"?>";

        XElement LangAlt(string value) =>
            new(rdf + "li", new XAttribute(XNamespace.Xml + "lang", "x-default"), value);
    }
}
