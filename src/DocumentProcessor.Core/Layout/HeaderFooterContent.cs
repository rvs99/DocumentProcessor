namespace DocumentProcessor.Core.Layout;

/// <summary>One piece of header/footer content — text, a dynamic field, or a logo image, laid out
/// left-to-right in the order supplied to <see cref="HeaderFooterService.SetHeaderContent"/>/
/// <see cref="HeaderFooterService.SetFooterContent"/>.</summary>
public abstract record HeaderFooterPart;

/// <summary>Literal text, or a <c>{{token}}</c>-templated string when a data dictionary is supplied.</summary>
public sealed record TextPart(string Text) : HeaderFooterPart;

public enum HeaderFooterFieldType { PageNumber, TotalPages, Date }

/// <summary>A live Word field (<c>{ PAGE }</c>, <c>{ NUMPAGES }</c>, or <c>{ DATE }</c>) rather than
/// static text — Word recomputes its displayed value on its own, including across page-count changes
/// after this document leaves the library's hands entirely.</summary>
public sealed record FieldPart(HeaderFooterFieldType FieldType) : HeaderFooterPart;

/// <summary>An inline logo image. PNG/JPEG bytes only (detected from the file signature) — vector
/// formats aren't supported since neither PDFsharp's PDF path nor Word's OOXML inline-picture model
/// this uses needs them for a logo-sized raster image.</summary>
public sealed record LogoPart(byte[] ImageBytes, double WidthPt, double HeightPt) : HeaderFooterPart;
