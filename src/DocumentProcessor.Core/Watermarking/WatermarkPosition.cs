namespace DocumentProcessor.Core.Watermarking;

/// <summary>Where a watermark sits on the page, shared by both <see cref="DocxWatermarkService"/> and <see cref="PdfWatermarkService"/>.</summary>
public enum WatermarkPosition
{
    Center,
    TopLeft, TopCenter, TopRight,
    MiddleLeft, MiddleRight,
    BottomLeft, BottomCenter, BottomRight
}
