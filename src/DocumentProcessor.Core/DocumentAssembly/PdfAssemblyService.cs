using DocumentProcessor.Core.PdfFonts;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace DocumentProcessor.Core.DocumentAssembly;

/// <summary>
/// Combines and subsets PDFs — merging a multi-document package, extracting a page range, or
/// appending an exhibit/schedule with page numbering that continues from the main document rather
/// than restarting. Uses PDFsharp's own cross-document page-import (<c>PdfDocument.AddPage</c> on a
/// page opened from another document, or <c>PdfPages.InsertRange</c> for a contiguous block) — no
/// new dependency, since PDFsharp is already used for watermarking/e-sign elsewhere in this library.
/// </summary>
public sealed class PdfAssemblyService
{
    /// <summary>Concatenates <paramref name="pdfPaths"/> in order into one PDF at <paramref name="outputPath"/>.</summary>
    public void MergePdfs(IReadOnlyList<string> pdfPaths, string outputPath)
    {
        if (pdfPaths.Count == 0)
            throw new ArgumentException("Must supply at least one PDF to merge.", nameof(pdfPaths));

        using var output = new PdfDocument();
        foreach (var path in pdfPaths)
        {
            using var input = PdfReader.Open(path, PdfDocumentOpenMode.Import);
            output.Pages.InsertRange(output.PageCount, input);
        }

        output.Save(outputPath);
    }

    /// <summary>
    /// Extracts pages [<paramref name="startPageIndex"/>, <paramref name="endPageIndex"/>] (inclusive,
    /// 0-based) from <paramref name="pdfPath"/> into a new PDF at <paramref name="outputPath"/>.
    /// </summary>
    public void ExtractPages(string pdfPath, int startPageIndex, int endPageIndex, string outputPath)
    {
        using var input = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Import);
        if (startPageIndex < 0 || endPageIndex >= input.PageCount || endPageIndex < startPageIndex)
        {
            throw new ArgumentOutOfRangeException(nameof(endPageIndex),
                $"Document has {input.PageCount} page(s); valid range is 0..{input.PageCount - 1}.");
        }

        using var output = new PdfDocument();
        output.Pages.InsertRange(0, input, startPageIndex, endPageIndex - startPageIndex + 1);
        output.Save(outputPath);
    }

    /// <summary>
    /// Appends <paramref name="exhibitPdfPath"/> after <paramref name="mainPdfPath"/> and stamps a
    /// running page number on every page of the combined result — the exhibit's numbering continues
    /// from the main document's page count rather than restarting at 1, e.g. a 10-page contract
    /// followed by a 3-page exhibit ends up numbered 1-13, not 1-10 then 1-3 again.
    /// </summary>
    public void AppendWithContinuedPageNumbers(
        string mainPdfPath, string exhibitPdfPath, string outputPath,
        int startingPageNumber = 1,
        Func<int, string>? formatPageNumber = null,
        double marginBottomPt = 24)
    {
        formatPageNumber ??= n => n.ToString();
        PdfFontResolver.EnsureRegistered();

        using var mainInput = PdfReader.Open(mainPdfPath, PdfDocumentOpenMode.Import);
        using var exhibitInput = PdfReader.Open(exhibitPdfPath, PdfDocumentOpenMode.Import);

        using var output = new PdfDocument();
        output.Pages.InsertRange(0, mainInput);
        output.Pages.InsertRange(output.PageCount, exhibitInput);

        var font = new XFont("Arial", 9);
        var brush = new XSolidBrush(XColor.FromArgb(0, 0, 0));

        for (var i = 0; i < output.PageCount; i++)
        {
            var page = output.Pages[i];
            var text = formatPageNumber(startingPageNumber + i);
            using var gfx = XGraphics.FromPdfPage(page);
            var size = gfx.MeasureString(text, font);
            gfx.DrawString(text, font, brush, new XPoint((page.Width.Point - size.Width) / 2, page.Height.Point - marginBottomPt));
        }

        output.Save(outputPath);
    }
}
