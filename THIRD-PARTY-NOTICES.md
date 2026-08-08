# Third-Party Notices

This project bundles or depends on the following third-party software. All are free of cost and
permissively licensed (or MPL-2.0 for LibreOffice); none require a paid license at any usage tier.

## NuGet dependencies

| Package | License | Used for |
|---|---|---|
| [DocumentFormat.OpenXml](https://github.com/dotnet/Open-XML-SDK) | MIT | Core .docx (OOXML) manipulation |
| [Clippit](https://github.com/sergey-tihon/Clippit) | MIT | docx comparison (WmlComparer), clause transplant (DocumentBuilder), track-changes acceptance |
| [PDFsharp](https://docs.pdfsharp.net/) | MIT | PDF watermarking, e-sign anchor stamping |
| [PdfPig](https://github.com/UglyToad/PdfPig) | Apache-2.0 | PDF text extraction |
| [PDFtoImage](https://github.com/sungaila/PDFtoImage) | MIT | PDF page rasterization (visual diff) |
| [SkiaSharp](https://github.com/mono/SkiaSharp) | MIT | Image/pixel handling |
| [DiffPlex](https://github.com/mmanela/diffplex) | Apache-2.0 | Text diffing |

> **Note:** `dotnet add package UglyToad.PdfPig` resolves to an unrelated, unlicensed package
> squatting on that name — the correct package ID is `PdfPig` (see table above). Verify package
> ownership before adding new dependencies.

## External tool

- **[LibreOffice](https://www.libreoffice.org/)** (MPL-2.0) — used headless (`soffice --headless --convert-to pdf`)
  for docx→PDF conversion. Not bundled; must be installed in the deployment environment (e.g. via
  `apt-get install libreoffice-writer` in a container).

## Bundled font

- **Roboto Mono** — `src/DocumentProcessor.Core/Assets/Fonts/RobotoMono-Regular.ttf`, licensed under the
  [SIL Open Font License 1.1](src/DocumentProcessor.Core/Assets/Fonts/OFL.txt). Used as the default
  fallback font for PDF text drawing (watermarks, e-sign anchors) so those features work without any
  environment font setup. Source: [google/fonts](https://github.com/google/fonts/tree/main/ofl/robotomono).
