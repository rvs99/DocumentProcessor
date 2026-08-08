# DocumentProcessor

A production-grade .NET module for programmatic Word/PDF document processing. Built entirely on
free, permissively-licensed libraries — no paid or revenue-gated dependencies anywhere in the stack.

## Capabilities

| Capability | Implementation |
|---|---|
| Content control replacement in .docx | `ContentControls/ContentControlService.cs` |
| Programmatic table generation/population | `Tables/TableGenerationService.cs` |
| Custom/embedded font support (docx) | `FontEmbedding/FontEmbeddingService.cs` |
| Custom/embedded font support (PDF) | `PdfFonts/PdfFontResolver.cs` |
| docx → PDF conversion | `Conversion/WordToPdfConverter.cs` (LibreOffice headless) |
| Redlining / docx-vs-docx comparison | `Redlining/DocumentComparisonService.cs` |
| PDF-vs-PDF comparison (text + visual) | `Comparison/PdfComparisonService.cs` |
| Clause/paragraph transplant between docx files | `Transplant/ClauseTransplantService.cs` |
| Watermarking (docx) | `Watermarking/DocxWatermarkService.cs` |
| Watermarking (PDF) | `Watermarking/PdfWatermarkService.cs` |
| E-sign field injection (docx + PDF) | `ESign/ESignFieldService.cs` |
| Track changes accept/reject | `TrackChanges/TrackChangesService.cs` |

All services live under `src/DocumentProcessor.Core`, one folder per capability, and are usable
independently — there's no shared "God object", just plain classes with a handful of public methods.

## Project layout

```
src/DocumentProcessor.Core/    Reusable library — all document-processing services
src/DocumentProcessor.Demo/    Console app: runs a full contract lifecycle through every capability
tests/DocumentProcessor.Tests/ 34 xUnit tests, several exercising the real LibreOffice conversion
```

## Requirements

- [.NET 10 SDK](https://dotnet.microsoft.com/download)
- [LibreOffice](https://www.libreoffice.org/) (headless), for the docx→PDF conversion step and
  anything downstream of it (PDF watermarking/e-sign/comparison in the demo)

### LibreOffice setup

**Production (Linux container):**
```dockerfile
RUN apt-get update && apt-get install -y --no-install-recommends libreoffice-writer \
    && rm -rf /var/lib/apt/lists/*
```
`WordToPdfConversionOptions.ExecutablePath` defaults to `/usr/bin/soffice` on non-Windows — no
further configuration needed.

**Local Windows development**, without installing the full LibreOffice desktop app:
```powershell
wsl --install -d Ubuntu           # if you don't already have a WSL distro
wsl -d Ubuntu -u root -- apt-get update
wsl -d Ubuntu -u root -- apt-get install -y --no-install-recommends libreoffice-writer
```
Then set `DOCPROC_LIBREOFFICE_WSL_DISTRO=Ubuntu` before running tests/the demo, which routes
conversions through `wsl.exe -d Ubuntu -- soffice ...` instead of a native executable
(see `WordToPdfConversionOptions.UseWslDistro`).

`src/DocumentProcessor.Demo/Properties/launchSettings.json` already sets this for you when running
via `dotnet run` or an IDE debug session (F5) — those launch a fresh process that does **not**
inherit environment variables set in your terminal, so without it the demo falls back to looking
for a native `soffice.exe` and skips the PDF steps. If your WSL distro isn't named `Ubuntu`, update
the value there — `dotnet run` applies the launch profile's environment variables on top of your
shell's, so editing that file is the reliable way to change it (a `$env:` override in your terminal
before `dotnet run` will *not* take effect, since the profile value wins).

## Running

```powershell
# Build everything
dotnet build

# Run the test suite (set the env var first if LibreOffice is only available via WSL)
$env:DOCPROC_LIBREOFFICE_WSL_DISTRO = "Ubuntu"
dotnet test

# Run the full lifecycle demo — produces artifacts under src/DocumentProcessor.Demo/bin/.../demo-output/
dotnet run --project src/DocumentProcessor.Demo
```

Without LibreOffice available, the docx-only capabilities (steps 1–10 of the demo) still run fine;
PDF-dependent steps are skipped with a clear message rather than failing the whole run.

## Design notes

- **E-sign field injection** uses the "anchor text" convention (e.g. `/sig1/`) that DocuSign/Adobe
  Sign's own auto-placement features rely on, for both docx and PDF — not native PDF AcroForm
  fields. This is deliberate: PDFsharp 6.2.4's AcroForm field constructors are non-public (it can
  fill existing form fields but not create new ones), and no free library offers a
  provider-agnostic way to create a "real" signature widget anyway, since the signing provider
  always owns that step once the document is uploaded to them.
- **Track changes accept/reject** covers run-level insertions/deletions (`w:ins`/`w:del`), which is
  the large majority of real-world tracked changes. Paragraph-mark and formatting-change revisions
  (`pPrChange`/`rPrChange`) are left as-is.
- Third-party licenses (including a bundled OFL-licensed font) are in
  [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md).
