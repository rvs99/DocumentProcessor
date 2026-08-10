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

## Conversion fidelity: what works and what doesn't

The docx→PDF conversion path (LibreOffice headless) is the piece most likely to surprise you, so
here's a direct accounting of what's been verified versus what's a known gap.

**High-fidelity, verified:**
- Plain/default-font documents (`demo-output/01b-contract-draft-normal-font.pdf` in the demo output)
  convert essentially as a mirror of the source docx — no font-substitution variable in play.
- Documents using an embedded custom font (`FontEmbedding/FontEmbeddingService.cs`) also convert
  cleanly, since the font travels inside the docx itself rather than depending on the host having it
  installed.
- Pagination now closely tracks Word's own line-breaking, after `SampleDocumentFactory` was given
  explicit `w:docDefaults` (spacing, line height, margins) — without those, LibreOffice and Word
  disagree noticeably on where pages break, because each falls back to its own built-in defaults
  rather than a shared standard.
- docx and PDF watermarking (removable/locked modes), redlining, track-changes accept/reject, and
  clause transplant have all been verified against real Word behavior, not just LibreOffice's own
  round-trip.

**Known, confirmed limitations — not configuration issues:**
- **VML WordArt (`v:textpath`) is unsupported by LibreOffice's VML importer, full stop.** Confirmed
  via an exhaustive from-scratch diagnostic sweep (size, position, rotation, z-index, header vs.
  body, shapetype presence all ruled out as the cause). Plain VML text boxes (`v:textbox`, what our
  watermarking uses) convert fine — only curved/stylized WordArt text is affected.
- **Pagination fidelity is "very close," not byte-identical to Word.** LibreOffice's layout engine is
  a different implementation from Word's; there's no way to reach pixel-perfect parity on a
  free-only stack, and no way to validate against real Word without a paid license (out of scope by
  design). Treat this as an accepted ceiling, not an open bug — see "Alternatives evaluated" below.
- **Conversion is an external-process dependency, not pure managed code.** It requires `soffice`
  (LibreOffice headless) to actually be installed and reachable wherever this runs — see
  "LibreOffice setup" above. This is a deployment requirement, not a NuGet reference.
- Untested surface area: headers/footers, embedded images/charts, nested tables, RTL text,
  formatting-only tracked changes (bold/italic toggles without text insert/delete), moved-text
  redlining, multi-signer e-sign flows. The demo exercises a realistic contract-editing path, not
  the full OOXML surface.
- `FontEmbeddingService.ApplyFontToAllRuns` (used by the demo) overwrites every run in the document
  uniformly — a demo convenience, not a library constraint. Selective per-run font targeting is a
  small caller-side change on top of `EmbedFontFamily`.
- E-signature support is anchor-tag injection only (`/sig1/`), not a digital-signature/certificate
  implementation — see the "E-sign field injection" design note below for why.

**Alternatives evaluated and rejected** (as replacements for LibreOffice headless):
- **OnlyOffice** — no fidelity advantage in testing, plus an unresolved licensing conflict: a
  dependency's copyright metadata claims "Commercial" despite the project's public AGPL license.
- **docx → HTML → Chromium headless print-to-PDF** — produced *worse* pagination fidelity than
  LibreOffice (fit more content per page, diverging further from Word) and introduced new
  text-rendering artifacts.

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
- **Docx watermarks** support two modes via `DocxWatermarkService.AddTextWatermark(..., removable:)`.
  Word's Design → Watermark UI (both "Remove Watermark" and the predefined gallery's
  replace-existing behavior) identifies a watermark purely by its shape id matching
  `PowerPlusWaterMarkObject<digits>` — not by appearance or position. `removable: true` (the
  default) uses that id, so end users can clear the watermark themselves through Word's own UI.
  `removable: false` uses a different id on purpose, so Word's Watermark commands don't recognize
  or manage it — appropriate for a disclaimer that shouldn't be a single click away from
  disappearing. (PDF watermarks don't currently offer an equivalent: PDFsharp's public API doesn't
  expose a way to attach a custom-appearance object that's cleanly one-click-removable in a viewer —
  the closest built-in mechanism, `PdfRubberStampAnnotation`, is restricted to a fixed set of 14
  standard PDF stamp names rendered in Acrobat's own style, not arbitrary custom text/rotation.)
- Third-party licenses (including a bundled OFL-licensed font) are in
  [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md).
