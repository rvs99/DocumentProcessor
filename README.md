# DocumentProcessor

A production-grade .NET module for programmatic Word/PDF document processing. Built entirely on
free, permissively-licensed libraries — no paid or revenue-gated dependencies anywhere in the stack.

## Capabilities

### Templating & document assembly

| Capability | Implementation |
|---|---|
| `{{token}}` mail-merge filling — run-merged scanning, so a token split across Word's own runs still matches | `Templating/TemplateEngine.cs`, `Templating/RunTextScanner.cs` |
| Missing-token policy: `Error` / `Redact` / `Highlight` (wraps the literal token, marked for review) | `Templating/MissingTokenPolicy.cs` |
| `{{html:token}}` rich-text injection, sanitized against an allow-list before conversion to OOXML | `Templating/HtmlToOoxmlConverter.cs` |
| `{{if:...}}` / `{{else}}` / `{{/if}}` conditional sections | `Templating/TemplateCondition.cs`, `TemplateEngine.cs` |
| `{{repeat:collection}}` / `{{/repeat}}` repeating sections, driven by a list of row dictionaries | `Templating/TemplateEngine.cs` |
| `{{clause:id}}` marker-based clause-library injection, with heading-numbering continuation | `Templating/ClauseLibrary.cs`, `TemplateEngine.cs` |
| Content control (SDT) replacement, type-aware (plain/rich-text/date/drop-down), by tag or in bulk | `ContentControls/ContentControlService.cs` |
| Content control locking after fill (`SetLock`, `ContentControlLockMode`) | `ContentControls/ContentControlService.cs` |
| Clause/paragraph transplant, removal, and replacement between docx files | `Transplant/ClauseTransplantService.cs` |
| Clause removal with cross-reference cleanup (dangling `REF`/`PAGEREF` reporting) | `Transplant/ClauseTransplantService.cs`, `DocumentAssembly/CrossReferenceValidator.cs` |
| Field dirtying / update-on-open, so Word recomputes TOC and cross-reference display text | `DocumentAssembly/FieldUpdateService.cs` |
| Programmatic table generation/population — column widths, borders, named styles, merged cells | `Tables/TableGenerationService.cs` |
| PDF assembly: merge, page-range extraction, exhibit append with continued page numbering | `DocumentAssembly/PdfAssemblyService.cs` |

### Conversion, redlining & review

| Capability | Implementation |
|---|---|
| docx → PDF conversion, batched and profile-pooled | `Conversion/WordToPdfConverter.cs` (LibreOffice headless) |
| Legacy `.doc` (binary OLE) → docx conversion, ahead of the normal OpenXml pipeline | `Format/LegacyDocConverter.cs` |
| Redlining / docx-vs-docx comparison with reviewer identity (Clippit `WmlComparer`) | `Redlining/DocumentComparisonService.cs` |
| Four-variant redline export: clean/redlined × docx/PDF, in one call | `Redlining/RedlineExportService.cs` |
| PDF-vs-PDF comparison (text + visual) | `Comparison/PdfComparisonService.cs` |
| Structured track-changes list (author/date/id/kind) and accept/reject by author or change id | `TrackChanges/TrackChangesService.cs` |
| Section-level edit restriction (comments-only / tracked-changes-only / forms-only), password-backed, with per-range exceptions | `Redlining/DocumentProtectionService.cs` |
| Word comments — read, add, reply, resolve, delete, with threading (`commentsEx`) | `Comments/DocumentCommentService.cs` |
| Plain-text and clause-level text extraction, tracked-deletion-aware | `Extraction/TextExtractionService.cs` |

### Branding, layout & watermarking

| Capability | Implementation |
|---|---|
| Page size, orientation, margins, columns, page breaks, default spacing, multi-section documents | `Layout/PageLayoutService.cs` |
| Headers/footers (default, first-page, even-page) | `Layout/HeaderFooterService.cs` |
| Tenant branding — header logo plus a document-wide accent color across headings | `Layout/BrandingService.cs` |
| Watermarking, docx and PDF, with position/size control | `Watermarking/DocxWatermarkService.cs`, `Watermarking/PdfWatermarkService.cs` |
| Status → watermark policy mapping (e.g. Draft/Final/Confidential → a configured watermark) | `Watermarking/StatusWatermarkPolicy.cs` |
| Custom/embedded font support (docx) | `Fonts/FontEmbeddingService.cs` |
| Custom/embedded font support (PDF) | `PdfFonts/PdfFontResolver.cs` |
| E-sign field injection (docx + PDF) | `ESign/ESignFieldService.cs` |

### Metadata, security & audit

| Capability | Implementation |
|---|---|
| Custom docx properties (string/bool/int/double/date) and standard core properties | `Metadata/DocumentMetadataService.cs` |
| XMP metadata embedding in PDF | `Metadata/PdfMetadataService.cs` |
| Macro/VBA detection, stripping, and macro-enabled template validation | `Security/MacroValidationService.cs` |
| PDF password protection (owner/user passwords, permission flags) | `Security/PdfProtectionService.cs` |
| Zip-bomb / decompression-ratio defence on every document opened, reading only the ZIP central directory | `Security/DocumentLimits.cs` |
| `ActivitySource` telemetry and `ILogger<T>` operation logging, opt-in and zero-overhead when unobserved | `Diagnostics/DocumentProcessorDiagnostics.cs` |

### API surface

| Capability | Implementation |
|---|---|
| Session/handle pattern: one open package, many operations, one save — byte[]-in/byte[]-out, no intermediate files | `Sessions/DocumentSession.cs`, `Sessions/SessionOperations.cs` |
| Dependency-injection registration for every service, `TryAdd` semantics throughout | `Abstractions/ServiceCollectionExtensions.cs` |
| A typed exception hierarchy (`DocumentProcessorException`) distinguishing retryable failures from deployment/input faults | `DocumentProcessorException.cs` |

All services live under `src/DocumentProcessor.Core`, one folder per capability, and are usable
independently — there's no shared "God object", just plain classes with a handful of public methods.
`DocumentSession` is the recommended entry point for any multi-step pipeline; the path-based service
classes above remain available directly for single-operation use.

## Project layout

```
src/DocumentProcessor.Core/    Reusable library — all document-processing services
src/DocumentProcessor.Demo/    Console app: runs a full contract lifecycle through every capability
tests/DocumentProcessor.Tests/ 331 xUnit tests, several exercising the real LibreOffice conversion
BACKLOG.md                     Deferred work, blocked items, known issues, and test baselines
```

[`BACKLOG.md`](BACKLOG.md) is worth reading before picking anything up: it records what has already
been measured or ruled out (and why), including the expected pass/fail counts for the test suite in
each environment.

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

### Conversion throughput

LibreOffice spends roughly 450 ms starting up before it looks at the first document, which for
small contracts is most of the conversion. Two things keep that off the request path, both on by
default and both configurable on `WordToPdfConversionOptions`:

- **Batching.** `ConvertBatchAsync` converts many documents in one LibreOffice invocation — the
  natural shape for a contract plus its exhibits. Measured on six documents: 6,421 ms one at a
  time against 2,600 ms as a batch. Concurrent `ConvertAsync` calls that end up queued behind a
  busy host are coalesced into shared invocations automatically, with no change to calling code —
  six concurrent calls completed in 1,636 ms. Nothing is ever held back waiting for a companion
  request, so an idle server behaves exactly as it did before.
- **Profile pooling.** LibreOffice user profiles are kept warm and reused rather than rebuilt per
  call, which is worth roughly 80–100 ms of a ~550 ms conversion. Profiles are still leased
  exclusively — two LibreOffice processes sharing one profile contend for its lock and hang — and
  any profile whose process had to be killed is destroyed rather than reused, since a killed
  LibreOffice leaves its lock file behind.

Batch conversion reports failure per document: one unreadable upload does not deny the rest of the
batch their results.

The process-wide cap on concurrent LibreOffice instances (`DOCPROC_MAX_CONCURRENT_CONVERSIONS`,
default: half the CPU count) still bounds everything above — batching changes how many documents
each process handles, not how many processes exist.

## Usage

### One document, several operations: `DocumentSession`

Every path-based service (`ContentControlService`, `TableGenerationService`, and so on) opens the
package, does one thing, and saves — fine standalone, wasteful chained. A five-step pipeline through
five services means five open/parse/rezip cycles. `DocumentSession` opens once, exposes the same
operations as properties, and saves once:

```csharp
using var session = DocumentSession.Open(uploadedBytes); // byte[]-in, no temp file required

session.ContentControls.ReplaceMany(new Dictionary<string, string>
{
    ["ClientName"] = "Acme Corporation",
    ["EffectiveDate"] = "2027-01-01",
});
session.Tables.AppendTable(pricingSpec);
session.Watermark.AddText("DRAFT");
session.Metadata.SetCustomProperties(new Dictionary<string, object?> { ["MatterNumber"] = "M-2027-0042" });
session.Protection.Restrict(EditRestriction.TrackedChanges, password: "…");

byte[] result = session.Save(); // byte[]-out
```

`DocumentSession.Open` validates the incoming package against `DocumentLimits.Default` (size,
entry count, and decompression-ratio limits) before anything else touches it — pass a different
`DocumentLimits` (or `.Unbounded`) for documents from a trusted source. Available operation groups:
`ContentControls`, `Metadata`, `Tables`, `TrackChanges`, `PageLayout`, `Watermark`, `Protection`,
`Fonts`, `Fields`, `Comments`, `Text`.

### Dependency injection

```csharp
services.AddDocumentProcessor(options => options.ExecutablePath = "/usr/bin/soffice");
```

Registers every service in the table above against its interface, as singletons (`TryAdd`
semantics, so an application's own registration of any interface takes precedence). Services hold
no per-call mutable state — per-document state lives entirely in `DocumentSession`, which callers
create and dispose per operation, not in DI.

### Error handling

Every exception this library throws derives from `DocumentProcessorException`, with
`IsRetryable` distinguishing failures worth an automatic retry from ones that will fail identically
every time:

```
DocumentProcessorException          IsRetryable => false by default
├── CorruptDocumentException        malformed or unreadable package
├── DocumentTooComplexException     structural limits (nesting, size) exceeded
│   └── HtmlTooComplexException     rich-text input specifically
├── TemplateException               template fill failed
│   └── MissingTemplateTokenException   MissingTokenPolicy.Error, token has no value
├── ConversionUnavailableException  LibreOffice missing/misconfigured — IsRetryable false
└── ConversionFailedException       the conversion itself failed — IsRetryable true
```

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

All 26 sections of the demo run without LibreOffice available — the docx-only capabilities in each
section run fine, and individual PDF-dependent steps within a section report a clear "SKIPPED"
message rather than failing the whole run.

## Conversion fidelity: what works and what doesn't

The docx→PDF conversion path (LibreOffice headless) is the piece most likely to surprise you, so
here's a direct accounting of what's been verified versus what's a known gap.

**High-fidelity, verified:**
- Plain/default-font documents (`demo-output/02-draft-normal-font.pdf` in the demo output)
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

## Fonts, page size, margins, and layout: what's supported

A direct accounting of what's configurable through the library versus what's a fixed default or
missing outright. Where something says "demo/sample-only," it means the value is hardcoded in
`Samples/SampleDocumentFactory.cs` (used to generate throwaway fixtures for the demo and tests) —
real callers bringing their own documents aren't bound by it.

### Fonts

| | Supported | Not supported |
|---|---|---|
| **docx embedding** (`Fonts/FontEmbeddingService.cs`) | `EmbedFontFamily` embeds up to 4 style variants per family (regular required; bold/italic/bold-italic optional) via `FontFamilyFiles`. `ApplyFontToAllRuns` sets one family across all run script slots (ascii/high-ansi/complex-script/east-asian) at once. `ListEmbeddedFonts` for introspection. | **TTF/TTC only** — and not fixable: Word's own "Embed fonts in file" feature only embeds TrueType-flavored fonts, so a `.otf` restriction here matches Word's own ceiling, not a gap in this library. No font subsetting either — no free .NET library in this stack does TTF/OTF subsetting (files embed whole). No per-script-type font mapping via `ApplyFontToAllRuns` (all 4 slots forced to the same family — callers wanting different East Asian/complex-script fonts need to set `RunFonts` themselves per run). |
| **PDF fonts** (`PdfFonts/PdfFontResolver.cs`) | `RegisterFont` for a single (regular-only) family, or `RegisterFontFamily` with a `PdfFontFamilyFiles` (mirroring the docx side's `FontFamilyFiles`) to register up to 4 style variants at once. `ResolveTypeface`/`GetFontBytes` now actually use `isBold`/`isItalic` to pick the matching registered face, degrading gracefully bold-italic → bold → italic → regular → bundled default when the exact style wasn't registered. Verified via `PdfFontResolverTests` (no bold/italic font asset is bundled in this repo, so this isn't shown in the demo — only unit-tested with marker files). One bundled fallback font (Roboto Mono, regular only) ships for when nothing custom is registered. | No standard-14 (Helvetica/Times/Courier) special-casing. No explicit Unicode-coverage/CID handling beyond what PDFsharp/SkiaSharp do natively with the supplied bytes. |

### Page size, orientation, margins, columns, breaks, and spacing — `Layout/PageLayoutService.cs`

All section-properties methods take an optional `sectionIndex` (defaults to every section in the
document — every document from this library has exactly one today, but the parameter means a future
multi-section document won't need this API reworked).

| | Supported | Not supported |
|---|---|---|
| **Page size & orientation** | `SetPageSize` with `PageSize.Letter()/.A4()/.Legal()` presets (each taking `PageOrientation.Portrait`/`.Landscape`, which swaps width/height and sets `w:pgSz/@w:orient`) or an explicit `PageSize(widthTwips, heightTwips)`. Verified surviving real LibreOffice conversion (`LayoutConversionTests`) — landscape Letter converts to a genuine 792×612pt PDF page, not just correct docx XML. | Any size beyond what you construct explicitly — there's no named-preset list beyond Letter/A4/Legal. |
| **Margins** | `SetMargins` with `PageMargins.FromInches(...)` or explicit twips, covering top/bottom/left/right/header/footer/gutter. Callers bringing their own docx can now set this directly instead of it being demo-only. | PDF still has no margin concept — the PDF services only stamp onto or read existing page geometry; nothing creates or resizes PDF pages directly (margins only apply on the docx side, before conversion). |
| **Columns** | `SetColumns(docxPath, columnCount, spacingTwips)` — equal-width `w:cols`. | Unequal column widths, or a column break mid-content (only the section-level column *count* is configurable). |
| **Page breaks** | `InsertPageBreak(docxPath, beforeParagraphIndex)` — same 0-based paragraph-index convention as `ClauseTransplantService.ListParagraphs`. | — |
| **Default paragraph/line spacing** | `SetDefaultParagraphSpacing(docxPath, afterTwips, lineTwips, lineRule)` — updates both `w:docDefaults` and the `Normal` style, the same mechanism `SampleDocumentFactory` now calls (rather than duplicating the logic) to match Word's Normal.dotm pagination baseline. | Per-paragraph spacing overrides — this sets the document-wide default, not individual paragraph properties. |
| **Multi-section documents** | `InsertSectionBreak(docxPath, beforeParagraphIndex, breakType?)` splits the document into two independently-laid-out sections at a paragraph boundary (0-based, same convention as `InsertPageBreak`) — call it repeatedly for more than two. The new earlier section starts as a copy of whatever the document's current last section looks like; call `SetPageSize`/`SetMargins`/`SetColumns` afterward with the relevant `sectionIndex` to differentiate them (e.g. a landscape exhibit followed by a portrait appendix, both in one docx — exactly the demo's step 6). Verified surviving real LibreOffice conversion with two genuinely different page geometries in one PDF, not just correct docx XML (`MultiSectionConversionTests`). | Per-section headers/footers and per-section watermarks aren't wired up — `DocxWatermarkService`/`HeaderFooterService` still only address the document's single trailing section reference; a multi-section document's earlier sections keep whatever header/footer (if any) they already had. |

### Headers & footers — `Layout/HeaderFooterService.cs`

General-purpose header/footer text, distinct from `DocxWatermarkService`'s header injection (which
is single-purpose — it only ever hosts the watermark shape).

| | Supported | Not supported |
|---|---|---|
| **Headers** | `SetHeaderText`/`RemoveHeader`, all three Word variants (`HeaderFooterValues.Default`/`.First`/`.Even`). Setting `.First` also turns on `w:titlePg`; setting `.Even` also turns on `w:evenAndOddHeaders` in document settings — both required for Word to actually honor that variant, not just accept the reference. | Non-text content (logos/images) — text only for now. |
| **Footers** | `SetFooterText`/`RemoveFooter`, same three variants — this closes a gap that had zero footer-related code anywhere in the repo before. | Same as headers — text only. |

### Tables — `Tables/TableGenerationService.cs`

| | Supported | Not supported |
|---|---|---|
| **Column widths** | `TableSpec.ColumnWidthsTwips` — explicit per-column widths (one entry per header column), applied to both `w:tblGrid` and every cell's `w:tcW`, with `w:tblLayout w:type="fixed"` so the renderer honors them instead of auto-fitting content. Omit for the prior auto-width behavior. | Mixed auto/fixed columns in one table — it's all-explicit or all-auto. |
| **Borders** | `TableSpec.Borders` (a `TableBorderSpec`) — style, size (eighths-of-a-point), and an optional color, applied uniformly to all six border positions (outer edges + inside lines). Omit for the prior default (single 0.5pt line, no explicit color). | Per-edge border variation (e.g. a heavier outer border than inside lines) — one spec applies everywhere. |
| **Named table styles** | `TableSpec.TableStyleId` sets `w:tblStyle` to reference a style already defined in the target document's styles part. | Not validated, and this library doesn't define/generate table styles itself — the caller is responsible for the style existing (e.g. one already present in a template docx). |
| **Merged cells** | `TableSpec.Merges` (a list of `TableCellMerge(RowIndex, ColumnIndex, Span, Direction)`, row 0 = header row) — horizontal merges via `w:gridSpan`, vertical merges via `w:vMerge` (restart/continue). Verified surviving real LibreOffice conversion with text intact (`TableLayoutConversionTests`). | Merges can't overlap each other (validated, throws if they do) — a cell can be part of at most one merge, so an L-shaped or nested merge region isn't supported in one call. |

### Watermark placement — `Watermarking/DocxWatermarkService.cs`, `Watermarking/PdfWatermarkService.cs`

| | Supported | Not supported |
|---|---|---|
| **Position** | `position: WatermarkPosition` on both services — `Center` (the prior fixed default) plus the 8 compass points (`TopLeft` … `BottomRight`). Docx maps directly to VML's own `mso-position-horizontal`/`-vertical` keywords; PDF insets off-center positions proportionally to page size, since the text is drawn rotated around its own anchor point. Verified with a real visual diff between placements (`WatermarkPositionAndSizeTests`), not just "it didn't throw." | Arbitrary absolute X/Y coordinates — 9 fixed positions only, no free placement. |
| **Size** | Docx: `widthPt`/`heightPt` (the shape's bounding box) and `fontSizePt`. PDF: `fontSizePt`. Both default to the prior fixed values (415×207.5pt box / 72pt text) when omitted. | PDF has no separate box-size concept (the box was always implicit, sized to fit the rasterized text) — only the font size is adjustable there. |
| **Everything else** | Text, font family, rotation angle, and color/opacity remain configurable as before. | Same treatment applied uniformly to every page — no per-page position/size variation within one call. |

All four phases from the original gap list are now closed. The only remaining edges are the ones
called out as genuine ceilings above (OTF/subsetting, per-section headers/footers/watermarks) rather
than open work.

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
