# Backlog and known issues

Work that has been deliberately deferred, and constraints that have been investigated and found
real. The point of this file is that nobody re-derives a conclusion that has already been paid for
— each entry records what was measured or checked, so a future decision starts from evidence rather
than from scratch.

Status meanings:

- **Deferred** — achievable, not currently justified. Has a trigger that would change that.
- **Blocked** — not achievable without changing a dependency or a stated constraint.
- **Known issue** — something is wrong or unverified today.

---

## Deferred

### PDF/A-2b archival conformance

**Trigger to revisit:** a customer contract or regulator requiring archival-format retention.

PDFsharp has no built-in PDF/A mode. Reaching PDF/A-2b means hand-constructing XMP metadata and an
`OutputIntent` dictionary through PDFsharp's low-level `PdfDictionary` API, guaranteeing every font
is fully embedded (already possible — no subsetting, so files are larger but conformant), and
avoiding transparency and other constructs PDF/A forbids. Validation would use veraPDF, which is a
Java CLI tool rather than a NuGet package, so it becomes a CI tool dependency.

Deferred by explicit decision when the feature checklist was scoped: the whole "Word→PDF Conversion
& Archival" section was excluded. Flagged here rather than dropped because contract retention is a
common CLM legal requirement, so the trigger is plausible rather than hypothetical.

Effort: substantial. Not a weekend item.

### A UNO sidecar for conversion

**Trigger to revisit:** conversion throughput becoming the bottleneck again after batching, i.e. a
workload dominated by *single* small conversions that cannot be batched.

LibreOffice spends roughly 450 ms starting up before it looks at the first document. Batching and
profile pooling (see `Conversion/ConversionQueue.cs`, `Conversion/LibreOfficeProfilePool.cs`) removed
that cost for grouped and concurrent work, but a lone conversion on an idle host still pays it.

Two cheaper alternatives were measured and **ruled out**:

- **Keeping a persistent `soffice` alive on a shared profile does nothing.** A subsequent
  `--convert-to` invocation still bootstraps its own process: 380 ms with a live instance against
  400 ms without one. It does not delegate to the running instance.
- **Profile pooling alone** is worth only ~80–100 ms of a ~550 ms conversion. It is implemented, but
  it is not the lever that removes the bootstrap.

The only approach that actually eliminates the 450 ms is a genuine UNO sidecar — one persistent
`soffice --accept=socket,...` process driven over UNO, the `unoserver` architecture. .NET has no
usable UNO bridge, so the client would be a Python script running against LibreOffice's bundled
`python3-uno`, talking to this library over stdin/stdout. That adds a Python script asset, a
`python3-uno` package requirement in the container, and a socket health-check/restart loop for when
the instance wedges.

Deferred because the deployment and failure surface is a poor trade for ~450 ms on the one workload
shape that batching does not already cover.

---

## Blocked

### PDF/UA accessibility tagging

Structure-tree tagging requires marked content tied to a structure tree across every content stream
PDFsharp emits, and PDFsharp's public API has no meaningful support for it. Doing this properly
means a different PDF library — realistically iText7, which is AGPL or commercial, in direct
tension with the permissive-licensing constraint the rest of the stack is built on.

This is the one item on the original feature checklist where the honest answer is "not achievable
here", not "not yet built". Revisiting means accepting a licensing change, not writing more code.

### Multi-targeting `net8.0`

Clippit is `net10.0`-only across every version checked (3.6.0 through 3.9.0), and it provides both
`WmlComparer` (redlining) and `DocumentBuilder` (cross-document style and numbering remapping) —
neither of which has a viable free replacement. Multi-targeting would mean shipping a `net8.0`
target with redlining and clause assembly silently missing, which is worse than requiring
`net10.0`.

Recorded in `src/DocumentProcessor.Core/DocumentProcessor.Core.csproj` alongside the
`TargetFramework` element. Revisit only if Clippit broadens its target frameworks.

### Single dependency / no subprocess

Free, permissively licensed, Word-comparable docx→PDF fidelity, pure managed code, one dependency —
these are not simultaneously satisfiable in today's .NET ecosystem. OnlyOffice was rejected on
licensing (a dependency's copyright metadata claims "Commercial" despite the project's public
AGPL), and Chromium print-to-PDF was rejected on fidelity (worse pagination than LibreOffice, plus
new text-rendering artifacts). The commercial renderers that would satisfy it — Aspose.Words,
Syncfusion DocIO, GemBox.Document — cost money and are still a dependency.

Satisfying this means giving up either the free-only constraint or the fidelity bar. It is not an
engineering-effort problem.

---

## Known issues

### The legacy `.doc` conversion path is unverified

`LegacyDocConverterTests.ConvertToDocx_produces_a_docx_the_OpenXml_pipeline_can_open` fails, and
will keep failing as written. Its fixture is an OLE Compound File signature followed by 512 zero
bytes — a file that *looks* like a `.doc` but contains no document, so LibreOffice correctly reports
"source file could not be loaded". The demo reports the same step as skipped.

The conversion code itself is almost certainly fine — it is the same `LibreOfficeRunner` path every
other conversion uses, differing only in the target extension — but that is an argument, not
evidence. Fixing this needs a real, small, redistributable `.doc` checked in as a test asset,
generated from a machine that has Word or a LibreOffice able to save the legacy format.

Until then: **one failing test is the expected state of the suite.**

### Untested OOXML surface area

Headers and footers, embedded images and charts, nested tables, RTL text, formatting-only tracked
changes (bold/italic toggles with no text inserted or deleted), moved-text redlining, and
multi-signer e-sign flows are all unexercised. The demo walks a realistic contract-editing path, not
the full OOXML surface. See the "Conversion fidelity" section of `README.md` for the fuller
accounting.

---

## Test baselines

The suite's expected state, so a regression is distinguishable from the environment:

| Environment | Expected |
|---|---|
| `DOCPROC_LIBREOFFICE_WSL_DISTRO=Ubuntu` set (or native `soffice` installed) | **330 passed, 1 failed** — the failure is the legacy `.doc` fixture above |
| No LibreOffice reachable | **291 passed, 40 failed** — every conversion test fails to start `soffice` |

Any other number is a real regression. Note that the conversion tests are skipped by *failing*, not
by `Skip`, which is why the second row exists at all.
