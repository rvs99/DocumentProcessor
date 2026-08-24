using DocumentFormat.OpenXml.Packaging;
using DocumentProcessor.Core.ContentControls;
using DocumentProcessor.Core.Diagnostics;
using DocumentProcessor.Core.Metadata;
using DocumentProcessor.Core.Security;
using DocumentProcessor.Core.Tables;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;

namespace DocumentProcessor.Core.Sessions;

/// <summary>
/// One open .docx package that many operations share, so a pipeline unzips, parses, serializes and
/// rezips the document <em>once</em> instead of once per step.
/// <para>
/// The path-based service APIs each open and save the whole package independently, which is fine
/// for a single call and badly wrong for a pipeline: a twelve-step assembly performed fifteen full
/// package cycles, and on a 200-page contract each cycle inflates ~15–25 MB of XML into a
/// 150–400 MB DOM before Deflate-compressing it back. Those services remain, unchanged, for
/// one-shot use; this is what you reach for when doing more than one thing to a document.
/// </para>
/// <example>
/// <code>
/// using var session = DocumentSession.Open(uploadedBytes);
/// session.ContentControls.ReplaceMany(fieldValues);
/// session.Tables.AppendTable(pricingSpec);
/// session.Metadata.SetCustomProperties(matterProperties);
/// byte[] filled = session.Save();
/// </code>
/// </example>
/// <para>
/// Not thread-safe: a session wraps a mutable DOM and is meant to be used by one logical operation
/// at a time. Concurrent requests each get their own session, which is the isolation that matters.
/// </para>
/// </summary>
public sealed class DocumentSession : IDisposable
{
    private readonly MemoryStream _stream;
    private readonly WordprocessingDocument _document;
    private readonly ILoggerFactory? _loggerFactory;
    private bool _disposed;

    private DocumentSession(MemoryStream stream, WordprocessingDocument document, ILoggerFactory? loggerFactory)
    {
        _stream = stream;
        _document = document;
        _loggerFactory = loggerFactory;

        ContentControls = new ContentControlOperations(this, Logger<ContentControlService>());
        Metadata = new DocumentMetadataOperations(this);
        Tables = new TableOperations(this);
        TrackChanges = new TrackChangesOperations(this);
        PageLayout = new PageLayoutOperations(this);
        Watermark = new WatermarkOperations(this);
        Protection = new DocumentProtectionOperations(this);
        Fonts = new FontOperations(this);
        Fields = new FieldOperations(this);
        Comments = new CommentOperations(this);
        Text = new TextOperations(this);
    }

    /// <summary>
    /// Opens a document held in memory — a web upload, a blob-store read — with no filesystem
    /// involvement at any point.
    /// <para>
    /// The package is checked against <paramref name="limits"/> before it is parsed. This is the
    /// entry point untrusted tenant uploads come through, so the bound is applied by default rather
    /// than being something a caller has to remember; pass <see cref="DocumentLimits.Unbounded"/>
    /// for documents this system produced itself.
    /// </para>
    /// </summary>
    /// <exception cref="DocumentTooComplexException">The package exceeds <paramref name="limits"/>.</exception>
    /// <exception cref="CorruptDocumentException">The bytes are not a readable Office package.</exception>
    public static DocumentSession Open(byte[] docxBytes, ILoggerFactory? loggerFactory = null, DocumentLimits? limits = null)
    {
        ArgumentNullException.ThrowIfNull(docxBytes);
        DocumentPackageGuard.Validate(docxBytes, limits);

        // Capacity is pre-sized: the package grows as parts are edited, and letting a
        // zero-capacity MemoryStream double its way up from nothing churns the large object heap
        // for the whole length of the document.
        var stream = new MemoryStream(docxBytes.Length + 64 * 1024);
        stream.Write(docxBytes, 0, docxBytes.Length);
        stream.Position = 0;

        return OpenStreamOwned(stream, loggerFactory);
    }

    /// <summary>Copies <paramref name="docxStream"/> into memory and opens it. The source stream is
    /// not retained, and is left at whatever position reading finished.</summary>
    public static DocumentSession Open(Stream docxStream, ILoggerFactory? loggerFactory = null, DocumentLimits? limits = null)
    {
        ArgumentNullException.ThrowIfNull(docxStream);

        var stream = new MemoryStream();
        docxStream.CopyTo(stream);

        // Validated from the buffered copy rather than the caller's stream, which may not be
        // seekable and must not be consumed twice.
        DocumentPackageGuard.Validate(stream.ToArray(), limits);
        stream.Position = 0;

        return OpenStreamOwned(stream, loggerFactory);
    }

    /// <summary>Reads a document from disk into memory and opens it. The file is closed
    /// immediately — nothing holds a handle for the life of the session.</summary>
    public static DocumentSession OpenFile(string docxPath, ILoggerFactory? loggerFactory = null, DocumentLimits? limits = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(docxPath);
        if (!File.Exists(docxPath))
            throw new FileNotFoundException("Input .docx file not found.", docxPath);

        // Checked against the file's own length before reading it in, so an oversized file is
        // rejected without being loaded into memory first.
        DocumentPackageGuard.ValidateFile(docxPath, limits);
        return Open(File.ReadAllBytes(docxPath), loggerFactory, DocumentLimits.Unbounded);
    }

    private static DocumentSession OpenStreamOwned(MemoryStream stream, ILoggerFactory? loggerFactory)
    {
        try
        {
            var document = WordprocessingDocument.Open(stream, isEditable: true);
            return new DocumentSession(stream, document, loggerFactory);
        }
        catch
        {
            // The stream is this method's responsibility until the session owns it.
            stream.Dispose();
            throw;
        }
    }

    /// <summary>The underlying package, for operations this session doesn't yet wrap. Prefer the
    /// typed operation groups; reach for this when you need something they don't cover, rather than
    /// dropping back to a path-based service and paying for a second full package cycle.</summary>
    public WordprocessingDocument Document
    {
        get
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            return _document;
        }
    }

    /// <summary>Content-control (structured document tag) operations.</summary>
    public ContentControlOperations ContentControls { get; }

    /// <summary>Core and custom document properties.</summary>
    public DocumentMetadataOperations Metadata { get; }

    /// <summary>Table generation and prototype-row population.</summary>
    public TableOperations Tables { get; }

    /// <summary>Accepting and rejecting tracked changes, and reading them as data.</summary>
    public TrackChangesOperations TrackChanges { get; }

    /// <summary>Page size, margins, columns, and section/page breaks.</summary>
    public PageLayoutOperations PageLayout { get; }

    /// <summary>Adding and removing the docx watermark.</summary>
    public WatermarkOperations Watermark { get; }

    /// <summary>Editing restrictions and permitted ranges.</summary>
    public DocumentProtectionOperations Protection { get; }

    /// <summary>Embedding font families and applying them.</summary>
    public FontOperations Fonts { get; }

    /// <summary>Field dirtying, update-on-open, and cross-reference validation.</summary>
    public FieldOperations Fields { get; }

    /// <summary>Reviewer comments: reading the thread, replying, resolving, deleting.</summary>
    public CommentOperations Comments { get; }

    /// <summary>Document text as data, for search indexing and clause-level review.</summary>
    public TextOperations Text { get; }

    /// <summary>Flushes every pending change into the package and returns it as bytes. The session
    /// stays usable afterwards, so this can be called more than once (e.g. to snapshot an
    /// intermediate version) at the cost of one serialization each time.</summary>
    public byte[] Save()
    {
        Flush();
        return _stream.ToArray();
    }

    /// <summary>Flushes and writes the package to <paramref name="path"/>.</summary>
    public void SaveAs(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        Flush();

        var directory = Path.GetDirectoryName(Path.GetFullPath(path));
        if (!string.IsNullOrEmpty(directory))
            Directory.CreateDirectory(directory);

        _stream.Position = 0;
        using var file = File.Create(path);
        _stream.CopyTo(file);
    }

    /// <summary>Flushes and copies the package into <paramref name="destination"/>.</summary>
    public void CopyTo(Stream destination)
    {
        ArgumentNullException.ThrowIfNull(destination);
        Flush();

        _stream.Position = 0;
        _stream.CopyTo(destination);
    }

    private void Flush()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        using var activity = DocumentProcessorDiagnostics.ActivitySource.StartActivity("DocumentSession.Save");
        _document.Save();
    }

    internal ILogger<T> Logger<T>() =>
        _loggerFactory?.CreateLogger<T>() ?? NullLogger<T>.Instance;

    public void Dispose()
    {
        if (_disposed)
            return;

        _disposed = true;
        _document.Dispose();
        _stream.Dispose();
    }
}
