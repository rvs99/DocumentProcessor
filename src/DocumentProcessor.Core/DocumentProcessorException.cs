namespace DocumentProcessor.Core;

/// <summary>
/// Base type for every error this library raises deliberately.
/// <para>
/// The point of the hierarchy is that a caller can decide what to <em>do</em> without reading
/// exception messages. Before it existed, a corrupt customer upload, a missing LibreOffice install,
/// and a crashed conversion all arrived as <see cref="InvalidOperationException"/> with different
/// text, so an API layer had to string-match to choose between 400, 500 and 503-with-retry.
/// </para>
/// <para>
/// Deriving from <see cref="Exception"/> rather than <see cref="InvalidOperationException"/> is
/// deliberate: inheriting from it would mean a caller's existing
/// <c>catch (InvalidOperationException)</c> — typically mapped to a 500 — silently swallowed errors
/// that are really the caller's own bad input.
/// </para>
/// </summary>
public abstract class DocumentProcessorException : Exception
{
    protected DocumentProcessorException(string message) : base(message) { }
    protected DocumentProcessorException(string message, Exception innerException) : base(message, innerException) { }

    /// <summary>
    /// Whether retrying the identical operation could plausibly succeed. False for anything caused
    /// by the input or the request itself — retrying a malformed document just fails again — and
    /// true only for transient infrastructure failures.
    /// </summary>
    public virtual bool IsRetryable => false;
}

/// <summary>
/// The supplied document could not be read as a valid OOXML/PDF package, or is missing parts the
/// operation requires. Caused by the input, so it maps to a client error and must not be retried.
/// </summary>
public sealed class CorruptDocumentException : DocumentProcessorException
{
    public CorruptDocumentException(string message) : base(message) { }
    public CorruptDocumentException(string message, Exception innerException) : base(message, innerException) { }

    /// <summary>Convenience for the most common case: a package that opened but has no usable body.</summary>
    public static CorruptDocumentException MissingBody() =>
        new("Document has no main part or body — it is not a readable Word document.");
}

/// <summary>
/// The input is well-formed but larger or more deeply nested than the configured limits allow.
/// Separate from <see cref="CorruptDocumentException"/> because the document may be perfectly valid;
/// it is the size or shape that is refused, and the caller may be able to act on that distinction
/// (e.g. surfacing "this file is too large" rather than "this file is broken").
/// </summary>
public class DocumentTooComplexException : DocumentProcessorException
{
    public DocumentTooComplexException(string message) : base(message) { }
}

/// <summary>
/// A template could not be processed because of how it is authored — unbalanced block markers, an
/// unparseable condition, a clause id the library does not contain. The caller's data or template
/// is at fault, not the document engine.
/// </summary>
public class TemplateException : DocumentProcessorException
{
    public TemplateException(string message) : base(message) { }
    public TemplateException(string message, Exception innerException) : base(message, innerException) { }
}

/// <summary>
/// The external converter could not be started at all — typically LibreOffice is not installed, or
/// the configured executable path is wrong. An environment/deployment fault: it will keep failing
/// until someone fixes the host, so retrying is pointless and this should page an operator.
/// </summary>
public sealed class ConversionUnavailableException : DocumentProcessorException
{
    public ConversionUnavailableException(string message) : base(message) { }
    public ConversionUnavailableException(string message, Exception innerException) : base(message, innerException) { }
}

/// <summary>
/// The converter ran but did not produce a usable result — a non-zero exit, or success reported
/// with no output file. Usually transient (resource pressure, a wedged worker), so
/// <see cref="IsRetryable"/> is true.
/// </summary>
public sealed class ConversionFailedException : DocumentProcessorException
{
    public ConversionFailedException(string message) : base(message) { }
    public ConversionFailedException(string message, Exception innerException) : base(message, innerException) { }

    /// <inheritdoc/>
    public override bool IsRetryable => true;
}
