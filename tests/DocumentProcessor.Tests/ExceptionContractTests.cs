using DocumentProcessor.Core;
using DocumentProcessor.Core.Conversion;
using DocumentProcessor.Core.Samples;
using DocumentProcessor.Core.Templating;

namespace DocumentProcessor.Tests;

/// <summary>
/// Pins the error contract a calling API layer depends on. The point of the hierarchy is that a
/// caller can pick a status code and a retry policy from the exception <em>type</em>, without
/// reading messages — previously a corrupt upload, a missing LibreOffice install and a crashed
/// conversion were all <see cref="InvalidOperationException"/> with different text.
/// </summary>
public class ExceptionContractTests
{
    [Fact]
    public void Library_errors_do_not_masquerade_as_framework_errors()
    {
        // Deriving from Exception rather than InvalidOperationException is the whole point: a
        // caller's catch(InvalidOperationException) -> 500 must not swallow their own bad input.
        Assert.False(typeof(InvalidOperationException).IsAssignableFrom(typeof(DocumentProcessorException)));
        Assert.True(typeof(Exception).IsAssignableFrom(typeof(DocumentProcessorException)));
    }

    [Theory]
    [InlineData(typeof(CorruptDocumentException))]
    [InlineData(typeof(DocumentTooComplexException))]
    [InlineData(typeof(TemplateException))]
    [InlineData(typeof(MissingTemplateTokenException))]
    [InlineData(typeof(HtmlTooComplexException))]
    [InlineData(typeof(ConversionFailedException))]
    [InlineData(typeof(ConversionUnavailableException))]
    public void Every_library_exception_is_catchable_through_the_common_base(Type exceptionType)
    {
        Assert.True(typeof(DocumentProcessorException).IsAssignableFrom(exceptionType));
    }

    [Fact]
    public void Template_authoring_errors_are_distinguishable_from_engine_faults()
    {
        // A malformed condition and a missing merge field are both the caller's problem (400), and
        // both are reachable through one catch.
        Assert.Throws<TemplateException>(() => TemplateCondition.Parse("this is not a condition"));
        Assert.True(typeof(TemplateException).IsAssignableFrom(typeof(MissingTemplateTokenException)));
    }

    [Fact]
    public void Only_transient_conversion_failures_are_marked_retryable()
    {
        // The retry decision is on the type, not in the message.
        Assert.True(new ConversionFailedException("worker died").IsRetryable);
        Assert.False(new ConversionUnavailableException("not installed").IsRetryable);
        Assert.False(new CorruptDocumentException("bad upload").IsRetryable);
        Assert.False(new DocumentTooComplexException("too big").IsRetryable);
    }

    [Fact]
    public void A_missing_converter_reports_unavailable_rather_than_a_raw_Win32Exception()
    {
        var converter = new WordToPdfConverter(new WordToPdfConversionOptions
        {
            ExecutablePath = Path.Combine(Path.GetTempPath(), $"definitely-not-soffice-{Guid.NewGuid():N}.exe"),
        });

        var docxPath = TestFiles.NewTempPath(".docx");
        var pdfPath = TestFiles.NewTempPath(".pdf");
        try
        {
            SampleDocumentFactory.CreateBasicDocument(docxPath, "Contract", ["Body."]);

            var ex = Assert.Throws<ConversionUnavailableException>(() => converter.Convert(docxPath, pdfPath));
            Assert.False(ex.IsRetryable, "a missing install fails identically on every retry");
            Assert.Contains("LibreOffice", ex.Message);
        }
        finally
        {
            File.Delete(docxPath);
            if (File.Exists(pdfPath)) File.Delete(pdfPath);
        }
    }

    [Fact]
    public void A_corrupt_document_is_reported_as_corrupt()
    {
        var path = TestFiles.NewTempPath(".docx");
        try
        {
            File.WriteAllBytes(path, [0x50, 0x4B, 0x03, 0x04, 0xFF, 0xFF, 0xFF, 0xFF]);   // ZIP magic, then garbage

            Assert.ThrowsAny<DocumentProcessorException>(() =>
                DocumentProcessor.Core.Security.DocumentPackageGuard.ValidateFile(path));
        }
        finally
        {
            File.Delete(path);
        }
    }
}
