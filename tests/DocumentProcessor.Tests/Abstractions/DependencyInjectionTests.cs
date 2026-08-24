using DocumentProcessor.Core;
using DocumentProcessor.Core.ContentControls;
using DocumentProcessor.Core.Conversion;
using DocumentProcessor.Core.DocumentAssembly;
using DocumentProcessor.Core.Security;
using DocumentProcessor.Core.Redlining;
using DocumentProcessor.Core.Templating;
using DocumentProcessor.Core.TrackChanges;
using Microsoft.Extensions.DependencyInjection;

namespace DocumentProcessor.Tests.Abstractions;

/// <summary>
/// The two things interfaces plus DI registration are meant to unblock for a consuming application:
/// resolving services from its own container, and substituting a fake for anything with an external
/// dependency so its tests don't inherit ours.
/// </summary>
public class DependencyInjectionTests
{
    [Fact]
    public void Every_registered_service_resolves()
    {
        var provider = new ServiceCollection().AddDocumentProcessor().BuildServiceProvider();

        // A representative span rather than all 25: templating, review, conversion, and a
        // PDF-side service, which between them cover every registration shape used above.
        Assert.NotNull(provider.GetRequiredService<ITemplateEngine>());
        Assert.NotNull(provider.GetRequiredService<IContentControlService>());
        Assert.NotNull(provider.GetRequiredService<IDocumentComparisonService>());
        Assert.NotNull(provider.GetRequiredService<ITrackChangesService>());
        Assert.NotNull(provider.GetRequiredService<IWordToPdfConverter>());
        Assert.NotNull(provider.GetRequiredService<IPdfAssemblyService>());
        Assert.NotNull(provider.GetRequiredService<IMacroValidationService>());
    }

    [Fact]
    public void Conversion_options_flow_from_the_container_into_the_converter()
    {
        // The LibreOffice path necessarily differs between a dev machine and a container, so it has
        // to be configurable through the standard options pipeline rather than only in code.
        var provider = new ServiceCollection()
            .AddDocumentProcessor(o => o.ExecutablePath = "/custom/soffice")
            .BuildServiceProvider();

        var converter = provider.GetRequiredService<IWordToPdfConverter>();

        // Observable through behaviour: the configured path is what it tries to start.
        var docx = TestFiles.NewTempPath(".docx");
        var pdf = TestFiles.NewTempPath(".pdf");
        try
        {
            Core.Samples.SampleDocumentFactory.CreateBasicDocument(docx, "Contract", ["Body."]);
            var ex = Assert.Throws<ConversionUnavailableException>(() => converter.Convert(docx, pdf));
            Assert.Contains("/custom/soffice", ex.Message);
        }
        finally
        {
            File.Delete(docx);
            if (File.Exists(pdf)) File.Delete(pdf);
        }
    }

    [Fact]
    public void An_application_can_substitute_its_own_implementation()
    {
        // TryAdd semantics: a registration the application made itself must win, which is what
        // lets it swap in a fake without the library overwriting it.
        var services = new ServiceCollection();
        services.AddSingleton<IWordToPdfConverter>(new FakeConverter());
        services.AddDocumentProcessor();

        var resolved = services.BuildServiceProvider().GetRequiredService<IWordToPdfConverter>();

        Assert.IsType<FakeConverter>(resolved);
    }

    [Fact]
    public async Task Conversion_can_be_faked_so_consumer_tests_do_not_need_LibreOffice()
    {
        // The concrete point of IWordToPdfConverter: before it existed, any consumer test whose
        // code path touched conversion had to shell out to a real soffice install, so their CI
        // needed apt-get install libreoffice-writer to exercise their own business rules.
        var fake = new FakeConverter();
        IWordToPdfConverter converter = fake;

        await converter.ConvertAsync("contract.docx", "contract.pdf");

        Assert.Equal([("contract.docx", "contract.pdf")], fake.Calls);
    }

    private sealed class FakeConverter : IWordToPdfConverter
    {
        public List<(string Input, string Output)> Calls { get; } = [];

        public void Convert(string docxPath, string outputPdfPath) => Calls.Add((docxPath, outputPdfPath));

        public Task ConvertAsync(string docxPath, string outputPdfPath, CancellationToken cancellationToken = default)
        {
            Calls.Add((docxPath, outputPdfPath));
            return Task.CompletedTask;
        }
    }
}
