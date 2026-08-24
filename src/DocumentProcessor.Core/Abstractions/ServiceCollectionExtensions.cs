using DocumentProcessor.Core.Comparison;
using DocumentProcessor.Core.ContentControls;
using DocumentProcessor.Core.Conversion;
using DocumentProcessor.Core.DocumentAssembly;
using DocumentProcessor.Core.ESign;
using DocumentProcessor.Core.FontEmbedding;
using DocumentProcessor.Core.Format;
using DocumentProcessor.Core.Layout;
using DocumentProcessor.Core.Metadata;
using DocumentProcessor.Core.Redlining;
using DocumentProcessor.Core.Security;
using DocumentProcessor.Core.Tables;
using DocumentProcessor.Core.Templating;
using DocumentProcessor.Core.TrackChanges;
using DocumentProcessor.Core.Transplant;
using DocumentProcessor.Core.Watermarking;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using Microsoft.Extensions.Options;

namespace DocumentProcessor.Core;

/// <summary>
/// Registers the library's services with a dependency-injection container, so a consuming
/// application does not have to hand-wire two dozen concrete types and their options.
/// </summary>
public static class ServiceCollectionExtensions
{
    /// <summary>
    /// Registers every document-processing service against its interface.
    /// <para>
    /// Services are registered as singletons: they hold no per-call mutable state — configuration
    /// and an optional logger are the only fields any of them carry — so a single instance is safe
    /// to share across concurrent requests. Per-document state lives in
    /// <see cref="Sessions.DocumentSession"/>, which callers create and dispose per operation.
    /// </para>
    /// <para>
    /// <see cref="TryAdd"/> semantics throughout: an application that has already registered its own
    /// implementation of any interface keeps it, so this can be called safely alongside custom
    /// substitutions.
    /// </para>
    /// </summary>
    /// <param name="services">The container to register into.</param>
    /// <param name="configureConversion">
    /// Optional conversion settings — most importantly the LibreOffice executable path, which
    /// necessarily differs between a developer machine and a Linux container. Omit this to bind
    /// <see cref="WordToPdfConversionOptions"/> from configuration instead (see the overload
    /// remarks), or to accept the platform defaults.
    /// </param>
    public static IServiceCollection AddDocumentProcessor(
        this IServiceCollection services,
        Action<WordToPdfConversionOptions>? configureConversion = null)
    {
        ArgumentNullException.ThrowIfNull(services);

        if (configureConversion is not null)
            services.Configure(configureConversion);

        // The converters take a plain options object rather than IOptions<T>, so they stay usable
        // without a container at all. Bridging here gives DI consumers the standard configuration
        // pipeline (appsettings.json, environment variables) without forcing that dependency on
        // everyone else.
        services.TryAddSingleton<IWordToPdfConverter>(sp =>
            new WordToPdfConverter(sp.GetService<IOptions<WordToPdfConversionOptions>>()?.Value));

        services.TryAddSingleton<ILegacyDocConverter>(sp =>
            new LegacyDocConverter(sp.GetService<IOptions<LegacyDocConversionOptions>>()?.Value));

        // Templating and assembly
        services.TryAddSingleton<ITemplateEngine, TemplateEngine>();
        services.TryAddSingleton<IContentControlService, ContentControlService>();
        services.TryAddSingleton<ITableGenerationService, TableGenerationService>();
        services.TryAddSingleton<IClauseTransplantService, ClauseTransplantService>();
        services.TryAddSingleton<ICrossReferenceValidator, CrossReferenceValidator>();
        services.TryAddSingleton<IFieldUpdateService, FieldUpdateService>();
        services.TryAddSingleton<IPdfAssemblyService, PdfAssemblyService>();

        // Review and negotiation
        services.TryAddSingleton<IDocumentComparisonService, DocumentComparisonService>();
        services.TryAddSingleton<ITrackChangesService, TrackChangesService>();
        services.TryAddSingleton<IRedlineExportService, RedlineExportService>();
        services.TryAddSingleton<IDocumentProtectionService, DocumentProtectionService>();

        // Presentation
        services.TryAddSingleton<IHeaderFooterService, HeaderFooterService>();
        services.TryAddSingleton<IPageLayoutService, PageLayoutService>();
        services.TryAddSingleton<IBrandingService, BrandingService>();
        services.TryAddSingleton<IDocxWatermarkService, DocxWatermarkService>();
        services.TryAddSingleton<IPdfWatermarkService, PdfWatermarkService>();
        services.TryAddSingleton<IFontEmbeddingService, FontEmbeddingService>();

        // Metadata, security, and inspection
        services.TryAddSingleton<IDocumentMetadataService, DocumentMetadataService>();
        services.TryAddSingleton<IPdfMetadataService, PdfMetadataService>();
        services.TryAddSingleton<IMacroValidationService, MacroValidationService>();
        services.TryAddSingleton<IPdfProtectionService, PdfProtectionService>();
        services.TryAddSingleton<IPdfComparisonService, PdfComparisonService>();
        services.TryAddSingleton<IESignFieldService, ESignFieldService>();

        return services;
    }
}
