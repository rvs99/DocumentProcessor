using System.Diagnostics;

namespace DocumentProcessor.Core.Diagnostics;

/// <summary>
/// Shared telemetry source for the library. A single <see cref="ActivitySource"/> per library (not
/// per class) is the standard .NET convention — consumers listen for
/// <c>"DocumentProcessor.Core"</c> once and see every operation, rather than needing to know each
/// service's name up front. No NuGet package is needed for this: <see cref="ActivitySource"/>/
/// <see cref="Activity"/> are part of the base class library (<c>System.Diagnostics</c>), and produce
/// no overhead at all when nothing is listening (the OpenTelemetry SDK, Application Insights, or any
/// other <see cref="System.Diagnostics.DiagnosticSource"/> listener — or none) — <see cref="StartActivity"/>
/// short-circuits to a no-op when there's no listener, so services can call it unconditionally.
/// </summary>
public static class DocumentProcessorDiagnostics
{
    public const string SourceName = "DocumentProcessor.Core";

    public static readonly ActivitySource ActivitySource = new(SourceName);
}
