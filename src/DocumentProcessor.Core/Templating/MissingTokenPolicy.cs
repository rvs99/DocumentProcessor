namespace DocumentProcessor.Core.Templating;

/// <summary>
/// How <see cref="TemplateEngine.Fill"/> handles a <c>{{token}}</c> whose field path isn't present
/// in the supplied data (or resolves to <see langword="null"/>).
/// </summary>
public enum MissingTokenPolicy
{
    /// <summary>Throw a <see cref="MissingTemplateTokenException"/> naming the token and its field path.</summary>
    Error,

    /// <summary>Replace the token with an empty string, silently.</summary>
    Redact,

    /// <summary>Replace the token with its own literal text (e.g. <c>{{Company.Name}}</c>), wrapped
    /// in a yellow-highlighted run so it's easy to spot when reviewing the filled document.</summary>
    Highlight
}

/// <summary>Thrown by <see cref="MissingTokenPolicy.Error"/> when a token's field path has no value.</summary>
public sealed class MissingTemplateTokenException(string fieldPath)
    : InvalidOperationException($"Template token references unresolved field path '{fieldPath}'.")
{
    public string FieldPath { get; } = fieldPath;
}
