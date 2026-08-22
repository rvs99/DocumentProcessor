namespace DocumentProcessor.Core.Templating;

/// <summary>
/// Resolves dotted field paths (e.g. <c>Company.Name</c>) against a data dictionary, with a parent
/// scope fallback — used so tokens inside a <c>{{repeat:Parties}}</c> body can reference both the
/// current item's fields and outer/top-level fields (e.g. a global <c>CompanyName</c>) by the same
/// <c>{{token}}</c> syntax.
/// </summary>
internal sealed class TemplateContext(IReadOnlyDictionary<string, object?> scope, TemplateContext? parent = null)
{
    public bool TryResolve(string fieldPath, out object? value)
    {
        var parts = fieldPath.Split('.');
        if (TryResolveInScope(scope, parts, out value))
            return true;

        return parent is not null && parent.TryResolve(fieldPath, out value);
    }

    public TemplateContext Push(IReadOnlyDictionary<string, object?> childScope) => new(childScope, this);

    private static bool TryResolveInScope(IReadOnlyDictionary<string, object?> scope, string[] parts, out object? value)
    {
        object? current = scope;
        foreach (var part in parts)
        {
            if (current is IReadOnlyDictionary<string, object?> dict && dict.TryGetValue(part, out var next))
                current = next;
            else
            {
                value = null;
                return false;
            }
        }

        value = current;
        return true;
    }
}
