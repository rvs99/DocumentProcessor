using System.Globalization;
using System.Text.RegularExpressions;

namespace DocumentProcessor.Core.Templating;

public enum ConditionOperator { Equal, NotEqual, GreaterThan, LessThan, GreaterThanOrEqual, LessThanOrEqual }

/// <summary>
/// Parses and evaluates the condition inside a <c>{{if:attr op value}}</c> marker — e.g.
/// <c>{{if:Amount &gt;= 1000}}</c> or <c>{{if:Status == "Final"}}</c>. Six operators are supported:
/// <c>==</c>, <c>!=</c>, <c>&gt;</c>, <c>&lt;</c>, <c>&gt;=</c>, <c>&lt;=</c>.
/// </summary>
public sealed partial record TemplateCondition(string FieldPath, ConditionOperator Operator, string ComparisonValue)
{
    [GeneratedRegex(@"^\s*(?<field>[A-Za-z0-9_.]+)\s*(?<op>==|!=|>=|<=|>|<)\s*(?<value>.+?)\s*$")]
    private static partial Regex Pattern();

    public static TemplateCondition Parse(string expression)
    {
        var match = Pattern().Match(expression);
        if (!match.Success)
        {
            throw new FormatException(
                $"Malformed condition '{expression}'. Expected 'field op value' with op in ==, !=, >=, <=, >, <.");
        }

        var op = match.Groups["op"].Value switch
        {
            "==" => ConditionOperator.Equal,
            "!=" => ConditionOperator.NotEqual,
            ">" => ConditionOperator.GreaterThan,
            "<" => ConditionOperator.LessThan,
            ">=" => ConditionOperator.GreaterThanOrEqual,
            "<=" => ConditionOperator.LessThanOrEqual,
            var unknown => throw new FormatException($"Unsupported operator '{unknown}'.")
        };

        var value = match.Groups["value"].Value.Trim();
        if (value.Length >= 2 && value[0] == '"' && value[^1] == '"')
            value = value[1..^1];

        return new TemplateCondition(match.Groups["field"].Value, op, value);
    }

    /// <summary>
    /// Evaluates this condition against <paramref name="actualValue"/> (the field's resolved value,
    /// or <see langword="null"/> if the field path wasn't found — a missing field always evaluates to
    /// <see langword="false"/>, regardless of operator, since there's no meaningful "redact"/"highlight"
    /// for a control-flow decision).
    /// </summary>
    public bool Evaluate(object? actualValue)
    {
        if (actualValue is null)
            return false;

        var actualText = TemplateValueFormatter.ToComparableString(actualValue);

        // Prefer numeric comparison when both sides parse as numbers; fall back to ordinal string
        // comparison for ==/!= and lexicographic comparison otherwise (e.g. date strings, names).
        if (double.TryParse(actualText, NumberStyles.Float, CultureInfo.InvariantCulture, out var actualNum) &&
            double.TryParse(ComparisonValue, NumberStyles.Float, CultureInfo.InvariantCulture, out var expectedNum))
        {
            return Operator switch
            {
                ConditionOperator.Equal => actualNum == expectedNum,
                ConditionOperator.NotEqual => actualNum != expectedNum,
                ConditionOperator.GreaterThan => actualNum > expectedNum,
                ConditionOperator.LessThan => actualNum < expectedNum,
                ConditionOperator.GreaterThanOrEqual => actualNum >= expectedNum,
                ConditionOperator.LessThanOrEqual => actualNum <= expectedNum,
                _ => throw new NotSupportedException(Operator.ToString())
            };
        }

        var comparison = string.CompareOrdinal(actualText, ComparisonValue);
        return Operator switch
        {
            ConditionOperator.Equal => comparison == 0,
            ConditionOperator.NotEqual => comparison != 0,
            ConditionOperator.GreaterThan => comparison > 0,
            ConditionOperator.LessThan => comparison < 0,
            ConditionOperator.GreaterThanOrEqual => comparison >= 0,
            ConditionOperator.LessThanOrEqual => comparison <= 0,
            _ => throw new NotSupportedException(Operator.ToString())
        };
    }
}

internal static class TemplateValueFormatter
{
    public static string ToComparableString(object value) => value switch
    {
        string s => s,
        bool b => b ? "true" : "false",
        DateTime dt => dt.ToString("O", CultureInfo.InvariantCulture),
        DateTimeOffset dto => dto.ToString("O", CultureInfo.InvariantCulture),
        IFormattable f => f.ToString(null, CultureInfo.InvariantCulture),
        _ => value.ToString() ?? string.Empty
    };
}
