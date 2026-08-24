using DocumentProcessor.Core;
using DocumentProcessor.Core.Templating;

namespace DocumentProcessor.Tests.Templating;

public class TemplateConditionTests
{
    [Theory]
    [InlineData("Amount == 1000", "1000", true)]
    [InlineData("Amount == 1000", "999", false)]
    [InlineData("Amount != 1000", "999", true)]
    [InlineData("Amount > 500", "1000", true)]
    [InlineData("Amount > 500", "100", false)]
    [InlineData("Amount < 500", "100", true)]
    [InlineData("Amount >= 1000", "1000", true)]
    [InlineData("Amount <= 1000", "1000", true)]
    [InlineData("Amount <= 1000", "1001", false)]
    public void Evaluate_supports_all_six_operators_numerically(string expression, string actual, bool expected)
    {
        var condition = TemplateCondition.Parse(expression);
        Assert.Equal(expected, condition.Evaluate(actual));
    }

    [Theory]
    [InlineData("Status == \"Final\"", "Final", true)]
    [InlineData("Status == \"Final\"", "Draft", false)]
    [InlineData("Status != \"Final\"", "Draft", true)]
    public void Evaluate_falls_back_to_string_comparison_for_non_numeric_values(string expression, string actual, bool expected)
    {
        var condition = TemplateCondition.Parse(expression);
        Assert.Equal(expected, condition.Evaluate(actual));
    }

    [Fact]
    public void Evaluate_returns_false_for_a_missing_field_regardless_of_operator()
    {
        var condition = TemplateCondition.Parse("Amount == 1000");
        Assert.False(condition.Evaluate(null));
    }

    [Fact]
    public void Parse_rejects_a_malformed_expression()
    {
        // TemplateException, not FormatException: a malformed condition is the template author's
        // mistake, and a caller needs to tell that apart from an unrelated int.Parse failure.
        Assert.Throws<TemplateException>(() => TemplateCondition.Parse("not a condition"));
    }

    [Fact]
    public void Parse_extracts_field_operator_and_value()
    {
        var condition = TemplateCondition.Parse("Company.Region >= \"West\"");

        Assert.Equal("Company.Region", condition.FieldPath);
        Assert.Equal(ConditionOperator.GreaterThanOrEqual, condition.Operator);
        Assert.Equal("West", condition.ComparisonValue);
    }
}
