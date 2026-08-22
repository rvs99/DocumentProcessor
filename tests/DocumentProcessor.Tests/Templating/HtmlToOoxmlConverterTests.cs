using DocumentProcessor.Core.Templating;

namespace DocumentProcessor.Tests.Templating;

public class HtmlToOoxmlConverterTests
{
    [Fact]
    public void Sanitize_removes_script_tags_and_their_content()
    {
        var result = HtmlToOoxmlConverter.Sanitize("<p>Hello <script>alert('xss')</script>World</p>");

        Assert.DoesNotContain("script", result, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("alert", result);
        Assert.Contains("Hello", result);
        Assert.Contains("World", result);
    }

    [Fact]
    public void Sanitize_strips_event_handler_attributes()
    {
        var result = HtmlToOoxmlConverter.Sanitize("<p onclick=\"doEvil()\">Hi</p>");

        Assert.DoesNotContain("onclick", result, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("doEvil", result);
    }

    [Fact]
    public void Sanitize_strips_a_javascript_url_but_keeps_a_safe_one()
    {
        var maliciousResult = HtmlToOoxmlConverter.Sanitize("<a href=\"javascript:alert(1)\">bad</a>");
        Assert.DoesNotContain("javascript:", maliciousResult, StringComparison.OrdinalIgnoreCase);

        var safeResult = HtmlToOoxmlConverter.Sanitize("<a href=\"https://example.com\">good</a>");
        Assert.Contains("https://example.com", safeResult);
    }

    [Fact]
    public void Sanitize_unwraps_disallowed_tags_but_keeps_their_text()
    {
        var result = HtmlToOoxmlConverter.Sanitize("<custom-tag>Hello</custom-tag>");

        Assert.Contains("Hello", result);
        Assert.DoesNotContain("custom-tag", result);
    }
}
