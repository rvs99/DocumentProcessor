using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.Samples;
using DocumentProcessor.Core.Security;

namespace DocumentProcessor.Tests.Security;

public class MacroValidationServiceTests : IDisposable
{
    private readonly string _path = TestFiles.NewTempPath(".docx");
    private readonly MacroValidationService _sut = new();

    private static void CreateMacroEnabledDocument(string path)
    {
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.MacroEnabledDocument);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("Body text.")))));
        mainPart.Document.Save();
    }

    [Fact]
    public void ContainsMacros_is_false_for_a_plain_document()
    {
        SampleDocumentFactory.CreateBasicDocument(_path, "Plain", ["Body text."]);

        Assert.False(_sut.ContainsMacros(_path));
    }

    [Fact]
    public void ContainsMacros_is_true_for_a_macro_enabled_document_type()
    {
        CreateMacroEnabledDocument(_path);

        Assert.True(_sut.ContainsMacros(_path));
    }

    [Fact]
    public void StripMacros_converts_the_document_type_back_to_plain()
    {
        CreateMacroEnabledDocument(_path);
        var outputPath = TestFiles.NewTempPath(".docx");

        try
        {
            _sut.StripMacros(_path, outputPath);

            Assert.False(_sut.ContainsMacros(outputPath));
            using var doc = WordprocessingDocument.Open(outputPath, isEditable: false);
            Assert.Equal(WordprocessingDocumentType.Document, doc.DocumentType);
        }
        finally
        {
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void ValidateTemplate_is_valid_for_a_normal_document()
    {
        SampleDocumentFactory.CreateBasicDocument(_path, "Plain", ["Body text."]);

        var result = _sut.ValidateTemplate(_path);

        Assert.True(result.IsValid);
        Assert.Empty(result.Issues);
    }

    [Fact]
    public void ValidateTemplate_is_invalid_for_a_missing_file()
    {
        var result = _sut.ValidateTemplate(TestFiles.NewTempPath(".docx"));

        Assert.False(result.IsValid);
        Assert.Contains(result.Issues, i => i.Contains("does not exist"));
    }

    [Fact]
    public void ValidateTemplate_is_invalid_for_a_macro_enabled_document()
    {
        CreateMacroEnabledDocument(_path);

        var result = _sut.ValidateTemplate(_path);

        Assert.False(result.IsValid);
        Assert.Contains(result.Issues, i => i.Contains("macro", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ValidateTemplate_is_invalid_for_a_corrupt_file()
    {
        File.WriteAllBytes(_path, [1, 2, 3, 4, 5]);

        var result = _sut.ValidateTemplate(_path);

        Assert.False(result.IsValid);
        Assert.NotEmpty(result.Issues);
    }

    public void Dispose() => File.Delete(_path);
}
