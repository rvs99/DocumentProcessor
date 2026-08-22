using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.DocumentAssembly;
using DocumentProcessor.Core.Samples;

namespace DocumentProcessor.Tests.DocumentAssembly;

public class CrossReferenceValidatorTests : IDisposable
{
    private readonly string _path = TestFiles.NewTempPath(".docx");
    private readonly CrossReferenceValidator _sut = new();

    [Fact]
    public void Validate_returns_empty_when_every_reference_resolves()
    {
        SampleDocumentFactory.CreateDocumentFromParagraphs(_path,
        [
            new Paragraph(new BookmarkStart { Id = "1", Name = "_Ref100" }, new Run(new Text("Section 2 heading")), new BookmarkEnd { Id = "1" }),
            new Paragraph(new SimpleField(new Run(new Text("Section 2"))) { Instruction = " REF _Ref100 \\h " })
        ]);

        Assert.Empty(_sut.Validate(_path));
    }

    [Fact]
    public void Validate_flags_a_simple_field_reference_to_a_missing_bookmark()
    {
        SampleDocumentFactory.CreateDocumentFromParagraphs(_path,
        [
            new Paragraph(new SimpleField(new Run(new Text("Section 2"))) { Instruction = " REF _Ref999 \\h " })
        ]);

        var dangling = _sut.Validate(_path);

        Assert.Single(dangling);
        Assert.Equal("_Ref999", dangling[0].BookmarkName);
        Assert.Equal("REF", dangling[0].FieldType);
    }

    [Fact]
    public void Validate_flags_a_complex_field_reference_to_a_missing_bookmark()
    {
        SampleDocumentFactory.CreateDocumentFromParagraphs(_path,
        [
            new Paragraph(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" PAGEREF _Ref200 \\h ")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text("5")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }))
        ]);

        var dangling = _sut.Validate(_path);

        Assert.Single(dangling);
        Assert.Equal("_Ref200", dangling[0].BookmarkName);
        Assert.Equal("PAGEREF", dangling[0].FieldType);
    }

    public void Dispose() => File.Delete(_path);
}
