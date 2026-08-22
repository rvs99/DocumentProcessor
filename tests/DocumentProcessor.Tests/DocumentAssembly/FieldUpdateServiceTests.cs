using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.DocumentAssembly;
using DocumentProcessor.Core.Samples;

namespace DocumentProcessor.Tests.DocumentAssembly;

public class FieldUpdateServiceTests : IDisposable
{
    private readonly string _path = TestFiles.NewTempPath(".docx");
    private readonly FieldUpdateService _sut = new();

    public FieldUpdateServiceTests()
    {
        SampleDocumentFactory.CreateDocumentFromParagraphs(_path,
        [
            new Paragraph(new SimpleField(new Run(new Text("Section 2"))) { Instruction = " REF _Ref100 \\h " }),
            new Paragraph(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" PAGEREF _Ref200 \\h ")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text("5")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }))
        ]);
    }

    [Fact]
    public void MarkAllFieldsDirty_sets_dirty_on_both_simple_and_complex_fields()
    {
        _sut.MarkAllFieldsDirty(_path);

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var body = doc.MainDocumentPart!.Document!.Body!;

        var simpleField = body.Descendants<SimpleField>().Single();
        Assert.True(simpleField.Dirty!.Value);

        var beginFieldChar = body.Descendants<FieldChar>().Single(f => f.FieldCharType?.Value == FieldCharValues.Begin);
        Assert.True(beginFieldChar.Dirty!.Value);
    }

    [Fact]
    public void SetUpdateFieldsOnOpen_sets_the_setting_and_clearing_removes_it()
    {
        _sut.SetUpdateFieldsOnOpen(_path, true);
        Assert.True(ReadUpdateFieldsOnOpen(_path));

        _sut.SetUpdateFieldsOnOpen(_path, false);
        Assert.False(ReadUpdateFieldsOnOpen(_path));
    }

    private static bool ReadUpdateFieldsOnOpen(string path)
    {
        using var doc = WordprocessingDocument.Open(path, isEditable: false);
        return doc.MainDocumentPart?.DocumentSettingsPart?.Settings?.Elements<UpdateFieldsOnOpen>().Any() ?? false;
    }

    public void Dispose() => File.Delete(_path);
}
