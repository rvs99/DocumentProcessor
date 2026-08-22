using DocumentProcessor.Core.Metadata;
using DocumentProcessor.Core.Samples;

namespace DocumentProcessor.Tests.Metadata;

public class DocumentMetadataServiceTests : IDisposable
{
    private readonly string _path = TestFiles.NewTempPath(".docx");
    private readonly DocumentMetadataService _sut = new();

    public DocumentMetadataServiceTests()
    {
        SampleDocumentFactory.CreateBasicDocument(_path, "Metadata Test", ["Body text."]);
    }

    [Fact]
    public void SetCustomProperty_string_roundtrips_through_GetCustomProperties()
    {
        _sut.SetCustomProperty(_path, "MatterNumber", "M-2026-0091");

        var properties = _sut.GetCustomProperties(_path);

        Assert.Equal("M-2026-0091", properties["MatterNumber"]);
    }

    [Fact]
    public void SetCustomProperty_supports_bool_int_double_and_date()
    {
        _sut.SetCustomProperty(_path, "IsApproved", true);
        _sut.SetCustomProperty(_path, "RevisionCount", 7);
        _sut.SetCustomProperty(_path, "ContractValue", 12500.50);
        _sut.SetCustomProperty(_path, "SignedOn", new DateTime(2026, 8, 22, 0, 0, 0, DateTimeKind.Utc));

        var properties = _sut.GetCustomProperties(_path);

        Assert.Equal("true", properties["IsApproved"]);
        Assert.Equal("7", properties["RevisionCount"]);
        Assert.Equal(12500.50, double.Parse(properties["ContractValue"]));
        Assert.Contains("2026-08-22", properties["SignedOn"]);
    }

    [Fact]
    public void SetCustomProperty_called_twice_for_the_same_name_replaces_not_duplicates()
    {
        _sut.SetCustomProperty(_path, "Status", "Draft");
        _sut.SetCustomProperty(_path, "Status", "Final");

        var properties = _sut.GetCustomProperties(_path);

        Assert.Equal("Final", properties["Status"]);
    }

    [Fact]
    public void RemoveCustomProperty_removes_an_existing_property_and_reports_true()
    {
        _sut.SetCustomProperty(_path, "Temp", "value");

        var removed = _sut.RemoveCustomProperty(_path, "Temp");

        Assert.True(removed);
        Assert.DoesNotContain("Temp", _sut.GetCustomProperties(_path).Keys);
    }

    [Fact]
    public void RemoveCustomProperty_on_a_missing_property_reports_false()
    {
        Assert.False(_sut.RemoveCustomProperty(_path, "DoesNotExist"));
    }

    [Fact]
    public void SetCoreProperties_sets_title_author_subject_and_keywords()
    {
        _sut.SetCoreProperties(_path, title: "Master Services Agreement", author: "Acme Legal", subject: "Contract", keywords: "contract, legal, msa");

        var props = ReadCoreProperties(_path);
        Assert.Equal("Master Services Agreement", props.Title);
        Assert.Equal("Acme Legal", props.Creator);
        Assert.Equal("Contract", props.Subject);
        Assert.Equal("contract, legal, msa", props.Keywords);
    }

    private static (string? Title, string? Creator, string? Subject, string? Keywords) ReadCoreProperties(string path)
    {
        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(path, isEditable: false);
#pragma warning disable OOXML0001
        var properties = doc.PackageProperties;
        return (properties.Title, properties.Creator, properties.Subject, properties.Keywords);
#pragma warning restore OOXML0001
    }

    public void Dispose() => File.Delete(_path);
}
