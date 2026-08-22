using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.Redlining;
using DocumentProcessor.Core.Samples;

namespace DocumentProcessor.Tests.Redlining;

public class DocumentProtectionServiceTests : IDisposable
{
    private readonly string _path = TestFiles.NewTempPath(".docx");
    private readonly DocumentProtectionService _sut = new();

    public DocumentProtectionServiceTests()
    {
        SampleDocumentFactory.CreateBasicDocument(_path, "Agreement",
            ["Paragraph one.", "Paragraph two.", "Paragraph three."]);
    }

    [Fact]
    public void SetDocumentProtection_without_password_sets_edit_restriction_and_no_hash()
    {
        _sut.SetDocumentProtection(_path, EditRestriction.TrackedChanges);

        var protection = ReadProtection(_path);
        Assert.NotNull(protection);
        Assert.Equal(DocumentProtectionValues.TrackedChanges, protection!.Edit!.Value);
        Assert.True(protection.Enforcement!.Value);
        Assert.Null(protection.Hash?.Value);
    }

    [Fact]
    public void SetDocumentProtection_with_password_writes_a_hash_matching_the_documented_algorithm()
    {
        _sut.SetDocumentProtection(_path, EditRestriction.ReadOnly, password: "correct horse battery staple");

        var protection = ReadProtection(_path);
        Assert.NotNull(protection?.Hash?.Value);
        Assert.NotNull(protection?.Salt?.Value);
        Assert.Equal(100_000, (int)protection!.CryptographicSpinCount!.Value);

        var salt = Convert.FromBase64String(protection.Salt!.Value!);
        var expectedHash = DocumentProtectionService.ComputePasswordHash("correct horse battery staple", salt, 100_000);
        Assert.Equal(Convert.ToBase64String(expectedHash), protection.Hash!.Value);
    }

    [Fact]
    public void RemoveDocumentProtection_clears_a_previously_set_protection()
    {
        _sut.SetDocumentProtection(_path, EditRestriction.Forms, password: "secret");

        _sut.RemoveDocumentProtection(_path);

        Assert.Null(ReadProtection(_path));
    }

    [Fact]
    public void AllowEditingInRange_wraps_the_given_paragraphs_with_matching_permStart_and_permEnd()
    {
        _sut.AllowEditingInRange(_path, startParagraphIndex: 1, endParagraphIndex: 2, EditorGroup.Everyone);

        using var doc = WordprocessingDocument.Open(_path, isEditable: false);
        var paragraphs = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().ToList();

        var permStart = paragraphs[1].GetFirstChild<PermStart>();
        var permEnd = paragraphs[2].GetFirstChild<PermEnd>();
        Assert.NotNull(permStart);
        Assert.NotNull(permEnd);
        Assert.Equal(permStart!.Id!.Value, permEnd!.Id!.Value);
        Assert.Equal(RangePermissionEditingGroupValues.Everyone, permStart.EditorGroup!.Value);

        // Paragraphs outside the range are untouched.
        Assert.Null(paragraphs[0].GetFirstChild<PermStart>());
    }

    [Fact]
    public void ComputePasswordHash_is_deterministic_for_the_same_salt_and_differs_for_a_different_password()
    {
        var salt = Convert.FromHexString("00112233445566778899AABBCCDDEEFF");

        var hash1 = DocumentProtectionService.ComputePasswordHash("password-a", salt, 1000);
        var hash2 = DocumentProtectionService.ComputePasswordHash("password-a", salt, 1000);
        var hash3 = DocumentProtectionService.ComputePasswordHash("password-b", salt, 1000);

        Assert.Equal(hash1, hash2);
        Assert.NotEqual(hash1, hash3);
    }

    private static DocumentProtection? ReadProtection(string path)
    {
        using var doc = WordprocessingDocument.Open(path, isEditable: false);
        return doc.MainDocumentPart?.DocumentSettingsPart?.Settings?.GetFirstChild<DocumentProtection>();
    }

    public void Dispose() => File.Delete(_path);
}
