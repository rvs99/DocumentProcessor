using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentProcessor.Core.Redlining;

/// <summary>What editing a protected document still allows — mirrors OOXML's <c>ST_DocProtect</c>.</summary>
public enum EditRestriction
{
    /// <summary>No changes at all.</summary>
    ReadOnly,
    /// <summary>Only comments may be added.</summary>
    Comments,
    /// <summary>Only tracked changes may be made — matches Word's "Track Changes" restriction mode.</summary>
    TrackedChanges,
    /// <summary>Only form-field content may be filled in.</summary>
    Forms
}

/// <summary>Which editors an <see cref="DocumentProtectionService.AllowEditingInRange"/> exception applies to.</summary>
public enum EditorGroup { Everyone, Administrators, Contributors, Editors, Owners, Current }

/// <summary>
/// Word's "Restrict Editing" feature: document-wide protection (optionally password-backed) plus
/// range exceptions that stay editable despite it — e.g. lock an entire contract to Tracked Changes
/// only, while leaving a specific signature-block paragraph range freely editable for both parties.
/// </summary>
public sealed class DocumentProtectionService : IDocumentProtectionService
{
    /// <summary>
    /// Restricts editing for the whole document to <paramref name="restriction"/>. When
    /// <paramref name="password"/> is supplied, protection is password-backed using the same
    /// iterated-SHA-1 hash (100,000 rounds over a random salt) Word itself writes for "Restrict
    /// Editing" — removing protection in Word then requires that password. Without a password,
    /// protection is advisory only (Word's UI will still offer to turn it off with one click), which
    /// is still useful as a signal to Word's own editing UI/behavior even when enforcement isn't
    /// cryptographically backed.
    /// </summary>
    public void SetDocumentProtection(string docxPath, EditRestriction restriction, string? password = null)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var mainPart = doc.MainDocumentPart ?? throw new CorruptDocumentException("Document has no main part.");
        var settingsPart = mainPart.DocumentSettingsPart ?? mainPart.AddNewPart<DocumentSettingsPart>();
        settingsPart.Settings ??= new Settings();

        settingsPart.Settings.RemoveAllChildren<DocumentProtection>();

        var protection = new DocumentProtection
        {
            Edit = ToProtectionValue(restriction),
            Enforcement = true
        };

        if (password is not null)
        {
            const int spinCount = 100_000;
            var (hash, salt) = ComputePasswordHash(password, spinCount);
            protection.CryptographicProviderType = CryptProviderValues.RsaAdvancedEncryptionStandard;
            protection.CryptographicAlgorithmClass = CryptAlgorithmClassValues.Hash;
            protection.CryptographicAlgorithmType = CryptAlgorithmValues.TypeAny;
            protection.CryptographicAlgorithmSid = 4; // SHA-1, per ECMA-376's ST_CryptAlgSid enumeration
            protection.CryptographicSpinCount = (uint)spinCount;
            protection.Hash = Convert.ToBase64String(hash);
            protection.Salt = Convert.ToBase64String(salt);
        }

        settingsPart.Settings.PrependChild(protection);
        settingsPart.Settings.Save();
    }

    /// <summary>Removes document-level protection entirely (no password required by this API — that
    /// check is Word's own UI concern, not a guarantee this library enforces).</summary>
    public void RemoveDocumentProtection(string docxPath)
    {
        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var settings = doc.MainDocumentPart?.DocumentSettingsPart?.Settings;
        if (settings is null)
            return;

        settings.RemoveAllChildren<DocumentProtection>();
        settings.Save();
    }

    /// <summary>
    /// Marks paragraphs [<paramref name="startParagraphIndex"/>, <paramref name="endParagraphIndex"/>]
    /// (inclusive, 0-based) as an editing exception via <c>w:permStart</c>/<c>w:permEnd</c> — the same
    /// mechanism Word's "Restrict Editing" uses for "Everyone can edit this section" ranges. Only
    /// meaningful alongside <see cref="SetDocumentProtection"/>: Word treats permission ranges as
    /// exceptions to document-wide protection, not as protection by themselves.
    /// </summary>
    public void AllowEditingInRange(string docxPath, int startParagraphIndex, int endParagraphIndex, EditorGroup editorGroup = EditorGroup.Everyone)
    {
        if (endParagraphIndex < startParagraphIndex)
            throw new ArgumentOutOfRangeException(nameof(endParagraphIndex), "End index must be >= start index.");

        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var body = doc.MainDocumentPart?.Document?.Body ?? throw new CorruptDocumentException("Document has no main part/body.");
        var paragraphs = body.Elements<Paragraph>().ToList();

        if (endParagraphIndex >= paragraphs.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(endParagraphIndex),
                $"Document has {paragraphs.Count} paragraphs; valid indices are 0..{paragraphs.Count - 1}.");
        }

        var id = Guid.NewGuid().GetHashCode() & 0x7FFFFFFF; // permStart/permEnd ids just need to match each other and be non-negative
        var startParagraph = paragraphs[startParagraphIndex];
        var endParagraph = paragraphs[endParagraphIndex];

        startParagraph.InsertAt(new PermStart { Id = id, EditorGroup = ToEditorGroupValue(editorGroup) }, 0);
        endParagraph.AppendChild(new PermEnd { Id = id });

        doc.MainDocumentPart!.Document!.Save();
    }

    /// <summary>
    /// Computes the same password verifier Word writes for <c>w:documentProtection</c>: SHA-1 of
    /// (salt + UTF-16LE password), then re-hashed <paramref name="spinCount"/> times as
    /// SHA-1(4-byte little-endian iteration counter + previous hash) — per ECMA-376's password
    /// hashing algorithm for legacy document/VBA project protection, reused here for "Restrict
    /// Editing" passwords. Exposed publicly so callers can independently verify a password without
    /// needing Word.
    /// </summary>
    public static (byte[] Hash, byte[] Salt) ComputePasswordHash(string password, int spinCount)
    {
        var salt = RandomNumberGenerator.GetBytes(16);
        var hash = ComputePasswordHash(password, salt, spinCount);
        return (hash, salt);
    }

    /// <summary>Overload for verifying an existing password against a known salt (e.g. to check a
    /// candidate password against a document's stored hash/salt).</summary>
    public static byte[] ComputePasswordHash(string password, byte[] salt, int spinCount)
    {
        var passwordBytes = Encoding.Unicode.GetBytes(password);
        var combined = new byte[salt.Length + passwordBytes.Length];
        Buffer.BlockCopy(salt, 0, combined, 0, salt.Length);
        Buffer.BlockCopy(passwordBytes, 0, combined, salt.Length, passwordBytes.Length);

        var hash = SHA1.HashData(combined);
        for (var i = 0; i < spinCount; i++)
        {
            var iterationBytes = BitConverter.GetBytes(i);
            var input = new byte[iterationBytes.Length + hash.Length];
            Buffer.BlockCopy(iterationBytes, 0, input, 0, iterationBytes.Length);
            Buffer.BlockCopy(hash, 0, input, iterationBytes.Length, hash.Length);
            hash = SHA1.HashData(input);
        }

        return hash;
    }

    private static DocumentProtectionValues ToProtectionValue(EditRestriction restriction) => restriction switch
    {
        EditRestriction.ReadOnly => DocumentProtectionValues.ReadOnly,
        EditRestriction.Comments => DocumentProtectionValues.Comments,
        EditRestriction.TrackedChanges => DocumentProtectionValues.TrackedChanges,
        EditRestriction.Forms => DocumentProtectionValues.Forms,
        _ => throw new ArgumentOutOfRangeException(nameof(restriction))
    };

    private static RangePermissionEditingGroupValues ToEditorGroupValue(EditorGroup group) => group switch
    {
        EditorGroup.Everyone => RangePermissionEditingGroupValues.Everyone,
        EditorGroup.Administrators => RangePermissionEditingGroupValues.Administrators,
        EditorGroup.Contributors => RangePermissionEditingGroupValues.Contributors,
        EditorGroup.Editors => RangePermissionEditingGroupValues.Editors,
        EditorGroup.Owners => RangePermissionEditingGroupValues.Owners,
        EditorGroup.Current => RangePermissionEditingGroupValues.Current,
        _ => throw new ArgumentOutOfRangeException(nameof(group))
    };
}
