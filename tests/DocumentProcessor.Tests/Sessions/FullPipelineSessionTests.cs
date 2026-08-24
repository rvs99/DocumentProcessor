using System.Diagnostics;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Core.ContentControls;
using DocumentProcessor.Core.FontEmbedding;
using DocumentProcessor.Core.Layout;
using DocumentProcessor.Core.Metadata;
using DocumentProcessor.Core.Redlining;
using DocumentProcessor.Core.Samples;
using DocumentProcessor.Core.Sessions;
using DocumentProcessor.Core.Tables;
using DocumentProcessor.Core.Watermarking;

namespace DocumentProcessor.Tests.Sessions;

/// <summary>
/// The realistic contract-assembly pipeline, run entirely inside one session. Previously each of
/// these steps opened, parsed, serialised and rezipped the whole package independently.
/// </summary>
public class FullPipelineSessionTests : IDisposable
{
    private readonly List<string> _cleanup = [];

    private string NewContract(int bodyParagraphs = 40)
    {
        var path = TestFiles.NewTempPath(".docx");
        _cleanup.Add(path);
        SampleDocumentFactory.CreateDocumentWithContentControls(path, new Dictionary<string, string>
        {
            ["ClientName"] = "[Client]",
            ["EffectiveDate"] = "[Date]",
        });
        SampleDocumentFactory.AppendParagraphs(path,
            Enumerable.Range(0, bodyParagraphs).Select(i => $"Clause {i}: contractual language for section {i}."));
        return path;
    }

    private static TableSpec Pricing() => new(["Item", "Amount"], [["Implementation", "$150,000"]]);

    [Fact]
    public void A_full_pipeline_runs_inside_one_session()
    {
        using var session = DocumentSession.Open(File.ReadAllBytes(NewContract()));

        session.ContentControls.ReplaceMany(new Dictionary<string, string>
        {
            ["ClientName"] = "Acme Corporation",
            ["EffectiveDate"] = "2027-01-01",
        });
        session.Tables.AppendTable(Pricing());
        session.PageLayout.SetMargins(PageMargins.FromInches(1, 1, 1, 1));
        session.Fonts.EmbedAndApply("Roboto Mono", new FontFamilyFiles(
            Path.Combine(AppContext.BaseDirectory, "Assets", "Fonts", "RobotoMono-Regular.ttf")));
        session.Watermark.AddText("DRAFT");
        session.Metadata.SetCustomProperties(new Dictionary<string, object?>
        {
            ["MatterNumber"] = "M-2027-0042",
            ["Value"] = 250000.0,
        });
        session.Fields.SetUpdateOnOpen();
        session.Protection.Restrict(EditRestriction.TrackedChanges, password: "secret");

        var result = session.Save();

        // Every step is observable in the single saved output.
        using var stream = new MemoryStream(result);
        using var doc = WordprocessingDocument.Open(stream, isEditable: false);
        var main = doc.MainDocumentPart!;

        Assert.Contains("Acme Corporation", main.Document!.Body!.InnerText);
        Assert.Single(main.Document.Body.Elements<Table>());
        Assert.NotEmpty(main.HeaderParts);                                   // watermark header
        Assert.NotNull(main.DocumentSettingsPart!.Settings!.GetFirstChild<DocumentProtection>());
        Assert.NotNull(main.DocumentSettingsPart.Settings.GetFirstChild<UpdateFieldsOnOpen>());
        // Embedding writes the font table; applying writes RunFonts into document.xml. Neither
        // touches the styles part, so both are checked where they actually land.
        Assert.NotNull(main.FontTablePart);
        Assert.Contains("Roboto Mono", main.Document.InnerXml);
    }

    [Fact]
    public void The_session_pipeline_matches_the_path_based_pipeline()
    {
        var viaPaths = NewContract();
        var viaSession = NewContract();

        new ContentControlService().ReplaceMany(viaPaths, new Dictionary<string, string> { ["ClientName"] = "Acme" });
        new TableGenerationService().AppendTable(viaPaths, Pricing());
        new PageLayoutService().SetMargins(viaPaths, PageMargins.FromInches(1, 1, 1, 1));
        new DocxWatermarkService().AddTextWatermark(viaPaths, "DRAFT");
        new DocumentMetadataService().SetCustomProperty(viaPaths, "MatterNumber", "M-1");

        byte[] sessionBytes;
        using (var session = DocumentSession.Open(File.ReadAllBytes(viaSession)))
        {
            session.ContentControls.ReplaceMany(new Dictionary<string, string> { ["ClientName"] = "Acme" });
            session.Tables.AppendTable(Pricing());
            session.PageLayout.SetMargins(PageMargins.FromInches(1, 1, 1, 1));
            session.Watermark.AddText("DRAFT");
            session.Metadata.SetCustomProperties(new Dictionary<string, object?> { ["MatterNumber"] = "M-1" });
            sessionBytes = session.Save();
        }

        Assert.Equal(BodyText(File.ReadAllBytes(viaPaths)), BodyText(sessionBytes));
    }

    [Fact]
    public void Tracked_changes_can_be_resolved_within_a_session()
    {
        var path = TestFiles.NewTempPath(".docx");
        _cleanup.Add(path);
        SampleDocumentFactory.CreateDocumentWithTrackedChanges(path, author: "Counsel");

        using var session = DocumentSession.Open(File.ReadAllBytes(path));

        Assert.True(session.TrackChanges.HasTrackedChanges());
        Assert.Equal(2, session.TrackChanges.GetTrackedChanges().Count);

        session.TrackChanges.AcceptAll();
        var accepted = session.Save();

        Assert.Contains("twenty-four", BodyText(accepted));
        Assert.DoesNotContain("twelve", BodyText(accepted));
    }

    [Fact]
    public void An_eight_step_pipeline_beats_the_same_work_through_paths()
    {
        var viaPaths = NewContract(bodyParagraphs: 300);
        var sourceBytes = File.ReadAllBytes(NewContract(bodyParagraphs: 300));

        var pathTimer = Stopwatch.StartNew();
        var metadata = new DocumentMetadataService();
        var layout = new PageLayoutService();
        new ContentControlService().ReplaceMany(viaPaths, new Dictionary<string, string> { ["ClientName"] = "Acme" });
        new TableGenerationService().AppendTable(viaPaths, Pricing());
        layout.SetMargins(viaPaths, PageMargins.FromInches(1, 1, 1, 1));
        layout.SetColumns(viaPaths, 1);
        new DocxWatermarkService().AddTextWatermark(viaPaths, "DRAFT");
        metadata.SetCustomProperty(viaPaths, "A", "1");
        metadata.SetCustomProperty(viaPaths, "B", "2");
        new DocumentProtectionService().SetDocumentProtection(viaPaths, EditRestriction.TrackedChanges);
        pathTimer.Stop();

        var sessionTimer = Stopwatch.StartNew();
        using (var session = DocumentSession.Open(sourceBytes))
        {
            session.ContentControls.ReplaceMany(new Dictionary<string, string> { ["ClientName"] = "Acme" });
            session.Tables.AppendTable(Pricing());
            session.PageLayout.SetMargins(PageMargins.FromInches(1, 1, 1, 1));
            session.PageLayout.SetColumns(1);
            session.Watermark.AddText("DRAFT");
            session.Metadata.SetCustomProperties(new Dictionary<string, object?> { ["A"] = "1", ["B"] = "2" });
            session.Protection.Restrict(EditRestriction.TrackedChanges);
            session.Save();
        }
        sessionTimer.Stop();

        Assert.True(sessionTimer.Elapsed < pathTimer.Elapsed,
            $"session {sessionTimer.ElapsedMilliseconds} ms vs paths {pathTimer.ElapsedMilliseconds} ms for eight operations.");
    }

    private static string BodyText(byte[] docxBytes)
    {
        using var stream = new MemoryStream(docxBytes);
        using var doc = WordprocessingDocument.Open(stream, isEditable: false);
        return doc.MainDocumentPart!.Document!.Body!.InnerText;
    }

    public void Dispose()
    {
        foreach (var path in _cleanup)
            if (File.Exists(path)) File.Delete(path);
    }
}
