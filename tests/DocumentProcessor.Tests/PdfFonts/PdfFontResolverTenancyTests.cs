using DocumentProcessor.Core.PdfFonts;

namespace DocumentProcessor.Tests.PdfFonts;

/// <summary>
/// The resolver is necessarily a process-wide singleton, because PDFsharp resolves fonts through
/// one global <c>GlobalFontSettings.FontResolver</c> with no per-document hook. These tests pin the
/// isolation that makes that safe in a shared multi-tenant host.
/// </summary>
public class PdfFontResolverTenancyTests
{
    // Distinct byte payloads stand in for two different real font files; the resolver stores and
    // returns bytes verbatim, so identity is all that needs checking here.
    private static byte[] FontA => [1, 1, 1, 1];
    private static byte[] FontB => [2, 2, 2, 2];

    private static string NewTenantId() => $"tenant-{Guid.NewGuid():N}";

    [Fact]
    public void Two_tenants_registering_the_same_family_name_do_not_overwrite_each_other()
    {
        var tenantA = NewTenantId();
        var tenantB = NewTenantId();

        using (PdfFontResolver.BeginTenantScope(tenantA))
            PdfFontResolver.Instance.RegisterFont("BrandSans", FontA);

        using (PdfFontResolver.BeginTenantScope(tenantB))
            PdfFontResolver.Instance.RegisterFont("BrandSans", FontB);

        using (PdfFontResolver.BeginTenantScope(tenantA))
            Assert.Equal(FontA, PdfFontResolver.Instance.GetFontBytes("BrandSans"));

        using (PdfFontResolver.BeginTenantScope(tenantB))
            Assert.Equal(FontB, PdfFontResolver.Instance.GetFontBytes("BrandSans"));

        PdfFontResolver.Instance.ClearTenant(tenantA);
        PdfFontResolver.Instance.ClearTenant(tenantB);
    }

    [Fact]
    public void The_face_name_handed_to_PdfSharp_is_tenant_qualified()
    {
        // PDFsharp caches glyph typefaces by the face name this returns, so two tenants must not
        // share one — otherwise the dictionary is partitioned but PDFsharp's own cache is not.
        var tenantA = NewTenantId();
        var tenantB = NewTenantId();

        using (PdfFontResolver.BeginTenantScope(tenantA))
            PdfFontResolver.Instance.RegisterFont("BrandSans", FontA);
        using (PdfFontResolver.BeginTenantScope(tenantB))
            PdfFontResolver.Instance.RegisterFont("BrandSans", FontB);

        string faceA, faceB;
        using (PdfFontResolver.BeginTenantScope(tenantA))
            faceA = PdfFontResolver.Instance.ResolveTypeface("BrandSans", false, false).FaceName;
        using (PdfFontResolver.BeginTenantScope(tenantB))
            faceB = PdfFontResolver.Instance.ResolveTypeface("BrandSans", false, false).FaceName;

        Assert.NotEqual(faceA, faceB);
        // And each face name round-trips back to that tenant's own bytes via PDFsharp's callback.
        Assert.Equal(FontA, PdfFontResolver.Instance.GetFont(faceA));
        Assert.Equal(FontB, PdfFontResolver.Instance.GetFont(faceB));

        PdfFontResolver.Instance.ClearTenant(tenantA);
        PdfFontResolver.Instance.ClearTenant(tenantB);
    }

    [Fact]
    public void A_tenants_font_is_invisible_outside_its_scope()
    {
        var tenantA = NewTenantId();
        using (PdfFontResolver.BeginTenantScope(tenantA))
            PdfFontResolver.Instance.RegisterFont("ScopedOnly", FontA);

        // Unscoped resolution must not see it — it falls back to the bundled default instead.
        Assert.NotEqual(FontA, PdfFontResolver.Instance.GetFontBytes("ScopedOnly"));

        PdfFontResolver.Instance.ClearTenant(tenantA);
    }

    [Fact]
    public void ClearTenant_releases_that_tenants_fonts_and_leaves_others_alone()
    {
        var tenantA = NewTenantId();
        var tenantB = NewTenantId();

        using (PdfFontResolver.BeginTenantScope(tenantA))
            PdfFontResolver.Instance.RegisterFont("Shared", FontA);
        using (PdfFontResolver.BeginTenantScope(tenantB))
            PdfFontResolver.Instance.RegisterFont("Shared", FontB);

        var removed = PdfFontResolver.Instance.ClearTenant(tenantA);

        Assert.Equal(1, removed);
        using (PdfFontResolver.BeginTenantScope(tenantA))
            Assert.NotEqual(FontA, PdfFontResolver.Instance.GetFontBytes("Shared"));
        using (PdfFontResolver.BeginTenantScope(tenantB))
            Assert.Equal(FontB, PdfFontResolver.Instance.GetFontBytes("Shared"));

        PdfFontResolver.Instance.ClearTenant(tenantB);
    }

    [Fact]
    public void Scopes_nest_and_restore_the_previous_tenant()
    {
        var outer = NewTenantId();
        var inner = NewTenantId();

        using (PdfFontResolver.BeginTenantScope(outer))
        {
            PdfFontResolver.Instance.RegisterFont("Nested", FontA);

            using (PdfFontResolver.BeginTenantScope(inner))
                PdfFontResolver.Instance.RegisterFont("Nested", FontB);

            // Back in the outer scope, the outer tenant's font is what resolves.
            Assert.Equal(FontA, PdfFontResolver.Instance.GetFontBytes("Nested"));
        }

        PdfFontResolver.Instance.ClearTenant(outer);
        PdfFontResolver.Instance.ClearTenant(inner);
    }

    [Fact]
    public async Task Concurrent_tenants_do_not_bleed_across_async_boundaries()
    {
        // The scope rides AsyncLocal, so it must survive awaits and stay independent per task —
        // this is the actual request-concurrency shape, not the sequential case above.
        var tenants = Enumerable.Range(0, 16).Select(i => ($"tenant-{Guid.NewGuid():N}", new byte[] { (byte)i, 9, 9 })).ToList();

        await Task.WhenAll(tenants.Select(t => Task.Run(async () =>
        {
            var (id, bytes) = t;
            using (PdfFontResolver.BeginTenantScope(id))
            {
                PdfFontResolver.Instance.RegisterFont("Concurrent", bytes);
                await Task.Yield();
                await Task.Delay(5);
                Assert.Equal(bytes, PdfFontResolver.Instance.GetFontBytes("Concurrent"));
            }
        })));

        foreach (var (id, _) in tenants)
            PdfFontResolver.Instance.ClearTenant(id);
    }
}
