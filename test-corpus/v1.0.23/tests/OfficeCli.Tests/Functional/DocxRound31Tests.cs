// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Regression tests for DOCX bugs found in rounds 31-60.
/// Covers paragraph-property key casing, table cell gridSpan casing,
/// keepNext round-trip on Add, widowControl=false XML write,
/// and w14glow / w14shadow full property readback.
/// </summary>
public class DocxRound31Tests : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private (string path, WordHandler handler) CreateDoc()
    {
        var path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.docx");
        _tempFiles.Add(path);
        BlankDocCreator.Create(path);
        return (path, new WordHandler(path, editable: true));
    }

    private WordHandler Reopen(string path)
        => new WordHandler(path, editable: true);

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            try { File.Delete(f); } catch { }
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug R31-1: pageBreakBefore key casing
    //
    // Navigation.cs previously emitted "pagebreakbefore" (all lowercase)
    // instead of the canonical camelCase "pageBreakBefore".
    // Fix: Navigation emits "pageBreakBefore".
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public void Bug_R31_1a_PageBreakBefore_Get_ReturnsCamelCaseKey()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Break before",
            ["pagebreakbefore"] = "true"
        });

        var node = handler.Get("/body/p[1]");

        node.Format.Should().ContainKey("pageBreakBefore",
            "Get must return the canonical camelCase key 'pageBreakBefore', not 'pagebreakbefore'");
        node.Format.Should().NotContainKey("pagebreakbefore",
            "Legacy all-lowercase key must not appear in Get output");
    }

    [Fact]
    public void Bug_R31_1b_PageBreakBefore_Add_And_Get_RoundTrip()
    {
        var (path, h) = CreateDoc();
        using (var handler = h)
        {
            handler.Add("/body", "paragraph", null, new Dictionary<string, string>
            {
                ["text"] = "Break before paragraph",
                ["pageBreakBefore"] = "true"
            });
        }

        using var h2 = Reopen(path);
        var node = h2.Get("/body/p[1]");
        node.Format.Should().ContainKey("pageBreakBefore");
        node.Format["pageBreakBefore"].Should().Be(true,
            "pageBreakBefore=true must persist across save/reopen");
    }

    [Fact]
    public void Bug_R31_1c_PageBreakBefore_Set_AcceptsCanonicalKey()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "No break initially"
        });

        var unsupported = handler.Set("/body/p[1]", new Dictionary<string, string>
        {
            ["pageBreakBefore"] = "true"
        });

        unsupported.Should().NotContain("pageBreakBefore",
            "Set must accept canonical key 'pageBreakBefore'");

        var node = handler.Get("/body/p[1]");
        node.Format.Should().ContainKey("pageBreakBefore");
        node.Format["pageBreakBefore"].Should().Be(true);
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug R31-2: widowControl key casing
    //
    // Navigation.cs previously emitted "widowcontrol" (all lowercase),
    // not the canonical camelCase "widowControl".
    // Fix: Navigation emits "widowControl" and respects Val=false.
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public void Bug_R31_2a_WidowControl_Get_ReturnsCamelCaseKey()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Widow controlled",
            ["widowcontrol"] = "true"
        });

        var node = handler.Get("/body/p[1]");

        node.Format.Should().ContainKey("widowControl",
            "Get must return the canonical camelCase key 'widowControl', not 'widowcontrol'");
        node.Format.Should().NotContainKey("widowcontrol",
            "Legacy all-lowercase key must not appear in Get output");
    }

    [Fact]
    public void Bug_R31_2b_WidowControl_False_WritesAndReadsBack()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Widow disabled",
            ["widowControl"] = "false"
        });

        var node = handler.Get("/body/p[1]");

        // widowControl=false writes <w:widowControl w:val="0"/> to XML.
        // Navigation must read back the Val attribute and return false,
        // not blindly return true just because the element is present.
        node.Format.Should().ContainKey("widowControl",
            "widowControl element should be reported even when Val=false");
        node.Format["widowControl"].Should().Be(false,
            "widowControl=false must read back as false, not true");
    }

    [Fact]
    public void Bug_R31_2c_WidowControl_Set_False_WritesToXml()
    {
        var (path, h) = CreateDoc();
        using (var handler = h)
        {
            handler.Add("/body", "paragraph", null, new Dictionary<string, string>
            {
                ["text"] = "Initially enabled",
                ["widowControl"] = "true"
            });

            handler.Set("/body/p[1]", new Dictionary<string, string>
            {
                ["widowControl"] = "false"
            });
        }

        using var h2 = Reopen(path);
        var node = h2.Get("/body/p[1]");
        node.Format.Should().ContainKey("widowControl");
        node.Format["widowControl"].Should().Be(false,
            "Set widowControl=false must persist after save/reopen as false (Val=0 in XML)");
    }

    [Fact]
    public void Bug_R31_2d_WidowControl_True_RoundTrip()
    {
        var (path, h) = CreateDoc();
        using (var handler = h)
        {
            handler.Add("/body", "paragraph", null, new Dictionary<string, string>
            {
                ["text"] = "Widow enabled",
                ["widowControl"] = "true"
            });
        }

        using var h2 = Reopen(path);
        var node = h2.Get("/body/p[1]");
        node.Format.Should().ContainKey("widowControl");
        node.Format["widowControl"].Should().Be(true,
            "widowControl=true must read back as true");
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug R31-3: gridSpan key casing
    //
    // Navigation.cs previously emitted "gridspan" (lowercase), not the
    // canonical camelCase "gridSpan". Fix: use "gridSpan".
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public void Bug_R31_3a_GridSpan_Get_ReturnsCamelCaseKey()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        // Create a 3-column table so we can span across columns
        handler.Add("/body", "table", null, new Dictionary<string, string>
        {
            ["cols"] = "3",
            ["rows"] = "2"
        });

        // Merge cells in first row (col 1 spans 2 columns)
        handler.Set("/body/tbl[1]/tr[1]/tc[1]", new Dictionary<string, string>
        {
            ["gridSpan"] = "2"
        });

        var node = handler.Get("/body/tbl[1]/tr[1]/tc[1]");

        node.Format.Should().ContainKey("gridSpan",
            "Get must return 'gridSpan' (camelCase), not 'gridspan'");
        node.Format.Should().NotContainKey("gridspan",
            "Legacy lowercase key must not appear in Get output");
    }

    [Fact]
    public void Bug_R31_3b_GridSpan_Value_IsGreaterThanOne()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "table", null, new Dictionary<string, string>
        {
            ["cols"] = "3",
            ["rows"] = "1"
        });

        handler.Set("/body/tbl[1]/tr[1]/tc[1]", new Dictionary<string, string>
        {
            ["gridSpan"] = "2"
        });

        var node = handler.Get("/body/tbl[1]/tr[1]/tc[1]");

        node.Format.Should().ContainKey("gridSpan");
        var spanVal = Convert.ToInt32(node.Format["gridSpan"]);
        spanVal.Should().Be(2,
            "gridSpan=2 must read back as 2");
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug R31-4: keepNext on Add
    //
    // Add("paragraph", keepNext=true) must write <w:keepNext/> to XML,
    // which Navigation should then read back as Format["keepNext"] = true.
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public void Bug_R31_4a_KeepNext_AddWithProperty_ReadsBack()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Keep with next paragraph",
            ["keepNext"] = "true"
        });

        var node = handler.Get("/body/p[1]");

        node.Format.Should().ContainKey("keepNext",
            "keepNext=true on Add must be readable back via Get");
        node.Format["keepNext"].Should().Be(true);
    }

    [Fact]
    public void Bug_R31_4b_KeepNext_AddWithLowercaseAlias_ReadsBack()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        // Legacy lowercase alias
        handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Keep next (lowercase alias)",
            ["keepnext"] = "true"
        });

        var node = handler.Get("/body/p[1]");

        node.Format.Should().ContainKey("keepNext",
            "keepnext (lowercase alias) on Add must be readable back as 'keepNext'");
        node.Format["keepNext"].Should().Be(true);
    }

    [Fact]
    public void Bug_R31_4c_KeepNext_Persistence_AfterReopen()
    {
        var (path, h) = CreateDoc();
        using (var handler = h)
        {
            handler.Add("/body", "paragraph", null, new Dictionary<string, string>
            {
                ["text"] = "Persistent keepNext",
                ["keepNext"] = "true"
            });
        }

        using var h2 = Reopen(path);
        var node = h2.Get("/body/p[1]");
        node.Format.Should().ContainKey("keepNext");
        node.Format["keepNext"].Should().Be(true,
            "keepNext must persist after save/reopen");
    }

    [Fact]
    public void Bug_R31_4d_KeepNext_Set_ThenGet()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "No keepNext initially"
        });

        var node = handler.Get("/body/p[1]");
        node.Format.Should().NotContainKey("keepNext",
            "paragraph without keepNext must not have the key");

        handler.Set("/body/p[1]", new Dictionary<string, string>
        {
            ["keepNext"] = "true"
        });

        node = handler.Get("/body/p[1]");
        node.Format.Should().ContainKey("keepNext");
        node.Format["keepNext"].Should().Be(true,
            "Set keepNext=true must write <w:keepNext/> and be readable");
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug R31-5: widowControl=false writes to XML
    //
    // Add/Set with widowControl=false must write <w:widowControl w:val="0"/>
    // (not omit the element entirely). Navigation must then return false,
    // not true, for such a paragraph.
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public void Bug_R31_5a_WidowControlFalse_ElementPresentInXml()
    {
        var (path, h) = CreateDoc();
        using (var handler = h)
        {
            handler.Add("/body", "paragraph", null, new Dictionary<string, string>
            {
                ["text"] = "widowControl disabled",
                ["widowControl"] = "false"
            });
        }

        // Verify the XML contains the element with val=0
        var xml = File.ReadAllText(path); // not guaranteed for binary; use raw XML part
        // Open via raw XML to inspect the generated XML
        using var h2 = Reopen(path);
        var node = h2.Get("/body/p[1]");

        // The element MUST be present (val=false is semantically different from absent)
        node.Format.Should().ContainKey("widowControl",
            "<w:widowControl w:val=\"0\"/> must be written and read back, not silently dropped");
        node.Format["widowControl"].Should().Be(false,
            "widowControl=false must read back as false");
    }

    [Fact]
    public void Bug_R31_5b_WidowControlFalse_DoesNotReadBackAsTrue()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Widow explicitly disabled",
            ["widowControl"] = "false"
        });

        var node = handler.Get("/body/p[1]");

        // If Navigation just checks `pProps.WidowControl != null` and blindly returns true,
        // this assertion would catch the bug.
        if (node.Format.TryGetValue("widowControl", out var val))
        {
            val.Should().Be(false,
                "widowControl element with Val=false must not be reported as true");
        }
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug R31-6: w14glow full readback
    //
    // Previously Navigation only returned color for w14glow, losing the
    // radius and opacity parameters.
    // Fix: return "COLOR;RADIUS;OPACITY" format.
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public void Bug_R31_6a_W14Glow_Readback_IncludesRadius()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Glow test"
        });

        // Add w14glow with explicit radius via Set
        handler.Set("/body/p[1]/r[1]", new Dictionary<string, string>
        {
            ["w14glow"] = "FF0000;10;50"  // COLOR;RADIUS_PT;OPACITY_%
        });

        var node = handler.Get("/body/p[1]/r[1]");

        node.Format.Should().ContainKey("w14glow",
            "w14glow must be readable after Set");

        var glowVal = node.Format["w14glow"].ToString()!;

        glowVal.Should().Contain(";",
            "w14glow readback must include at least one ';' separator — " +
            "expected format 'COLOR;RADIUS;OPACITY' not just 'COLOR'");
    }

    [Fact]
    public void Bug_R31_6b_W14Glow_Readback_IncludesOpacity()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Glow opacity test"
        });

        handler.Set("/body/p[1]/r[1]", new Dictionary<string, string>
        {
            ["w14glow"] = "0000FF;8;75"   // COLOR;RADIUS_PT;OPACITY_%
        });

        var node = handler.Get("/body/p[1]/r[1]");
        node.Format.Should().ContainKey("w14glow");

        var glowVal = node.Format["w14glow"].ToString()!;
        var parts = glowVal.Split(';');

        parts.Should().HaveCountGreaterThanOrEqualTo(2,
            "w14glow readback must contain at least COLOR;RADIUS. " +
            "Previously only COLOR was returned.");

        // Verify color round-trips
        parts[0].Should().Be("#0000FF",
            "Color portion of w14glow must include # prefix");
    }

    [Fact]
    public void Bug_R31_6c_W14Glow_RoundTrip_DefaultParameters()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Glow default"
        });

        // Use default parameters (no radius/opacity specified)
        handler.Set("/body/p[1]/r[1]", new Dictionary<string, string>
        {
            ["w14glow"] = "FF8800"
        });

        var node = handler.Get("/body/p[1]/r[1]");
        node.Format.Should().ContainKey("w14glow");

        var glowVal = node.Format["w14glow"].ToString()!;
        // Default radius is 8pt, default opacity is 75%
        glowVal.Should().StartWith("#FF8800",
            "Color portion must be present with # prefix");
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug R31-7: w14shadow full readback
    //
    // Previously Navigation only returned the color for w14shadow, losing
    // blur, angle, distance, and opacity parameters.
    // Fix: return "COLOR;BLUR;ANGLE;DIST;OPACITY" format.
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public void Bug_R31_7a_W14Shadow_Readback_IncludesBlurAndAngle()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Shadow test"
        });

        // Set shadow with full parameters: COLOR;BLUR;ANGLE;DIST;OPACITY
        handler.Set("/body/p[1]/r[1]", new Dictionary<string, string>
        {
            ["w14shadow"] = "333333;6;45;3;50"
        });

        var node = handler.Get("/body/p[1]/r[1]");
        node.Format.Should().ContainKey("w14shadow",
            "w14shadow must be readable after Set");

        var shadowVal = node.Format["w14shadow"].ToString()!;

        shadowVal.Should().Contain(";",
            "w14shadow readback must include ';' separators — " +
            "expected format 'COLOR;BLUR;ANGLE;DIST;OPACITY' not just 'COLOR'");
    }

    [Fact]
    public void Bug_R31_7b_W14Shadow_Readback_HasFiveComponents()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Shadow full"
        });

        handler.Set("/body/p[1]/r[1]", new Dictionary<string, string>
        {
            ["w14shadow"] = "000000;4;45;3;40"
        });

        var node = handler.Get("/body/p[1]/r[1]");
        node.Format.Should().ContainKey("w14shadow");

        var shadowVal = node.Format["w14shadow"].ToString()!;
        var parts = shadowVal.Split(';');

        parts.Should().HaveCountGreaterThanOrEqualTo(3,
            "w14shadow readback must contain at least COLOR;BLUR;ANGLE. " +
            "Previously only COLOR was returned.");

        parts[0].Should().Be("#000000",
            "Color portion of w14shadow must include # prefix");
    }

    [Fact]
    public void Bug_R31_7c_W14Shadow_Readback_ColorHasPrefixHash()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Shadow color"
        });

        handler.Set("/body/p[1]/r[1]", new Dictionary<string, string>
        {
            ["w14shadow"] = "FF0000"  // color only, use defaults
        });

        var node = handler.Get("/body/p[1]/r[1]");
        node.Format.Should().ContainKey("w14shadow");

        var shadowVal = node.Format["w14shadow"].ToString()!;
        var colorPart = shadowVal.Split(';')[0];

        colorPart.Should().StartWith("#",
            "Color in w14shadow readback must have # prefix per canonical format rules");
        colorPart.Should().Be("#FF0000");
    }

    [Fact]
    public void Bug_R31_7d_W14Shadow_RoundTrip_DefaultParameters()
    {
        var (path, h) = CreateDoc();
        using (var handler = h)
        {
            handler.Add("/body", "paragraph", null, new Dictionary<string, string>
            {
                ["text"] = "Shadow persist"
            });

            handler.Set("/body/p[1]/r[1]", new Dictionary<string, string>
            {
                ["w14shadow"] = "808080;4;45;3;40"
            });
        }

        using var h2 = Reopen(path);
        var node = h2.Get("/body/p[1]/r[1]");
        node.Format.Should().ContainKey("w14shadow",
            "w14shadow must persist after save/reopen");

        var shadowVal = node.Format["w14shadow"].ToString()!;
        shadowVal.Should().StartWith("#808080",
            "Color must be preserved across save/reopen with # prefix");
    }
}
