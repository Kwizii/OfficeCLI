// Black-box tests (Round 2) targeting:
//   - Arrow/hexagon preset shapes: Add→Get→Verify→Set→Verify (commit 11df1af)
//   - Text inset (margin) on arrow/hexagon shapes
//   - Watermark schema validity after EvenAndOddHeaders/TitlePage fix (commit 0600aa0)
//
// Key API rule: Add("/", "slide", null, new()) with NO title ⇒ no title placeholder ⇒
//   first user shape is shape[1].  Adding ["title"] creates a title placeholder at shape[1].

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;
using Xunit.Abstractions;

namespace OfficeCli.Tests.Functional;

public class BtBlackBoxRound2 : IDisposable
{
    private readonly List<string> _tempFiles = new();
    private readonly ITestOutputHelper _output;

    public BtBlackBoxRound2(ITestOutputHelper output)
    {
        _output = output;
    }

    private string CreateTemp(string ext)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bt2_test_{Guid.NewGuid():N}.{ext}");
        _tempFiles.Add(path);
        BlankDocCreator.Create(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            if (File.Exists(f)) File.Delete(f);
    }

    // ── Helpers ──────────────────────────────────────────────────────────────

    private void AssertValidDocx(string path, string step)
    {
        using var doc = WordprocessingDocument.Open(path, false);
        var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
        var errors = validator.Validate(doc).ToList();
        foreach (var e in errors)
            _output.WriteLine($"[{step}] {e.ErrorType}: {e.Description} @ {e.Path?.XPath}");
        errors.Should().BeEmpty($"DOCX must be schema-valid after step: {step}");
    }

    private void AssertValidPptx(string path, string step)
    {
        using var doc = PresentationDocument.Open(path, false);
        var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
        var errors = validator.Validate(doc).ToList();
        foreach (var e in errors)
            _output.WriteLine($"[{step}] {e.ErrorType}: {e.Description} @ {e.Path?.XPath}");
        errors.Should().BeEmpty($"PPTX must be schema-valid after step: {step}");
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // SECTION 1 — Arrow shapes: preset Add + Get + Set lifecycle
    // ═══════════════════════════════════════════════════════════════════════════

    [Fact]
    public void Pptx_RightArrow_AddGetVerify()
    {
        // No title on slide ⇒ shape[1] is the user shape
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());

        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Go Right",
            ["preset"] = "rightArrow",
            ["fill"] = "4472C4",
            ["width"] = "5cm",
            ["height"] = "2cm"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node.Text.Should().Be("Go Right");
        node.Format.Should().ContainKey("preset");
        node.Format["preset"].ToString().Should().Be("rightArrow");
    }

    [Fact]
    public void Pptx_LeftArrow_AddGetVerify()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());

        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Go Left",
            ["preset"] = "leftArrow",
            ["fill"] = "FF0000"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("Go Left");
        node.Format.Should().ContainKey("preset");
        node.Format["preset"].ToString().Should().Be("leftArrow");
    }

    [Fact]
    public void Pptx_UpArrow_AddGetVerify()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());

        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Go Up",
            ["preset"] = "upArrow",
            ["fill"] = "00B050"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("Go Up");
        node.Format["preset"].ToString().Should().Be("upArrow");
    }

    [Fact]
    public void Pptx_DownArrow_AddGetVerify()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());

        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Go Down",
            ["preset"] = "downArrow",
            ["fill"] = "FFC000"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("Go Down");
        node.Format["preset"].ToString().Should().Be("downArrow");
    }

    [Fact]
    public void Pptx_Hexagon_AddGetVerify()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());

        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Hex Cell",
            ["preset"] = "hexagon",
            ["fill"] = "7030A0"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("Hex Cell");
        node.Format["preset"].ToString().Should().Be("hexagon");
    }

    [Fact]
    public void Pptx_RightArrow_PresetPersistsAfterReopen()
    {
        var path = CreateTemp("pptx");

        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new());
            handler.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = "Right",
                ["preset"] = "rightArrow"
            });
        }

        using (var handler = new PowerPointHandler(path, editable: false))
        {
            var node = handler.Get("/slide[1]/shape[1]");
            node.Format.Should().ContainKey("preset", "preset must survive save/reopen");
            node.Format["preset"].ToString().Should().Be("rightArrow");
            node.Text.Should().Be("Right");
        }
    }

    [Fact]
    public void Pptx_Hexagon_PresetPersistsAfterReopen()
    {
        var path = CreateTemp("pptx");

        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new());
            handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Hex", ["preset"] = "hexagon" });
        }

        using (var handler = new PowerPointHandler(path, editable: false))
        {
            var node = handler.Get("/slide[1]/shape[1]");
            node.Format.Should().ContainKey("preset");
            node.Format["preset"].ToString().Should().Be("hexagon");
            node.Text.Should().Be("Hex");
        }
    }

    [Fact]
    public void Pptx_Arrow_SetPresetChangesGeometry()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Arrow", ["preset"] = "rightArrow" });

        // Verify initial
        var node = handler.Get("/slide[1]/shape[1]");
        node.Format["preset"].ToString().Should().Be("rightArrow");

        // Change to leftArrow
        handler.Set("/slide[1]/shape[1]", new() { ["preset"] = "leftArrow" });

        // Verify updated
        node = handler.Get("/slide[1]/shape[1]");
        node.Format["preset"].ToString().Should().Be("leftArrow", "Set should update preset geometry");
    }

    [Fact]
    public void Pptx_Arrow_SetPresetToHexagon()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Shape", ["preset"] = "rightArrow" });

        handler.Set("/slide[1]/shape[1]", new() { ["preset"] = "hexagon" });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format["preset"].ToString().Should().Be("hexagon");
    }

    [Fact]
    public void Pptx_Hexagon_SetPresetToArrow()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Shape", ["preset"] = "hexagon" });

        handler.Set("/slide[1]/shape[1]", new() { ["preset"] = "rightArrow" });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format["preset"].ToString().Should().Be("rightArrow");
    }

    [Fact]
    public void Pptx_ArrowShape_SetTextUpdatesAndPreservePreset()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Initial",
            ["preset"] = "rightArrow"
        });

        // Verify initial
        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("Initial");
        node.Format["preset"].ToString().Should().Be("rightArrow");

        // Update text only
        handler.Set("/slide[1]/shape[1]", new() { ["text"] = "Updated Arrow" });

        // Verify updated
        node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("Updated Arrow");
        node.Format["preset"].ToString().Should().Be("rightArrow", "preset must remain after text Set");
    }

    [Fact]
    public void Pptx_ArrowShape_FillColorRoundtrip()
    {
        var path = CreateTemp("pptx");

        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new());
            handler.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = "Arrow",
                ["preset"] = "rightArrow",
                ["fill"] = "#4472C4"
            });
        }

        using (var handler = new PowerPointHandler(path, editable: false))
        {
            var node = handler.Get("/slide[1]/shape[1]");
            node.Format.Should().ContainKey("fill");
            node.Format["fill"].ToString().Should().Be("#4472C4");
        }
    }

    [Fact]
    public void Pptx_ArrowShape_BoldTextRoundtrip()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Bold Arrow",
            ["preset"] = "rightArrow",
            ["bold"] = "true",
            ["size"] = "18"
        });

        // Verify Add
        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("Bold Arrow");
        node.Format.Should().ContainKey("bold");
        node.Format["bold"].Should().Be(true);

        // Set bold again (idempotent)
        handler.Set("/slide[1]/shape[1]", new() { ["bold"] = "true" });

        node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("Bold Arrow");
        node.Format["bold"].Should().Be(true);
    }

    [Fact]
    public void Pptx_Arrow_EmptyTextNoException()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());

        var act = () => handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "",
            ["preset"] = "rightArrow"
        });

        act.Should().NotThrow("empty text in arrow shape should be allowed");

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format["preset"].ToString().Should().Be("rightArrow");
    }

    [Fact]
    public void Pptx_Arrow_LongTextRoundtrip()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());

        var longText = new string('A', 200);
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = longText,
            ["preset"] = "rightArrow"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be(longText);
    }

    [Fact]
    public void Pptx_Arrow_SchemaValidAfterAddMultiple()
    {
        var path = CreateTemp("pptx");

        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new());
            handler.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = "Right",
                ["preset"] = "rightArrow",
                ["fill"] = "4472C4"
            });
            handler.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = "Left",
                ["preset"] = "leftArrow",
                ["fill"] = "FF0000"
            });
            handler.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = "Hex",
                ["preset"] = "hexagon",
                ["fill"] = "7030A0"
            });
        }

        AssertValidPptx(path, "Add right/left arrows and hexagon");
    }

    [Fact]
    public void Pptx_FourArrows_AllIndexedCorrectly()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());

        var presets = new[] { "rightArrow", "leftArrow", "upArrow", "downArrow" };
        var texts = new[] { "Right", "Left", "Up", "Down" };

        for (int i = 0; i < presets.Length; i++)
        {
            handler.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = texts[i],
                ["preset"] = presets[i]
            });
        }

        for (int i = 1; i <= 4; i++)
        {
            var node = handler.Get($"/slide[1]/shape[{i}]");
            node.Text.Should().Be(texts[i - 1], $"shape[{i}] text mismatch");
            node.Format["preset"].ToString().Should().Be(presets[i - 1], $"shape[{i}] preset mismatch");
        }
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // SECTION 2 — Text inset (margin) on arrow/hexagon shapes
    // ═══════════════════════════════════════════════════════════════════════════

    [Fact]
    public void Pptx_ArrowShape_MarginRoundtrip()
    {
        var path = CreateTemp("pptx");

        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new());
            handler.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = "Arrow",
                ["preset"] = "rightArrow",
                ["margin"] = "0.2cm"
            });
        }

        using (var handler = new PowerPointHandler(path, editable: false))
        {
            var node = handler.Get("/slide[1]/shape[1]");
            node.Format.Should().ContainKey("margin", "margin/inset must be readable after reopen");
        }
    }

    [Fact]
    public void Pptx_HexagonShape_MarginRoundtrip()
    {
        var path = CreateTemp("pptx");

        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new());
            handler.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = "Hex",
                ["preset"] = "hexagon",
                ["margin"] = "0.3cm"
            });
        }

        using (var handler = new PowerPointHandler(path, editable: false))
        {
            var node = handler.Get("/slide[1]/shape[1]");
            node.Format.Should().ContainKey("margin", "margin/inset must be readable after reopen");
        }
    }

    [Fact]
    public void Pptx_HexagonShape_SchemaValidAfterSave()
    {
        var path = CreateTemp("pptx");

        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new());
            handler.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = "Hexagon Shape",
                ["preset"] = "hexagon",
                ["fill"] = "#7030A0",
                ["bold"] = "true"
            });
        }

        AssertValidPptx(path, "Hexagon shape with fill and bold");
    }

    [Fact]
    public void Pptx_NotchedRightArrow_AddGetVerify()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());

        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Notched Arrow",
            ["preset"] = "notchedRightArrow"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("Notched Arrow");
        node.Format["preset"].ToString().Should().Be("notchedRightArrow");
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // SECTION 3 — Word Watermark: schema validity (EvenAndOddHeaders/TitlePage fix)
    // ═══════════════════════════════════════════════════════════════════════════

    [Fact]
    public void Word_Watermark_SchemaValidAfterAdd()
    {
        var path = CreateTemp("docx");

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Body content" });
            handler.Add("/", "watermark", null, new() { ["text"] = "DRAFT" });
        }

        AssertValidDocx(path, "Add watermark");
    }

    [Fact]
    public void Word_Watermark_TextRoundtrip()
    {
        var path = CreateTemp("docx");

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Content" });
            handler.Add("/", "watermark", null, new() { ["text"] = "SAMPLE" });
        }

        using (var handler = new WordHandler(path, editable: false))
        {
            var node = handler.Get("/watermark");
            node.Should().NotBeNull();
            node.Type.Should().Be("watermark");
            node.Format["text"].Should().Be("SAMPLE");
        }
    }

    [Fact]
    public void Word_Watermark_ReplaceSchemaValid()
    {
        var path = CreateTemp("docx");

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Body" });
            handler.Add("/", "watermark", null, new() { ["text"] = "DRAFT" });
            handler.Add("/", "watermark", null, new() { ["text"] = "FINAL" });
        }

        AssertValidDocx(path, "Double watermark replace");

        using (var handler = new WordHandler(path, editable: false))
        {
            var node = handler.Get("/watermark");
            node.Format["text"].Should().Be("FINAL", "second watermark replaces first");
        }
    }

    [Fact]
    public void Word_Watermark_SchemaValidWithAllProperties()
    {
        var path = CreateTemp("docx");

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Document body" });
            handler.Add("/", "watermark", null, new()
            {
                ["text"] = "DRAFT",
                ["color"] = "C0C0C0",
                ["font"] = "Arial",
                ["opacity"] = ".3",
                ["rotation"] = "315",
                ["width"] = "400pt",
                ["height"] = "200pt"
            });
        }

        AssertValidDocx(path, "Watermark with all properties");
    }

    [Fact]
    public void Word_Watermark_AfterFirstPageHeader_SchemaValid()
    {
        var path = CreateTemp("docx");

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Body" });
            handler.Add("/", "header", null, new() { ["type"] = "first", ["text"] = "First Page Header" });
            handler.Add("/", "watermark", null, new() { ["text"] = "DRAFT" });
        }

        AssertValidDocx(path, "First-page header then watermark");
    }

    [Fact]
    public void Word_Watermark_AfterDefaultHeader_SchemaValid()
    {
        var path = CreateTemp("docx");

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Body" });
            handler.Add("/", "header", null, new() { ["type"] = "default", ["text"] = "Default Header" });
            handler.Add("/", "watermark", null, new() { ["text"] = "DRAFT" });
        }

        AssertValidDocx(path, "Default header then watermark");
    }

    [Fact]
    public void Word_Watermark_SetColorUpdates()
    {
        var path = CreateTemp("docx");

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Content" });
            handler.Add("/", "watermark", null, new() { ["text"] = "DRAFT", ["color"] = "silver" });

            // Set + Verify
            handler.Set("/watermark", new() { ["color"] = "FF0000" });
        }

        AssertValidDocx(path, "Watermark Set color");

        using (var handler = new WordHandler(path, editable: false))
        {
            var node = handler.Get("/watermark");
            node.Format["color"].ToString().Should().ContainEquivalentOf("FF0000",
                "color should be updated to red");
        }
    }

    [Fact]
    public void Word_Watermark_RemoveSchemaValid()
    {
        var path = CreateTemp("docx");

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Content" });
            handler.Add("/", "watermark", null, new() { ["text"] = "DRAFT" });
            handler.Remove("/watermark");
        }

        AssertValidDocx(path, "After Remove watermark");
    }

    [Fact]
    public void Word_Watermark_SpecialCharsNoException()
    {
        var path = CreateTemp("docx");

        var act = () =>
        {
            using var handler = new WordHandler(path, editable: true);
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Content" });
            handler.Add("/", "watermark", null, new() { ["text"] = "Draft & <Test>" });
        };

        act.Should().NotThrow("special XML chars in watermark text must be handled");
    }

    [Fact]
    public void Word_Watermark_EmptyTextNoException()
    {
        var path = CreateTemp("docx");

        var act = () =>
        {
            using var handler = new WordHandler(path, editable: true);
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Content" });
            handler.Add("/", "watermark", null, new() { ["text"] = "" });
        };

        act.Should().NotThrow("empty watermark text should not crash");
    }

    [Fact]
    public void Word_WatermarkAfterMultipleParagraphs_SchemaValid()
    {
        var path = CreateTemp("docx");

        using (var handler = new WordHandler(path, editable: true))
        {
            for (int i = 1; i <= 5; i++)
                handler.Add("/body", "paragraph", null, new() { ["text"] = $"Paragraph {i}" });
            handler.Add("/", "watermark", null, new() { ["text"] = "DRAFT" });
        }

        AssertValidDocx(path, "Watermark after 5 paragraphs");
    }

    [Fact]
    public void Word_WatermarkWithPageSize_SchemaValid()
    {
        var path = CreateTemp("docx");

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Content" });
            handler.Set("/", new() { ["pageWidth"] = "21cm", ["pageHeight"] = "29.7cm" });
            handler.Add("/", "watermark", null, new() { ["text"] = "DRAFT" });
        }

        AssertValidDocx(path, "Page size then watermark");
    }

    [Fact]
    public void Word_Watermark_ThreeHeadersCreated()
    {
        var path = CreateTemp("docx");

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Body" });
            handler.Add("/", "watermark", null, new() { ["text"] = "DRAFT" });
        }

        using (var handler = new WordHandler(path, editable: false))
        {
            var root = handler.Get("/");
            var headerCount = root.Children.Count(c => c.Type == "header");
            headerCount.Should().BeGreaterOrEqualTo(3,
                "watermark creates default, first, and even page headers");
        }
    }

    [Fact]
    public void Word_Watermark_SchemaValidAfterReopen()
    {
        var path = CreateTemp("docx");

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Body" });
            handler.Add("/", "watermark", null, new() { ["text"] = "CONFIDENTIAL" });
        }

        AssertValidDocx(path, "After first save");

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/", "watermark", null, new() { ["text"] = "FINAL" });
        }

        AssertValidDocx(path, "After replacing watermark");
    }
}
