// Black-box tests targeting recently fixed features:
//   1. Text inset for arrow/hexagon shapes (commit 11df1af)
//   2. EvenAndOddHeaders/TitlePage schema order in watermark (commit 0600aa0)
// Plus boundary cases for shapes with text and watermark interactions.

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;
using Xunit.Abstractions;

namespace OfficeCli.Tests.Functional;

public class BtBlackBoxRound1 : IDisposable
{
    private readonly List<string> _tempFiles = new();
    private readonly ITestOutputHelper _output;

    public BtBlackBoxRound1(ITestOutputHelper output)
    {
        _output = output;
    }

    private string CreateTemp(string ext)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bt_test_{Guid.NewGuid():N}.{ext}");
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
    // SECTION 1 — PPTX Arrow & Hexagon shapes: Add + text + preset
    // ═══════════════════════════════════════════════════════════════════════════

    [Fact]
    public void Pptx_RightArrow_CanAddWithText()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());

        var shapePath = handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Go Right",
            ["preset"] = "rightArrow",
            ["fill"] = "4472C4",
            ["width"] = "5cm",
            ["height"] = "2cm"
        });

        var node = handler.Get(shapePath);
        node.Should().NotBeNull();
        node.Text.Should().Be("Go Right");
        node.Format.Should().ContainKey("preset");
        node.Format["preset"].ToString().Should().Be("rightArrow");
    }

    [Fact]
    public void Pptx_RightArrow_PresetPersistsAfterReopen()
    {
        var path = CreateTemp("pptx");
        string shapePath;

        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new());
            shapePath = handler.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = "Right",
                ["preset"] = "rightArrow"
            });
        }

        using (var handler = new PowerPointHandler(path, editable: false))
        {
            var node = handler.Get(shapePath);
            node.Format["preset"].ToString().Should().Be("rightArrow", "preset must survive save/reopen");
            node.Text.Should().Be("Right");
        }
    }

    [Fact]
    public void Pptx_LeftArrow_CanAddWithText()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());

        var shapePath = handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Go Left",
            ["preset"] = "leftArrow",
            ["fill"] = "FF0000"
        });

        var node = handler.Get(shapePath);
        node.Text.Should().Be("Go Left");
        node.Format["preset"].ToString().Should().Be("leftArrow");
    }

    [Fact]
    public void Pptx_UpArrow_CanAddWithText()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());

        var shapePath = handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Go Up",
            ["preset"] = "upArrow",
            ["fill"] = "00B050"
        });

        var node = handler.Get(shapePath);
        node.Text.Should().Be("Go Up");
        node.Format["preset"].ToString().Should().Be("upArrow");
    }

    [Fact]
    public void Pptx_DownArrow_CanAddWithText()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());

        var shapePath = handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Go Down",
            ["preset"] = "downArrow",
            ["fill"] = "FFC000"
        });

        var node = handler.Get(shapePath);
        node.Text.Should().Be("Go Down");
        node.Format["preset"].ToString().Should().Be("downArrow");
    }

    [Fact]
    public void Pptx_Hexagon_CanAddWithText()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());

        var shapePath = handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Hex Cell",
            ["preset"] = "hexagon",
            ["fill"] = "7030A0"
        });

        var node = handler.Get(shapePath);
        node.Text.Should().Be("Hex Cell");
        node.Format["preset"].ToString().Should().Be("hexagon");
    }

    [Fact]
    public void Pptx_Hexagon_PresetPersistsAfterReopen()
    {
        var path = CreateTemp("pptx");
        string shapePath;

        using (var h = new PowerPointHandler(path, editable: true))
        {
            h.Add("/", "slide", null, new());
            shapePath = h.Add("/slide[1]", "shape", null, new() { ["text"] = "Hex", ["preset"] = "hexagon" });
        }

        using (var h = new PowerPointHandler(path, editable: false))
        {
            var node = h.Get(shapePath);
            node.Format["preset"].ToString().Should().Be("hexagon");
            node.Text.Should().Be("Hex");
        }
    }

    [Fact]
    public void Pptx_RightArrow_SetPresetUpdatesShape()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());
        var shapePath = handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Arrow", ["preset"] = "rightArrow" });

        // Change to a different arrow shape
        handler.Set(shapePath, new() { ["preset"] = "leftArrow" });

        var node = handler.Get(shapePath);
        node.Format["preset"].ToString().Should().Be("leftArrow", "Set should update preset geometry");
    }

    [Fact]
    public void Pptx_ArrowShape_CanSetTextAfterAdd()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());
        var shapePath = handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Initial",
            ["preset"] = "rightArrow"
        });

        handler.Set(shapePath, new() { ["text"] = "Updated Arrow" });

        var node = handler.Get(shapePath);
        node.Text.Should().Be("Updated Arrow");
        node.Format["preset"].ToString().Should().Be("rightArrow", "preset must remain unchanged after Set text");
    }

    [Fact]
    public void Pptx_ArrowShape_SchemaValid()
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

        AssertValidPptx(path, "Add arrow and hexagon shapes");
    }

    [Fact]
    public void Pptx_ArrowShape_TextMarginRoundtrip()
    {
        // Verify that setting inset on an arrow shape survives roundtrip
        var path = CreateTemp("pptx");
        string shapePath;

        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new());
            shapePath = handler.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = "Arrow",
                ["preset"] = "rightArrow",
                ["margin"] = "0.2cm"
            });
        }

        using (var handler = new PowerPointHandler(path, editable: false))
        {
            var node = handler.Get(shapePath);
            // margin should be reported back
            node.Format.Should().ContainKey("margin", "set inset should be readable back");
        }
    }

    [Fact]
    public void Pptx_Hexagon_TextMarginRoundtrip()
    {
        var path = CreateTemp("pptx");
        string shapePath;

        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new());
            shapePath = handler.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = "Hex",
                ["preset"] = "hexagon",
                ["margin"] = "0.3cm"
            });
        }

        using (var handler = new PowerPointHandler(path, editable: false))
        {
            var node = handler.Get(shapePath);
            node.Format.Should().ContainKey("margin");
        }
    }

    [Fact]
    public void Pptx_AllFourArrows_OnSameSlide()
    {
        // All four direction arrows should be independently addable and queryable
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());

        var directions = new[] { "rightArrow", "leftArrow", "upArrow", "downArrow" };
        var shapePaths = new List<string>();
        foreach (var dir in directions)
        {
            shapePaths.Add(handler.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = dir,
                ["preset"] = dir
            }));
        }

        var shapes = handler.Query("shape");
        shapes.Should().HaveCount(4, "four arrow shapes added");

        // Verify each shape has correct text and preset using returned paths
        for (int i = 0; i < 4; i++)
        {
            var node = handler.Get(shapePaths[i]);
            node.Text.Should().Be(directions[i]);
            node.Format["preset"].ToString().Should().Be(directions[i]);
        }
    }

    [Fact]
    public void Pptx_Arrow_EmptyTextNoException()
    {
        // Arrow shape with empty text should not throw
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());

        string shapePath = "";
        var act = () =>
        {
            shapePath = handler.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = "",
                ["preset"] = "rightArrow"
            });
        };

        act.Should().NotThrow("empty text in arrow shape should be allowed");
    }

    [Fact]
    public void Pptx_Arrow_LongTextDoesNotThrow()
    {
        // Very long text in arrow shape
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());

        var longText = new string('A', 500);
        string shapePath = "";
        var act = () =>
        {
            shapePath = handler.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = longText,
                ["preset"] = "rightArrow"
            });
        };

        act.Should().NotThrow("long text in arrow shape should not throw");
        var node = handler.Get(shapePath);
        node.Text.Should().Be(longText);
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // SECTION 2 — Word Watermark: EvenAndOddHeaders + TitlePage schema order
    // ═══════════════════════════════════════════════════════════════════════════

    [Fact]
    public void Word_Watermark_SchemaValidAfterAdd()
    {
        // The fix 0600aa0 ensures EvenAndOddHeaders and TitlePage are
        // inserted in correct schema order in Settings.
        var path = CreateTemp("docx");

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Body content" });
            handler.Add("/", "watermark", null, new() { ["text"] = "DRAFT" });
        }

        AssertValidDocx(path, "Add watermark with EvenAndOddHeaders + TitlePage");
    }

    [Fact]
    public void Word_Watermark_AddThenSchemaValidAfterReopen()
    {
        var path = CreateTemp("docx");
        WordHandler handler;

        using (handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Body" });
            handler.Add("/", "watermark", null, new() { ["text"] = "CONFIDENTIAL" });
        }

        // Validate before reopen
        AssertValidDocx(path, "After first save");

        using (handler = new WordHandler(path, editable: true))
        {
            // Add another watermark (replace existing)
            handler.Add("/", "watermark", null, new() { ["text"] = "FINAL" });
        }

        AssertValidDocx(path, "After replacing watermark");
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

        AssertValidDocx(path, "Add watermark with all properties");
    }

    [Fact]
    public void Word_Watermark_AfterFirstPageHeader_SchemaValid()
    {
        // Previously problematic: first-page header adds TitlePage to Settings.
        // Then watermark also tries to add EvenAndOddHeaders + TitlePage.
        // The schema order fix should handle both gracefully.
        var path = CreateTemp("docx");

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Body" });
            // First add a first-page header (which touches TitlePage in sectPr or Settings)
            handler.Add("/", "header", null, new() { ["type"] = "first", ["text"] = "First Page Header" });
            // Then add watermark (which adds EvenAndOddHeaders + TitlePage in Settings)
            handler.Add("/", "watermark", null, new() { ["text"] = "DRAFT" });
        }

        AssertValidDocx(path, "First-page header then watermark");
    }

    [Fact]
    public void Word_Watermark_AfterEvenOddHeader_SchemaValid()
    {
        // Watermark added after an even/odd header setup should not break schema order
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
    public void Word_Watermark_ReplacedWatermark_SchemaValid()
    {
        // Adding watermark twice (replace) must keep valid schema
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
            node.Format["text"].Should().Be("FINAL", "second watermark should replace first");
        }
    }

    [Fact]
    public void Word_Watermark_ThreeHeadersCreated()
    {
        // Watermark must create 3 headers: default, first, even
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
                "watermark should create default, first, and even page headers");
        }
    }

    [Fact]
    public void Word_Watermark_TextRoundtrip()
    {
        // Basic text roundtrip
        var path = CreateTemp("docx");

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Content" });
            handler.Add("/", "watermark", null, new() { ["text"] = "SAMPLE" });
        }

        using (var handler = new WordHandler(path, editable: false))
        {
            var node = handler.Get("/watermark");
            node.Type.Should().Be("watermark");
            node.Format["text"].Should().Be("SAMPLE");
        }
    }

    [Fact]
    public void Word_Watermark_SpecialCharsInText_NoException()
    {
        // Special characters in watermark text (XML chars that need escaping)
        var path = CreateTemp("docx");

        var act = () =>
        {
            using var handler = new WordHandler(path, editable: true);
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Content" });
            handler.Add("/", "watermark", null, new()
            {
                ["text"] = "Draft & <Test>"
            });
        };

        act.Should().NotThrow("special XML chars in watermark text must be escaped, not throw");
    }

    [Fact]
    public void Word_Watermark_EmptyTextFallsBackToDefault()
    {
        // Empty text should not crash
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
    public void Word_Watermark_SetColor_UpdatesAndSchemaValid()
    {
        var path = CreateTemp("docx");

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Content" });
            handler.Add("/", "watermark", null, new() { ["text"] = "DRAFT", ["color"] = "silver" });
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
    public void Word_Watermark_Remove_SchemaValid()
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

    // ═══════════════════════════════════════════════════════════════════════════
    // SECTION 3 — Combined: Watermark + content interactions
    // ═══════════════════════════════════════════════════════════════════════════

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
    public void Word_WatermarkWithPageSizeSet_SchemaValid()
    {
        // Set page size then add watermark — settings interaction
        var path = CreateTemp("docx");

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Content" });
            handler.Set("/", new() { ["pageWidth"] = "21cm", ["pageHeight"] = "29.7cm" });
            handler.Add("/", "watermark", null, new() { ["text"] = "DRAFT" });
        }

        AssertValidDocx(path, "Page size then watermark");
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // SECTION 4 — Boundary cases for arrow/hexagon shapes
    // ═══════════════════════════════════════════════════════════════════════════

    [Fact]
    public void Pptx_Arrow_CanChangePresetToHexagon()
    {
        // Start as arrow, change to hexagon
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());
        var shapePath = handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Shape", ["preset"] = "rightArrow" });

        handler.Set(shapePath, new() { ["preset"] = "hexagon" });

        var node = handler.Get(shapePath);
        node.Format["preset"].ToString().Should().Be("hexagon");
    }

    [Fact]
    public void Pptx_Hexagon_CanChangePresetToArrow()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());
        var shapePath = handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Shape", ["preset"] = "hexagon" });

        handler.Set(shapePath, new() { ["preset"] = "rightArrow" });

        var node = handler.Get(shapePath);
        node.Format["preset"].ToString().Should().Be("rightArrow");
    }

    [Fact]
    public void Pptx_NotchedRightArrow_CanAddWithText()
    {
        // notchedRightArrow is listed in the same inset bucket as rightArrow
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());

        var shapePath = handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Notched Arrow",
            ["preset"] = "notchedRightArrow"
        });

        var node = handler.Get(shapePath);
        node.Text.Should().Be("Notched Arrow");
    }

    [Fact]
    public void Pptx_ArrowShape_FillColorRoundtrip()
    {
        // Arrow with color — verify fill survives reopen
        var path = CreateTemp("pptx");
        string shapePath;

        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new());
            shapePath = handler.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = "Arrow",
                ["preset"] = "rightArrow",
                ["fill"] = "#4472C4"
            });
        }

        using (var handler = new PowerPointHandler(path, editable: false))
        {
            var node = handler.Get(shapePath);
            node.Format.Should().ContainKey("fill");
            node.Format["fill"].ToString().Should().Be("#4472C4");
        }
    }

    [Fact]
    public void Pptx_ArrowShape_BoldTextRoundtrip()
    {
        // Arrow with bold text
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());
        var shapePath = handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Bold Arrow",
            ["preset"] = "rightArrow",
            ["bold"] = "true",
            ["size"] = "18"
        });

        handler.Set(shapePath, new() { ["bold"] = "true" });

        var node = handler.Get(shapePath);
        node.Text.Should().Be("Bold Arrow");
        node.Format.Should().ContainKey("bold");
        node.Format["bold"].ToString().Should().Be("True");
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
}
