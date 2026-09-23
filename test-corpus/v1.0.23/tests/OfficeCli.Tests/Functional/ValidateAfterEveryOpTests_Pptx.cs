// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;
using Xunit.Abstractions;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Validate-after-every-operation tests for the PowerPoint handler.
/// Each Add or Set is followed by a flush (dispose) + OpenXML validation + reopen,
/// so that schema errors are pinpointed to the exact operation that introduced them.
/// </summary>
public class ValidateAfterEveryOpTests_Pptx : IDisposable
{
    private readonly List<string> _tempFiles = new();
    private readonly ITestOutputHelper _output;

    public ValidateAfterEveryOpTests_Pptx(ITestOutputHelper output)
    {
        _output = output;
    }

    private string CreateTemp()
    {
        var path = Path.Combine(Path.GetTempPath(), $"validate_op_{Guid.NewGuid():N}.pptx");
        _tempFiles.Add(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            try { File.Delete(f); } catch { }
    }

    /// <summary>
    /// Validates the PPTX file using OpenXmlValidator (Office2019).
    /// Throws assertion failure if any validation errors are found,
    /// with the operation name in the failure message.
    /// </summary>
    private void AssertValid(string path, string operation)
    {
        using var doc = PresentationDocument.Open(path, false);
        var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
        var errors = validator.Validate(doc).ToList();
        foreach (var e in errors)
            _output.WriteLine($"[Step: {operation}] {e.ErrorType}: {e.Description} @ {e.Path?.XPath}");
        errors.Should().BeEmpty($"after step: {operation}");
    }

    /// <summary>
    /// Flush (dispose handler), validate, then reopen and return new handler.
    /// </summary>
    private PowerPointHandler FlushValidateReopen(PowerPointHandler handler, string path, string operation)
    {
        handler.Dispose();
        AssertValid(path, operation);
        return new PowerPointHandler(path, editable: true);
    }

    // =====================================================================
    // Test 1: Sequential operations across 5 slides — validate after each
    // =====================================================================

    [Fact]
    public void SequentialOperations_ValidateAfterEachStep()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        var h = new PowerPointHandler(path, editable: true);

        // ---- Slide 1: Various shapes ----

        // Step 1: Add slide
        h.Add("/", "slide", null, new() { ["title"] = "Slide 1 - Shapes" });
        h = FlushValidateReopen(h, path, "1. Add slide 1");

        // Step 2: Add shape (rectangle) with text, fill, position, size
        h.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Rectangle",
            ["fill"] = "4472C4",
            ["x"] = "2cm",
            ["y"] = "3cm",
            ["width"] = "6cm",
            ["height"] = "3cm"
        });
        h = FlushValidateReopen(h, path, "2. Add shape with text/fill/position/size");

        // Step 3: Set shape text formatting: bold, italic, underline, color, size, font
        h.Set("/slide[1]/shape[1]", new()
        {
            ["bold"] = "true",
            ["italic"] = "true",
            ["underline"] = "true",
            ["color"] = "FF0000",
            ["size"] = "18",
            ["font"] = "Arial"
        });
        h = FlushValidateReopen(h, path, "3. Set shape bold/italic/underline/color/size/font");

        // Step 4: Set shape gradient fill
        h.Set("/slide[1]/shape[1]", new()
        {
            ["gradient"] = "radial:FF0000-0000FF-center"
        });
        h = FlushValidateReopen(h, path, "4. Set shape gradient fill");

        // Step 5: Set shape shadow effect
        h.Set("/slide[1]/shape[1]", new()
        {
            ["shadow"] = "333333"
        });
        h = FlushValidateReopen(h, path, "5. Set shape shadow effect");

        // Step 6: Set shape line color, lineWidth, lineDash
        h.Set("/slide[1]/shape[1]", new()
        {
            ["lineColor"] = "000000",
            ["lineWidth"] = "2pt",
            ["lineDash"] = "dash"
        });
        h = FlushValidateReopen(h, path, "6. Set shape lineColor/lineWidth/lineDash");

        // Step 7: Set shape opacity
        h.Set("/slide[1]/shape[1]", new()
        {
            ["opacity"] = "0.7"
        });
        h = FlushValidateReopen(h, path, "7. Set shape opacity");

        // Step 8: Set shape rotation
        h.Set("/slide[1]/shape[1]", new()
        {
            ["rotation"] = "45"
        });
        h = FlushValidateReopen(h, path, "8. Set shape rotation");

        // Step 9: Add second shape (as textbox-like shape) with text
        h.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "A textbox-like shape",
            ["x"] = "10cm",
            ["y"] = "3cm",
            ["width"] = "8cm",
            ["height"] = "2cm"
        });
        h = FlushValidateReopen(h, path, "9. Add second shape with text");

        // Step 10: Set second shape alignment
        h.Set("/slide[1]/shape[2]", new()
        {
            ["align"] = "center",
            ["valign"] = "middle"
        });
        h = FlushValidateReopen(h, path, "10. Set shape align=center, valign=middle");

        // ---- Slide 2: Table ----

        // Step 11: Add slide 2
        h.Add("/", "slide", null, new() { ["title"] = "Slide 2 - Table" });
        h = FlushValidateReopen(h, path, "11. Add slide 2");

        // Step 12: Add table (4 rows, 3 cols)
        h.Add("/slide[2]", "table", null, new()
        {
            ["rows"] = "4",
            ["cols"] = "3",
            ["width"] = "18cm",
            ["height"] = "8cm"
        });
        h = FlushValidateReopen(h, path, "12. Add table 4x3");

        // Step 13: Set table cell text
        h.Set("/slide[2]/table[1]/tr[1]/tc[1]", new() { ["text"] = "Header 1" });
        h = FlushValidateReopen(h, path, "13. Set table cell text");

        // Step 14: Set table cell fill, bold, align
        h.Set("/slide[2]/table[1]/tr[1]/tc[1]", new()
        {
            ["fill"] = "4472C4",
            ["bold"] = "true",
            ["align"] = "center"
        });
        h = FlushValidateReopen(h, path, "14. Set table cell fill/bold/align");

        // Step 15: Set table cell colspan
        h.Set("/slide[2]/table[1]/tr[2]/tc[1]", new()
        {
            ["text"] = "Merged Cell",
            ["colspan"] = "2"
        });
        h.Set("/slide[2]/table[1]/tr[2]/tc[2]", new()
        {
            ["hmerge"] = "true"
        });
        h = FlushValidateReopen(h, path, "15. Set table cell colspan=2");

        // Step 16: Set table cell border
        h.Set("/slide[2]/table[1]/tr[1]/tc[2]", new()
        {
            ["border.top"] = "2pt solid FF0000"
        });
        h = FlushValidateReopen(h, path, "16. Set table cell border.top");

        // ---- Slide 3: Chart ----

        // Step 17: Add slide 3
        h.Add("/", "slide", null, new() { ["title"] = "Slide 3 - Chart" });
        h = FlushValidateReopen(h, path, "17. Add slide 3");

        // Step 18: Add column chart with data and legend
        h.Add("/slide[3]", "chart", null, new()
        {
            ["charttype"] = "column",
            ["title"] = "Sales Report",
            ["categories"] = "Q1,Q2,Q3,Q4",
            ["series1"] = "Revenue:100,200,150,300",
            ["series2"] = "Cost:80,150,120,250",
            ["legend"] = "right"
        });
        h = FlushValidateReopen(h, path, "18. Add column chart with data and legend");

        // Step 19: Set chart title properties
        h.Set("/slide[3]/chart[1]", new()
        {
            ["title"] = "Updated Sales Report",
            ["title.bold"] = "true",
            ["title.size"] = "18"
        });
        h = FlushValidateReopen(h, path, "19. Set chart title/title.bold/title.size");

        // Step 20: Set chart legend position
        h.Set("/slide[3]/chart[1]", new()
        {
            ["legend"] = "right"
        });
        h = FlushValidateReopen(h, path, "20. Set chart legend=right");

        // ---- Slide 4: Shape with animations and notes ----

        // Step 21: Add slide 4
        h.Add("/", "slide", null, new() { ["title"] = "Slide 4 - Animations" });
        h = FlushValidateReopen(h, path, "21. Add slide 4");

        // Step 22: Add shape (ellipse) with text
        h.Add("/slide[4]", "shape", null, new()
        {
            ["text"] = "Animated Ellipse",
            ["preset"] = "ellipse",
            ["fill"] = "FFC000",
            ["x"] = "5cm",
            ["y"] = "5cm",
            ["width"] = "5cm",
            ["height"] = "5cm"
        });
        h = FlushValidateReopen(h, path, "22. Add ellipse shape with text");

        // Step 23: Add animation (fade entrance)
        h.Add("/slide[4]/shape[1]", "animation", null, new()
        {
            ["effect"] = "fade",
            ["trigger"] = "onclick"
        });
        h = FlushValidateReopen(h, path, "23. Add animation fade-entrance");

        // Step 24: Set animation (replace with fly entrance)
        h.Set("/slide[4]/shape[1]", new()
        {
            ["animation"] = "fly-entrance-left-400"
        });
        h = FlushValidateReopen(h, path, "24. Set animation=fly-entrance-left-400");

        // Step 25: Add notes to slide 4
        h.Add("/slide[4]", "notes", null, new() { ["text"] = "Speaker notes for slide 4" });
        h = FlushValidateReopen(h, path, "25. Add notes to slide 4");

        // ---- Slide 5: Group and connector ----

        // Step 26: Add slide 5
        h.Add("/", "slide", null, new() { ["title"] = "Slide 5 - Group & Connector" });
        h = FlushValidateReopen(h, path, "26. Add slide 5");

        // Step 27: Add 2 shapes for grouping
        h.Add("/slide[5]", "shape", null, new()
        {
            ["text"] = "Shape A",
            ["x"] = "2cm", ["y"] = "2cm",
            ["width"] = "3cm", ["height"] = "2cm",
            ["fill"] = "FF6347"
        });
        h.Add("/slide[5]", "shape", null, new()
        {
            ["text"] = "Shape B",
            ["x"] = "7cm", ["y"] = "2cm",
            ["width"] = "3cm", ["height"] = "2cm",
            ["fill"] = "4682B4"
        });
        h = FlushValidateReopen(h, path, "27. Add 2 shapes for grouping");

        // Step 28: Add group (shapes 1,2)
        h.Add("/slide[5]", "group", null, new() { ["shapes"] = "1,2" });
        h = FlushValidateReopen(h, path, "28. Add group from shapes 1,2");

        // Step 29: Add connector
        h.Add("/slide[5]", "connector", null, new()
        {
            ["line"] = "000000",
            ["linewidth"] = "2pt"
        });
        h = FlushValidateReopen(h, path, "29. Add connector");

        // Step 30: Set connector properties
        h.Set("/slide[5]/connector[1]", new()
        {
            ["lineColor"] = "FF0000",
            ["lineWidth"] = "3pt",
            ["lineDash"] = "dash",
            ["headEnd"] = "arrow",
            ["tailEnd"] = "diamond"
        });
        h = FlushValidateReopen(h, path, "30. Set connector line/lineWidth/lineDash/headEnd/tailEnd");

        // ---- Global operations ----

        // Step 31: Set slide 1 transition
        h.Set("/slide[1]", new() { ["transition"] = "fade" });
        h = FlushValidateReopen(h, path, "31. Set slide 1 transition=fade");

        // Step 32: Set slide 1 background color
        h.Set("/slide[1]", new() { ["background"] = "1A1A2E" });
        h = FlushValidateReopen(h, path, "32. Set slide 1 background color");

        // Step 33: Move slide 5 to position 2
        h.Move("/slide[5]", null, 1);
        h = FlushValidateReopen(h, path, "33. Move slide 5 to position 2");

        // Final verification: all 5 slides exist
        var root = h.Get("/", depth: 1);
        root.Should().NotBeNull();
        root!.Children.Count(c => c.Type == "slide").Should().Be(5,
            "all 5 slides should still exist after all operations");

        h.Dispose();
    }

    // =====================================================================
    // Test 2: Combined property Set — multiple properties in one call
    // =====================================================================

    [Fact]
    public void CombinedPropertySet_ValidateAfterEachCombo()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        var h = new PowerPointHandler(path, editable: true);

        // Setup: add a slide and a shape
        h.Add("/", "slide", null, new() { ["title"] = "Combined Props" });
        h.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Multi-Set Test",
            ["x"] = "2cm", ["y"] = "2cm",
            ["width"] = "8cm", ["height"] = "4cm"
        });
        h = FlushValidateReopen(h, path, "Setup: add slide and shape");

        // Combo 1: text formatting + fill in one call
        h.Set("/slide[1]/shape[1]", new()
        {
            ["bold"] = "true",
            ["italic"] = "true",
            ["color"] = "FFFFFF",
            ["fill"] = "2E86C1",
            ["size"] = "24"
        });
        h = FlushValidateReopen(h, path, "Combo 1: bold+italic+color+fill+size");

        // Combo 2: line + shadow in one call
        h.Set("/slide[1]/shape[1]", new()
        {
            ["lineColor"] = "000000",
            ["lineWidth"] = "1pt",
            ["shadow"] = "404040"
        });
        h = FlushValidateReopen(h, path, "Combo 2: lineColor+lineWidth+shadow");

        // Combo 3: gradient + rotation in one call
        h.Set("/slide[1]/shape[1]", new()
        {
            ["gradient"] = "radial:FF6347-4682B4-tr",
            ["rotation"] = "30"
        });
        h = FlushValidateReopen(h, path, "Combo 3: gradient+rotation");

        // Combo 4: alignment + underline + font in one call
        h.Set("/slide[1]/shape[1]", new()
        {
            ["align"] = "center",
            ["valign"] = "middle",
            ["underline"] = "true",
            ["font"] = "Calibri"
        });
        h = FlushValidateReopen(h, path, "Combo 4: align+valign+underline+font");

        // Combo 5: Add table and set multiple cell properties at once
        h.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "2",
            ["width"] = "10cm",
            ["height"] = "4cm"
        });
        h = FlushValidateReopen(h, path, "Combo 5a: add table");

        h.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Bold Header",
            ["bold"] = "true",
            ["fill"] = "4472C4",
            ["align"] = "center",
            ["valign"] = "center"
        });
        h = FlushValidateReopen(h, path, "Combo 5b: table cell text+bold+fill+align+valign");

        // Combo 6: slide-level combined properties
        h.Set("/slide[1]", new()
        {
            ["background"] = "F0F0F0",
            ["transition"] = "push"
        });
        h = FlushValidateReopen(h, path, "Combo 6: slide background+transition");

        // Combo 7: connector with all line properties at once
        h.Add("/slide[1]", "connector", null, new()
        {
            ["line"] = "000000",
            ["linewidth"] = "1pt"
        });
        h = FlushValidateReopen(h, path, "Combo 7a: add connector");

        h.Set("/slide[1]/connector[1]", new()
        {
            ["lineColor"] = "0000FF",
            ["lineWidth"] = "2pt",
            ["lineDash"] = "dot",
            ["headEnd"] = "diamond",
            ["tailEnd"] = "arrow"
        });
        h = FlushValidateReopen(h, path, "Combo 7b: connector lineColor+lineWidth+lineDash+headEnd+tailEnd");

        // Combo 8: chart with title and legend in one Set
        h.Add("/slide[1]", "chart", null, new()
        {
            ["charttype"] = "column",
            ["title"] = "Combo Chart",
            ["categories"] = "A,B,C",
            ["series1"] = "Data:10,20,30"
        });
        h = FlushValidateReopen(h, path, "Combo 8a: add chart");

        h.Set("/slide[1]/chart[1]", new()
        {
            ["title"] = "Updated Combo Chart",
            ["title.bold"] = "true",
            ["legend"] = "right"
        });
        h = FlushValidateReopen(h, path, "Combo 8b: chart title+title.bold+legend");

        h.Dispose();
    }
}
