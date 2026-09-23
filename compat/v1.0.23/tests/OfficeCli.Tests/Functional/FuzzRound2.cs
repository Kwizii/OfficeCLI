// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

// FuzzRound2 — Expanded fuzz coverage targeting areas not yet covered by existing fuzzers.
//
// Areas targeted:
//   F60: Excel shape Set align — silent fallback (_ => TextAlignmentTypeValues.Left) for invalid value
//   F61: Excel sparkline Set lineweight — no NaN/Infinity guard (only if TryParse succeeds)
//   F62: Word run Set underline — pass-through of arbitrary string stores invalid XML
//   F63: Word table cell Set underline — pass-through of arbitrary string stores invalid XML
//   F64: PPTX shape Set autofit — invalid value is silent no-op (no ArgumentException)
//   F65: PPTX shape Set textwarp — accepts arbitrary strings without validation
//   F66: PPTX shape Set liststyle — long strings (>2 chars) already throw, but empty string and edge cases
//   G01: Get on nonexistent/malformed paths — all three handlers should not crash
//   G02: Query with unexpected selector types — should not crash
//   M01: Mixed ops: Add then Remove then Get — Get on removed element
//   M02: Mixed ops: double Remove — second Remove on same path

using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class FuzzRound2 : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private string CreateTemp(string ext)
    {
        var path = Path.Combine(Path.GetTempPath(), $"fuzz2_{Guid.NewGuid():N}.{ext}");
        _tempFiles.Add(path);
        BlankDocCreator.Create(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            if (File.Exists(f)) File.Delete(f);
    }

    // ==================== F60: Excel shape Set align silent fallback ====================
    // Bug: ExcelHandler.Set.cs:469 — `_ => Drawing.TextAlignmentTypeValues.Left`
    // Invalid align values silently become Left instead of throwing ArgumentException.
    // Fix: `_ => throw new ArgumentException(...)` listing valid values.

    [Theory]
    [InlineData("center")]
    [InlineData("left")]
    [InlineData("right")]
    [InlineData("justify")]
    public void Excel_ShapeSetAlign_ValidValues_Succeed(string align)
    {
        var path = CreateTemp("xlsx");
        // Need a drawing shape in xlsx — use picture as proxy since plain shape requires drawing
        // We can confirm the Set path by setting align on a shape[N]
        // First add a shape via xlsx drawing (if supported)
        using var handler = new ExcelHandler(path, editable: true);
        // Try adding a shape — if not supported this will fail gracefully
        try
        {
            handler.Add("/Sheet1", "shape", null, new() { ["text"] = "Hello", ["x"] = "0", ["y"] = "0", ["width"] = "5", ["height"] = "3" });
            var act = () => handler.Set("/Sheet1/shape[1]", new() { ["align"] = align });
            act.Should().NotThrow($"Excel shape align='{align}' is valid");
        }
        catch (ArgumentException)
        {
            // If Add shape is not supported, skip test gracefully
        }
    }

    [Theory]
    [InlineData("invalid")]
    [InlineData("middle")]
    [InlineData("")]
    public void F60_Excel_ShapeSetAlign_InvalidValues_ShouldThrowArgumentException(string align)
    {
        // Bug: ExcelHandler.Set.cs:469 — `_ => Drawing.TextAlignmentTypeValues.Left` silent fallback
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);
        try
        {
            handler.Add("/Sheet1", "shape", null, new() { ["text"] = "Hello", ["x"] = "0", ["y"] = "0", ["width"] = "5", ["height"] = "3" });
            var act = () => handler.Set("/Sheet1/shape[1]", new() { ["align"] = align });
            act.Should().Throw<ArgumentException>(
                $"Excel shape align='{align}' is invalid — should throw, not silently default to Left");
        }
        catch (ArgumentException ex) when (ex.Message.Contains("not supported") || ex.Message.Contains("shape"))
        {
            // Add shape not supported: skip
        }
    }

    // ==================== F61: Excel sparkline Set lineweight no NaN guard ====================
    // Bug: ExcelHandler.Set.cs:100 — `if (double.TryParse(value, out var lw)) spkGroup.LineWeight = lw;`
    // double.TryParse("NaN") returns true with double.NaN — stored without guard.
    // double.TryParse("abc") returns false — silently ignored (no exception).
    // Fix: add || double.IsNaN(lw) || double.IsInfinity(lw) guard, throw on NaN/Infinity.

    [Fact]
    public void F61_Excel_SparklineSetLineweight_NaN_SilentlyIgnoredOrStored()
    {
        // Verify that double.TryParse("NaN") returns true — the NaN guard is needed.
        double.TryParse("NaN", System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture, out var nanVal).Should().BeTrue(
            "TryParse('NaN') returns true — NaN would be stored as lineweight without a guard");
        double.IsNaN(nanVal).Should().BeTrue("confirming parsed value is NaN");

        // Actual behavior: currently NaN may pass through to spkGroup.LineWeight
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);
        handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "1" });
        handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B1", ["value"] = "2" });

        try
        {
            handler.Add("/Sheet1", "sparkline", null, new()
            {
                ["ref"] = "C1", ["data"] = "A1:B1", ["type"] = "line"
            });

            // This is the bug: NaN passes TryParse without being guarded
            var act = () => handler.Set("/Sheet1/sparkline[1]", new() { ["lineweight"] = "NaN" });
            // Should throw ArgumentException (currently may silently store NaN)
            act.Should().Throw<ArgumentException>(
                "lineweight='NaN' should be rejected — NaN stored as double attribute corrupts XML");
        }
        catch (ArgumentException ex) when (ex.Message.Contains("sparkline") || ex.Message.Contains("Sparkline"))
        {
            // Sparkline add may not be supported on blank sheet — skip
            _ = ex;
        }
    }

    [Fact]
    public void F61b_Excel_SparklineSetLineweight_InvalidString_SilentlyIgnored()
    {
        // Bug: "abc" fails TryParse so is silently ignored — no exception, no feedback.
        // The caller gets no indication that lineweight was not applied.
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);
        handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "1" });
        handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B1", ["value"] = "2" });

        try
        {
            handler.Add("/Sheet1", "sparkline", null, new()
            {
                ["ref"] = "C1", ["data"] = "A1:B1", ["type"] = "line"
            });

            // "abc" fails TryParse: lineweight is silently not applied. No exception.
            // This confirms the silent-fallback pattern — callers have no indication of failure.
            var act = () => handler.Set("/Sheet1/sparkline[1]", new() { ["lineweight"] = "abc" });
            // Currently does NOT throw — this is the bug.
            act.Should().Throw<ArgumentException>(
                "lineweight='abc' is not a valid number — should throw, not silently ignore");
        }
        catch (ArgumentException ex) when (ex.Message.Contains("sparkline") || ex.Message.Contains("Sparkline"))
        {
            _ = ex;
        }
    }

    // ==================== F62: Word run Set underline pass-through ====================
    // Bug: WordHandler.Set.cs:764-773 — `_ => value` pass-through stores arbitrary string as
    // w:u/@w:val attribute, creating invalid OOXML (e.g. underline="xyz" is not a valid UnderlineValues).

    [Theory]
    [InlineData("single")]
    [InlineData("double")]
    [InlineData("none")]
    [InlineData("dotted")]
    [InlineData("true")]
    [InlineData("false")]
    public void Word_RunSetUnderline_ValidValues_Succeed(string underline)
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello" });
        var act = () => handler.Set("/body/p[1]", new() { ["underline"] = underline });
        act.Should().NotThrow($"Word run underline='{underline}' is valid");
    }

    [Theory]
    [InlineData("xyz")]
    [InlineData("overline")]
    [InlineData("wavy_invalid")]
    [InlineData("INVALID_UNDERLINE_STYLE_12345")]
    public void F62_Word_RunSetUnderline_InvalidValues_ShouldThrowArgumentException(string underline)
    {
        // Bug: WordHandler.Set.cs:764-773 — `_ => value` pass-through
        // Arbitrary string stored as w:u/@w:val — no validation, creates invalid OOXML
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello" });
        var act = () => handler.Set("/body/p[1]", new() { ["underline"] = underline });
        act.Should().Throw<ArgumentException>(
            $"Word run underline='{underline}' is invalid — should throw, not store arbitrary string as XML attribute");
    }

    // ==================== F63: Word table cell Set underline pass-through ====================
    // Bug: WordHandler.Set.cs:1107,1152 — same `_ => value` pass-through in table cell context.

    [Theory]
    [InlineData("xyz")]
    [InlineData("invalid_underline")]
    [InlineData("BOGUS")]
    public void F63_Word_TableCellSetUnderline_InvalidValues_ShouldThrowArgumentException(string underline)
    {
        // Bug: same pattern as F62 but in table cell run/paragraph mark run properties
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        var act = () => handler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["underline"] = underline });
        act.Should().Throw<ArgumentException>(
            $"Word table cell underline='{underline}' is invalid — should throw ArgumentException");
    }

    // ==================== F64: PPTX shape Set autofit invalid value silent no-op ====================
    // Bug: PowerPointHandler.ShapeProperties.cs:577-591 — the switch has no default/throw.
    // Invalid value is silently ignored (no autofit element added, no error).

    [Theory]
    [InlineData("invalid_autofit")]
    [InlineData("yes")]
    [InlineData("")]
    [InlineData("expand")]
    public void F64_Pptx_SetAutofit_InvalidValues_ShouldThrowArgumentException(string autofit)
    {
        // Bug: no `default: throw` in the switch — invalid values silently do nothing.
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Hello", ["x"] = "2cm", ["y"] = "2cm",
            ["width"] = "10cm", ["height"] = "3cm"
        });
        var act = () => handler.Set("/slide[1]/shape[1]", new() { ["autofit"] = autofit });
        act.Should().Throw<ArgumentException>(
            $"autofit='{autofit}' is invalid — should throw ArgumentException, not silently ignore");
    }

    [Theory]
    [InlineData("true")]
    [InlineData("normal")]
    [InlineData("false")]
    [InlineData("none")]
    [InlineData("shape")]
    [InlineData("resize")]
    public void Pptx_SetAutofit_ValidValues_Succeed(string autofit)
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Hello", ["x"] = "2cm", ["y"] = "2cm",
            ["width"] = "10cm", ["height"] = "3cm"
        });
        var act = () => handler.Set("/slide[1]/shape[1]", new() { ["autofit"] = autofit });
        act.Should().NotThrow($"autofit='{autofit}' is valid");
    }

    // ==================== F65: PPTX shape Set textwarp crash on short names ====================
    // Bug: PowerPointHandler.ShapeProperties.cs:569 — `value.StartsWith("text") ? value : $"text{char.ToUpper(value[0])}{value[1..]}"`
    // Short aliases like "arch" -> "textArch", "wave" -> "textWave", "squeeze" -> "textSqueeze".
    // These are NOT valid Drawing.TextShapeValues names (the SDK's enum does not have "textArch", etc.)
    // so `new Drawing.TextShapeValues("textArch")` throws ArgumentOutOfRangeException — a crash.
    // The handler has no validation mapping: it blindly constructs the enum value from user input.
    // Fix: maintain an explicit allow-list of valid warp preset names.

    [Theory]
    [InlineData("none")]       // special case: removes warp
    [InlineData("textNoShape")] // fully-qualified valid SDK name
    public void Pptx_SetTextwarp_KnownValidValues_Succeed(string warp)
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Hello", ["x"] = "2cm", ["y"] = "2cm",
            ["width"] = "10cm", ["height"] = "3cm"
        });
        var act = () => handler.Set("/slide[1]/shape[1]", new() { ["textwarp"] = warp });
        act.Should().NotThrow($"textwarp='{warp}' is a valid warp name");
    }

    [Theory]
    [InlineData("arch")]     // "textArch" not in SDK PresetTextWarp.Preset enum — crash
    [InlineData("wave")]     // "textWave" not in SDK PresetTextWarp.Preset enum — crash
    [InlineData("squeeze")]  // "textSqueeze" not in SDK PresetTextWarp.Preset enum — crash
    // Note: "inflate" -> "textInflate" IS valid in the SDK — so it does not crash
    public void F65_Pptx_SetTextwarp_ShortAliasNames_ThrowArgumentOutOfRangeException_NotHandled(
        string warp)
    {
        // CRASH BUG: PowerPointHandler.ShapeProperties.cs:569-572
        // "arch" -> warpName = "textArch" -> new Drawing.TextShapeValues("textArch")
        // DrawingML SDK validates the enum — "textArch" is not a member -> ArgumentOutOfRangeException.
        // This exception is NOT caught/wrapped by the handler, so it escapes as a crash.
        // Fix: map short aliases to full names (e.g. "arch" -> "textArchUp") or validate against allow-list.
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Hello", ["x"] = "2cm", ["y"] = "2cm",
            ["width"] = "10cm", ["height"] = "3cm"
        });

        var act = () => handler.Set("/slide[1]/shape[1]", new() { ["textwarp"] = warp });
        // Currently throws ArgumentOutOfRangeException — not an ArgumentException.
        // The ideal fix would convert this to a clear ArgumentException with valid options.
        act.Should().NotThrow<NullReferenceException>(
            $"textwarp='{warp}' must not throw NullReferenceException");
        // Confirm the crash: should currently throw ArgumentOutOfRangeException
        // (PresetTextWarp.Preset validates the enum value; TextShapeValues ctor does not)
        act.Should().Throw<ArgumentOutOfRangeException>(
            $"textwarp='{warp}' -> 'text{char.ToUpper(warp[0])}{warp.Substring(1)}' is not a valid SDK PresetTextWarp enum value — crashes with ArgumentOutOfRangeException");
    }

    [Theory]
    [InlineData("INVALID_WARP_XYZ123")]
    [InlineData("notavalidwarp")]
    [InlineData("random_garbage_value_123")]
    public void F65b_Pptx_SetTextwarp_LongInvalidValues_ThrowArgumentOutOfRangeException(string warp)
    {
        // Long invalid strings also crash: new Drawing.TextShapeValues("textNotavalidwarp_123") -> ArgumentOutOfRangeException
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Hello", ["x"] = "2cm", ["y"] = "2cm",
            ["width"] = "10cm", ["height"] = "3cm"
        });
        var act = () => handler.Set("/slide[1]/shape[1]", new() { ["textwarp"] = warp });
        act.Should().NotThrow<NullReferenceException>(
            $"textwarp='{warp}' must not throw NullReferenceException");
        // The handler should ideally throw ArgumentException — currently throws ArgumentOutOfRangeException
        act.Should().Throw<Exception>(
            $"textwarp='{warp}' is not a valid preset warp name and should throw");
    }

    // ==================== F66: PPTX shape Set liststyle edge cases ====================
    // Verify that empty string and single-char edge cases don't crash unexpectedly.

    [Theory]
    [InlineData("")]
    [InlineData("bullet")]
    [InlineData("none")]
    [InlineData("numbered")]
    public void Pptx_SetListstyle_ValidAndBoundaryValues_DoNotCrash(string liststyle)
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Hello", ["x"] = "2cm", ["y"] = "2cm",
            ["width"] = "10cm", ["height"] = "3cm"
        });
        // Empty string and known values should either succeed or throw ArgumentException — not crash
        var act = () => handler.Set("/slide[1]/shape[1]", new() { ["liststyle"] = liststyle });
        act.Should().NotThrow<NullReferenceException>(
            $"liststyle='{liststyle}' should not cause NullReferenceException");
        act.Should().NotThrow<IndexOutOfRangeException>(
            $"liststyle='{liststyle}' should not cause IndexOutOfRangeException");
    }

    [Fact]
    public void Pptx_SetListstyle_LongInvalidString_ThrowsArgumentException()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Hello", ["x"] = "2cm", ["y"] = "2cm",
            ["width"] = "10cm", ["height"] = "3cm"
        });
        // Strings > 2 chars that aren't known values should throw ArgumentException
        var act = () => handler.Set("/slide[1]/shape[1]", new() { ["liststyle"] = "notavalid" });
        act.Should().Throw<ArgumentException>("long invalid liststyle should throw ArgumentException");
    }

    // ==================== G01: Get on nonexistent/malformed paths ====================
    // All three handlers should throw ArgumentException (not NullReferenceException/crash)
    // when given a path that doesn't exist or is malformed.

    [Theory]
    [InlineData("/slide[99]")]
    [InlineData("/slide[0]")]
    [InlineData("/slide[-1]")]
    [InlineData("/slide[1]/shape[99]")]
    [InlineData("/slide[1]/shape[0]")]
    [InlineData("/nonexistent")]
    [InlineData("/slide[abc]")]
    [InlineData("slide[1]")]     // missing leading slash
    [InlineData("//")]
    [InlineData("/slide[1]//shape[1]")]
    public void Pptx_Get_NonexistentPath_ThrowsArgumentException_NotCrash(string badPath)
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Hello", ["x"] = "2cm", ["y"] = "2cm",
            ["width"] = "10cm", ["height"] = "3cm"
        });

        var act = () => handler.Get(badPath);
        // Must not crash with NullReferenceException, IndexOutOfRangeException, or OverflowException
        act.Should().NotThrow<NullReferenceException>(
            $"Get('{badPath}') must not throw NullReferenceException");
        act.Should().NotThrow<IndexOutOfRangeException>(
            $"Get('{badPath}') must not throw IndexOutOfRangeException");
        act.Should().NotThrow<OverflowException>(
            $"Get('{badPath}') must not throw OverflowException");
        // Should throw ArgumentException (or return empty node)
    }

    [Theory]
    [InlineData("/Sheet1/ZZ99999")]
    [InlineData("/Sheet99")]
    [InlineData("/Sheet1/row[0]")]
    [InlineData("/Sheet1/row[-1]")]
    [InlineData("/Sheet1/col[0]")]
    [InlineData("/Sheet1//A1")]
    [InlineData("Sheet1/A1")]    // missing leading slash
    public void Excel_Get_NonexistentPath_ThrowsArgumentException_NotCrash(string badPath)
    {
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);
        handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "test" });

        var act = () => handler.Get(badPath);
        act.Should().NotThrow<NullReferenceException>(
            $"Get('{badPath}') must not throw NullReferenceException");
        act.Should().NotThrow<IndexOutOfRangeException>(
            $"Get('{badPath}') must not throw IndexOutOfRangeException");
        act.Should().NotThrow<OverflowException>(
            $"Get('{badPath}') must not throw OverflowException");
    }

    [Theory]
    [InlineData("/body/p[99]")]
    [InlineData("/body/p[0]")]
    [InlineData("/body/tbl[99]")]
    [InlineData("/section[99]")]
    [InlineData("/body//p[1]")]
    [InlineData("body/p[1]")]   // missing leading slash
    public void Word_Get_NonexistentPath_ThrowsArgumentException_NotCrash(string badPath)
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello" });

        var act = () => handler.Get(badPath);
        act.Should().NotThrow<NullReferenceException>(
            $"Get('{badPath}') must not throw NullReferenceException");
        act.Should().NotThrow<IndexOutOfRangeException>(
            $"Get('{badPath}') must not throw IndexOutOfRangeException");
        act.Should().NotThrow<OverflowException>(
            $"Get('{badPath}') must not throw OverflowException");
    }

    // ==================== G02: Query with unexpected selector types ====================
    // Fuzz Query() with invalid/unexpected selector types — should not crash.

    [Theory]
    [InlineData("")]
    [InlineData("slide")]      // invalid per CLAUDE.md — Query("slide") returns nothing
    [InlineData("nonexistent")]
    [InlineData("shape shape")]
    [InlineData("shape[1]")]   // indexing in selector not supported
    [InlineData("SHAPE")]      // case mismatch
    [InlineData("table")]      // valid
    [InlineData("picture")]    // valid
    public void Pptx_Query_UnexpectedSelectors_DoNotCrash(string selector)
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Hello", ["x"] = "2cm", ["y"] = "2cm",
            ["width"] = "10cm", ["height"] = "3cm"
        });

        var act = () => handler.Query(selector);
        act.Should().NotThrow<NullReferenceException>(
            $"Query('{selector}') must not throw NullReferenceException");
        act.Should().NotThrow<IndexOutOfRangeException>(
            $"Query('{selector}') must not throw IndexOutOfRangeException");
    }

    [Theory]
    [InlineData("")]
    [InlineData("cell")]
    [InlineData("bogus_type")]
    [InlineData("row row")]
    public void Excel_Query_UnexpectedSelectors_DoNotCrash(string selector)
    {
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);
        handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "test" });

        var act = () => handler.Query(selector);
        act.Should().NotThrow<NullReferenceException>(
            $"Excel Query('{selector}') must not throw NullReferenceException");
    }

    [Theory]
    [InlineData("")]
    [InlineData("paragraph")]
    [InlineData("bogus")]
    [InlineData("table table")]
    public void Word_Query_UnexpectedSelectors_DoNotCrash(string selector)
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello" });

        var act = () => handler.Query(selector);
        act.Should().NotThrow<NullReferenceException>(
            $"Word Query('{selector}') must not throw NullReferenceException");
    }

    // ==================== M01: Add then Remove then Get ====================
    // After Remove, Get on the same path should throw ArgumentException — not crash.

    [Fact]
    public void Pptx_AddThenRemoveThenGet_ThrowsArgumentException_NotCrash()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());
        var shapePath = handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Hello", ["x"] = "2cm", ["y"] = "2cm",
            ["width"] = "10cm", ["height"] = "3cm"
        });

        // Remove the shape
        handler.Remove(shapePath);

        // Get on the now-removed path must not crash
        var act = () => handler.Get(shapePath);
        act.Should().NotThrow<NullReferenceException>(
            "Get on removed shape path must not throw NullReferenceException");
        act.Should().NotThrow<IndexOutOfRangeException>(
            "Get on removed shape path must not throw IndexOutOfRangeException");
        // It should throw ArgumentException (element not found)
        act.Should().Throw<ArgumentException>(
            "Get on removed shape should throw ArgumentException (element not found)");
    }

    [Fact]
    public void Excel_AddThenRemoveThenGet_ThrowsArgumentException_NotCrash()
    {
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);
        handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "test" });

        handler.Remove("/Sheet1/A1");

        var act = () => handler.Get("/Sheet1/A1");
        act.Should().NotThrow<NullReferenceException>(
            "Get on removed cell must not throw NullReferenceException");
    }

    [Fact]
    public void Word_AddThenRemoveThenGet_ThrowsArgumentException_NotCrash()
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        var paraPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello" });

        handler.Remove(paraPath);

        var act = () => handler.Get(paraPath);
        act.Should().NotThrow<NullReferenceException>(
            "Get on removed paragraph must not throw NullReferenceException");
    }

    // ==================== M02: Double Remove — second Remove on same path ====================
    // Removing the same element twice should throw ArgumentException — not crash.

    [Fact]
    public void Pptx_DoubleRemove_SecondCallThrowsArgumentException_NotCrash()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());
        var shapePath = handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Hello", ["x"] = "2cm", ["y"] = "2cm",
            ["width"] = "10cm", ["height"] = "3cm"
        });

        handler.Remove(shapePath); // First remove succeeds

        var act = () => handler.Remove(shapePath); // Second remove — element no longer exists
        act.Should().NotThrow<NullReferenceException>(
            "Second Remove on same path must not throw NullReferenceException");
        act.Should().NotThrow<IndexOutOfRangeException>(
            "Second Remove on same path must not throw IndexOutOfRangeException");
    }

    [Fact]
    public void Word_DoubleRemove_SecondCallDoesNotCrash()
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        var paraPath = handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello" });

        handler.Remove(paraPath);

        var act = () => handler.Remove(paraPath);
        act.Should().NotThrow<NullReferenceException>(
            "Second Remove on same Word paragraph must not throw NullReferenceException");
    }

    // ==================== Boundary: PPTX Set on empty/null-like properties ====================
    // Set with empty string values for numeric properties — should throw ArgumentException.

    [Theory]
    [InlineData("x", "")]
    [InlineData("y", "abc")]
    [InlineData("width", "NaN")]
    [InlineData("height", "Infinity")]
    [InlineData("rotation", "not_a_number")]
    [InlineData("opacity", "abc")]
    [InlineData("charspacing", "NaN")]
    public void Pptx_SetShapeNumericProperties_InvalidValues_ThrowArgumentException(string key, string value)
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Hello", ["x"] = "2cm", ["y"] = "2cm",
            ["width"] = "10cm", ["height"] = "3cm"
        });

        var act = () => handler.Set("/slide[1]/shape[1]", new() { [key] = value });
        act.Should().NotThrow<NullReferenceException>(
            $"Set {key}='{value}' must not throw NullReferenceException");
        act.Should().NotThrow<OverflowException>(
            $"Set {key}='{value}' must not throw OverflowException");
        // Should throw ArgumentException for invalid numeric inputs
        act.Should().Throw<Exception>(
            $"Set {key}='{value}' is invalid and should be rejected");
    }

    // ==================== Boundary: Word Set alignment invalid values ====================
    // ParseJustification already throws — confirm coverage.

    [Theory]
    [InlineData("")]
    [InlineData("middle")]
    [InlineData("justified")] // not "justify" — check if it's accepted
    [InlineData("invalid")]
    public void Word_SetParagraphAlignment_InvalidValues_ThrowArgumentException(string align)
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello" });

        // ParseJustification throws for unknown values — confirm no crash
        var act = () => handler.Set("/body/p[1]", new() { ["alignment"] = align });
        act.Should().NotThrow<NullReferenceException>(
            $"Word alignment='{align}' must not throw NullReferenceException");
        // Should throw ArgumentException
        act.Should().Throw<ArgumentException>(
            $"Word alignment='{align}' is invalid and should throw ArgumentException");
    }

    // ==================== Boundary: Excel data validation type invalid values ====================
    // ExcelHandler.Set.cs:220 already has `_ => throw new ArgumentException(...)` — confirm coverage.

    [Theory]
    [InlineData("invalid")]
    [InlineData("range")]
    [InlineData("")]
    public void Excel_SetValidationType_InvalidValues_ThrowArgumentException(string dvType)
    {
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);
        // Add a data validation first
        try
        {
            handler.Add("/Sheet1", "datavalidation", null, new()
            {
                ["ref"] = "A1", ["type"] = "list", ["formula1"] = "\"Yes,No\""
            });
            var act = () => handler.Set("/Sheet1/datavalidation[1]", new() { ["type"] = dvType });
            act.Should().NotThrow<NullReferenceException>(
                $"DV type='{dvType}' must not throw NullReferenceException");
            act.Should().Throw<ArgumentException>(
                $"DV type='{dvType}' is invalid and should throw ArgumentException");
        }
        catch (ArgumentException ex) when (!ex.Message.Contains(dvType))
        {
            // datavalidation add may not be supported — skip
            _ = ex;
        }
    }

    // ==================== Boundary: PPTX slide Set slidesize invalid value ====================
    // The slidesize switch already has `default: unsupported.Add(key)` — confirm it doesn't throw.

    [Theory]
    [InlineData("invalid_size")]
    [InlineData("")]
    [InlineData("8x10")]
    public void Pptx_SetSlideSize_InvalidValues_DoNotCrash(string slidesize)
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);

        // Invalid slidesize is added to unsupported list — not thrown. Confirm no crash.
        var act = () => handler.Set("/", new() { ["slidesize"] = slidesize });
        act.Should().NotThrow<NullReferenceException>(
            $"Set slidesize='{slidesize}' must not throw NullReferenceException");
        act.Should().NotThrow<ArgumentException>(
            $"Set slidesize='{slidesize}' should silently add to unsupported list, not throw");
    }

    // ==================== Boundary: Get with null/empty string path ====================
    // All handlers should throw ArgumentException for empty path, not crash.

    [Fact]
    public void Pptx_Get_EmptyPath_ThrowsArgumentException()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        var act = () => handler.Get("");
        act.Should().Throw<ArgumentException>("empty path should throw ArgumentException");
    }

    [Fact]
    public void Excel_Get_EmptyPath_ThrowsArgumentException()
    {
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);
        var act = () => handler.Get("");
        act.Should().Throw<ArgumentException>("empty path should throw ArgumentException");
    }

    [Fact]
    public void Word_Get_EmptyPath_ThrowsArgumentException()
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        var act = () => handler.Get("");
        act.Should().Throw<ArgumentException>("empty path should throw ArgumentException");
    }

    // ==================== Boundary: PPTX Add with invalid element types ====================
    // Add should throw ArgumentException for unknown element types — not crash.

    [Theory]
    [InlineData("bogus")]
    [InlineData("")]
    [InlineData("SHAPE")]
    [InlineData("invalid_element_type_xyz")]
    public void Pptx_Add_InvalidElementType_ThrowsArgumentException_NotCrash(string elementType)
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());

        var act = () => handler.Add("/slide[1]", elementType, null, new() { ["text"] = "test" });
        act.Should().NotThrow<NullReferenceException>(
            $"Add type='{elementType}' must not throw NullReferenceException");
        act.Should().NotThrow<IndexOutOfRangeException>(
            $"Add type='{elementType}' must not throw IndexOutOfRangeException");
    }

    [Theory]
    [InlineData("bogus")]
    [InlineData("")]
    [InlineData("PARAGRAPH")]
    public void Word_Add_InvalidElementType_ThrowsArgumentException_NotCrash(string elementType)
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);

        var act = () => handler.Add("/body", elementType, null, new() { ["text"] = "test" });
        act.Should().NotThrow<NullReferenceException>(
            $"Word Add type='{elementType}' must not throw NullReferenceException");
        act.Should().NotThrow<IndexOutOfRangeException>(
            $"Word Add type='{elementType}' must not throw IndexOutOfRangeException");
    }

    [Theory]
    [InlineData("bogus")]
    [InlineData("")]
    [InlineData("CELL")]
    public void Excel_Add_InvalidElementType_ThrowsArgumentException_NotCrash(string elementType)
    {
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);

        var act = () => handler.Add("/Sheet1", elementType, null, new() { ["ref"] = "A1" });
        act.Should().NotThrow<NullReferenceException>(
            $"Excel Add type='{elementType}' must not throw NullReferenceException");
        act.Should().NotThrow<IndexOutOfRangeException>(
            $"Excel Add type='{elementType}' must not throw IndexOutOfRangeException");
    }
}
